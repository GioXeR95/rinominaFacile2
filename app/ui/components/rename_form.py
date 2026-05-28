import os
import re
import time
from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QTextEdit,
    QDateEdit,
    QCalendarWidget,
    QPushButton,
    QFrame,
    QGroupBox,
    QSizePolicy,
    QSpacerItem,
    QMessageBox,
)
from PySide6.QtCore import Qt, QDate, Signal, QCoreApplication
from PySide6.QtGui import QFont
from core.config import config


class RenameForm(QWidget):
    """Component for entering file rename metadata"""

    # Signal emitted when rename is requested
    rename_requested = Signal(str, str)  # current_file_path, new_filename

    # Signal emitted when form data changes
    form_changed = Signal(str)  # preview_filename

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file_path = None
        self.current_extension = ""
        self._max_field_length = self._load_max_field_length_config()
        self._syncing_date = False
        self.setMinimumWidth(380)
        self._setup_ui()

    def _load_max_field_length_config(self):
        """Read max field length from config. Returns int or None for unlimited."""
        default_limit = 50
        raw_value = config.get("rename.max_field_length", default_limit)

        if raw_value is None:
            return None

        # bool is a subclass of int, so handle it explicitly.
        if isinstance(raw_value, bool):
            return default_limit

        if isinstance(raw_value, (int, float)):
            parsed = int(raw_value)
            return parsed if parsed > 0 else None

        if isinstance(raw_value, str):
            normalized = raw_value.strip().lower()
            unlimited_tokens = {
                "senza limiti",
                "senza limite",
                "illimitato",
                "illimitata",
                "unlimited",
                "no limit",
                "none",
                "off",
                "0",
            }
            if normalized in unlimited_tokens:
                return None

            if normalized.isdigit():
                parsed = int(normalized)
                return parsed if parsed > 0 else None

        return default_limit

    def _setup_ui(self):
        """Setup the rename form UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Set size policy - expanding in both directions for middle column
        self.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Expanding,
        )

        # Form group (removed title since it will be handled by parent QGroupBox)
        self._form_group = QGroupBox(self.tr("Document Details"))
        form_layout = QVBoxLayout(self._form_group)

        # Date picker
        self._setup_date_field(form_layout)

        # Organization name
        self._setup_organization_field(form_layout)

        # Subject
        self._setup_subject_field(form_layout)

        # Receiver name
        self._setup_receiver_field(form_layout)

        layout.addWidget(self._form_group)

        # Destination folder selector, shown only when a default storage folder exists
        self._setup_destination_folder_field(layout)

        # Preview and action area
        self._setup_preview_area(layout)

        # Initially disable the form
        self._set_form_enabled(False)

    def _setup_date_field(self, parent_layout):
        """Setup the date picker field"""
        date_container = QVBoxLayout()

        date_header = QHBoxLayout()
        self._date_label = QLabel(self.tr("Date:"))
        self._date_label.setMinimumWidth(100)
        date_header.addWidget(self._date_label)

        self._date_edit = QDateEdit()
        self._date_edit.setDate(QDate.currentDate())
        self._date_edit.setCalendarPopup(False)
        self._date_edit.setDisplayFormat("yyyy-MM-dd")
        self._date_edit.setReadOnly(True)
        self._date_edit.dateChanged.connect(self._on_date_changed)
        date_header.addWidget(self._date_edit)
        date_header.addStretch()
        date_container.addLayout(date_header)

        self._calendar = QCalendarWidget()
        self._calendar.setGridVisible(True)
        self._calendar.setMaximumHeight(260)
        self._calendar.setSelectedDate(self._date_edit.date())
        self._calendar.selectionChanged.connect(self._on_calendar_changed)
        date_container.addWidget(self._calendar)

        parent_layout.addLayout(date_container)

    def _on_date_changed(self, new_date):
        """Sync calendar when the date field changes"""
        if self._syncing_date:
            return
        self._syncing_date = True
        if self._calendar.selectedDate() != new_date:
            self._calendar.setSelectedDate(new_date)
        self._syncing_date = False
        self._on_form_changed()

    def _on_calendar_changed(self):
        """Sync date field when the calendar selection changes"""
        if self._syncing_date:
            return
        self._syncing_date = True
        self._date_edit.setDate(self._calendar.selectedDate())
        self._syncing_date = False
        self._on_form_changed()

    def _setup_organization_field(self, parent_layout):
        """Setup the organization name field"""
        org_layout = QVBoxLayout()
        self._org_label = QLabel(self.tr("Organization:"))
        org_layout.addWidget(self._org_label)

        self._organization_edit = QTextEdit()
        self._organization_edit.setPlaceholderText(self.tr("Enter organization name"))
        self._organization_edit.setFixedHeight(56)
        self._organization_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._organization_edit)
        )
        org_layout.addWidget(self._organization_edit)

        parent_layout.addLayout(org_layout)

    def _setup_subject_field(self, parent_layout):
        """Setup the subject field"""
        subject_layout = QVBoxLayout()
        self._subject_label = QLabel(self.tr("Subject:"))
        subject_layout.addWidget(self._subject_label)

        self._subject_edit = QTextEdit()
        self._subject_edit.setPlaceholderText(self.tr("Enter document subject or description"))
        self._subject_edit.setFixedHeight(56)
        self._subject_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._subject_edit)
        )
        subject_layout.addWidget(self._subject_edit)

        parent_layout.addLayout(subject_layout)

    def _setup_receiver_field(self, parent_layout):
        """Setup the receiver name field"""
        receiver_layout = QVBoxLayout()
        self._receiver_label = QLabel(self.tr("Receiver:"))
        receiver_layout.addWidget(self._receiver_label)

        self._receiver_edit = QTextEdit()
        self._receiver_edit.setPlaceholderText(self.tr("Enter receiver name"))
        self._receiver_edit.setFixedHeight(56)
        self._receiver_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._receiver_edit)
        )
        receiver_layout.addWidget(self._receiver_edit)

        parent_layout.addLayout(receiver_layout)

    def _setup_destination_folder_field(self, parent_layout):
        """Setup the destination folder selector."""
        self._destination_group = QGroupBox(self.tr("Destination folder"))
        destination_layout = QVBoxLayout(self._destination_group)

        self._destination_info_label = QLabel(
            self.tr(
                "Leave the selection empty to create the destination folder automatically."
            )
        )
        self._destination_info_label.setWordWrap(True)
        self._destination_info_label.setStyleSheet("color: gray; font-size: 11px;")
        destination_layout.addWidget(self._destination_info_label)

        destination_row = QHBoxLayout()
        destination_label = QLabel(self.tr("Folder:"))
        destination_label.setMinimumWidth(100)
        destination_row.addWidget(destination_label)

        self._destination_folder_combo = QComboBox()
        self._destination_folder_combo.currentIndexChanged.connect(
            self._on_destination_folder_changed
        )
        destination_row.addWidget(self._destination_folder_combo)
        destination_row.addStretch()

        destination_layout.addLayout(destination_row)
        parent_layout.addWidget(self._destination_group)

        self._refresh_destination_folder_options()

    def _setup_preview_area(self, parent_layout):
        """Setup the filename preview and action area"""
        # New name group (styled like Document Details)
        self._name_group = QGroupBox(self.tr("New Filename"))
        name_layout = QVBoxLayout(self._name_group)

        self._target_path_title_label = QLabel(self.tr("Destination Path:"))
        name_layout.addWidget(self._target_path_title_label)

        self._target_path_label = QLabel(
            self.tr("The path where the renamed file will be saved will be shown here")
        )
        self._target_path_label.setStyleSheet("color: gray; font-size: 11px;")
        self._target_path_label.setWordWrap(True)
        name_layout.addWidget(self._target_path_label)

        # Filename preview
        self._preview_label = QLabel(self.tr("Select a document to see new filename"))
        self._preview_label.setWordWrap(True)
        name_layout.addWidget(self._preview_label)

        parent_layout.addWidget(self._name_group)

        # Action buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self._clear_button = QPushButton(self.tr("Clear Form"))
        self._clear_button.clicked.connect(self._clear_form)
        button_layout.addWidget(self._clear_button)

        self._rename_button = QPushButton(self.tr("Rename File"))
        self._rename_button.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:disabled {
                background-color: #6c757d;
            }
        """)
        self._rename_button.clicked.connect(self._on_rename_clicked)
        button_layout.addWidget(self._rename_button)

        parent_layout.addLayout(button_layout)

    def set_current_file(self, file_path):
        """Set the current file to be renamed"""
        self.current_file_path = file_path
        if file_path:
            # Extract extension
            self.current_extension = os.path.splitext(file_path)[1]
            # Clear form when loading a new file
            self._clear_form()
            self._refresh_destination_folder_options()
            self._set_form_enabled(True)
            self._update_preview()
        else:
            self.current_extension = ""
            self._refresh_destination_folder_options()
            self._set_form_enabled(False)
            self._target_path_label.setText(
                self.tr(
                    "The path where the renamed file will be saved will be shown here"
                )
            )
            self._target_path_label.setStyleSheet("color: gray; font-size: 11px;")
            self._preview_label.setText(self.tr("Select a document to see new filename"))

    def _set_form_enabled(self, enabled):
        """Enable or disable the form"""
        self._date_edit.setEnabled(enabled)
        self._calendar.setEnabled(enabled)
        self._organization_edit.setEnabled(enabled)
        self._subject_edit.setEnabled(enabled)
        self._receiver_edit.setEnabled(enabled)
        if hasattr(self, "_destination_folder_combo"):
            self._destination_folder_combo.setEnabled(
                enabled and self._destination_group.isVisible()
            )
        self._clear_button.setEnabled(enabled)
        self._rename_button.setEnabled(enabled and self._is_form_valid())

    def _is_form_valid(self):
        """Check if form is valid - now always returns True since no fields are required"""
        return True

    def _on_form_changed(self):
        """Handle form field changes"""
        self._update_preview()
        self._rename_button.setEnabled(self._is_form_valid() and bool(self.current_file_path))

        # Emit preview signal
        if self.current_file_path:
            preview_filename = self._generate_filename()
            self.form_changed.emit(preview_filename)

    def _on_limited_text_changed(self, field: QTextEdit):
        """Sanitize and clamp text directly in the input field while typing.

        Keep user-typed spaces intact here; full normalization happens when building
        the final filename.
        """
        value = field.toPlainText()

        # Keep spaces as typed in the editor, only strip filesystem-invalid chars.
        cleaned_value = value.replace("\n", " ").replace("\r", " ").replace("\t", " ")
        cleaned_value = re.sub(r'[<>:"/\\|?*,]+', " ", cleaned_value)
        cleaned_value = re.sub(r"[\x00-\x1f]", " ", cleaned_value)

        if (
            self._max_field_length is not None
            and len(cleaned_value) > self._max_field_length
        ):
            cleaned_value = cleaned_value[: self._max_field_length]

        if cleaned_value != value:
            field.blockSignals(True)
            field.setPlainText(cleaned_value)
            field.moveCursor(field.textCursor().MoveOperation.End)
            field.blockSignals(False)

        self._on_form_changed()

    def _on_destination_folder_changed(self, *_args):
        """Refresh the preview when the destination folder selection changes."""
        self._on_form_changed()

    def _refresh_destination_folder_options(self):
        """Load the destination folder dropdown from the configured storage folder."""
        configured_storage_raw = config.get("default_storage_folder", "")
        configured_storage_folder = (
            configured_storage_raw.strip()
            if isinstance(configured_storage_raw, str)
            else ""
        )

        if not hasattr(self, "_destination_group"):
            return

        if not configured_storage_folder:
            self._destination_group.setVisible(False)
            self._destination_folder_combo.blockSignals(True)
            self._destination_folder_combo.clear()
            self._destination_folder_combo.blockSignals(False)
            return

        previous_selection = self.get_selected_destination_folder()
        base_folder = Path(configured_storage_folder).expanduser()

        self._destination_group.setVisible(True)
        self._destination_folder_combo.blockSignals(True)
        self._destination_folder_combo.clear()
        self._destination_folder_combo.addItem("", None)

        if base_folder.exists() and base_folder.is_dir():
            try:
                subfolders = sorted(
                    (entry for entry in base_folder.iterdir() if entry.is_dir()),
                    key=lambda path: path.name.casefold(),
                )
            except OSError:
                subfolders = []

            for folder in subfolders:
                self._destination_folder_combo.addItem(folder.name, str(folder))

        self._destination_folder_combo.blockSignals(False)

        if previous_selection:
            restored_index = self._find_destination_folder_index(previous_selection)
            if restored_index is not None:
                self._destination_folder_combo.setCurrentIndex(restored_index)
                return

        self._destination_folder_combo.setCurrentIndex(0)

    def refresh_destination_folder_options(self):
        """Public wrapper to reload destination folder options and preview."""
        self._refresh_destination_folder_options()
        if self.current_file_path:
            self._update_preview()

    def _find_destination_folder_index(self, folder_path: Path | str):
        """Find the combo index for a destination folder path."""
        if not hasattr(self, "_destination_folder_combo"):
            return None

        target_path = str(folder_path)
        for index in range(self._destination_folder_combo.count()):
            item_path = self._destination_folder_combo.itemData(index)
            if item_path == target_path:
                return index
        return None

    def get_selected_destination_folder(self):
        """Return the selected destination folder, or None for automatic selection."""
        if not hasattr(self, "_destination_folder_combo"):
            return None

        selected_path = self._destination_folder_combo.currentData()
        if not selected_path:
            return None

        return Path(str(selected_path))

    def _update_preview(self):
        """Update the filename preview"""
        if not self.current_file_path:
            return

        preview_filename = self._generate_filename()
        target_path = self._build_target_preview_path(preview_filename)
        # remove the filename from the target path for cleaner display, since it's already shown in the preview
        target_dir = os.path.dirname(target_path)
        self._target_path_label.setText(target_dir)
        self._preview_label.setText(preview_filename)

    def _build_target_preview_path(self, filename: str) -> str:
        """Build destination path preview based on current config and form values."""
        if not self.current_file_path:
            return ""

        current_path = Path(self.current_file_path)
        configured_storage_raw = config.get("default_storage_folder", "")
        configured_storage_folder = (
            configured_storage_raw.strip()
            if isinstance(configured_storage_raw, str)
            else ""
        )

        if configured_storage_folder:
            selected_folder = self.get_selected_destination_folder()
            if selected_folder is not None:
                target_dir = selected_folder
            else:
                target_dir = Path(configured_storage_folder).expanduser()
                organization_folder = self.get_sanitized_organization()
                if organization_folder:
                    target_dir = target_dir / organization_folder
        else:
            target_dir = current_path.parent

        return str(target_dir / filename)

    def _generate_filename(self):
        """Generate the new filename based on form data"""
        if not self.current_file_path:
            return ""

        # Get form data
        date_str = self._date_edit.date().toString("yyyy-MM-dd")
        organization = self._sanitize_filename(
            self._organization_edit.toPlainText().strip()
        )
        subject = self._sanitize_filename(self._subject_edit.toPlainText().strip())
        receiver = self._sanitize_filename(self._receiver_edit.toPlainText().strip())

        # Generate filename parts
        parts = []
        if date_str:
            parts.append(date_str)
        if organization:
            parts.append(organization)
        if subject:
            parts.append(subject)
        if receiver:
            parts.append(receiver)

        if parts:
            filename = " - ".join(parts) + self.current_extension
        else:
            filename = f"document{self.current_extension}"

        return filename

    def _sanitize_filename(self, text):
        """Clean filename component by removing only filesystem-invalid chars."""
        if not text:
            return ""

        # Normalize line breaks and tabs so multiline fields become one filename token.
        text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")

        # Remove characters invalid on Windows (also safest cross-platform subset).
        # This intentionally keeps valid symbols like parentheses.
        text = re.sub(r'[<>:"/\\|?*,]+', " ", text)
        text = re.sub(r"[\x00-\x1f]", " ", text)

        # Normalize repeated spaces but keep space as a valid character.
        text = " ".join(text.split())

        # Keep components short enough for broad filesystem compatibility.
        if self._max_field_length is not None and len(text) > self._max_field_length:
            text = text[: self._max_field_length]

        return text

    def _clear_form(self):
        """Clear all form fields"""
        self._date_edit.setDate(QDate.currentDate())
        self._organization_edit.clear()
        self._subject_edit.clear()
        self._receiver_edit.clear()
        if hasattr(self, "_destination_folder_combo"):
            self._destination_folder_combo.setCurrentIndex(0)
        self._update_preview()

    def clear_form(self):
        """Public method to clear the form"""
        self._clear_form()

    def get_sanitized_organization(self):
        """Return organization value sanitized for safe folder names."""
        return self._sanitize_filename(self._organization_edit.toPlainText().strip())

    def _on_rename_clicked(self):
        """Handle rename button click"""
        if not self.current_file_path:
            return

        new_filename = self._generate_filename()

        confirm_message = (
            self.tr("Rename file to:")
            + "<br><br>"
            + new_filename
            + "<br><br>"
            + self.tr("Are you sure?")
        )

        # Show confirmation dialog
        reply = QMessageBox.question(
            self,
            self.tr("Confirm Rename"),
            confirm_message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.rename_requested.emit(self.current_file_path, new_filename)

    def get_form_data(self):
        """Get current form data as dictionary"""
        return {
            "date": self._date_edit.date().toString("yyyy-MM-dd"),
            "organization": self._organization_edit.toPlainText().strip(),
            "subject": self._subject_edit.toPlainText().strip(),
            "receiver": self._receiver_edit.toPlainText().strip(),
        }

    def set_form_data(self, data):
        """Set form data from dictionary"""
        if 'date' in data:
            date = QDate.fromString(data['date'], "yyyy-MM-dd")
            if date.isValid():
                self._date_edit.setDate(date)

        if 'organization' in data:
            self._organization_edit.setPlainText(data["organization"])

        if 'subject' in data:
            self._subject_edit.setPlainText(data["subject"])

        if 'receiver' in data:
            self._receiver_edit.setPlainText(data["receiver"])

        self._update_preview()

    def retranslate_ui(self):
        """Retranslate all UI elements"""
        # Update group titles
        self._form_group.setTitle(self.tr("Document Details"))
        if hasattr(self, "_destination_group"):
            self._destination_group.setTitle(self.tr("Destination folder"))
        self._name_group.setTitle(self.tr("New Filename"))
        self._target_path_title_label.setText(self.tr("Destination Path:"))

        # Update labels
        self._date_label.setText(self.tr("Date:"))
        self._org_label.setText(self.tr("Organization:"))
        self._subject_label.setText(self.tr("Subject:"))
        self._receiver_label.setText(self.tr("Receiver:"))
        if hasattr(self, "_destination_info_label"):
            self._destination_info_label.setText(
                self.tr(
                    "Leave the selection empty to create the destination folder automatically."
                )
            )

        # Update placeholders
        self._organization_edit.setPlaceholderText(self.tr("Enter organization name"))
        self._subject_edit.setPlaceholderText(
            self.tr("Enter document subject or description")
        )
        self._receiver_edit.setPlaceholderText(self.tr("Enter receiver name"))

        # Update button texts
        self._clear_button.setText(self.tr("Clear Form"))
        self._rename_button.setText(self.tr("Rename File"))

        # Update preview label if no file is selected
        if not self.current_file_path:
            self._target_path_label.setText(
                self.tr(
                    "The path where the renamed file will be saved will be shown here"
                )
            )
            self._target_path_label.setStyleSheet("color: gray; font-size: 11px;")
            self._preview_label.setText(
                self.tr("Select a document to see new filename")
            )
        else:
            # Update preview with current data
            self._update_preview()

    def tr(self, text):
        """Translation method - uses QCoreApplication.translate with class context"""
        return QCoreApplication.translate("RenameForm", text)
