import os
import re
import time
from datetime import datetime
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
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
        self._custom_folders = self._load_custom_folders_config()
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
        form_layout.setContentsMargins(12, 16, 12, 12)
        form_layout.setSpacing(8)
        form_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._form_group.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Maximum,
        )

        # Date picker
        self._setup_date_field(form_layout)

        # Custom folder
        self._setup_custom_folder_field(form_layout)

        # Organization / Subject / Receiver in two columns
        self._details_grid = QGridLayout()
        self._details_grid.setContentsMargins(0, 0, 0, 0)
        self._details_grid.setHorizontalSpacing(10)
        self._details_grid.setVerticalSpacing(8)
        self._details_grid.setColumnStretch(0, 0)
        self._details_grid.setColumnStretch(1, 1)

        # Organization name
        self._setup_organization_field(self._details_grid, 0)

        # Subject
        self._setup_subject_field(self._details_grid, 1)

        # Receiver name
        self._setup_receiver_field(self._details_grid, 2)

        form_layout.addLayout(self._details_grid)

        layout.addWidget(self._form_group)

        # Preview and action area
        self._setup_preview_area(layout)

        # Initially disable the form
        self._set_form_enabled(False)

    def _setup_date_field(self, parent_layout):
        """Setup the date picker field"""
        date_container = QVBoxLayout()
        date_container.setSpacing(6)

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

    def _setup_organization_field(self, parent_layout, row: int):
        """Setup the organization name field"""
        self._org_label = QLabel(self.tr("Organization:"))
        self._org_label.setMinimumWidth(100)
        parent_layout.addWidget(self._org_label, row, 0)

        self._organization_edit = QTextEdit()
        self._organization_edit.setPlaceholderText(self.tr("Enter organization name"))
        self._organization_edit.setFixedHeight(56)
        self._organization_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._organization_edit)
        )
        parent_layout.addWidget(self._organization_edit, row, 1)

    def _load_custom_folders_config(self):
        """Load saved custom destination folders from config."""
        raw_value = config.get("rename.custom_folders", [])
        if not isinstance(raw_value, list):
            return []

        cleaned: list[str] = []
        seen: set[str] = set()
        for item in raw_value:
            if not isinstance(item, str):
                continue
            normalized = self._sanitize_filename(item.strip())
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            cleaned.append(normalized)
        return cleaned

    def _save_custom_folders_config(self):
        """Persist the custom destination folders list."""
        config.set("rename.custom_folders", self._custom_folders)

    def _setup_custom_folder_field(self, parent_layout):
        """Setup the custom folder selector and editor."""
        custom_layout = QGridLayout()
        custom_layout.setContentsMargins(0, 0, 0, 0)
        custom_layout.setHorizontalSpacing(10)
        custom_layout.setVerticalSpacing(8)
        custom_layout.setColumnStretch(0, 0)
        custom_layout.setColumnStretch(1, 1)

        self._custom_folder_label = QLabel(self.tr("Custom folder:"))
        self._custom_folder_label.setMinimumWidth(100)
        custom_layout.addWidget(self._custom_folder_label, 0, 0)

        editor_layout = QHBoxLayout()
        self._custom_folder_edit = QLineEdit()
        self._custom_folder_edit.setPlaceholderText(
            self.tr("Enter a folder name to add")
        )
        self._custom_folder_edit.returnPressed.connect(self._add_custom_folder)
        editor_layout.addWidget(self._custom_folder_edit)

        self._add_custom_folder_button = QPushButton(self.tr("Add"))
        self._add_custom_folder_button.clicked.connect(self._add_custom_folder)
        editor_layout.addWidget(self._add_custom_folder_button)

        custom_layout.addLayout(editor_layout, 0, 1)

        self._custom_folder_list_label = QLabel(self.tr("Selected:"))
        self._custom_folder_list_label.setMinimumWidth(100)
        custom_layout.addWidget(self._custom_folder_list_label, 1, 0)

        selection_layout = QHBoxLayout()

        self._custom_folder_combo = QComboBox()
        self._custom_folder_combo.addItem("")
        self._custom_folder_combo.addItems(self._custom_folders)
        self._custom_folder_combo.currentIndexChanged.connect(
            self._on_custom_folder_selection_changed
        )
        selection_layout.addWidget(self._custom_folder_combo)

        self._delete_custom_folder_button = QPushButton(self.tr("Delete"))
        self._delete_custom_folder_button.setText("-")
        self._delete_custom_folder_button.setToolTip(self.tr("Delete selected custom folder"))
        self._delete_custom_folder_button.setFixedWidth(
            self._add_custom_folder_button.sizeHint().width()
        )
        self._delete_custom_folder_button.clicked.connect(
            self._delete_selected_custom_folder
        )
        selection_layout.addWidget(self._delete_custom_folder_button)

        custom_layout.addLayout(selection_layout, 1, 1)

        self._update_custom_folder_controls()

        parent_layout.addLayout(custom_layout)

    def _add_custom_folder(self):
        """Add the typed custom folder to the saved list and select it."""
        folder_name = self._sanitize_filename(self._custom_folder_edit.text().strip())
        if not folder_name:
            return

        if folder_name not in self._custom_folders:
            self._custom_folders.append(folder_name)
            self._save_custom_folders_config()

        self._refresh_custom_folder_combo(folder_name)
        self._custom_folder_edit.clear()

    def _refresh_custom_folder_combo(self, selected_folder=None):
        """Reload the custom folder combo while keeping the blank option first."""
        if selected_folder is None:
            current_text = self.get_selected_custom_folder()
        else:
            current_text = selected_folder.strip()

        self._custom_folder_combo.blockSignals(True)
        self._custom_folder_combo.clear()
        self._custom_folder_combo.addItem("")
        self._custom_folder_combo.addItems(self._custom_folders)

        if current_text and current_text in self._custom_folders:
            self._custom_folder_combo.setCurrentText(current_text)
        else:
            self._custom_folder_combo.setCurrentIndex(0)

        self._custom_folder_combo.blockSignals(False)
        self._on_custom_folder_selection_changed()

    def _update_custom_folder_controls(self):
        """Enable delete only when a non-empty custom folder is selected."""
        has_selection = bool(self.get_selected_custom_folder())
        if hasattr(self, "_delete_custom_folder_button"):
            self._delete_custom_folder_button.setEnabled(
                self._custom_folder_combo.isEnabled() and has_selection
            )

    def _on_custom_folder_selection_changed(self):
        """Handle custom folder selection changes."""
        self._update_custom_folder_controls()
        self._on_form_changed()

    def _delete_selected_custom_folder(self):
        """Delete the currently selected custom folder after confirmation."""
        selected_folder = self.get_selected_custom_folder()
        if not selected_folder:
            self._update_custom_folder_controls()
            return

        reply = QMessageBox.question(
            self,
            self.tr("Delete Custom Folder"),
            self.tr("Delete '{folder}' from the custom folder list?").format(
                folder=selected_folder
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        if selected_folder in self._custom_folders:
            self._custom_folders.remove(selected_folder)
            self._save_custom_folders_config()

        self._custom_folder_combo.blockSignals(True)
        self._custom_folder_combo.clear()
        self._custom_folder_combo.addItem("")
        self._custom_folder_combo.addItems(self._custom_folders)
        self._custom_folder_combo.setCurrentIndex(0)
        self._custom_folder_combo.blockSignals(False)

        self._update_custom_folder_controls()
        self._on_form_changed()

    def _setup_subject_field(self, parent_layout, row: int):
        """Setup the subject field"""
        self._subject_label = QLabel(self.tr("Subject:"))
        self._subject_label.setMinimumWidth(100)
        parent_layout.addWidget(self._subject_label, row, 0)

        self._subject_edit = QTextEdit()
        self._subject_edit.setPlaceholderText(self.tr("Enter document subject or description"))
        self._subject_edit.setFixedHeight(56)
        self._subject_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._subject_edit)
        )
        parent_layout.addWidget(self._subject_edit, row, 1)

    def _setup_receiver_field(self, parent_layout, row: int):
        """Setup the receiver name field"""
        self._receiver_label = QLabel(self.tr("Receiver:"))
        self._receiver_label.setMinimumWidth(100)
        parent_layout.addWidget(self._receiver_label, row, 0)

        self._receiver_edit = QTextEdit()
        self._receiver_edit.setPlaceholderText(self.tr("Enter receiver name"))
        self._receiver_edit.setFixedHeight(56)
        self._receiver_edit.textChanged.connect(
            lambda: self._on_limited_text_changed(self._receiver_edit)
        )
        parent_layout.addWidget(self._receiver_edit, row, 1)

    def _setup_preview_area(self, parent_layout):
        """Setup the filename preview and action area"""
        # New name group (styled like Document Details)
        self._name_group = QGroupBox(self.tr("New Filename"))
        name_layout = QVBoxLayout(self._name_group)
        name_layout.setContentsMargins(12, 16, 12, 12)
        name_layout.setSpacing(8)
        name_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._name_group.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Maximum,
        )

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
            # Clear form when loading a new file
            self._clear_form()
            # Extract extension
            self.current_extension = os.path.splitext(file_path)[1]
            self._set_form_enabled(True)
            self._update_preview()
        else:
            self.current_extension = ""
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
        self._custom_folder_edit.setEnabled(enabled)
        self._add_custom_folder_button.setEnabled(enabled)
        self._custom_folder_combo.setEnabled(enabled)
        self._delete_custom_folder_button.setEnabled(
            enabled and bool(self.get_selected_custom_folder())
        )
        self._organization_edit.setEnabled(enabled)
        self._subject_edit.setEnabled(enabled)
        self._receiver_edit.setEnabled(enabled)
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
            target_dir = Path(configured_storage_folder).expanduser()
            destination_folder = self.get_destination_folder_name()
            if destination_folder:
                existing_organization_dir = self._find_organization_folder_in_depth(
                    target_dir, destination_folder
                )
                if existing_organization_dir is not None:
                    target_dir = existing_organization_dir
                else:
                    target_dir = target_dir / destination_folder
        else:
            target_dir = current_path.parent

        return str(target_dir / filename)

    def _find_organization_folder_in_depth(
        self, base_folder: Path, organization_folder: str
    ) -> Path | None:
        """Find best matching organization folder under base_folder.

        Supports exact matches and names containing extra chars, e.g.:
        - "01 asdlol"
        - "asdlol - 001"
        """
        if not organization_folder:
            return None

        if not base_folder.exists() or not base_folder.is_dir():
            return None

        target_name = self._normalize_folder_match_value(organization_folder)
        if not target_name:
            return None

        timeout_raw = config.get("rename.folder_search_timeout_seconds", 5)
        timeout_seconds = 5.0
        if isinstance(timeout_raw, (int, float)):
            timeout_seconds = float(timeout_raw)
        elif isinstance(timeout_raw, str):
            try:
                timeout_seconds = float(timeout_raw.strip())
            except ValueError:
                timeout_seconds = 5.0

        if timeout_seconds <= 0:
            timeout_seconds = 5.0
        deadline = time.monotonic() + timeout_seconds

        exact_matches: list[Path] = []
        partial_matches: list[Path] = []
        pending_dirs: list[Path] = [base_folder]
        while pending_dirs:
            if time.monotonic() >= deadline:
                break

            current_dir = pending_dirs.pop()

            # Use scandir with per-directory error handling to better tolerate
            # transient failures and permission issues on network shares.
            try:
                with os.scandir(current_dir) as entries:
                    for entry in entries:
                        if time.monotonic() >= deadline:
                            pending_dirs.clear()
                            break

                        try:
                            if not entry.is_dir(follow_symlinks=False):
                                continue
                        except OSError:
                            continue

                        candidate = Path(entry.path)
                        pending_dirs.append(candidate)

                        candidate_name = self._normalize_folder_match_value(
                            candidate.name
                        )
                        if not candidate_name:
                            continue

                        if candidate_name == target_name:
                            exact_matches.append(candidate)
                            continue

                        if f" {target_name} " in f" {candidate_name} ":
                            partial_matches.append(candidate)
            except OSError:
                continue

        # Prefer exact matches first; among many, prefer the closest/shortest path.
        if exact_matches:
            return min(exact_matches, key=lambda p: (len(p.parts), len(str(p)), str(p)))

        if partial_matches:
            return min(
                partial_matches, key=lambda p: (len(p.parts), len(str(p)), str(p))
            )

        return None

    def _normalize_folder_match_value(self, value: str) -> str:
        """Normalize a folder name for tolerant comparisons."""
        chars = [ch.casefold() if ch.isalnum() else " " for ch in value]
        return " ".join("".join(chars).split())

    def get_selected_custom_folder(self):
        """Return the selected custom folder, or empty string when not selected."""
        if not hasattr(self, "_custom_folder_combo"):
            return ""

        selected_folder = self._custom_folder_combo.currentText().strip()
        return selected_folder if selected_folder else ""

    def get_destination_folder_name(self):
        """Return the folder name used for destination storage."""
        selected_custom_folder = self.get_selected_custom_folder()
        if selected_custom_folder:
            return self._sanitize_filename(selected_custom_folder)
        return self.get_sanitized_organization()

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
        self._custom_folder_edit.clear()
        self._refresh_custom_folder_combo("")
        self._organization_edit.clear()
        self._subject_edit.clear()
        self._receiver_edit.clear()
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
            "custom_folder": self.get_selected_custom_folder(),
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

        if 'custom_folder' in data:
            custom_folder = self._sanitize_filename(data['custom_folder'].strip())
            if custom_folder:
                if custom_folder not in self._custom_folders:
                    self._custom_folders.append(custom_folder)
                    self._save_custom_folders_config()
                self._refresh_custom_folder_combo(custom_folder)
            else:
                self._refresh_custom_folder_combo("")

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
        self._name_group.setTitle(self.tr("New Filename"))
        self._target_path_title_label.setText(self.tr("Destination Path:"))

        # Update labels
        self._date_label.setText(self.tr("Date:"))
        self._custom_folder_label.setText(self.tr("Custom folder:"))
        self._custom_folder_list_label.setText(self.tr("Selected:"))
        self._org_label.setText(self.tr("Organization:"))
        self._subject_label.setText(self.tr("Subject:"))
        self._receiver_label.setText(self.tr("Receiver:"))

        # Update placeholders
        self._custom_folder_edit.setPlaceholderText(
            self.tr("Enter a folder name to add")
        )
        self._organization_edit.setPlaceholderText(self.tr("Enter organization name"))
        self._subject_edit.setPlaceholderText(
            self.tr("Enter document subject or description")
        )
        self._receiver_edit.setPlaceholderText(self.tr("Enter receiver name"))

        # Update button texts
        self._add_custom_folder_button.setText(self.tr("Add"))
        self._delete_custom_folder_button.setText(self.tr("Delete"))
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
