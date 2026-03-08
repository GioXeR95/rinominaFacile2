import os
import re
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
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
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

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

    def _setup_preview_area(self, parent_layout):
        """Setup the filename preview and action area"""
        # New name group (styled like Document Details)
        self._name_group = QGroupBox(self.tr("New Filename"))
        name_layout = QVBoxLayout(self._name_group)

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
            self._preview_label.setText(self.tr("Select a document to see new filename"))

    def _set_form_enabled(self, enabled):
        """Enable or disable the form"""
        self._date_edit.setEnabled(enabled)
        self._calendar.setEnabled(enabled)
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
        """Sanitize and clamp text directly in the input field while typing."""
        value = field.toPlainText()
        cleaned_value = self._sanitize_filename(value)

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
        self._preview_label.setText(preview_filename)

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

        # Replace multiple spaces with single space
        text = " ".join(text.split()).strip(" -_")

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

        # Show confirmation dialog
        reply = QMessageBox.question(
            self, 
            self.tr("Confirm Rename"),
            self.tr(f"Rename file to:\n\n{new_filename}\n\nAre you sure?"),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
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
        self._name_group.setTitle(self.tr("New Filename"))

        # Update labels
        self._date_label.setText(self.tr("Date:"))
        self._org_label.setText(self.tr("Organization:"))
        self._subject_label.setText(self.tr("Subject:"))
        self._receiver_label.setText(self.tr("Receiver:"))

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
            self._preview_label.setText(
                self.tr("Select a document to see new filename")
            )
        else:
            # Update preview with current data
            self._update_preview()

    def tr(self, text):
        """Translation method - uses QCoreApplication.translate with class context"""
        return QCoreApplication.translate("RenameForm", text)
