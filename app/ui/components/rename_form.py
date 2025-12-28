import os
from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
    QTextEdit, QDateEdit, QPushButton, QFrame, QGroupBox,
    QSizePolicy, QSpacerItem, QMessageBox
)
from PySide6.QtCore import Qt, QDate, Signal
from PySide6.QtGui import QFont


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
        self._setup_ui()
        
    def _setup_ui(self):
        """Setup the rename form UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Set size policy - expanding in both directions for middle column
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        
        # Title
        title_label = QLabel(self.tr("Rename Information"))
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(11)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # Form group
        form_group = QGroupBox(self.tr("Document Details"))
        form_layout = QVBoxLayout(form_group)
        
        # Date picker
        self._setup_date_field(form_layout)
        
        # Organization name
        self._setup_organization_field(form_layout)
        
        # Subject
        self._setup_subject_field(form_layout)
        
        # Receiver name
        self._setup_receiver_field(form_layout)
        
        layout.addWidget(form_group)
        
        # Preview and action area
        self._setup_preview_area(layout)
        
        # Initially disable the form
        self._set_form_enabled(False)
        
    def _setup_date_field(self, parent_layout):
        """Setup the date picker field"""
        date_layout = QHBoxLayout()
        
        date_label = QLabel(self.tr("Date:"))
        date_label.setMinimumWidth(100)
        date_layout.addWidget(date_label)
        
        self._date_edit = QDateEdit()
        self._date_edit.setDate(QDate.currentDate())
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDisplayFormat("yyyy-MM-dd")
        self._date_edit.dateChanged.connect(self._on_form_changed)
        date_layout.addWidget(self._date_edit)
        
        date_layout.addStretch()
        parent_layout.addLayout(date_layout)
        
    def _setup_organization_field(self, parent_layout):
        """Setup the organization name field"""
        org_layout = QHBoxLayout()
        
        org_label = QLabel(self.tr("Organization:"))
        org_label.setMinimumWidth(100)
        org_layout.addWidget(org_label)
        
        self._organization_edit = QLineEdit()
        self._organization_edit.setPlaceholderText(self.tr("Enter organization name"))
        self._organization_edit.textChanged.connect(self._on_form_changed)
        org_layout.addWidget(self._organization_edit)
        
        parent_layout.addLayout(org_layout)
        
    def _setup_subject_field(self, parent_layout):
        """Setup the subject field"""
        subject_layout = QHBoxLayout()
        
        subject_label = QLabel(self.tr("Subject:"))
        subject_label.setMinimumWidth(100)
        subject_layout.addWidget(subject_label)
        
        self._subject_edit = QLineEdit()
        self._subject_edit.setPlaceholderText(self.tr("Enter document subject or description"))
        self._subject_edit.textChanged.connect(self._on_form_changed)
        subject_layout.addWidget(self._subject_edit)
        
        parent_layout.addLayout(subject_layout)
        
    def _setup_receiver_field(self, parent_layout):
        """Setup the receiver name field"""
        receiver_layout = QHBoxLayout()
        
        receiver_label = QLabel(self.tr("Receiver:"))
        receiver_label.setMinimumWidth(100)
        receiver_layout.addWidget(receiver_label)
        
        self._receiver_edit = QLineEdit()
        self._receiver_edit.setPlaceholderText(self.tr("Enter receiver name"))
        self._receiver_edit.textChanged.connect(self._on_form_changed)
        receiver_layout.addWidget(self._receiver_edit)
        
        parent_layout.addLayout(receiver_layout)
        
    def _setup_preview_area(self, parent_layout):
        """Setup the filename preview and action area"""
        # New name section
        name_layout = QHBoxLayout()
        
        name_label = QLabel(self.tr("New name:"))
        name_label.setMinimumWidth(100)
        name_layout.addWidget(name_label)
        
        # Simple filename preview
        self._preview_label = QLabel(self.tr("Select a document to see new filename"))
        self._preview_label.setWordWrap(True)
        name_layout.addWidget(self._preview_label)
        
        parent_layout.addLayout(name_layout)
        
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
            self._set_form_enabled(True)
            self._update_preview()
        else:
            self.current_extension = ""
            self._set_form_enabled(False)
            self._preview_label.setText(self.tr("Select a document to see new filename"))
            
    def _set_form_enabled(self, enabled):
        """Enable or disable the form"""
        self._date_edit.setEnabled(enabled)
        self._organization_edit.setEnabled(enabled)
        self._subject_edit.setEnabled(enabled)
        self._receiver_edit.setEnabled(enabled)
        self._clear_button.setEnabled(enabled)
        self._rename_button.setEnabled(enabled and self._is_form_valid())
        
    def _is_form_valid(self):
        """Check if all required fields are filled"""
        return (
            bool(self._organization_edit.text().strip()) and
            bool(self._subject_edit.text().strip()) and
            bool(self._receiver_edit.text().strip())
        )
        
    def _on_form_changed(self):
        """Handle form field changes"""
        self._update_preview()
        self._rename_button.setEnabled(self._is_form_valid() and bool(self.current_file_path))
        
        # Emit preview signal
        if self.current_file_path:
            preview_filename = self._generate_filename()
            self.form_changed.emit(preview_filename)
            
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
        organization = self._sanitize_filename(self._organization_edit.text().strip())
        subject = self._sanitize_filename(self._subject_edit.text().strip())
        receiver = self._sanitize_filename(self._receiver_edit.text().strip())
        
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
        """Remove invalid characters from filename component"""
        if not text:
            return ""
            
        # Replace invalid filename characters
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            text = text.replace(char, '_')
            
        # Replace multiple spaces with single space
        text = ' '.join(text.split())
        
        # Limit length
        if len(text) > 50:
            text = text[:47] + "..."
            
        return text
        
    def _clear_form(self):
        """Clear all form fields"""
        self._date_edit.setDate(QDate.currentDate())
        self._organization_edit.clear()
        self._subject_edit.clear()
        self._receiver_edit.clear()
        self._update_preview()
        
    def _on_rename_clicked(self):
        """Handle rename button click"""
        if not self.current_file_path or not self._is_form_valid():
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
            'date': self._date_edit.date().toString("yyyy-MM-dd"),
            'organization': self._organization_edit.text().strip(),
            'subject': self._subject_edit.text().strip(),
            'receiver': self._receiver_edit.text().strip()
        }
        
    def set_form_data(self, data):
        """Set form data from dictionary"""
        if 'date' in data:
            date = QDate.fromString(data['date'], "yyyy-MM-dd")
            if date.isValid():
                self._date_edit.setDate(date)
                
        if 'organization' in data:
            self._organization_edit.setText(data['organization'])
            
        if 'subject' in data:
            self._subject_edit.setText(data['subject'])
            
        if 'receiver' in data:
            self._receiver_edit.setText(data['receiver'])
            
        self._update_preview()
        
    def retranslate_ui(self):
        """Retranslate all UI elements"""
        self._update_preview()
        
    def tr(self, text):
        """Translation method - uses parent's tr if available"""
        if self.parent() and hasattr(self.parent(), 'tr'):
            return self.parent().tr(text)
        return text
