import os
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QTextEdit, 
    QFrame, QHBoxLayout, QPushButton, QSizePolicy
)
from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QPixmap, QFont, QIcon


class FilePreview(QWidget):
    """Component to preview different types of documents"""
    
    # Signal emitted when preview content changes
    content_changed = Signal(str)  # file_path
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file_path = None
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup the preview UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Set size policy to expand in both directions
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Header with file info
        self._setup_header(layout)
        
        # Preview content area
        self._setup_content_area(layout)
        
        # Initially show placeholder
        self._show_placeholder()
    
    def _setup_header(self, parent_layout):
        """Setup the header section with file info"""
        header_frame = QFrame()
        header_frame.setFrameStyle(QFrame.Box)
        header_frame.setMaximumHeight(80)
        
        header_layout = QVBoxLayout(header_frame)
        
        # File name label
        self._file_name_label = QLabel(self.tr("No document selected"))
        font = QFont()
        font.setBold(True)
        font.setPointSize(12)
        self._file_name_label.setFont(font)
        self._file_name_label.setWordWrap(True)
        header_layout.addWidget(self._file_name_label)
        
        # File info label
        self._file_info_label = QLabel("")
        self._file_info_label.setStyleSheet("color: #666;")
        header_layout.addWidget(self._file_info_label)
        
        parent_layout.addWidget(header_frame)
    
    def _setup_content_area(self, parent_layout):
        """Setup the scrollable content area"""
        # Create scroll area
        self._scroll_area = QScrollArea()
        self._scroll_area.setWidgetResizable(True)
        self._scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self._scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Content widget inside scroll area
        self._content_widget = QWidget()
        self._content_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._content_layout = QVBoxLayout(self._content_widget)
        
        # Preview label/widget
        self._preview_widget = QLabel()
        self._preview_widget.setAlignment(Qt.AlignCenter)
        self._preview_widget.setMinimumSize(300, 400)
        self._preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._preview_widget.setScaledContents(False)  # Important for images
        self._content_layout.addWidget(self._preview_widget)
        
        self._scroll_area.setWidget(self._content_widget)
        parent_layout.addWidget(self._scroll_area)
    
    def preview_file(self, file_path):
        """Preview the given file"""
        if not file_path or not os.path.exists(file_path):
            self._show_error(self.tr("File not found"))
            return
        
        self.current_file_path = file_path
        file_info = Path(file_path)
        
        # Update header
        self._file_name_label.setText(file_info.name)
        
        # Get file size and extension
        file_size = self._format_file_size(file_info.stat().st_size)
        file_ext = file_info.suffix.upper()
        self._file_info_label.setText(f"{file_ext} • {file_size}")
        
        # Preview based on file type
        self._preview_by_type(file_path, file_info.suffix.lower())
        
        # Emit signal
        self.content_changed.emit(file_path)
    
    def _preview_by_type(self, file_path, extension):
        """Preview file based on its type"""
        if extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']:
            self._preview_image(file_path)
        elif extension in ['.txt', '.rtf']:
            self._preview_text(file_path)
        elif extension in ['.pdf']:
            self._preview_pdf(file_path)
        elif extension in ['.doc', '.docx', '.odt']:
            self._preview_document(file_path)
        else:
            self._show_unsupported(extension)
    
    def _preview_image(self, file_path):
        """Preview image files"""
        try:
            pixmap = QPixmap(file_path)
            if pixmap.isNull():
                self._show_error(self.tr("Cannot load image"))
                return
            
            # Scale image to fit the available space while maintaining aspect ratio
            # Get current widget size, with some padding
            widget_size = self._preview_widget.size()
            available_width = max(widget_size.width() - 40, 300)  # Min 300px
            available_height = max(widget_size.height() - 40, 400)  # Min 400px
            max_size = QSize(available_width, available_height)
            
            scaled_pixmap = pixmap.scaled(max_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            
            self._preview_widget.setPixmap(scaled_pixmap)
            self._preview_widget.setText("")
            
        except Exception as e:
            self._show_error(f"{self.tr('Error loading image')}: {str(e)}")
    
    def _preview_text(self, file_path):
        """Preview text files"""
        try:
            # Replace the label with a text edit for better text display
            if not isinstance(self._preview_widget, QTextEdit):
                self._content_layout.removeWidget(self._preview_widget)
                self._preview_widget.deleteLater()
                
                self._preview_widget = QTextEdit()
                self._preview_widget.setReadOnly(True)
                self._preview_widget.setMinimumSize(300, 400)
                self._preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                self._content_layout.addWidget(self._preview_widget)
            
            # Try different encodings
            encodings = ['utf-8', 'latin-1', 'cp1252']
            content = None
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as file:
                        content = file.read(10000)  # Limit to first 10KB
                        break
                except UnicodeDecodeError:
                    continue
            
            if content is not None:
                if len(content) >= 10000:
                    content += f"\n\n{self.tr('[File truncated - showing first 10KB]')}"
                self._preview_widget.setPlainText(content)
            else:
                self._show_error(self.tr("Cannot decode text file"))
                
        except Exception as e:
            self._show_error(f"{self.tr('Error reading text file')}: {str(e)}")
    
    def _preview_pdf(self, file_path):
        """Preview PDF files - placeholder for now"""
        self._ensure_label_widget()
        self._preview_widget.setText(
            f"📄 {self.tr('PDF Document')}\n\n"
            f"{self.tr('PDF preview not yet implemented.')}\n"
            f"{self.tr('File:')} {os.path.basename(file_path)}"
        )
        self._preview_widget.setStyleSheet("""
            QLabel {
                background-color: #f5f5f5;
                border: 2px dashed #ccc;
                border-radius: 10px;
                padding: 20px;
                color: #666;
            }
        """)
    
    def _preview_document(self, file_path):
        """Preview Word documents - placeholder for now"""
        self._ensure_label_widget()
        file_type = "Word Document" if file_path.lower().endswith(('.doc', '.docx')) else "Document"
        self._preview_widget.setText(
            f"📄 {self.tr(file_type)}\n\n"
            f"{self.tr('Document preview not yet implemented.')}\n"
            f"{self.tr('File:')} {os.path.basename(file_path)}"
        )
        self._preview_widget.setStyleSheet("""
            QLabel {
                background-color: #f5f5f5;
                border: 2px dashed #ccc;
                border-radius: 10px;
                padding: 20px;
                color: #666;
            }
        """)
    
    def _show_unsupported(self, extension):
        """Show unsupported file type message"""
        self._ensure_label_widget()
        self._preview_widget.setText(
            f"❓ {self.tr('Unsupported Format')}\n\n"
            f"{self.tr('Cannot preview')} {extension.upper()} {self.tr('files')}"
        )
        self._preview_widget.setStyleSheet("""
            QLabel {
                background-color: #fff3cd;
                border: 2px dashed #ffc107;
                border-radius: 10px;
                padding: 20px;
                color: #856404;
            }
        """)
    
    def _show_error(self, message):
        """Show error message"""
        self._ensure_label_widget()
        self._preview_widget.setText(f"❌ {self.tr('Error')}\n\n{message}")
        self._preview_widget.setStyleSheet("""
            QLabel {
                background-color: #f8d7da;
                border: 2px dashed #dc3545;
                border-radius: 10px;
                padding: 20px;
                color: #721c24;
            }
        """)
    
    def _show_placeholder(self):
        """Show placeholder when no file is selected"""
        self._ensure_label_widget()
        self._preview_widget.setText(
            f"📄 {self.tr('Document Preview')}\n\n"
            f"{self.tr('Select a document from the list to preview it here')}"
        )
        self._preview_widget.setStyleSheet("""
            QLabel {
                background-color: #e9ecef;
                border: 2px dashed #6c757d;
                border-radius: 10px;
                padding: 20px;
                color: #495057;
            }
        """)
    
    def _ensure_label_widget(self):
        """Ensure the preview widget is a QLabel (not QTextEdit)"""
        if not isinstance(self._preview_widget, QLabel):
            self._content_layout.removeWidget(self._preview_widget)
            self._preview_widget.deleteLater()
            
            self._preview_widget = QLabel()
            self._preview_widget.setAlignment(Qt.AlignCenter)
            self._preview_widget.setMinimumSize(300, 400)
            self._preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self._preview_widget.setWordWrap(True)
            self._preview_widget.setScaledContents(False)
            self._content_layout.addWidget(self._preview_widget)
    
    def _format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        size = float(size_bytes)
        
        while size >= 1024.0 and i < len(size_names) - 1:
            size /= 1024.0
            i += 1
        
        return f"{size:.1f} {size_names[i]}"
    
    def clear_preview(self):
        """Clear the current preview"""
        self.current_file_path = None
        self._file_name_label.setText(self.tr("No document selected"))
        self._file_info_label.setText("")
        self._show_placeholder()
    
    def retranslate_ui(self):
        """Retranslate all UI elements"""
        if self.current_file_path:
            # Re-preview the current file to update language
            self.preview_file(self.current_file_path)
        else:
            self._file_name_label.setText(self.tr("No document selected"))
            self._show_placeholder()
    
    def resizeEvent(self, event):
        """Handle resize events to rescale images"""
        super().resizeEvent(event)
        
        # If we're currently previewing an image, rescale it
        if (self.current_file_path and 
            isinstance(self._preview_widget, QLabel) and 
            self._preview_widget.pixmap() and 
            not self._preview_widget.pixmap().isNull()):
            
            # Get file extension to check if it's an image
            file_info = Path(self.current_file_path)
            extension = file_info.suffix.lower()
            if extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']:
                # Rescale the image for the new size
                self._preview_image(self.current_file_path)

    def tr(self, text):
        """Translation method - uses parent's tr if available"""
        if self.parent() and hasattr(self.parent(), 'tr'):
            return self.parent().tr(text)
        return text