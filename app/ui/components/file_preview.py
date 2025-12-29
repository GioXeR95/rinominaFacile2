import os
from pathlib import Path
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QTextEdit, 
    QFrame, QHBoxLayout, QPushButton, QSizePolicy
)
from PySide6.QtCore import Qt, QSize, Signal
from PySide6.QtGui import QPixmap, QFont, QIcon

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False


class FilePreview(QWidget):
    """Component to preview different types of documents"""
    
    # Signal emitted when preview content changes
    content_changed = Signal(str)  # file_path
    
    # Signals emitted when user wants to send text to form fields
    send_to_date_requested = Signal(str)  # selected_text
    send_to_organization_requested = Signal(str)  # selected_text
    send_to_subject_requested = Signal(str)  # selected_text
    send_to_receiver_requested = Signal(str)  # selected_text
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_file_path = None
        self.current_pdf_doc = None  # Store PDF document for navigation
        self.current_page_num = 0  # Current page number (0-based)
        self.total_pages = 0  # Total number of pages
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
        header_frame.setMaximumHeight(120)  # Increased to accommodate navigation buttons
        
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
        
        # PDF navigation controls
        self._setup_pdf_navigation(header_layout)
        
        parent_layout.addWidget(header_frame)
    
    def _setup_pdf_navigation(self, parent_layout):
        """Setup PDF page navigation controls"""
        self._nav_widget = QWidget()
        nav_layout = QHBoxLayout(self._nav_widget)
        nav_layout.setContentsMargins(0, 5, 0, 0)
        
        # Previous page button
        self._prev_page_btn = QPushButton("◀ " + self.tr("Previous"))
        self._prev_page_btn.setMaximumHeight(30)
        self._prev_page_btn.clicked.connect(self._go_to_previous_page)
        self._prev_page_btn.setEnabled(False)
        nav_layout.addWidget(self._prev_page_btn)
        
        # Page info label
        self._page_info_label = QLabel("")
        self._page_info_label.setAlignment(Qt.AlignCenter)
        self._page_info_label.setStyleSheet("font-weight: bold; color: #333;")
        nav_layout.addWidget(self._page_info_label)
        
        # Next page button
        self._next_page_btn = QPushButton(self.tr("Next") + " ▶")
        self._next_page_btn.setMaximumHeight(30)
        self._next_page_btn.clicked.connect(self._go_to_next_page)
        self._next_page_btn.setEnabled(False)
        nav_layout.addWidget(self._next_page_btn)
        
        # OCR button
        self._ocr_btn = QPushButton("🔤 " + self.tr("Extract Text"))
        self._ocr_btn.setMaximumHeight(30)
        self._ocr_btn.setStyleSheet("""
            QPushButton {
                background-color: #007acc;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:pressed {
                background-color: #004578;
            }
        """)
        self._ocr_btn.clicked.connect(self._extract_pdf_text)
        self._ocr_btn.setEnabled(False)
        nav_layout.addWidget(self._ocr_btn)
        
        # Initially hide navigation controls
        self._nav_widget.setVisible(False)
        parent_layout.addWidget(self._nav_widget)
    
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
        # Reset any previous state (OCR buttons, etc.)
        self._reset_preview_state()
        
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
    
    def _reset_preview_state(self):
        """Reset preview state when switching documents"""
        # Remove return button if it exists
        if hasattr(self, '_return_btn'):
            self._return_btn.setVisible(False)
            self._return_btn.deleteLater()
            delattr(self, '_return_btn')
        
        # Hide navigation controls initially
        self._nav_widget.setVisible(False)
    
    def _go_to_previous_page(self):
        """Go to the previous page of the PDF"""
        if self.current_page_num > 0:
            self.current_page_num -= 1
            self._render_current_pdf_page()
    
    def _go_to_next_page(self):
        """Go to the next page of the PDF"""
        if self.current_page_num < self.total_pages - 1:
            self.current_page_num += 1
            self._render_current_pdf_page()
    
    def _extract_pdf_text(self):
        """Extract text from the current PDF page or image and show it in a text window"""
        if not self.current_file_path:
            return
        
        file_info = Path(self.current_file_path)
        extension = file_info.suffix.lower()
        
        if extension == '.pdf' and self.current_pdf_doc:
            # Extract text from PDF
            self._extract_text_from_pdf()
        elif extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']:
            # Extract text from image using OCR
            self._extract_text_from_image()
        else:
            self._show_error(self.tr("Text extraction not supported for this file type"))
    
    def _extract_text_from_pdf(self):
        """Extract text from PDF page"""
        try:
            # Get the current page
            page = self.current_pdf_doc[self.current_page_num]
            
            # Extract text content
            text_content = page.get_text()
            
            if text_content.strip():
                # Show extracted text
                header = f"📄 {self.tr('Extracted Text')} - {self.tr('Page')} {self.current_page_num + 1}\n"
                self._show_extracted_text(text_content, header)
            else:
                self._show_error(self.tr("No text found on this page"))
                
        except Exception as e:
            self._show_error(f"{self.tr('Error extracting text from PDF')}: {str(e)}")
    
    def _extract_text_from_image(self):
        """Extract text from image using OCR (placeholder - requires OCR library)"""
        try:
            # This is a placeholder implementation
            # In a real application, you would use an OCR library like pytesseract
            
            # For now, show a message about OCR requirements
            message = (
                f"{self.tr('OCR (Optical Character Recognition) functionality requires additional setup.')}\n\n"
                f"{self.tr('To enable text extraction from images, you need to:')}\n"
                f"1. {self.tr('Install pytesseract: pip install pytesseract')}\n"
                f"2. {self.tr('Install Tesseract OCR engine')}\n\n"
                f"{self.tr('Image file:')} {Path(self.current_file_path).name}"
            )
            
            # Show the OCR setup message in the text view
            header = f"🔤 {self.tr('OCR Setup Required')}\n"
            self._show_extracted_text(message, header)
            
        except Exception as e:
            self._show_error(f"{self.tr('Error setting up OCR for image')}: {str(e)}")
    
    def _show_extracted_text(self, text_content, header_text=None):
        """Show extracted text in a separate window/dialog"""
        # For now, replace the current preview with the text
        # In the future, this could open a separate dialog
        self._ensure_textedit_widget()
        
        # Set the text content with proper styling
        self._preview_widget.setStyleSheet("""
            QTextEdit {
                background-color: white;
                color: black;
                border: 2px solid #007acc;
                border-radius: 8px;
                padding: 15px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 12pt;
                line-height: 1.5;
                selection-background-color: rgba(0, 120, 215, 0.3);
                selection-color: black;
            }
        """)
        
        # Add header to indicate this is extracted text
        if header_text is None:
            if self.current_pdf_doc:
                header_text = f"📄 {self.tr('Extracted Text')} - {self.tr('Page')} {self.current_page_num + 1}\n"
            else:
                header_text = f"🔤 {self.tr('Extracted Text')} - {Path(self.current_file_path).name}\n"
        
        header = header_text + "=" * 50 + "\n\n"
        
        self._preview_widget.setPlainText(header + text_content)
        
        # Add action buttons and return button
        self._add_text_action_buttons()
        self._add_return_button()
    
    def _add_text_action_buttons(self):
        """Add buttons to send selected text to form fields"""
        # Remove existing action buttons if they exist
        if hasattr(self, '_action_buttons_widget'):
            self._action_buttons_widget.setVisible(False)
            self._action_buttons_widget.deleteLater()
            delattr(self, '_action_buttons_widget')
        
        # Create a widget to hold the action buttons
        self._action_buttons_widget = QWidget()
        button_layout = QHBoxLayout(self._action_buttons_widget)
        button_layout.setContentsMargins(5, 5, 5, 5)
        button_layout.setSpacing(10)
        
        # Style for action buttons
        button_style = """
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
            QPushButton:disabled {
                background-color: #6c757d;
                color: #adb5bd;
            }
        """
        
        # Send to Date button
        self._send_to_date_btn = QPushButton("📅 " + self.tr("Send to Date"))
        self._send_to_date_btn.setStyleSheet(button_style)
        self._send_to_date_btn.clicked.connect(self._on_send_to_date)
        button_layout.addWidget(self._send_to_date_btn)
        
        # Send to Organization button
        self._send_to_org_btn = QPushButton("🏢 " + self.tr("Send to Organization"))
        self._send_to_org_btn.setStyleSheet(button_style)
        self._send_to_org_btn.clicked.connect(self._on_send_to_organization)
        button_layout.addWidget(self._send_to_org_btn)
        
        # Send to Subject button
        self._send_to_subject_btn = QPushButton("📝 " + self.tr("Send to Subject"))
        self._send_to_subject_btn.setStyleSheet(button_style)
        self._send_to_subject_btn.clicked.connect(self._on_send_to_subject)
        button_layout.addWidget(self._send_to_subject_btn)
        
        # Send to Receiver button
        self._send_to_receiver_btn = QPushButton("👤 " + self.tr("Send to Receiver"))
        self._send_to_receiver_btn.setStyleSheet(button_style)
        self._send_to_receiver_btn.clicked.connect(self._on_send_to_receiver)
        button_layout.addWidget(self._send_to_receiver_btn)
        
        # Add to main layout at the bottom
        main_layout = self.layout()
        main_layout.addWidget(self._action_buttons_widget)
    
    def _on_send_to_date(self):
        """Handle send to date button click"""
        selected_text = self._get_selected_text()
        if selected_text:
            self.send_to_date_requested.emit(selected_text)
        else:
            self._show_no_selection_message()
    
    def _on_send_to_organization(self):
        """Handle send to organization button click"""
        selected_text = self._get_selected_text()
        if selected_text:
            self.send_to_organization_requested.emit(selected_text)
        else:
            self._show_no_selection_message()
    
    def _on_send_to_subject(self):
        """Handle send to subject button click"""
        selected_text = self._get_selected_text()
        if selected_text:
            self.send_to_subject_requested.emit(selected_text)
        else:
            self._show_no_selection_message()
    
    def _on_send_to_receiver(self):
        """Handle send to receiver button click"""
        selected_text = self._get_selected_text()
        if selected_text:
            self.send_to_receiver_requested.emit(selected_text)
        else:
            self._show_no_selection_message()
    
    def _get_selected_text(self):
        """Get the currently selected text from the preview widget"""
        if isinstance(self._preview_widget, QTextEdit):
            cursor = self._preview_widget.textCursor()
            return cursor.selectedText().strip()
        return ""
    
    def _show_no_selection_message(self):
        """Show message when no text is selected"""
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(
            self,
            self.tr("No Text Selected"),
            self.tr("Please select some text first before using this button.")
        )
    
    def _add_return_button(self):
        """Add a button to return to original document view"""
        if hasattr(self, '_return_btn'):
            return  # Button already exists
            
        # Create a return button and add it to the navigation
        if self.current_pdf_doc:
            button_text = "📄 " + self.tr("Back to PDF")
        else:
            button_text = "�️ " + self.tr("Back to Image")
            
        self._return_btn = QPushButton(button_text)
        self._return_btn.setMaximumHeight(30)
        self._return_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1e7e34;
            }
        """)
        self._return_btn.clicked.connect(self._return_to_original_view)
        
        # Add to navigation layout
        nav_layout = self._nav_widget.layout()
        nav_layout.addWidget(self._return_btn)
    
    def _return_to_original_view(self):
        """Return to original document view"""
        # Clean up action buttons
        if hasattr(self, '_action_buttons_widget'):
            self._action_buttons_widget.setVisible(False)
            self._action_buttons_widget.deleteLater()
            delattr(self, '_action_buttons_widget')
        
        # Clean up return button
        if hasattr(self, '_return_btn'):
            self._return_btn.setVisible(False)
            self._return_btn.deleteLater()
            delattr(self, '_return_btn')
        
        # Re-render the current document
        if self.current_file_path:
            file_info = Path(self.current_file_path)
            extension = file_info.suffix.lower()
            
            if extension == '.pdf':
                self._render_current_pdf_page()
            elif extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']:
                self._preview_image(self.current_file_path)
    
    def _update_navigation_buttons(self):
        """Update the state of navigation buttons"""
        if not PYMUPDF_AVAILABLE or not self.current_pdf_doc:
            self._nav_widget.setVisible(False)
            return
        
        # Always show navigation controls for PDFs (including single page for OCR)
        self._nav_widget.setVisible(True)
        
        # Show all PDF navigation elements
        self._prev_page_btn.setVisible(True)
        self._next_page_btn.setVisible(True)
        self._page_info_label.setVisible(True)
        self._ocr_btn.setVisible(True)
        
        # Update OCR button text for PDFs
        self._ocr_btn.setText("🔤 " + self.tr("Extract Text"))
        
        # Update page navigation buttons
        if self.total_pages > 1:
            self._prev_page_btn.setEnabled(self.current_page_num > 0)
            self._next_page_btn.setEnabled(self.current_page_num < self.total_pages - 1)
            self._page_info_label.setText(f"{self.tr('Page')} {self.current_page_num + 1} {self.tr('of')} {self.total_pages}")
        else:
            self._prev_page_btn.setEnabled(False)
            self._next_page_btn.setEnabled(False)
            self._page_info_label.setText(f"{self.tr('Page')} 1 {self.tr('of')} 1")
        
        # Enable OCR button when PDF is loaded
        self._ocr_btn.setEnabled(True)
    
    def _preview_image(self, file_path):
        """Preview image files"""
        try:
            # Ensure we have a QLabel widget for displaying images
            self._ensure_label_widget()
            
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
            
            # Show OCR controls for images (useful for scanned documents)
            self._show_image_ocr_controls()
            
        except Exception as e:
            self._show_error(f"{self.tr('Error loading image')}: {str(e)}")
    
    def _show_image_ocr_controls(self):
        """Show OCR controls for image files"""
        # Show navigation widget with just the OCR button
        self._nav_widget.setVisible(True)
        
        # Hide page navigation buttons for images
        self._prev_page_btn.setVisible(False)
        self._next_page_btn.setVisible(False)
        self._page_info_label.setVisible(False)
        
        # Show and enable OCR button
        self._ocr_btn.setVisible(True)
        self._ocr_btn.setEnabled(True)
        
        # Update OCR button text for images
        self._ocr_btn.setText("🔤 " + self.tr("Extract Text from Image"))
    
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
        """Preview PDF files using PyMuPDF"""
        if not PYMUPDF_AVAILABLE:
            self._show_pdf_fallback(file_path)
            return
            
        try:
            # Close previous PDF document if exists
            if self.current_pdf_doc:
                self.current_pdf_doc.close()
                self.current_pdf_doc = None
            
            # Open PDF document
            self.current_pdf_doc = fitz.open(file_path)
            
            if len(self.current_pdf_doc) == 0:
                self._show_error(self.tr("PDF file appears to be empty"))
                self.current_pdf_doc.close()
                self.current_pdf_doc = None
                return
            
            # Set page navigation info
            self.current_page_num = 0  # Start with first page
            self.total_pages = len(self.current_pdf_doc)
            
            # Render the first page
            self._render_current_pdf_page()
            
            # Update navigation buttons
            self._update_navigation_buttons()
            
        except Exception as e:
            self._show_error(f"{self.tr('Error loading PDF')}: {str(e)}")
            if self.current_pdf_doc:
                self.current_pdf_doc.close()
                self.current_pdf_doc = None
    
    def _render_current_pdf_page(self):
        """Render the current page of the PDF as image"""
        if not self.current_pdf_doc or self.current_page_num >= len(self.current_pdf_doc):
            return
            
        try:
            # Get the current page
            page = self.current_pdf_doc[self.current_page_num]
            
            # Always render as image for consistent preview
            self._render_pdf_as_image(page)
            
            # Update navigation buttons
            self._update_navigation_buttons()
            
            # Add tooltip with page information
            if self.total_pages > 1:
                self._preview_widget.setToolTip(
                    self.tr(f"PDF Document - Page {self.current_page_num + 1} of {self.total_pages}")
                )
            else:
                self._preview_widget.setToolTip(self.tr("PDF Document - 1 page"))
                
        except Exception as e:
            self._show_error(f"{self.tr('Error rendering PDF page')}: {str(e)}")
    
    def _render_pdf_as_image(self, page):
        """Render PDF page as image when no text is available"""
        try:
            # Get page dimensions
            page_rect = page.rect
            
            # Calculate appropriate zoom level based on available widget size
            widget_size = self._preview_widget.size()
            available_width = max(widget_size.width() - 40, 300)
            available_height = max(widget_size.height() - 40, 400)
            
            # Calculate zoom to fit the available space
            zoom_x = available_width / page_rect.width
            zoom_y = available_height / page_rect.height
            zoom = min(zoom_x, zoom_y, 2.0)  # Cap zoom at 2.0 for readability
            
            # Render page to pixmap
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to QPixmap
            img_data = pix.tobytes("ppm")
            qpixmap = QPixmap()
            qpixmap.loadFromData(img_data)
            
            # Clean up pixmap
            pix = None
            
            if qpixmap.isNull():
                self._show_error(self.tr("Failed to render PDF page"))
                return
            
            # Ensure we have a label widget for displaying the pixmap
            self._ensure_label_widget()
            
            # Clear any existing styling and display the PDF
            self._preview_widget.setStyleSheet("")
            self._preview_widget.setPixmap(qpixmap)
                
        except Exception as e:
            self._show_error(f"{self.tr('Error rendering PDF as image')}: {str(e)}")
    
    def _ensure_textedit_widget(self):
        """Ensure the preview widget is a QTextEdit for selectable text"""
        if not isinstance(self._preview_widget, QTextEdit):
            self._content_layout.removeWidget(self._preview_widget)
            self._preview_widget.deleteLater()
            
            self._preview_widget = QTextEdit()
            self._preview_widget.setReadOnly(True)
            self._preview_widget.setMinimumSize(300, 400)
            self._preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self._preview_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self._preview_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self._content_layout.addWidget(self._preview_widget)
    
    def _show_pdf_fallback(self, file_path):
        """Show PDF fallback message when PyMuPDF is not available"""
        self._ensure_label_widget()
        self._preview_widget.setText(
            f"📄 {self.tr('PDF Document')}\n\n"
            f"{self.tr('PDF preview requires PyMuPDF library.')}\n"
            f"{self.tr('Install with: pip install PyMuPDF')}\n\n"
            f"{self.tr('File:')} {os.path.basename(file_path)}"
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
    
    def _ensure_textedit_widget(self):
        """Ensure the preview widget is a QTextEdit for selectable text"""
        if not isinstance(self._preview_widget, QTextEdit):
            self._content_layout.removeWidget(self._preview_widget)
            self._preview_widget.deleteLater()
            
            self._preview_widget = QTextEdit()
            self._preview_widget.setReadOnly(True)
            self._preview_widget.setMinimumSize(300, 400)
            self._preview_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self._preview_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self._preview_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
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
        # Close PDF document if open
        if self.current_pdf_doc:
            self.current_pdf_doc.close()
            self.current_pdf_doc = None
        
        # Reset navigation state
        self.current_page_num = 0
        self.total_pages = 0
        self._nav_widget.setVisible(False)
        
        # Clear other state
        self.current_file_path = None
        self._file_name_label.setText(self.tr("No document selected"))
        self._file_info_label.setText("")
        self._show_placeholder()
    
    def retranslate_ui(self):
        """Retranslate all UI elements"""
        # Update button texts
        if hasattr(self, '_prev_page_btn'):
            self._prev_page_btn.setText("◀ " + self.tr("Previous"))
        if hasattr(self, '_next_page_btn'):
            self._next_page_btn.setText(self.tr("Next") + " ▶")
        if hasattr(self, '_ocr_btn'):
            # Set appropriate OCR button text based on current document type
            if self.current_file_path:
                file_info = Path(self.current_file_path)
                extension = file_info.suffix.lower()
                if extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif']:
                    self._ocr_btn.setText("🔤 " + self.tr("Extract Text from Image"))
                else:
                    self._ocr_btn.setText("🔤 " + self.tr("Extract Text"))
            else:
                self._ocr_btn.setText("🔤 " + self.tr("Extract Text"))
        if hasattr(self, '_return_btn'):
            if self.current_pdf_doc:
                self._return_btn.setText("📄 " + self.tr("Back to PDF"))
            else:
                self._return_btn.setText("🖼️ " + self.tr("Back to Image"))
        
        # Update navigation if PDF is loaded
        if self.current_pdf_doc and self.total_pages > 1:
            self._page_info_label.setText(f"{self.tr('Page')} {self.current_page_num + 1} {self.tr('of')} {self.total_pages}")
        elif self.current_pdf_doc:
            self._page_info_label.setText(f"{self.tr('Page')} 1 {self.tr('of')} 1")
        
        # Update preview content
        if self.current_file_path:
            # Re-preview the current file to update language
            self.preview_file(self.current_file_path)
        else:
            self._file_name_label.setText(self.tr("No document selected"))
            self._show_placeholder()
    
    def resizeEvent(self, event):
        """Handle resize events to rescale images and PDFs"""
        super().resizeEvent(event)
        
        # If we're currently previewing content, re-render it for the new size
        if self.current_file_path:
            file_info = Path(self.current_file_path)
            extension = file_info.suffix.lower()
            
            # Check if it's an image and we have a QLabel widget
            if (extension in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'] and
                isinstance(self._preview_widget, QLabel) and
                self._preview_widget.pixmap() and not self._preview_widget.pixmap().isNull()):
                # Rescale the image for the new size
                self._preview_image(self.current_file_path)
            
            # Check if it's a PDF
            elif (extension == '.pdf' and PYMUPDF_AVAILABLE and self.current_pdf_doc):
                # Re-render the current PDF page for the new size
                self._render_current_pdf_page()

    def tr(self, text):
        """Translation method - uses parent's tr if available"""
        if self.parent() and hasattr(self.parent(), 'tr'):
            return self.parent().tr(text)
        return text