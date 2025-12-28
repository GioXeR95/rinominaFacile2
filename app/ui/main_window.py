import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout, 
    QPushButton, QFileDialog, QListWidget, QListWidgetItem, QSplitter
)
from PySide6.QtCore import QEvent, QTranslator, Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent

from ui.preferences_window import PreferencesWindow
from ui.toolbar.menu_bar import MenuBar
from ui.components.file_preview import FilePreview
from ui.components.rename_form import RenameForm
from core.config import config

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(self.tr("Easy Renamer 2"))
        self.resize(1000, 600)  # Increased width to accommodate preview
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        
        # Store selected files
        self.selected_files = []
        
        self._setup_ui()
        
        # Create menu bar using the reusable component
        self.menu_bar_component = MenuBar(parent_window=self)
        self.menu_bar_component.create_menu_bar(self)
        
        # Load saved language AFTER all UI elements are created
        self._load_saved_language()
    
    def _setup_ui(self):
        """Setup the main user interface"""
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # Title label - compact header
        self._main_label = QLabel(self.tr("Select documents to rename"))
        self._main_label.setAlignment(Qt.AlignCenter)
        self._main_label.setMaximumHeight(30)  # Limit height to save space
        self._main_label.setStyleSheet("font-weight: bold; padding: 5px;")
        main_layout.addWidget(self._main_label)
        
        # File selection buttons - compact layout
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self._select_file_btn = QPushButton(self.tr("Select Documents"))
        self._select_file_btn.clicked.connect(self._select_files)
        self._select_file_btn.setMaximumHeight(35)  # Compact button height
        button_layout.addWidget(self._select_file_btn)
        
        self._clear_files_btn = QPushButton(self.tr("Clear Documents"))
        self._clear_files_btn.clicked.connect(self._clear_files)
        self._clear_files_btn.setEnabled(False)  # Initially disabled
        self._clear_files_btn.setMaximumHeight(35)  # Compact button height
        button_layout.addWidget(self._clear_files_btn)
        
        # Create a widget to contain buttons and limit its height
        button_widget = QWidget()
        button_widget.setLayout(button_layout)
        button_widget.setMaximumHeight(50)  # Limit total button area height
        main_layout.addWidget(button_widget)
        
        # Create horizontal splitter for three-column layout: file list, rename form, preview
        content_splitter = QSplitter(Qt.Horizontal)
        content_splitter.setHandleWidth(8)  # Make splitter handle more visible (increased from 5 to 8)
        
        # Style the splitter handles to make them more noticeable
        content_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #d0d0d0;
                border: 1px solid #b0b0b0;
                border-radius: 4px;
                margin: 2px;
            }
            QSplitter::handle:hover {
                background-color: #c0c0c0;
                border-radius: 4px;
            }
            QSplitter::handle:pressed {
                background-color: #a0a0a0;
                border-radius: 4px;
            }
        """)
        
        # Left side - File list and drop zone
        left_widget = QWidget()
        left_widget.setMinimumWidth(250)  # Reduce minimum width to make room for middle column
        left_layout = QVBoxLayout(left_widget)
        
        # File list display
        self._file_list = QListWidget()
        self._file_list.setMinimumHeight(200)
        self._file_list.itemClicked.connect(self._on_file_selected)
        left_layout.addWidget(self._file_list)
        
        # Drop zone label
        self._drop_label = QLabel(self.tr("Or drag and drop documents here"))
        self._drop_label.setAlignment(Qt.AlignCenter)
        self._drop_label.setMaximumHeight(80)  # Limit height so it doesn't grow too much
        self._drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                color: #666;
                background-color: #f9f9f9;
            }
        """)
        left_layout.addWidget(self._drop_label)
        
        # Middle - Rename form (vertical layout)
        self._rename_form = RenameForm(self)
        self._rename_form.setMinimumWidth(280)  # Reduced minimum width
        # Remove maximum width constraint to allow horizontal resizing
        self._rename_form.rename_requested.connect(self._on_rename_requested)
        
        # Right side - File preview
        self._file_preview = FilePreview(self)
        self._file_preview.setMinimumWidth(350)  # Ensure minimum width for preview
        
        # Add widgets to main horizontal splitter (three columns)
        content_splitter.addWidget(left_widget)        # Left: File list
        content_splitter.addWidget(self._rename_form)   # Middle: Rename form
        content_splitter.addWidget(self._file_preview)  # Right: Preview
        
        # Set initial splitter sizes (25% left, 30% middle, 45% right)
        content_splitter.setSizes([250, 300, 450])
        content_splitter.setChildrenCollapsible(False)
        content_splitter.setStretchFactor(0, 0)  # Left side doesn't stretch much
        content_splitter.setStretchFactor(1, 1)  # Middle (form) can stretch and shrink
        content_splitter.setStretchFactor(2, 1)  # Right side (preview) can also stretch and shrink
        
        # Add splitter to main layout with stretch to take most space
        main_layout.addWidget(content_splitter, stretch=1)  # Give splitter all remaining space
        
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
    
    def _retranslate_ui(self):
        """Retranslate all UI elements"""
        self.setWindowTitle(self.tr("Easy Renamer 2"))
        
        # Update buttons
        self._select_file_btn.setText(self.tr("Select Documents"))
        self._clear_files_btn.setText(self.tr("Clear Documents"))
        self._drop_label.setText(self.tr("Or drag and drop documents here"))
        
        # Update status
        self._update_status()
        
        # Update file preview
        if hasattr(self, '_file_preview'):
            self._file_preview.retranslate_ui()
        
        # Update rename form
        if hasattr(self, '_rename_form'):
            self._rename_form.retranslate_ui()
        
        # Update menu bar using the component
        if hasattr(self, 'menu_bar_component'):
            self.menu_bar_component.retranslate_ui()
    
    def changeEvent(self, event):
        """Handle language change events"""
        if event.type() == QEvent.Type.LanguageChange:
            self._retranslate_ui()
        super().changeEvent(event)
    
    def _load_saved_language(self):
        """Load and apply saved language"""
        from pathlib import Path
        from PySide6.QtWidgets import QApplication
        
        saved_language = config.language
        if saved_language == "en":
            return  # English is default, no translation needed
        
        # Find translations directory
        cur = Path(__file__).resolve()
        trans_dir = None
        for p in cur.parents:
            candidate = p / "translations"
            if candidate.exists() and any(candidate.glob("*.qm")):
                trans_dir = candidate
                break
        
        if trans_dir:
            translation_file = trans_dir / f"{saved_language}.qm"
            if translation_file.exists():
                self._translator = QTranslator(self)
                if self._translator.load(str(translation_file)):
                    app = QApplication.instance()
                    app.installTranslator(self._translator)
                    self._retranslate_ui()
    
    def _select_files(self):
        """Open file dialog to select files"""
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setWindowTitle(self.tr("Select Documents to Rename"))
        
        # Set file filters focused on documents and images (for scanned documents)
        file_dialog.setNameFilter(self.tr("Documents (*.pdf *.doc *.docx *.txt *.rtf *.odt);; Images (*.png *.jpg *.jpeg *.gif *.bmp *.tiff);; All Files (*.*)"))
        
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            self._add_files(selected_files)
    
    def _clear_files(self):
        """Clear all selected files"""
        self.selected_files.clear()
        self._file_list.clear()
        self._clear_files_btn.setEnabled(False)
        self._file_preview.clear_preview()
        self._update_status()
    
    def _on_file_selected(self, item):
        """Handle file selection from the list"""
        if item:
            file_path = item.text()
            self._file_preview.preview_file(file_path)
            self._rename_form.set_current_file(file_path)
    
    def _is_supported_document(self, file_path):
        """Check if the file is a supported document type"""
        supported_extensions = {
            # Document formats
            '.pdf', '.doc', '.docx', '.txt', '.rtf', '.odt',
            # Image formats (for scanned documents)
            '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'
        }
        
        file_extension = os.path.splitext(file_path.lower())[1]
        return file_extension in supported_extensions
    
    def _add_files(self, file_paths):
        """Add files to the selection"""
        for file_path in file_paths:
            if file_path not in self.selected_files and self._is_supported_document(file_path):
                self.selected_files.append(file_path)
                
                # Add to list widget
                item = QListWidgetItem(file_path)
                self._file_list.addItem(item)
        
        # Enable clear button if we have files
        self._clear_files_btn.setEnabled(len(self.selected_files) > 0)
        self._update_status()
    
    def _update_status(self):
        """Update the status label"""
        count = len(self.selected_files)
        if count == 0:
            self._main_label.setText(self.tr("Select documents to rename"))
        elif count == 1:
            self._main_label.setText(self.tr("1 document selected"))
        else:
            self._main_label.setText(self.tr("{} documents selected").format(count))
    
    # Drag and Drop Events
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            # Add visual feedback
            self._drop_label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #0078d4;
                    border-radius: 10px;
                    padding: 20px;
                    color: #0078d4;
                    background-color: #e6f3ff;
                }
            """)
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """Handle drag leave event"""
        # Reset visual feedback
        self._drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 10px;
                padding: 20px;
                color: #666;
                background-color: #f9f9f9;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event"""
        # Reset visual feedback
        self.dragLeaveEvent(event)
        
        if event.mimeData().hasUrls():
            file_paths = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    # Only add files (not directories)
                    if os.path.isfile(file_path):
                        file_paths.append(file_path)
            
            if file_paths:
                self._add_files(file_paths)
            
            event.acceptProposedAction()
        else:
            event.ignore()
    
    def get_selected_files(self):
        """Get the list of currently selected documents"""
        return self.selected_files.copy()
    
    def has_files(self):
        """Check if any documents are selected"""
        return len(self.selected_files) > 0
    
    def _on_rename_requested(self, current_file_path, new_filename):
        """Handle rename request from the rename form"""
        try:
            import shutil
            from pathlib import Path
            from PySide6.QtWidgets import QMessageBox
            
            # Get the directory of the current file
            current_path = Path(current_file_path)
            target_path = current_path.parent / new_filename
            
            # Check if target already exists
            if target_path.exists():
                reply = QMessageBox.question(
                    self,
                    self.tr("File Exists"),
                    self.tr(f"A file with the name '{new_filename}' already exists.\n\nDo you want to replace it?"),
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if reply != QMessageBox.Yes:
                    return
            
            # Perform the rename
            shutil.move(str(current_path), str(target_path))
            
            # Update the file list
            for i in range(self._file_list.count()):
                item = self._file_list.item(i)
                if item.text() == current_file_path:
                    # Update the item text
                    item.setText(str(target_path))
                    # Update our selected_files list
                    file_index = self.selected_files.index(current_file_path)
                    self.selected_files[file_index] = str(target_path)
                    break
            
            # Update the rename form with the new path
            self._rename_form.set_current_file(str(target_path))
            
            # Update preview
            self._file_preview.preview_file(str(target_path))
            
            # Show success message
            QMessageBox.information(
                self,
                self.tr("Rename Successful"),
                self.tr(f"File successfully renamed to:\n{new_filename}")
            )
            
        except Exception as e:
            QMessageBox.critical(
                self,
                self.tr("Rename Failed"),
                self.tr(f"Failed to rename file:\n{str(e)}")
            )