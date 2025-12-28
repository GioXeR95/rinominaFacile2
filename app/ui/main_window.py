from PySide6.QtWidgets import (
    QMainWindow, QWidget, QLabel, QVBoxLayout
)
from PySide6.QtCore import QEvent, QTranslator

from ui.preferences_window import PreferencesWindow
from ui.toolbar.menu_bar import MenuBar
from core.config import config

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(self.tr("Easy Renamer 2"))
        self.resize(600, 400)

        central_widget = QWidget()
        layout = QVBoxLayout()

        self._main_label = QLabel(self.tr("Drag files here"))
        layout.addWidget(self._main_label)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Create menu bar using the reusable component
        self.menu_bar_component = MenuBar(parent_window=self)
        self.menu_bar_component.create_menu_bar(self)
        
        # Load saved language AFTER all UI elements are created
        self._load_saved_language()
    
    def _retranslate_ui(self):
        """Retranslate all UI elements"""
        self.setWindowTitle(self.tr("Easy Renamer 2"))
        self._main_label.setText(self.tr("Drag files here"))
        
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