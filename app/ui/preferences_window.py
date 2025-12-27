from pathlib import Path
from PySide6.QtWidgets import QWidget, QVBoxLayout, QComboBox, QPushButton, QHBoxLayout, QMessageBox, QApplication, QLabel
from PySide6.QtCore import QTranslator, QLocale, Qt, QEvent

from core.config import config


class PreferencesWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self._translator = None
        
        # Make sure this window is independent
        self.setWindowFlags(Qt.Window)
        
        self.setWindowTitle(self.tr("Preferences"))
        self.resize(300, 100)
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Add info label showing config location
        config_info = QLabel(f"Config file: {config.get_config_location()}")
        config_info.setWordWrap(True)
        config_info.setStyleSheet("color: gray; font-size: 9px;")
        layout.addWidget(config_info)
        
        # Find translations directory
        cur = Path(__file__).resolve()
        trans_dir = None
        for p in cur.parents:
            candidate = p / "translations"
            if candidate.exists() and any(candidate.glob("*.qm")):
                trans_dir = candidate
                break
        
        self.combo = QComboBox()
        if trans_dir:
            qm_files = sorted(trans_dir.glob("*.qm"))
            current_language = config.language
            for i, f in enumerate(qm_files):
                code = f.stem
                name = QLocale(code).nativeLanguageName() or code
                self.combo.addItem(f"{name} ({code})", str(f))
                # Select current language
                if code == current_language:
                    self.combo.setCurrentIndex(i)
        else:
            self.combo.addItem(self.tr("No translations found"))
            self.combo.setEnabled(False)
        
        self.apply_btn = QPushButton(self.tr("Apply"))
        self.apply_btn.clicked.connect(self._apply_locale)
        
        hl = QHBoxLayout()
        hl.addWidget(self.combo)
        hl.addWidget(self.apply_btn)
        layout.addLayout(hl)
    
    def _apply_locale(self):
        if not self.combo.isEnabled():
            return
        path = self.combo.currentData()
        if not path:
            return
        
        # Extract language code from path
        language_code = Path(path).stem
        
        translator = QTranslator(self)
        if not translator.load(path):
            QMessageBox.warning(self, self.tr("Error"), self.tr("Failed to load translation file."))
            return
        
        app = QApplication.instance()
        # Remove previous translator if present
        old = getattr(self.parent_window, "_translator", None)
        if old:
            app.removeTranslator(old)
        app.installTranslator(translator)
        
        # Store translator in parent window
        if self.parent_window:
            self.parent_window._translator = translator
        
        # Save language to config
        config.language = language_code
        
        # Send language change event to trigger retranslation
        if self.parent_window:
            event = QEvent(QEvent.Type.LanguageChange)
            app.sendEvent(self.parent_window, event)
            
        # Send language change event to this window
        event = QEvent(QEvent.Type.LanguageChange)
        app.sendEvent(self, event)
        
        QMessageBox.information(self, self.tr("Language changed"),
                               self.tr("Language changed. Some strings may require restarting the application."))
    
    def _retranslate_ui(self):
        """Retranslate all UI elements"""
        self.setWindowTitle(self.tr("Preferences"))
        
        # Update button text using direct reference
        self.apply_btn.setText(self.tr("Apply"))
            
        # Update combo box if no translations found
        if not self.combo.isEnabled():
            self.combo.setItemText(0, self.tr("No translations found"))
    
    def changeEvent(self, event):
        """Handle language change events"""
        if event.type() == QEvent.Type.LanguageChange:
            self._retranslate_ui()
        super().changeEvent(event)