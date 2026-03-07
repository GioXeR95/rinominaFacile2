from pathlib import Path
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QComboBox,
    QPushButton,
    QHBoxLayout,
    QMessageBox,
    QApplication,
    QLabel,
    QLineEdit,
    QGroupBox,
    QSpinBox,
    QCheckBox,
)
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
        self.resize(420, 260)

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
        self.cancel_btn = QPushButton(self.tr("Cancel"))
        self.cancel_btn.clicked.connect(self.close)

        hl = QHBoxLayout()
        hl.addWidget(self.combo)
        layout.addLayout(hl)

        # --- Rename filename length section ---
        self._rename_group = QGroupBox(self.tr("Rename Limits"))
        rename_layout = QVBoxLayout(self._rename_group)

        self._max_length_label = QLabel(self.tr("Max characters per field:"))
        rename_layout.addWidget(self._max_length_label)

        length_row = QHBoxLayout()
        self._max_length_spin = QSpinBox()
        self._max_length_spin.setRange(1, 255)
        self._max_length_spin.setValue(50)
        length_row.addWidget(self._max_length_spin)

        self._no_limit_check = QCheckBox(self.tr("No limit"))
        self._no_limit_check.toggled.connect(
            lambda checked: self._max_length_spin.setEnabled(not checked)
        )
        length_row.addWidget(self._no_limit_check)
        length_row.addStretch()
        rename_layout.addLayout(length_row)

        layout.addWidget(self._rename_group)

        # --- Gemini API Key section ---
        self._api_group = QGroupBox(self.tr("Gemini API Key"))
        api_layout = QVBoxLayout(self._api_group)

        # Input field (masked)
        self._api_key_edit = QLineEdit()
        self._api_key_edit.setPlaceholderText(self.tr("Enter Gemini API key"))
        self._api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        api_layout.addWidget(self._api_key_edit)

        # Buttons row: Show
        btn_row = QHBoxLayout()
        self._toggle_show_btn = QPushButton(self.tr("Show"))
        self._toggle_show_btn.setCheckable(True)
        self._toggle_show_btn.clicked.connect(self._on_toggle_partial)
        btn_row.addWidget(self._toggle_show_btn)

        btn_row.addStretch()
        api_layout.addLayout(btn_row)

        layout.addWidget(self._api_group)

        # Bottom button bar: Cancel / Apply
        bottom_buttons = QHBoxLayout()
        bottom_buttons.addStretch()
        bottom_buttons.addWidget(self.cancel_btn)
        bottom_buttons.addWidget(self.apply_btn)
        layout.addLayout(bottom_buttons)

        # Initialize input field with stored key
        self._load_saved_rename_limit()
        self._load_saved_key()

    def _load_saved_rename_limit(self):
        """Load the current rename limit preference into controls."""
        raw_value = config.get("rename.max_field_length", 50)

        if raw_value is None:
            self._no_limit_check.setChecked(True)
            return

        if isinstance(raw_value, str):
            normalized = raw_value.strip().lower()
            if normalized in {
                "senza limiti",
                "senza limite",
                "illimitato",
                "illimitata",
                "unlimited",
                "no limit",
                "none",
                "off",
                "0",
            }:
                self._no_limit_check.setChecked(True)
                return
            if normalized.isdigit():
                raw_value = int(normalized)

        try:
            value = int(raw_value)
        except (TypeError, ValueError):
            value = 50

        if value <= 0:
            self._no_limit_check.setChecked(True)
            return

        self._max_length_spin.setValue(max(1, min(value, 255)))
        self._no_limit_check.setChecked(False)

    def _apply_locale(self):
        """Apply all preferences (language and API key)."""
        language_changed = False
        # Apply language if translations are available
        if self.combo.isEnabled():
            path = self.combo.currentData()
            if path:
                language_code = Path(path).stem
                translator = QTranslator(self)
                if not translator.load(path):
                    QMessageBox.warning(self, self.tr("Error"), self.tr("Failed to load translation file."))
                else:
                    app = QApplication.instance()
                    old = getattr(self.parent_window, "_translator", None)
                    if old:
                        app.removeTranslator(old)
                    app.installTranslator(translator)
                    if self.parent_window:
                        self.parent_window._translator = translator
                    # Save language to config
                    config.language = language_code
                    # Send language change events
                    if self.parent_window:
                        event = QEvent(QEvent.Type.LanguageChange)
                        app.sendEvent(self.parent_window, event)
                    event = QEvent(QEvent.Type.LanguageChange)
                    app.sendEvent(self, event)
                    language_changed = True

        # Save rename length setting
        if self._no_limit_check.isChecked():
            config.set("rename.max_field_length", "senza limiti")
            applied_limit = None
        else:
            config.set("rename.max_field_length", int(self._max_length_spin.value()))
            applied_limit = int(self._max_length_spin.value())

        # Apply setting live to open rename form when available
        if self.parent_window and hasattr(self.parent_window, "_rename_form"):
            self.parent_window._rename_form._max_field_length = applied_limit
            self.parent_window._rename_form._on_form_changed()

        # Save API key if provided
        key = self._api_key_edit.text().strip()
        if key:
            ok, error_msg = config.set_gemini_api_key_plain(key)
            if not ok:
                QMessageBox.critical(self, self.tr("Error"), 
                                   self.tr("Failed to save API key.") + "\n" + error_msg)

        # Unified confirmation
        QMessageBox.information(self, self.tr("Preferences saved"),
                                self.tr("Preferences saved"))

    def _retranslate_ui(self):
        """Retranslate all UI elements"""
        self.setWindowTitle(self.tr("Preferences"))

        # Update button text using direct reference
        self.apply_btn.setText(self.tr("Apply"))
        self.cancel_btn.setText(self.tr("Cancel"))
        # API section texts
        if hasattr(self, '_api_group'):
            self._api_group.setTitle(self.tr("Gemini API Key"))
        if hasattr(self, "_rename_group"):
            self._rename_group.setTitle(self.tr("Rename Limits"))
        if hasattr(self, "_max_length_label"):
            self._max_length_label.setText(self.tr("Max characters per field:"))
        if hasattr(self, "_no_limit_check"):
            self._no_limit_check.setText(self.tr("No limit"))
        self._api_key_edit.setPlaceholderText(self.tr("Enter Gemini API key"))
        if self._toggle_show_btn.isChecked():
            self._toggle_show_btn.setText(self.tr("Hide"))
        else:
            self._toggle_show_btn.setText(self.tr("Show"))

        # Update combo box if no translations found
        if not self.combo.isEnabled():
            self.combo.setItemText(0, self.tr("No translations found"))

    def _load_saved_key(self):
        """Load saved API key into the input field."""
        plain = config.get_gemini_api_key_plain()
        if plain:
            self._api_key_edit.setText(plain)

    def _on_toggle_partial(self):
        """Toggle between showing and hiding the API key in the input field."""
        if self._toggle_show_btn.isChecked():
            self._api_key_edit.setEchoMode(QLineEdit.EchoMode.Normal)
            self._toggle_show_btn.setText(self.tr("Hide"))
        else:
            self._api_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
            self._toggle_show_btn.setText(self.tr("Show"))

    def changeEvent(self, event):
        """Handle language change events"""
        if event.type() == QEvent.Type.LanguageChange:
            self._retranslate_ui()
        super().changeEvent(event)
