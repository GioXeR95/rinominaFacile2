from PySide6.QtWidgets import QMenuBar, QMenu
from PySide6.QtCore import QObject, Signal

from ui.preferences_window import PreferencesWindow


class MenuBar(QObject):
    """Reusable menu bar component that can be used across different windows"""
    
    # Signals for menu actions
    exit_requested = Signal()
    preferences_requested = Signal()
    
    def __init__(self, parent_window=None):
        super().__init__()
        self.parent_window = parent_window
        self._menu_bar = None
        self._menus = {}
        self._actions = {}
    
    def create_menu_bar(self, main_window):
        """Create and setup the menu bar for a QMainWindow"""
        self._menu_bar = main_window.menuBar()
        self._setup_menus()
        return self._menu_bar
    
    def _setup_menus(self):
        """Setup all menus and actions"""
        self._create_file_menu()
        self._create_settings_menu()
    
    def _create_file_menu(self):
        """Create the File menu"""
        if not self._menu_bar:
            return
            
        self._menus['file'] = self._menu_bar.addMenu(self.tr("File"))
        
        # Exit action
        self._actions['exit'] = self._menus['file'].addAction(self.tr("Exit"))
        self._actions['exit'].triggered.connect(self._on_exit)
    
    def _create_settings_menu(self):
        """Create the Settings menu"""
        if not self._menu_bar:
            return
            
        self._menus['settings'] = self._menu_bar.addMenu(self.tr("Settings"))
        
        # Preferences action
        self._actions['preferences'] = self._menus['settings'].addAction(self.tr("Preferences"))
        self._actions['preferences'].triggered.connect(self._on_preferences)
    
    def _on_exit(self):
        """Handle exit action"""
        if self.parent_window:
            self.parent_window.close()
        self.exit_requested.emit()
    
    def _on_preferences(self):
        """Handle preferences action"""
        if self.parent_window:
            if not hasattr(self.parent_window, '_prefs_window') or not self.parent_window._prefs_window:
                self.parent_window._prefs_window = PreferencesWindow(parent=self.parent_window)
            self.parent_window._prefs_window.show()
        self.preferences_requested.emit()
    
    def retranslate_ui(self):
        """Retranslate all menu items"""
        if not self._menu_bar:
            return
            
        # Update menu titles
        if 'file' in self._menus:
            self._menus['file'].setTitle(self.tr("File"))
        if 'settings' in self._menus:
            self._menus['settings'].setTitle(self.tr("Settings"))
        
        # Update action texts
        if 'exit' in self._actions:
            self._actions['exit'].setText(self.tr("Exit"))
        if 'preferences' in self._actions:
            self._actions['preferences'].setText(self.tr("Preferences"))
    
    def tr(self, text):
        """Translation method - uses parent window's tr if available"""
        if self.parent_window and hasattr(self.parent_window, 'tr'):
            return self.parent_window.tr(text)
        return text
    
    # Property accessors for easy access to menus and actions
    @property
    def file_menu(self):
        """Get the File menu"""
        return self._menus.get('file')
    
    @property
    def settings_menu(self):
        """Get the Settings menu"""
        return self._menus.get('settings')
    
    @property
    def exit_action(self):
        """Get the Exit action"""
        return self._actions.get('exit')
    
    @property
    def preferences_action(self):
        """Get the Preferences action"""
        return self._actions.get('preferences')
