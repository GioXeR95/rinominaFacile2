import sys
import os
from pathlib import Path

# Add the app directory to the Python path for development and PyInstaller
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    app_dir = Path(sys.executable).parent
else:
    # Running as script
    app_dir = Path(__file__).parent

sys.path.insert(0, str(app_dir))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTranslator, QLocale
from ui.main_window import MainWindow

app = QApplication(sys.argv)

# The MainWindow now handles loading saved language automatically
window = MainWindow()
window.show()

sys.exit(app.exec())