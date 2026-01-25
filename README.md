# rinominaFacile2
A software that helps you rename file recognizing them with AI for a first scan.

## Configuration
The application saves user preferences in a `config.json` file in the appropriate system directory:

### Configuration File Locations:
- **Windows**: `%APPDATA%\RinominaFacile2\config.json`
  - Example: `C:\Users\YourName\AppData\Roaming\RinominaFacile2\config.json`
- **macOS**: `~/Library/Application Support/RinominaFacile2/config.json`
- **Linux**: `~/.config/RinominaFacile2/config.json`


## Test Usage:

### On Windows:
```powershell
# Enable script execution (run this once)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Activate virtual environment
venv\Scripts\activate

# Launch the application
python app/main.py
```

### On Linux/macOS:
```bash
# Activate virtual environment
source venv/bin/activate

# Launch the application  
python app/main.py
```

## Translation Files:

### Update translation strings:
```powershell
# Extract translatable strings from source code
pyside6-lupdate app/ui/main_window.py app/ui/preferences_window.py -ts app/translations/it.ts

# Compile translation files
pyside6-lrelease app/translations/it.ts
```

## Building Executable

The application is designed to work correctly as a standalone executable:
- Configuration files are stored in appropriate system directories
- No need to bundle config files with the executable
- User preferences persist across updates

### Build with PyInstaller (example):
```bash
pyinstaller --onefile --windowed app/main.py --name "RinominaFacile2" --collect-all ui --collect-all core --collect-all ai --collect-all translations
```

or with a correct .spec file setting

Windows
```bash
pyhton generate_spec.py
pyinstaller RinominaFacile2.spec
```
