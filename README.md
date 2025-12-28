# rinominaFacile2
A software that helps you rename file recognizing them with AI for a first scan.

## Configuration
The application saves user preferences in a `config.json` file in the appropriate system directory:

### Configuration File Locations:
- **Windows**: `%APPDATA%\RinominaFacile2\config.json`
  - Example: `C:\Users\YourName\AppData\Roaming\RinominaFacile2\config.json`
- **macOS**: `~/Library/Application Support/RinominaFacile2/config.json`
- **Linux**: `~/.config/RinominaFacile2/config.json`

### Saved Settings:
- Language selection

## Features

### Document Selection:
- **Select Documents Button**: Browse and select multiple documents at once
- **Drag & Drop Support**: Drag documents directly onto the application window
- **Document List Display**: View all selected documents in an organized list
- **Clear Documents**: Remove all selected documents with one click
- **Visual Feedback**: Drop zone changes color when dragging documents over it
- **Smart Filtering**: Only accepts supported document and image formats

### Document Preview:
- **Real-time Preview**: Click any document in the list to preview it instantly
- **Multi-format Support**: Preview text files, images, and document metadata
- **Smart Layout**: Resizable split-pane interface with document list on left, preview on right
- **File Information**: Shows file name, type, and size in the preview header
- **Scrollable Content**: Large documents are scrollable within the preview pane

### Supported Document Types:
- **Text Documents**: PDF, DOC, DOCX, TXT, RTF, ODT
- **Images**: PNG, JPG, JPEG, GIF, BMP, TIFF (for scanned documents)
- **Auto-validation**: Automatically filters out unsupported file types like audio/video

### Document Operations:
- Multiple document selection
- Document type filtering and validation
- Real-time document count display
- Cross-platform file path handling

## Usage:

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
pip install pyinstaller
pyinstaller --onefile --windowed app/main.py --name "RinominaFacile2"
```
