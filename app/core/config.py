import json
import os
from pathlib import Path
from typing import Dict, Any


class Config:
    """Configuration manager for the application"""
    
    def __init__(self, config_file: str = None):
        if config_file is None:
            # Use appropriate config directory based on OS
            self.config_file = self._get_config_path()
        else:
            self.config_file = Path(config_file)
        
        # Ensure config directory exists
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        
        self._config = self._load_config()
    
    def _get_config_path(self) -> Path:
        """Get appropriate config file path based on OS"""
        app_name = "RinominaFacile2"
        
        if os.name == 'nt':  # Windows
            # Use %APPDATA%\RinominaFacile2\config.json
            config_dir = Path(os.environ.get('APPDATA', Path.home() / 'AppData' / 'Roaming'))
        elif os.name == 'posix':
            if os.uname().sysname == 'Darwin':  # macOS
                # Use ~/Library/Application Support/RinominaFacile2/config.json
                config_dir = Path.home() / 'Library' / 'Application Support'
            else:  # Linux and other Unix-like
                # Use ~/.config/RinominaFacile2/config.json (XDG Base Directory)
                config_dir = Path(os.environ.get('XDG_CONFIG_HOME', Path.home() / '.config'))
        else:
            # Fallback to home directory
            config_dir = Path.home()
            
        return config_dir / app_name / 'config.json'
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                # If config file is corrupted, return default config
                return self._get_default_config()
        else:
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration"""
        return {
            "language": "en",  # Default language
            "ai": {
                "gemini_api_key": ""  # Encrypted & base64-encoded when set
            }
        }
    
    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2, ensure_ascii=False)
        except IOError as e:
            print(f"Error saving config: {e}")
    
    def get(self, key: str, default=None):
        """Get configuration value"""
        keys = key.split('.')
        value = self._config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any):
        """Set configuration value"""
        keys = key.split('.')
        config = self._config
        
        # Navigate to the parent of the target key
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        # Set the value
        config[keys[-1]] = value
        self.save_config()

    # -------------------- Gemini API Key helpers --------------------
    def set_gemini_api_key_plain(self, api_key: str) -> tuple[bool, str]:
        """Encrypt and save the Gemini API key. Returns (success, error_message)."""
        try:
            if not isinstance(api_key, str):
                return (False, "Invalid API key type")
            from base64 import b64encode
            from core.secure_storage import encrypt_bytes
            token = encrypt_bytes(api_key.encode('utf-8'))
            enc_b64 = b64encode(token).decode('ascii')
            self.set('ai.gemini_api_key', enc_b64)
            return (True, "")
        except ImportError as e:
            return (False, f"Missing dependency: {e}")
        except Exception as e:
            return (False, f"Encryption error: {e}")

    def get_gemini_api_key_plain(self) -> str:
        """Decrypt and return the stored Gemini API key, or empty string if unavailable."""
        from base64 import b64decode
        from core.secure_storage import decrypt_bytes
        enc_b64 = self.get('ai.gemini_api_key', '')
        if not enc_b64:
            return ""
        try:
            token = b64decode(enc_b64)
            data = decrypt_bytes(token)
            if data is None:
                return ""
            return data.decode('utf-8')
        except Exception:
            return ""
    
    @property
    def language(self) -> str:
        """Get current language"""
        return self.get("language", "en")
    
    @language.setter
    def language(self, value: str):
        """Set current language"""
        self.set("language", value)
    
    def get_config_location(self) -> str:
        """Get the full path of the configuration file"""
        return str(self.config_file)


# Global config instance
config = Config()
