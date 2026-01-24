import os
from pathlib import Path
from typing import Optional


def _get_app_config_dir() -> Path:
    app_name = "RinominaFacile2"
    if os.name == 'nt':
        base = Path(os.environ.get('APPDATA', Path.home() / 'AppData' / 'Roaming'))
    elif os.name == 'posix':
        try:
            import platform
            if platform.system() == 'Darwin':
                base = Path.home() / 'Library' / 'Application Support'
            else:
                base = Path(os.environ.get('XDG_CONFIG_HOME', Path.home() / '.config'))
        except Exception:
            base = Path.home() / '.config'
    else:
        base = Path.home()
    cfg_dir = base / app_name
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir


# --- Windows DPAPI (preferred on Windows) ---

def _windows_encrypt(data: bytes) -> bytes:
    import win32crypt  # type: ignore
    # CryptProtectData returns encrypted bytes directly
    return win32crypt.CryptProtectData(data, "GeminiAPIKey")


def _windows_decrypt(blob: bytes) -> bytes:
    import win32crypt  # type: ignore
    # CryptUnprotectData returns (description, decrypted_data)
    return win32crypt.CryptUnprotectData(blob, None)[1]


# --- Cross-platform fallback using Fernet ---

def _get_fernet():
    from cryptography.fernet import Fernet
    key_path = _get_app_config_dir() / 'secret.key'
    if not key_path.exists():
        key = Fernet.generate_key()
        try:
            key_path.write_bytes(key)
        except Exception:
            # Fallback to creating parent then writing
            key_path.parent.mkdir(parents=True, exist_ok=True)
            key_path.write_bytes(key)
    else:
        key = key_path.read_bytes()
    return Fernet(key)


def _fernet_encrypt(data: bytes) -> bytes:
    f = _get_fernet()
    return f.encrypt(data)


def _fernet_decrypt(token: bytes) -> bytes:
    f = _get_fernet()
    return f.decrypt(token)


def encrypt_bytes(data: bytes) -> bytes:
    """Encrypt bytes using OS-specific mechanism.
    - Windows: DPAPI (user-bound)
    - Others: Fernet with local key file
    """
    if os.name == 'nt':
        try:
            return _windows_encrypt(data)
        except Exception:
            # Fallback to Fernet
            return _fernet_encrypt(data)
    else:
        return _fernet_encrypt(data)


def decrypt_bytes(token: bytes) -> Optional[bytes]:
    """Decrypt bytes; returns None if decryption fails."""
    try:
        if os.name == 'nt':
            try:
                return _windows_decrypt(token)
            except Exception:
                # Fallback to Fernet
                return _fernet_decrypt(token)
        else:
            return _fernet_decrypt(token)
    except Exception:
        return None
