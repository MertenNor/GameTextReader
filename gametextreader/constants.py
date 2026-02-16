"""
Constants and configuration for GameTextReader
"""
import os

# App information
APP_NAME = "GameTextReader"
APP_VERSION = "0.9.4.1"
APP_SLUG = APP_NAME.lower().replace(" ", "")

# Determine the Documents directory path reliably on Windows (respecting redirection)
def get_documents_dir():
    if os.name == 'nt':
        try:
            import ctypes
            from ctypes import wintypes
            CSIDL_PERSONAL = 5       # My Documents
            SHGFP_TYPE_CURRENT = 0   # Current, not default value
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
            return buf.value
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'Documents')
    return os.path.join(os.path.expanduser('~'), 'Documents')

APP_DOCUMENTS_DIR = os.path.join(get_documents_dir(), APP_NAME)
APP_SETTINGS_FILENAME = f"{APP_SLUG}_settings.json"
APP_SETTINGS_PATH = os.path.join(APP_DOCUMENTS_DIR, APP_SETTINGS_FILENAME)
APP_SETTINGS_BACKUP_FILENAME = f".{APP_SETTINGS_FILENAME}.backup"
APP_AUTO_READ_SETTINGS_PATH = os.path.join(APP_DOCUMENTS_DIR, 'auto_read_settings.json')
APP_LAYOUTS_DIR = os.path.join(APP_DOCUMENTS_DIR, 'Layouts')

# GitHub repository configuration
GITHUB_REPO = "MertenNor/GameTextReader"  # Format: "username/repository-name"

# Update server configuration (Google Apps Script)
UPDATE_SERVER_URL = ""

# Testing: Set to True to always show update popup (for testing UI). Set to False for release (only shows when update is actually available)
SHOW_UPDATE_POPUP_FOR_TESTING = False

