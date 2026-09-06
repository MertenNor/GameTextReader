"""
Constants and configuration for GameTextReader
"""
import os

# App information
APP_NAME = "GameTextReader"
APP_VERSION = "0.9.5.1"
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
APP_AI_VOICES_DIR     = os.path.join(APP_DOCUMENTS_DIR, 'ai_voices')
APP_PIPER_DIR         = os.path.join(APP_AI_VOICES_DIR, 'piper')
APP_PIPER_BIN_DIR     = os.path.join(APP_PIPER_DIR, 'piper_bin')
APP_PIPER_PRESETS_DIR = os.path.join(APP_PIPER_DIR, 'piper_presets')
APP_PIPER_VOICES_DIR  = os.path.join(APP_PIPER_DIR, 'piper_voices')
APP_KOKORO_DIR        = os.path.join(APP_AI_VOICES_DIR, 'kokoro')
APP_KOKORO_CUSTOM_DIR = os.path.join(APP_KOKORO_DIR, 'custom')
APP_SAPI_PRESETS_DIR  = os.path.join(APP_DOCUMENTS_DIR, 'sapi_presets')

PRESS_ANY_KEY = "Press any key..."

# GitHub repository configuration
GITHUB_REPO = "MertenNor/GameTextReader"  # Format: "username/repository-name"

# Update server configuration (Google Apps Script)
UPDATE_SERVER_URL = ""

# Testing: Set to True to always show update popup (for testing UI). Set to False for release (only shows when update is actually available)
SHOW_UPDATE_POPUP_FOR_TESTING = False

# Debug: force get_dpi_scale() to return this value instead of the real Windows
# display scale, so DPI/high-scaling behavior (window sizing, checkbox/font
# rendering, etc.) can be tested without changing actual display settings.
# Set to a number (e.g. 3.0 for 300%) to test; set to False to use the real
# scale. Only takes effect when running from source - ignored in a compiled
# build even if accidentally left non-False, so this can't ship enabled.
DEBUG_FORCE_DPI_SCALE = False

