"""
Constants and configuration for GameTextReader
"""
import os

# App information
APP_NAME = "GameTextReader"
APP_VERSION = "0.9.3.1"
APP_SLUG = APP_NAME.lower().replace(" ", "")

# Paths
APP_DOCUMENTS_DIR = os.path.join(os.path.expanduser('~'), 'Documents', APP_NAME)
APP_SETTINGS_FILENAME = f"{APP_SLUG}_settings.json"
APP_SETTINGS_PATH = os.path.join(APP_DOCUMENTS_DIR, APP_SETTINGS_FILENAME)
APP_SETTINGS_BACKUP_FILENAME = f".{APP_SETTINGS_FILENAME}.backup"
APP_SETTINGS_BACKUP_FILENAME = f".{APP_SETTINGS_FILENAME}.backup"
APP_AUTO_READ_SETTINGS_PATH = os.path.join(APP_DOCUMENTS_DIR, 'auto_read_settings.json')
APP_LAYOUTS_DIR = os.path.join(APP_DOCUMENTS_DIR, 'Layouts')

# GitHub repository configuration
GITHUB_REPO = "MertenNor/GameTextReader"  # Format: "username/repository-name"

# Update server configuration (Google Apps Script)
UPDATE_SERVER_URL = ""

# Testing: Set to True to always show update popup (for testing UI). Set to False for release (only shows when update is actually available)
SHOW_UPDATE_POPUP_FOR_TESTING = False

