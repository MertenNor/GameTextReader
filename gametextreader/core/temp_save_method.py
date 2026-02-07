import os
import json
from gametextreader.constants import APP_SETTINGS_PATH

def save_tesseract_settings(self):
    """Save Tesseract-related settings to the global settings file."""
    try:
        temp_path = APP_SETTINGS_PATH
        
        # Load existing settings
        settings = {}
        if os.path.exists(temp_path):
            try:
                with open(temp_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            except:
                pass
        
        # Update Tesseract language
        settings['tesseract_language'] = self.tesseract_language_var.get()
        
        # Save back
        with open(temp_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=4)
            
    except Exception as e:
        print(f"Error saving Tesseract settings: {e}")
