import os
import requests
import shutil
import threading
from pathlib import Path
import pytesseract

# URL for Tesseract trained data (best quality)
# For standard quality use: https://github.com/tesseract-ocr/tessdata
TESSDATA_REPO_URL = "https://github.com/tesseract-ocr/tessdata_best/raw/main"

# Mapping of common language codes to display names
LANGUAGE_MAP = {
    "eng": "English",
    "deu": "German",
    "fra": "French",
    "ita": "Italian",
    "spa": "Spanish",
    "por": "Portuguese",
    "rus": "Russian",
    "jpn": "Japanese",
    "chi_sim": "Chinese (Simplified)",
    "chi_tra": "Chinese (Traditional)",
    "kor": "Korean",
    "nld": "Dutch",
    "pol": "Polish",
    "tur": "Turkish",
    "ukr": "Ukrainian",
    "vie": "Vietnamese",
    "ara": "Arabic",
    "hin": "Hindi",
    "tha": "Thai",
    "swe": "Swedish",
    "dan": "Danish",
    "nor": "Norwegian",
    "fin": "Finnish",
    # Add more as needed
}

class TesseractManager:
    def __init__(self, documents_dir):
        """
        Initialize the TesseractManager.
        :param documents_dir: The application's documents directory (e.g., in My Documents).
        """
        self.custom_tessdata_dir = os.path.join(documents_dir, "tessdata")
        self._cached_installed_languages = None
        self._cached_system_tessdata_dir = None
        self.ensure_tessdata_dir()

    def ensure_tessdata_dir(self):
        """Ensure the custom tessdata directory exists."""
        if not os.path.exists(self.custom_tessdata_dir):
            try:
                os.makedirs(self.custom_tessdata_dir)
            except Exception as e:
                print(f"[ERROR] OCR Manager: Directory creation failed: {e}")
        
        # Ensure 'eng.traineddata' exists in the custom directory to prevent it from "disappearing"
        # when Tesseract is forced to use the custom directory
        self.ensure_english_available()

    def find_system_tessdata_dir(self):
        """Locate the tessdata folder of the detected Tesseract installation."""
        if self._cached_system_tessdata_dir is not None:
            return self._cached_system_tessdata_dir or None

        cmds = [
            pytesseract.pytesseract.tesseract_cmd,
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            os.path.join(os.getenv('LOCALAPPDATA', ''), 'Tesseract-OCR', 'tesseract.exe')
        ]

        system_tessdata = None
        for cmd in cmds:
            if cmd and os.path.exists(cmd) and os.path.isfile(cmd):
                parent = os.path.dirname(cmd)
                candidate = os.path.join(parent, 'tessdata')
                if os.path.exists(candidate) and os.path.isdir(candidate):
                    system_tessdata = candidate
                    break

        # Cache '' rather than None so a lookup failure isn't retried every call.
        self._cached_system_tessdata_dir = system_tessdata or ''
        return system_tessdata

    def ensure_english_available(self):
        """Copies eng.traineddata from system install to custom folder if missing."""
        eng_dest = os.path.join(self.custom_tessdata_dir, "eng.traineddata")
        if os.path.exists(eng_dest):
            return

        try:
            system_tessdata = self.find_system_tessdata_dir()
            if system_tessdata:
                eng_src = os.path.join(system_tessdata, "eng.traineddata")
                if os.path.exists(eng_src):
                    print(f"Copying system English language data to custom folder: {eng_dest}")
                    shutil.copy2(eng_src, eng_dest)
                    # Invalidate cache to ensure it appears as 'custom' (or just present)
                    self._cached_installed_languages = None
        except Exception as e:
            print(f"[ERROR] Failed to copy system English data: {e}")

    def get_tessdata_dir_param(self, lang_code):
        """
        Return the --tessdata-dir argument pointing at wherever lang_code's
        traineddata file actually lives, or None if it can't be resolved there.

        Languages downloaded through this app live in custom_tessdata_dir;
        languages installed via the Tesseract installer live in the system
        install's own tessdata folder. Picking the wrong one causes Tesseract
        to fail loading a language that the installer clearly provided.
        """
        if self.is_custom_language(lang_code):
            directory = self.custom_tessdata_dir
        else:
            system_dir = self.find_system_tessdata_dir()
            if system_dir and os.path.exists(os.path.join(system_dir, f"{lang_code}.traineddata")):
                directory = system_dir
            else:
                return None

        # Use forward slashes to handle cases correctly on Windows.
        # Only use quotes if the path contains spaces to avoid Tesseract misinterpreting quotes.
        path = os.path.normpath(directory).replace('\\', '/')
        if ' ' in path:
            return f'--tessdata-dir "{path}"'
        return f'--tessdata-dir {path}'

    def get_installed_languages(self, force_refresh=False):
        """
        Returns a dictionary of installed language codes and their display names.
        Combines system languages and custom downloaded ones.
        """
        if self._cached_installed_languages is not None and not force_refresh:
            return self._cached_installed_languages

        installed = {}
        
        # 1. Get system languages via pytesseract
        try:
            # Prevent environment variable from forcing tesseract to look elsewhere
            # while we specifically want the SYSTEM default languages.
            # Some environments or user configs might set this, causing system languages to 'disappear'.
            old_prefix = os.environ.pop('TESSDATA_PREFIX', None)
            
            # Prevent console window flash on Windows
            system_langs = pytesseract.get_languages(config='')
            
            # Restore environment if it was set
            if old_prefix:
                os.environ['TESSDATA_PREFIX'] = old_prefix
            
            for lang in system_langs:
                name = LANGUAGE_MAP.get(lang, lang)
                installed[lang] = name
        except Exception as e:
            # Pytesseract might fail if tesseract not found or other issues
            print(f"[ERROR] Failed to list system OCR languages: {e}")

        # 2. Get custom languages
        if os.path.exists(self.custom_tessdata_dir):
            for filename in os.listdir(self.custom_tessdata_dir):
                if filename.endswith(".traineddata"):
                    lang_code = filename.replace(".traineddata", "")
                    name = LANGUAGE_MAP.get(lang_code, lang_code)
                    installed[lang_code] = f"{name}"

        # Ensure we always have at least English if detection fails but "eng" is standard
        if not installed:
            installed["eng"] = "English"
        
        self._cached_installed_languages = installed
        return installed

    def get_available_languages(self):
        """Returns a dict of all downloadable languages supported by this manager."""
        return LANGUAGE_MAP

    def download_language(self, lang_code):
        """
        Downloads the traineddata file for the given language code.
        Blocking call. Returns True if successful, False otherwise.
        """
        # Ensure directory exists before downloading
        self.ensure_tessdata_dir()
        
        url = f"{TESSDATA_REPO_URL}/{lang_code}.traineddata"
        target_path = os.path.join(self.custom_tessdata_dir, f"{lang_code}.traineddata")
        
        try:
            # print(f"Downloading OCR: {LANGUAGE_MAP.get(lang_code, lang_code)}...")
            response = requests.get(url, stream=True)
            response.raise_for_status()
            
            with open(target_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            print("Download complete.")
            self._cached_installed_languages = None  # Invalidate cache
            return True, None
                
        except Exception as e:
            print(f"[ERROR] OCR Manager: Download failed: {e}")
            error_msg = str(e)
            if "Permission denied" in error_msg or isinstance(e, PermissionError):
                error_msg = "Permission Denied! Check folder permissions."
            if os.path.exists(target_path):
                try:
                    os.remove(target_path)
                except:
                    pass
            return False, error_msg

    def is_custom_language(self, lang_code):
        """Check if a language is installed in the custom folder (and thus deletable)."""
        return os.path.exists(os.path.join(self.custom_tessdata_dir, f"{lang_code}.traineddata"))

    def delete_language(self, lang_code):
        """Delete a custom language file."""
        if not self.is_custom_language(lang_code):
            return False, "Not a custom language"
            
        try:
            os.remove(os.path.join(self.custom_tessdata_dir, f"{lang_code}.traineddata"))
            self._cached_installed_languages = None  # Invalidate cache
            return True, None
        except Exception as e:
            print(f"[ERROR] OCR Manager: Delete failed: {e}")
            error_msg = str(e)
            if "Permission denied" in error_msg or isinstance(e, PermissionError):
                error_msg = "Permission Denied! Check folder permissions."
            return False, error_msg

    def is_language_installed(self, lang_code):
        """Check if a language ID is locally available."""
        # Use cache if available, but don't force refresh just for this check
        if self._cached_installed_languages:
            return lang_code in self._cached_installed_languages
            
        # Check custom dir
        if self.is_custom_language(lang_code):
            return True
        
        # Check system (approximate)
        try:
            if lang_code in pytesseract.get_languages(config=''):
                return True
        except:
            pass
            
        return False
