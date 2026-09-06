
import os
import threading
import json
import argostranslate.package
import argostranslate.translate
from ..constants import APP_DOCUMENTS_DIR
import argostranslate.settings

class TranslationManager:
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(TranslationManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if self._initialized:
            return
            
        self._initialized = True
        
        # Set custom package directory
        # User requested: ~/.local/share/argos-translate
        user_home = os.path.expanduser("~")
        self.package_dir = os.path.join(user_home, ".local", "share", "argos-translate")
        
        # Ensure the directory exists
        if not os.path.exists(self.package_dir):
            try:
                os.makedirs(self.package_dir)
            except OSError as e:
                print(f"[ERROR] Translation Manager: Directory creation failed: {e}")
        
        # Tell argostranslate to use this directory
        os.environ["ARGOS_PACKAGES_DIR"] = self.package_dir
        
        # Update index is needed to know what's available
        self.packages_updated = False
        
    def update_package_index(self):
        """Downloads the latest package index from the official repository."""
        try:
            argostranslate.package.update_package_index()
            self.packages_updated = True
            return True
        except Exception as e:
            print(f"[ERROR] Translation Manager: Index update failed: {e}")
            return False

    def get_installed_languages(self):
        """Returns a list of installed languages."""
        try:
            return argostranslate.package.get_installed_packages()
        except Exception as e:
            print(f"[ERROR] Translation Manager: Lookup failed: {e}")
            return []

    def get_available_packages(self):
        """Returns a list of available packages for download (from local index)."""
        try:
            return argostranslate.package.get_available_packages()
        except Exception as e:
            print(f"[ERROR] Translation Manager: Index lookup failed: {e}")
            return []
            
    def install_package(self, package, progress_callback=None):
        """
        Installs a specific package with optional progress reporting.
        """
        try:
            download_path = None
            
            # If progress callback provided and package has links, try manual download for progress bar
            if progress_callback and hasattr(package, 'links') and package.links:
                try:
                    import requests
                    import tempfile
                    
                    url = package.links[0]
                    # Create unique temp file
                    fd, download_path = tempfile.mkstemp(suffix=".argosmodel")
                    os.close(fd)
                    
                    # print(f"Downloading Translation: {package.from_name} -> {package.to_name}: 0%")
                    
                    response = requests.get(url, stream=True)
                    total_length = response.headers.get('content-length')

                    if total_length is None: # no content length header
                        # Fallback to standard
                        pass
                    else:
                        dl = 0
                        total_length = int(total_length)
                        with open(download_path, 'wb') as f:
                            for data in response.iter_content(chunk_size=4096):
                                dl += len(data)
                                f.write(data)
                                percentage = int(100 * dl / total_length)
                                # Only update occasionally to avoid spamming GUI queue
                                if dl % (4096 * 10) == 0 or percentage == 100:
                                    # print(f"Downloading Translation: {package.from_name} -> {package.to_name}: {percentage}%")
                                    progress_callback(percentage)
                        
                        # Ensure we hit 100 on valid download
                        print("Download complete.")
                        progress_callback(100)
                except Exception as e:
                    print("[WARNING] Translation Manager: Manual download failed, falling back to built-in.")
                    if download_path and os.path.exists(download_path):
                        try:
                            os.remove(download_path)
                        except:
                            pass
                    download_path = None

            # Fallback to default download if manual failed or wasn't attempted
            if not download_path:
                download_path = package.download()
                
            argostranslate.package.install_from_path(download_path)
            
            # Cleanup manual download if we created one
            if progress_callback and download_path and os.path.exists(download_path):
                 # Note: install_from_path usually unzips/copies, so we can likely delete the source zip.
                 # However, built-in download() usually puts it in cache. 
                 # Our temp file is truly temp.
                 try:
                     os.remove(download_path)
                 except:
                     pass
                     
            return True
        except Exception as e:
            print(f"[ERROR] Translation Manager: Install failed ({package}): {e}")
            return False

    def uninstall_package(self, package):
        """Uninstalls a package."""
        try:
            argostranslate.package.uninstall(package)
            return True
        except Exception as e:
            print(f"[ERROR] Translation Manager: Uninstall failed: {e}")
            return False

    def translate(self, text, from_code, to_code):
        """
        Translates text from source language to target language.
        """
        if not text:
            return ""
            
        try:
            # FIX: Use translate.get_installed_languages to get Language objects for translation
            # The self.get_installed_languages() returns Package objects which don't have .code attribute or translation methods
            installed_langs = argostranslate.translate.get_installed_languages()
            
            from_lang = next((x for x in installed_langs if x.code == from_code), None)
            to_lang = next((x for x in installed_langs if x.code == to_code), None)

            if from_lang and to_lang:
                translation = from_lang.get_translation(to_lang)
                if translation:
                    return translation.translate(text)
                else:
                    # Try generic translate if specific pair not found directly
                    return argostranslate.translate.translate(text, from_code, to_code)
            else:
                 return argostranslate.translate.translate(text, from_code, to_code)
        except Exception as e:
            print(f"[ERROR] Translation: {e}")
            return text

    def apply_settings(self, settings):
        """Apply performance and other settings to argostranslate"""
        try:
            device = settings.get("device", "cpu")
            if device:
                argostranslate.settings.device = device
                
            compute_type = settings.get("compute_type", "float32")
            if compute_type:
                argostranslate.settings.compute_type = compute_type
            
            beam_size = settings.get("beam_size")
            if beam_size is not None:
                argostranslate.settings.beam_size = int(beam_size)
            
            debug = settings.get("debug")
            if debug is not None:
                argostranslate.settings.debug = bool(debug)
                
            print("[NOTE] Translation Manager: Settings applied.")
            
            # Reload internal CTranslate2 translator/generator if needed? 
            # Usually Argos loads models on demand, so changing settings *before* first translation should work.
            # If models are already loaded, this might not affect them unless we clear cache.
            # But for now this is the best we can do without deep diving into Argos internals.
            
        except Exception as e:
            print(f"[ERROR] Translation Manager: Settings application failed: {e}")
