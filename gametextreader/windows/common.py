import os
from PIL import ImageTk
import sys

def set_window_icon(root, icon_path=None, register=False):
    # Set the window icon
    if icon_path is None:
        icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
    try:
        if os.path.exists(icon_path):
            # root.iconbitmap(icon_path) only works on Windows, so use PIL for cross-platform support
            icon_photo = ImageTk.PhotoImage(file=icon_path)
            root.iconphoto(True, icon_photo)

            # Register the AUMID in the registry so Windows shows the correct
            # display name and icon in the taskbar jump list instead of "Python".
            if sys.platform.startswith('win') and register:
                try:
                    import winreg
                    key_path = r"Software\Classes\AppUserModelId\GameTextReader.App"
                    key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path,
                                             0, winreg.KEY_SET_VALUE)
                    winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "GameTextReader")
                    winreg.SetValueEx(key, "IconUri",     0, winreg.REG_SZ, icon_path)
                    winreg.CloseKey(key)
                except Exception:
                    pass
        else:
            print(f"Icon file not found at: {icon_path}")
        print(f"Set window icon to: {icon_path}")
    except Exception as e:
        print(f"Error setting window icon: {e}")