import os
import tkinter as tk
from PIL import ImageTk
import sys


class ToolTip:
    """Create a tooltip for a given widget"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<Motion>", self.on_motion)

    def on_enter(self, event=None):
        self.show_tooltip()

    def on_leave(self, event=None):
        self.hide_tooltip()

    def on_motion(self, event=None):
        if self.tooltip_window:
            self.update_position(event)

    def show_tooltip(self):
        if self.tooltip_window:
            return

        x, y, _, _ = self.widget.bbox("all") if hasattr(self.widget, 'bbox') else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            background="#ffffe0",
            relief='solid',
            borderwidth=1,
            font=("Helvetica", 9),
            justify='left',
            wraplength=300
        )
        label.pack()

    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def update_position(self, event):
        if self.tooltip_window:
            x = self.widget.winfo_rootx() + 25
            y = self.widget.winfo_rooty() + 20
            self.tooltip_window.wm_geometry(f"+{x}+{y}")


def add_info_icon(parent, text, **pack_opts):
    """Pack a small Windows-style info icon (blue circle with white 'i') into
    `parent` that shows `text` in a tooltip on hover. Returns the icon widget."""
    info_canvas = tk.Canvas(parent, width=16, height=16, highlightthickness=0, cursor="hand2")
    info_canvas.create_oval(2, 2, 14, 14, fill='#0078D4', outline='#005A9E', width=1)
    info_canvas.create_text(8, 8, text="i", font=("Helvetica", 9, "bold"), fill="white")
    pack_kwargs = {'side': 'left', 'padx': (5, 0)}
    pack_kwargs.update(pack_opts)
    info_canvas.pack(**pack_kwargs)
    ToolTip(info_canvas, text)
    return info_canvas


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