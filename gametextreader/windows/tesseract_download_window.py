
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import webbrowser

class TesseractDownloadWindow:
    def __init__(self, parent, tesseract_manager):
        self.parent = parent
        self.tesseract_manager = tesseract_manager
        self.window = tk.Toplevel(parent)
        self.window.title("Tesseract Language Manager")
        self.window.geometry("700x500")
        self.window.resizable(False, False)
        
        # Center the window
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            pass
        
        # Disable hotkeys in main window while this is open? 
        # Actually proper way is to register it, but we can just make it modal-ish or rely on focus.
        # Ideally we registers it to the main app if possible, but we don't have easy access to the main app register method here 
        # unless we pass it or make it global. 
        # For now, let's just keep it simple.
        
        self.setup_ui()
        self.refresh_installed_languages()
        
        # Start loading available languages
        self.status_label.config(text="Loading available languages...", fg="blue")
        threading.Thread(target=self.load_available_languages, daemon=True).start()

    def setup_ui(self):
        # Main container
        main_frame = tk.Frame(self.window, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_label = tk.Label(main_frame, text="Manage Tesseract OCR Languages", font=("Helvetica", 12, "bold"))
        header_label.pack(pady=(0, 5))
        
        # Link to GitHub
        link_frame = tk.Frame(main_frame)
        link_frame.pack(pady=(0, 10))
        
        tk.Label(link_frame, text="The list is sourced from Tesseract GitHub:", font=("Helvetica", 9)).pack(side=tk.LEFT)
        
        link = tk.Label(link_frame, text="https://github.com/tesseract-ocr/tessdata_best", font=("Helvetica", 9, "underline"), fg="blue", cursor="hand2")
        link.pack(side=tk.LEFT, padx=5)
        link.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/tesseract-ocr/tessdata_best"))
        
        tk.Label(main_frame, text="Download additional languages for better text recognition", font=("Helvetica", 9)).pack(pady=(0, 10))

        # Content area (Two lists)
        content_frame = tk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Left side: Available Languages
        left_frame = tk.LabelFrame(content_frame, text="Available Online", padx=5, pady=5)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.available_listbox = tk.Listbox(left_frame, selectmode=tk.SINGLE, font=("Helvetica", 10))
        self.available_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        available_scrollbar = tk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.available_listbox.yview)
        available_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.available_listbox.config(yscrollcommand=available_scrollbar.set)
        
        self.available_listbox.bind('<<ListboxSelect>>', self.on_available_select)

        # Buttons in the middle
        btn_frame = tk.Frame(content_frame)
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)

        self.install_btn = tk.Button(btn_frame, text="Download >>", command=self.install_selected, state=tk.DISABLED, width=12)
        self.install_btn.pack(side=tk.TOP, pady=20)

        self.delete_btn = tk.Button(btn_frame, text="Delete", command=self.delete_selected, state=tk.DISABLED, width=12)
        self.delete_btn.pack(side=tk.TOP, pady=20)

        # Right side: Installed Languages
        right_frame = tk.LabelFrame(content_frame, text="Installed Locally", padx=5, pady=5)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.installed_listbox = tk.Listbox(right_frame, selectmode=tk.SINGLE, font=("Helvetica", 10))
        self.installed_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        installed_scrollbar = tk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.installed_listbox.yview)
        installed_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.installed_listbox.config(yscrollcommand=installed_scrollbar.set)
        
        self.installed_listbox.bind('<<ListboxSelect>>', self.on_installed_select)
        
        # Refresh Button (Placed under the lists, above the status bar)
        refresh_frame = tk.Frame(main_frame)
        refresh_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Using a frame to align it under the "Available Online" list (left side)
        refresh_align_frame = tk.Frame(refresh_frame)
        refresh_align_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Make the button fill the width of the left column roughly
        self.refresh_btn = tk.Button(refresh_align_frame, text="Refresh List", command=self.start_refresh, width=12)
        self.refresh_btn.pack(side=tk.LEFT)
        
        # Spacer for right side
        tk.Frame(refresh_frame).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # Bottom area: Status and Close
        bottom_frame = tk.Frame(main_frame, pady=10)
        bottom_frame.pack(fill=tk.X)

        # Right side buttons
        close_btn = tk.Button(bottom_frame, text="Close", command=self.window.destroy, width=10)
        close_btn.pack(side=tk.RIGHT, padx=5)

        open_folder_btn = tk.Button(bottom_frame, text="Open Tessdata Folder", command=self.open_local_folder, width=20)
        open_folder_btn.pack(side=tk.RIGHT)

        # Progress bar and status (Left side)
        self.progress = ttk.Progressbar(bottom_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.status_label = tk.Label(bottom_frame, text="Ready", font=("Helvetica", 9), anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Dictionaries to map listbox indices to code/name
        self.available_map = {}
        self.installed_map = {} # index -> code

    def start_refresh(self):
        """Manually refresh the available languages list."""
        self.available_listbox.delete(0, tk.END)
        self.available_map.clear()
        self.status_label.config(text="Refreshing available languages...", fg="blue")
        self.progress.start(10)
        threading.Thread(target=self.load_available_languages, daemon=True).start()

    def load_available_languages(self):
        try:
            # We don't have a remote index to fetch for Tesseract like Argos does,
            # but we can force refresh the manager's installed cache and reload the UI.
            self.tesseract_manager.get_installed_languages(force_refresh=True)
            langs = self.tesseract_manager.get_available_languages()
            self.window.after(0, self.update_available_list, langs)
        except Exception as e:
            self.window.after(0, self._on_load_error, str(e))

    def _on_load_error(self, error_msg):
        self.progress.stop()
        self.status_label.config(text=f"Error: {error_msg}", fg="red")

    def update_available_list(self, langs):
        self.progress.stop()
        self.available_listbox.delete(0, tk.END)
        self.available_map.clear()
        
        # Filter out already installed ones
        installed_codes = self.tesseract_manager.get_installed_languages().keys()
        
        # Convert dict to list of tuples for sorting
        lang_list = []
        for code, name in langs.items():
            if code not in installed_codes:
                lang_list.append((code, name))
        
        # Sort alphanumerically by display name
        lang_list.sort(key=lambda x: x[1])
        
        count = 0
        for code, name in lang_list:
             display_text = name
             self.available_listbox.insert(tk.END, display_text)
             self.available_map[count] = code
             count += 1
        
        self.status_label.config(text="Review available languages.", fg="green")

    def refresh_installed_languages(self):
        self.installed_listbox.delete(0, tk.END)
        self.installed_map.clear()
        
        installed = self.tesseract_manager.get_installed_languages()
        
        # Convert to list for sorting
        lang_list = []
        for code, name in installed.items():
            lang_list.append((code, name))
            
        lang_list.sort(key=lambda x: x[1])
        
        count = 0
        for code, name in lang_list:
            display_text = n = name # name includes code from TesseractManager.get_installed_languages
            
            # Skip showing OSD in the list since it's internal
            if code == 'osd':
                continue
            
            # Highlight if it's default
            if "(Default)" in n:
                self.installed_listbox.insert(tk.END, n)
                self.installed_listbox.itemconfig(tk.END, {'bg':'#e6f3ff'})
            else:
                self.installed_listbox.insert(tk.END, n)
                
            self.installed_map[count] = code
            count += 1
        
        self.delete_btn.config(state=tk.DISABLED)

    def on_available_select(self, event):
        selection = self.available_listbox.curselection()
        if selection:
            self.install_btn.config(state=tk.NORMAL)
        else:
            self.install_btn.config(state=tk.DISABLED)

    def install_selected(self):
        selection = self.available_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        code = self.available_map.get(index)
        
        if code:
            self.install_btn.config(state=tk.DISABLED)
            self.status_label.config(text=f"Downloading {code}...", fg="blue")
            self.progress['mode'] = 'indeterminate'
            self.progress.start(10)
            
            threading.Thread(target=self.run_install, args=(code,), daemon=True).start()

    def run_install(self, code):
        success, error_msg = self.tesseract_manager.download_language(code)
        self.window.after(0, self.on_install_complete, success, error_msg, code)

    def on_install_complete(self, success, error_msg, code):
        self.progress.stop()
        self.install_btn.config(state=tk.DISABLED)
        
        if success:
            self.status_label.config(text=f"Successfully installed {code}.", fg="green")
            self.refresh_installed_languages()
            # Reload available to remove the installed one
            threading.Thread(target=self.load_available_languages, daemon=True).start()
        else:
            self.status_label.config(text="Download failed.", fg="red")
            msg = f"Failed to download language: {code}"
            if error_msg:
                msg += f"\n\n{error_msg}"
            messagebox.showerror("Error", msg)

    def on_installed_select(self, event):
        selection = self.installed_listbox.curselection()
        if selection:
             # Check if it is a deletable custom language (not default system ones)
             index = selection[0]
             code = self.installed_map.get(index)
             
             # Protected default languages
             protected_langs = ['eng', 'osd']
             
             # Logic: Must be custom AND not in protected list
             if self.tesseract_manager.is_custom_language(code) and code not in protected_langs:
                 self.delete_btn.config(state=tk.NORMAL)
                 # Reset status if it was showing a warning
                 current_status = self.status_label.cget("text")
                 if "Cannot delete" in current_status:
                     self.status_label.config(text="Ready", fg="black")
             else:
                 self.delete_btn.config(state=tk.DISABLED)
                 if code in protected_langs:
                     self.status_label.config(text=f"Cannot delete default language ({code}).", fg="orange")
                 else:
                     self.status_label.config(text="Cannot delete default system languages.", fg="orange")
        else:
            self.delete_btn.config(state=tk.DISABLED)

    def delete_selected(self):
        selection = self.installed_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        code = self.installed_map.get(index)
        
        # Safeguard against deleting defaults
        if code in ['eng', 'osd']:
            messagebox.showwarning("Warning", f"Cannot delete default language: {code}")
            return
        
        if code:
            # Get full name using the manager's map
            full_name = self.tesseract_manager.get_available_languages().get(code, code)
            confirm_text = f"{full_name} ({code})"
            
            confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the language data for:\n{confirm_text}?")
            if confirm:
                self.delete_btn.config(state=tk.DISABLED)
                self.status_label.config(text=f"Deleting {confirm_text}...", fg="blue")
                
                threading.Thread(target=self.run_uninstall, args=(code,), daemon=True).start()

    def run_uninstall(self, code):
        success, error_msg = self.tesseract_manager.delete_language(code)
        self.window.after(0, self.on_uninstall_complete, success, error_msg, code)

    def on_uninstall_complete(self, success, error_msg, code):
        if success:
            self.status_label.config(text=f"Successfully deleted {code}", fg="green")
            self.refresh_installed_languages()
            threading.Thread(target=self.load_available_languages, daemon=True).start()
        else:
            self.status_label.config(text="Deletion failed.", fg="red")
            msg = "Failed to delete language file."
            if error_msg:
                msg += f"\n\n{error_msg}"
            messagebox.showerror("Error", msg)

    def open_local_folder(self):
        """Open the local tessdata folder."""
        folder_path = self.tesseract_manager.custom_tessdata_dir
        if not os.path.exists(folder_path):
             # Try creating it or fallback to checking usage
             os.makedirs(folder_path, exist_ok=True)
             
        try:
            os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {str(e)}")
