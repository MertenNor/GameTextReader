
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import webbrowser
from ..translation.translation_manager import TranslationManager
from ..window_geometry import apply_window_geometry

class TranslationWindow:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        self.window.title("Language Manager")
        self.window.resizable(False, False)
        apply_window_geometry(self.window, 'language_manager', 700, 500)

        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            pass

        self.translation_manager = TranslationManager()
        
        self.setup_ui()
        self.refresh_installed_languages()
        
        # Start loading available languages (from local cache first)
        self.status_label.config(text="Checking local package index...", fg="blue")
        threading.Thread(target=lambda: self.load_available_languages(update_index=False), daemon=True).start()
        self.window.deiconify()

    def setup_ui(self):
        # Main container
        main_frame = tk.Frame(self.window, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_label = tk.Label(main_frame, text="Manage Translation Packages", font=("Helvetica", 12, "bold"))
        header_label.pack(pady=(0, 5))
        
        # Link to GitHub
        link_frame = tk.Frame(main_frame)
        link_frame.pack(pady=(0, 10))
        
        tk.Label(link_frame, text="Powered by Argos Translate:", font=("Helvetica", 9)).pack(side=tk.LEFT)
        
        link = tk.Label(link_frame, text="https://github.com/argosopentech/argos-translate", font=("Helvetica", 9, "underline"), fg="blue", cursor="hand2")
        link.pack(side=tk.LEFT, padx=5)
        link.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/argosopentech/argos-translate"))

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

        self.install_btn = tk.Button(btn_frame, text="Install >>", command=self.install_selected, state=tk.DISABLED, width=12)
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
        # The lists are 50/50, so we can put it in a left aligned container
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

        open_folder_btn = tk.Button(bottom_frame, text="Open Local Install Folder", command=self.open_local_folder, width=20)
        open_folder_btn.pack(side=tk.RIGHT)




        
        # Progress bar and status (Left side)
        self.progress = ttk.Progressbar(bottom_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.status_label = tk.Label(bottom_frame, text="Ready", font=("Helvetica", 9), anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Dictionaries to map display names to objects
        self.available_map = {}
        self.installed_map = {}

    def start_refresh(self):
        self.available_listbox.delete(0, tk.END)
        self.available_map.clear()
        self.status_label.config(text="Updating package index...", fg="blue")
        threading.Thread(target=lambda: self.load_available_languages(update_index=True), daemon=True).start()

    def load_available_languages(self, update_index=False):
        try:
            if update_index:
                self.translation_manager.update_package_index()
            packages = self.translation_manager.get_available_packages()
            self.window.after(0, self.update_available_list, packages)
        except Exception as e:
            self.window.after(0, self.status_label.config, {"text": f"Error: {str(e)}", "fg": "red"})

    def update_available_list(self, packages):
        self.available_listbox.delete(0, tk.END)
        self.available_map.clear()
        
        # Sort packages
        packages.sort(key=lambda p: str(p))
        
        count = 0
        for pkg in packages:
             # Check if already installed (simplistic check by string representation or code)
            is_installed = any(str(i) == str(pkg) for i in self.translation_manager.get_installed_languages())
            if not is_installed:
                display_text = f"{pkg.from_name} -> {pkg.to_name}"
                self.available_listbox.insert(tk.END, display_text)
                self.available_map[count] = pkg
                count += 1
        
        if count == 0:
            self.status_label.config(text="No packages found. Click 'Refresh List' to update index.", fg="orange")
        elif not getattr(self, '_suppress_index_update_status', False):
            self.status_label.config(text="Loaded available packages from local index.", fg="green")

    def refresh_installed_languages(self):
        self.installed_listbox.delete(0, tk.END)
        self.installed_map.clear()
        
        installed = self.translation_manager.get_installed_languages()
        # Ensure we have a way to uniquely identify them or display them
        # argostranslate package objects have standard str representation
        
        count = 0
        for pkg in installed:
            display_text = f"{pkg.from_name} -> {pkg.to_name}"
            self.installed_listbox.insert(tk.END, display_text)
            self.installed_map[count] = pkg
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
        package = self.available_map.get(index)
        
        if package:
            self.install_btn.config(state=tk.DISABLED)
            
            self.status_label.config(text=f"Downloading:\n{package.from_name} -> {package.to_name}...", fg="blue")
            
            # Reset progress bar
            self.progress['mode'] = 'determinate'
            self.progress['value'] = 0
            
            threading.Thread(target=self.run_install, args=(package,), daemon=True).start()

    def run_install(self, package):
        def update_progress(percentage):
            self.window.after(0, lambda: self.progress.configure(value=percentage))
            
        success = self.translation_manager.install_package(package, progress_callback=update_progress)
        self.window.after(0, self.on_install_complete, success, package)

    def on_install_complete(self, success, package):
        self.progress.stop()
        self.install_btn.config(state=tk.DISABLED) # Will be re-enabled if selection changes
        
        if success:
            self.progress['value'] = 100
            self.status_label.config(text="Installation complete.", fg="green")
            self.refresh_installed_languages()
            
            # Suppress the "Index updated" message from the re-fetch
            self._suppress_index_update_status = True
            
            # Revert message after 3 seconds
            def revert_status():
                self.status_label.config(text="Index updated.", fg="green")
                self._suppress_index_update_status = False
                self.progress['value'] = 0 
                
            self.window.after(3000, revert_status)

            # Remove from available list
            # Re-fetch or just remove locally? 
            # Re-fetching is safer but slower. 
            # Let's just reload the list by filtering installed
            threading.Thread(target=lambda: self.load_available_languages(update_index=False), daemon=True).start()
        else:
            self.status_label.config(text="Installation failed.", fg="red")
            messagebox.showerror("Error", "Failed to install language package.")

    def on_installed_select(self, event):
        selection = self.installed_listbox.curselection()
        if selection:
            self.delete_btn.config(state=tk.NORMAL)
        else:
            self.delete_btn.config(state=tk.DISABLED)

    def delete_selected(self):
        selection = self.installed_listbox.curselection()
        if not selection:
            return
            
        index = selection[0]
        package = self.installed_map.get(index)
        
        if package:
            confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the package:\n{package.from_name} -> {package.to_name}?")
            if confirm:
                self.delete_btn.config(state=tk.DISABLED)
                self.status_label.config(text=f"Deleting {package.from_name} -> {package.to_name}...", fg="blue")
                self.progress.start(10)
                
                threading.Thread(target=self.run_uninstall, args=(package,), daemon=True).start()

    def run_uninstall(self, package):
        success = self.translation_manager.uninstall_package(package)
        self.window.after(0, self.on_uninstall_complete, success, package)

    def on_uninstall_complete(self, success, package):
        self.progress.stop()
        
        if success:
            self.status_label.config(text=f"Successfully deleted {package.from_name} -> {package.to_name}", fg="green")
            self.refresh_installed_languages()
            # Update available list as the deleted package is now available again
            threading.Thread(target=lambda: self.load_available_languages(update_index=False), daemon=True).start()
        else:
            self.status_label.config(text="Deletion failed.", fg="red")
            messagebox.showerror("Error", "Failed to delete language package.")

    def open_local_folder(self):
        """Open the local translation package folder."""
        import os
        import subprocess
        
        folder_path = self.translation_manager.package_dir
        if not os.path.exists(folder_path):
            os.makedirs(folder_path, exist_ok=True)
            
        try:
            os.startfile(folder_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder: {str(e)}")

