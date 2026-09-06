import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import argostranslate.settings
from ..constants import APP_SETTINGS_PATH
from ..translation.translation_manager import TranslationManager
from ..window_geometry import apply_window_geometry
from .common import set_window_icon

class PerformanceSettingsWindow:
    def __init__(self, game_reader):
        self.game_reader = game_reader
        self.window = tk.Toplevel(game_reader.root)
        self.window.withdraw()
        self.window.title("Argos-Translate Performance Settings")
        self.window.resizable(False, False)
        apply_window_geometry(self.window, 'performance_settings', 450, 550)

        # Register as hotkey disabling window
        if hasattr(self.game_reader, 'register_hotkey_disabling_window'):
            self.game_reader.register_hotkey_disabling_window("Performance Settings", self.window)
            
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Set the window icon
        set_window_icon(self.window)

        # Load current settings
        self.current_settings = self.load_settings()
        
        self.setup_ui()
        self.window.deiconify()

    def on_close(self):
        if hasattr(self.game_reader, 'unregister_hotkey_disabling_window'):
            self.game_reader.unregister_hotkey_disabling_window("Performance Settings")
        self.window.destroy()
        
    def load_settings(self):
        """Load translation settings from total settings file"""
        defaults = {
            "device": "cpu",
            "compute_type": "float32",
            "beam_size": 2, # Default is usually 1 or 2
            "token_batch_size": 0, # 0 means default
            "debug": False
        }
        
        try:
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    all_settings = json.load(f)
                    return all_settings.get("translation_settings", defaults)
        except Exception as e:
            print(f"[ERROR] Settings window: Load failed: {e}")
            
        return defaults

    def save_settings(self):
        """Save settings to global file and apply them"""
        new_settings = {
            "device": self.device_var.get(),
            "compute_type": self.compute_type_var.get(),
            "beam_size": int(self.beam_size_var.get()),
            # "token_batch_size": int(self.batch_size_var.get()),
            "debug": self.debug_var.get()
        }
        
        try:
            # Read existing
            all_settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    all_settings = json.load(f)
            
            # Update translation section
            all_settings["translation_settings"] = new_settings
            
            # Write back
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(all_settings, f, indent=4)
                
            # Apply immediately
            TranslationManager().apply_settings(new_settings)
            
            messagebox.showinfo("Success", "Settings saved and applied.")
            self.on_close()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {e}")

    def setup_ui(self):
        main_frame = tk.Frame(self.window, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Device
        tk.Label(main_frame, text="Device:", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(main_frame, text="• cpu: Uses computer processor\n• cuda: Uses Nvidia GPU (Requires installed Nvidia graphics drivers)\n• auto: Auto-detects best available", 
                 justify="left", fg="#555555", font=("Helvetica", 8)).pack(anchor="w", pady=(0, 5))
        self.device_var = tk.StringVar(value=self.current_settings.get("device", "cpu"))
        device_combo = ttk.Combobox(main_frame, textvariable=self.device_var, values=["cpu", "cuda", "auto"], state="readonly")
        device_combo.pack(fill="x", pady=(0, 15))
        
        # Compute Type
        tk.Label(main_frame, text="Compute Type (Precision):", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(main_frame, text="Controls calculation precision. Lower precision (e.g. int8) is faster with minimal quality loss. 'float32' is standard.", 
                 justify="left", fg="#555555", font=("Helvetica", 8), wraplength=400).pack(anchor="w", pady=(0, 5))
        self.compute_type_var = tk.StringVar(value=self.current_settings.get("compute_type", "float32"))
        compute_values = ["int8", "int8_float16", "int16", "float16", "float32", "auto"]
        compute_combo = ttk.Combobox(main_frame, textvariable=self.compute_type_var, values=compute_values, state="readonly")
        compute_combo.pack(fill="x", pady=(0, 15))
        
        # Beam Size
        tk.Label(main_frame, text="Beam Size:", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(0, 2))
        tk.Label(main_frame, text="Controls translation quality vs speed. Higher number = checks more translation possibilities (Better Quality, Slower). Lower number = faster.", 
                 justify="left", fg="#555555", font=("Helvetica", 8), wraplength=400).pack(anchor="w", pady=(0, 5))
        self.beam_size_var = tk.StringVar(value=str(self.current_settings.get("beam_size", 2)))
        beam_spin = ttk.Spinbox(main_frame, from_=1, to=20, textvariable=self.beam_size_var)
        beam_spin.pack(fill="x", pady=(0, 15))

        # Debug
        self.debug_var = tk.BooleanVar(value=self.current_settings.get("debug", False))
        tk.Checkbutton(main_frame, text="Enable Debug Logging (Output to Console Window)", variable=self.debug_var).pack(anchor="w", pady=(0, 15))
        
        # Save Button
        tk.Button(main_frame, text="Save & Apply", command=self.save_settings, width=15, height=2).pack(side="bottom", pady=10)
