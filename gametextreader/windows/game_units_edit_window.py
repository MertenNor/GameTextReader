"""
Game units editor window for managing game unit mappings
"""
import json
import os
import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client

from ..constants import APP_DOCUMENTS_DIR


class GameUnitsEditWindow:
    def __init__(self, root, game_text_reader):
        self.root = root
        self.game_text_reader = game_text_reader
        self.window = tk.Toplevel(root)
        self.window.title("Edit Gamer Units")
        self.window.geometry("500x600")
        self.window.resizable(True, True)
        
        # Register this window as one that disables hotkeys
        self.game_text_reader.register_hotkey_disabling_window("Gamer Units", self.window)
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting game units editor icon: {e}")
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.window.winfo_screenheight() // 2) - (600 // 2)
        self.window.geometry(f"500x600+{x}+{y}")
        
        # Load game units data
        self.game_units = self.game_text_reader.load_game_units()
        self.original_units = self.game_units.copy()
        
        # Get default units as a list to preserve order
        default_units_dict = self.get_default_units()
        self.default_units_list = [(short, full) for short, full in default_units_dict.items()]
        
        # Store entry widgets and variables
        self.entry_widgets = []  # List of (short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame)
        
        # Voice selection variables
        self.selected_voice = None
        self.current_speaker = None
        
        # Set up protocol to handle window closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Create UI
        self.create_ui()
        
        # Populate with existing data
        self.populate_entries()
    
    def on_close(self):
        """Handle window closing."""
        self.cancel_edit()
    
    def create_ui(self):
        """Create the user interface for the editor."""
        # Top frame with voice selection and Stop button
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill='x', padx=10, pady=10)
        
        # Voice selection label
        tk.Label(top_frame, text="Preview Voice:", font=("Helvetica", 10)).pack(side='left', padx=5)
        
        # Voice selection dropdown
        voice_display_names = []
        voice_full_names = {}
        
        if hasattr(self.game_text_reader, 'voices') and self.game_text_reader.voices:
            try:
                for i, voice in enumerate(self.game_text_reader.voices, 1):
                    full_name = voice.GetDescription()
                    
                    # Create abbreviated display name with numbering
                    if "Microsoft" in full_name and " - " in full_name:
                        parts = full_name.split(" - ")
                        if len(parts) == 2:
                            voice_part = parts[0].replace("Microsoft ", "")
                            lang_part = parts[1]
                            display_name = f"{i}. {voice_part} ({lang_part})"
                        else:
                            display_name = f"{i}. {full_name}"
                    elif " - " in full_name:
                        parts = full_name.split(" - ")
                        if len(parts) == 2:
                            display_name = f"{i}. {parts[0]} ({parts[1]})"
                        else:
                            display_name = f"{i}. {full_name}"
                    else:
                        display_name = f"{i}. {full_name}"
                    
                    voice_display_names.append(display_name)
                    voice_full_names[display_name] = full_name
            except Exception as e:
                print(f"Warning: Could not get voice descriptions: {e}")
        
        # Set default to first voice or fallback
        default_voice_display = voice_display_names[0] if voice_display_names else "No voices available"
        self.voice_var = tk.StringVar(value=default_voice_display)
        if default_voice_display in voice_full_names:
            self.selected_voice = voice_full_names[default_voice_display]
        
        # Function to update the actual voice when display name is selected
        def on_voice_selection(*args):
            selected_display = self.voice_var.get()
            if selected_display in voice_full_names:
                self.selected_voice = voice_full_names[selected_display]
            else:
                self.selected_voice = selected_display
        
        # Create the OptionMenu with voices
        voice_menu = tk.OptionMenu(
            top_frame,
            self.voice_var,
            *voice_display_names if voice_display_names else ["No voices available"],
            command=on_voice_selection
        )
        voice_menu.config(width=30, anchor="w")
        voice_menu.pack(side='left', padx=5)
        
        # Stop button
        stop_button = tk.Button(top_frame, text="Stop", command=self.stop_speech, width=8)
        stop_button.pack(side='left', padx=10)
        
        # Separator
        ttk.Separator(self.window, orient='horizontal').pack(fill='x', padx=10, pady=5)
        
        # Scrollable frame for entries
        canvas_frame = tk.Frame(self.window)
        canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Create window that fills the canvas width
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Make the scrollable frame fill the canvas width
        def configure_scroll_region(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind('<Configure>', configure_scroll_region)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mousewheel to canvas and canvas_frame for better coverage
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind to canvas and canvas_frame, and also use bind_all for global capture
        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas_frame.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Also bind Linux mousewheel events
        def _on_mousewheel_linux(event):
            canvas.yview_scroll(int(-1*(event.delta)), "units")
        canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))
        
        # Set focus to canvas when mouse enters
        def _on_enter(event):
            canvas.focus_set()
        canvas.bind("<Enter>", _on_enter)
        canvas_frame.bind("<Enter>", _on_enter)
        
        # Headers
        header_frame = tk.Frame(self.scrollable_frame)
        header_frame.pack(fill='x', padx=0, pady=5)
        tk.Label(header_frame, text="Detected phrase:", font=("Helvetica", 10, "bold"), width=15, anchor='w').pack(side='left', padx=5)
        tk.Label(header_frame, text="Will be read as:", font=("Helvetica", 10, "bold"), width=20, anchor='w').pack(side='left', padx=5)
        tk.Label(header_frame, text="Actions", font=("Helvetica", 10, "bold"), width=15, anchor='w').pack(side='left', padx=5)
        
        # Store canvas and scrollable_frame for later use
        self.canvas = canvas
        self.scrollable_frame = self.scrollable_frame
        
        # Separator
        ttk.Separator(self.window, orient='horizontal').pack(fill='x', padx=10, pady=5)
        
        # Bottom frame with Add New, Delete All, Reset to Default, Save, and Cancel buttons
        bottom_frame = tk.Frame(self.window)
        bottom_frame.pack(fill='x', padx=10, pady=10)
        
        # Left side buttons
        left_frame = tk.Frame(bottom_frame)
        left_frame.pack(side='left', padx=5)
        
        # Add New button
        add_button = tk.Button(left_frame, text="Add New", command=self.add_new_entry, width=10)
        add_button.pack(side='left', padx=2)
        
        # Delete All button
        delete_all_button = tk.Button(left_frame, text="Delete All", command=self.delete_all_entries, width=10)
        delete_all_button.pack(side='left', padx=2)
        
        # Reset to Default button
        reset_default_button = tk.Button(left_frame, text="Reset to Default", command=self.reset_to_default, width=12)
        reset_default_button.pack(side='left', padx=2)
        
        # Spacer
        tk.Frame(bottom_frame).pack(side='left', expand=True)
        
        # Right side buttons
        right_frame = tk.Frame(bottom_frame)
        right_frame.pack(side='right', padx=5)
        
        # Save button
        save_button = tk.Button(right_frame, text="Save", command=self.save_units, width=10)
        save_button.pack(side='right', padx=2)
        
        # Cancel button
        cancel_button = tk.Button(right_frame, text="Cancel", command=self.cancel_edit, width=10)
        cancel_button.pack(side='right', padx=2)
    
    def load_case_sensitive_settings(self):
        """Load case-sensitive settings from JSON file."""
        try:
            temp_path = APP_DOCUMENTS_DIR
            os.makedirs(temp_path, exist_ok=True)
            
            file_path = os.path.join(temp_path, 'gamer_units_case_sensitive.json')
            
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {}
        except Exception as e:
            print(f"Error loading case-sensitive settings: {e}")
            return {}
    
    def save_case_sensitive_settings(self, case_sensitive_dict):
        """Save case-sensitive settings to JSON file."""
        try:
            temp_path = APP_DOCUMENTS_DIR
            os.makedirs(temp_path, exist_ok=True)
            
            file_path = os.path.join(temp_path, 'gamer_units_case_sensitive.json')
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(case_sensitive_dict, f, indent=4, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving case-sensitive settings: {e}")
    
    def populate_entries(self):
        """Populate the scrollable frame with existing game units."""
        # Load case-sensitive settings if they exist
        case_sensitive_settings = self.load_case_sensitive_settings()
        
        for short_name, full_name in self.game_units.items():
            case_sensitive = case_sensitive_settings.get(short_name, False)
            self.add_entry_row(short_name, full_name, case_sensitive)
    
    def add_entry_row(self, short_name="", full_name="", case_sensitive=False):
        """Add a new row for editing a game unit entry."""
        row_frame = tk.Frame(self.scrollable_frame)
        row_frame.pack(fill='x', padx=0, pady=2)
        
        # Check if this row will be within the default list range
        current_row_index = len(self.entry_widgets)
        has_default = current_row_index < len(self.default_units_list)
        
        # Short name entry
        short_name_var = tk.StringVar(value=short_name)
        short_entry = tk.Entry(row_frame, textvariable=short_name_var, width=18)
        short_entry.pack(side='left', padx=5)
        
        # Full name entry
        full_name_var = tk.StringVar(value=full_name)
        full_entry = tk.Entry(row_frame, textvariable=full_name_var, width=20)
        full_entry.pack(side='left', padx=5)
        
        # Case sensitive checkbox
        case_sensitive_var = tk.BooleanVar(value=case_sensitive)
        case_frame = tk.Frame(row_frame)
        case_frame.pack(side='left', padx=0)
        case_checkbox = tk.Checkbutton(case_frame, variable=case_sensitive_var, text="Case\nSensitive")
        case_checkbox.pack(side='left')
        
        # Actions frame
        actions_frame = tk.Frame(row_frame)
        actions_frame.pack(side='left', padx=10)
        
        # Listen button - use lambda with default argument to capture current value
        listen_btn = tk.Button(actions_frame, text="Listen", command=lambda var=full_name_var: self.listen_to_text(var.get()), width=7)
        listen_btn.pack(side='left', padx=2)
        
        # Delete button
        delete_btn = tk.Button(actions_frame, text="Delete", command=lambda: self.delete_entry(row_frame, short_name_var, full_name_var, case_sensitive_var), width=7)
        delete_btn.pack(side='left', padx=2)
        
        # Default button - only add if this row is within the default list range
        default_btn = None
        if has_default:
            default_btn = tk.Button(actions_frame, text="Default", command=lambda: self.restore_default(short_name_var, full_name_var, case_sensitive_var, row_frame), width=7)
            default_btn.pack(side='left', padx=(2, 5))
        else:
            # Add padding to match spacing when there's no Default button
            tk.Frame(actions_frame, width=7).pack(side='left', padx=(2, 5))
        
        # Store widgets
        self.entry_widgets.append((short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame))
    
    def delete_all_entries(self):
        """Delete all entries from the editor."""
        if messagebox.askyesno("Delete All", "Are you sure you want to delete all game units? This action cannot be undone."):
            # Clear all entry widgets
            for widget in self.entry_widgets:
                row_frame = widget[10]  # row_frame is at index 10 now
                row_frame.destroy()
            self.entry_widgets.clear()
            
            # Add one empty entry row
            self.add_entry_row("", "")
    
    def reset_to_default(self):
        """Reset all entries to default values."""
        if messagebox.askyesno("Reset to Default", "This will replace all current entries with the default values. Any custom entries will be lost. Continue?"):
            # Clear all existing entries
            for widget in self.entry_widgets:
                row_frame = widget[10]  # row_frame is at index 10 now
                row_frame.destroy()
            self.entry_widgets.clear()
            
            # Add all default entries
            for short_name, full_name in self.default_units_list:
                self.add_entry_row(short_name, full_name)
            
            # Scroll to top
            self.canvas.update_idletasks()
            self.canvas.yview_moveto(0.0)
    
    def add_new_entry(self):
        """Add a new empty entry row."""
        self.add_entry_row("", "")
        # Scroll to bottom
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)
    
    def get_default_units(self):
        """Get the default game units from the source code."""
        return {
            'xp': 'Experience Points',
            'hp': 'Health Points',
            'mp': 'Mana Points',
            'gp': 'Gold Pieces',
            'pp': 'Platinum Pieces',
            'sp': 'Skill Points',
            'ep': 'Energy Points',
            'ap': 'Action Points',
            'bp': 'Battle Points',
            'lp': 'Loyalty Points',
            'cp': 'Challenge Points',
            'vp': 'Victory Points',
            'rp': 'Reputation Points',
            'tp': 'Talent Points',
            'ar': 'Armor Rating',
            'dmg': 'Damage',
            'dps': 'Damage Per Second',
            'def': 'Defense',
            'mat': 'Materials',
            'exp': 'Exploration Points',
            '§': 'Simoliance',
            'v-bucks': 'Virtual Bucks',
            'r$': 'Robux',
            'nmt': 'Nook Miles Tickets',
            'be': 'Blue Essence',
            'radianite': 'Radianite Points',
            'ow coins': 'Overwatch Coins',
            '₽': 'PokeDollars',
            '€$': 'Eurodollars',
            'z': 'Zenny',
            'l': 'Lunas',
            'e': 'Eve',
            'i': 'Isk',
            'j': 'Jewel',
            'sc': 'Star Coins',
            'o2': 'Oxygen',
            'pu': 'Power Units',
            'mc': 'Mana Crystals',
            'es': 'Essence',
            'sh': 'Shards',
            'st': 'Stars',
            'mu': 'Munny',
            'b': 'Bolts',
            'r': 'Rings',
            'ca': 'Caps',
            'rns': 'Runes',
            'sl': 'Souls',
            'fav': 'Favor',
            'am': 'Amber',
            'cc': 'Crystal Cores',
            'fg': 'Fragments'
        }
    
    def restore_default(self, short_name_var, full_name_var, case_sensitive_var, row_frame):
        """Restore the default value for a game unit entry based on its position in the list."""
        # Find the index of this row in the entry_widgets list
        row_index = None
        for i, (s_var, f_var, cs_var, s_entry, f_entry, cs_frame, cs_checkbox, l_btn, d_btn, def_btn, r_frame) in enumerate(self.entry_widgets):
            if r_frame == row_frame:
                row_index = i
                break
        
        if row_index is None:
            messagebox.showerror("Error", "Could not find row position.")
            return
        
        # Check if there's a default value for this position
        if row_index >= len(self.default_units_list):
            messagebox.showwarning("No Default", f"No default value available for position {row_index + 1}.")
            return
        
        # Get the default values for this position
        default_short_name, default_full_name = self.default_units_list[row_index]
        
        current_short_name = short_name_var.get().strip()
        current_full_name = full_name_var.get().strip()
        current_case_sensitive = case_sensitive_var.get()
        
        # Check if already at default
        if (current_short_name == default_short_name and 
            current_full_name == default_full_name and 
            not current_case_sensitive):
            messagebox.showinfo("Already Default", f"This row is already set to its default values:\nShort: '{default_short_name}'\nFull: '{default_full_name}'\nCase Sensitive: No")
            return
        
        # Prompt before applying
        if messagebox.askyesno("Restore Default", 
                               f"Restore this row to default values (position {row_index + 1})?\n\n"
                               f"Current:\n  Short: {current_short_name or '(empty)'}\n  Full: {current_full_name or '(empty)'}\n  Case Sensitive: {current_case_sensitive}\n\n"
                               f"Default:\n  Short: {default_short_name}\n  Full: {default_full_name}\n  Case Sensitive: No"):
            short_name_var.set(default_short_name)
            full_name_var.set(default_full_name)
            case_sensitive_var.set(False)
    
    def delete_entry(self, row_frame, short_name_var, full_name_var, case_sensitive_var):
        """Delete an entry row."""
        # Remove from entry_widgets list
        for i, (s_var, f_var, cs_var, s_entry, f_entry, cs_frame, cs_checkbox, l_btn, d_btn, def_btn, r_frame) in enumerate(self.entry_widgets):
            if r_frame == row_frame:
                self.entry_widgets.pop(i)
                break
        
        # Destroy the row frame
        row_frame.destroy()
    
    def listen_to_text(self, text):
        """Read the given text aloud using the selected voice."""
        if not text:
            return
        
        # Stop any current speech
        self.stop_speech()
        
        # Get the selected voice
        voice = self.selected_voice
        if not voice and hasattr(self.game_text_reader, 'voices') and self.game_text_reader.voices:
            # Use first available voice if none selected
            try:
                voice = self.game_text_reader.voices[0].GetDescription()
            except (IndexError, AttributeError, Exception):
                # Silently fail if no voices available or voice object doesn't have GetDescription
                pass
        
        if not voice:
            messagebox.showwarning("No Voice Selected", "Please select a voice from the dropdown.")
            return
        
        # Create a temporary speaker for this window
        try:
            self.current_speaker = win32com.client.Dispatch("SAPI.SpVoice")
            
            # Set the voice
            for v in self.game_text_reader.voices:
                try:
                    if v.GetDescription() == voice:
                        self.current_speaker.Voice = v
                        break
                except (AttributeError, Exception):
                    # Voice object may not have GetDescription or setting voice may fail
                    continue
            
            # Set volume
            if hasattr(self.game_text_reader, 'volume'):
                self.current_speaker.Volume = int(self.game_text_reader.volume.get())
            
            # Speak the text
            self.current_speaker.Speak(text, 1)  # 1 is SVSFlagsAsync
        except Exception as e:
            print(f"Error speaking text: {e}")
            messagebox.showerror("Error", f"Could not read text: {e}")
    
    def stop_speech(self):
        """Stop any ongoing speech."""
        try:
            if self.current_speaker:
                self.current_speaker.Speak("", 2)  # 2 is SVSFPurgeBeforeSpeak
                self.current_speaker = None
        except Exception as e:
            print(f"Error stopping speech: {e}")
        
        # Also stop main window speech if needed
        if hasattr(self.game_text_reader, 'stop_speaking'):
            self.game_text_reader.stop_speaking()
    
    def save_units(self):
        """Save the game units to the JSON file."""
        # Collect data from all entries
        new_units = {}
        case_sensitive_settings = {}
        errors = []
        
        for short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame in self.entry_widgets:
            short_name = short_name_var.get().strip()
            full_name = full_name_var.get().strip()
            case_sensitive = case_sensitive_var.get()
            
            # Skip empty entries
            if not short_name and not full_name:
                continue
            
            # Validate
            if not short_name:
                errors.append("One or more entries have empty short names.")
                continue
            
            if not full_name:
                errors.append("One or more entries have empty full names.")
                continue
            
            # Check for duplicate short names
            if short_name in new_units:
                errors.append(f"Duplicate short name: '{short_name}'")
                continue
            
            new_units[short_name] = full_name
            case_sensitive_settings[short_name] = case_sensitive
        
        # Show errors if any
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return
        
        # Save main units file
        try:
            temp_path = APP_DOCUMENTS_DIR
            os.makedirs(temp_path, exist_ok=True)
            
            file_path = os.path.join(temp_path, 'gamer_units.json')
            
            with open(file_path, 'w', encoding='utf-8') as f:
                if new_units:
                    # Only write header and content if there are units to save
                    header = '''//  Game Units Configuration
//  Format: "short_name": "Full Name"
//  Example: "xp" will be read as "Experience Points"
//  Enable "Read gamer units" in the main window to use this feature

'''
                    f.write(header)
                    json.dump(new_units, f, indent=4, ensure_ascii=False)
                else:
                    # Write empty JSON object for completely empty file
                    f.write('{}')
            
            # Save case-sensitive settings
            self.save_case_sensitive_settings(case_sensitive_settings)
            
            # Update the game_text_reader's game_units
            self.game_text_reader.game_units = new_units
            self.game_units = new_units
            
            # Show success message
            messagebox.showinfo("Success", "Game units saved successfully!")
            
            # Unregister this window and clean up reference in game_text_reader
            self.game_text_reader.unregister_hotkey_disabling_window("Gamer Units")
            if hasattr(self.game_text_reader, '_game_units_editor'):
                self.game_text_reader._game_units_editor = None
            
            # Close the window
            self.window.destroy()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save game units: {str(e)}")
            print(f"Error saving game units: {e}")
    
    def cancel_edit(self):
        """Cancel editing and close the window."""
        # Check if there are unsaved changes
        current_units = {}
        current_case_sensitive = {}
        for short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame in self.entry_widgets:
            short_name = short_name_var.get().strip()
            full_name = full_name_var.get().strip()
            case_sensitive = case_sensitive_var.get()
            if short_name and full_name:
                current_units[short_name] = full_name
                current_case_sensitive[short_name] = case_sensitive
        
        # Load original case-sensitive settings for comparison
        original_case_sensitive = self.load_case_sensitive_settings()
        
        if current_units != self.original_units or current_case_sensitive != original_case_sensitive:
            if messagebox.askyesno("Unsaved Changes", "You have unsaved changes. Are you sure you want to cancel?"):
                # User clicked Yes - proceed with cancel
                pass
            else:
                # User clicked No - don't cancel, return to the window
                return
        
        # Stop any speech
        self.stop_speech()
        
        # Unregister this window and clean up reference in game_text_reader
        self.game_text_reader.unregister_hotkey_disabling_window("Gamer Units")
        if hasattr(self.game_text_reader, '_game_units_editor'):
            self.game_text_reader._game_units_editor = None
        
        # Close the window
        self.window.destroy()

