"""
Game units editor window for managing game unit mappings
"""
import json
import os
import re
import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client

from ..constants import APP_DOCUMENTS_DIR
from ..window_geometry import apply_window_geometry

# Ready-made regex rules for the "Presets" dropdown, so common cases don't
# require hand-writing a pattern. Each adds a new Regex-mode row prefilled
# with pattern/replacement; users can still tweak it afterward like any
# other row. Add more entries here as common requests come up.
REGEX_PRESETS = [
    {
        "name": "Suppress DC roll text (e.g. \"DC: 45 Failure\")",
        "pattern": r"DC:\s*\d+\s*(Failure|Success)",
        "replacement": ".",
        "example": "DC: 45 Failure",
    },
]

# Plain-English cheatsheet shown by the "?" button in the pattern tester, for
# users who have never written a regex pattern before. Each row is
# (symbol, meaning, example). Kept short on purpose - this is a jumping-off
# point, not a full regex reference.
REGEX_GUIDE_ROWS = [
    (r"\d+", "one or more digits (a number)", "matches \"45\" in \"DC: 45\""),
    (r"\s*", "any number of spaces, or none", "lets \"DC:45\" and \"DC: 45\" both match"),
    (r"(A|B)", "either A or B", "(Failure|Success) matches either word"),
    (r".", "any single character", "D.C matches \"DAC\", \"DBC\", etc."),
    (r"*", "zero or more of what came before it", "\\d* matches \"\", \"5\", \"50\", ..."),
    (r"+", "one or more of what came before it", "\\d+ matches \"5\" or \"50\", but not \"\""),
    (r"?", "optional - zero or one of what came before it", "colou?r matches \"color\" and \"colour\""),
    (r"\b", "a word boundary (whole words only)", "\\bcat\\b won't match \"category\""),
    (r"\1  \2", "reuse a (parenthesized) part in the replacement", "swap \"(\\w+), (\\w+)\" to \"\\2 \\1\""),
]


class GameUnitsEditWindow:
    def __init__(self, root, game_text_reader):
        self.root = root
        self.game_text_reader = game_text_reader
        self.window = tk.Toplevel(root)
        self.window.withdraw()
        self.window.title("Text Replacement Editor")
        self.window.resizable(True, True)
        # 700 (not the old 620) to comfortably fit the Test button added
        # alongside Listen/Delete/Default in the Actions column.
        apply_window_geometry(self.window, 'text_replacement_editor', 700, 600)

        # Register this window as one that disables hotkeys
        self.game_text_reader.register_hotkey_disabling_window("Gamer Units", self.window)

        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            pass

        # Load game units data
        self.game_units = self.game_text_reader.load_game_units()
        self.original_units = self.game_units.copy()
        self.original_regex_settings = self.load_regex_settings()
        
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
        self.window.deiconify()

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
                print(f"[WARNING] Game Units Editor: Voice lookup failed: {e}")
        
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
        tk.Label(header_frame, text="Detected text:", font=("Helvetica", 10, "bold"), width=15, anchor='w').pack(side='left', padx=5)
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

        # Presets button - drops down a menu of ready-made regex rules so
        # users don't have to hand-write a pattern for common cases.
        presets_button = tk.Menubutton(left_frame, text="Presets", relief='raised', width=10)
        presets_menu = tk.Menu(presets_button, tearoff=0)
        for preset in REGEX_PRESETS:
            presets_menu.add_command(label=preset["name"], command=lambda p=preset: self.add_preset_entry(p))
        presets_button.config(menu=presets_menu)
        presets_button.pack(side='left', padx=2)

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
            print(f"[ERROR] Game Units Editor: Load failed: {e}")
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
            print(f"[ERROR] Game Units Editor: Save failed: {e}")

    def load_regex_settings(self):
        """Load per-entry regex-mode settings from JSON file."""
        try:
            temp_path = APP_DOCUMENTS_DIR
            os.makedirs(temp_path, exist_ok=True)

            file_path = os.path.join(temp_path, 'gamer_units_regex.json')

            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {}
        except Exception as e:
            print(f"[ERROR] Game Units Editor: Load failed: {e}")
            return {}

    def save_regex_settings(self, regex_dict):
        """Save per-entry regex-mode settings to JSON file."""
        try:
            temp_path = APP_DOCUMENTS_DIR
            os.makedirs(temp_path, exist_ok=True)

            file_path = os.path.join(temp_path, 'gamer_units_regex.json')

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(regex_dict, f, indent=4, ensure_ascii=False)

        except Exception as e:
            print(f"[ERROR] Game Units Editor: Save failed: {e}")

    def populate_entries(self):
        """Populate the scrollable frame with existing game units."""
        # Load case-sensitive and regex settings if they exist
        case_sensitive_settings = self.load_case_sensitive_settings()
        regex_settings = self.original_regex_settings

        for short_name, full_name in self.game_units.items():
            case_sensitive = case_sensitive_settings.get(short_name, False)
            is_regex = regex_settings.get(short_name, False)
            self.add_entry_row(short_name, full_name, case_sensitive, is_regex)

    def add_entry_row(self, short_name="", full_name="", case_sensitive=False, is_regex=False):
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
        case_checkbox = tk.Checkbutton(case_frame, variable=case_sensitive_var, text="Case Sensitive\non Scan")
        case_checkbox.pack(side='left')

        # Regex checkbox - when enabled, "Detected text" is a regex pattern
        # (matched against the raw OCR text) instead of a literal word, and
        # "Will be read as" is its replacement (supports \1, \2 backreferences).
        is_regex_var = tk.BooleanVar(value=is_regex)
        regex_frame = tk.Frame(row_frame)
        regex_frame.pack(side='left', padx=0)
        regex_checkbox = tk.Checkbutton(regex_frame, variable=is_regex_var, text="Regex")
        regex_checkbox.pack(side='left')
        # Stashed on the row frame so save_units()/cancel_edit() can read it
        # back without changing the shape of entry_widgets everywhere else.
        row_frame.is_regex_var = is_regex_var

        # Actions frame
        actions_frame = tk.Frame(row_frame)
        actions_frame.pack(side='left', padx=10)
        
        # Listen button - use lambda with default argument to capture current value
        listen_btn = tk.Button(actions_frame, text="Listen", command=lambda var=full_name_var: self.listen_to_text(var.get()), width=7)
        listen_btn.pack(side='left', padx=2)

        # Test Regex button - opens a small dialog to try this row's pattern
        # against sample text without needing to save and run an actual OCR
        # scan. Only enabled while Regex is checked, both because testing
        # matters most there and because a disabled "Test Regex" makes it
        # obvious at a glance that this button is a regex-only feature.
        # getattr(row_frame, 'example_text', ...) is looked up at click time
        # (not row-creation time) so it still finds an example set by
        # add_preset_entry() right after this row was created.
        test_btn = tk.Button(actions_frame, text="Test Regex", command=lambda: self.open_pattern_tester(short_name_var, full_name_var, is_regex_var, case_sensitive_var, getattr(row_frame, 'example_text', None)), width=10)
        test_btn.pack(side='left', padx=2)

        def update_test_button_state(*_args):
            test_btn.config(state='normal' if is_regex_var.get() else 'disabled')

        is_regex_var.trace_add('write', update_test_button_state)
        update_test_button_state()

        # Delete button
        delete_btn = tk.Button(actions_frame, text="Delete", command=lambda: self.delete_entry(row_frame, short_name_var, full_name_var, case_sensitive_var), width=7)
        delete_btn.pack(side='left', padx=2)
        
        # Default button - only add if this row is within the default list range
        default_btn = None
        if has_default:
            default_btn = tk.Button(actions_frame, text="Default", command=lambda: self.restore_default(short_name_var, full_name_var, case_sensitive_var, is_regex_var, row_frame), width=7)
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

    def add_preset_entry(self, preset):
        """Add a new row prefilled from a Presets menu selection."""
        self.add_entry_row(preset["pattern"], preset["replacement"], case_sensitive=False, is_regex=True)
        # Stash the preset's sample text on the row so its Test Match button
        # can prefill the tester dialog with a working example.
        self.entry_widgets[-1][10].example_text = preset.get("example")
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)
    
    def get_default_units(self):
        """Get the default game units (empty - users add their own)."""
        return {}
    
    def restore_default(self, short_name_var, full_name_var, case_sensitive_var, is_regex_var, row_frame):
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
        current_is_regex = is_regex_var.get()

        # Check if already at default
        if (current_short_name == default_short_name and
            current_full_name == default_full_name and
            not current_case_sensitive and
            not current_is_regex):
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
            is_regex_var.set(False)
    
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
            print(f"[ERROR] Game Units Editor: Play failed: {e}")
            messagebox.showerror("Error", f"Could not read text: {e}")
    
    def stop_speech(self):
        """Stop any ongoing speech."""
        try:
            if self.current_speaker:
                self.current_speaker.Speak("", 2)  # 2 is SVSFPurgeBeforeSpeak
                self.current_speaker = None
        except Exception as e:
            print(f"[ERROR] Game Units Editor: Stop failed: {e}")
        
        # Also stop main window speech if needed
        if hasattr(self.game_text_reader, 'stop_speaking'):
            self.game_text_reader.stop_speaking()

    def open_pattern_tester(self, short_name_var, full_name_var, is_regex_var, case_sensitive_var, example_text=None):
        """Open a small dialog to try this row's pattern against sample text.

        Reads the row's StringVars live, so edits made after the dialog is
        open (including toggling Regex/Case Sensitive) are reflected without
        needing to reopen it. `example_text`, when known (e.g. from a
        Presets row), prefills the sample box so the result is visible the
        instant the dialog opens instead of requiring the user to type
        something first.
        """
        tester = tk.Toplevel(self.window)
        tester.title("Test This Rule")
        tester.transient(self.window)
        # Resizable in height too, since toggling the guide below changes
        # how tall the window needs to be.
        tester.resizable(True, True)
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                tester.iconbitmap(icon_path)
        except Exception:
            pass

        DIALOG_WRAP = 560

        tk.Label(tester, text="Sample text to test against:", font=("Helvetica", 9, "bold")).pack(anchor='w', padx=10, pady=(10, 0))
        tk.Label(tester, text="This only tests the pattern already saved in this row's \"Detected text\" / \"Will be read as\" fields - edit those in the main list, not here.", fg="#666666", wraplength=DIALOG_WRAP, justify='left', anchor='w').pack(fill='x', padx=10, pady=(0, 4))

        sample_var = tk.StringVar()
        sample_entry = tk.Entry(tester, textvariable=sample_var)
        sample_entry.pack(fill='x', padx=10)

        # Gray "sample text here" placeholder shown only while the field is
        # empty and unfocused - cleared the moment the user clicks in or
        # types, and restored if they click away leaving it empty again.
        PLACEHOLDER_TEXT = "sample text here"
        normal_fg = sample_entry.cget('fg')
        placeholder_active = {'value': False}

        def show_placeholder():
            placeholder_active['value'] = True
            sample_entry.config(fg='gray')
            sample_entry.insert(0, PLACEHOLDER_TEXT)

        def clear_placeholder(*_args):
            if placeholder_active['value']:
                placeholder_active['value'] = False
                sample_entry.delete(0, tk.END)
                sample_entry.config(fg=normal_fg)

        def restore_placeholder_if_empty(*_args):
            if not sample_var.get():
                show_placeholder()

        sample_entry.bind('<FocusIn>', clear_placeholder)
        sample_entry.bind('<FocusOut>', restore_placeholder_if_empty)

        # Literal (non-regex) mode: "Detected text" is the exact word/phrase,
        # so it trivially works as its own example. Regex mode: only prefill
        # when a known example was supplied (e.g. from a preset) - a
        # made-up regex pattern can't be turned into sample text reliably.
        if not example_text and not is_regex_var.get():
            example_text = short_name_var.get()

        tk.Label(tester, text="Tip: type (or paste) a line the way it'd actually appear in-game, including any parts that change, like numbers.", fg="#666666", wraplength=DIALOG_WRAP, justify='left', anchor='w').pack(fill='x', padx=10, pady=(4, 0))

        guide_btn = tk.Button(tester, text="Show Guide", width=12)
        guide_btn.pack(anchor='w', padx=10, pady=(4, 4))

        result_label = tk.Label(tester, text="Type sample text above to see the result.", wraplength=DIALOG_WRAP, justify='left', anchor='w')
        result_label.pack(fill='x', padx=10, pady=10)

        # Regex quick guide - built once, hidden until "Show Guide" is
        # clicked, so people who already know regex never see it and people
        # who don't have it one click away without a separate window to
        # manage. Laid out as two side-by-side columns instead of one long
        # list so the dialog stays short rather than very tall.
        guide_frame = tk.Frame(tester, bd=1, relief='groove')
        tk.Label(guide_frame, text="Regex quick guide", font=("Helvetica", 9, "bold")).pack(anchor='w', padx=8, pady=(6, 0))
        tk.Label(guide_frame, text="Uses Python's regex syntax (the \"re\" module) - almost identical to most other regex flavors.", font=("Helvetica", 8), fg="#666666", wraplength=DIALOG_WRAP, justify='left', anchor='w').pack(fill='x', padx=8, pady=(0, 4))

        columns_frame = tk.Frame(guide_frame)
        columns_frame.pack(fill='x', padx=8)
        split = (len(REGEX_GUIDE_ROWS) + 1) // 2
        column_rows = [REGEX_GUIDE_ROWS[:split], REGEX_GUIDE_ROWS[split:]]
        for i, rows in enumerate(column_rows):
            col_frame = tk.Frame(columns_frame)
            col_frame.pack(side='left', fill='both', expand=True, padx=(0 if i == 0 else 6, 6 if i == 0 else 0))
            for symbol, meaning, example in rows:
                row = tk.Frame(col_frame)
                row.pack(fill='x', pady=2)
                tk.Label(row, text=symbol, font=("Consolas", 9, "bold"), width=7, anchor='nw', justify='left').pack(side='left')
                text_col = tk.Frame(row)
                text_col.pack(side='left', fill='x', expand=True)
                tk.Label(text_col, text=meaning, font=("Helvetica", 8), anchor='w', justify='left', wraplength=200).pack(fill='x')
                tk.Label(text_col, text=example, font=("Helvetica", 8, "italic"), fg="#666666", anchor='w', justify='left', wraplength=200).pack(fill='x')

        tk.Label(guide_frame, text="For anything trickier, a free site like regex101.com will explain any pattern you paste in.", font=("Helvetica", 8), fg="#666666", wraplength=DIALOG_WRAP, justify='left', anchor='w').pack(fill='x', padx=8, pady=(2, 6))

        # A complete, working example people can copy straight into the
        # Detected text / Will be read as fields (and this dialog's sample
        # box) to see a real regex rule work end to end before writing
        # their own. Lives inside the guide so it only takes up space once
        # someone actually asks to see the guide.
        example_frame = tk.LabelFrame(guide_frame, text="Example you can copy", font=("Helvetica", 8, "bold"), fg="#444444")
        example_frame.pack(fill='x', padx=8, pady=(0, 8))

        def add_copyable_row(parent, label_text, value):
            row = tk.Frame(parent)
            row.pack(fill='x', padx=6, pady=2)
            tk.Label(row, text=label_text, font=("Helvetica", 8), width=13, anchor='w').pack(side='left')
            field = tk.Entry(row, font=("Consolas", 8))
            field.insert(0, value)
            field.config(state='readonly', readonlybackground='white')
            field.pack(side='left', fill='x', expand=True, padx=(2, 4))

            copy_btn = tk.Button(row, text="Copy", width=8)
            copy_btn.pack(side='left')

            def copy_value(btn=copy_btn):
                tester.clipboard_clear()
                tester.clipboard_append(value)
                # Brief "Copied!" confirmation so clicking the button isn't
                # silent - reverts on its own after a moment.
                btn.config(text="Copied!")
                tester.after(1200, lambda: btn.config(text="Copy"))

            copy_btn.config(command=copy_value)

        add_copyable_row(example_frame, "Detected text:", r"Damage:\s*\d+")
        add_copyable_row(example_frame, "Will be read as:", "some damage")
        add_copyable_row(example_frame, "Sample text:", "You dealt Damage: 32 to the enemy")

        def toggle_guide():
            if guide_frame.winfo_ismapped():
                guide_frame.pack_forget()
                guide_btn.config(text="Show Guide")
            else:
                guide_frame.pack(fill='x', padx=10, pady=(0, 8), before=result_label)
                guide_btn.config(text="Hide Guide")
            tester.update_idletasks()
            tester.geometry(f"{tester.winfo_width()}x{tester.winfo_reqheight()}")

        guide_btn.config(command=toggle_guide)

        def update_result(*_args):
            if placeholder_active['value']:
                result_label.config(text="Type sample text above to see the result.", fg="black")
                return

            pattern = short_name_var.get()
            replacement = full_name_var.get()
            sample = sample_var.get()
            is_regex = is_regex_var.get()
            case_sensitive = case_sensitive_var.get()

            if not pattern:
                result_label.config(text="Enter something in \"Detected text\" first.", fg="black")
                return
            if not sample:
                result_label.config(text="Type sample text above to see the result.", fg="black")
                return

            flags = 0 if case_sensitive else re.IGNORECASE
            # Literal mode matches on word boundaries, approximating the
            # real whole-word matching used at OCR time.
            search_pattern = pattern if is_regex else r'\b' + re.escape(pattern) + r'\b'

            try:
                new_text, count = re.subn(search_pattern, replacement, sample, flags=flags)
            except re.error as e:
                result_label.config(text=f"Invalid regex: {e}", fg="red")
                return

            if count:
                result_label.config(text=f"✓ Matches ({count}x) → would be read as:\n\"{new_text}\"", fg="#1a7f37")
            else:
                result_label.config(text="✗ No match — this sample would be read as-is.", fg="#b71c1c")

        sample_var.trace_add('write', update_result)
        if example_text:
            sample_var.set(example_text)
            # Auto-focus + select only when there's real prefilled text, so
            # typing immediately overwrites it. Skip this when showing the
            # placeholder instead, since focusing it would immediately
            # trigger <FocusIn> and clear the placeholder before it's seen.
            sample_entry.focus_set()
            sample_entry.selection_range(0, tk.END)
        else:
            show_placeholder()

        tk.Button(tester, text="Close", command=tester.destroy, width=10).pack(pady=(0, 10))

        apply_window_geometry(tester, 'text_replacement_tester', 600, 260, parent=self.window)

    def save_units(self):
        """Save the game units to the JSON file."""
        # Collect data from all entries
        new_units = {}
        case_sensitive_settings = {}
        regex_settings = {}
        errors = []

        for short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame in self.entry_widgets:
            short_name = short_name_var.get().strip()
            full_name = full_name_var.get().strip()
            case_sensitive = case_sensitive_var.get()
            is_regex = getattr(row_frame, 'is_regex_var', None)
            is_regex = is_regex.get() if is_regex is not None else False

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

            # Regex entries must compile - catch typos here instead of at OCR time
            if is_regex:
                try:
                    re.compile(short_name)
                except re.error as e:
                    errors.append(f"Invalid regex pattern '{short_name}': {e}")
                    continue

            new_units[short_name] = full_name
            case_sensitive_settings[short_name] = case_sensitive
            regex_settings[short_name] = is_regex

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
//  Enable the "Regex" checkbox on an entry to match a regex pattern instead
//  of a literal word - "Full Name" is then the replacement text and may use
//  \\1, \\2, etc. to reference capture groups from the pattern.

'''
                    f.write(header)
                    json.dump(new_units, f, indent=4, ensure_ascii=False)
                else:
                    # Write empty JSON object for completely empty file
                    f.write('{}')

            # Save case-sensitive and regex settings
            self.save_case_sensitive_settings(case_sensitive_settings)
            self.save_regex_settings(regex_settings)

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
            print(f"[ERROR] Game Units Editor: Save failed: {e}")
    
    def cancel_edit(self):
        """Cancel editing and close the window."""
        # Check if there are unsaved changes
        current_units = {}
        current_case_sensitive = {}
        current_is_regex = {}
        for short_name_var, full_name_var, case_sensitive_var, short_entry, full_entry, case_frame, case_checkbox, listen_btn, delete_btn, default_btn, row_frame in self.entry_widgets:
            short_name = short_name_var.get().strip()
            full_name = full_name_var.get().strip()
            case_sensitive = case_sensitive_var.get()
            is_regex_var = getattr(row_frame, 'is_regex_var', None)
            is_regex = is_regex_var.get() if is_regex_var is not None else False
            if short_name and full_name:
                current_units[short_name] = full_name
                current_case_sensitive[short_name] = case_sensitive
                current_is_regex[short_name] = is_regex

        # Load original case-sensitive settings for comparison
        original_case_sensitive = self.load_case_sensitive_settings()

        if (current_units != self.original_units
                or current_case_sensitive != original_case_sensitive
                or current_is_regex != self.original_regex_settings):
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

