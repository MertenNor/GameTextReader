"""
Automations window for setting up if-then scenarios based on image detection
"""
import os
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageStat
import pytesseract

from ..screen_capture import capture_screen_area

# Try to import numpy for better image comparison (optional)
try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False
    # numpy is optional - methods will work without it


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


class AutomationsWindow:
    def __init__(self, root, game_text_reader):
        self.root = root
        self.game_text_reader = game_text_reader
        
        self.window = tk.Toplevel(self.root)
        self.window.title("Automations")
        self.window.geometry("800x600")
        self.window.resizable(True, True)
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting automations window icon: {e}")
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (840 // 2)
        y = (self.window.winfo_screenheight() // 2) - (600 // 2)
        self.window.geometry(f"840x600+{x}+{y}")
        
        # Store automation rules
        self.automations = []  # List of automation dictionaries
        
        # Store hotkey combos (hotkey + list of areas with timers)
        self.hotkey_combos = []  # List of hotkey combo dictionaries
        
        # Registry to store combo callbacks by hotkey name for reliable lookup
        self.combo_callbacks_by_hotkey = {}  # {hotkey_name: callback_function}
        
        # Registry to store automation callbacks by hotkey name for reliable lookup
        self.automation_callbacks_by_hotkey = {}  # {hotkey_name: callback_function}
        
        # Background polling control
        # Restore polling state from game_text_reader if it exists (persists across window close/reopen)
        # Check for old instance that was preserved before this new instance was created
        old_instance = None
        if hasattr(game_text_reader, '_old_automations_window'):
            old_instance = game_text_reader._old_automations_window
        elif hasattr(game_text_reader, '_automations_window') and game_text_reader._automations_window:
            # Fallback: check current instance (might be the old one if we're being created before assignment)
            old_instance = game_text_reader._automations_window
            # Only use it if it's not the same as self (which would be the case if we're being recreated)
            if old_instance is self:
                old_instance = None
        
        # Check if old instance has polling active and thread is still running
        if old_instance:
            if (hasattr(old_instance, 'polling_active') and old_instance.polling_active and
                hasattr(old_instance, 'polling_thread') and old_instance.polling_thread and
                old_instance.polling_thread.is_alive()):
                # Polling is actually running - transfer state
                self.polling_active = True
                self.polling_thread = old_instance.polling_thread
                game_text_reader._automations_polling_active = True
                print("AUTOMATION: Restored polling state - thread was still running")
            else:
                # Polling was not active or thread died, check stored state
                if hasattr(game_text_reader, '_automations_polling_active'):
                    self.polling_active = game_text_reader._automations_polling_active
                else:
                    self.polling_active = False
                    game_text_reader._automations_polling_active = False
                self.polling_thread = None
        else:
            # No old instance, check stored state in game_text_reader
            if hasattr(game_text_reader, '_automations_polling_active'):
                self.polling_active = game_text_reader._automations_polling_active
            else:
                self.polling_active = False
                game_text_reader._automations_polling_active = False
            self.polling_thread = None
        self.polling_interval = 0.1  # Check every 100ms
        
        # Track unsaved changes
        self._has_unsaved_changes = False
        self._initial_automations_state = None  # Snapshot of state when window opens or after save
        
        # Set up protocol to handle window closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Add remove_focus method - same as main window (simple and always works)
        def remove_focus(event):
            widget = event.widget
            # Only remove focus if clicking on something that's NOT an Entry
            # This allows Entry widgets to work normally
            if not isinstance(widget, tk.Entry):
                # Use after_idle to ensure Entry can receive focus if clicked
                def check_and_remove():
                    current_focus = self.window.focus_get()
                    # Only remove if an Entry has focus and we didn't click on an Entry
                    if isinstance(current_focus, tk.Entry) and not isinstance(widget, tk.Entry):
                        self.window.focus()
                self.window.after_idle(check_and_remove)
        
        # Bind click to remove focus from entry fields when clicking outside
        self.window.bind("<Button-1>", remove_focus)
        
        # Create UI
        self.create_ui()
    
    def _update_polling_button_state(self):
        """Update the polling button state to match the actual polling state"""
        try:
            # Use shared state in game_text_reader as source of truth (persists across window close/reopen)
            is_active = False
            if hasattr(self.game_text_reader, '_automations_polling_active'):
                is_active = self.game_text_reader._automations_polling_active
            
            # Sync local state with shared state
            self.polling_active = is_active
            
            # Update button based on shared state (source of truth)
            if is_active:
                self.polling_button.config(text="Stop Monitoring", bg="#FFB6C1")
                print(f"POLLING: Button updated to 'Stop Monitoring' (is_active={is_active})")
            else:
                self.polling_button.config(text="Start Monitor Detections", bg="#90EE90")
                print(f"POLLING: Button updated to 'Start Monitor Detections' (is_active={is_active})")
        except Exception as e:
            print(f"POLLING: Error updating button state: {e}")
            pass  # Button might not exist yet
    
    def _mark_unsaved_changes(self):
        """Mark that there are unsaved changes and auto-save if layout is loaded"""
        self._has_unsaved_changes = True
        
        # Auto-save if a layout is loaded
        if self.game_text_reader.layout_file.get():
            try:
                self.game_text_reader.save_layout_auto()
                # Reset flag after successful save
                self._has_unsaved_changes = False
            except Exception as e:
                print(f"Error auto-saving automations: {e}")
                # Keep flag set if save failed
    
    def _update_main_status(self, message, color="black", duration=3000):
        """Update the main window status label with a message"""
        if hasattr(self.game_text_reader, 'status_label'):
            try:
                self.game_text_reader.status_label.config(text=message, fg=color, font=("Helvetica", 10, "bold"))
                # Clear after duration
                if hasattr(self.game_text_reader, '_automation_status_timer'):
                    self.game_text_reader.root.after_cancel(self.game_text_reader._automation_status_timer)
                def clear_and_restore_monitoring():
                    self.game_text_reader.status_label.config(text="", font=("Helvetica", 10))
                    # After clearing, show monitoring status if active
                    if hasattr(self.game_text_reader, '_show_monitoring_status_if_active'):
                        self.game_text_reader._show_monitoring_status_if_active()
                self.game_text_reader._automation_status_timer = self.game_text_reader.root.after(duration, clear_and_restore_monitoring)
            except:
                pass  # Status label might not exist or be destroyed
    
    def on_close(self):
        """Handle window closing - check for unsaved changes, stop polling, clean up hotkeys, and destroy window"""
        # Check if there are automations or combos that need to be saved
        has_automations = len(self.automations) > 0 or len(self.hotkey_combos) > 0
        has_layout_file = bool(self.game_text_reader.layout_file.get())
        
        # If there are automations/combos but no layout file, prompt to create one
        if has_automations and not has_layout_file:
            response = messagebox.askyesnocancel(
                "Save Automations",
                "You have automations configured but no layout file is set.\n\n"
                "Automations will be lost if you close without saving.\n\n"
                "Create a layout file to save your automations?\n\n"
                "(Yes = Create Layout File, No = Close without Saving, Cancel = Don't Close)"
            )
            
            if response is None:  # Cancel - don't close
                return
            elif response:  # Yes - Create and save layout file
                try:
                    # Call save_layout which will prompt for file location
                    self.game_text_reader.save_layout()
                    # Reset unsaved changes flag after save
                    self._has_unsaved_changes = False
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save automations: {str(e)}")
                    # Don't close if save failed
                    return
            # If response is False (No), continue with closing without saving
        
        # Check for unsaved changes if a layout is loaded
        if self._has_unsaved_changes and has_layout_file:
            response = messagebox.askyesnocancel(
                "Unsaved Automation Changes",
                "You have unsaved changes to automations.\n\n"
                "Save changes before closing?\n\n"
                "(Yes = Save and Close, No = Close without Saving, Cancel = Don't Close)"
            )
            
            if response is None:  # Cancel - don't close
                return
            elif response:  # Yes - Save and close
                # Auto-save the layout
                try:
                    self.game_text_reader.save_layout_auto()
                    # Reset unsaved changes flag
                    self._has_unsaved_changes = False
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save automations: {str(e)}")
                    # Don't close if save failed
                    return
            # If response is False (No), continue with closing without saving
        
        # DON'T stop polling - monitoring should continue even when window is closed
        # self.stop_polling()
        
        # Cancel any pending after() calls that might try to update widgets
        try:
            # Cancel any pending status updates
            for automation in self.automations:
                if automation.get('_status_update_id'):
                    try:
                        self.root.after_cancel(automation['_status_update_id'])
                    except:
                        pass
        except:
            pass
        
        # Stop any active triggering for combos, but DON'T clean up hotkeys - they should persist globally
        # Hotkeys are registered globally via keyboard.add_hotkey() and should continue working
        # even when the window is closed
        print(f"HOTKEY COMBO DEBUG: Window closing - checking registry state...")
        print(f"HOTKEY COMBO DEBUG: Registry has {len(self.combo_callbacks_by_hotkey)} combo callback(s): {list(self.combo_callbacks_by_hotkey.keys())}")
        print(f"HOTKEY COMBO DEBUG: Registry has {len(self.automation_callbacks_by_hotkey)} automation callback(s): {list(self.automation_callbacks_by_hotkey.keys())}")
        
        for combo in self.hotkey_combos[:]:  # Use slice copy to avoid modification during iteration
            # Stop any active triggering
            combo['is_triggering'] = False
            # NOTE: We intentionally do NOT clean up keyboard hooks here
            # The hotkeys are registered globally and should persist when the window closes
            # The callbacks are stored in combo_callbacks_by_hotkey registry for reliable lookup
            hotkey_button = combo.get('hotkey_button')
            hotkey_name = 'N/A'
            if hotkey_button and hasattr(hotkey_button, 'hotkey'):
                hotkey_name = hotkey_button.hotkey
            print(f"HOTKEY COMBO DEBUG: Keeping hotkey '{hotkey_name}' active (window closing but hotkey persists)")
            # Verify callback is in registry
            if hotkey_name in self.combo_callbacks_by_hotkey:
                print(f"HOTKEY COMBO DEBUG: ✓ Callback for '{hotkey_name}' confirmed in registry")
            else:
                print(f"HOTKEY COMBO DEBUG: ✗ WARNING - Callback for '{hotkey_name}' NOT in registry!")
        
        print(f"HOTKEY COMBO DEBUG: Window being destroyed, but registries will persist in game_text_reader._automations_window")
        self.window.destroy()
    
    def create_ui(self):
        """Create the user interface"""
        # Top frame container
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill='x', padx=10, pady=10)
        
        # Line 1: Left side buttons, Right side Start Monitoring
        line1_frame = tk.Frame(top_frame)
        line1_frame.pack(fill='x', pady=2)
        
        # Left side of line 1
        add_automation_button = tk.Button(
            line1_frame, 
            text="🖼 Add Detection Area", 
            command=self.add_automation,
            font=("Helvetica", 10)
        )
        add_automation_button.pack(side='left')
        
        # Set Image Area button (at top level)
        set_image_area_button = tk.Button(
            line1_frame,
            text="Set a detection area",
            command=self.set_image_area_top_level,
            font=("Helvetica", 10)
        )
        set_image_area_button.pack(side='left', padx=(10, 0))
        # Store reference to prevent accidental triggering
        self.set_image_area_button = set_image_area_button
        
        # Freeze Screen checkbox (at top level)
        self.freeze_screen_var = tk.BooleanVar(value=False)
        # Track changes to freeze screen checkbox to trigger auto-save
        def on_freeze_screen_change(*args):
            self._mark_unsaved_changes()
        self.freeze_screen_var.trace('w', on_freeze_screen_change)
        freeze_screen_checkbox = tk.Checkbutton(
            line1_frame,
            text="Freeze Screen",
            variable=self.freeze_screen_var,
            font=("Helvetica", 10)
        )
        freeze_screen_checkbox.pack(side='left', padx=(10, 0))
        
        # Set Hotkey button (at top level)
        self.set_hotkey_button = tk.Button(
            line1_frame,
            text="Set Hotkey: [ None ]",
            command=self.set_hotkey_top_level,
            font=("Helvetica", 10)
        )
        self.set_hotkey_button.pack(side='left', padx=(10, 0))
        
        # Right side of line 1: Start/Stop monitoring button
        # Initial state will be set by _update_polling_button_state after UI creation
        self.polling_button = tk.Button(
            line1_frame,
            text="Start Monitor Detections",
            command=self.toggle_polling,
            font=("Helvetica", 10),
            bg="#90EE90"
        )
        self.polling_button.pack(side='right', padx=10)
        
        # Line 2: Add Area Combo button
        line2_frame = tk.Frame(top_frame)
        line2_frame.pack(fill='x', pady=2)
        
        add_hotkey_combo_button = tk.Button(
            line2_frame,
            text="➕ Add Area Combo",
            command=self.add_hotkey_combo,
            font=("Helvetica", 10)
        )
        add_hotkey_combo_button.pack(side='left')
        
        # Separator
        ttk.Separator(self.window, orient='horizontal').pack(fill='x', padx=10, pady=5)
        
        # Scrollable frame for automations
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
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mouse wheel scrolling anywhere in the window
        def _on_mousewheel(event):
            # Check if the canvas is scrollable
            if canvas.bbox("all") and canvas.winfo_height() < (canvas.bbox("all")[3] - canvas.bbox("all")[1]):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        # Bind mouse wheel to the entire window so it works anywhere
        def bind_mousewheel(event):
            self.window.bind_all("<MouseWheel>", _on_mousewheel)
        
        def unbind_mousewheel(event):
            self.window.unbind_all("<MouseWheel>")
        
        # Bind when mouse enters the window
        self.window.bind("<Enter>", bind_mousewheel)
        # Keep it bound - don't unbind when leaving, so it works everywhere
        # Also bind directly to canvas and scrollable frame
        canvas.bind("<MouseWheel>", _on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        # Bind to the entire window
        self.window.bind("<MouseWheel>", _on_mousewheel)
        
        self.canvas = canvas
        self.scrollable_frame_ref = self.scrollable_frame
        
        # Update polling button state to reflect actual monitoring status
        self._update_polling_button_state()
    
    def get_available_trigger_options(self, exclude_automation=None, exclude_combo=None):
        """Get all available trigger options: areas, automations, and area combos
        
        Args:
            exclude_automation: Automation to exclude from the list (to prevent self-reference)
            exclude_combo: Area combo to exclude from the list (to prevent self-reference)
        
        Returns:
            List of option names with separators
        """
        options = []
        
        # Add regular read areas (including individual Auto Read 1, Auto Read 2, etc.)
        area_names = []
        for area in self.game_text_reader.areas:
            area_name = area[3].get()  # area_name_var
            if area_name:
                area_names.append(area_name)
        
        if area_names:
            options.append("─── Areas ───")
            options.extend(area_names)
        
        # Add other automations (excluding the current one if specified) with separator
        automation_names = []
        for automation in self.automations:
            if exclude_automation is None or automation['id'] != exclude_automation['id']:
                automation_names.append(automation['name'])
        
        if automation_names:
            options.append("─── Automations ───")
            options.extend(automation_names)
        
        # Add area combos (excluding the current one if specified) with separator
        combo_names = []
        for combo in self.hotkey_combos:
            if exclude_combo is None or combo['id'] != exclude_combo['id']:
                combo_names.append(combo['name'])
        
        if combo_names:
            options.append("─── Area Combos ───")
            options.extend(combo_names)
        
        return options if options else ["No options available"]
    
    def add_automation(self):
        """Add a new automation rule"""
        automation_id = len(self.automations)
        # Name automations as "Detection Area A", "Detection Area B", etc.
        automation_letter = chr(65 + automation_id)  # 65 is 'A' in ASCII
        automation_name = f"Detection Area: {automation_letter}"
        
        automation = {
            'id': automation_id,
            'name': automation_name,
            'image_area_coords': None,  # (x1, y1, x2, y2)
            'reference_image': None,  # PIL Image
            'hotkey': None,  # Hotkey string
            'hotkey_button': None,  # Button widget for hotkey (not used in UI anymore)
            'match_percent': tk.DoubleVar(value=80.0),  # Default 80%
            'comparison_method': tk.StringVar(value="SSIM"),  # Comparison method: "Pixel", "Histogram", "SSIM", "Perceptual"
            'target_read_area': tk.StringVar(value=""),  # Selected read area name
            'only_read_if_text': tk.BooleanVar(value=False),
            'read_after_ms': tk.IntVar(value=0),  # Timer in milliseconds
            'timer_active': False,  # Whether countdown is active
            'timer_start_time': None,  # When timer started
            'was_matching': False,  # Previous match state (for toggle behavior)
            'has_triggered': False,  # Whether we've triggered for current match state
            'text_last_found_time': None,  # When text was last detected (for debouncing OCR misses)
            'frame': None  # UI frame for this automation
        }
        
        self.automations.append(automation)
        self.create_automation_ui(automation)
        self._mark_unsaved_changes()
    
    def create_automation_ui(self, automation):
        """Create UI elements for a single automation rule"""
        # Main frame for this automation
        automation_frame = tk.Frame(self.scrollable_frame_ref, relief='ridge', bd=2, padx=10, pady=10)
        automation_frame.pack(fill='x', padx=5, pady=5)
        automation['frame'] = automation_frame
        
        # Top row: Automation name, Freeze Screen checkbox, Remove button
        top_row = tk.Frame(automation_frame)
        top_row.pack(fill='x', pady=5)
        
        # Automation name label
        name_label = tk.Label(
            top_row,
            text=automation['name'],
            font=("Helvetica", 10, "bold")
        )
        name_label.pack(side='left', padx=5)
        automation['name_label'] = name_label
        
        # Help text: "Start Monitoring Areas must be on for it to trigger"
        help_text_label = tk.Label(
            top_row,
            text="(Start Monitoring Areas must be on for it to trigger)",
            font=("Helvetica", 8),
            fg="gray"
        )
        help_text_label.pack(side='left', padx=5)
        automation['help_text_label'] = help_text_label
        
        # Show hotkey if set
        hotkey_label = tk.Label(
            top_row,
            text="",
            font=("Helvetica", 9),
            fg="gray"
        )
        hotkey_label.pack(side='left', padx=5)
        automation['hotkey_label'] = hotkey_label
        self.update_hotkey_display(automation)
        
        # Remove button
        remove_button = tk.Button(
            top_row,
            text="❌ Remove",
            command=lambda: self.remove_automation(automation),
            font=("Helvetica", 9),
            fg="red"
        )
        remove_button.pack(side='right', padx=5)
        
        # Line 2: Method and "Only read if text exists" checkbox
        line2_row = tk.Frame(automation_frame)
        line2_row.pack(fill='x', pady=2)
        
        # Comparison method selection
        method_label = tk.Label(line2_row, text="Method:", font=("Helvetica", 9))
        method_label.pack(side='left', padx=5)
        
        method_dropdown = tk.OptionMenu(
            line2_row,
            automation['comparison_method'],
            "Pixel",      # Pixel-by-pixel (strict, exact colors)
            "Histogram",  # Color histogram (forgiving to shifts)
            "SSIM",       # Structural similarity (best for games)
            "Perceptual", # Perceptual hash (very forgiving)
            "Edge"        # Edge detection (ignores colors, detects shapes)
        )
        method_dropdown.config(width=12, font=("Helvetica", 8))
        method_dropdown.pack(side='left', padx=5)
        
        # Windows-style info icon with tooltip (circular blue button with "i")
        info_canvas = tk.Canvas(
            line2_row,
            width=16,
            height=16,
            highlightthickness=0,
            cursor="hand2"
        )
        info_canvas.pack(side='left', padx=(5, 0))
        
        # Draw Windows-style info icon: blue circle with white "i"
        info_canvas.create_oval(2, 2, 14, 14, fill='#0078D4', outline='#005A9E', width=1)
        info_canvas.create_text(8, 8, text="i", font=("Helvetica", 9, "bold"), fill="white")
        
        # Tooltip text explaining each method
        tooltip_text = (
            "Comparison Methods:\n\n"
            "• Pixel: Strict pixel-by-pixel comparison. Requires exact colors.\n\n"
            "• Histogram: Compares color distribution. Forgiving to slight shifts.\n\n"
            "• SSIM: Structural Similarity Index. Best for games, handles lighting changes.\n\n"
            "• Perceptual: Perceptual hash. Very forgiving, ignores minor differences.\n\n"
            "• Edge: Edge detection. Ignores colors, detects shapes and outlines."
        )
        ToolTip(info_canvas, tooltip_text)
        
        # "Only read if text exists" checkbox
        text_checkbox = tk.Checkbutton(
            line2_row,
            text="Only trigger if detection area is a match and there is set text in the triggered area",
            variable=automation['only_read_if_text'],
            font=("Helvetica", 9)
        )
        text_checkbox.pack(side='left', padx=10)
        
        # Line 3: Image preview, Match %, Timer, Read Area dropdown
        line3_row = tk.Frame(automation_frame)
        line3_row.pack(fill='x', pady=2)
        
        # Image preview (40x40px)
        preview_frame = tk.Frame(line3_row, width=40, height=40, relief='sunken', bd=1)
        preview_frame.pack(side='left', padx=5)
        preview_frame.pack_propagate(False)
        
        preview_label = tk.Label(preview_frame, text="No image", font=("Helvetica", 6))
        preview_label.pack(expand=True)
        automation['preview_label'] = preview_label
        
        # Show status if image is set
        status_label = tk.Label(line3_row, text="", font=("Helvetica", 8), fg="gray")
        status_label.pack(side='left', padx=5)
        automation['status_label'] = status_label
        
        # Match % slider
        match_frame = tk.Frame(line3_row)
        match_frame.pack(side='left', padx=5)
        
        tk.Label(match_frame, text="Match %:", font=("Helvetica", 9)).pack(side='left')
        match_slider = tk.Scale(
            match_frame,
            from_=5,
            to=100,
            orient='horizontal',
            variable=automation['match_percent'],
            length=150,
            font=("Helvetica", 8)
        )
        match_slider.pack(side='left', padx=5)
        match_value_label = tk.Label(match_frame, text="80.0%", font=("Helvetica", 8))
        match_value_label.pack(side='left', padx=5)
        automation['match_value_label'] = match_value_label
        
        # Update label when slider changes
        def update_match_label(*args):
            match_value_label.config(text=f"{automation['match_percent'].get():.1f}%")
            self._mark_unsaved_changes()
        automation['match_percent'].trace('w', update_match_label)
        
        # Track changes to comparison method
        def on_comparison_method_change(*args):
            self._mark_unsaved_changes()
        automation['comparison_method'].trace('w', on_comparison_method_change)
        
        # Track changes to target read area
        def on_target_read_area_change(*args):
            self._mark_unsaved_changes()
        automation['target_read_area'].trace('w', on_target_read_area_change)
        
        # Track changes to only_read_if_text
        def on_only_read_if_text_change(*args):
            self._mark_unsaved_changes()
        automation['only_read_if_text'].trace('w', on_only_read_if_text_change)
        
        # Track changes to read_after_ms
        def on_read_after_ms_change(*args):
            self._mark_unsaved_changes()
        automation['read_after_ms'].trace('w', on_read_after_ms_change)
        
        # Timer input
        timer_frame = tk.Frame(line3_row)
        timer_frame.pack(side='left', padx=5)
        
        tk.Label(timer_frame, text="Trigger after", font=("Helvetica", 9)).pack(side='left')
        timer_entry = tk.Entry(timer_frame, textvariable=automation['read_after_ms'], width=6, font=("Helvetica", 9))
        timer_entry.pack(side='left', padx=2)
        tk.Label(timer_frame, text="ms", font=("Helvetica", 9)).pack(side='left')
        
        # Store reference to timer entry for focus removal
        automation['timer_entry'] = timer_entry
        
        # Ensure entry field always stays enabled and can receive focus
        # This prevents it from becoming uneditable after image updates or combo triggers
        def ensure_entry_enabled():
            """Ensure entry field is always enabled and can receive focus"""
            try:
                if timer_entry.winfo_exists():
                    timer_entry.config(state='normal')
            except:
                pass
        
        # Re-enable after any window updates - bind to multiple events for robustness
        timer_entry.bind("<FocusIn>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<Button-1>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<Enter>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<KeyPress>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<KeyRelease>", lambda e: ensure_entry_enabled())
        
        # Periodic check to ensure entry stays enabled (every 500ms)
        def periodic_enable_check():
            try:
                if self.window.winfo_exists() and timer_entry.winfo_exists():
                    ensure_entry_enabled()
                    self.window.after(500, periodic_enable_check)
            except:
                pass
        
        # Start periodic check
        self.window.after(500, periodic_enable_check)
        
        # Read Area dropdown
        read_area_frame = tk.Frame(line3_row)
        read_area_frame.pack(side='left', padx=5)
        
        tk.Label(read_area_frame, text="Trigger Area:", font=("Helvetica", 9)).pack(side='left')
        
        # Create combobox that refreshes when opened
        def refresh_automation_dropdown():
            """Refresh the dropdown options when opened"""
            trigger_options = self.get_available_trigger_options(exclude_automation=automation)
            if not trigger_options or trigger_options == ["No options available"]:
                trigger_options = ["No options available"]
            read_area_dropdown['values'] = trigger_options
            # Preserve current selection if it's still valid
            current_value = automation['target_read_area'].get()
            if current_value not in trigger_options and current_value:
                automation['target_read_area'].set("")
        
        # Get initial list of available trigger options
        trigger_options = self.get_available_trigger_options(exclude_automation=automation)
        if not trigger_options or trigger_options == ["No options available"]:
            trigger_options = ["No options available"]
            automation['target_read_area'].set("No options available")
        
        read_area_dropdown = ttk.Combobox(
            read_area_frame,
            textvariable=automation['target_read_area'],
            values=trigger_options,
            state='readonly',
            width=18,
            font=("Helvetica", 8),
            postcommand=refresh_automation_dropdown
        )
        read_area_dropdown.pack(side='left', padx=5)
        
        # Status indicators row: Found Image, Found Text, Timer Progress
        status_row = tk.Frame(automation_frame)
        status_row.pack(fill='x', pady=5)
        
        # Found Image indicator
        image_status_frame = tk.Frame(status_row)
        image_status_frame.pack(side='left', padx=10)
        
        # Circle for image status (gray by default when monitoring not active)
        image_circle = tk.Canvas(image_status_frame, width=12, height=12, highlightthickness=0)
        image_circle.pack(side='left', padx=2)
        initial_color = "gray" if not self.polling_active else "red"
        image_circle.create_oval(2, 2, 10, 10, fill=initial_color, outline="black", width=1)
        automation['image_status_circle'] = image_circle
        
        tk.Label(image_status_frame, text="Image", font=("Helvetica", 8)).pack(side='left')
        automation['image_status_label'] = tk.Label(image_status_frame, text="", font=("Helvetica", 8))
        automation['image_status_label'].pack(side='left', padx=2)
        
        # Found Text indicator
        text_status_frame = tk.Frame(status_row)
        text_status_frame.pack(side='left', padx=10)
        
        # Circle for text status (gray by default when monitoring not active)
        text_circle = tk.Canvas(text_status_frame, width=12, height=12, highlightthickness=0)
        text_circle.pack(side='left', padx=2)
        initial_color = "gray" if not self.polling_active else "red"
        text_circle.create_oval(2, 2, 10, 10, fill=initial_color, outline="black", width=1)
        automation['text_status_circle'] = text_circle
        
        tk.Label(text_status_frame, text="Text", font=("Helvetica", 8)).pack(side='left')
        automation['text_status_label'] = tk.Label(text_status_frame, text="", font=("Helvetica", 8))
        automation['text_status_label'].pack(side='left', padx=2)
        
        # Timer progress bar
        timer_progress_frame = tk.Frame(status_row)
        timer_progress_frame.pack(side='left', padx=10)
        
        tk.Label(timer_progress_frame, text="Timer:", font=("Helvetica", 8)).pack(side='left', padx=2)
        
        # Progress bar for timer
        progress_bar = ttk.Progressbar(
            timer_progress_frame,
            mode='determinate',
            length=150,
            maximum=100
        )
        progress_bar.pack(side='left', padx=5)
        automation['timer_progress_bar'] = progress_bar
        
        automation['timer_progress_label'] = tk.Label(timer_progress_frame, text="0ms", font=("Helvetica", 8))
        automation['timer_progress_label'].pack(side='left', padx=2)
        
        # Initialize status indicators
        self.update_automation_status(automation, False, False, 0, 0)
    
    def set_image_area_top_level(self):
        """Set image area from top-level button - immediately starts area selection"""
        print("=" * 60)
        print("AUTOMATION: set_image_area_top_level() called")
        
        if not self.automations:
            print("AUTOMATION: No automations found, showing info message")
            messagebox.showinfo("No Automations", "Please add an automation first.")
            return
        
        # Check if area selection is already in progress
        if hasattr(self.game_text_reader, 'area_selection_in_progress') and self.game_text_reader.area_selection_in_progress:
            print("AUTOMATION: Area selection already in progress")
            messagebox.showwarning("Area Selection Active", "Please complete or cancel the current area selection first.")
            return
        
        # Immediately start area selection - dialog will show AFTER image is captured
        print("AUTOMATION: Starting area selection immediately...")
        self.start_area_selection_for_automations()
        print("=" * 60)
    
    def set_hotkey_top_level(self):
        """Set hotkey from top-level button - assigns a hotkey that triggers area selection"""
        print("AUTOMATION: set_hotkey_top_level() called")
        
        # If a hotkey is already set, clear it first to allow setting a new one
        if hasattr(self.set_hotkey_button, 'hotkey') and self.set_hotkey_button.hotkey:
            print("AUTOMATION: Clearing existing hotkey before setting new one")
            # Clear the old hotkey
            old_hotkey = self.set_hotkey_button.hotkey
            if hasattr(self.game_text_reader, 'hotkeys') and old_hotkey in self.game_text_reader.hotkeys:
                try:
                    self.game_text_reader.hotkeys[old_hotkey].unhook()
                    del self.game_text_reader.hotkeys[old_hotkey]
                except:
                    pass
            self.set_hotkey_button.hotkey = None
            self.set_hotkey_button.config(text="Set Hotkey: [ None ]")
        
        # Ensure we're not in the middle of area selection
        if hasattr(self.game_text_reader, 'area_selection_in_progress') and self.game_text_reader.area_selection_in_progress:
            print("AUTOMATION: Area selection in progress, cannot set hotkey")
            messagebox.showwarning("Area Selection Active", "Please complete or cancel the area selection first.")
            return
        
        # Create a temporary frame for compatibility with hotkey system
        temp_frame = tk.Frame()
        temp_frame._is_automation_area_hotkey = True
        temp_frame._automation_window = self  # Store reference to automations window
        
        # Store callback for when hotkey is pressed - triggers area selection
        # Make it a method that can be restored if lost
        def hotkey_callback():
            # When hotkey is pressed, trigger area selection
            print("=" * 60)
            print("AUTOMATION: Area selection hotkey pressed, triggering area selection")
            print(f"AUTOMATION: Button: {self.set_hotkey_button}")
            print(f"AUTOMATION: Has callback attr: {hasattr(self.set_hotkey_button, '_automation_callback')}")
            if hasattr(self.set_hotkey_button, '_automation_callback'):
                print(f"AUTOMATION: Callback value: {self.set_hotkey_button._automation_callback}")
            # Check if area selection is already in progress
            if hasattr(self.game_text_reader, 'area_selection_in_progress') and self.game_text_reader.area_selection_in_progress:
                print("AUTOMATION: Area selection already in progress, ignoring hotkey")
                return
            self.start_area_selection_for_automations()
            # Re-set callback after use to ensure it persists
            # This is critical - the callback must persist after being called
            if hasattr(self, 'set_hotkey_button') and self.set_hotkey_button:
                self.set_hotkey_button._automation_callback = hotkey_callback
                print("AUTOMATION: Callback re-set after execution")
            print("=" * 60)
        
        # Store callback on button IMMEDIATELY so it's available when setup_hotkey creates the handler
        # Also store it on the window as a backup in case button reference changes
        print(f"AUTOMATION: Setting callback on button: {self.set_hotkey_button}")
        self.set_hotkey_button._automation_callback = hotkey_callback
        self.set_hotkey_button._automation_temp_frame = temp_frame
        # Store as backup on window
        self._area_selection_hotkey_callback = hotkey_callback
        self._area_selection_hotkey_button = self.set_hotkey_button
        print(f"AUTOMATION: Callback set: {hasattr(self.set_hotkey_button, '_automation_callback')}")
        print(f"AUTOMATION: Callback value: {self.set_hotkey_button._automation_callback}")
        
        # Use existing hotkey system to start hotkey assignment
        print("AUTOMATION: Starting hotkey assignment mode for area selection...")
        self.game_text_reader.set_hotkey(self.set_hotkey_button, temp_frame)
        
        # After hotkey is assigned, ensure callback is set and update the display
        # setup_hotkey is called immediately after hotkey assignment in _finalize_hotkey
        # We need to ensure the callback is still there after setup_hotkey completes
        def update_after_hotkey_set():
            """Update button display after hotkey is set and ensure callback is preserved"""
            if hasattr(self.set_hotkey_button, 'hotkey') and self.set_hotkey_button.hotkey:
                print(f"AUTOMATION: Area selection hotkey assigned: {self.set_hotkey_button.hotkey}")
                
                # CRITICAL: Re-set the callback after setup_hotkey completes
                # The handler in setup_hotkey checks for this callback at runtime, so it must exist
                print("AUTOMATION: Re-setting callback after hotkey setup completes")
                self.set_hotkey_button._automation_callback = hotkey_callback
                # Also update backup
                self._area_selection_hotkey_callback = hotkey_callback
                self._area_selection_hotkey_button = self.set_hotkey_button
                
                # Verify it's set
                if hasattr(self.set_hotkey_button, '_automation_callback') and self.set_hotkey_button._automation_callback:
                    print("AUTOMATION: ✓ Callback is set and ready")
                    print(f"AUTOMATION: Callback function: {self.set_hotkey_button._automation_callback}")
                else:
                    print("AUTOMATION: ✗ ERROR - Callback is NOT set!")
                    print(f"AUTOMATION: Button: {self.set_hotkey_button}")
                    print(f"AUTOMATION: Has attr: {hasattr(self.set_hotkey_button, '_automation_callback')}")
                    if hasattr(self.set_hotkey_button, '_automation_callback'):
                        print(f"AUTOMATION: Callback value: {self.set_hotkey_button._automation_callback}")
                
                # Also ensure the temp_frame reference is preserved
                if not hasattr(self.set_hotkey_button, '_automation_temp_frame'):
                    self.set_hotkey_button._automation_temp_frame = temp_frame
                
                display_name = self.set_hotkey_button.hotkey.replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                self.set_hotkey_button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                
                # Mark as unsaved to trigger auto-save
                self._mark_unsaved_changes()
                
                # Set up periodic callback restoration to ensure it never gets lost
                def ensure_callback_persists():
                    """Periodically ensure callback is set - prevents it from being lost"""
                    if hasattr(self, 'set_hotkey_button') and self.set_hotkey_button:
                        if hasattr(self.set_hotkey_button, 'hotkey') and self.set_hotkey_button.hotkey:
                            # Check if callback is missing or None
                            if not hasattr(self.set_hotkey_button, '_automation_callback') or not self.set_hotkey_button._automation_callback:
                                # Restore from backup
                                if hasattr(self, '_area_selection_hotkey_callback'):
                                    print("AUTOMATION: Restoring lost callback from backup")
                                    self.set_hotkey_button._automation_callback = self._area_selection_hotkey_callback
                            # Schedule next check
                            self.window.after(2000, ensure_callback_persists)
                
                # Start periodic checking
                self.window.after(2000, ensure_callback_persists)
            else:
                # Still waiting for assignment, check again
                self.window.after(100, update_after_hotkey_set)
        
        # Start checking after a delay to allow setup_hotkey to complete
        # setup_hotkey is called in _finalize_hotkey which happens ~250ms after key press
        # So we check after 500ms to be safe
        self.window.after(500, update_after_hotkey_set)
        
        # Also set it again after a longer delay to be absolutely sure
        self.window.after(1000, lambda: setattr(self.set_hotkey_button, '_automation_callback', hotkey_callback) if hasattr(self.set_hotkey_button, 'hotkey') and self.set_hotkey_button.hotkey else None)
    
    def start_area_selection_for_automations(self):
        """Start area selection - after capture, show dialog to select which automation(s) to assign image to"""
        print("=" * 60)
        print("AUTOMATION: start_area_selection_for_automations() called")
        
        # Check if window exists - if not, just return (don't open window)
        window_exists = False
        try:
            window_exists = self.window.winfo_exists()
        except:
            pass
        
        if not window_exists:
            print("AUTOMATION: Window is closed, cannot start area selection")
            return
        
        print(f"AUTOMATION: Freeze screen setting: {self.freeze_screen_var.get()}")
        
        # Create a temporary frame to store coordinates
        temp_frame = tk.Frame()
        print(f"AUTOMATION: Created temp_frame: {temp_frame}")
        
        # Create callback to handle area selection completion
        def on_area_selected(frame):
            print("=" * 60)
            print(f"AUTOMATION: on_area_selected() callback triggered")
            print(f"AUTOMATION: Frame object: {frame}")
            print(f"AUTOMATION: Frame has area_coords attr: {hasattr(frame, 'area_coords')}")
            
            try:
                if hasattr(frame, 'area_coords') and frame.area_coords:
                    coords = frame.area_coords
                    print(f"AUTOMATION: Area coordinates found: {coords}")
                    x1, y1, x2, y2 = coords
                    
                    # Capture reference image first
                    print(f"AUTOMATION: Capturing reference image at ({x1}, {y1}, {x2}, {y2})")
                    freeze_screen = self.freeze_screen_var.get()
                    print(f"AUTOMATION: Freeze screen for capture: {freeze_screen}")
                    captured_image = self.capture_image_for_automations(x1, y1, x2, y2, freeze_screen, frame)
                    
                    if captured_image:
                        # Now show dialog to select which automation(s) to assign this image to
                        # Add a delay to ensure area selection window is fully closed first
                        # Use root window for scheduling to avoid issues if automation window is closed
                        print("AUTOMATION: Showing automation selection dialog...")
                        self.root.after(200, lambda: self.show_automation_selection_dialog(captured_image, coords))
                    else:
                        print("AUTOMATION: Failed to capture image")
                else:
                    print(f"AUTOMATION: WARNING - No area coordinates found in frame")
                    if hasattr(frame, 'area_coords'):
                        print(f"AUTOMATION: Frame.area_coords exists but is: {frame.area_coords}")
                    else:
                        print("AUTOMATION: Frame does not have area_coords attribute")
            except Exception as e:
                print(f"AUTOMATION: ERROR in on_area_selected callback: {e}")
                import traceback
                traceback.print_exc()
            finally:
                # Don't restore focus immediately - let the dialog show first
                # Focus will be restored after dialog is closed
                pass
            print("=" * 60)
        
        # Store the callback in the frame for later use
        temp_frame._automation_callback = on_area_selected
        print(f"AUTOMATION: Stored callback in temp_frame")
        
        # Use the helper method from game_text_reader with top-level freeze screen setting
        try:
            print(f"AUTOMATION: Calling game_text_reader.set_area_for_automation()...")
            print(f"AUTOMATION: - temp_frame: {temp_frame}")
            print(f"AUTOMATION: - callback: {on_area_selected}")
            print(f"AUTOMATION: - freeze_screen: {self.freeze_screen_var.get()}")
            
            self.game_text_reader.set_area_for_automation(temp_frame, on_area_selected, self.freeze_screen_var.get())
            
            print(f"AUTOMATION: set_area_for_automation() returned successfully")
            print(f"AUTOMATION: Area selection should start now - window should appear")
        except Exception as e:
            print(f"AUTOMATION: ERROR starting area selection: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to start area selection: {e}")
        print("=" * 60)
    
    def capture_image_for_automations(self, x1, y1, x2, y2, freeze_screen, frame=None):
        """Capture image for automation assignment - returns the captured image
        If freeze_screen is True and frame has frozen_screenshot, extracts from frozen screenshot"""
        try:
            # Check if we should use frozen screenshot
            if freeze_screen and frame and hasattr(frame, 'frozen_screenshot') and frame.frozen_screenshot is not None:
                print(f"AUTOMATION: Using frozen screenshot for capture")
                frozen_img = frame.frozen_screenshot
                
                # Get frozen screenshot bounds
                if hasattr(frame, 'frozen_screenshot_bounds'):
                    frozen_min_x, frozen_min_y, frozen_width, frozen_height = frame.frozen_screenshot_bounds
                else:
                    # Fallback: assume full screenshot bounds
                    frozen_width, frozen_height = frozen_img.size
                    frozen_min_x, frozen_min_y = 0, 0
                
                # Convert screen coordinates to coordinates relative to the frozen screenshot
                rel_x1 = x1 - frozen_min_x
                rel_y1 = y1 - frozen_min_y
                rel_x2 = x2 - frozen_min_x
                rel_y2 = y2 - frozen_min_y
                
                # Ensure coordinates are within bounds
                rel_x1 = max(0, min(frozen_width, rel_x1))
                rel_y1 = max(0, min(frozen_height, rel_y1))
                rel_x2 = max(0, min(frozen_width, rel_x2))
                rel_y2 = max(0, min(frozen_height, rel_y2))
                
                # Ensure valid area
                crop_x1, crop_x2 = min(rel_x1, rel_x2), max(rel_x1, rel_x2)
                crop_y1, crop_y2 = min(rel_y1, rel_y2), max(rel_y1, rel_y2)
                
                # Extract the region from the frozen screenshot
                if crop_x2 > crop_x1 and crop_y2 > crop_y1:
                    image = frozen_img.crop((crop_x1, crop_y1, crop_x2, crop_y2))
                    print(f"AUTOMATION: Extracted region from frozen screenshot: ({crop_x1}, {crop_y1}, {crop_x2}, {crop_y2})")
                    return image.copy()
                else:
                    print(f"AUTOMATION: Invalid crop area from frozen screenshot, falling back to live capture")
                    # Fall through to live capture
            
            # Default: capture from live screen
            image = capture_screen_area(x1, y1, x2, y2)
            print(f"AUTOMATION: Image captured from live screen: {image.size}")
            return image.copy()  # Return a copy
        except Exception as e:
            print(f"AUTOMATION: Error capturing image: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to capture image: {e}")
            return None
    
    def show_automation_selection_dialog(self, captured_image, coords):
        """Show dialog to select which automation(s) to assign the captured image to"""
        print("=" * 60)
        print("AUTOMATION: show_automation_selection_dialog() called")
        print(f"AUTOMATION: Number of automations: {len(self.automations)}")
        
        # Check if window still exists - if not, reopen it
        window_exists = False
        try:
            window_exists = self.window.winfo_exists()
        except:
            pass
        
        # Get the current window instance (might be different if window was reopened)
        current_window_instance = self
        if not window_exists:
            print("AUTOMATION: Window no longer exists, reopening it...")
            # Reopen the automation window
            self.game_text_reader.open_automations_window()
            # Wait a bit for window to be created
            self.root.update()
            # Get reference to the current window instance
            if hasattr(self.game_text_reader, '_automations_window') and self.game_text_reader._automations_window:
                current_window_instance = self.game_text_reader._automations_window
                print("AUTOMATION: Using current window instance")
        
        # Use the current window instance for all operations
        window_instance = current_window_instance
        
        # Ensure the automation window is visible and focused first
        try:
            window_instance.window.update_idletasks()
            window_instance.window.lift()
            window_instance.window.focus_force()
        except:
            pass
        
        # Create selection dialog
        try:
            selection_window = tk.Toplevel(window_instance.window)
        except Exception as e:
            print(f"AUTOMATION: Error creating selection dialog: {e}")
            return
        selection_window.title("Select Automation(s)")
        # Make window taller to accommodate multiple checkboxes
        selection_window.geometry("300x300")
        selection_window.transient(window_instance.window)
        selection_window.grab_set()
        
        # Ensure dialog appears on top
        selection_window.lift()
        selection_window.focus_force()
        selection_window.attributes("-topmost", True)
        
        # Center the dialog
        selection_window.update_idletasks()
        x = window_instance.window.winfo_x() + (window_instance.window.winfo_width() // 2) - 150
        y = window_instance.window.winfo_y() + (window_instance.window.winfo_height() // 2) - 150
        selection_window.geometry(f"300x300+{x}+{y}")
        
        # Remove topmost after positioning (so it doesn't stay on top forever)
        selection_window.after(100, lambda: selection_window.attributes("-topmost", False))
        
        tk.Label(selection_window, text="Select which automation(s) to assign this image to:", 
                font=("Helvetica", 10)).pack(pady=10)
        
        selected_automations = []  # List to store selected automations
        create_new_automation = False  # Flag for "New automation area" checkbox
        
        # Create checkboxes for each automation
        checkbox_vars = {}  # Dictionary to store checkbox variables
        print(f"AUTOMATION: Creating checkboxes for {len(self.automations)} automations")
        
        # Create a frame with scrollbar if needed
        checkbox_frame = tk.Frame(selection_window)
        checkbox_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Add "New automation area" checkbox at the top
        new_automation_var = tk.BooleanVar(value=False)
        new_automation_cb = tk.Checkbutton(
            checkbox_frame,
            text="New automation area",
            variable=new_automation_var,
            font=("Helvetica", 9, "bold")
        )
        new_automation_cb.pack(anchor='w', padx=20, pady=5)
        
        # Add separator if there are existing automations
        if window_instance.automations:
            separator = tk.Frame(checkbox_frame, height=2, bg="gray")
            separator.pack(fill='x', padx=20, pady=5)
        
        # Add checkboxes for existing automations
        for automation in window_instance.automations:
            print(f"AUTOMATION: Adding checkbox for {automation['name']}")
            var = tk.BooleanVar(value=False)  # Default to unchecked
            checkbox_vars[automation['id']] = var
            cb = tk.Checkbutton(
                checkbox_frame,
                text=automation['name'],
                variable=var,
                font=("Helvetica", 9)
            )
            cb.pack(anchor='w', padx=20, pady=2)
        
        def confirm_selection():
            print("AUTOMATION: confirm_selection() called")
            selected_automations.clear()
            create_new_automation = new_automation_var.get()
            
            # Collect selected existing automations
            for automation in window_instance.automations:
                if checkbox_vars[automation['id']].get():
                    print(f"AUTOMATION: Automation selected: {automation['name']} (ID: {automation['id']})")
                    selected_automations.append(automation)
            
            # Check if at least one option is selected
            if not selected_automations and not create_new_automation:
                print("AUTOMATION: WARNING - No automations selected")
                messagebox.showwarning("No Selection", "Please select at least one automation or 'New automation area'.")
                return  # Don't close the dialog if nothing is selected
            
            print("AUTOMATION: Destroying selection window...")
            selection_window.destroy()
            print("AUTOMATION: Selection window destroyed")
        
        def cancel_selection():
            print("AUTOMATION: cancel_selection() called")
            selection_window.destroy()
        
        button_frame = tk.Frame(selection_window)
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="OK", command=confirm_selection, width=10).pack(side='left', padx=5)
        tk.Button(button_frame, text="Cancel", command=cancel_selection, width=10).pack(side='left', padx=5)
        
        print("AUTOMATION: Waiting for selection window to close...")
        selection_window.wait_window()
        print("AUTOMATION: Selection window closed, wait_window() returned")
        
        # Update the main window to ensure it processes events
        print("AUTOMATION: Updating main window...")
        window_instance.window.update()
        print("AUTOMATION: Main window updated")
        
        # Check if "New automation area" was selected
        create_new = new_automation_var.get()
        
        if create_new:
            print("AUTOMATION: Creating new automation area...")
            # Create a new automation using the current window instance
            window_instance.add_automation()
            # Get the newly created automation (it's the last one in the list)
            new_automation = window_instance.automations[-1]
            # Add to selected automations list so it gets processed below
            selected_automations.append(new_automation)
            print(f"AUTOMATION: Created new automation: {new_automation['name']}")
        
        if selected_automations:
            print(f"AUTOMATION: {len(selected_automations)} automation(s) selected")
            # Assign the same image and coordinates to all selected automations
            for automation in selected_automations:
                print(f"AUTOMATION: Assigning image to {automation['name']} (ID: {automation['id']})")
                automation['image_area_coords'] = coords
                automation['reference_image'] = captured_image.copy()  # Use same image for all
                # Update preview
                window_instance.update_preview(automation, captured_image)
            print("AUTOMATION: All selected automations updated")
            window_instance._mark_unsaved_changes()
        else:
            print("AUTOMATION: No automations selected")
        
        # Restore focus to automation window after dialog is closed
        def restore_window_focus():
            try:
                if window_instance.window.winfo_exists():
                    window_instance.window.focus_force()
                    window_instance.window.lift()
                    window_instance.window.update_idletasks()
                    print(f"AUTOMATION: Restored focus to automation window")
            except Exception as e:
                print(f"AUTOMATION: Error restoring window focus: {e}")
        
        window_instance.window.after(100, restore_window_focus)
        print("=" * 60)
    
    def set_image_area(self, automation):
        """Set the image area for detection using existing area selection method"""
        # This method is kept for backward compatibility but is no longer used
        # The new flow uses start_area_selection_for_automations instead
        pass
    
    def capture_reference_image(self, automation, x1, y1, x2, y2, freeze_screen):
        """Capture and store the reference image for comparison"""
        try:
            # Capture the area
            # For reference image, we always capture directly (freeze screen is for checking, not reference)
            image = capture_screen_area(x1, y1, x2, y2)
            automation['reference_image'] = image.copy()
            
            # Update preview (scale to 40x40px)
            self.update_preview(automation, image)
            
            print(f"Reference image captured for automation {automation['id']}: {image.size}")
        except Exception as e:
            print(f"Error capturing reference image: {e}")
            messagebox.showerror("Error", f"Failed to capture reference image: {e}")
    
    def update_preview(self, automation, image):
        """Update the preview image (scaled to fit 40x40px)"""
        try:
            # Scale image to fit within 40x40px while maintaining aspect ratio
            preview_size = 40
            img_width, img_height = image.size
            
            # Calculate scaling to fit within 40x40
            scale = min(preview_size / img_width, preview_size / img_height)
            new_width = int(img_width * scale)
            new_height = int(img_height * scale)
            
            # Resize image
            preview_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(preview_image)
            
            # Update label
            automation['preview_label'].config(image=photo, text="")
            automation['preview_label'].image = photo  # Keep reference
            
            # Update status label
            if automation.get('status_label'):
                automation['status_label'].config(text="Image set", fg="green")
            
            # Ensure entry field remains enabled after update
            if automation.get('timer_entry'):
                try:
                    if automation['timer_entry'].winfo_exists():
                        automation['timer_entry'].config(state='normal')
                except:
                    pass
            
            # Ensure automation window has focus after image update
            # This allows entry fields to be immediately editable
            def ensure_window_focus():
                try:
                    if self.window.winfo_exists():
                        # Use focus_force() for more aggressive focus restoration
                        self.window.focus_force()
                        self.window.lift()
                        self.window.update_idletasks()
                        
                        # Also ensure entry field can receive focus
                        if automation.get('timer_entry'):
                            try:
                                if automation['timer_entry'].winfo_exists():
                                    automation['timer_entry'].config(state='normal')
                                    # Try to give focus to entry field after a short delay
                                    self.window.after(50, lambda: automation['timer_entry'].focus_set() if automation['timer_entry'].winfo_exists() else None)
                            except:
                                pass
                except:
                    pass
            
            # Restore focus after delays to ensure all updates are complete
            # Try multiple times in case something else is grabbing focus
            self.window.after(150, ensure_window_focus)
            self.window.after(400, ensure_window_focus)
            self.window.after(700, ensure_window_focus)
        except Exception as e:
            print(f"Error updating preview: {e}")
    
    def set_automation_hotkey(self, button, automation):
        """Set hotkey for automation - the hotkey will trigger image area selection when pressed"""
        print(f"AUTOMATION: set_automation_hotkey() called for {automation['name']}")
        print(f"AUTOMATION: Button: {button}")
        
        # Create a temporary frame for compatibility with hotkey system
        temp_frame = tk.Frame()
        temp_frame._is_automation_hotkey = True
        temp_frame._automation_ref = automation
        temp_frame._automation_window = self  # Store reference to automations window
        
        # Store callback for when hotkey is pressed (not when setting it)
        def hotkey_callback():
            # When hotkey is pressed, trigger area selection
            print(f"AUTOMATION: Hotkey pressed for {automation['name']}, triggering area selection")
            # Check if window exists - if not, reopen it
            window_exists = False
            try:
                window_exists = self.window.winfo_exists()
            except:
                pass
            
            if not window_exists:
                print("AUTOMATION: Window closed, reopening it...")
                # Reopen the automation window
                self.game_text_reader.open_automations_window()
                # Wait a bit for window to be created
                self.root.update()
                # Get reference to the current window instance
                if hasattr(self.game_text_reader, '_automations_window') and self.game_text_reader._automations_window:
                    current_window_instance = self.game_text_reader._automations_window
                    # Update self references to use current instance
                    self.window = current_window_instance.window
                    self.automations = current_window_instance.automations
                    # Find the automation in the current instance
                    for auto in current_window_instance.automations:
                        if auto.get('id') == automation.get('id'):
                            automation = auto
                            break
            
            # Now trigger area selection (works whether window was open or just reopened)
            self.set_image_area(automation)
        
        automation['hotkey_callback'] = hotkey_callback
        
        # Store automation reference in button and frame for later use
        button._automation_ref = automation
        button._automation_temp_frame = temp_frame
        # Set the callback on the button so setup_hotkey can use it
        button._automation_callback = hotkey_callback
        
        # Also store callback in registry for persistence (works even when window is closed)
        if hasattr(button, 'hotkey') and button.hotkey:
            self.automation_callbacks_by_hotkey[button.hotkey] = hotkey_callback
            print(f"AUTOMATION DEBUG: Stored callback in registry for hotkey '{button.hotkey}'")
            print(f"AUTOMATION DEBUG: Registry now has {len(self.automation_callbacks_by_hotkey)} automation callback(s)")
            print(f"AUTOMATION DEBUG: Registry keys: {list(self.automation_callbacks_by_hotkey.keys())}")
        else:
            print(f"AUTOMATION DEBUG: Button has no hotkey yet, will store in registry after hotkey is assigned")
        
        # Use existing hotkey system to start hotkey assignment
        # This will put the system in hotkey assignment mode
        print(f"AUTOMATION: Starting hotkey assignment mode...")
        self.game_text_reader.set_hotkey(button, temp_frame)
        
        # After hotkey is assigned, update the automation and displays
        # We'll check for this in a callback after setup_hotkey is called
        def update_after_hotkey_set():
            """Update automation after hotkey is set"""
            if hasattr(button, 'hotkey') and button.hotkey:
                print(f"AUTOMATION DEBUG: Hotkey assigned: {button.hotkey}")
                automation['hotkey'] = button.hotkey
                # Store callback in registry for persistence (works even when window is closed)
                self.automation_callbacks_by_hotkey[button.hotkey] = hotkey_callback
                print(f"AUTOMATION DEBUG: Stored callback in registry for hotkey '{button.hotkey}' after assignment")
                print(f"AUTOMATION DEBUG: Registry now has {len(self.automation_callbacks_by_hotkey)} automation callback(s)")
                print(f"AUTOMATION DEBUG: Registry keys: {list(self.automation_callbacks_by_hotkey.keys())}")
                print(f"AUTOMATION DEBUG: Callback function: {hotkey_callback}")
                print(f"AUTOMATION DEBUG: Window exists: {self.window.winfo_exists() if hasattr(self, 'window') else 'N/A'}")
                # Update displays
                self.update_hotkey_display(automation)
                self.update_top_hotkey_button()
                self._mark_unsaved_changes()
        
        # Check periodically if hotkey was set (setup_hotkey is called after assignment)
        def check_hotkey_set():
            if hasattr(button, 'hotkey') and button.hotkey and automation.get('hotkey') != button.hotkey:
                update_after_hotkey_set()
            elif not hasattr(button, 'hotkey') or not button.hotkey:
                # Still waiting for assignment, check again
                self.window.after(100, check_hotkey_set)
        
        # Start checking after a short delay
        self.window.after(300, check_hotkey_set)
    
    def update_hotkey_display(self, automation):
        """Update the hotkey display in the automation UI"""
        if automation.get('hotkey_label'):
            if automation.get('hotkey'):
                display_name = automation['hotkey'].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                automation['hotkey_label'].config(text=f"Hotkey: [ {display_name.upper()} ]", fg="green")
            else:
                automation['hotkey_label'].config(text="", fg="gray")
    
    def update_top_hotkey_button(self):
        """Update the top-level hotkey button to show if any automation has a hotkey"""
        # Check if any automation has a hotkey
        has_hotkey = any(automation.get('hotkey') for automation in self.automations)
        if has_hotkey:
            # Find the first automation with a hotkey
            for automation in self.automations:
                if automation.get('hotkey'):
                    display_name = automation['hotkey'].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                    self.set_hotkey_button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                    break
        else:
            self.set_hotkey_button.config(text="Set Hotkey: [ None ]")
    
    def remove_automation(self, automation):
        """Remove an automation rule"""
        # Stop any active timer
        automation['timer_active'] = False
        
        # Remove hotkey if set
        if automation.get('hotkey') and hasattr(self.game_text_reader, 'hotkeys'):
            # Hotkey cleanup will be handled by the hotkey system
            pass
        
        # Remove from list
        if automation in self.automations:
            self.automations.remove(automation)
        
        # Destroy UI frame
        if automation.get('frame'):
            automation['frame'].destroy()
        
        self._mark_unsaved_changes()
        
        # Rename remaining automations to maintain A, B, C order (keep "Detection Area:" format)
        for i, auto in enumerate(self.automations):
            new_letter = chr(65 + i)  # 65 is 'A' in ASCII
            new_name = f"Detection Area: {new_letter}"
            auto['name'] = new_name
            # Update the name label if it exists
            if auto.get('name_label'):
                auto['name_label'].config(text=new_name)
    
    def add_hotkey_combo(self):
        """Add a new hotkey combo"""
        combo_id = len(self.hotkey_combos)
        combo_letter = chr(65 + combo_id)  # 65 is 'A' in ASCII
        combo_name = f"Area Combo: {combo_letter}"
        
        combo = {
            'id': combo_id,
            'name': combo_name,
            'hotkey': None,  # Hotkey string
            'areas': [],  # List of {'area_name': str, 'timer_ms': int}
            'frame': None,  # UI frame for this combo
            'is_triggering': False,  # Whether combo is currently being triggered
            'current_area_index': 0  # Current area being processed
        }
        
        self.hotkey_combos.append(combo)
        self.create_hotkey_combo_ui(combo)
        # Automatically add the first step
        self.add_area_to_combo(combo)
        self._mark_unsaved_changes()
    
    def create_hotkey_combo_ui(self, combo):
        """Create UI elements for a single hotkey combo"""
        # Main frame for this combo
        combo_frame = tk.Frame(self.scrollable_frame_ref, relief='ridge', bd=2, padx=10, pady=10)
        combo_frame.pack(fill='x', padx=5, pady=5)
        combo['frame'] = combo_frame
        
        # Top row: Combo name, Hotkey button, Remove button
        top_row = tk.Frame(combo_frame)
        top_row.pack(fill='x', pady=5)
        
        # Combo name label
        name_label = tk.Label(
            top_row,
            text=combo['name'],
            font=("Helvetica", 10, "bold")
        )
        name_label.pack(side='left', padx=5)
        combo['name_label'] = name_label
        
        # Hotkey button
        hotkey_button = tk.Button(
            top_row,
            text="Set Hotkey: [ None ]",
            command=lambda: self.set_combo_hotkey(hotkey_button, combo),
            font=("Helvetica", 9)
        )
        hotkey_button.pack(side='left', padx=5)
        combo['hotkey_button'] = hotkey_button
        
        # Remove button
        remove_button = tk.Button(
            top_row,
            text="❌ Remove",
            command=lambda: self.remove_hotkey_combo(combo),
            font=("Helvetica", 9),
            fg="red"
        )
        remove_button.pack(side='right', padx=5)
        
        # Areas frame - will contain area entries
        areas_frame = tk.Frame(combo_frame)
        areas_frame.pack(fill='x', pady=5)
        combo['areas_frame'] = areas_frame
        
        # Add first area button
        add_area_button = tk.Button(
            combo_frame,
            text="+Add New Step",
            command=lambda: self.add_area_to_combo(combo),
            font=("Helvetica", 9)
        )
        add_area_button.pack(side='left', padx=5, pady=5)
        combo['add_area_button'] = add_area_button
    
    def set_combo_hotkey(self, button, combo):
        """Set hotkey for a combo - uses same robust pattern as automation hotkey"""
        print(f"HOTKEY COMBO: set_combo_hotkey() called for {combo['name']}")
        
        # Create a temporary frame for compatibility with hotkey system
        temp_frame = tk.Frame()
        temp_frame._is_hotkey_combo = True
        temp_frame._combo_ref = combo
        temp_frame._combo_window = self
        
        # Store callback for when hotkey is pressed - triggers the combo
        # Make it a method that can be restored if lost
        def hotkey_callback():
            # When hotkey is pressed, trigger the combo
            print("=" * 60)
            print(f"HOTKEY COMBO: Hotkey pressed for {combo['name']}")
            print(f"HOTKEY COMBO: Button: {button}")
            print(f"HOTKEY COMBO: Has callback attr: {hasattr(button, '_combo_callback')}")
            if hasattr(button, '_combo_callback'):
                print(f"HOTKEY COMBO: Callback value: {button._combo_callback}")
            try:
                self.trigger_hotkey_combo(combo)
            finally:
                # ALWAYS re-set callback after use to ensure it persists
                # This is critical - the callback must persist after being called
                # Use the backup from combo dictionary as the source of truth
                if combo.get('_hotkey_callback_backup'):
                    button._combo_callback = combo['_hotkey_callback_backup']
                    print("HOTKEY COMBO: Callback restored from backup after execution")
                elif hasattr(button, '_combo_callback'):
                    button._combo_callback = hotkey_callback
                    print("HOTKEY COMBO: Callback re-set after execution")
            print("=" * 60)
        
        combo['hotkey_callback'] = hotkey_callback
        
        # Store callback on button IMMEDIATELY so it's available when setup_hotkey creates the handler
        # Also store it on the combo as a backup in case button reference changes
        print(f"HOTKEY COMBO: Setting callback on button: {button}")
        button._combo_callback = hotkey_callback
        button._combo_temp_frame = temp_frame
        button._combo_ref = combo
        # Store as backup on combo
        combo['_hotkey_callback_backup'] = hotkey_callback
        combo['_hotkey_button'] = button
        print(f"HOTKEY COMBO: Callback set: {hasattr(button, '_combo_callback')}")
        print(f"HOTKEY COMBO: Callback value: {button._combo_callback}")
        
        # Use existing hotkey system to start hotkey assignment
        print("HOTKEY COMBO: Starting hotkey assignment mode...")
        self.game_text_reader.set_hotkey(button, temp_frame)
        
        # After hotkey is assigned, ensure callback is set and update the display
        # setup_hotkey is called immediately after hotkey assignment in _finalize_hotkey
        # We need to ensure the callback is still there after setup_hotkey completes
        def update_after_hotkey_set():
            """Update button display after hotkey is set and ensure callback is preserved"""
            if hasattr(button, 'hotkey') and button.hotkey:
                print(f"HOTKEY COMBO: Hotkey assigned: {button.hotkey}")
                
                # CRITICAL: Re-set the callback after setup_hotkey completes
                # The handler in setup_hotkey checks for this callback at runtime, so it must exist
                print("HOTKEY COMBO: Re-setting callback after hotkey setup completes")
                button._combo_callback = hotkey_callback
                # Also update backup
                combo['_hotkey_callback_backup'] = hotkey_callback
                combo['_hotkey_button'] = button
                
                # Also ensure the temp_frame reference is preserved
                if not hasattr(button, '_combo_temp_frame'):
                    button._combo_temp_frame = temp_frame
                if not hasattr(button, '_combo_ref'):
                    button._combo_ref = combo
                
                # CRITICAL: Explicitly call setup_hotkey to ensure it's registered globally
                # setup_hotkey is already called in _finalize_hotkey, but we call it again
                # to ensure the callback is properly set when the handler is created
                # However, we need to clean up any existing hook first to avoid duplicates
                print("HOTKEY COMBO: Ensuring hotkey is properly registered with callback")
                try:
                    # Clean up any existing hook first to avoid duplicates
                    if hasattr(button, 'keyboard_hook') and button.keyboard_hook:
                        try:
                            import keyboard
                            if hasattr(button.keyboard_hook, 'remove'):
                                keyboard.remove_hotkey(button.keyboard_hook)
                            else:
                                keyboard.unhook(button.keyboard_hook)
                            print("HOTKEY COMBO: Cleaned up existing hook before re-registering")
                        except Exception as e:
                            print(f"HOTKEY COMBO: Warning: Error cleaning up existing hook: {e}")
                        finally:
                            button.keyboard_hook = None
                    
                    # Ensure callback is set before calling setup_hotkey
                    button._combo_callback = hotkey_callback
                    self.game_text_reader.setup_hotkey(button, None)
                    # Re-set callback after setup_hotkey in case it got cleared
                    button._combo_callback = hotkey_callback
                    print("HOTKEY COMBO: setup_hotkey called successfully, callback preserved")
                except Exception as e:
                    print(f"HOTKEY COMBO: Error calling setup_hotkey: {e}")
                    import traceback
                    traceback.print_exc()
                
                # Verify it's set
                if hasattr(button, '_combo_callback') and button._combo_callback:
                    print("HOTKEY COMBO: ✓ Callback is set and ready")
                    print(f"HOTKEY COMBO: Callback function: {button._combo_callback}")
                else:
                    print("HOTKEY COMBO: ✗ ERROR - Callback is NOT set!")
                    print(f"HOTKEY COMBO: Button: {button}")
                    print(f"HOTKEY COMBO: Has attr: {hasattr(button, '_combo_callback')}")
                    if hasattr(button, '_combo_callback'):
                        print(f"HOTKEY COMBO: Callback value: {button._combo_callback}")
                
                combo['hotkey'] = button.hotkey
                # Store callback in registry by hotkey name for reliable lookup
                self.combo_callbacks_by_hotkey[button.hotkey] = hotkey_callback
                print(f"HOTKEY COMBO: Stored callback in registry for hotkey '{button.hotkey}'")
                
                display_name = button.hotkey.replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                self._mark_unsaved_changes()
                
                # Set up periodic callback restoration to ensure it never gets lost
                def ensure_callback_persists():
                    """Periodically ensure callback is set - prevents it from being lost"""
                    if hasattr(combo, '_hotkey_button') and combo.get('_hotkey_button'):
                        button_ref = combo['_hotkey_button']
                        if hasattr(button_ref, 'hotkey') and button_ref.hotkey:
                            # Check if callback is missing or None
                            if not hasattr(button_ref, '_combo_callback') or not button_ref._combo_callback:
                                # Restore from backup
                                if combo.get('_hotkey_callback_backup'):
                                    print("HOTKEY COMBO: Restoring lost callback from backup")
                                    button_ref._combo_callback = combo['_hotkey_callback_backup']
                            # Schedule next check
                            self.window.after(2000, ensure_callback_persists)
                
                # Start periodic checking
                self.window.after(2000, ensure_callback_persists)
            else:
                # Still waiting for assignment, check again
                self.window.after(100, update_after_hotkey_set)
        
        # Start checking after a delay to allow setup_hotkey to complete
        # setup_hotkey is called in _finalize_hotkey which happens ~250ms after key press
        # So we check after 500ms to be safe
        self.window.after(500, update_after_hotkey_set)
        
        # Also set it again after a longer delay to be absolutely sure
        self.window.after(1000, lambda: setattr(button, '_combo_callback', hotkey_callback) if hasattr(button, 'hotkey') and button.hotkey else None)
    
    def add_area_to_combo(self, combo):
        """Add an area entry to a hotkey combo"""
        area_entry = {
            'area_name': tk.StringVar(value=""),
            'timer_ms': tk.IntVar(value=0),
            'frame': None
        }
        combo['areas'].append(area_entry)
        self.create_area_entry_ui(combo, area_entry, len(combo['areas']) - 1)
        self._mark_unsaved_changes()
    
    def create_area_entry_ui(self, combo, area_entry, index):
        """Create UI for a single area entry in a combo"""
        # Frame for this area entry
        entry_frame = tk.Frame(combo['areas_frame'], relief='sunken', bd=1, padx=5, pady=5)
        entry_frame.pack(fill='x', padx=5, pady=2)
        area_entry['frame'] = entry_frame
        
        # Row 1: Area selection
        area_row = tk.Frame(entry_frame)
        area_row.pack(fill='x', pady=2)
        
        # Area label
        tk.Label(area_row, text=f"Step {index + 1}:", font=("Helvetica", 9)).pack(side='left', padx=5)
        
        # Create combobox that refreshes when opened
        def refresh_combo_dropdown():
            """Refresh the dropdown options when opened"""
            trigger_options = self.get_available_trigger_options(exclude_combo=combo)
            if not trigger_options or trigger_options == ["No options available"]:
                trigger_options = ["No options available"]
            area_dropdown['values'] = trigger_options
            # Preserve current selection if it's still valid
            current_value = area_entry['area_name'].get()
            if current_value not in trigger_options and current_value:
                area_entry['area_name'].set("")
        
        # Get initial list of available trigger options (areas, automations, area combos)
        # Exclude the current combo to prevent self-reference
        trigger_options = self.get_available_trigger_options(exclude_combo=combo)
        
        if not trigger_options or trigger_options == ["No options available"]:
            trigger_options = ["No options available"]
            area_entry['area_name'].set("No options available")
        
        # Area dropdown
        area_dropdown = ttk.Combobox(
            area_row,
            textvariable=area_entry['area_name'],
            values=trigger_options,
            state='readonly',
            width=18,
            font=("Helvetica", 8),
            postcommand=refresh_combo_dropdown
        )
        area_dropdown.pack(side='left', padx=5)
        
        # Remove area button (only if not the first area)
        if index > 0:
            remove_area_button = tk.Button(
                area_row,
                text="❌",
                command=lambda: self.remove_area_from_combo(combo, area_entry),
                font=("Helvetica", 8),
                fg="red",
                width=3
            )
            remove_area_button.pack(side='right', padx=5)
        
        # Row 2: Timer input and progress bar (appears under the area selection)
        timer_row = tk.Frame(entry_frame)
        timer_row.pack(fill='x', pady=2, padx=(30, 0))  # Indent to align under dropdown
        
        # Track changes to area_name
        def on_area_name_change(*args):
            self._mark_unsaved_changes()
        area_entry['area_name'].trace('w', on_area_name_change)
        
        # Track changes to timer_ms
        def on_timer_ms_change(*args):
            self._mark_unsaved_changes()
            # Update progress label to show 0/total ms format when timer is set
            try:
                timer_ms_value = area_entry['timer_ms'].get()
                if timer_progress_label and timer_ms_value > 0:
                    # Only update if timer is not currently active (not counting down)
                    if area_entry.get('timer_progress') and area_entry['timer_progress']['value'] == 0:
                        timer_progress_label.config(text=f"0/{int(timer_ms_value)}ms")
                elif timer_progress_label and timer_ms_value == 0:
                    timer_progress_label.config(text="0ms")
            except:
                pass
        area_entry['timer_ms'].trace('w', on_timer_ms_change)
        
        # Timer input
        timer_frame = tk.Frame(timer_row)
        timer_frame.pack(side='left', padx=5)
        
        tk.Label(timer_frame, text="Step delay", font=("Helvetica", 9)).pack(side='left')
        timer_entry = tk.Entry(timer_frame, textvariable=area_entry['timer_ms'], width=6, font=("Helvetica", 9))
        timer_entry.pack(side='left', padx=2)
        tk.Label(timer_frame, text="ms", font=("Helvetica", 9)).pack(side='left')
        
        # Store reference to timer entry for focus removal and state management
        area_entry['timer_entry'] = timer_entry
        
        # Ensure entry field always stays enabled and can receive focus
        # This prevents it from becoming uneditable after UI updates or combo triggers
        def ensure_entry_enabled():
            """Ensure entry field is always enabled and can receive focus"""
            try:
                if timer_entry.winfo_exists():
                    timer_entry.config(state='normal')
            except:
                pass
        
        # Re-enable after any window updates - bind to multiple events for robustness
        timer_entry.bind("<FocusIn>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<Button-1>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<Enter>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<KeyPress>", lambda e: ensure_entry_enabled())
        timer_entry.bind("<KeyRelease>", lambda e: ensure_entry_enabled())
        
        # Periodic check to ensure entry stays enabled (every 500ms)
        def periodic_enable_check():
            try:
                if self.window.winfo_exists() and timer_entry.winfo_exists():
                    ensure_entry_enabled()
                    self.window.after(500, periodic_enable_check)
            except:
                pass
        
        # Start periodic check
        self.window.after(500, periodic_enable_check)
        
        # Progress bar for timer countdown
        progress_frame = tk.Frame(timer_row)
        progress_frame.pack(side='left', padx=10, fill='x', expand=True)
        
        timer_progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=150,
            maximum=100
        )
        timer_progress.pack(side='left', padx=5)
        area_entry['timer_progress'] = timer_progress
        
        timer_progress_label = tk.Label(progress_frame, text="", font=("Helvetica", 8), width=8)
        timer_progress_label.pack(side='left', padx=2)
        area_entry['timer_progress_label'] = timer_progress_label
        
        # Initialize progress bar - show 0/total ms format if timer is set
        timer_progress['value'] = 0
        try:
            timer_ms_value = area_entry['timer_ms'].get()
            if timer_ms_value > 0:
                timer_progress_label.config(text=f"0/{int(timer_ms_value)}ms")
            else:
                timer_progress_label.config(text="0ms")
        except:
            timer_progress_label.config(text="0ms")
    
    def remove_area_from_combo(self, combo, area_entry):
        """Remove an area entry from a combo"""
        if area_entry in combo['areas']:
            combo['areas'].remove(area_entry)
            if area_entry.get('frame'):
                area_entry['frame'].destroy()
            # Recreate UI to update indices
            self.recreate_combo_areas_ui(combo)
            self._mark_unsaved_changes()
    
    def recreate_combo_areas_ui(self, combo):
        """Recreate the areas UI for a combo (to update indices)"""
        # Destroy all area entry frames
        for area_entry in combo['areas']:
            if area_entry.get('frame'):
                area_entry['frame'].destroy()
        
        # Recreate all area entries
        for i, area_entry in enumerate(combo['areas']):
            self.create_area_entry_ui(combo, area_entry, i)
        
        # Ensure all timer entries are enabled after recreation
        def ensure_timer_entries_enabled():
            """Ensure all timer entries are enabled after UI recreation"""
            for area_entry in combo['areas']:
                if area_entry.get('timer_entry'):
                    try:
                        if area_entry['timer_entry'].winfo_exists():
                            area_entry['timer_entry'].config(state='normal')
                    except:
                        pass
        
        # Re-enable after a short delay to ensure UI is fully created
        self.window.after(50, ensure_timer_entries_enabled)
    
    def remove_hotkey_combo(self, combo):
        """Remove a hotkey combo"""
        # Stop any active triggering
        combo['is_triggering'] = False
        
        # Remove from callback registry
        if combo.get('hotkey'):
            if combo['hotkey'] in self.combo_callbacks_by_hotkey:
                del self.combo_callbacks_by_hotkey[combo['hotkey']]
                print(f"HOTKEY COMBO: Removed callback from registry for hotkey '{combo['hotkey']}'")
        
        # Remove from game_text_reader.hotkeys dictionary if it exists
        hotkey_button = combo.get('hotkey_button')
        if hotkey_button and hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey:
            if hasattr(self.game_text_reader, 'hotkeys') and hotkey_button.hotkey in self.game_text_reader.hotkeys:
                try:
                    self.game_text_reader.hotkeys[hotkey_button.hotkey].unhook()
                    del self.game_text_reader.hotkeys[hotkey_button.hotkey]
                    print(f"HOTKEY COMBO: Removed from game_text_reader.hotkeys: {hotkey_button.hotkey}")
                except Exception as e:
                    print(f"HOTKEY COMBO: Error removing from game_text_reader.hotkeys: {e}")
        
        # Clean up hotkey hook if it exists - try multiple methods to ensure it's removed
        if hotkey_button and hasattr(hotkey_button, 'hotkey') and hotkey_button.hotkey:
            try:
                import keyboard
                # Method 1: Remove by hotkey name (most reliable for add_hotkey)
                try:
                    keyboard.remove_hotkey(hotkey_button.hotkey)
                    print(f"HOTKEY COMBO: Removed hotkey '{hotkey_button.hotkey}' by name")
                except Exception as e:
                    print(f"HOTKEY COMBO: Could not remove hotkey by name: {e}")
                
                # Method 2: Remove by hook object if it exists
                if hasattr(hotkey_button, 'keyboard_hook') and hotkey_button.keyboard_hook:
                    try:
                        # Check if this is a custom ctrl hook or a regular add_hotkey hook
                        if hasattr(hotkey_button.keyboard_hook, 'remove'):
                            # This is an add_hotkey hook
                            keyboard.remove_hotkey(hotkey_button.keyboard_hook)
                            print(f"HOTKEY COMBO: Removed hotkey hook object for {combo['name']}")
                        else:
                            # This is a custom on_press hook or hook ID
                            keyboard.unhook(hotkey_button.keyboard_hook)
                            print(f"HOTKEY COMBO: Unhooked custom keyboard hook for {combo['name']}")
                    except Exception as e:
                        print(f"HOTKEY COMBO: Warning: Error removing hook object: {e}")
                    finally:
                        # Always set to None to prevent future errors
                        hotkey_button.keyboard_hook = None
            except Exception as e:
                print(f"HOTKEY COMBO: Warning: Error cleaning up keyboard hook: {e}")
        
        # Clean up mouse hook if it exists
        if hotkey_button and hasattr(hotkey_button, 'mouse_hook_id') and hotkey_button.mouse_hook_id:
            try:
                import mouse
                mouse.unhook(hotkey_button.mouse_hook_id)
                print(f"HOTKEY COMBO: Removed mouse hook for {combo['name']}")
            except Exception as e:
                print(f"HOTKEY COMBO: Warning: Error cleaning up mouse hook: {e}")
            finally:
                hotkey_button.mouse_hook_id = None
        
        # Clear callback references
        if hotkey_button:
            if hasattr(hotkey_button, '_combo_callback'):
                delattr(hotkey_button, '_combo_callback')
            if hasattr(hotkey_button, '_combo_temp_frame'):
                delattr(hotkey_button, '_combo_temp_frame')
            if hasattr(hotkey_button, '_combo_ref'):
                delattr(hotkey_button, '_combo_ref')
        
        # Remove from list
        if combo in self.hotkey_combos:
            self.hotkey_combos.remove(combo)
        
        # Destroy UI frame
        if combo.get('frame'):
            combo['frame'].destroy()
        
        self._mark_unsaved_changes()
        
        # Rename remaining combos to maintain A, B, C order (keep "Area Combo:" format)
        for i, c in enumerate(self.hotkey_combos):
            new_letter = chr(65 + i)  # 65 is 'A' in ASCII
            new_name = f"Area Combo: {new_letter}"
            c['name'] = new_name
            if c.get('name_label'):
                c['name_label'].config(text=new_name)
    
    def trigger_hotkey_combo(self, combo):
        """Trigger a hotkey combo - sequentially read areas with timers"""
        # Check if combo is already running - prevent re-triggering
        if combo.get('is_triggering', False):
            print(f"HOTKEY COMBO: Already triggering {combo['name']}, ignoring duplicate trigger")
            return
        
        if not combo['areas']:
            print(f"HOTKEY COMBO: No areas configured for {combo['name']}")
            return
        
        # Filter out invalid triggers (areas, automations, or area combos)
        valid_triggers = []
        for area_entry in combo['areas']:
            trigger_name = area_entry['area_name'].get()
            if trigger_name and trigger_name not in ["No areas available", "No options available"]:
                # Skip separator entries
                if trigger_name.startswith("───"):
                    continue
                
                trigger_info = None
                
                # Check if it's a regular area
                area_frame = None
                for area in self.game_text_reader.areas:
                    if area[3].get() == trigger_name:  # area_name_var
                        area_frame = area[0]  # area_frame
                        break
                
                if area_frame:
                    trigger_info = {
                        'type': 'area',
                        'area_frame': area_frame,
                        'name': trigger_name,
                        'timer_ms': area_entry['timer_ms'].get()
                    }
                
                if not trigger_info:
                    # Check if it's an automation
                    for automation in self.automations:
                        if automation['name'] == trigger_name:
                            trigger_info = {
                                'type': 'automation',
                                'automation': automation,
                                'name': trigger_name,
                                'timer_ms': area_entry['timer_ms'].get()
                            }
                            break
                    
                    # Check if it's an area combo
                    if not trigger_info:
                        for other_combo in self.hotkey_combos:
                            if other_combo['name'] == trigger_name:
                                trigger_info = {
                                    'type': 'combo',
                                    'combo': other_combo,
                                    'name': trigger_name,
                                    'timer_ms': area_entry['timer_ms'].get()
                                }
                                break
                
                if trigger_info:
                    valid_triggers.append(trigger_info)
        
        if not valid_triggers:
            print(f"HOTKEY COMBO: No valid triggers found for {combo['name']}")
            # Show warning dialog
            messagebox.showwarning(
                "No Steps Configured",
                f"Area Combo '{combo['name']}' has no steps configured.\n\n"
                "Please add at least one step (area, automation, or combo) before using this hotkey.",
                parent=self.window
            )
            return
        
        print(f"HOTKEY COMBO: Triggering {combo['name']} with {len(valid_triggers)} triggers")
        # Update main window status
        self._update_main_status(f"{combo['name']} started ({len(valid_triggers)} steps)", "blue", 2000)
        # Set flag BEFORE starting to prevent race conditions
        combo['is_triggering'] = True
        combo['current_area_index'] = 0
        combo['_valid_areas'] = valid_triggers  # Keep same variable name for compatibility
        
        # Start with the first trigger
        self._process_next_area_in_combo(combo)
    
    def _process_next_area_in_combo(self, combo):
        """Process the next area in a combo"""
        # Double-check flag is still set
        if not combo.get('is_triggering', False):
            print(f"HOTKEY COMBO: Combo {combo['name']} is no longer triggering, stopping")
            return
        
        valid_areas = combo.get('_valid_areas', [])
        current_index = combo['current_area_index']
        
        if current_index >= len(valid_areas):
            # All areas processed
            print(f"HOTKEY COMBO: Completed {combo['name']}")
            # Update main window status
            self._update_main_status(f"{combo['name']} completed", "green", 2000)
            # Clear flag immediately to allow re-triggering
            combo['is_triggering'] = False
            combo['current_area_index'] = 0
            combo['_valid_areas'] = []
            print(f"HOTKEY COMBO: Flag cleared for {combo['name']}, ready for next trigger")
            
            # Re-enable all timer entries in this combo to ensure they're editable
            def re_enable_combo_entries():
                try:
                    for area_entry in combo.get('areas', []):
                        if area_entry.get('timer_entry'):
                            try:
                                if area_entry['timer_entry'].winfo_exists():
                                    area_entry['timer_entry'].config(state='normal')
                            except:
                                pass
                except:
                    pass
            
            # Re-enable entries after a short delay to ensure UI is ready
            self.root.after(100, re_enable_combo_entries)
            return
        
        trigger_info = valid_areas[current_index]
        trigger_name = trigger_info['name']
        trigger_type = trigger_info['type']
        timer_ms = trigger_info['timer_ms']
        
        print(f"HOTKEY COMBO: Processing trigger {current_index + 1}/{len(valid_areas)}: {trigger_name} (type: {trigger_type})")
        # Update main window status for combo step
        combo_name = combo.get('name', 'Unknown')
        self._update_main_status(f"{combo_name}: Step {current_index + 1}/{len(valid_areas)} - {trigger_name}", "blue", 1500)
        
        # Trigger based on type
        if trigger_type == 'area':
            # Trigger the area read
            area_frame = trigger_info['area_frame']
            area_name = trigger_info['name']
            
            # Check if this is an Auto Read area - if so, trigger area selection dialog
            if area_name.startswith("Auto Read"):
                print(f"HOTKEY COMBO: Triggering Auto Read area '{area_name}' - opening area selection dialog")
                # Find the area info to get area_name_var and set_area_button
                # Match by area name (more reliable than frame reference)
                area_name_var = None
                set_area_button = None
                found_area_frame = None
                for area in self.game_text_reader.areas:
                    if len(area) >= 4:
                        current_area_name_var = area[3]  # area_name_var
                        if hasattr(current_area_name_var, 'get') and current_area_name_var.get() == area_name:
                            # Found matching area by name
                            found_area_frame = area[0]  # area_frame
                            if len(area) >= 9:
                                _, _, set_area_button, area_name_var, _, _, _, _, _ = area[:9]
                            else:
                                _, _, set_area_button, area_name_var, _, _, _, _ = area[:8]
                            break
                
                if area_name_var and found_area_frame:
                    # Mark the frame to indicate it's being called from a combo
                    # This allows set_auto_read_area to notify us when area selection completes
                    # Also mark that it's not cancelled initially
                    found_area_frame._combo_callback = lambda: self._start_speech_monitoring_for_combo(combo, timer_ms)
                    found_area_frame._combo_cancelled = False  # Will be set to True if cancelled
                    # Trigger area selection dialog (same as pressing the Auto Read hotkey)
                    # Note: set_area_button is None for Auto Read areas, which is fine - set_auto_read_area doesn't use it
                    def trigger_auto_read():
                        try:
                            print(f"HOTKEY COMBO: Calling set_auto_read_area for '{area_name}' (frame={found_area_frame}, area_name_var={area_name_var})")
                            # Use game_text_reader.root to ensure we're using the correct root window
                            self.game_text_reader.set_auto_read_area(found_area_frame, area_name_var, set_area_button)
                            print(f"HOTKEY COMBO: set_auto_read_area call completed for '{area_name}'")
                        except Exception as e:
                            print(f"HOTKEY COMBO: Error calling set_auto_read_area for '{area_name}': {e}")
                            import traceback
                            traceback.print_exc()
                            # If Auto Read fails, skip to next step
                            combo['current_area_index'] += 1
                            self.game_text_reader.root.after(0, lambda: self._process_next_area_in_combo(combo))
                    # Use game_text_reader.root to ensure it works even when automation window is closed
                    self.game_text_reader.root.after(0, trigger_auto_read)
                    # Don't start speech monitoring yet - wait for area selection to complete
                    # The combo callback will start monitoring after read_area() is called
                    return
                else:
                    print(f"HOTKEY COMBO: Warning - Could not find area info for Auto Read area '{area_name}' (area_name_var={area_name_var}, set_area_button={set_area_button}, found_area_frame={found_area_frame}), skipping")
                    # Move to next area if we can't trigger this one
                    combo['current_area_index'] += 1
                    self.root.after(0, lambda: self._process_next_area_in_combo(combo))
                    return
            else:
                # Regular area - set up callback to start speech monitoring after read_area starts
                # This ensures speech monitoring starts after speech actually begins
                area_frame._combo_callback = lambda: self._start_speech_monitoring_for_combo(combo, timer_ms)
                self.root.after(0, lambda: self.game_text_reader.read_area(area_frame))
                # Don't start speech monitoring here - wait for callback from read_area
                return
        elif trigger_type == 'automation':
            # Trigger the automation
            automation = trigger_info['automation']
            self.root.after(0, lambda: self.check_automation(automation))
        elif trigger_type == 'combo':
            # Trigger the area combo
            other_combo = trigger_info['combo']
            self.root.after(0, lambda: self.trigger_hotkey_combo(other_combo))
        
        # Wait for speech to complete, then wait for timer, then move to next area
        # (Only for automations and nested combos - regular areas use callback above)
        self._start_speech_monitoring_for_combo(combo, timer_ms)
    
    def _start_speech_monitoring_for_combo(self, combo, timer_ms):
        """Start speech monitoring for a combo step"""
        # Check if this was triggered from a cancelled Auto Read
        # If so, skip speech monitoring and go straight to timer
        valid_areas = combo.get('_valid_areas', [])
        current_index = combo.get('current_area_index', 0)
        was_cancelled = False
        
        if current_index < len(valid_areas):
            trigger_info = valid_areas[current_index]
            if trigger_info.get('type') == 'area' and trigger_info.get('name', '').startswith("Auto Read"):
                # This is an Auto Read area - check if it was cancelled
                area_frame = trigger_info.get('area_frame')
                if area_frame and hasattr(area_frame, '_combo_cancelled') and area_frame._combo_cancelled:
                    was_cancelled = True
                    print(f"HOTKEY COMBO: Auto Read was cancelled for step {current_index + 1}, skipping speech wait")
        
        if was_cancelled:
            # Auto Read was cancelled - no speech to wait for, go straight to timer
            print(f"HOTKEY COMBO: Auto Read cancelled, starting timer immediately (if set)")
            # Mark this step as cancelled in the combo to prevent any speech monitoring
            combo['_current_step_cancelled'] = True
            if timer_ms > 0:
                print(f"HOTKEY COMBO: Waiting {timer_ms}ms before next area")
                self._start_timer_countdown(combo, current_index, timer_ms)
            else:
                # No timer, move to next area immediately
                print(f"HOTKEY COMBO: Moving to next area immediately (no timer)")
                self._move_to_next_area(combo)
            return
        
        # Clear cancellation flag if not cancelled (in case it was set from a previous step)
        combo['_current_step_cancelled'] = False
        
        speech_start_time = time.time()
        max_wait_for_speech_start = 2.0  # Wait up to 2 seconds for speech to start
        max_wait_for_speech_finish = 60.0  # Maximum 60 seconds to wait for speech to finish (prevents infinite loops)
        max_wait_for_stuck_sapi = 15.0  # If SAPI reports running for this long, assume it's stuck and force continue
        
        # Track if we've confirmed speech has started (to prevent starting timer too early)
        speech_confirmed_started = False
        
        def check_speech_and_continue():
            # Get current_index from combo (may have changed if combo advanced)
            current_index = combo.get('current_area_index', 0)
            
            # First, check if this step was cancelled (Auto Read cancellation)
            # This prevents speech monitoring from continuing if cancellation happened
            if combo.get('_current_step_cancelled', False):
                # This step was cancelled - stop speech monitoring immediately
                print(f"HOTKEY COMBO: Step {current_index + 1} was cancelled, stopping speech monitoring checks")
                return
            
            # Also check the frame directly as a backup
            valid_areas = combo.get('_valid_areas', [])
            if current_index < len(valid_areas):
                trigger_info = valid_areas[current_index]
                if trigger_info.get('type') == 'area' and trigger_info.get('name', '').startswith("Auto Read"):
                    area_frame = trigger_info.get('area_frame')
                    if area_frame and hasattr(area_frame, '_combo_cancelled') and area_frame._combo_cancelled:
                        # This was cancelled - stop speech monitoring immediately
                        print(f"HOTKEY COMBO: Detected cancellation via frame check for step {current_index + 1}, stopping speech checks")
                        combo['_current_step_cancelled'] = True
                        return
            
            # Check if still speaking using multiple methods for reliability
            is_still_speaking = False
            speech_has_started = False
            
            # Check timeout - if we've been waiting too long, force continue
            time_since_start = time.time() - speech_start_time
            
            # Method 1: Check the is_speaking flag first (most reliable - updated by speech monitor thread)
            # The speech monitor thread updates this flag when speech actually finishes
            if self.game_text_reader.is_speaking:
                is_still_speaking = True
                speech_has_started = True
                # Mark that we've confirmed speech has started
                nonlocal speech_confirmed_started
                speech_confirmed_started = True
            else:
                # Flag says not speaking - but we need to check SAPI to see if speech is actually running
                # because the flag might be False if speech hasn't started yet
                # We'll check SAPI below to determine if speech is actually done or just hasn't started
                is_still_speaking = None  # Unknown - need to check SAPI
                speech_has_started = False  # Don't assume it started if flag is False
            
            # Early exit: If is_speaking is False and enough time has passed, trust the flag
            # The flag is updated by the speech monitor thread which is more reliable than SAPI
            # SAPI can report "running" even when there's no text (empty string or similar)
            if not self.game_text_reader.is_speaking:
                # If enough time has passed (0.4s), assume no text or speech finished
                # We use 0.4s to give speech time to start, but not wait too long if there's no text
                if time_since_start >= 0.4:
                    # Flag says not speaking and enough time has passed - trust it and skip speech monitoring
                    # This handles cases where there's no text OR speech finished quickly
                    if speech_confirmed_started:
                        print(f"HOTKEY COMBO: Speech finished for area {current_index + 1} (is_speaking=False after {time_since_start:.1f}s)")
                    else:
                        print(f"HOTKEY COMBO: No text detected for area {current_index + 1} (is_speaking=False after {time_since_start:.1f}s), skipping speech wait")
                    # Go straight to timer or next step
                    if timer_ms > 0:
                        print(f"HOTKEY COMBO: Waiting {timer_ms}ms before next area")
                        self._start_timer_countdown(combo, current_index, timer_ms)
                    else:
                        print(f"HOTKEY COMBO: Moving to next area immediately (no timer)")
                        self._move_to_next_area(combo)
                    return
                # If less than 0.4s has passed and flag is False, wait a bit more
                # (This handles the case where callback runs before speech actually starts)
                elif time_since_start < 0.4:
                    # Very early - speech might not have started yet, wait a bit more
                    self.root.after(100, check_speech_and_continue)
                    return
            
            # Method 2: Check SAPI to determine actual speech status
            # This is needed because the flag might be False if speech hasn't started yet
            if hasattr(self.game_text_reader, 'speaker') and self.game_text_reader.speaker:
                try:
                    # SAPI SpVoice has a Status property with RunningState
                    # RunningState can be: 0=Not running, 1=Running
                    # This is the most reliable way to check if speech is actually happening
                    try:
                        status = self.game_text_reader.speaker.Status
                        if hasattr(status, 'RunningState'):
                            running_state = status.RunningState
                            # Only print if SAPI says running (to reduce console spam when no text)
                            if running_state == 1:
                                print(f"HOTKEY COMBO: SAPI RunningState={running_state} for area {current_index + 1}")
                            if running_state == 1:  # 1 = SPEVSF_RUNNING
                                # SAPI confirms speaking - but check if flag disagrees
                                # If is_speaking is False and enough time has passed, trust the flag
                                # (SAPI can report "running" even when there's no text)
                                if not self.game_text_reader.is_speaking and time_since_start > 0.6:
                                    # Flag says not speaking - trust it over SAPI (flag is more reliable)
                                    print(f"HOTKEY COMBO: SAPI says running but is_speaking flag is False (after {time_since_start:.1f}s) - trusting flag, no text detected")
                                    is_still_speaking = False
                                    speech_has_started = True
                                else:
                                    # SAPI confirms speaking and flag agrees (or too early to tell)
                                    is_still_speaking = True
                                    speech_has_started = True
                                    # Only print if flag also says speaking (to reduce spam when no text)
                                    if self.game_text_reader.is_speaking:
                                        print(f"HOTKEY COMBO: SAPI confirms speech is running")
                            else:
                                # SAPI says not running
                                if is_still_speaking is None:
                                    # Flag was False and we didn't know status - now SAPI confirms not running
                                    # BUT: SAPI can incorrectly report "not running" briefly during speech
                                    # So we need to be more careful - check if flag also says not speaking
                                    # and wait longer to ensure speech really finished
                                    if time_since_start < 0.5:
                                        # Very early - speech probably hasn't started yet
                                        is_still_speaking = None  # Still unknown, wait a bit more
                                        speech_has_started = False
                                        print(f"HOTKEY COMBO: Too early ({time_since_start:.2f}s) - speech may not have started yet, waiting...")
                                    elif not self.game_text_reader.is_speaking and time_since_start >= 0.8:
                                        # Flag says not speaking AND SAPI says not running AND enough time passed
                                        # This is more reliable - both agree and enough time has passed
                                        is_still_speaking = False
                                        speech_has_started = True
                                        print(f"HOTKEY COMBO: Both flag and SAPI confirm speech complete (RunningState={running_state}, elapsed: {time_since_start:.1f}s)")
                                    else:
                                        # SAPI says not running but flag might still be True or not enough time passed
                                        # Wait a bit more to be sure
                                        is_still_speaking = None  # Still unknown, wait more
                                        speech_has_started = False
                                        print(f"HOTKEY COMBO: SAPI says not running but waiting to confirm (elapsed: {time_since_start:.2f}s, is_speaking={self.game_text_reader.is_speaking})...")
                                else:
                                    # We already knew status - but double-check flag before trusting SAPI
                                    if self.game_text_reader.is_speaking:
                                        # Flag says speaking but SAPI says not running - trust flag (SAPI can be wrong)
                                        is_still_speaking = True
                                        speech_has_started = True
                                        # Only log periodically to reduce console spam (every 0.5 seconds)
                                        if int(time_since_start * 2) % 2 == 0:
                                            print(f"HOTKEY COMBO: Flag says speaking but SAPI says not running - trusting flag (elapsed: {time_since_start:.1f}s)")
                                    else:
                                        # Both agree - speech is done
                                        is_still_speaking = False
                                        speech_has_started = True
                                        print(f"HOTKEY COMBO: Both flag and SAPI confirm speech complete")
                    except AttributeError:
                        # Status might not have RunningState, try alternative method
                        # Use WaitUntilDone with 0 timeout (non-blocking check)
                        try:
                            import pythoncom
                            pythoncom.CoInitialize()
                            try:
                                # WaitUntilDone(0) returns True if done, False if still speaking
                                is_done = self.game_text_reader.speaker.WaitUntilDone(0)
                                if not is_done:
                                    # Still speaking
                                    is_still_speaking = True
                                    speech_has_started = True
                                    print(f"HOTKEY COMBO: SAPI WaitUntilDone confirms speech is running")
                                else:
                                    # Speech is done
                                    if is_still_speaking is None:
                                        # Flag was False and we didn't know status
                                        if time_since_start < 0.5:
                                            # Too early - speech probably hasn't started
                                            is_still_speaking = None
                                            speech_has_started = False
                                            print(f"HOTKEY COMBO: Too early ({time_since_start:.2f}s) - speech may not have started yet")
                                        else:
                                            # Enough time passed - speech is done
                                            is_still_speaking = False
                                            speech_has_started = True
                                            print(f"HOTKEY COMBO: SAPI WaitUntilDone reports speech complete")
                                    else:
                                        is_still_speaking = False
                                        speech_has_started = True
                            finally:
                                pythoncom.CoUninitialize()
                        except Exception as e:
                            # If both methods fail, check timing
                            if is_still_speaking is None:
                                # Can't check SAPI - use timing to determine if speech started
                                if time_since_start < 0.5:
                                    is_still_speaking = None
                                    speech_has_started = False
                                else:
                                    # Assume speech finished if enough time passed
                                    is_still_speaking = False
                                    speech_has_started = True
                except Exception as e:
                    # If we can't check SAPI status, use timing
                    if is_still_speaking is None:
                        if time_since_start < 0.5:
                            is_still_speaking = None
                            speech_has_started = False
                        else:
                            is_still_speaking = False
                            speech_has_started = True
            
            # Check for stuck SAPI state - if SAPI has been reporting "running" for too long, force continue
            # This handles cases where SAPI gets stuck in a running state even though speech finished
            if is_still_speaking and speech_has_started and time_since_start >= max_wait_for_stuck_sapi:
                print(f"HOTKEY COMBO: SAPI stuck detection - been reporting 'running' for {time_since_start:.1f}s (max: {max_wait_for_stuck_sapi}s)")
                print(f"HOTKEY COMBO: is_speaking flag = {self.game_text_reader.is_speaking}")
                # If the is_speaking flag is False but SAPI says running, trust the flag (it's more reliable)
                if not self.game_text_reader.is_speaking:
                    print(f"HOTKEY COMBO: is_speaking flag is False but SAPI says running - trusting flag and continuing")
                    is_still_speaking = False
                else:
                    # Both say running but it's been too long - force continue anyway
                    print(f"HOTKEY COMBO: Forcing continue due to stuck SAPI state (clearing flags)")
                    self.game_text_reader.is_speaking = False
                    is_still_speaking = False
                    speech_has_started = True
            
            # Check timeout - if we've been waiting too long, force continue regardless of SAPI status
            if time_since_start >= max_wait_for_speech_finish:
                print(f"HOTKEY COMBO: Timeout reached ({max_wait_for_speech_finish}s) for area {current_index + 1}, forcing continue...")
                # Force clear the speaking flag to prevent getting stuck
                if self.game_text_reader.is_speaking:
                    print(f"HOTKEY COMBO: Clearing stuck is_speaking flag due to timeout")
                    self.game_text_reader.is_speaking = False
                # Continue to next area despite timeout
                is_still_speaking = False
                speech_has_started = True
            
            # If speech status is still unknown or hasn't started yet, wait a bit longer
            if is_still_speaking is None or (not speech_has_started and time_since_start < max_wait_for_speech_start):
                # Speech status unknown or hasn't started yet, wait a bit longer
                if is_still_speaking is None:
                    print(f"HOTKEY COMBO: Speech status unknown for area {current_index + 1}, waiting... (elapsed: {time_since_start:.2f}s)")
                else:
                    print(f"HOTKEY COMBO: Speech hasn't started yet for area {current_index + 1}, waiting... (elapsed: {time_since_start:.2f}s)")
                self.root.after(100, check_speech_and_continue)
                return
            
            if is_still_speaking:
                # Still speaking - wait for it to finish
                # Only log periodically to reduce console spam (every 1 second)
                if int(time_since_start * 10) % 10 == 0:  # Log roughly every 1 second
                    print(f"HOTKEY COMBO: Still speaking area {current_index + 1}... (waited {time_since_start:.1f}s)")
                self.root.after(100, check_speech_and_continue)
            else:
                # Speech appears to be done - but only start timer if we confirmed speech started first
                # This prevents starting timer too early if speech hasn't actually started yet
                if not speech_confirmed_started and time_since_start < max_wait_for_speech_start:
                    # Speech hasn't started yet - wait a bit more
                    print(f"HOTKEY COMBO: Speech hasn't started yet for area {current_index + 1}, waiting... (elapsed: {time_since_start:.2f}s)")
                    self.root.after(100, check_speech_and_continue)
                    return
                
                # Speech is done (or never started and enough time passed) - NOW start the timer
                print(f"HOTKEY COMBO: Speech confirmed done for area {current_index + 1}")
                if timer_ms > 0:
                    print(f"HOTKEY COMBO: Waiting {timer_ms}ms before next area")
                    # Start timer countdown with progress bar
                    self._start_timer_countdown(combo, current_index, timer_ms)
                else:
                    # No timer, move to next area immediately
                    print(f"HOTKEY COMBO: Moving to next area immediately (no timer)")
                    self._move_to_next_area(combo)
        
        # Start checking after a short delay (give speech time to start)
        self.root.after(200, check_speech_and_continue)
    
    def _start_timer_countdown(self, combo, area_index, timer_ms):
        """Start timer countdown with progress bar visualization"""
        # Find the area entry to update its progress bar
        area_entry = None
        if area_index < len(combo['areas']):
            area_entry = combo['areas'][area_index]
        
        start_time = time.time()
        
        def update_timer_progress():
            # Check if combo was cancelled
            if not combo.get('is_triggering'):
                # Combo was cancelled, reset progress
                if area_entry and area_entry.get('timer_progress'):
                    try:
                        area_entry['timer_progress']['value'] = 0
                        if area_entry.get('timer_progress_label'):
                            # Show initial state: 0/total ms
                            if timer_ms > 0:
                                area_entry['timer_progress_label'].config(text=f"0/{int(timer_ms)}ms")
                            else:
                                area_entry['timer_progress_label'].config(text="0ms")
                    except:
                        pass  # Widget was destroyed
                return
            
            elapsed_ms = (time.time() - start_time) * 1000
            remaining_ms = max(0, timer_ms - elapsed_ms)
            
            # Update UI only if window exists
            window_exists = False
            try:
                window_exists = self.window.winfo_exists()
            except:
                pass
            
            if window_exists:
                # Window exists - update UI
                if area_entry and area_entry.get('timer_progress'):
                    try:
                        progress = min(100, (elapsed_ms / timer_ms) * 100) if timer_ms > 0 else 0
                        area_entry['timer_progress']['value'] = progress
                        
                        # Update label - use same format as automation: elapsed/total ms
                        if area_entry.get('timer_progress_label'):
                            if timer_ms > 0:
                                area_entry['timer_progress_label'].config(text=f"{int(elapsed_ms)}/{int(timer_ms)}ms")
                            else:
                                area_entry['timer_progress_label'].config(text="0ms")
                    except:
                        pass  # Widget was destroyed
            
            
            # Continue timer countdown regardless of window state
            if remaining_ms > 0:
                # Still counting down, update again in 50ms for smooth progress
                # Use root instead of window so it works even when window is closed
                self.root.after(50, update_timer_progress)
            else:
                # Timer complete, reset progress and move to next area
                if window_exists and area_entry and area_entry.get('timer_progress'):
                    try:
                        area_entry['timer_progress']['value'] = 0
                        if area_entry.get('timer_progress_label'):
                            # Show final state: elapsed/total ms
                            if timer_ms > 0:
                                area_entry['timer_progress_label'].config(text=f"{int(timer_ms)}/{int(timer_ms)}ms")
                            else:
                                area_entry['timer_progress_label'].config(text="0ms")
                    except:
                        pass  # Widget was destroyed
                self._move_to_next_area(combo)
        
        # Start the countdown
        update_timer_progress()
    
    def _move_to_next_area(self, combo):
        """Move to the next area in a combo"""
        # Clear cancellation flag when moving to next step (so it doesn't affect the next step)
        combo['_current_step_cancelled'] = False
        combo['current_area_index'] += 1
        self._process_next_area_in_combo(combo)
    
    def toggle_polling(self):
        """Start or stop background polling"""
        # Check shared state as source of truth (works even if window was closed/reopened)
        is_active = False
        if hasattr(self.game_text_reader, '_automations_polling_active'):
            is_active = self.game_text_reader._automations_polling_active
        
        # Also check if thread is actually running
        if hasattr(self, 'polling_thread') and self.polling_thread and self.polling_thread.is_alive():
            is_active = True
        
        if is_active:
            self.stop_polling()
        else:
            self.start_polling()
    
    def start_polling(self):
        """Start background polling thread"""
        # Check if already active - use shared state as source of truth
        is_already_active = False
        if hasattr(self.game_text_reader, '_automations_polling_active'):
            is_already_active = self.game_text_reader._automations_polling_active
        
        # Also check if thread is actually running
        if hasattr(self, 'polling_thread') and self.polling_thread and self.polling_thread.is_alive():
            is_already_active = True
        
        if is_already_active:
            return
        
        # Validate that at least one automation has both detection area and read area set
        has_valid_automation = False
        validation_issues = []
        
        for automation in self.automations:
            # Check if detection area is set
            has_coords = automation.get('image_area_coords') is not None
            has_image = automation.get('reference_image') is not None
            has_detection_area = has_coords and has_image
            
            # Check if read area is selected
            target_read_area_var = automation.get('target_read_area')
            if target_read_area_var:
                if hasattr(target_read_area_var, 'get'):
                    target_area = target_read_area_var.get()
                else:
                    target_area = target_read_area_var
            else:
                target_area = ""
            
            has_read_area = target_area and target_area != "" and target_area != "No areas available" and not target_area.startswith("───")
            
            # Debug output
            print(f"Validation for {automation.get('name', 'Unknown')}:")
            print(f"  - Has coords: {has_coords}")
            print(f"  - Has image: {has_image}")
            print(f"  - Has detection area: {has_detection_area}")
            print(f"  - Target area: '{target_area}'")
            print(f"  - Has read area: {has_read_area}")
            
            if has_detection_area and has_read_area:
                has_valid_automation = True
                print(f"  ✓ Valid automation found: {automation.get('name', 'Unknown')}")
                break
            else:
                if not has_detection_area:
                    if not has_coords:
                        validation_issues.append(f"{automation.get('name', 'Unknown')}: Missing detection area coordinates")
                    if not has_image:
                        validation_issues.append(f"{automation.get('name', 'Unknown')}: Missing reference image")
                if not has_read_area:
                    validation_issues.append(f"{automation.get('name', 'Unknown')}: No target read area selected")
        
        if not has_valid_automation:
            if validation_issues:
                issues_text = "\n".join(validation_issues[:5])  # Show first 5 issues
                if len(validation_issues) > 5:
                    issues_text += f"\n... and {len(validation_issues) - 5} more"
                message = (
                    "Cannot Start Monitoring\n\n"
                    "Please set up at least one automation with:\n"
                    "- A detection area (use 'Set a detection area' button)\n"
                    "- A target read area (select from dropdown)\n\n"
                    f"Issues found:\n{issues_text}"
                )
            else:
                message = (
                    "Cannot Start Monitoring\n\n"
                    "Please set up at least one automation with:\n"
                    "- A detection area (use 'Set a detection area' button)\n"
                    "- A target read area (select from dropdown)"
                )
            messagebox.showwarning("Cannot Start Monitoring", message)
            return
        
        self.polling_active = True
        # Store state in game_text_reader for persistence
        self.game_text_reader._automations_polling_active = True
        # Update button state using the method to ensure consistency
        self._update_polling_button_state()
        
        # Show monitoring status in main window
        if hasattr(self.game_text_reader, '_show_monitoring_status_if_active'):
            self.game_text_reader._show_monitoring_status_if_active()
        
        # Update all automation status circles to show they're active (red/green instead of gray)
        for automation in self.automations:
            if automation.get('reference_image') and automation.get('image_area_coords'):
                # Trigger a status update to refresh circles from gray to red/green
                self.root.after(0, lambda a=automation: self.update_automation_status(a, False, False, 0, 0))
        
        def polling_loop():
            # Check shared state in game_text_reader as source of truth
            # This allows the thread to continue even when window is closed/reopened
            # IMPORTANT: Access automations through game_text_reader._automations_window
            # so we always use the current window instance's automations, not the old one
            while True:
                # Check shared state as source of truth
                should_continue = False
                if hasattr(self.game_text_reader, '_automations_polling_active'):
                    should_continue = self.game_text_reader._automations_polling_active
                
                if not should_continue:
                    print("POLLING: Stopping polling loop - shared state is False")
                    break
                try:
                    # Use current window instance's automations (works even if window was closed/reopened)
                    current_window = getattr(self.game_text_reader, '_automations_window', None)
                    if current_window and hasattr(current_window, 'check_all_automations'):
                        current_window.check_all_automations()
                    else:
                        # Fallback to self if current window not available
                        self.check_all_automations()
                    time.sleep(self.polling_interval)
                except Exception as e:
                    print(f"Error in polling loop: {e}")
                    import traceback
                    traceback.print_exc()
                    time.sleep(self.polling_interval)
        
        self.polling_thread = threading.Thread(target=polling_loop, daemon=True)
        self.polling_thread.start()
    
    def stop_polling(self):
        """Stop background polling"""
        print("POLLING: stop_polling() called")
        self.polling_active = False
        # Store state in game_text_reader for persistence (this is what the polling loop checks)
        self.game_text_reader._automations_polling_active = False
        print(f"POLLING: Set _automations_polling_active to False")
        # Update button state using the method to ensure consistency
        # Use root.after(0) to ensure it happens immediately on the main thread
        self.root.after(0, self._update_polling_button_state)
        
        # Clear monitoring status from main window if no other message is showing
        if hasattr(self.game_text_reader, 'status_label'):
            try:
                current_text = self.game_text_reader.status_label.cget('text')
                if current_text == "Area monitoring is active":
                    self.game_text_reader.status_label.config(text="", font=("Helvetica", 10))
            except:
                pass
        
        # Update all automation status circles to gray when monitoring stops
        for automation in self.automations:
            if automation.get('image_status_circle') or automation.get('text_status_circle'):
                # Update status to show gray circles
                self.root.after(0, lambda a=automation: self.update_automation_status(a, False, False, 0, 0))
    
    def check_all_automations(self):
        """Check all automation rules"""
        for automation in self.automations:
            if not automation.get('reference_image') or not automation.get('image_area_coords'):
                continue  # Skip if not configured
            
            self.check_automation(automation)
    
    def update_automation_status(self, automation, image_match, text_found, elapsed_ms, total_ms):
        """Update the status indicators for an automation"""
        try:
            # Check if window still exists before accessing widgets
            if not self.window.winfo_exists():
                return
            
            # Check shared state as source of truth for monitoring status
            is_monitoring_active = False
            if hasattr(self.game_text_reader, '_automations_polling_active'):
                is_monitoring_active = self.game_text_reader._automations_polling_active
            
            # Update image status circle
            if automation.get('image_status_circle'):
                try:
                    automation['image_status_circle'].delete("all")
                    if not is_monitoring_active:
                        color = "gray"  # Gray when monitoring not active
                    else:
                        color = "green" if image_match else "red"
                    automation['image_status_circle'].create_oval(2, 2, 10, 10, fill=color, outline="black", width=1)
                except:
                    pass  # Widget was destroyed
            
            # Update image status label with match percentage (always show, even if below threshold)
            if automation.get('image_status_label'):
                try:
                    match_percent = automation.get('_last_match_percent', 0)
                    if image_match:
                        # Above threshold - show in green
                        automation['image_status_label'].config(text=f"({match_percent:.1f}%)", fg="green")
                    else:
                        # Below threshold - show in red/orange so you can see how close you are
                        automation['image_status_label'].config(text=f"({match_percent:.1f}%)", fg="orange")
                except:
                    pass  # Widget was destroyed
            
            # Update text status circle
            if automation.get('text_status_circle'):
                try:
                    automation['text_status_circle'].delete("all")
                    if not is_monitoring_active:
                        color = "gray"  # Gray when monitoring not active
                    elif text_found is None:
                        color = "gray"  # Gray when text detection is disabled
                    else:
                        color = "green" if text_found else "red"
                    automation['text_status_circle'].create_oval(2, 2, 10, 10, fill=color, outline="black", width=1)
                except:
                    pass  # Widget was destroyed
            
            # Update text status label
            if automation.get('text_status_label'):
                try:
                    if text_found is None:
                        # Text detection disabled - show nothing
                        automation['text_status_label'].config(text="", fg="black")
                    elif text_found:
                        automation['text_status_label'].config(text="✓", fg="green")
                    else:
                        automation['text_status_label'].config(text="", fg="black")
                except:
                    pass  # Widget was destroyed
            
            # Update timer progress bar
            if automation.get('timer_progress_bar') and automation.get('timer_progress_label'):
                try:
                    if total_ms > 0:
                        progress = min(100, (elapsed_ms / total_ms) * 100)
                        automation['timer_progress_bar']['value'] = progress
                        automation['timer_progress_label'].config(text=f"{int(elapsed_ms)}/{int(total_ms)}ms")
                    else:
                        automation['timer_progress_bar']['value'] = 0
                        automation['timer_progress_label'].config(text="0ms")
                except:
                    pass  # Widgets were destroyed
        except Exception as e:
            print(f"Error updating automation status: {e}")
    
    def check_automation(self, automation):
        """Check a single automation rule"""
        try:
            # Capture current screen area
            x1, y1, x2, y2 = automation['image_area_coords']
            
            # TODO: Add freeze screen support here
            current_image = capture_screen_area(x1, y1, x2, y2)
            reference_image = automation['reference_image']
            
            # Compare images using selected method
            comparison_method = automation['comparison_method'].get()
            if comparison_method == "Pixel":
                match_percent = self.compare_images_pixel_by_pixel(current_image, reference_image)
            elif comparison_method == "Histogram":
                match_percent = self.compare_images_histogram(current_image, reference_image)
            elif comparison_method == "SSIM":
                match_percent = self.compare_images_ssim(current_image, reference_image)
            elif comparison_method == "Perceptual":
                match_percent = self.compare_images_perceptual(current_image, reference_image)
            elif comparison_method == "Edge":
                match_percent = self.compare_images_edge(current_image, reference_image)
            else:
                # Default to pixel comparison
                match_percent = self.compare_images_pixel_by_pixel(current_image, reference_image)
            
            threshold = automation['match_percent'].get()
            
            # Store match percent for status display
            automation['_last_match_percent'] = match_percent
            
            is_matching = match_percent >= threshold
            was_matching = automation.get('was_matching', False)
            
            # Check for text only if "Only read if text exists" is enabled
            # If disabled, skip text detection entirely and set to None (will show gray)
            text_found = None  # None means text detection is disabled
            if automation['only_read_if_text'].get():
                # Text detection is enabled - check for text as long as we're in matching state
                # This prevents stopping text detection when match % fluctuates slightly
                text_found = False
                if was_matching or is_matching:
                    # We're in matching state - check for text in the detection area
                    text_found = self.has_text_in_area(current_image)
            
            # Calculate timer progress
            elapsed_ms = 0
            total_ms = automation['read_after_ms'].get()
            if automation.get('timer_active') and automation.get('timer_start_time'):
                elapsed_ms = (time.time() - automation['timer_start_time']) * 1000
            
            # Update status indicators on main thread
            self.root.after(0, lambda: self.update_automation_status(automation, is_matching, text_found, elapsed_ms, total_ms))
            
            if is_matching:
                # Image matches - check if this is a new match (state transition)
                if not was_matching:
                    # State transition: not_matching -> matching
                    # Reset trigger flag and start timer
                    automation['has_triggered'] = False
                    automation['was_matching'] = True
                    
                    # Check if timer is already active (shouldn't be, but check anyway)
                    if automation['timer_active']:
                        # Timer is counting down - check if it's expired
                        read_after_ms = automation['read_after_ms'].get()
                        timer_elapsed_ms = (time.time() - automation['timer_start_time']) * 1000
                        
                        if timer_elapsed_ms >= read_after_ms:
                            # Timer expired - trigger read (only if not already triggered)
                            if not automation['has_triggered']:
                                self.trigger_read_area(automation)
                                automation['has_triggered'] = True
                            automation['timer_active'] = False
                            # Reset has_triggered to allow re-triggering if conditions are still met
                            automation['has_triggered'] = False
                    else:
                        # Start new timer if conditions are met
                        if automation['only_read_if_text'].get():
                            # "Only read if text exists" is ON: Require BOTH image match AND text found
                            # Use the text_found value we already checked above
                            if text_found:
                                # Both image and text are "green" - start timer and record when text was found
                                automation['timer_active'] = True
                                automation['timer_start_time'] = time.time()
                                automation['text_last_found_time'] = time.time()
                            # If no text, don't start timer (wait for both conditions)
                        else:
                            # "Only read if text exists" is OFF: Trigger only on image detection (image "green")
                            # No text check needed - start timer immediately when image matches
                            automation['timer_active'] = True
                            automation['timer_start_time'] = time.time()
                else:
                    # Still matching - check if we need to handle text requirement
                    if automation['only_read_if_text'].get():
                        # "Only read if text exists" is ON: Require BOTH image match AND text found (both "green")
                        if not text_found:
                            # Text not found - but be lenient: only cancel timer if text has been missing for >500ms
                            # This prevents timer reset from brief OCR misses
                            current_time = time.time()
                            if automation.get('text_last_found_time'):
                                time_since_text_found = (current_time - automation['text_last_found_time']) * 1000
                                if time_since_text_found > 500:  # 500ms tolerance for OCR misses
                                    # Text has been missing for >500ms - cancel timer
                                    if automation['timer_active']:
                                        automation['timer_active'] = False
                                        automation['timer_start_time'] = None
                            else:
                                # Text was never found - cancel timer immediately
                                if automation['timer_active']:
                                    automation['timer_active'] = False
                                    automation['timer_start_time'] = None
                            # Don't trigger if text is missing
                        else:
                            # Text found - update last found time
                            automation['text_last_found_time'] = time.time()
                            # Both image and text found - check timer if active
                            if automation['timer_active']:
                                # Timer is counting down - check if it's expired
                                timer_elapsed_ms = (time.time() - automation['timer_start_time']) * 1000
                                read_after_ms = automation['read_after_ms'].get()
                                
                                if timer_elapsed_ms >= read_after_ms:
                                    # Timer expired - trigger read
                                    if not automation['has_triggered']:
                                        self.trigger_read_area(automation)
                                        automation['has_triggered'] = True
                                    automation['timer_active'] = False
                                    # Reset has_triggered to allow re-triggering if conditions are still met
                                    automation['has_triggered'] = False
                            else:
                                # Timer not active yet but both conditions met - start timer
                                # (only if we haven't already triggered for this match state)
                                if not automation['has_triggered']:
                                    automation['timer_active'] = True
                                    automation['timer_start_time'] = time.time()
                                    automation['text_last_found_time'] = time.time()
                    else:
                        # "Only read if text exists" is OFF: Trigger only on image detection (image "green")
                        # No text check needed - just check timer if active
                        if automation['timer_active']:
                            # Timer is counting down - check if it's expired
                            timer_elapsed_ms = (time.time() - automation['timer_start_time']) * 1000
                            read_after_ms = automation['read_after_ms'].get()
                            
                            if timer_elapsed_ms >= read_after_ms:
                                # Timer expired - trigger read
                                if not automation['has_triggered']:
                                    self.trigger_read_area(automation)
                                    automation['has_triggered'] = True
                                automation['timer_active'] = False
                                # Reset has_triggered to allow re-triggering if conditions are still met
                                automation['has_triggered'] = False
                    # If timer not active and already triggered, start a new timer cycle if conditions are still met
                    if not automation['timer_active'] and automation['has_triggered']:
                        # Reset and start new cycle
                        automation['has_triggered'] = False
                        if automation['only_read_if_text'].get():
                            # "Only read if text exists" is ON: Require BOTH image match AND text found (with leniency for OCR misses)
                            if text_found:
                                automation['timer_active'] = True
                                automation['timer_start_time'] = time.time()
                                automation['text_last_found_time'] = time.time()
                            elif automation.get('text_last_found_time'):
                                # Text not found now, but check if it was found recently (within 500ms)
                                time_since_text_found = (time.time() - automation['text_last_found_time']) * 1000
                                if time_since_text_found <= 500:
                                    # Text was found recently - still allow timer to start
                                    automation['timer_active'] = True
                                    automation['timer_start_time'] = time.time()
                        else:
                            automation['timer_active'] = True
                            automation['timer_start_time'] = time.time()
            else:
                # Image doesn't match - state transition: matching -> not_matching
                if was_matching:
                    # Reset state for next match
                    automation['was_matching'] = False
                    automation['has_triggered'] = False
                    automation['text_last_found_time'] = None  # Reset text tracking
                
                # Cancel timer if active
                if automation['timer_active']:
                    automation['timer_active'] = False
        except Exception as e:
            print(f"Error checking automation {automation['id']}: {e}")
    
    def compare_images_pixel_by_pixel(self, img1, img2):
        """Compare two images pixel-by-pixel and return match percentage"""
        try:
            # Resize images to same size if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)
            
            # Convert to RGB if needed
            if img1.mode != 'RGB':
                img1 = img1.convert('RGB')
            if img2.mode != 'RGB':
                img2 = img2.convert('RGB')
            
            # Get pixel data
            pixels1 = list(img1.getdata())
            pixels2 = list(img2.getdata())
            
            if len(pixels1) != len(pixels2):
                return 0.0
            
            # Count matching pixels (within tolerance)
            matching_pixels = 0
            total_pixels = len(pixels1)
            tolerance = 10  # Allow small color differences
            
            for p1, p2 in zip(pixels1, pixels2):
                # Calculate color distance
                r_diff = abs(p1[0] - p2[0])
                g_diff = abs(p1[1] - p2[1])
                b_diff = abs(p1[2] - p2[2])
                
                # If all color channels are within tolerance, consider it a match
                if r_diff <= tolerance and g_diff <= tolerance and b_diff <= tolerance:
                    matching_pixels += 1
            
            match_percent = (matching_pixels / total_pixels) * 100.0
            return match_percent
        except Exception as e:
            print(f"Error comparing images: {e}")
            return 0.0
    
    def compare_images_histogram(self, img1, img2):
        """Compare images using color histogram - more forgiving to pixel shifts"""
        try:
            # Resize images to same size if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)
            
            # Convert to RGB if needed
            if img1.mode != 'RGB':
                img1 = img1.convert('RGB')
            if img2.mode != 'RGB':
                img2 = img2.convert('RGB')
            
            # Calculate histograms for each channel
            hist1_r = img1.histogram()[0:256]
            hist1_g = img1.histogram()[256:512]
            hist1_b = img1.histogram()[512:768]
            
            hist2_r = img2.histogram()[0:256]
            hist2_g = img2.histogram()[256:512]
            hist2_b = img2.histogram()[512:768]
            
            # Calculate correlation for each channel
            def histogram_correlation(h1, h2):
                """Calculate correlation coefficient between two histograms"""
                n = len(h1)
                if n == 0:
                    return 0.0
                
                mean1 = sum(h1) / n
                mean2 = sum(h2) / n
                
                numerator = sum((h1[i] - mean1) * (h2[i] - mean2) for i in range(n))
                denom1 = sum((h1[i] - mean1) ** 2 for i in range(n))
                denom2 = sum((h2[i] - mean2) ** 2 for i in range(n))
                
                if denom1 == 0 or denom2 == 0:
                    return 0.0
                
                return numerator / ((denom1 * denom2) ** 0.5)
            
            # Average correlation across all channels
            corr_r = histogram_correlation(hist1_r, hist2_r)
            corr_g = histogram_correlation(hist1_g, hist2_g)
            corr_b = histogram_correlation(hist1_b, hist2_b)
            
            # Convert correlation (-1 to 1) to match percentage (0 to 100)
            avg_corr = (corr_r + corr_g + corr_b) / 3.0
            match_percent = ((avg_corr + 1) / 2.0) * 100.0  # Scale from [-1,1] to [0,100]
            
            return max(0.0, min(100.0, match_percent))
        except Exception as e:
            print(f"Error comparing images with histogram: {e}")
            return 0.0
    
    def compare_images_ssim(self, img1, img2):
        """Compare images using SSIM-like structural similarity - best for games"""
        try:
            # Resize images to same size if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)
            
            # Convert to RGB if needed
            if img1.mode != 'RGB':
                img1 = img1.convert('RGB')
            if img2.mode != 'RGB':
                img2 = img2.convert('RGB')
            
            # Convert to grayscale for SSIM calculation (simpler and faster)
            img1_gray = img1.convert('L')
            img2_gray = img2.convert('L')
            
            # Get pixel data as arrays
            pixels1 = list(img1_gray.getdata())
            pixels2 = list(img2_gray.getdata())
            
            if len(pixels1) != len(pixels2):
                return 0.0
            
            # Calculate mean
            mean1 = sum(pixels1) / len(pixels1)
            mean2 = sum(pixels2) / len(pixels2)
            
            # Calculate variance and covariance
            var1 = sum((p - mean1) ** 2 for p in pixels1) / len(pixels1)
            var2 = sum((p - mean2) ** 2 for p in pixels2) / len(pixels2)
            covar = sum((pixels1[i] - mean1) * (pixels2[i] - mean2) for i in range(len(pixels1))) / len(pixels1)
            
            # SSIM constants
            c1 = (0.01 * 255) ** 2
            c2 = (0.03 * 255) ** 2
            
            # Calculate SSIM
            numerator = (2 * mean1 * mean2 + c1) * (2 * covar + c2)
            denominator = (mean1 ** 2 + mean2 ** 2 + c1) * (var1 + var2 + c2)
            
            if denominator == 0:
                return 0.0
            
            ssim = numerator / denominator
            # Convert SSIM (0 to 1) to match percentage (0 to 100)
            match_percent = ssim * 100.0
            
            return max(0.0, min(100.0, match_percent))
        except Exception as e:
            print(f"Error comparing images with SSIM: {e}")
            return 0.0
    
    def compare_images_perceptual(self, img1, img2):
        """Compare images using perceptual hash - very forgiving to minor changes"""
        try:
            # Resize images to same size if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)
            
            # Convert to RGB if needed
            if img1.mode != 'RGB':
                img1 = img1.convert('RGB')
            if img2.mode != 'RGB':
                img2 = img2.convert('RGB')
            
            # Resize to small size for hash comparison (8x8 is standard for pHash)
            size = 8
            img1_small = img1.resize((size, size), Image.Resampling.LANCZOS).convert('L')
            img2_small = img2.resize((size, size), Image.Resampling.LANCZOS).convert('L')
            
            # Calculate average pixel value
            pixels1 = list(img1_small.getdata())
            pixels2 = list(img2_small.getdata())
            
            avg1 = sum(pixels1) / len(pixels1)
            avg2 = sum(pixels2) / len(pixels2)
            
            # Create hash: 1 if pixel > average, 0 otherwise
            hash1 = [1 if p > avg1 else 0 for p in pixels1]
            hash2 = [1 if p > avg2 else 0 for p in pixels2]
            
            # Calculate Hamming distance (number of different bits)
            hamming_distance = sum(h1 != h2 for h1, h2 in zip(hash1, hash2))
            
            # Convert Hamming distance to similarity percentage
            # Maximum distance is size*size, similarity is inverse
            max_distance = size * size
            similarity = 1.0 - (hamming_distance / max_distance)
            match_percent = similarity * 100.0
            
            return max(0.0, min(100.0, match_percent))
        except Exception as e:
            print(f"Error comparing images with perceptual hash: {e}")
            return 0.0
    
    def compare_images_edge(self, img1, img2):
        """Compare images using edge detection - ignores colors, detects shapes/structures"""
        try:
            from PIL import ImageFilter
            
            # Resize images to same size if needed
            if img1.size != img2.size:
                img2 = img2.resize(img1.size, Image.Resampling.LANCZOS)
            
            # Convert to grayscale
            img1_gray = img1.convert('L')
            img2_gray = img2.convert('L')
            
            # Apply edge detection (FIND_EDGES filter)
            img1_edges = img1_gray.filter(ImageFilter.FIND_EDGES)
            img2_edges = img2_gray.filter(ImageFilter.FIND_EDGES)
            
            # Get pixel data
            pixels1 = list(img1_edges.getdata())
            pixels2 = list(img2_edges.getdata())
            
            if len(pixels1) != len(pixels2):
                return 0.0
            
            # Compare edge pixels (threshold to binary: edge or not)
            # Edges are typically bright pixels in FIND_EDGES output
            edge_threshold = 50  # Pixels brighter than this are considered edges
            
            matching_edges = 0
            total_edges = 0
            
            for p1, p2 in zip(pixels1, pixels2):
                is_edge1 = p1 > edge_threshold
                is_edge2 = p2 > edge_threshold
                
                # Only compare where at least one image has an edge
                if is_edge1 or is_edge2:
                    total_edges += 1
                    if is_edge1 == is_edge2:
                        matching_edges += 1
            
            if total_edges == 0:
                # No edges in either image - consider it a match if both are blank
                return 100.0
            
            match_percent = (matching_edges / total_edges) * 100.0
            return max(0.0, min(100.0, match_percent))
        except Exception as e:
            print(f"Error comparing images with edge detection: {e}")
            return 0.0
    
    def has_text_in_area(self, image):
        """Check if text exists in the image using OCR"""
        try:
            # Use basic OCR to detect text
            text = pytesseract.image_to_string(image, config='--psm 6')
            # Remove whitespace and check if any text remains
            text = text.strip()
            return len(text) > 0
        except Exception as e:
            print(f"Error checking for text: {e}")
            return False
    
    def trigger_read_area(self, automation):
        """Trigger the selected read area, automation, or area combo"""
        try:
            target_name = automation['target_read_area'].get()
            
            if not target_name or target_name in ["No areas available", "No options available"]:
                return
            
            # Skip separator entries
            if target_name.startswith("───"):
                return
            
            # Check if it's a regular area
            target_area_frame = None
            area_name_var = None
            set_area_button = None
            for area in self.game_text_reader.areas:
                area_name = area[3].get()  # area_name_var
                if area_name == target_name:
                    target_area_frame = area[0]  # area_frame
                    # Get area_name_var and set_area_button for Auto Read areas
                    if len(area) >= 9:
                        _, _, set_area_button, area_name_var, _, _, _, _, _ = area[:9]
                    elif len(area) >= 8:
                        _, _, set_area_button, area_name_var, _, _, _, _ = area[:8]
                    break
            
            if target_area_frame:
                # Check if this is an Auto Read area - if so, trigger area selection dialog
                if target_name.startswith("Auto Read"):
                    print(f"Automation triggered read area: {target_name} (Auto Read - opening area selection dialog)")
                    automation_name = automation.get('name', 'Unknown')
                    self._update_main_status(f"Automation '{automation_name}' → Auto Read: {target_name}", "blue", 2000)
                    if area_name_var and target_area_frame:
                        # Trigger area selection dialog (same as pressing the Auto Read hotkey)
                        # Note: set_area_button is None for Auto Read areas, which is fine - set_auto_read_area doesn't use it
                        self.root.after(0, lambda f=target_area_frame, n=area_name_var, s=set_area_button: self.game_text_reader.set_auto_read_area(f, n, s))
                    else:
                        print(f"Automation: Warning - Could not find area info for Auto Read area '{target_name}'")
                else:
                    # Regular area - just read it
                    automation_name = automation.get('name', 'Unknown')
                    self._update_main_status(f"Automation '{automation_name}' → Reading: {target_name}", "blue", 2000)
                    self.root.after(0, lambda: self.game_text_reader.read_area(target_area_frame))
                    print(f"Automation triggered read area: {target_name}")
                return
            
            # Check if it's an automation
            for other_automation in self.automations:
                if other_automation['name'] == target_name:
                    # Trigger the automation by checking it
                    print(f"Automation triggered automation: {target_name}")
                    automation_name = automation.get('name', 'Unknown')
                    self._update_main_status(f"Automation '{automation_name}' → Triggering: {target_name}", "blue", 2000)
                    self.root.after(0, lambda a=other_automation: self.check_automation(a))
                    return
            
            # Check if it's an area combo
            for combo in self.hotkey_combos:
                if combo['name'] == target_name:
                    # Check if combo is already running before triggering
                    is_triggering = combo.get('is_triggering', False)
                    print(f"Automation: Checking combo '{target_name}' - is_triggering={is_triggering}")
                    if is_triggering:
                        print(f"Automation: Combo '{target_name}' is already running, ignoring trigger")
                        self._update_main_status(f"{target_name} already running", "orange", 2000)
                        return
                    # Trigger the area combo
                    print(f"Automation triggered area combo: {target_name}")
                    automation_name = automation.get('name', 'Unknown')
                    self._update_main_status(f"Automation '{automation_name}' → Combo '{target_name}'", "blue", 2000)
                    self.root.after(0, lambda c=combo: self.trigger_hotkey_combo(c))
                    return
            
            print(f"Could not find trigger target: {target_name}")
        except Exception as e:
            print(f"Error triggering read area: {e}")


