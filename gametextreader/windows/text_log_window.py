"""
Text log window for viewing scan history and repeating text
"""
import asyncio
import os
import re
import time
import tkinter as tk
from tkinter import messagebox, ttk

from ..utils import _ensure_uwp_available


class TextLogWindow:
    def __init__(self, root, game_text_reader):
        self.root = root
        self.game_text_reader = game_text_reader
        self.window = tk.Toplevel(root)
        self.window.title("Scan History")
        self.window.geometry("700x300")
        self.window.resizable(True, True)
        
        # Register this window as one that disables hotkeys
        self.game_text_reader.register_hotkey_disabling_window("Scan History", self.window)
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting Scan History window icon: {e}")
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (700 // 2)
        y = (self.window.winfo_screenheight() // 2) - (300 // 2)
        self.window.geometry(f"700x300+{x}+{y}")
        
        # Create main frame
        main_frame = tk.Frame(self.window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        
        # Hotkey button frame for repeating latest area text
        hotkey_frame = tk.Frame(main_frame)
        hotkey_frame.pack(fill='x', pady=(0, 12))
        
        # Title label on the left
        title_label = tk.Label(hotkey_frame, text="History", 
                              font=("Helvetica", 12, "bold"))
        title_label.pack(side='left')
        
        # Frame for the hotkey controls on the right
        hotkey_controls_frame = tk.Frame(hotkey_frame)
        hotkey_controls_frame.pack(side='right')
        
        # Label for hotkey button
        hotkey_label = tk.Label(hotkey_controls_frame, text="Repeat latest scan:", 
                               font=("Helvetica", 10))
        hotkey_label.pack(side='left', padx=(0, 10))
        
        # Button to set hotkey (use persistent button from GameTextReader for hotkey registration)
        # Matches the style of hotkey buttons in the main window
        self.repeat_latest_hotkey_button = tk.Button(
            hotkey_controls_frame,
            text="Set Hotkey",
            command=lambda: self.game_text_reader.set_hotkey(self.game_text_reader.repeat_latest_hotkey_button, None)
        )
        self.repeat_latest_hotkey_button.pack(side='left')
        
        # Sync the persistent button's hotkey
        self.game_text_reader.repeat_latest_hotkey_button.hotkey = self.game_text_reader.repeat_latest_hotkey
        
        # Store reference to display button for updating
        self.game_text_reader.repeat_latest_hotkey_button._display_button = self.repeat_latest_hotkey_button
        
        # Override the config method on the persistent button to update the display button
        # Also convert "Set Hotkey:" to "Hotkey:" to match main window format
        def update_display_button(**kwargs):
            if hasattr(self.game_text_reader.repeat_latest_hotkey_button, '_display_button'):
                display_btn = self.game_text_reader.repeat_latest_hotkey_button._display_button
                if hasattr(display_btn, 'config'):
                    # Convert "Set Hotkey:" to "Hotkey:" to match main window format
                    # Only convert when hotkey is actually set (text contains brackets with content)
                    if 'text' in kwargs:
                        text = kwargs['text']
                        if text.startswith("Set Hotkey:") and "[" in text and "]" in text:
                            # Check if there's actual hotkey content (not just empty brackets or preview)
                            match = re.search(r'\[([^\]]+)\]', text)
                            if match and match.group(1).strip() and not match.group(1).strip().endswith("+"):
                                # Has actual hotkey content, convert to "Hotkey:"
                                kwargs['text'] = text.replace("Set Hotkey:", "Hotkey:", 1)
                    display_btn.config(**kwargs)
        
        self.game_text_reader.repeat_latest_hotkey_button.config = update_display_button
        
        # Restore hotkey if it was previously set
        if self.game_text_reader.repeat_latest_hotkey:
            # Update button text to match main window format: "Hotkey: [ {display_name.upper()} ]"
            display_name = self.game_text_reader._hotkey_to_display_name(self.game_text_reader.repeat_latest_hotkey)
            self.repeat_latest_hotkey_button.config(text=f"Hotkey: [ {display_name.upper()} ]")
            # Only set up the hotkey if it's not already registered
            # This prevents removing and re-registering an already working hotkey
            if not hasattr(self.game_text_reader.repeat_latest_hotkey_button, 'keyboard_hook') and not hasattr(self.game_text_reader.repeat_latest_hotkey_button, 'mouse_hook_id'):
                # Set up the hotkey using the persistent button
                self.game_text_reader.setup_hotkey(self.game_text_reader.repeat_latest_hotkey_button, None)
        
        # Add separator line under hotkey button
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.pack(fill='x', pady=(10, 10))
        
        # Create scrollable frame
        canvas_frame = tk.Frame(main_frame)
        canvas_frame.pack(fill='both', expand=True)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind mousewheel to canvas - only when mouse is over canvas
        def _on_canvas_mousewheel(event):
            if canvas.bbox('all') and canvas.winfo_height() < (canvas.bbox('all')[3] - canvas.bbox('all')[1]):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        
        def _bind_canvas_mousewheel(event):
            canvas.bind_all('<MouseWheel>', _on_canvas_mousewheel)
        
        def _unbind_canvas_mousewheel(event):
            canvas.unbind_all('<MouseWheel>')
        
        canvas.bind('<Enter>', _bind_canvas_mousewheel)
        canvas.bind('<Leave>', _unbind_canvas_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Store canvas for updates
        self.canvas = canvas
        self.scrollable_frame_inner = self.scrollable_frame
        
        # Track last history count for auto-refresh
        self.last_history_count = 0
        
        # Store entry widgets for voice updates
        self.entry_widgets = []
        
        # Update display
        self.update_display()
        
        # Set up protocol to handle window closing
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Start auto-refresh
        self.auto_refresh()
    
    def on_close(self):
        """Handle window closing."""
        # Stop auto-refresh
        if hasattr(self, '_auto_refresh_id'):
            self.window.after_cancel(self._auto_refresh_id)
        # Clear display button reference (hotkey will still work via persistent button)
        if hasattr(self.game_text_reader, 'repeat_latest_hotkey_button'):
            self.game_text_reader.repeat_latest_hotkey_button._display_button = None
            # Reset config to dummy method
            self.game_text_reader.repeat_latest_hotkey_button.config = lambda **kwargs: None
        # Unregister this window and clear reference in game_text_reader
        self.game_text_reader.unregister_hotkey_disabling_window("Scan History")
        if hasattr(self.game_text_reader, 'text_log_window'):
            self.game_text_reader.text_log_window = None
        self.window.destroy()
    
    def auto_refresh(self):
        """Auto-refresh the display when new entries are added."""
        if not self.window.winfo_exists():
            return
        
        current_count = len(self.game_text_reader.text_log_history)
        if current_count != self.last_history_count:
            self.update_display()
            self.last_history_count = current_count
        
        # Schedule next check (every 500ms)
        self._auto_refresh_id = self.window.after(500, self.auto_refresh)
    
    def update_display(self):
        """Update the display with current Scan History history."""
        # Clear existing widgets
        for widget in self.scrollable_frame_inner.winfo_children():
            widget.destroy()
        
        # Clear entry widgets list
        self.entry_widgets = []
        
        # Get reversed history (most recent first)
        history = list(reversed(self.game_text_reader.text_log_history))
        
        if not history:
            no_text_label = tk.Label(self.scrollable_frame_inner, 
                                    text="No converted texts yet.", 
                                    font=("Helvetica", 10),
                                    fg="gray")
            no_text_label.pack(pady=20)
            return
        
        # Create header
        header_frame = tk.Frame(self.scrollable_frame_inner)
        header_frame.pack(fill='x', padx=5, pady=5)
        tk.Label(header_frame, text="Area", font=("Helvetica", 10, "bold"), width=15, anchor='w').pack(side='left', padx=5)
        tk.Label(header_frame, text="Voice", font=("Helvetica", 10, "bold"), width=25, anchor='w').pack(side='left', padx=5)
        tk.Label(header_frame, text="Text", font=("Helvetica", 10, "bold"), anchor='w').pack(side='left', padx=5, fill='x', expand=True)
        tk.Label(header_frame, text="Action", font=("Helvetica", 10, "bold"), width=10, anchor='w').pack(side='left', padx=5)
        
        # Separator
        ttk.Separator(self.scrollable_frame_inner, orient='horizontal').pack(fill='x', padx=5, pady=2)
        
        # Add entries
        for entry in history:
            self.create_entry_row(entry)
        
        # Update canvas scroll region
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Update last history count
        self.last_history_count = len(history)
    
    def create_entry_row(self, entry):
        """Create a row for a Scan History entry."""
        row_frame = tk.Frame(self.scrollable_frame_inner)
        row_frame.pack(fill='x', padx=5, pady=3)
        
        # Area name label
        area_name = entry.get('area_name', 'Unknown')
        area_label = tk.Label(row_frame, text=area_name, width=10, anchor='w', 
                             font=("Helvetica", 9))
        area_label.pack(side='left', padx=5)
        
        # Voice dropdown
        voice_frame = tk.Frame(row_frame)
        voice_frame.pack(side='left', padx=5)
        
        voice_var = tk.StringVar()
        voice_full_names = {}
        voice_display_names = []
        
        # Get available voices
        if hasattr(self.game_text_reader, 'voices') and self.game_text_reader.voices:
            try:
                for i, voice in enumerate(self.game_text_reader.voices, 1):
                    full_name = voice.GetDescription()
                    
                    # Create abbreviated display name
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
                print(f"Error getting voices for dropdown: {e}")
        
        # Find current voice in list
        current_voice = entry.get('voice', None)
        selected_display = "Select Voice"
        if current_voice:
            # Try to find matching voice
            for display_name, full_name in voice_full_names.items():
                if full_name == current_voice:
                    selected_display = display_name
                    break
            # If not found, show the full name
            if selected_display == "Select Voice" and current_voice:
                selected_display = current_voice
        
        voice_var.set(selected_display)
        voice_var._full_name = current_voice if current_voice else None
        
        # Create dropdown
        voice_menu = tk.OptionMenu(voice_frame, voice_var, *voice_display_names if voice_display_names else ["No voices available"])
        voice_menu.config(width=22, anchor="w")  # Left-align text like in main window
        voice_menu.pack(side='left')
        
        # Update entry when voice changes
        def on_voice_change(*args):
            selected_display = voice_var.get()
            if selected_display in voice_full_names:
                voice_var._full_name = voice_full_names[selected_display]
            else:
                voice_var._full_name = selected_display
            # Update the entry
            entry['voice'] = voice_var._full_name
        
        voice_var.trace('w', on_voice_change)
        
        # Text widget (editable for copying)
        text = entry.get('text', '')
        text_frame = tk.Frame(row_frame)
        text_frame.pack(side='left', padx=5)
        
        text_widget = tk.Text(text_frame, height=2, width=40, wrap=tk.WORD, 
                            font=("Helvetica", 9), relief=tk.SUNKEN, borderwidth=1)
        text_widget.insert('1.0', text)
        text_widget.config(state=tk.NORMAL)  # Allow editing/copying
        text_widget.pack(fill='y')  # Only fill vertically, not horizontally
        
        # Enable mouse wheel scrolling for text widget only if it needs scrolling
        # Store reference to canvas for fallback scrolling
        def _on_text_mousewheel(event):
            # Check if the text widget actually needs scrolling
            try:
                # Get scroll position info
                first_visible = text_widget.index('@0,0')
                last_visible = text_widget.index('@0,%d' % text_widget.winfo_height())
                end_index = text_widget.index('end-1c')
                
                first_line = float(first_visible.split('.')[0])
                last_line = float(last_visible.split('.')[0])
                total_lines = float(end_index.split('.')[0])
                
                # Check if we can scroll (content extends beyond visible area)
                can_scroll_down = last_line < total_lines
                can_scroll_up = first_line > 1.0
                
                # Only handle scroll if content is scrollable
                if can_scroll_down or can_scroll_up:
                    text_widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
                    return "break"
            except Exception:
                # If check fails, try to scroll anyway
                try:
                    text_widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
                    return "break"
                except Exception:
                    pass
            # If no scrolling needed, don't capture the event
            # Don't return "break" so event can propagate to canvas
            return None
        
        # Use widget-specific binding instead of bind_all to avoid conflicts
        # This way, the canvas binding can still work when text widget doesn't need scrolling
        text_widget.bind('<MouseWheel>', _on_text_mousewheel)
        
        # Right-click context menu for text widget
        def show_text_context_menu(event):
            context_menu = tk.Menu(text_widget, tearoff=0)
            context_menu.add_command(label="Select All", 
                                  command=lambda: text_widget.tag_add('sel', '1.0', 'end'))
            context_menu.add_command(label="Copy", 
                                  command=lambda: text_widget.event_generate('<<Copy>>'))
            context_menu.add_command(label="Paste", 
                                  command=lambda: text_widget.event_generate('<<Paste>>'))
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
        
        text_widget.bind("<Button-3>", show_text_context_menu)
        
        # Update entry when text changes
        def on_text_change(event=None):
            entry['text'] = text_widget.get('1.0', tk.END).rstrip('\n')
        
        text_widget.bind('<KeyRelease>', on_text_change)
        text_widget.bind('<FocusOut>', on_text_change)
        
        # Play button
        play_button = tk.Button(row_frame, text="▶ Play", width=10,
                               command=lambda e=entry: self.play_text(e))
        play_button.pack(side='left', padx=5)
        
        # Store widgets for reference
        self.entry_widgets.append({
            'entry': entry,
            'voice_var': voice_var,
            'text_widget': text_widget,
            'row_frame': row_frame
        })
    
    def play_text(self, entry):
        """Play the text using the voice assigned to its area or the selected voice in the dropdown."""
        # Get text from entry (may have been edited)
        text = entry.get('text', '').strip()
        area_name = entry.get('area_name', '')
        # Use the voice from entry (may have been changed in dropdown)
        voice_name = entry.get('voice', None)
        speed_value = entry.get('speed', None)
        
        if not text.strip():
            return
        
        # Stop any current speech
        self.game_text_reader.stop_speaking()
        
        # Find the area and get its voice/speed settings
        # First try to find in areas list
        voice_var = None
        speed_var = None
        
        for area in self.game_text_reader.areas:
            # Handle both 8 and 9 element tuples (for backward compatibility)
            if len(area) >= 9:
                area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var_item, speed_var_item, psm_var, freeze_screen_var = area[:9]
            else:
                area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var_item, speed_var_item, psm_var = area[:8]
            if area_name_var.get() == area_name:
                voice_var = voice_var_item
                speed_var = speed_var_item
                break
        
        # Check if no voice is selected
        has_voice = False
        invalid_voices = ["Select Voice", "No voices available"]
        
        # Check voice from entry
        if voice_name and voice_name.strip() and voice_name not in invalid_voices:
            has_voice = True
        # Check voice from area
        if not has_voice and voice_var:
            actual_voice_name = getattr(voice_var, '_full_name', voice_var.get())
            if actual_voice_name and actual_voice_name.strip() and actual_voice_name not in invalid_voices:
                has_voice = True
        
        # If no voice is selected, show error
        if not has_voice:
            messagebox.showerror("Error", "No voice selected. Please select a voice.")
            print("Error: Did not speak, Reason: No voice selected in scan history window.")
            return
        
        # Prioritize voice from entry (scan history window selection) over area voice
        # First try to use the voice from entry if it's valid
        voice_to_use = None
        if voice_name and voice_name.strip() and voice_name not in invalid_voices:
            voice_to_use = voice_name
        # Fall back to area voice if entry doesn't have a valid voice
        elif voice_var:
            voice_to_use = getattr(voice_var, '_full_name', voice_var.get())
            if not voice_to_use or not voice_to_use.strip() or voice_to_use in invalid_voices:
                voice_to_use = None
        
        # Set the voice if we have one
        if voice_to_use:
            selected_voice = None
            
            # Try to find voice in available voices
            try:
                voices = self.game_text_reader.speaker.GetVoices()
                for voice in voices:
                    if voice.GetDescription() == voice_to_use:
                        selected_voice = voice
                        break
            except Exception:
                pass
            
            # If not found in SAPI, try combined voice list
            if not selected_voice and hasattr(self.game_text_reader, 'voices'):
                for voice in self.game_text_reader.voices:
                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == voice_to_use:
                        if hasattr(voice, 'GetId'):
                            selected_voice = voice
                            break
            
            # Set voice if found
            if selected_voice and selected_voice != "mock_voice":
                try:
                    if hasattr(selected_voice, 'GetToken'):
                        # OneCore voice - use UWP
                        if _ensure_uwp_available():
                            loop = None
                            old_loop = None
                            try:
                                loop = asyncio.new_event_loop()
                                old_loop = asyncio.get_event_loop()
                                asyncio.set_event_loop(loop)
                                loop.run_until_complete(self.game_text_reader._speak_with_uwp(text, preferred_desc=voice_to_use))
                                return
                            except Exception as e:
                                print(f"UWP fallback failed: {e}")
                            finally:
                                # Restore previous event loop
                                if old_loop is not None:
                                    try:
                                        asyncio.set_event_loop(old_loop)
                                    except Exception:
                                        pass
                                # Only close loop if it was created and run successfully
                                if loop is not None:
                                    try:
                                        # Only close if loop was actually started (run_until_complete was called)
                                        if not loop.is_closed() and loop.is_running() == False:
                                            # Check if loop has been run by checking if it's in a valid state
                                            try:
                                                loop.close()
                                            except RuntimeError as e:
                                                if "run loop not started" not in str(e).lower():
                                                    raise
                                    except Exception:
                                        pass
                    else:
                        # Regular SAPI voice
                        self.game_text_reader.speaker.Voice = selected_voice
                except Exception as e:
                    print(f"Error setting voice: {e}")
            elif selected_voice == "mock_voice" and _ensure_uwp_available():
                # Mock voice - use UWP
                loop = None
                old_loop = None
                try:
                    loop = asyncio.new_event_loop()
                    try:
                        old_loop = asyncio.get_event_loop()
                    except RuntimeError:
                        old_loop = None
                    asyncio.set_event_loop(loop)
                    loop.run_until_complete(self.game_text_reader._speak_with_uwp(text, preferred_desc=voice_to_use))
                    return
                except Exception as e:
                    print(f"UWP fallback failed: {e}")
                finally:
                    # Restore previous event loop
                    if old_loop is not None:
                        try:
                            asyncio.set_event_loop(old_loop)
                        except Exception:
                            pass
                    # Only close loop if it was created and run successfully
                    if loop is not None:
                        try:
                            # Only close if loop was actually started (run_until_complete was called)
                            if not loop.is_closed():
                                try:
                                    loop.close()
                                except RuntimeError as e:
                                    if "run loop not started" not in str(e).lower():
                                        raise
                        except Exception:
                            pass
        
        # Set speed if available
        if speed_var:
            try:
                speed = int(speed_var.get())
                if speed > 0:
                    self.game_text_reader.speaker.Rate = (speed - 100) // 10
            except ValueError:
                pass
        elif speed_value is not None:
            try:
                if speed_value > 0:
                    self.game_text_reader.speaker.Rate = (speed_value - 100) // 10
            except ValueError:
                pass
        
        # Set volume
        try:
            vol = int(self.game_text_reader.volume.get())
            if 0 <= vol <= 100:
                self.game_text_reader.speaker.Volume = vol
            else:
                self.game_text_reader.speaker.Volume = 100
        except ValueError:
            self.game_text_reader.speaker.Volume = 100
        
        # Speak the text
        try:
            self.game_text_reader._ensure_speech_ready()
            # Track text and start time for pause/resume functionality
            self.game_text_reader.current_speech_text = text
            self.game_text_reader.speech_start_time = time.time()
            self.game_text_reader.paused_text = None
            self.game_text_reader.paused_position = 0
            self.game_text_reader.is_speaking = True
            self.game_text_reader.speaker.Speak(text, 1)  # 1 is SVSFlagsAsync
            print(f"Playing text from area '{area_name}'")
            # Start monitoring speech completion
            self.game_text_reader._start_speech_monitor()
        except Exception as e:
            print(f"Error playing text: {e}")
            self.game_text_reader.is_speaking = False
    
    def repeat_latest_area_text(self):
        """Repeat the latest area text from the Scan History."""
        if not hasattr(self.game_text_reader, 'text_log_history') or not self.game_text_reader.text_log_history:
            print("No Scan History available to repeat.")
            return
        
        # Get the latest entry (last one in the list)
        latest_entry = self.game_text_reader.text_log_history[-1]
        
        # Play the text using the existing play_text method
        self.play_text(latest_entry)

