"""
Image processing settings window for adjusting OCR preprocessing parameters
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageEnhance, ImageFilter, ImageTk

from ..image_processing import preprocess_image


class ImageProcessingWindow:
    def __init__(self, root, area_name, latest_images, settings, game_text_reader):
        self.window = tk.Toplevel(root)
        self.window.title(f"Image Processing for: {area_name}")
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting image processing window icon: {e}")
        
        self.area_name = area_name
        self.latest_images = latest_images
        self.settings = settings
        self.game_text_reader = game_text_reader
        
        # Set up protocol to re-enable hotkeys when window closes
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # Check if there is an image for the area
        if area_name not in latest_images:
            messagebox.showerror("Error", "No image to process, generate an image by pressing the hotkey.")
            # Unregister this window before destroying window since on_close won't be called
            self.game_text_reader.unregister_hotkey_disabling_window("Image Processing")
            self.window.destroy()
            return

        self.image = latest_images[area_name]
        self.processed_image = self.image.copy()

        # Add note about hotkeys being disabled
        hotkey_note = ttk.Label(self.window, text="Note: Hotkeys (including controller hotkeys) are disabled while this window is open.", 
                               font=("Helvetica", 10, "bold"), foreground='#666666')
        hotkey_note.grid(row=0, column=0, columnspan=5, padx=10, pady=(10, 5), sticky='w')
        
        # Register this window as one that disables hotkeys
        self.game_text_reader.register_hotkey_disabling_window("Image Processing", self.window)
        
        # Create a canvas to display the image
        self.image_frame = ttk.Frame(self.window)
        self.image_frame.grid(row=1, column=0, columnspan=5, padx=10, pady=5)
        self.canvas = tk.Canvas(self.image_frame, width=self.image.width, height=self.image.height)
        self.canvas.pack()

        # Display the image on the canvas
        self.photo_image = ImageTk.PhotoImage(self.image)
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        # Add a label under the image with larger text - centered
        info_text = f"Showing previous image captured in area: {area_name}\n\nProcessing applies to unprocessed images; results may differ if the preview is already processed."
        info_label = ttk.Label(self.image_frame, text=info_text, font=("Helvetica", 12), justify='center')
        info_label.pack(pady=(10, 0), fill='x')

        # Create a frame for bottom controls
        control_frame = ttk.Frame(self.window)
        control_frame.grid(row=2, column=0, columnspan=5, pady=10)

        # Add scale dropdown
        scale_frame = ttk.Frame(control_frame)
        scale_frame.pack(side='left', padx=10)
        
        ttk.Label(scale_frame, text="Preview Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]
        scale_menu = tk.OptionMenu(scale_frame, self.scale_var, *scales, command=self.update_preview)
        scale_menu.pack(side='left')
        ttk.Label(scale_frame, text="%").pack(side='left')

        # Add buttons
        ttk.Button(control_frame, text="Apply img. processing", command=self.save_settings).pack(side='left', padx=10)
        ttk.Button(control_frame, text="Reset to default", command=self.reset_all).pack(side='left', padx=10)

        # Add sliders for image processing
        self.brightness_var = tk.DoubleVar(value=settings.get('brightness', 1.0))
        self.contrast_var = tk.DoubleVar(value=settings.get('contrast', 1.0))
        self.saturation_var = tk.DoubleVar(value=settings.get('saturation', 1.0))
        self.sharpness_var = tk.DoubleVar(value=settings.get('sharpness', 1.0))
        self.blur_var = tk.DoubleVar(value=settings.get('blur', 0.0))
        self.threshold_var = tk.IntVar(value=settings.get('threshold', 128))
        self.hue_var = tk.DoubleVar(value=settings.get('hue', 0.0))
        self.exposure_var = tk.DoubleVar(value=settings.get('exposure', 1.0))
        self.threshold_enabled_var = tk.BooleanVar(value=settings.get('threshold_enabled', False))

        self.create_slider("Brightness", self.brightness_var, 0.1, 2.0, 1.0, 3, 0)
        self.create_slider("Contrast", self.contrast_var, 0.1, 2.0, 1.0, 3, 1)
        self.create_slider("Saturation", self.saturation_var, 0.1, 2.0, 1.0, 3, 2)
        self.create_slider("Sharpness", self.sharpness_var, 0.1, 2.0, 1.0, 3, 3)
        self.create_slider("Blur", self.blur_var, 0.0, 10.0, 0.0, 3, 4)
        self.create_slider("Threshold", self.threshold_var, 0, 255, 128, 4, 0, self.threshold_enabled_var)
        self.create_slider("Hue", self.hue_var, -1.0, 1.0, 0.0, 4, 1)
        self.create_slider("Exposure", self.exposure_var, 0.1, 2.0, 1.0, 4, 2)

    def create_slider(self, label, variable, from_, to, initial, row, col, enabled_var=None):
        frame = ttk.Frame(self.window)
        frame.grid(row=row, column=col, padx=10, pady=5)

        # Use a label frame for consistent structure
        label_frame = ttk.LabelFrame(frame, text=label)
        label_frame.pack(fill='both', expand=True)
    
        ttk.Label(label_frame, text=label).pack()

        entry_var = tk.StringVar(value=f'{initial:.2f}')
        # Add trace to variable to update entry field
        variable.trace_add('write', lambda *args: entry_var.set(f'{variable.get():.2f}'))

        slider = ttk.Scale(label_frame, from_=from_, to=to, orient='horizontal', variable=variable, command=self.update_image)
        slider.set(initial)
        slider.pack()

        # Create entry with context menu
        entry = ttk.Entry(label_frame, textvariable=entry_var)
        entry.pack()
        
        # Add context menu for copy/paste
        entry_menu = tk.Menu(entry, tearoff=0)
        entry_menu.add_command(label="Cut", command=lambda: entry.event_generate('<<Cut>>'))
        entry_menu.add_command(label="Copy", command=lambda: entry.event_generate('<<Copy>>'))
        entry_menu.add_command(label="Paste", command=lambda: entry.event_generate('<<Paste>>'))
        entry_menu.add_separator()
        entry_menu.add_command(label="Select All", command=lambda: entry.selection_range(0, 'end'))
        
        def show_entry_menu(event):
            entry_menu.post(event.x_root, event.y_root)
            
        entry.bind('<Button-3>', show_entry_menu)
    
        ttk.Button(label_frame, text="Reset", command=lambda: self.reset_slider(slider, entry, initial, variable)).pack()

        # Create checkbox for threshold slider
        if label == "Threshold":
            checkbox_frame = ttk.Frame(label_frame)
            checkbox_frame.pack(anchor='w')
        
            checkbox = ttk.Checkbutton(checkbox_frame, variable=enabled_var, command=self.update_image)
            checkbox.pack(side=tk.LEFT)

            ttk.Label(checkbox_frame, text="Enabled").pack(side=tk.LEFT, padx=(5, 0))

        setattr(self, f"{label.lower()}_slider", frame)
        frame.slider, frame.entry = slider, entry
        
    
    def reset_slider(self, slider, entry, initial, variable):
        slider.set(initial)
        variable.set(initial)
        entry.delete(0, tk.END)
        entry.insert(0, str(round(float(initial), 2)))
        self.update_image()
        

    def reset_all(self):
        self.brightness_var.set(1.0)
        self.contrast_var.set(1.0)
        self.saturation_var.set(1.0)
        self.sharpness_var.set(1.0)
        self.blur_var.set(0.0)
        self.threshold_var.set(128)
        self.hue_var.set(0.0)
        self.exposure_var.set(1.0)
        self.threshold_enabled_var.set(False)
        self.update_image()


    def update_image(self, _=None):
        if self.image:
            # Clean up previous processed image if it exists
            if self.processed_image:
                self.processed_image.close()
            self.processed_image = self.image.copy()

            # Apply brightness
            enhancer = ImageEnhance.Brightness(self.processed_image)
            self.processed_image = enhancer.enhance(self.brightness_var.get())

            # Apply contrast
            enhancer = ImageEnhance.Contrast(self.processed_image)
            self.processed_image = enhancer.enhance(self.contrast_var.get())

            # Apply saturation
            enhancer = ImageEnhance.Color(self.processed_image)
            self.processed_image = enhancer.enhance(self.saturation_var.get())

            # Apply sharpness
            enhancer = ImageEnhance.Sharpness(self.processed_image)
            self.processed_image = enhancer.enhance(self.sharpness_var.get())

            # Apply blur
            if self.blur_var.get() > 0:
                self.processed_image = self.processed_image.filter(ImageFilter.GaussianBlur(self.blur_var.get()))

            # Apply threshold if enabled
            if self.threshold_enabled_var.get():
                self.processed_image = self.processed_image.point(lambda p: p > self.threshold_var.get() and 255)

            # Apply hue (simplified, for demonstration purposes)
            self.processed_image = self.processed_image.convert('HSV')
            channels = list(self.processed_image.split())
            channels[0] = channels[0].point(lambda p: (p + int(self.hue_var.get() * 255)) % 256)
            self.processed_image = Image.merge('HSV', channels).convert('RGB')

            # Apply exposure (simplified, for demonstration purposes)
            enhancer = ImageEnhance.Brightness(self.processed_image)
            self.processed_image = enhancer.enhance(self.exposure_var.get())

            # Clean up previous photo_image if it exists
            if self.photo_image:
                del self.photo_image
            self.photo_image = ImageTk.PhotoImage(self.processed_image)
            self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)

    def save_settings(self):
        # First, update all settings in the processing_settings dictionary
        self.settings['brightness'] = self.brightness_var.get()
        self.settings['contrast'] = self.contrast_var.get()
        self.settings['saturation'] = self.saturation_var.get()
        self.settings['sharpness'] = self.sharpness_var.get()
        self.settings['blur'] = self.blur_var.get()
        self.settings['hue'] = self.hue_var.get()
        self.settings['exposure'] = self.exposure_var.get()
        if self.threshold_enabled_var.get():
            self.settings['threshold'] = self.threshold_var.get()
        else:
            self.settings['threshold'] = None
        self.settings['threshold_enabled'] = self.threshold_enabled_var.get()

        # Ensure the settings are properly stored in the game_text_reader's processing_settings
        area_name = self.area_name
        self.game_text_reader.processing_settings[area_name] = self.settings.copy()

        # Save Auto Read settings to file immediately
        if area_name.startswith("Auto Read"):
            # First save the processing settings
            self.game_text_reader.processing_settings[area_name] = self.settings.copy()
            
            # Call the existing save_auto_read_settings function to save all settings
            # This will include hotkey, checkboxes, and other settings
            if hasattr(self.game_text_reader, 'save_auto_read_settings'):
                # Get a reference to the save_auto_read_settings function
                # Find the first Auto Read area (they all have the same save function)
                save_func = None
                for area in self.game_text_reader.areas:
                    area_frame2, _, _, area_name_var2, _, _, _, _ = area
                    if area_name_var2.get().startswith("Auto Read"):
                        # This is a bit of a hack - we're accessing the nested function through the frame's children
                        for child in area[0].winfo_children():
                            if hasattr(child, '_name') and child._name == 'save_auto_read_settings':
                                save_func = child
                                break
                        if save_func:
                            break
                
                if save_func:
                    # Call the save function
                    save_func()
                    
                    # Show feedback in status label if available
                    if hasattr(self.game_text_reader, 'status_label'):
                        self.game_text_reader.status_label.config(text="Auto Read settings saved", fg="black")
                        if hasattr(self.game_text_reader, '_feedback_timer') and self.game_text_reader._feedback_timer:
                            self.game_text_reader.root.after_cancel(self.game_text_reader._feedback_timer)
                        self.game_text_reader._feedback_timer = self.game_text_reader.root.after(2000, 
                            lambda: self.game_text_reader.status_label.config(text=""))

        # Find and enable the preprocess checkbox for this area
        for area_frame, _, _, area_name_var, preprocess_var, _, _, _, _ in self.game_text_reader.areas:
            if area_name_var.get() == area_name:
                preprocess_var.set(True)  # Enable the checkbox
                break

        # Check if this is the Auto Read area or if there's a layout file
        is_auto_read = self.area_name.startswith("Auto Read")
        has_layout_file = bool(self.game_text_reader.layout_file.get())
        
        if not has_layout_file and not is_auto_read:
            # Create custom dialog for non-Auto Read areas without a layout file
            dialog = tk.Toplevel(self.window)
            dialog.title("No Save File")
            dialog.geometry("400x150")
            
            # Set the window icon
            try:
                icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
                if os.path.exists(icon_path):
                    dialog.iconbitmap(icon_path)
            except Exception as e:
                print(f"Error setting dialog icon: {e}")
            
            dialog.transient(self.window)  # Make dialog modal
            dialog.grab_set()  # Make dialog modal
            
            # Center the dialog on the screen
            dialog.geometry("+%d+%d" % (
                self.window.winfo_rootx() + self.window.winfo_width()/2 - 200,
                self.window.winfo_rooty() + self.window.winfo_height()/2 - 75))
            
            # Add message
            message = tk.Label(dialog, 
                text="No save file exists. You need to save the layout\nto preserve these settings.\n\nCreate save file now?",
                pady=20)
            message.pack()
            
            # Add buttons frame
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=10)
            
            # Create Yes button
            def on_yes():
                dialog.destroy()
                self.game_text_reader.save_layout()
                
            # Create No button
            def on_no():
                dialog.destroy()
                return
                
            yes_button = tk.Button(button_frame, text="Yes", command=on_yes, width=10)
            yes_button.pack(side='left', padx=10)
            
            no_button = tk.Button(button_frame, text="No", command=on_no, width=10)
            no_button.pack(side='left', padx=10)
            
            # Center the dialog on the screen
            dialog.update_idletasks()
            width = dialog.winfo_width()
            height = dialog.winfo_height()
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
            dialog.geometry(f'{width}x{height}+{x}+{y}')
            
            # Make the dialog modal
            dialog.transient(self.window)
            dialog.grab_set()
            
            # Wait for dialog to close
            self.window.wait_window(dialog)
            
            # If we get here, the user closed the dialog without clicking a button
            return

        # Store a reference to game_text_reader before destroying window
        game_text_reader = self.game_text_reader

        # --- AUTO SAVE for Auto Read area ---
        if area_name.startswith("Auto Read"):
            # Processing settings are already saved to self.game_text_reader.processing_settings[area_name] above
            # Now save to layout file if one exists, otherwise just update the in-memory settings
            current_layout_file = game_text_reader.layout_file.get()
            if current_layout_file and os.path.exists(current_layout_file):
                # Auto-save the layout file to include the updated processing settings
                try:
                    game_text_reader.save_layout_auto()
                    # Show status message if available
                    if hasattr(game_text_reader, 'status_label'):
                        game_text_reader.status_label.config(text="Auto Read area settings saved (auto)", fg="black")
                        if hasattr(game_text_reader, '_feedback_timer') and game_text_reader._feedback_timer:
                            game_text_reader.root.after_cancel(game_text_reader._feedback_timer)
                        game_text_reader._feedback_timer = game_text_reader.root.after(2000, lambda: game_text_reader.status_label.config(text=""))
                except Exception as e:
                    print(f"Warning: Could not auto-save layout file: {e}")
                    # Still show success message since settings are saved in memory
                    if hasattr(game_text_reader, 'status_label'):
                        game_text_reader.status_label.config(text="Settings saved (layout file not available)", fg="orange")
                        if hasattr(game_text_reader, '_feedback_timer') and game_text_reader._feedback_timer:
                            game_text_reader.root.after_cancel(game_text_reader._feedback_timer)
                        game_text_reader._feedback_timer = game_text_reader.root.after(2000, lambda: game_text_reader.status_label.config(text=""))
            else:
                # No layout file loaded, just update in-memory settings
                # Show status message
                if hasattr(game_text_reader, 'status_label'):
                    game_text_reader.status_label.config(text="Settings saved (save layout to persist)", fg="orange")
                    if hasattr(game_text_reader, '_feedback_timer') and game_text_reader._feedback_timer:
                        game_text_reader.root.after_cancel(game_text_reader._feedback_timer)
                    game_text_reader._feedback_timer = game_text_reader.root.after(2000, lambda: game_text_reader.status_label.config(text=""))
            # Unregister this window before destroying window (since on_close won't be called)
            self.game_text_reader.unregister_hotkey_disabling_window("Image Processing")
            # Destroy window (if not already destroyed)
            self.window.destroy()
            return

        # For all other areas, continue with manual/dialog save logic
        # Unregister this window before destroying window (since on_close won't be called)
        self.game_text_reader.unregister_hotkey_disabling_window("Image Processing")
        # Destroy window
        self.window.destroy()
        # Now that everything is properly synchronized, save the layout
        game_text_reader.save_layout()


    def update_preview(self, *args):
        """Update the preview with current settings and scale"""
        # Apply current processing settings
        self.processed_image = preprocess_image(
            self.image,
            brightness=self.brightness_var.get(),
            contrast=self.contrast_var.get(),
            saturation=self.saturation_var.get(),
            sharpness=self.sharpness_var.get(),
            blur=self.blur_var.get(),
            threshold=self.threshold_var.get() if self.threshold_enabled_var.get() else None,
            hue=self.hue_var.get(),
            exposure=self.exposure_var.get()
        )

        # Scale the image according to the selected percentage
        scale_factor = int(self.scale_var.get()) / 100
        if scale_factor != 1:
            new_width = int(self.processed_image.width * scale_factor)
            new_height = int(self.processed_image.height * scale_factor)
            display_image = self.processed_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        else:
            display_image = self.processed_image

        # Update the canvas size
        self.canvas.config(width=display_image.width, height=display_image.height)
        
        # Update the displayed image
        self.photo_image = ImageTk.PhotoImage(display_image)
        self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)
    
    def on_close(self):
        """Re-enable hotkeys when the window is closed and cleanup PhotoImage"""
        # Cleanup PhotoImage to prevent memory leaks
        try:
            if hasattr(self, 'photo_image') and self.photo_image is not None:
                del self.photo_image
                self.photo_image = None
        except Exception:
            pass
        # Cleanup processed image
        try:
            if hasattr(self, 'processed_image') and self.processed_image is not None:
                if hasattr(self.processed_image, 'close'):
                    self.processed_image.close()
        except Exception:
            pass
        self.game_text_reader.unregister_hotkey_disabling_window("Image Processing")
        self.window.destroy()

