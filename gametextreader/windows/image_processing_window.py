"""
Image processing settings window for adjusting OCR preprocessing parameters
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageEnhance, ImageFilter, ImageTk
from tkinter import font as tkfont
import time

from ..image_processing import preprocess_image, apply_color_mask


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
            # Don't need to unregister since we didn't register
            self.window.destroy()
            return

        # Use original image for preview to match debug console behavior
        if area_name in self.game_text_reader.original_images:
            self.original_image = self.game_text_reader.original_images[area_name]
            # Check if image is still valid, if not use latest_images
            try:
                self.original_image.load()  # Test if image is still valid
                self.image = self.original_image.copy()  # Start with original, not processed
                self.using_fallback_original = False
                print(f"Using original image for preview: {area_name}")
            except (ValueError, AttributeError):
                print(f"Original image closed, using latest_images for: {area_name}")
                self.image = latest_images[area_name].copy()
                self.original_image = self.image.copy()
                self.using_fallback_original = True
        else:
            # Try to capture a fresh original image for this area
            try:
                # Find the area configuration to get its coordinates
                area_coords = None
                for area in self.game_text_reader.areas:
                    area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var = area[:8]
                    if area_name_var.get() == area_name:
                        # Get coordinates from the area's stored data
                        if hasattr(area_frame, 'area_coords'):
                            area_coords = area_frame.area_coords
                        break
                
                if area_coords:
                    # Capture fresh image for original
                    from ..screen_capture import capture_screen_area
                    x1, y1, x2, y2 = area_coords
                    fresh_capture = capture_screen_area(x1, y1, x2, y2)
                    self.original_image = fresh_capture
                    self.using_fallback_original = False
                    print(f"Captured fresh original image for {area_name}")
                else:
                    raise Exception("Area coordinates not found")
                    
            except Exception as e:
                print(f"Could not capture fresh original image for {area_name}: {e}")
                # Fallback: use current processed image (but mark as fallback)
                self.original_image = self.image.copy()
                self.using_fallback_original = True
                print(f"Warning: Using processed image as fallback for {area_name} - no true original available")
        
        # Create a copy for processed image (with error handling for closed images)
        try:
            self.processed_image = self.image.copy()  # Start with the processed image that was used for OCR
        except Exception as e:
            print(f"Warning: Could not copy processed image (may be closed): {e}")
            # Try to reload from original if available
            if hasattr(self, 'original_image') and self.original_image:
                try:
                    self.processed_image = self.original_image.copy()
                    self.image = self.original_image.copy()
                    print("Reloaded image from original")
                except Exception as e2:
                    print(f"Error: Could not reload from original: {e2}")
                    # Create a blank fallback image
                    from PIL import Image
                    self.processed_image = Image.new('RGB', (100, 100), color='white')
                    self.image = Image.new('RGB', (100, 100), color='white')
                    print("Created blank fallback image")
            else:
                # Create a blank fallback image
                from PIL import Image
                self.processed_image = Image.new('RGB', (100, 100), color='white')
                self.image = Image.new('RGB', (100, 100), color='white')
                print("Created blank fallback image")
        
        # Toggle state for reset button
        self.showing_original = False
        
        # Flag to track if window is closing
        self.is_closing = False
        
        # Create custom fonts for bold/normal text
        self.bold_font = tkfont.Font(family="Helvetica", size=9, weight="bold")
        self.normal_font = tkfont.Font(family="Helvetica", size=9, weight="normal")

        # Debouncing attributes for UI updates
        self._update_pending = False
        self._last_update_time = 0
        self._debounce_delay = 0.05  # 50ms debounce delay

        # Don't disable hotkeys for image processing window
        # self.game_text_reader.register_hotkey_disabling_window("Image Processing", self.window)
        
        # Create a scrollable frame for all content
        self.create_scrollable_frame()
        
        # Create a canvas to display the image in the fixed image section
        self.image_frame = ttk.Frame(self.image_section)
        self.image_frame.pack(padx=10, pady=5)
        self.canvas = tk.Canvas(self.image_frame, width=self.image.width, height=self.image.height)
        self.canvas.pack()

        # Display the image on the canvas
        self.photo_image = ImageTk.PhotoImage(self.image)
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        # Add a label under the image with larger text - centered
        info_text = f"Showing previous image captured in area: {area_name}\n\nProcessing applies to unprocessed images; results may differ if the preview is already processed."
        if self.using_fallback_original:
            info_text += "\n⚠️ Warning: No true original image available - using processed image as fallback"
        info_label = ttk.Label(self.image_frame, text=info_text, font=("Helvetica", 12), justify='center')
        info_label.pack(pady=(10, 0), fill='x')

        # Create a frame for bottom controls in the fixed image section
        control_frame = ttk.Frame(self.image_section)
        control_frame.pack(pady=10)

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
        self.reset_button_frame = ttk.Frame(control_frame, relief="raised", borderwidth=2)
        self.reset_button_frame.pack(side='left', padx=10)
        self.reset_button_frame.bind("<Button-1>", self.toggle_original_processed)
        self.reset_button_frame.bind("<Enter>", lambda e: self.reset_button_frame.config(relief="sunken"))
        self.reset_button_frame.bind("<Leave>", lambda e: self.reset_button_frame.config(relief="raised"))
        
        # Create two labels for mixed formatting
        self.processed_label = tk.Label(self.reset_button_frame, text="Processed", font=self.bold_font, padx=5, pady=5)
        self.processed_label.pack(side='left')
        
        self.separator_label = tk.Label(self.reset_button_frame, text=" / ", font=self.normal_font, padx=0, pady=5)
        self.separator_label.pack(side='left')
        
        self.unprocessed_label = tk.Label(self.reset_button_frame, text="Unprocessed", font=self.normal_font, padx=5, pady=5)
        self.unprocessed_label.pack(side='left')
        
        # Bind click events to all parts
        for label in [self.processed_label, self.separator_label, self.unprocessed_label]:
            label.bind("<Button-1>", self.toggle_original_processed)
        
        ttk.Button(control_frame, text="Reset Sliders", command=self.reset_sliders_only).pack(side='left', padx=10)
        ttk.Button(control_frame, text="Reset Edits", command=self.reset_edits).pack(side='left', padx=10)
        ttk.Button(control_frame, text="Apply img. processing", command=self.save_settings).pack(side='left', padx=10)

        # Add sliders for image processing
        self.brightness_var = tk.DoubleVar(value=settings.get('brightness', 1.0))
        self.contrast_var = tk.DoubleVar(value=settings.get('contrast', 1.0))
        self.saturation_var = tk.DoubleVar(value=settings.get('saturation', 1.0))
        self.sharpness_var = tk.DoubleVar(value=settings.get('sharpness', 1.0))
        self.threshold_var = tk.IntVar(value=settings.get('threshold', 128))
        self.hue_var = tk.DoubleVar(value=settings.get('hue', 0.0))
        self.exposure_var = tk.DoubleVar(value=settings.get('exposure', 1.0))
        self.threshold_enabled_var = tk.BooleanVar(value=settings.get('threshold_enabled', False))
        
        # Color mask variables
        self.color_mask_enabled_var = tk.BooleanVar(value=settings.get('color_mask_enabled', False))
        self.color_mask_color_var = tk.StringVar(value=settings.get('color_mask_color', '#FF0000'))
        self.color_mask_tolerance_var = tk.IntVar(value=settings.get('color_mask_tolerance', 15))
        self.color_mask_background_var = tk.StringVar(value=settings.get('color_mask_background', 'black'))
        self.color_mask_position_var = tk.StringVar(value=settings.get('color_mask_position', 'after'))
        
        # Text-specific color mask options
        self.color_mask_text_mode_var = tk.BooleanVar(value=settings.get('color_mask_text_mode', False))
        self.color_mask_preserve_edges_var = tk.BooleanVar(value=settings.get('color_mask_preserve_edges', False))
        self.color_mask_enhance_contrast_var = tk.BooleanVar(value=settings.get('color_mask_enhance_contrast', False))

        self.create_slider("Brightness", self.brightness_var, 0.1, 2.0, 1.0, 0)
        self.create_slider("Contrast", self.contrast_var, 0.1, 2.0, 1.0, 1)
        self.create_slider("Saturation", self.saturation_var, 0.1, 2.0, 1.0, 2)
        self.create_slider("Sharpness", self.sharpness_var, 0.1, 2.0, 1.0, 3)
        self.create_slider("Exposure", self.exposure_var, 0.1, 2.0, 1.0, 4)
        self.create_slider("Hue", self.hue_var, -1.0, 1.0, 0.0, 5)
        self.create_slider("Threshold", self.threshold_var, 0, 255, 128, 6, self.threshold_enabled_var)
        
        # Create color mask UI
        self.create_color_mask_ui()
        
        # Store initial settings to detect changes (after all variables are created)
        self.initial_settings = {
            'brightness': self.brightness_var.get(),
            'contrast': self.contrast_var.get(),
            'saturation': self.saturation_var.get(),
            'sharpness': self.sharpness_var.get(),
            'threshold': self.threshold_var.get(),
            'hue': self.hue_var.get(),
            'exposure': self.exposure_var.get(),
            'threshold_enabled': self.threshold_enabled_var.get(),
            'color_mask_enabled': self.color_mask_enabled_var.get(),
            'color_mask_color': self.color_mask_color_var.get(),
            'color_mask_tolerance': self.color_mask_tolerance_var.get(),
            'color_mask_background': self.color_mask_background_var.get(),
            'color_mask_position': self.color_mask_position_var.get(),
            'color_mask_text_mode': self.color_mask_text_mode_var.get(),
            'color_mask_preserve_edges': self.color_mask_preserve_edges_var.get(),
            'color_mask_enhance_contrast': self.color_mask_enhance_contrast_var.get()
        }

    def create_scrollable_frame(self):
        """Create a layout with fixed image area and scrollable settings section"""
        # Main container
        main_container = ttk.Frame(self.window)
        main_container.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        
        # Fixed image section at the top
        self.image_section = ttk.Frame(main_container)
        self.image_section.pack(fill='x', pady=(0, 10))
        
        # Scrollable settings section at the bottom
        self.create_scrollable_settings(main_container)
        
        # Set initial window size
        self.window.geometry("910x600")
        
    def create_scrollable_settings(self, parent):
        """Create a scrollable frame for settings sliders and color mask"""
        # Create frame for scrollable settings
        settings_container = ttk.LabelFrame(parent, text="Image Processing Settings")
        settings_container.pack(fill='both', expand=True)
        
        # Create canvas with scrollbars for settings
        self.settings_canvas = tk.Canvas(settings_container, highlightthickness=0)
        self.settings_canvas.pack(side='left', fill='both', expand=True)
        
        # Create vertical scrollbar for settings
        v_scrollbar = ttk.Scrollbar(settings_container, orient='vertical', command=self.settings_canvas.yview)
        v_scrollbar.pack(side='right', fill='y')
        
        # Configure canvas to use scrollbar
        self.settings_canvas.configure(yscrollcommand=v_scrollbar.set)
        
        # Create scrollable frame inside canvas
        self.scrollable_frame = ttk.Frame(self.settings_canvas)
        self.settings_canvas_window = self.settings_canvas.create_window((0, 0), window=self.scrollable_frame, anchor='nw')
        
        # Bind events to update scrollregion when frame changes size
        self.scrollable_frame.bind('<Configure>', self.on_settings_frame_configure)
        self.settings_canvas.bind('<Configure>', self.on_settings_canvas_configure)
        
        # Enable mouse wheel scrolling for settings - bind to entire window and scrollable frame
        self.window.bind('<MouseWheel>', self.on_settings_mousewheel)
        self.window.bind('<Button-4>', self.on_settings_mousewheel)  # Linux scroll up
        self.window.bind('<Button-5>', self.on_settings_mousewheel)  # Linux scroll down
        self.scrollable_frame.bind('<MouseWheel>', self.on_settings_mousewheel)
        self.scrollable_frame.bind('<Button-4>', self.on_settings_mousewheel)  # Linux scroll up
        self.scrollable_frame.bind('<Button-5>', self.on_settings_mousewheel)  # Linux scroll down
        self.settings_canvas.bind('<MouseWheel>', self.on_settings_mousewheel)
        self.settings_canvas.bind('<Button-4>', self.on_settings_mousewheel)  # Linux scroll up
        self.settings_canvas.bind('<Button-5>', self.on_settings_mousewheel)  # Linux scroll down
        
    def on_settings_frame_configure(self, event):
        """Update scrollregion when the settings frame changes size"""
        self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox('all'))
        
    def on_settings_canvas_configure(self, event):
        """Update frame width when settings canvas changes size"""
        canvas_width = event.width
        self.settings_canvas.itemconfig(self.settings_canvas_window, width=canvas_width)
        
    def on_settings_mousewheel(self, event):
        """Handle mouse wheel scrolling in settings area"""
        # Get the current scroll position
        if event.delta:  # Windows
            delta = -1 * (event.delta // 120)
        elif event.num == 4:  # Linux scroll up
            delta = -1
        elif event.num == 5:  # Linux scroll down
            delta = 1
        else:
            delta = 0
            
        # Scroll vertically
        self.settings_canvas.yview_scroll(delta, 'units')
        

    def has_settings_changed(self):
        """Check if any settings have changed from initial values"""
        current_settings = {
            'brightness': self.brightness_var.get(),
            'contrast': self.contrast_var.get(),
            'saturation': self.saturation_var.get(),
            'sharpness': self.sharpness_var.get(),
            'threshold': self.threshold_var.get(),
            'hue': self.hue_var.get(),
            'exposure': self.exposure_var.get(),
            'threshold_enabled': self.threshold_enabled_var.get(),
            'color_mask_enabled': self.color_mask_enabled_var.get(),
            'color_mask_color': self.color_mask_color_var.get(),
            'color_mask_tolerance': self.color_mask_tolerance_var.get(),
            'color_mask_background': self.color_mask_background_var.get(),
            'color_mask_position': self.color_mask_position_var.get(),
            'color_mask_text_mode': self.color_mask_text_mode_var.get(),
            'color_mask_preserve_edges': self.color_mask_preserve_edges_var.get(),
            'color_mask_enhance_contrast': self.color_mask_enhance_contrast_var.get()
        }
        return current_settings != self.initial_settings

    def create_slider(self, label, variable, from_, to, initial, position, enabled_var=None):
        # Calculate grid position (7 columns per row so all fit in one row)
        row = position // 7
        col = position % 7
        print(f"Creating slider '{label}' at position {position} -> row {row}, col {col}")
        
        # Create slider directly in scrollable frame without extra frame
        label_frame = ttk.Frame(self.scrollable_frame, relief='solid', borderwidth=1)
        label_frame.grid(row=row, column=col, padx=2, pady=2, sticky='ew')  # Use grid layout
        
        # Don't set fixed width - let grid control it
        
        # Add title label inside the frame
        title_label = ttk.Label(label_frame, text=label, font=('TkDefaultFont', 7, 'bold'))  # Even smaller font
        title_label.pack(pady=(1, 1))  # Minimal padding

        entry_var = tk.StringVar(value=f'{initial:.2f}')
        # Add trace to variable to update entry field
        variable.trace_add('write', lambda *args: entry_var.set(f'{variable.get():.2f}'))

        slider = ttk.Scale(label_frame, from_=from_, to=to, orient='horizontal', variable=variable, command=self.debounced_update_image)
        slider.set(initial)
        slider.pack(fill='x', padx=2)  # Minimal padding

        # Create a frame for entry and reset button side by side
        entry_reset_frame = ttk.Frame(label_frame)
        entry_reset_frame.pack(fill='x', pady=(0, 1))
        
        # Create entry with context menu - much narrower
        entry = ttk.Entry(entry_reset_frame, textvariable=entry_var, width=5)  # Much smaller width
        entry.pack(side='left', padx=(2, 1))
        
        # Add reset button to the right of entry - full text
        reset_button = ttk.Button(entry_reset_frame, text="Reset", width=6, command=lambda: self.reset_slider(slider, entry, initial, variable))
        reset_button.pack(side='right', padx=(1, 2))
        
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
        
        # Enable/disable slider and entry based on enabled_var
        if enabled_var:
            def update_state(*args):
                state = 'normal' if enabled_var.get() else 'disabled'
                slider.config(state=state)
                entry.config(state=state)
                reset_button.config(state=state)
            
            enabled_var.trace_add('write', update_state)
            update_state()  # Set initial state

        # Create checkbox for threshold slider
        if label == "Threshold":
            checkbox_frame = ttk.Frame(label_frame)
            checkbox_frame.pack(anchor='w', padx=2)
        
            checkbox = ttk.Checkbutton(checkbox_frame, variable=enabled_var, command=self.update_image)
            checkbox.pack(side=tk.LEFT)

            ttk.Label(checkbox_frame, text="Enabled", font=('TkDefaultFont', 6)).pack(side=tk.LEFT, padx=(1, 0))  # Even smaller font

        setattr(self, f"{label.lower()}_slider", label_frame)
        label_frame.slider, label_frame.entry = slider, entry
        
        # Configure grid columns for narrower spacing (now 7 columns)
        for i in range(7):
            self.scrollable_frame.columnconfigure(i, weight=1, minsize=80, uniform="slider")
        
    def create_color_mask_ui(self):
        """Create color mask UI elements"""
        # Color mask frame - place it in the grid after all sliders
        color_mask_frame = ttk.Frame(self.scrollable_frame)
        # Place it in row 1 (after the 1 row of sliders), spanning all 7 columns
        color_mask_frame.grid(row=1, column=0, columnspan=7, padx=5, pady=10, sticky='ew')
        
        # Label frame for color mask
        color_mask_label_frame = ttk.LabelFrame(color_mask_frame, text="Color Mask")
        color_mask_label_frame.pack(fill='both', expand=True)
        
        # Enable checkbox
        checkbox_frame = ttk.Frame(color_mask_label_frame)
        checkbox_frame.pack(anchor='w', padx=5, pady=2)
        
        enable_checkbox = ttk.Checkbutton(
            checkbox_frame, 
            variable=self.color_mask_enabled_var, 
            command=self.update_image,
            text="Enable Color Mask"
        )
        enable_checkbox.pack(side=tk.LEFT)
        
        # Create note labels with proper layout
        note_frame = ttk.Frame(checkbox_frame)
        note_frame.pack(side=tk.LEFT, padx=(10, 0))
        
        # Red "Note:" label
        note_label = ttk.Label(
            note_frame,
            text="Note:",
            font=("Helvetica", 8, "bold"),
            foreground="red"
        )
        note_label.pack(side=tk.LEFT)
        
        # Black text label
        help_label = ttk.Label(
            note_frame,
            text="Adjusting the tolerance will change the spoken text output even if no visual changes\nappear in the preview. The image-to-text processing is highly sensitive to this setting.\nLook in the debug window for the image that is being processed.",
            font=("Helvetica", 8),
            foreground="black"
        )
        help_label.pack(side=tk.LEFT)
        
        # Color picker frame
        color_picker_frame = ttk.Frame(color_mask_label_frame)
        color_picker_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(color_picker_frame, text="Edit Mask Color:").pack(side=tk.LEFT, padx=(0, 5))
        self.color_button = tk.Button(
            color_picker_frame,
            text="    ",  # Space to show color
            bg=self.color_mask_color_var.get(),
            width=3,
            height=1,
            relief='sunken',
            bd=1,
            command=self.pick_color_mask_color
        )
        self.color_button.pack(side=tk.LEFT, padx=2)
        
        # Pick from image button
        pick_from_image_button = tk.Button(
            color_picker_frame,
            text="Select color from image",
            width=20,
            height=1,
            relief='raised',
            font=("Helvetica", 8),
            command=self.pick_color_from_image
        )
        pick_from_image_button.pack(side=tk.LEFT, padx=2)
        
        # Clear color button
        clear_color_button = tk.Button(
            color_picker_frame,
            text="✕",
            width=2,
            height=1,
            relief='flat',
            font=("Helvetica", 7),
            command=self.clear_color_mask
        )
        clear_color_button.pack(side=tk.LEFT, padx=1)
        
        # Tolerance slider frame
        tolerance_frame = ttk.Frame(color_mask_label_frame)
        tolerance_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(tolerance_frame, text="Tolerance:").pack(side=tk.LEFT)
        
        # Debounce timer for smoother slider updates
        self._tolerance_update_timer = None
        
        tolerance_slider = ttk.Scale(
            tolerance_frame,
            from_=1,
            to=50,
            orient='horizontal',
            variable=self.color_mask_tolerance_var,
            command=lambda x: self._debounced_tolerance_update(),
            length=150
        )
        tolerance_slider.pack(side=tk.LEFT, padx=5)
        
        tolerance_value_label = ttk.Label(tolerance_frame, text=f"{self.color_mask_tolerance_var.get()}")
        tolerance_value_label.pack(side=tk.LEFT)
        
        # Debounced update function for smoother slider experience
        def _debounced_tolerance_update():
            # Cancel any pending update
            if hasattr(self, '_tolerance_update_timer') and self._tolerance_update_timer:
                self.window.after_cancel(self._tolerance_update_timer)
            
            # Update the label immediately for responsive feedback
            tolerance_value_label.config(text=f"{self.color_mask_tolerance_var.get()}")
            
            # Debounce the actual image update
            self._tolerance_update_timer = self.window.after(50, self.update_image)
        
        self._debounced_tolerance_update = _debounced_tolerance_update
        
        # Background color frame
        background_frame = ttk.Frame(color_mask_label_frame)
        background_frame.pack(fill='x', padx=5, pady=2)
        
        ttk.Label(background_frame, text="Surrounding color:").pack(side=tk.LEFT)
        
        # Background color options
        background_options = ["black", "white"]
        background_menu = tk.OptionMenu(
            background_frame,
            self.color_mask_background_var,
            *background_options,
            command=self.update_image
        )
        background_menu.pack(side=tk.LEFT, padx=5)
        
        # Text-specific options frame
        text_options_frame = ttk.LabelFrame(color_mask_label_frame, text="Text Optimization Options")
        text_options_frame.pack(fill='x', padx=5, pady=5)
        
        # Text mode checkbox
        text_mode_frame = ttk.Frame(text_options_frame)
        text_mode_frame.pack(fill='x', padx=5, pady=2)
        
        text_mode_checkbox = ttk.Checkbutton(
            text_mode_frame,
            variable=self.color_mask_text_mode_var,
            command=self.update_image,
            text="Enable text mode (optimizes for OCR)"
        )
        text_mode_checkbox.pack(side=tk.LEFT)
        
        # Preserve edges checkbox
        preserve_edges_frame = ttk.Frame(text_options_frame)
        preserve_edges_frame.pack(fill='x', padx=5, pady=2)
        
        preserve_edges_checkbox = ttk.Checkbutton(
            preserve_edges_frame,
            variable=self.color_mask_preserve_edges_var,
            command=self.update_image,
            text="Preserve character edges (removes font artifacts)"
        )
        preserve_edges_checkbox.pack(side=tk.LEFT)
        
        # Enhance contrast checkbox
        enhance_contrast_frame = ttk.Frame(text_options_frame)
        enhance_contrast_frame.pack(fill='x', padx=5, pady=2)
        
        enhance_contrast_checkbox = ttk.Checkbutton(
            enhance_contrast_frame,
            variable=self.color_mask_enhance_contrast_var,
            command=self.update_image,
            text="Enhance text contrast (improves OCR accuracy)"
        )
        enhance_contrast_checkbox.pack(side=tk.LEFT)
        
        # Store reference
        self.color_mask_frame = color_mask_frame
        
    def pick_color_mask_color(self):
        """Open color picker dialog for color mask"""
        try:
            from tkinter import colorchooser
            
            # Get current color
            current_color = self.color_mask_color_var.get()
            
            # Open color picker
            color = colorchooser.askcolor(initialcolor=current_color, parent=self.window)
            
            if color[1]:  # User didn't cancel
                hex_color = color[1]
                self.color_mask_color_var.set(hex_color)
                self.color_button.config(bg=hex_color)
                self.update_image()
        except Exception as e:
            print(f"Error picking color: {e}")
    
    def clear_color_mask(self):
        """Clear the color mask"""
        self.color_mask_enabled_var.set(False)
        self.color_mask_color_var.set('#FF0000')
        self.color_mask_tolerance_var.set(15)
        self.color_mask_background_var.set('black')
        self.color_mask_position_var.set('after')
        self.color_mask_text_mode_var.set(False)
        self.color_mask_preserve_edges_var.set(False)
        self.color_mask_enhance_contrast_var.set(False)
        self.color_button.config(bg="#FF0000")
        self.update_image()
        
    def confirm_color_selection(self, hex_color, picker_window):
        """Fallback method to confirm color selection"""
        try:
            if hex_color and hex_color != "#FFFFFF":
                # Update main window color
                self.color_mask_color_var.set(hex_color)
                self.color_button.config(bg=hex_color)
                
                # Enable color mask if not already enabled
                if not self.color_mask_enabled_var.get():
                    self.color_mask_enabled_var.set(True)
                
                # Update the main image
                self.update_image()
                
                # Close picker window
                picker_window.destroy()
                
        except Exception as e:
            print(f"Error confirming color selection: {e}")

    def pick_color_from_image(self):
        """Open a new window to pick color from the image"""
        try:
            # Create a new window for color picking
            picker_window = tk.Toplevel(self.window)
            picker_window.title("Pick Color from Image")
            picker_window.geometry("600x500")
            picker_window.resizable(True, True)
            
            # Set the window icon
            try:
                icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
                if os.path.exists(icon_path):
                    picker_window.iconbitmap(icon_path)
            except Exception as e:
                print(f"Error setting color picker window icon: {e}")
            
            # Center the window
            picker_window.update_idletasks()
            x = (picker_window.winfo_screenwidth() // 2) - (600 // 2)
            y = (picker_window.winfo_screenheight() // 2) - (500 // 2)
            picker_window.geometry(f"600x500+{x}+{y}")
            
            # Add instruction label
            instruction_label = tk.Label(
                picker_window, 
                text="Click on any pixel in the image to select its color",
                font=("Helvetica", 10),
                fg="blue"
            )
            instruction_label.pack(pady=5)
            
            # Create a frame for controls
            controls_frame = tk.Frame(picker_window)
            controls_frame.pack(fill='x', padx=10, pady=5)
            
            # Zoom controls
            tk.Label(controls_frame, text="Zoom:").pack(side=tk.LEFT, padx=(0, 5))
            
            zoom_var = tk.DoubleVar(value=100.0)
            zoom_options = [25, 50, 75, 100, 125, 150, 200, 300, 400]
            zoom_menu = tk.OptionMenu(
                controls_frame,
                zoom_var,
                *zoom_options,
                command=lambda value: update_zoom(float(value))
            )
            zoom_menu.pack(side=tk.LEFT, padx=5)
            
            tk.Label(controls_frame, text="%").pack(side=tk.LEFT, padx=(0, 10))
            
            # Magnified preview
            tk.Label(controls_frame, text="Magnified:").pack(side=tk.LEFT, padx=(0, 5))
            
            magnified_canvas = tk.Canvas(
                controls_frame,
                width=60,
                height=60,
                bg="white",
                relief="solid",
                borderwidth=1,
                highlightthickness=0
            )
            magnified_canvas.pack(side=tk.LEFT, padx=5)
            
            # Create a frame for the image with scrollbars
            image_frame = tk.Frame(picker_window)
            image_frame.pack(fill='both', expand=True, padx=10, pady=5)
            
            # Create scrollbars
            h_scrollbar = tk.Scrollbar(image_frame, orient=tk.HORIZONTAL)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            v_scrollbar = tk.Scrollbar(image_frame, orient=tk.VERTICAL)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Create canvas with scrollbars
            canvas = tk.Canvas(
                image_frame, 
                width=580, 
                height=400,
                xscrollcommand=h_scrollbar.set,
                yscrollcommand=v_scrollbar.set
            )
            canvas.pack(side=tk.LEFT, fill='both', expand=True)
            
            h_scrollbar.config(command=canvas.xview)
            v_scrollbar.config(command=canvas.yview)
            
            # Use the original unedited image for color picking
            if hasattr(self, 'original_image') and self.original_image:
                source_image = self.original_image
            elif hasattr(self, 'image') and self.image:
                source_image = self.image
            else:
                messagebox.showerror("Error", "No image available for color picking")
                picker_window.destroy()
                return
            
            # Function to update zoom
            def update_zoom(zoom_percent):
                nonlocal display_image, photo_image, image_on_canvas
                
                zoom_scale = zoom_percent / 100.0
                new_width = int(source_image.width * zoom_scale)
                new_height = int(source_image.height * zoom_scale)
                
                display_image = source_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                photo_image = ImageTk.PhotoImage(display_image)
                
                # Update canvas size and image
                canvas.config(scrollregion=(0, 0, new_width, new_height))
                image_on_canvas = canvas.create_image(0, 0, anchor=tk.NW, image=photo_image)
                
                # Store updated references
                canvas.photo_image = photo_image
                canvas.display_image = display_image
                canvas.scale = zoom_scale
            
            # Initial display at 100% zoom
            zoom_scale = zoom_var.get() / 100.0
            new_width = int(source_image.width * zoom_scale)
            new_height = int(source_image.height * zoom_scale)
            display_image = source_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Display the image
            photo_image = ImageTk.PhotoImage(display_image)
            image_on_canvas = canvas.create_image(0, 0, anchor=tk.NW, image=photo_image)
            
            # Set scroll region
            canvas.config(scrollregion=(0, 0, new_width, new_height))
            
            # Store reference to prevent garbage collection
            canvas.photo_image = photo_image
            canvas.source_image = source_image
            canvas.scale = zoom_scale
            
            # Color preview frame
            preview_frame = tk.Frame(picker_window)
            preview_frame.pack(fill='x', padx=10, pady=5)
            
            tk.Label(preview_frame, text="Selected Color:").pack(side=tk.LEFT, padx=(0, 5))
            
            color_preview = tk.Label(preview_frame, width=10, height=2, bg="#FFFFFF", relief="solid", borderwidth=2)
            color_preview.pack(side=tk.LEFT, padx=5)
            
            hex_label = tk.Label(preview_frame, text="#FFFFFF", font=("Courier", 10))
            hex_label.pack(side=tk.LEFT, padx=5)
            
            # Add confirm button as fallback
            confirm_button = tk.Button(
                preview_frame,
                text="Use This Color",
                command=lambda: self.confirm_color_selection(hex_label['text'], picker_window)
            )
            confirm_button.pack(side=tk.LEFT, padx=10)
            
            # Add close button
            close_button = tk.Button(
                preview_frame,
                text="Cancel",
                command=picker_window.destroy
            )
            close_button.pack(side=tk.LEFT, padx=5)
            
            # Function to handle color selection
            def pick_color(event):
                try:
                    # Get canvas coordinates
                    canvas_x = canvas.canvasx(event.x)
                    canvas_y = canvas.canvasy(event.y)
                    
                    # Convert to image coordinates
                    image_x = int(canvas_x / canvas.scale)
                    image_y = int(canvas_y / canvas.scale)
                    
                    # Check bounds
                    if (0 <= image_x < canvas.source_image.width and 
                        0 <= image_y < canvas.source_image.height):
                        
                        # Get pixel color
                        pixel_color = canvas.source_image.getpixel((image_x, image_y))
                        
                        # Convert to hex
                        if len(pixel_color) == 3:  # RGB
                            r, g, b = pixel_color
                        else:  # RGBA or other
                            r, g, b = pixel_color[:3]
                        
                        hex_color = f"#{r:02x}{g:02x}{b:02x}".upper()
                        
                        # Update preview
                        color_preview.config(bg=hex_color)
                        hex_label.config(text=hex_color)
                        
                        # Update main window color
                        self.color_mask_color_var.set(hex_color)
                        self.color_button.config(bg=hex_color)
                        
                        # Enable color mask if not already enabled
                        if not self.color_mask_enabled_var.get():
                            self.color_mask_enabled_var.set(True)
                        
                        # Update the main image
                        self.update_image()
                        
                        # Close picker window
                        picker_window.destroy()
                        
                except Exception as e:
                    print(f"Error picking color: {e}")
            
            # Function to handle mouse motion for live preview
            def on_mouse_motion(event):
                try:
                    # Get canvas coordinates
                    canvas_x = canvas.canvasx(event.x)
                    canvas_y = canvas.canvasy(event.y)
                    
                    # Convert to image coordinates
                    image_x = int(canvas_x / canvas.scale)
                    image_y = int(canvas_y / canvas.scale)
                    
                    # Check bounds
                    if (0 <= image_x < canvas.source_image.width and 
                        0 <= image_y < canvas.source_image.height):
                        
                        # Get pixel color for preview
                        pixel_color = canvas.source_image.getpixel((image_x, image_y))
                        
                        # Convert to hex
                        if len(pixel_color) == 3:  # RGB
                            r, g, b = pixel_color
                        else:  # RGBA or other
                            r, g, b = pixel_color[:3]
                        
                        hex_color = f"#{r:02x}{g:02x}{b:02x}".upper()
                        
                        # Update preview
                        color_preview.config(bg=hex_color)
                        hex_label.config(text=hex_color)
                        
                        # Update magnified preview
                        update_magnified_preview(image_x, image_y)
                    else:
                        # Outside image bounds - clear magnified preview
                        magnified_canvas.delete("all")
                        
                except Exception as e:
                    # On error, clear the canvas
                    magnified_canvas.delete("all")
            
            # Function to update magnified preview
            def update_magnified_preview(center_x, center_y):
                try:
                    # Clear previous magnified view
                    magnified_canvas.delete("all")
                    
                    # Magnification settings
                    magnification = 10  # 10x zoom
                    preview_size = 60  # 60x60 pixels
                    source_size = preview_size // magnification  # 6x6 pixel area from source
                    
                    # Calculate source area (centered on cursor)
                    half_source = source_size // 2
                    src_x1 = max(0, center_x - half_source)
                    src_y1 = max(0, center_y - half_source)
                    src_x2 = min(canvas.source_image.width, center_x + half_source + 1)
                    src_y2 = min(canvas.source_image.height, center_y + half_source + 1)
                    
                    # Extract the source area
                    source_area = canvas.source_image.crop((src_x1, src_y1, src_x2, src_y2))
                    
                    # Resize for magnified view
                    magnified_area = source_area.resize((preview_size, preview_size), Image.Resampling.NEAREST)
                    
                    # Convert to PhotoImage and display
                    magnified_photo = ImageTk.PhotoImage(magnified_area)
                    magnified_canvas.create_image(0, 0, anchor=tk.NW, image=magnified_photo)
                    
                    # Draw crosshair in center
                    center = preview_size // 2
                    crosshair_color = "red"
                    magnified_canvas.create_line(center-5, center, center+5, center, fill=crosshair_color, width=1)
                    magnified_canvas.create_line(center, center-5, center, center+5, fill=crosshair_color, width=1)
                    
                    # Store reference to prevent garbage collection
                    magnified_canvas.magnified_photo = magnified_photo
                    
                except Exception as e:
                    # On error, clear the canvas
                    magnified_canvas.delete("all")
            
            # Bind events
            canvas.bind("<Button-1>", pick_color)
            canvas.bind("<Motion>", on_mouse_motion)
            
            # Center the window on screen
            picker_window.update_idletasks()
            x = (picker_window.winfo_screenwidth() // 2) - (picker_window.winfo_width() // 2)
            y = (picker_window.winfo_screenheight() // 2) - (picker_window.winfo_height() // 2)
            picker_window.geometry(f"+{x}+{y}")
            
            # Make window modal
            picker_window.transient(self.window)
            picker_window.grab_set()
            picker_window.focus_set()
            
        except Exception as e:
            print(f"Error opening color picker window: {e}")
            if 'picker_window' in locals():
                picker_window.destroy()
        
    
    def reset_slider(self, slider, entry, initial, variable):
        slider.set(initial)
        variable.set(initial)
        entry.delete(0, tk.END)
        entry.insert(0, str(round(float(initial), 2)))
        self.update_image()
        

    def toggle_original_processed(self, event=None):
        """Toggle between original unprocessed image and processed image"""
        if not self.showing_original:
            # Show original unprocessed image
            self.showing_original = True
            
            # Update fonts to show Unprocessed as bold
            self.processed_label.config(font=self.normal_font)
            self.unprocessed_label.config(font=self.bold_font)
            
            # Show warning if using fallback original
            if self.using_fallback_original:
                self.unprocessed_label.config(fg="orange")  # Orange warning color
            else:
                self.unprocessed_label.config(fg="black")  # Normal color
            
            # Display original image
            scale_factor = int(self.scale_var.get()) / 100
            if scale_factor != 1:
                new_width = int(self.original_image.width * scale_factor)
                new_height = int(self.original_image.height * scale_factor)
                display_image = self.original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            else:
                display_image = self.original_image
            
            # Update canvas size if needed
            self.canvas.config(width=display_image.width, height=display_image.height)
            
            # Update the displayed image
            if hasattr(self, 'photo_image') and self.photo_image:
                del self.photo_image
            self.photo_image = ImageTk.PhotoImage(display_image)
            self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)
        else:
            # Show processed image with current settings
            self.showing_original = False
            
            # Update fonts to show Processed as bold
            self.processed_label.config(font=self.bold_font)
            self.processed_label.config(fg="black")  # Normal color
            self.unprocessed_label.config(font=self.normal_font)
            self.unprocessed_label.config(fg="black")  # Normal color
            
            self.update_preview()

    def reset_sliders_only(self):
        """Reset all sliders to default values without changing the view"""
        self.brightness_var.set(1.0)
        self.contrast_var.set(1.0)
        self.saturation_var.set(1.0)
        self.sharpness_var.set(1.0)
        self.threshold_var.set(128)
        self.hue_var.set(0.0)
        self.exposure_var.set(1.0)
        self.threshold_enabled_var.set(False)
        
        # Reset color mask settings
        self.color_mask_enabled_var.set(False)
        self.color_mask_color_var.set('#FF0000')
        self.color_mask_tolerance_var.set(15)
        self.color_mask_background_var.set('black')
        self.color_mask_position_var.set('after')
        self.color_button.config(bg="#FF0000")
        
        # Update image if showing processed view
        if not self.showing_original:
            self.update_image()

    def reset_edits(self):
        """Copy the unedited image to replace the edited one, making both identical"""
        # Show confirmation dialog
        from tkinter import messagebox
        result = messagebox.askyesno(
            "Reset Edits",
            "This will replace the edited image with the unedited one and reset all settings.\n\nDo you want to continue?",
            icon='warning'
        )
        
        if not result:
            return  # User clicked Cancel
        
        # Copy the original (unedited) image to replace the base image
        if hasattr(self, 'original_image') and self.original_image:
            # Close current base image to free memory
            if hasattr(self, 'image') and self.image:
                self.image.close()
                
            # Replace the base image with the original (unedited) image
            self.image = self.original_image.copy()
        
        # Reset all sliders to default values
        self.reset_sliders_only()
        
        # Force showing processed view (which now shows the unedited image)
        self.showing_original = False
        self.processed_label.config(font=self.bold_font)
        self.unprocessed_label.config(font=self.normal_font)
        
        # Update to show the new state
        self.update_image()
        
        print("Reset edits - unedited image copied to replace edited image")

    def reset_all(self):
        """Reset all sliders to default values and show processed image"""
        self.reset_sliders_only()
        
        # Reset to showing processed image
        self.showing_original = False
        self.processed_label.config(font=self.bold_font)
        self.unprocessed_label.config(font=self.normal_font)
        self.update_image()


    def debounced_update_image(self, _=None):
        """Debounced version of update_image to prevent excessive updates during slider dragging"""
        current_time = time.time()
        
        # If we already have a pending update, cancel it
        if self._update_pending:
            self._update_pending = False
        
        # Check if enough time has passed since last update
        if current_time - self._last_update_time >= self._debounce_delay:
            # Update immediately if enough time has passed
            self._last_update_time = current_time
            self.update_image()
        else:
            # Schedule a delayed update
            self._update_pending = True
            delay_ms = int((self._debounce_delay - (current_time - self._last_update_time)) * 1000)
            self.window.after(delay_ms, self._check_and_update)

    def _check_and_update(self):
        """Check if we should still update (prevents stale updates)"""
        if self._update_pending:
            self._update_pending = False
            self._last_update_time = time.time()
            self.update_image()

    def update_image(self, _=None):
        # Don't update if window is closing
        if self.is_closing:
            return
            
        # If showing original image, switch back to processed when editing
        if self.showing_original:
            self.showing_original = False
            self.processed_label.config(font=self.bold_font)
            self.unprocessed_label.config(font=self.normal_font)
            
        if self.image:
            # Clean up previous processed image if it exists
            if self.processed_image:
                self.processed_image.close()
            
            # Use the optimized preprocess_image function
            self.processed_image = preprocess_image(
                self.image.copy(),
                brightness=self.brightness_var.get(),
                contrast=self.contrast_var.get(),
                saturation=self.saturation_var.get(),
                sharpness=self.sharpness_var.get(),
                threshold=self.threshold_var.get() if self.threshold_enabled_var.get() else None,
                hue=self.hue_var.get(),
                exposure=self.exposure_var.get(),
                color_mask_enabled=self.color_mask_enabled_var.get(),
                color_mask_color=self.color_mask_color_var.get(),
                color_mask_tolerance=self.color_mask_tolerance_var.get(),
                color_mask_background=self.color_mask_background_var.get(),
                color_mask_position=self.color_mask_position_var.get(),
                color_mask_text_mode=self.color_mask_text_mode_var.get(),
                color_mask_preserve_edges=self.color_mask_preserve_edges_var.get(),
                color_mask_enhance_contrast=self.color_mask_enhance_contrast_var.get()
            )

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
        self.settings['hue'] = self.hue_var.get()
        self.settings['exposure'] = self.exposure_var.get()
        if self.threshold_enabled_var.get():
            self.settings['threshold'] = self.threshold_var.get()
        else:
            self.settings['threshold'] = None
        self.settings['threshold_enabled'] = self.threshold_enabled_var.get()
        
        # Save color mask settings
        self.settings['color_mask_enabled'] = self.color_mask_enabled_var.get()
        self.settings['color_mask_color'] = self.color_mask_color_var.get()
        self.settings['color_mask_tolerance'] = self.color_mask_tolerance_var.get()
        self.settings['color_mask_background'] = self.color_mask_background_var.get()
        self.settings['color_mask_position'] = self.color_mask_position_var.get()
        
        # Save text-specific color mask options
        self.settings['color_mask_text_mode'] = self.color_mask_text_mode_var.get()
        self.settings['color_mask_preserve_edges'] = self.color_mask_preserve_edges_var.get()
        self.settings['color_mask_enhance_contrast'] = self.color_mask_enhance_contrast_var.get()

        # Ensure the settings are properly stored in the game_text_reader's processing_settings
        area_name = self.area_name
        self.game_text_reader.processing_settings[area_name] = self.settings.copy()
        
        # IMPORTANT: Automatically enable the preprocessing checkbox for this area
        # Find the area and enable its preprocess checkbox
        for area in self.game_text_reader.areas:
            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var = area[:8]
            if area_name_var.get() == area_name:
                preprocess_var.set(True)  # Enable preprocessing
                break

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

        # Allow image processing for all areas, not just Auto Read areas
        # The processing settings will be stored in memory and can be saved when layout is created

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

        # --- AUTO SAVE for normal areas when layout is loaded ---
        current_layout_file = game_text_reader.layout_file.get()
        if current_layout_file and os.path.exists(current_layout_file):
            # Auto-save the layout file to include the updated processing settings
            try:
                game_text_reader.save_layout_auto()
                # Show status message if available
                if hasattr(game_text_reader, 'status_label'):
                    game_text_reader.status_label.config(text=f"{area_name} settings saved (auto)", fg="black")
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
        # Don't update if window is closing
        if self.is_closing:
            return
            
        # If showing original, just update the scale
        if self.showing_original:
            scale_factor = int(self.scale_var.get()) / 100
            if scale_factor != 1:
                new_width = int(self.original_image.width * scale_factor)
                new_height = int(self.original_image.height * scale_factor)
                display_image = self.original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            else:
                display_image = self.original_image
            
            # Update the canvas size
            self.canvas.config(width=display_image.width, height=display_image.height)
            
            # Update the displayed image
            if hasattr(self, 'photo_image') and self.photo_image:
                del self.photo_image
            self.photo_image = ImageTk.PhotoImage(display_image)
            self.canvas.itemconfig(self.image_on_canvas, image=self.photo_image)
            return
        
        # Apply current processing settings
        self.processed_image = preprocess_image(
            self.image,
            brightness=self.brightness_var.get(),
            contrast=self.contrast_var.get(),
            saturation=self.saturation_var.get(),
            sharpness=self.sharpness_var.get(),
            threshold=self.threshold_var.get() if self.threshold_enabled_var.get() else None,
            hue=self.hue_var.get(),
            exposure=self.exposure_var.get(),
            color_mask_enabled=self.color_mask_enabled_var.get(),
            color_mask_color=self.color_mask_color_var.get(),
            color_mask_tolerance=self.color_mask_tolerance_var.get(),
            color_mask_background=self.color_mask_background_var.get(),
            color_mask_position=self.color_mask_position_var.get()
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
        # Check for unsaved changes
        if self.has_settings_changed():
            result = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes to the image processing settings.\n\n"
                "• Yes: Save changes and close\n"
                "• No: Discard changes and close\n"
                "• Cancel: Keep window open",
                icon='warning'
            )
            
            if result is None:  # Cancel
                return  # Don't close the window
            elif result:  # Yes - save changes
                # Set closing flag to prevent further updates
                self.is_closing = True
                # Save settings before closing
                self.save_settings()
                return  # save_settings will handle cleanup and closing
        
        # Set closing flag to prevent further updates
        self.is_closing = True
        
        # Cleanup tolerance update timer
        try:
            if hasattr(self, '_tolerance_update_timer') and self._tolerance_update_timer:
                self.window.after_cancel(self._tolerance_update_timer)
                self._tolerance_update_timer = None
        except Exception:
            pass
        
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
        # Cleanup original image
        try:
            if hasattr(self, 'original_image') and self.original_image is not None:
                if hasattr(self.original_image, 'close'):
                    self.original_image.close()
        except Exception:
            pass
        # Cleanup main image
        try:
            if hasattr(self, 'image') and self.image is not None:
                if hasattr(self.image, 'close'):
                    self.image.close()
        except Exception:
            pass
        
        # No need to unregister since we didn't register
        # self.game_text_reader.unregister_hotkey_disabling_window("Image Processing")
        self.window.destroy()

