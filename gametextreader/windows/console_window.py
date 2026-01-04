"""
Debug console window for viewing logs and processed images
"""
import datetime
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

# Maximum buffer size to prevent memory issues (10MB)
MAX_LOG_BUFFER_SIZE = 10 * 1024 * 1024


class ConsoleWindow:
    def __init__(self, root, log_buffer, layout_file_var, latest_images, latest_area_name_var):
        self.window = tk.Toplevel(root)
        self.window.title("Debug Console")
        
        # Set the window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting console window icon: {e}")
        
        self.latest_images = latest_images
        self.window.geometry("690x500")  # Initial size, will adjust based on image

        # Create a top frame for controls
        top_frame = tk.Frame(self.window)
        top_frame.pack(fill='x', padx=10, pady=5)

        # Add checkbox for image display
        self.show_image_var = tk.BooleanVar(value=True)
        self.image_checkbox = tk.Checkbutton(
            top_frame,
            text="Show last processed image",
            variable=self.show_image_var,
            command=self.update_image_display
        )
        self.image_checkbox.pack(side='left')
        
        # Add scale dropdown
        scale_frame = tk.Frame(top_frame)
        scale_frame.pack(side='left', padx=10)
        
        tk.Label(scale_frame, text="Scale:").pack(side='left')
        self.scale_var = tk.StringVar(value="100")
        scales = [str(i) for i in range(10, 101, 10)]  # Creates ["10", "20", ..., "100"]
        scale_menu = tk.OptionMenu(scale_frame, self.scale_var, *scales, command=self.update_image_display)
        scale_menu.pack(side='left')
        tk.Label(scale_frame, text="%").pack(side='left')

        # Add Save Log button
        save_log_button = tk.Button(top_frame, text="Save Log", command=self.save_log)
        save_log_button.pack(side='left', padx=(10, 0))

        # Add Clear Console button
        clear_console_button = tk.Button(top_frame, text="Clear Console", command=self.clear_console)
        clear_console_button.pack(side='left', padx=(10, 0))

        # Add Save Image button
        save_image_button = tk.Button(top_frame, text="Save Image", command=self.save_image)
        save_image_button.pack(side='left', padx=(10, 0))

        # Create a middle frame for image display
        image_frame = tk.Frame(self.window)
        image_frame.pack(fill='x', padx=10, pady=5)
        
        # Add image label to the middle frame
        self.image_label = tk.Label(image_frame)
        self.image_label.pack(fill='x')

        # Create a bottom frame for the log output
        log_frame = tk.Frame(self.window)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)

        # Add text widget for log output
        self.text_widget = tk.Text(log_frame)
        self.text_widget.pack(fill='both', expand=True)
        self.text_widget.config(state=tk.DISABLED)
        
        # Configure text tags for formatting
        self.text_widget.tag_configure('bold', font=("Helvetica", 9, "bold"))

        # Enable mouse wheel scrolling for the debug log
        def _on_mousewheel_debug(event):
            self.text_widget.yview_scroll(int(-1 * (event.delta / 120)), 'units')
            return "break"
        def _bind_mousewheel_debug(event):
            self.text_widget.bind_all('<MouseWheel>', _on_mousewheel_debug)
        def _unbind_mousewheel_debug(event):
            self.text_widget.unbind_all('<MouseWheel>')
        self.text_widget.bind('<Enter>', _bind_mousewheel_debug)
        self.text_widget.bind('<Leave>', _unbind_mousewheel_debug)

        # Add right-click context menu
        self.context_menu = tk.Menu(self.text_widget, tearoff=0)
        self.context_menu.add_command(label="Copy", command=self.copy_selection)
        self.context_menu.add_command(label="Select All", command=self.select_all)
        self.text_widget.bind("<Button-3>", self.show_context_menu)

        self.log_buffer = log_buffer
        self.layout_file_var = layout_file_var
        self.latest_area_name_var = latest_area_name_var
        self.photo = None  # Keep a reference to prevent garbage collection

        # Add line limit constant
        self.MAX_LINES = 250
        
        # Set up cleanup on window close
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)
        
        self.update_console()
    
    def on_close(self):
        """Cleanup PhotoImage on window close to prevent memory leaks"""
        try:
            if hasattr(self, 'photo') and self.photo is not None:
                del self.photo
                self.photo = None
        except Exception:
            pass
        self.window.destroy()

    def show_context_menu(self, event):
        """Show the context menu at the mouse position."""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_selection(self):
        """Copy selected text to clipboard."""
        try:
            selected_text = self.text_widget.get("sel.first", "sel.last")
            self.window.clipboard_clear()
            self.window.clipboard_append(selected_text)
        except tk.TclError:
            pass  # No text selected

    def select_all(self):
        """Select all text in the widget."""
        self.text_widget.tag_add("sel", "1.0", "end")

    def update_image_display(self, *args):
        if not self.window.winfo_exists():
            return
            
        area_name = self.latest_area_name_var.get()
        if self.show_image_var.get() and area_name in self.latest_images:
            image = self.latest_images[area_name]
            
            try:
                # Scale the image according to the selected percentage
                scale_factor = int(self.scale_var.get()) / 100
                if scale_factor != 1:
                    new_width = int(image.width * scale_factor)
                    new_height = int(image.height * scale_factor)
                    image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    
                # Calculate new window height based on scaled image height
                window_height = image.height + 300  # Add space for controls and log
                window_height = max(500, window_height)
                
                # Get current window position and width
                window_x = self.window.winfo_x()
                window_y = self.window.winfo_y()
                window_width = self.window.winfo_width()
                
                # Update window geometry
                self.window.geometry(f"{window_width}x{window_height}+{window_x}+{window_y}")
                
                # Create new photo before deleting old one to prevent AttributeError
                new_photo = ImageTk.PhotoImage(image)
                
                # Clean up previous photo if it exists (after creating new one)
                if hasattr(self, 'photo') and self.photo is not None:
                    del self.photo
                
                self.photo = new_photo
                if self.image_label.winfo_exists():
                    self.image_label.config(image=self.photo)
            except Exception as e:
                # If anything goes wrong, ensure photo attribute exists
                if not hasattr(self, 'photo'):
                    self.photo = None
                print(f"Error updating image display: {e}")
        else:
            if self.image_label.winfo_exists():
                self.image_label.config(image='')
            if hasattr(self, 'photo'):
                del self.photo

    def update_console(self):
        if not hasattr(self, 'text_widget') or not self.text_widget.winfo_exists():
            return
            
        self.text_widget.config(state=tk.NORMAL)
        
        # Get all text and split into lines
        text = self.log_buffer.getvalue()
        lines = text.splitlines()
        
        # Keep only the last MAX_LINES
        if len(lines) > self.MAX_LINES:
            # Join the last MAX_LINES with newlines
            text = '\n'.join(lines[-self.MAX_LINES:]) + '\n'
            # Update the buffer with truncated text
            self.log_buffer.truncate(0)
            self.log_buffer.seek(0)
            self.log_buffer.write(text)
        
        # Update the text widget with formatting support
        self.text_widget.delete(1.0, tk.END)
        
        # Parse text for [BOLD]...[/BOLD] markers and apply formatting
        pattern = r'\[BOLD\](.*?)\[/BOLD\]'
        last_end = 0
        
        for match in re.finditer(pattern, text):
            # Insert text before the bold marker
            if match.start() > last_end:
                self.text_widget.insert(tk.END, text[last_end:match.start()])
            # Insert bold text
            self.text_widget.insert(tk.END, match.group(1), 'bold')
            last_end = match.end()
        
        # Insert remaining text after last match
        if last_end < len(text):
            self.text_widget.insert(tk.END, text[last_end:])
        
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.see(tk.END)

    def write(self, message):
        """Write to the console window if it exists"""
        if not self.window.winfo_exists():
            return
        
        # Check buffer size and truncate if too large to prevent memory issues
        try:
            buffer_size = len(self.log_buffer.getvalue().encode('utf-8'))
            if buffer_size > MAX_LOG_BUFFER_SIZE:
                # Keep only last MAX_LINES to prevent memory issues
                text = self.log_buffer.getvalue()
                lines = text.splitlines()
                if len(lines) > self.MAX_LINES:
                    text = '\n'.join(lines[-self.MAX_LINES:]) + '\n'
                    self.log_buffer.truncate(0)
                    self.log_buffer.seek(0)
                    self.log_buffer.write(text)
        except Exception:
            pass  # If buffer check fails, continue anyway
            
        self.log_buffer.write(message)  # Write to the buffer
        self.update_console()  # Update the console window with line limit
        if self.show_image_var.get():  # Update image if checkbox is checked
            self.update_image_display()

    def flush(self):
        pass

    def save_log(self):
        # Get the current date and time
        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # Get the name of the save file
        save_file_name = self.layout_file_var.get().split('/')[-1].split('.')[0]
        # Suggest a file name
        suggested_name = f"Log_{save_file_name}_{current_time}.txt"
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=suggested_name, filetypes=[("Text files", "*.txt")])
        if file_path:
            with open(file_path, 'w') as f:
                f.write(self.log_buffer.getvalue())
            print(f"Log saved to {file_path}\n--------------------------")
     
            
    def save_image(self):
        """Save the currently displayed image"""
        if not self.window.winfo_exists():
            return
            
        area_name = self.latest_area_name_var.get()
        latest_image = self.latest_images.get(area_name)  # Access the image for the current area
        if not isinstance(latest_image, Image.Image):
            messagebox.showerror("Error", "No image to save.")
            return

        current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        suggested_name = f"{area_name}_{current_time}.png"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=suggested_name,
            filetypes=[("PNG files", "*.png")]
        )
        if file_path:
            latest_image.save(file_path, "PNG")
            print(f"Image saved to {file_path}\n--------------------------")

    def clear_console(self):
        """Clear the console text widget and log buffer"""
        if not self.window.winfo_exists():
            return
            
        # Clear the text widget
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.delete(1.0, tk.END)
        self.text_widget.config(state=tk.DISABLED)
        
        # Clear the log buffer
        self.log_buffer.seek(0)
        self.log_buffer.truncate(0)
        
        # Add a confirmation message

