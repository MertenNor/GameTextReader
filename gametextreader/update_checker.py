"""
Update checking functionality for GameTextReader
"""
import io
import json
import os
import re
import threading
import webbrowser
import requests
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

from .constants import APP_NAME, APP_VERSION, GITHUB_REPO, UPDATE_SERVER_URL, SHOW_UPDATE_POPUP_FOR_TESTING


def version_tuple(v):
    """Convert a version string like '0.6.1' to a tuple of ints: (0,6,1)"""
    return tuple(int(x) for x in v.split('.') if x.isdigit())


def show_update_popup(root, local_version, remote_version, remote_changelog, download_url=None, is_news_update=False):
    """
    Show the update popup window. Must be called from the main thread.
    download_url: Optional custom download URL. If None, uses GitHub releases.
    is_news_update: If True, shows as news update instead of version update.
    """
    popup = tk.Toplevel(root)
    popup.title("News Update" if is_news_update else "Update Available")
    popup.geometry("750x500")  # Set initial size
    popup.minsize(400, 150)    # Set minimum size
    
    # Set the window icon
    try:
        icon_path = os.path.join(os.path.dirname(__file__), '..', 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            popup.iconbitmap(icon_path)
    except Exception as e:
        print(f"Error setting update popup icon: {e}")
    
    # Make window resizable
    popup.resizable(True, True)
    
    # Configure grid weights
    popup.grid_rowconfigure(0, weight=1)
    popup.grid_columnconfigure(0, weight=1)
    
    # Create main frame with padding
    main_frame = ttk.Frame(popup, padding="20")
    main_frame.grid(row=0, column=0, sticky='nsew')
    main_frame.grid_rowconfigure(3, weight=1)  # Make text area expandable
    main_frame.grid_columnconfigure(0, weight=1)
    
    # Load and display logo in top-right corner as overlay
    try:
        icon_path = os.path.join(os.path.dirname(__file__), '..', 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            # Load icon and convert to image
            icon_img = Image.open(icon_path)
            # Resize to fit nicely (max 80x80 for top-right corner)
            icon_img.thumbnail((70, 70), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(icon_img)
            
            # Create label for logo and place it in top-right corner
            logo_label = tk.Label(popup, image=logo_photo, bg='white')
            logo_label.image = logo_photo  # Keep reference to prevent garbage collection
            # Place in top-right corner: 30px from right, 25px from top
            logo_label.place(relx=1.0, rely=0.0, anchor='ne', x=-30, y=25)
            # Raise to top of stacking order
            logo_label.lift()
    except Exception as e:
        print(f"Error loading logo for update popup: {e}")
    
    # Version info
    version_frame = ttk.Frame(main_frame)
    version_frame.grid(row=0, column=0, sticky='w', pady=(0, 15))
    
    # Set title based on whether it's a version update or news update
    if is_news_update:
        title_text = f"News Update for {APP_NAME}!"
    else:
        title_text = f"A new version of {APP_NAME} is available!"
    
    ttk.Label(version_frame, text=title_text, 
             font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
    
    if not is_news_update:
        ttk.Label(version_frame, text=f"Current version: {local_version}", font=('Helvetica', 10)).grid(row=1, column=0, sticky='w')
        ttk.Label(version_frame, text=f"Latest version: {remote_version}", font=('Helvetica', 10)).grid(row=2, column=0, sticky='w')
    
    # Changelog section
    ttk.Label(main_frame, text="What's new:", font=('Helvetica', 10, 'bold')).grid(row=1, column=0, sticky='nw', pady=(10, 5))
    
    # Add separator line before scroll field
    separator = ttk.Separator(main_frame, orient='horizontal')
    separator.grid(row=2, column=0, sticky='ew', pady=(5, 10))
    
    # Create a frame for the text widget and scrollbar
    text_frame = ttk.Frame(main_frame)
    text_frame.grid(row=3, column=0, sticky='nsew', pady=(0, 15))
    text_frame.grid_rowconfigure(0, weight=1)
    text_frame.grid_columnconfigure(0, weight=1)
    
    # Add text widget with scrollbar
    text = tk.Text(text_frame, wrap=tk.WORD, width=60, height=10, 
                 padx=10, pady=10, relief='flat', bg='#f0f0f0')
    scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=scrollbar.set)
    
    text.grid(row=0, column=0, sticky='nsew')
    scrollbar.grid(row=0, column=1, sticky='ns')
    
    # Insert changelog text with image support
    changelog = remote_changelog if remote_changelog else "No changelog available."
    
    def insert_changelog_with_images(text_widget, changelog_text):
        """Insert changelog text and replace image markers with actual images."""
        # Pattern to match [IMAGE:url] or ![alt](url) markdown-style
        # Supports: [IMAGE:url], [IMAGE:url:width], ![alt](url)
        image_pattern = r'\[IMAGE:([^\]]+)\]|!\[([^\]]*)\]\(([^\)]+)\)'
        
        parts = []
        last_end = 0
        
        for match in re.finditer(image_pattern, changelog_text):
            # Add text before the image marker
            if match.start() > last_end:
                parts.append(('text', changelog_text[last_end:match.start()]))
            
            # Extract image URL
            if match.group(1):  # [IMAGE:url] format
                url_part = match.group(1)
                # Check if width is specified: [IMAGE:url:width]
                if ':' in url_part:
                    url, width = url_part.rsplit(':', 1)
                    try:
                        width = int(width)
                    except ValueError:
                        width = 400  # Default width
                else:
                    url = url_part
                    width = 400  # Default width
            else:  # ![alt](url) markdown format
                url = match.group(3)
                width = 400  # Default width
            
            parts.append(('image', url, width))
            last_end = match.end()
        
        # Add remaining text
        if last_end < len(changelog_text):
            parts.append(('text', changelog_text[last_end:]))
        
        # Helper function to insert text with font formatting
        def insert_text_with_fonts(text_widget, text):
            """Insert text and apply font tags like [FONT:FontName]text[/FONT] or [FONT:FontName:Size]text[/FONT]"""
            font_pattern = r'\[FONT\s*:\s*([^:\]]+?)(?:\s*:\s*(\d+))?(?:\s*:\s*(bold|normal))?\s*\](.*?)\[/FONT\]'
            
            last_end = 0
            
            # Ensure text widget is enabled for tag operations
            current_state = text_widget.cget('state')
            if current_state == 'disabled':
                text_widget.config(state='normal')
            
            matches = list(re.finditer(font_pattern, text, re.DOTALL | re.IGNORECASE))
            
            if not matches:
                # No font tags found, just insert the text as-is
                text_widget.insert('end', text)
                if current_state == 'disabled':
                    text_widget.config(state='disabled')
                return
            
            for match in matches:
                # Insert text before the font tag
                if match.start() > last_end:
                    text_widget.insert('end', text[last_end:match.start()])
                
                # Get font name, optional size, and optional weight (bold/normal)
                font_name = match.group(1).strip()
                font_size = match.group(2)
                font_weight = match.group(3)  # 'bold' or 'normal' or None
                font_text = match.group(4)
                
                # Create a unique tag name for this font
                weight_suffix = f"_{font_weight}" if font_weight else ""
                tag_name = f"font_{font_name}_{font_size or 'default'}{weight_suffix}"
                
                # Configure the tag
                try:
                    # Build font tuple: (name, size, weight) or (name, size) or (name,)
                    if font_size:
                        size = int(font_size)
                        if font_weight and font_weight.lower() == 'bold':
                            font_tuple = (font_name, size, 'bold')
                        else:
                            font_tuple = (font_name, size)
                    else:
                        # When no size specified, use default size from widget or 12
                        default_font = text_widget.cget('font')
                        if isinstance(default_font, tuple) and len(default_font) >= 2:
                            default_size = default_font[1]
                        else:
                            default_size = 12
                        if font_weight and font_weight.lower() == 'bold':
                            font_tuple = (font_name, default_size, 'bold')
                        else:
                            font_tuple = (font_name, default_size)
                    
                    text_widget.tag_config(tag_name, font=font_tuple)
                except Exception as e:
                    print(f"Warning: Could not configure font tag {tag_name}: {e}")
                    # Fallback: insert without formatting
                    text_widget.insert('end', font_text)
                    last_end = match.end()
                    continue
                
                # Insert text with the tag
                start_pos = text_widget.index('end-1c') if text_widget.get('1.0', 'end-1c').strip() else '1.0'
                text_widget.insert('end', font_text)
                end_pos = text_widget.index('end-1c')
                # Apply the tag to the inserted text
                if start_pos != end_pos:  # Only apply if there's actual text
                    text_widget.tag_add(tag_name, start_pos, end_pos)
                
                last_end = match.end()
            
            # Insert remaining text
            if last_end < len(text):
                text_widget.insert('end', text[last_end:])
            
            # Restore original state
            if current_state == 'disabled':
                text_widget.config(state='disabled')
        
        # Insert parts into text widget
        for part in parts:
            if part[0] == 'text':
                if part[1]:  # Only insert if text is not empty
                    insert_text_with_fonts(text_widget, part[1])
            elif part[0] == 'image':
                url, width = part[1], part[2]
                # Insert placeholder text
                placeholder = f"\n[Loading image from {url}...]\n"
                text_widget.insert('end', placeholder)
                image_start = text_widget.index(f'end-{len(placeholder)}c')
                
                # Load image in background (non-blocking)
                def load_and_insert_image(img_url, img_width, start_pos, placeholder_text):
                    try:
                        print(f"Loading image from: {img_url}")
                        # Download image with proper headers to avoid rate limiting
                        headers = {
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                            'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
                            'Accept-Language': 'en-US,en;q=0.9',
                            'Referer': 'https://www.google.com/'
                        }
                        img_resp = requests.get(img_url, timeout=10, stream=True, headers=headers)
                        print(f"Image response status: {img_resp.status_code}")
                        if img_resp.status_code == 200:
                            # Load image
                            img_data = img_resp.content
                            img = Image.open(io.BytesIO(img_data))
                            print(f"Image loaded: {img.size}")
                            
                            # Resize if needed (maintain aspect ratio)
                            img_width_px = img_width
                            aspect_ratio = img.height / img.width
                            img_height_px = int(img_width_px * aspect_ratio)
                            
                            # Limit max dimensions
                            max_width, max_height = 600, 400
                            if img_width_px > max_width:
                                img_width_px = max_width
                                img_height_px = int(max_width * aspect_ratio)
                            if img_height_px > max_height:
                                img_height_px = max_height
                                img_width_px = int(max_height / aspect_ratio)
                            
                            # Resize to target dimensions (static)
                            img = img.resize((img_width_px, img_height_px), Image.Resampling.LANCZOS)
                            photo = ImageTk.PhotoImage(img)
                            
                            # Replace placeholder with image (must be on main thread)
                            def insert_image():
                                try:
                                    text_widget.config(state='normal')  # Enable to modify
                                    # Find the placeholder text
                                    content = text_widget.get('1.0', 'end')
                                    placeholder_idx = content.find(placeholder_text)
                                    if placeholder_idx >= 0:
                                        # Calculate line and column
                                        lines_before = content[:placeholder_idx].count('\n')
                                        line_start = content[:placeholder_idx].rfind('\n') + 1
                                        col_start = placeholder_idx - line_start
                                        start_index = f"{lines_before + 1}.{col_start}"
                                        end_index = f"{start_index}+{len(placeholder_text)}c"
                                        
                                        text_widget.delete(start_index, end_index)
                                        text_widget.image_create(start_index, image=photo)
                                        text_widget.insert(start_index, "\n\n")
                                        # Keep reference to prevent garbage collection
                                        if not hasattr(text_widget, '_images'):
                                            text_widget._images = []
                                        text_widget._images.append(photo)
                                    text_widget.config(state='disabled')  # Disable again
                                except Exception as e:
                                    print(f"Error inserting image: {e}")
                                    text_widget.config(state='normal')
                                    content = text_widget.get('1.0', 'end')
                                    placeholder_idx = content.find(placeholder_text)
                                    if placeholder_idx >= 0:
                                        lines_before = content[:placeholder_idx].count('\n')
                                        line_start = content[:placeholder_idx].rfind('\n') + 1
                                        col_start = placeholder_idx - line_start
                                        start_index = f"{lines_before + 1}.{col_start}"
                                        end_index = f"{start_index}+{len(placeholder_text)}c"
                                        text_widget.delete(start_index, end_index)
                                        text_widget.insert(start_index, f"\n[Image loaded but failed to display: {str(e)[:50]}]\n\n")
                                    text_widget.config(state='disabled')
                            
                            root.after(0, insert_image)
                        else:
                            # Replace placeholder with error message
                            def show_error():
                                text_widget.config(state='normal')
                                content = text_widget.get('1.0', 'end')
                                placeholder_idx = content.find(placeholder_text)
                                if placeholder_idx >= 0:
                                    lines_before = content[:placeholder_idx].count('\n')
                                    line_start = content[:placeholder_idx].rfind('\n') + 1
                                    col_start = placeholder_idx - line_start
                                    start_index = f"{lines_before + 1}.{col_start}"
                                    end_index = f"{start_index}+{len(placeholder_text)}c"
                                    text_widget.delete(start_index, end_index)
                                    text_widget.insert(start_index, f"\n[Image failed to load (HTTP {img_resp.status_code}): {img_url}]\n\n")
                                text_widget.config(state='disabled')
                            root.after(0, show_error)
                    except Exception as e:
                        print(f"Exception loading image: {e}")
                        import traceback
                        traceback.print_exc()
                        # Replace placeholder with error message
                        def show_error():
                            text_widget.config(state='normal')
                            content = text_widget.get('1.0', 'end')
                            placeholder_idx = content.find(placeholder_text)
                            if placeholder_idx >= 0:
                                lines_before = content[:placeholder_idx].count('\n')
                                line_start = content[:placeholder_idx].rfind('\n') + 1
                                col_start = placeholder_idx - line_start
                                start_index = f"{lines_before + 1}.{col_start}"
                                end_index = f"{start_index}+{len(placeholder_text)}c"
                                text_widget.delete(start_index, end_index)
                                text_widget.insert(start_index, f"\n[Image error: {str(e)[:100]}]\n\n")
                            text_widget.config(state='disabled')
                        root.after(0, show_error)
                
                # Start loading image in background thread
                threading.Thread(target=load_and_insert_image, args=(url, width, image_start, placeholder), daemon=True).start()
        
        # If no image markers found, still parse and insert text with font formatting
        if not parts:
            insert_text_with_fonts(text_widget, changelog_text)
    
    # Initialize images list for garbage collection
    text._images = []
    # Don't disable text widget yet - images need to load first
    insert_changelog_with_images(text, changelog)
    # Disable after a short delay to allow images to start loading
    root.after(100, lambda: text.config(state='disabled'))  # Make text read-only after images start loading
    
    # Buttons frame
    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=4, column=0, sticky='e')
    
    def open_github():
        url = f'https://github.com/{GITHUB_REPO}/releases'
        webbrowser.open(url)
        popup.destroy()
    
    def close_popup():
        popup.destroy()
    
    ttk.Button(button_frame, text="Close", command=close_popup).pack(side='right', padx=5)
    # Only show download button if it's a version update (not just news)
    if not is_news_update:
        ttk.Button(button_frame, text="Go to download page", command=open_github).pack(side='right', padx=5)
    
    # Center the popup on screen
    popup.update_idletasks()
    width = popup.winfo_width()
    height = popup.winfo_height()
    x = (popup.winfo_screenwidth() // 2) - (width // 2)
    y = (popup.winfo_screenheight() // 2) - (height // 2)
    popup.geometry(f'{width}x{height}+{x}+{y}')
    
    # Make popup modal
    popup.transient(root)
    popup.grab_set()
    popup.wait_window()


def check_for_update(root, local_version, force=False):
    """
    Fetch version info from update server (Google Apps Script), compare to local_version.
    If remote version is newer or force=True, show a popup.
    Must be called from a background thread. The popup will be scheduled on the main thread.
    """
    if UPDATE_SERVER_URL and "YOUR_SCRIPT_ID" not in UPDATE_SERVER_URL:
        # Use Google Apps Script update server
        try:
            # Google Apps Script may need specific headers to return JSON properly
            headers = {
                'User-Agent': f'{APP_NAME}/{APP_VERSION}',
                'Accept': 'application/json'
            }
            resp = requests.get(UPDATE_SERVER_URL, timeout=10, allow_redirects=True, headers=headers)
            
            if resp.status_code == 200:
                try:
                    # Get response text and clean it
                    response_text = resp.text
                    
                    # Remove BOM (Byte Order Mark) if present
                    if response_text.startswith('\ufeff'):
                        response_text = response_text[1:]
                    
                    # Strip whitespace
                    response_text = response_text.strip()
                    
                    # Sometimes Google Apps Script returns HTML redirect page first
                    # Try to extract JSON if it's wrapped in HTML
                    if '<html' in response_text.lower() or '<!doctype' in response_text.lower():
                        # Try to find JSON in the response
                        json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', response_text, re.DOTALL)
                        if json_match:
                            response_text = json_match.group(0)
                        else:
                            raise ValueError("Response is HTML, no JSON found")
                    
                    # Ensure response starts with { (valid JSON object)
                    if not response_text.startswith('{'):
                        # Try to find where JSON starts
                        json_start = response_text.find('{')
                        if json_start > 0:
                            response_text = response_text[json_start:]
                    
                    # Try to parse JSON - use resp.json() first, fallback to json.loads()
                    try:
                        update_info = resp.json()
                    except (ValueError, json.JSONDecodeError):
                        # If resp.json() fails, try manual parsing with cleaned text
                        update_info = json.loads(response_text)
                    
                    # Validate required fields
                    if not isinstance(update_info, dict):
                        raise ValueError("Response is not a JSON object")
                    
                    remote_version = update_info.get('version')
                    remote_changelog = update_info.get('changelog', '')
                    # Always use GITHUB_REPO for download URL, ignore download_url from Google Script
                    download_url = f'https://github.com/{GITHUB_REPO}/releases'
                    
                    update_available = remote_version and version_tuple(remote_version) > version_tuple(local_version)
                    
                    # Show popup if: force is True, update is available, OR testing flag is enabled
                    if force or update_available or SHOW_UPDATE_POPUP_FOR_TESTING:
                        # Schedule popup creation on main thread
                        root.after(100, lambda: show_update_popup(
                            root, local_version, remote_version or "Unknown", remote_changelog, download_url
                        ))
                except (ValueError, KeyError, json.JSONDecodeError, TypeError) as e:
                    # JSON parsing error or missing keys
                    print(f"Error parsing update server response: {e}")
                    if force:
                        error_display = f"Unable to parse update information.\n\nError: {str(e)[:100]}\n\nCheck console for details."
                        root.after(100, lambda msg=error_display: show_update_popup(
                            root, local_version, "Unknown", msg
                        ))
            elif force:
                root.after(100, lambda: show_update_popup(
                    root, local_version, "Unknown", 
                    "Unable to fetch update information. Please check your internet connection."
                ))
        except Exception as e:
            if force:
                root.after(100, lambda: show_update_popup(
                    root, local_version, "Unknown", 
                    "Unable to fetch update information. Please check your internet connection."
                ))
            # Otherwise fail silently if no internet or any error
            pass
    else:
        # If update checking is disabled/misconfigured
        if force:
            root.after(100, lambda: show_update_popup(
                root, local_version, "Unknown", 
                "Update check not configured. Please check UPDATE_SERVER_URL in constants.py."
            ))

