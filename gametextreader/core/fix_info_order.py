import os

file_path = r'c:\Users\Morten\Desktop\Game reader debug\Fresh_new_tr\gametextreader\core\game_text_reader.py'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Verify the anchor lines to ensure we are editing the right place
start_line = 4053 # 1-based, index 4052
end_line = 4310   # 1-based, index 4309

# Re-read to confirm content matches
# The content I want to replace spans from the start of the async check block
# to the start of the update section frame creation.

# Line 4053: "        # Check Tesseract installation status ASYNCHRONOUSLY to prevent UI freeze"
# Line 4310: "            foreground='black',"

# I want to insert tesseract_content_frame creation before line 4053 (or just after comments)
# And then update all parent references in the big block.

# Let's locate the block by specific unique lines since line numbers might be off slightly.
start_marker = "        # Check Tesseract installation status ASYNCHRONOUSLY to prevent UI freeze"
end_marker = "        # Title - will be updated when changelog is displayed"

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if start_marker in line:
        start_idx = i
    if end_marker in line:
        end_idx = i
        break

if start_idx == -1 or end_idx == -1:
    print(f"Could not find markers. Start: {start_idx}, End: {end_idx}")
    # Fallback to direct line access if markers fail but we trust line numbers roughly?
    # No, risky. Let's try to match partials.
    if start_idx == -1:
        for i, line in enumerate(lines):
            if "loading_label = ttk.Label" in line and "Checking Tesseract status" in line:
                 start_idx = i - 3 # Go back a bit to catch comments
                 break

if start_idx == -1 or end_idx == -1:
    print("FATAL: Could not locate code block.")
    exit(1)

# Now we construct the new content.
# We need to construct the entire block from scratch or modify lines?
# Since it's a large block with many replacements (parent frame change), it's safer to rewrite the block using the known good structure.

new_block = [
    '        # Helper frame to ensure Tesseract status appears ABOVE updates\n',
    '        tesseract_content_frame = ttk.Frame(tesseract_status_frame)\n',
    '        tesseract_content_frame.pack(anchor="w", fill="x", pady=0)\n',
    '        \n',
    '        # Check Tesseract installation status ASYNCHRONOUSLY to prevent UI freeze\n',
    '        # Show loading state in the correct position (top)\n',
    '        loading_label = ttk.Label(tesseract_content_frame, text="Checking Tesseract status...", font=("Helvetica", 10, "italic"))\n',
    '        loading_label.pack(anchor="w", pady=10)\n',
    '        \n',
    '        def update_tesseract_ui(tesseract_installed, tesseract_message):\n',
    '            if not info_window.winfo_exists():\n',
    '                return\n',
    '            \n',
    '            try:\n',
    '                loading_label.destroy()\n',
    '            except:\n',
    '                pass\n',
    '\n',
    '            \n',
    '            # Status label with appropriate color\n',
    '            status_color = "green" if tesseract_installed else "red"\n',
    '            status_text = "✓ " if tesseract_installed else "✗ "\n',
    '            \n',
    '            if tesseract_installed:\n',
    '                # Simple status when installed - improved layout for narrow width\n',
    '                status_row = ttk.Frame(tesseract_content_frame)\n',
    '                status_row.pack(anchor="w", pady=(0, 8), fill="x")\n',
    '                \n',
    '                # Status line\n',
    '                status_line = ttk.Frame(status_row)\n',
    '                status_line.pack(anchor="w", fill="x")\n',
    '                \n',
    '                # Black text for main status\n',
    '                main_status_label = ttk.Label(\n',
    '                    status_line,\n',
    '                    text="Tesseract OCR Status: ",\n',
    '                    font=("Helvetica", 10, "bold"),\n',
    '                    foreground="black"\n',
    '                )\n',
    '                main_status_label.pack(side="left")\n',
    '                \n',
    '                # Green checkmark\n',
    '                checkmark_label = ttk.Label(\n',
    '                    status_line,\n',
    '                    text=status_text,\n',
    '                    font=("Helvetica", 10, "bold"),\n',
    '                    foreground=status_color\n',
    '                )\n',
    '                checkmark_label.pack(side="left")\n',
    '                \n',
    '                # Green text for (Installed)\n',
    '                installed_label = ttk.Label(\n',
    '                    status_line,\n',
    '                    text="(Installed)",\n',
    '                    font=("Helvetica", 10, "bold"),\n',
    '                    foreground="green"\n',
    '                )\n',
    '                installed_label.pack(side="left")\n',
    '                \n',
    '                # Add "Locate Tesseract" button on new line if needed\n',
    '                locate_button = ttk.Button(\n',
    '                    status_row,\n',
    '                    text="Set custom path... ",\n',
    '                    command=self.locate_tesseract_executable\n',
    '                )\n',
    '                locate_button.pack(anchor="w", pady=(5, 0))\n',
    '            else:\n',
    '                # Detailed status when not installed - improved layout for narrow width\n',
    '                status_row = ttk.Frame(tesseract_content_frame)\n',
    '                status_row.pack(anchor="w", pady=(0, 8), fill="x")\n',
    '                \n',
    '                # Status line\n',
    '                status_line = ttk.Frame(status_row)\n',
    '                status_line.pack(anchor="w", fill="x")\n',
    '                \n',
    '                # Black text for main status\n',
    '                main_status_label = ttk.Label(\n',
    '                    status_line,\n',
    '                    text="Tesseract OCR Status: ",\n',
    '                    font=("Helvetica", 10, "bold"),\n',
    '                    foreground="black"\n',
    '                )\n',
    '                main_status_label.pack(side="left")\n',
    '                \n',
    '                # Red X\n',
    '                x_label = ttk.Label(\n',
    '                    status_line,\n',
    '                    text=status_text,\n',
    '                    font=("Helvetica", 10, "bold"),\n',
    '                    foreground=status_color\n',
    '                )\n',
    '                x_label.pack(side="left")\n',
    '                \n',
    '                # Required text on new line for better wrapping\n',
    '                required_label = ttk.Label(\n',
    '                    status_row,\n',
    '                    text=f"(Required for {APP_NAME} to fully function)",\n',
    '                    font=("Helvetica", 9, "bold"),\n',
    '                    foreground="red",\n',
    '                    wraplength=text_wraplength,\n',
    '                    justify="left"\n',
    '                )\n',
    '                required_label.pack(anchor="w", pady=(3, 0))\n',
    '                \n',
    '                # Add "Locate Tesseract" button\n',
    '                locate_button_not_installed = ttk.Button(\n',
    '                    status_row,\n',
    '                    text="Set custom path...",\n',
    '                    command=self.locate_tesseract_executable\n',
    '                )\n',
    '                locate_button_not_installed.pack(anchor="w", pady=(5, 0))\n',
    '                \n',
    '                # Reason label - wrap text with better formatting\n',
    '                reason_label = ttk.Label(\n',
    '                    tesseract_content_frame,\n',
    '                    text=f"Reason: {tesseract_message}",\n',
    '                    font=("Helvetica", 10),\n',
    '                    foreground="red",\n',
    '                    wraplength=text_wraplength,\n',
    '                    justify="left"\n',
    '                )\n',
    '                reason_label.pack(anchor="w", pady=(0, 8))\n',
    '            \n',
    '            # Download instruction and clickable URLs - improved formatting for narrow width\n',
    '            download_label = ttk.Label(tesseract_content_frame,\n',
    '                                       text="Tesseract OCR Download links:",\n',
    '                                       font=("Helvetica", 10, "bold"),\n',
    '                                       foreground="black",\n',
    '                                       wraplength=text_wraplength,\n',
    '                                       justify="left")\n',
    '            download_label.pack(anchor="w", pady=(0, 5))\n',
    '            \n',
    '            # Links stacked vertically for better readability in narrow column\n',
    '            links_container = ttk.Frame(tesseract_content_frame)\n',
    '            links_container.pack(anchor="w", fill="x", pady=(0, 10))\n',
    '            \n',
    '            # First link to Tesseract releases page\n',
    '            releases_frame = ttk.Frame(links_container)\n',
    '            releases_frame.pack(anchor="w", pady=(0, 3), fill="x")\n',
    '            \n',
    '            releases_text = ttk.Label(releases_frame,\n',
    '                                       text="Releases page:",\n',
    '                                       font=("Helvetica", 9),\n',
    '                                       foreground="black")\n',
    '            releases_text.pack(side="left")\n',
    '            \n',
    '            tesseract_link = ttk.Label(releases_frame,\n',
    '                                       text="https://github.com/tesseract-ocr/tesseract/releases",\n',
    '                                       font=("Helvetica", 9),\n',
    '                                       foreground="blue",\n',
    '                                       cursor="hand2")\n',
    '            tesseract_link.pack(side="left", padx=(5, 0))\n',
    '            tesseract_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases"))\n',
    '            tesseract_link.bind("<Enter>", lambda e: tesseract_link.configure(font=("Helvetica", 9, "underline")))\n',
    '            tesseract_link.bind("<Leave>", lambda e: tesseract_link.configure(font=("Helvetica", 9)))\n',
    '            \n',
    '            # Direct download link for Windows installer\n',
    '            installer_frame = ttk.Frame(links_container)\n',
    '            installer_frame.pack(anchor="w", pady=(0, 0), fill="x")\n',
    '            \n',
    '            installer_text = ttk.Label(installer_frame,\n',
    '                                   text="Direct download link to installer:",\n',
    '                                   font=("Helvetica", 9),\n',
    '                                   foreground="black")\n',
    '            installer_text.pack(side="left")\n',
    '            \n',
    '            direct_link = ttk.Label(installer_frame,\n',
    '                                   text="tesseract-ocr-w64-setup-5.5.0.20241111.exe",\n',
    '                                   font=("Helvetica", 9),\n',
    '                                   foreground="blue",\n',
    '                                   cursor="hand2")\n',
    '            direct_link.pack(side="left", padx=(5, 0))\n',
    '            direct_link.bind("<Button-1>", lambda e: open_url("https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"))\n',
    '            direct_link.bind("<Enter>", lambda e: direct_link.configure(font=("Helvetica", 9, "underline")))\n',
    '            direct_link.bind("<Leave>", lambda e: direct_link.configure(font=("Helvetica", 9)))\n',
    '            \n',
    '            # Add NaturalVoiceSAPIAdapter information with reduced spacing\n',
    '            \n',
    '            # NaturalVoiceSAPIAdapter section - improved formatting\n',
    '            natural_voice_frame = ttk.Frame(tesseract_content_frame)\n',
    '            natural_voice_frame.pack(anchor="w", pady=(20, 0))\n',
    '            \n',
    '            natural_voice_title = ttk.Label(\n',
    '                natural_voice_frame,\n',
    '                text="More Voice Options:",\n',
    '                font=("Helvetica", 11, "bold"),\n',
    '                foreground="black",\n',
    '                wraplength=text_wraplength,\n',
    '                justify="left"\n',
    '            )\n',
    '            natural_voice_title.pack(anchor="w", pady=(0, 5))\n',
    '            \n',
    '            natural_voice_label = ttk.Label(\n',
    '                natural_voice_frame,\n',
    '                text="NaturalVoiceSAPIAdapter by gexgd0419",\n',
    '                font=("Helvetica", 9),\n',
    '                foreground="black",\n',
    '                wraplength=text_wraplength,\n',
    '                justify="left"\n',
    '            )\n',
    '            natural_voice_label.pack(anchor="w", pady=(0, 5))\n',
    '            \n',
    '            # Download text and link - on the same line\n',
    '            download_frame = ttk.Frame(natural_voice_frame)\n',
    '            download_frame.pack(anchor="w", pady=(0, 5), fill="x")\n',
    '            \n',
    '            download_text_label = ttk.Label(\n',
    '                download_frame,\n',
    '                text="Download can be found here:",\n',
    '                font=("Helvetica", 9),\n',
    '                foreground="black"\n',
    '            )\n',
    '            download_text_label.pack(side="left")\n',
    '            \n',
    '            natural_voice_link = ttk.Label(\n',
    '                download_frame,\n',
    '                text="https://github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases",\n',
    '                font=("Helvetica", 9),\n',
    '                foreground="blue",\n',
    '                cursor="hand2"\n',
    '            )\n',
    '            natural_voice_link.pack(side="left", padx=(5, 0))\n',
    '            natural_voice_link.bind("<Button-1>", lambda e: open_url("https://github.com/gexgd0419/NaturalVoiceSAPIAdapter/releases"))\n',
    '            natural_voice_link.bind("<Enter>", lambda e: natural_voice_link.configure(font=("Helvetica", 9, "underline")))\n',
    '            natural_voice_link.bind("<Leave>", lambda e: natural_voice_link.configure(font=("Helvetica", 9)))\n',
    '            \n',
    '            natural_voice_note = ttk.Label(\n',
    '                natural_voice_frame,\n',
    '                text="Note! Online voices may take a moment to load when first activated.",\n',
    '                font=("Helvetica", 9),\n',
    '                foreground="black",\n',
    '                wraplength=text_wraplength,\n',
    '                justify="left"\n',
    '            )\n',
    '            natural_voice_note.pack(anchor="w", pady=(0, 0))\n',
    '\n',
    '        def run_check():\n',
    '            try:\n',
    '                # Add a small delay regarding UI creation\n',
    '                import time\n',
    '                time.sleep(0.1)\n',
    '                installed, message = self.check_tesseract_installed()\n',
    '                self.root.after(0, update_tesseract_ui, installed, message)\n',
    '            except Exception as e:\n',
    '                print(f"Error checking Tesseract: {e}")\n',
    '                self.root.after(0, update_tesseract_ui, False, str(e))\n',
    '        \n',
    '        # Start the check in a background thread\n',
    '        threading.Thread(target=run_check, daemon=True).start()\n',
    '        # "News / Updates:" section (moved into tesseract_status_frame to start higher, alongside banners)\n',
    '        update_section_frame = ttk.Frame(tesseract_status_frame)\n',
    '        update_section_frame.pack(anchor="w", pady=(5, 0), fill="x")\n',
    '        # Don\'t lift tesseract_status_frame - it would hide the banner container\n',
    '        # The banner container should be visible above the text\n',
    '        \n',
    '        # Title - will be updated when changelog is displayed\n'
]

# We need to preserve the indent level for `update_title` if we cut it off.
# The original file has `update_title` definition after `update_section_frame` creation.
# My `end_idx` points to "        # Title - will be updated when changelog is displayed"
# So I should stop replacing BEFORE this line.

replacement_lines = new_block

# Construct final file content
final_lines = lines[:start_idx] + replacement_lines + lines[end_idx:]

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(final_lines)

print("Successfully updated file with correct ordering.")
