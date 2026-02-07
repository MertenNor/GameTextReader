import os

file_path = r'c:\Users\Morten\Desktop\Game reader debug\Fresh_new_tr\gametextreader\core\game_text_reader.py'

with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Verify the anchor lines to ensure we are editing the right place
start_line = 4053 # 1-based
end_line = 4272   # 1-based

# Adjust to 0-based
start_idx = start_line - 1
end_idx = end_line - 1

# Check content matches expectation (roughly)
if "Check Tesseract installation status" not in lines[start_idx]:
    print(f"Error: Line {start_line} content mismatch: {lines[start_idx].strip()}")
    exit(1)

if "natural_voice_note.pack" not in lines[end_idx - 1] and "natural_voice_note.pack" not in lines[end_idx]:
    print(f"Error: Line {end_line} content mismatch: {lines[end_idx].strip()}")
    # Check nearby
    if "natural_voice_note.pack" in lines[end_idx-2]:
        print("Found it nearby at -2")
        end_idx -= 1
    else:
        exit(1)

# Define the new header code
header_code = [
    '        # Check Tesseract installation status ASYNCHRONOUSLY to prevent UI freeze\n',
    '        # Show loading state\n',
    '        loading_label = ttk.Label(tesseract_status_frame, text="Checking Tesseract status...", font=("Helvetica", 10, "italic"))\n',
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
    '\n'
]

# Define the footer code
footer_code = [
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
    '        threading.Thread(target=run_check, daemon=True).start()\n'
]

# Processing:
# 1. Keep lines before `start_idx`.
# 2. Insert header_code.
# 3. Take lines from `start_idx + 2` (skipping the synchronous check lines) up to `end_idx`.
# 4. Indent them by 4 spaces.
# 5. Insert footer_code.
# 6. Keep remaining lines.

# Lines to be indented start from 4056 (original file), which is index 4055.
# Lines 4053 (4052) and 4054 (4053) are the check lines we remove.
# Line 4055 (4054) is a blank line. We can keep it or indent it.

body_start_idx = start_idx + 2 # Skip the comment and the check call
# Actually, the original line 4053 was comment: "# Check Tesseract installation status"
# Line 4054 was call.
# I want to remove these two.

body_lines = lines[body_start_idx : end_idx + 1] # Inclusive of end_line
indented_body = ['    ' + line for line in body_lines]

new_lines = lines[:start_idx] + header_code + indented_body + footer_code + lines[end_idx + 1:]

with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Successfully updated file.")
