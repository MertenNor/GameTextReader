"""
Main entry point for GameTextReader application
"""
import json
import os
import shutil
import tempfile
import datetime
from tkinter import messagebox

import tkinter as tk
from PIL import Image, ImageTk

from gametextreader.constants import (
    APP_NAME, APP_VERSION, APP_DOCUMENTS_DIR, APP_LAYOUTS_DIR,
    APP_SETTINGS_PATH, APP_AUTO_READ_SETTINGS_PATH, APP_SETTINGS_FILENAME
)
from gametextreader.core.game_text_reader import GameTextReader, show_thinkr_warning, show_hotkey_conflict_warning

# Try to import tkinterdnd2 for drag and drop functionality
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False
    print("Warning: tkinterdnd2 not available. Drag and drop functionality will be disabled.")


def migrate_legacy_settings_file(root=None, app=None):
    """Offer to copy legacy settings into the new save location on first launch."""
    try:
        def copy_directory_contents(src, dst):
            """Copy all files from src into dst, preserving structure."""
            for dirpath, dirnames, filenames in os.walk(src):
                rel = os.path.relpath(dirpath, src)
                target_dir = os.path.join(dst, rel) if rel != "." else dst
                os.makedirs(target_dir, exist_ok=True)
                for name in filenames:
                    src_file = os.path.join(dirpath, name)
                    dst_file = os.path.join(target_dir, name)
                    shutil.copy2(src_file, dst_file)

        def copy_json_files(src, dst, skip_names):
            """Copy top-level JSON files from src to dst unless skipped; do not overwrite existing."""
            if not os.path.isdir(src):
                return
            os.makedirs(dst, exist_ok=True)
            for name in os.listdir(src):
                if not name.lower().endswith(".json"):
                    continue
                if name in skip_names:
                    continue
                src_file = os.path.join(src, name)
                if not os.path.isfile(src_file):
                    continue
                dst_file = os.path.join(dst, name)
                if os.path.exists(dst_file):
                    print(f"Skipped copying {name} (already exists at destination).")
                    continue
                try:
                    shutil.copy2(src_file, dst_file)
                    print(f"Copied legacy file {name} to {dst_file}")
                except Exception as e:
                    print(f"Warning: Failed to copy {name}: {e}")

        os.makedirs(APP_DOCUMENTS_DIR, exist_ok=True)
        new_path = APP_SETTINGS_PATH
        # If current app folder already has the required files, skip legacy search/prompt
        current_settings_exists = os.path.exists(new_path)
        current_units_exists = os.path.exists(os.path.join(APP_DOCUMENTS_DIR, "gamer_units.json"))
        current_layouts_exist = os.path.isdir(APP_LAYOUTS_DIR) and any(
            os.path.isfile(os.path.join(APP_LAYOUTS_DIR, f)) for f in os.listdir(APP_LAYOUTS_DIR)
        )
        if current_settings_exists and current_units_exists and current_layouts_exist:
            return
        
        # List of legacy app names to search for
        LEGACY_APP_NAMES = ["GameReader"]
        
        legacy_paths = [
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', 'gamereader_settings.json'),
            os.path.join(tempfile.gettempdir(), 'GameReader', 'gamereader_settings.json'),
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', APP_SETTINGS_FILENAME),
            os.path.join(tempfile.gettempdir(), 'GameReader', APP_SETTINGS_FILENAME),
            # Nested inside old GameReader\<APP_NAME>
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', APP_NAME, 'gamereader_settings.json'),
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', APP_NAME, APP_SETTINGS_FILENAME),
        ]
        
        # Add paths for each legacy app name
        for legacy_name in LEGACY_APP_NAMES:
            # Direct Documents folder for legacy app name
            legacy_paths.extend([
                os.path.join(os.path.expanduser('~'), 'Documents', legacy_name, 'gamereader_settings.json'),
                os.path.join(os.path.expanduser('~'), 'Documents', legacy_name, APP_SETTINGS_FILENAME),
            ])
            # Nested inside GameReader\<legacy_name>
            legacy_paths.extend([
                os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', legacy_name, 'gamereader_settings.json'),
                os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', legacy_name, APP_SETTINGS_FILENAME),
            ])
            # Temp folder paths
            legacy_paths.extend([
                os.path.join(tempfile.gettempdir(), 'GameReader', legacy_name, 'gamereader_settings.json'),
                os.path.join(tempfile.gettempdir(), 'GameReader', legacy_name, APP_SETTINGS_FILENAME),
            ])

        candidate_dirs = {
            os.path.dirname(p) for p in legacy_paths
        }
        # Add common legacy folders even if no settings file exists there
        candidate_dirs.update({
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader'),
            os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', APP_NAME),
            os.path.join(tempfile.gettempdir(), 'GameReader'),
            os.path.join(tempfile.gettempdir(), 'GameReader', APP_NAME),
        })
        # Add directories for each legacy app name
        for legacy_name in LEGACY_APP_NAMES:
            candidate_dirs.update({
                os.path.join(os.path.expanduser('~'), 'Documents', legacy_name),
                os.path.join(os.path.expanduser('~'), 'Documents', 'GameReader', legacy_name),
                os.path.join(tempfile.gettempdir(), 'GameReader', legacy_name),
            })
        # Exclude the current app directory from legacy search
        candidate_dirs = {d for d in candidate_dirs if d and os.path.abspath(d) != os.path.abspath(APP_DOCUMENTS_DIR)}

        available = [p for p in legacy_paths if os.path.exists(p)]

        selected_path = None
        copy_layouts = False
        copy_units = False
        delete_after_copy = False
        existing_settings = os.path.exists(new_path)
        copied_settings = False
        copied_units = False

        def has_assets(folder):
            return (
                os.path.isdir(os.path.join(folder, "Layouts")) or
                os.path.exists(os.path.join(folder, "gamer_units.json")) or
                any(
                    name.lower().endswith(".json")
                    for name in os.listdir(folder) if os.path.isfile(os.path.join(folder, name))
                )
            )

        def pick_best_dir(dirs):
            """Prefer a dir that is not the current app folder and that has layouts or gamer_units."""
            dirs = [d for d in dirs if os.path.isdir(d)]
            if not dirs:
                return None
            # 1) prefer with assets and not current app folder
            for d in dirs:
                if d != APP_DOCUMENTS_DIR and has_assets(d):
                    return d
            # 2) prefer any with assets
            for d in dirs:
                if has_assets(d):
                    return d
            # 3) prefer not current app folder
            for d in dirs:
                if d != APP_DOCUMENTS_DIR:
                    return d
            # 4) fallback first
            return dirs[0]

        # Collect all candidate dirs that actually have assets (layouts/units/json), excluding current app dir
        asset_dirs = [d for d in candidate_dirs if os.path.isdir(d) and has_assets(d) and os.path.abspath(d) != os.path.abspath(APP_DOCUMENTS_DIR)]

        # If we have no settings files and no assets, nothing to migrate
        if not available and not asset_dirs:
            return

        # If we have no settings files but still have legacy assets, pick the best folder anyway
        best_dir_from_assets = pick_best_dir(asset_dirs) if not available else None

        best_dir = None

        options_for_prompt = []
        for p in available:
            options_for_prompt.append(("file", p))
        for d in asset_dirs:
            options_for_prompt.append(("dir", d))

        reload_requested = False

        if root is not None and options_for_prompt:
            # Ask the user which legacy file or asset folder to import
            win = tk.Toplevel(root)
            win.title("Import Previous Settings")
            win.resizable(False, False)
            win.transient(root)
            win.grab_set()
            win.lift()
            try:
                icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
                if os.path.exists(icon_path):
                    win.iconbitmap(icon_path)
            except Exception:
                pass

            tk.Label(
                win,
                text=(
                    "Detected settings from an older install.\n"
                    "Choose which file to copy into the new saves folder, "
                    "or skip to start fresh."
                ),
                justify="left",
                wraplength=420,
            ).pack(padx=15, pady=(15, 10))

            # Pick initial option
            initial_value = None
            if available:
                initial_dir = pick_best_dir([os.path.dirname(p) for p in available])
                initial_path = next((p for p in available if os.path.dirname(p) == initial_dir), available[0])
                initial_value = f"file:{initial_path}"
            elif asset_dirs:
                initial_dir = pick_best_dir(asset_dirs)
                initial_value = f"dir:{initial_dir}"

            choice_var = tk.StringVar(value=initial_value)

            for kind, path in options_for_prompt:
                label = path
                try:
                    if kind == "file":
                        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
                        label = f"{path} (last updated {mtime})"
                    else:
                        label = f"{path} (folder with layouts/JSON)"
                except Exception:
                    pass

                tk.Radiobutton(
                    win,
                    text=label,
                    variable=choice_var,
                    value=f"{kind}:{path}",
                    anchor="w",
                    justify="left",
                    wraplength=420
                ).pack(fill="x", padx=20, pady=2)

            tk.Label(
                win,
                text=f"Selected file will be copied to:\n{new_path}",
                fg="#555555",
                justify="left",
                wraplength=420,
            ).pack(padx=15, pady=(10, 10))

            copy_layouts_var = tk.BooleanVar(value=True)
            tk.Checkbutton(
                win,
                text="Copy layouts from the old app folder (Layouts/)",
                variable=copy_layouts_var,
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(fill="x", padx=15, pady=(0, 4))

            copy_units_var = tk.BooleanVar(value=True)
            tk.Checkbutton(
                win,
                text="Copy gamer_units.json from the old app folder",
                variable=copy_units_var,
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(fill="x", padx=15, pady=(0, 4))

            delete_var = tk.BooleanVar(value=False)
            tk.Checkbutton(
                win,
                text="Delete the old settings file after it is copied",
                variable=delete_var,
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(fill="x", padx=15, pady=(0, 4))

            delete_folder_var = tk.BooleanVar(value=False)
            tk.Checkbutton(
                win,
                text="Delete the old app folder after copying (if possible)",
                variable=delete_folder_var,
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(fill="x", padx=15, pady=(0, 10))

            decision = {"path": None, "delete": False, "dir": None}

            def confirm_copy():
                sel = choice_var.get()
                if sel.startswith("file:"):
                    decision["path"] = sel[len("file:"):]
                    decision["dir"] = os.path.dirname(decision["path"])
                elif sel.startswith("dir:"):
                    decision["path"] = None
                    decision["dir"] = sel[len("dir:"):]
                decision["delete"] = delete_var.get()
                decision["copy_layouts"] = copy_layouts_var.get()
                decision["copy_units"] = copy_units_var.get()
                decision["delete_folder"] = delete_folder_var.get()
                nonlocal reload_requested
                reload_requested = True
                win.destroy()

            def skip_copy():
                decision["path"] = None
                decision["dir"] = None
                decision["delete"] = False
                decision["copy_layouts"] = False
                decision["copy_units"] = False
                decision["delete_folder"] = False
                win.destroy()

            button_frame = tk.Frame(win)
            button_frame.pack(fill="x", pady=(0, 15))
            tk.Button(button_frame, text="Copy selected", command=confirm_copy, width=14).pack(side="right", padx=(5, 15))
            tk.Button(button_frame, text="Skip", command=skip_copy, width=10).pack(side="right")

            win.protocol("WM_DELETE_WINDOW", skip_copy)
            root.wait_window(win)
            selected_path = decision["path"]
            best_dir = decision.get("dir")
            delete_after_copy = decision.get("delete", False)
            copy_layouts = decision.get("copy_layouts", False)
            copy_units = decision.get("copy_units", False)
            delete_folder = decision.get("delete_folder", False)
        elif available:
            # Fallback: automatically pick the best available legacy file
            selected_path = pick_best_dir([os.path.dirname(p) for p in available])
            selected_path = selected_path and next((p for p in available if os.path.dirname(p) == selected_path), None)
            delete_after_copy = False
            copy_layouts = True
            copy_units = True
            delete_folder = False
            best_dir = os.path.dirname(selected_path) if selected_path else None
            reload_requested = True
        else:
            # No settings file found, but we have assets to copy
            selected_path = None
            best_dir = best_dir_from_assets
            delete_after_copy = False
            # Only copy layouts/units if they exist in any asset folder
            copy_layouts = any(os.path.isdir(os.path.join(d, "Layouts")) for d in asset_dirs)
            copy_units = any(os.path.exists(os.path.join(d, "gamer_units.json")) for d in asset_dirs)
            delete_folder = False
            reload_requested = True

        # Build list of asset dirs to copy from (include selected dir if not already)
        asset_dirs_to_copy = []
        for d in asset_dirs:
            if d not in asset_dirs_to_copy:
                asset_dirs_to_copy.append(d)
        if best_dir and best_dir not in asset_dirs_to_copy and os.path.isdir(best_dir):
            asset_dirs_to_copy.append(best_dir)

        def try_copy_settings_from_dir(src_dir):
            """Copy a legacy settings file from a directory into the new settings path."""
            candidates = [
                os.path.join(src_dir, APP_SETTINGS_FILENAME),
                os.path.join(src_dir, "gamereader_settings.json"),
            ]
            # Add any *_settings.json files in the directory
            try:
                for name in os.listdir(src_dir):
                    if name.lower().endswith("_settings.json"):
                        candidates.append(os.path.join(src_dir, name))
            except Exception:
                pass

            for cand in candidates:
                if os.path.exists(cand) and os.path.isfile(cand):
                    try:
                        shutil.copy2(cand, new_path)
                        print(f"Migrated legacy settings file from {cand} to {new_path}")
                        # Remove legacy filename if it's different from the new one
                        if os.path.basename(cand) != APP_SETTINGS_FILENAME:
                            try:
                                os.remove(cand)
                                print(f"Removed old settings file: {cand}")
                            except Exception as e:
                                print(f"Warning: Copied settings but could not delete old file {cand}: {e}")
                        return True
                    except Exception as e:
                        print(f"Warning: Failed to copy settings from {cand}: {e}")
            return False

        if selected_path and os.path.exists(selected_path):
            # Always copy/rename the legacy settings into the new filename
            shutil.copy2(selected_path, new_path)
            print(f"Migrated legacy settings file from {selected_path} to {new_path}")
            copied_settings = True
            # Remove legacy filename if it's different from the new one
            legacy_name = os.path.basename(selected_path)
            if legacy_name != APP_SETTINGS_FILENAME:
                try:
                    os.remove(selected_path)
                    print(f"Removed old settings file: {selected_path}")
                except Exception as e:
                    print(f"Warning: Copied settings but could not delete old file {selected_path}: {e}")
        elif selected_path and not os.path.exists(selected_path):
            print(f"Selected legacy settings file not found at {selected_path}, skipping settings copy.")

        # If no explicit file was chosen but a dir was, try to pull settings from that dir
        if not copied_settings and best_dir and os.path.isdir(best_dir):
            if try_copy_settings_from_dir(best_dir):
                copied_settings = True
                # Attempt to remove legacy-named file in that dir
                old_file = os.path.join(best_dir, "gamereader_settings.json")
                if os.path.exists(old_file) and os.path.basename(old_file) != APP_SETTINGS_FILENAME:
                    try:
                        os.remove(old_file)
                        print(f"Removed old settings file: {old_file}")
                    except Exception as e:
                        print(f"Warning: Copied settings but could not delete old file {old_file}: {e}")

        # As a fallback, scan all asset dirs for a settings file if none copied yet
        if not copied_settings:
            for d in asset_dirs_to_copy:
                if try_copy_settings_from_dir(d):
                    copied_settings = True
                    reload_requested = True
                    # Attempt to remove legacy-named file in that dir
                    old_file = os.path.join(d, "gamereader_settings.json")
                    if os.path.exists(old_file) and os.path.basename(old_file) != APP_SETTINGS_FILENAME:
                        try:
                            os.remove(old_file)
                            print(f"Removed old settings file: {old_file}")
                        except Exception as e:
                            print(f"Warning: Copied settings but could not delete old file {old_file}: {e}")
                    break

        if asset_dirs_to_copy and copy_layouts:
            found_layouts = False
            for d in asset_dirs_to_copy:
                legacy_layouts = os.path.join(d, "Layouts")
                if os.path.isdir(legacy_layouts):
                    found_layouts = True
                    try:
                        os.makedirs(APP_LAYOUTS_DIR, exist_ok=True)
                        copy_directory_contents(legacy_layouts, APP_LAYOUTS_DIR)
                        print(f"Copied legacy layouts from {legacy_layouts} to {APP_LAYOUTS_DIR}")
                    except Exception as e:
                        print(f"Warning: Failed to copy legacy layouts from {legacy_layouts}: {e}")
                else:
                    sibling_layouts = os.path.join(os.path.dirname(d), "GameReader", "Layouts")
                    if os.path.isdir(sibling_layouts):
                        found_layouts = True
                        try:
                            os.makedirs(APP_LAYOUTS_DIR, exist_ok=True)
                            copy_directory_contents(sibling_layouts, APP_LAYOUTS_DIR)
                            print(f"Copied legacy layouts from {sibling_layouts} to {APP_LAYOUTS_DIR}")
                        except Exception as e:
                            print(f"Warning: Failed to copy legacy layouts from sibling folder {sibling_layouts}: {e}")
            if not found_layouts:
                print("No legacy Layouts folder found to copy.")

        if asset_dirs_to_copy and copy_units:
            found_units = False
            for d in asset_dirs_to_copy:
                legacy_units = os.path.join(d, "gamer_units.json")
                if os.path.exists(legacy_units):
                    found_units = True
                    try:
                        dest_units = os.path.join(APP_DOCUMENTS_DIR, "gamer_units.json")
                        shutil.copy2(legacy_units, dest_units)
                        print(f"Copied gamer_units.json from {legacy_units} to {dest_units}")
                        copied_units = True
                    except Exception as e:
                        print(f"Warning: Failed to copy gamer_units.json from {legacy_units}: {e}")
            if not found_units:
                print("No gamer_units.json found to copy.")

        if asset_dirs_to_copy:
            # Copy any other JSON files from each legacy app folder (top-level only) without overwriting
            for d in asset_dirs_to_copy:
                try:
                    skip = {
                        os.path.basename(new_path),
                        "gamer_units.json",
                    }
                    copy_json_files(d, APP_DOCUMENTS_DIR, skip_names=skip)
                except Exception as e:
                    print(f"Warning: Failed to copy additional legacy JSON files from {d}: {e}")

        if selected_path and delete_after_copy and os.path.exists(selected_path):
            try:
                os.remove(selected_path)
                print(f"Deleted legacy settings file at {selected_path}")
            except Exception as e:
                print(f"Warning: Copied settings but could not delete legacy file: {e}")

        if delete_folder and best_dir:
            try:
                if best_dir and os.path.isdir(best_dir) and best_dir != APP_DOCUMENTS_DIR:
                    shutil.rmtree(best_dir)
                    print(f"Deleted legacy app folder at {best_dir}")
                else:
                    print("Skipped deleting legacy folder (same as current folder or invalid).")
            except Exception as e:
                print(f"Warning: Could not delete legacy app folder: {e}")
        if not asset_dirs_to_copy and not best_dir:
            print("No legacy folder found to copy.")

        # If nothing was copied (skip or none found), ensure fresh settings and gamer_units exist
        if not copied_settings and not os.path.exists(new_path):
            try:
                os.makedirs(APP_DOCUMENTS_DIR, exist_ok=True)
                with open(new_path, 'w', encoding='utf-8') as f:
                    json.dump({}, f, indent=4)
                print(f"Created new settings file at {new_path}")
            except Exception as e:
                print(f"Warning: Could not create new settings file: {e}")

        gamer_units_path = os.path.join(APP_DOCUMENTS_DIR, "gamer_units.json")
        if not copied_units and not os.path.exists(gamer_units_path):
            try:
                os.makedirs(APP_DOCUMENTS_DIR, exist_ok=True)
                with open(gamer_units_path, 'w', encoding='utf-8') as f:
                    json.dump({}, f, indent=4)
                print(f"Created new gamer_units.json at {gamer_units_path}")
            except Exception as e:
                print(f"Warning: Could not create new gamer_units.json: {e}")

        # After migration, reload settings and game units into the running app (delayed to allow file writes)
        if app and root and reload_requested:
            def _reload_after_copy():
                try:
                    app.load_edit_view_settings()
                except Exception as e:
                    print(f"Warning: Failed to reload settings after migration: {e}")
                try:
                    app.game_units = app.load_game_units()
                except Exception as e:
                    print(f"Warning: Failed to reload game units after migration: {e}")
                try:
                    last_layout = app.load_last_layout_path()
                    if last_layout and os.path.exists(last_layout):
                        app._load_layout_file(last_layout)
                        print(f"Reloaded layout after migration: {last_layout}")
                except Exception as e:
                    print(f"Warning: Failed to reload layout after migration: {e}")
            try:
                root.after(1000, _reload_after_copy)
            except Exception as e:
                print(f"Warning: Could not schedule settings reload: {e}")
    except Exception as e:
        print(f"Warning: Could not migrate legacy settings file: {e}")


def main():
    """Main entry point for the application"""
    # Use TkinterDnD's Tk if available, otherwise fall back to regular tkinter
    if TKDND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except Exception as _tkdnd_error:
            print(f"Warning: TkinterDnD failed to initialize ({_tkdnd_error}). Falling back to Tk.")
            root = tk.Tk()
    else:
        root = tk.Tk()
    
    # Hide the main window during setup to prevent the "stretching" effect
    root.withdraw()
    
    # Create loading window as a Toplevel of the main root
    loading_window = tk.Toplevel(root)
    loading_window.title("Loading")
    loading_window.geometry("200x100")
    loading_window.resizable(False, False)
    # Center the loading window
    loading_window.update_idletasks()
    x = (loading_window.winfo_screenwidth() // 2) - (300 // 2)
    y = (loading_window.winfo_screenheight() // 2) - (100 // 2)
    loading_window.geometry(f"300x100+{x}+{y}")
    # Remove window decorations for a cleaner look
    loading_window.overrideredirect(True)
    
    # Add top border bar
    top_border = tk.Frame(loading_window, bg="#545252", height=2)
    top_border.pack(fill=tk.X, side=tk.TOP)
    
    # Load and display logo
    try:
        icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            # Load icon and resize to 25x25px
            icon_img = Image.open(icon_path)
            icon_img = icon_img.resize((40, 40), Image.Resampling.LANCZOS)
            icon_photo = ImageTk.PhotoImage(icon_img)
            logo_label = tk.Label(loading_window, image=icon_photo)
            logo_label.image = icon_photo  # Keep a reference
            logo_label.pack(pady=(15, 5))
    except Exception as e:
        print(f"Error loading logo for loading window: {e}")
    
    # Create label with loading text
    loading_label = tk.Label(loading_window, text=f"Loading {APP_NAME}...", font=("Helvetica", 12, "bold"))
    loading_label.pack(expand=True)
    
    # Add bottom border bar
    bottom_border = tk.Frame(loading_window, bg="#545252", height=2)
    bottom_border.pack(fill=tk.X, side=tk.BOTTOM)
    
    loading_window.update()
    
    # Set the window icon
    try:
        icon_path = os.path.join(os.path.dirname(__file__), 'Assets', 'icon.ico')
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
            print(f"Set window icon to: {icon_path}")
        else:
            print(f"Icon file not found at: {icon_path}")
    except Exception as e:
        print(f"Error setting window icon: {e}")
    
    app = GameTextReader(root)
    
    # Window close handler is set in GameTextReader.__init__ to on_window_close
    # which properly checks for unsaved changes before closing
    
    # Set the proper window size before it becomes visible
    app.root.update_idletasks()  # Ensure all widgets are properly sized
    app.resize_window(force=True)  # Calculate and set the optimal window size
    app._ensure_window_width()  # Expand width if loaded hotkeys need more space
    
    # Destroy loading window before showing main window
    loading_window.destroy()
    
    # Prompt to import legacy settings after the UI is visible to avoid blocking on the loading screen
    # DISABLED: root.after(300, lambda: migrate_legacy_settings_file(root, app))
    
    # Now show the window at the correct size
    app.root.deiconify()
    
    # Check if Tesseract OCR is installed and show warning if not
    def check_tesseract_on_startup():
        tesseract_installed, tesseract_message = app.check_tesseract_installed()
        if not tesseract_installed:
            # Prompt with OK/Cancel so users can dismiss without opening help
            open_help = messagebox.askokcancel(
                "Tesseract OCR Not Found",
                f"{APP_NAME} requires Tesseract OCR to function properly.\n\n"
                "Press OK to open the Info/Help window for installation steps, "
                "or Cancel to dismiss this notice.",
                icon="warning",
                default="ok"
            )
            if open_help:
                app.show_info()
    
    # Schedule the check after a short delay to ensure window is fully displayed
    app.root.after(500, check_tesseract_on_startup)
    
    # Check for updates on startup if auto-check is enabled
    def check_updates_on_startup():
        app.check_for_updates_on_startup()
    
    # Schedule update check after a delay to ensure window is fully displayed
    app.root.after(1000, check_updates_on_startup)
    
    # Try to load settings for Auto Read areas from temp folder (backward compatibility only)
    # Note: Settings are now stored in the layout file. This code only runs if no layout file exists.
    # If a layout file is loaded later, its settings will take precedence.
    temp_path = APP_AUTO_READ_SETTINGS_PATH
    # Check if a layout file path exists and load it if it does
    last_layout_path = app.load_last_layout_path()
    if last_layout_path and os.path.exists(last_layout_path):
        # Load the last layout file on startup
        def load_last_layout():
            try:
                app._load_layout_file(last_layout_path)
                print(f"Loaded last layout on startup: {last_layout_path}")
            except Exception as e:
                print(f"Warning: Failed to load last layout on startup: {e}")
        # Schedule layout loading after a short delay to ensure window is ready
        app.root.after(600, load_last_layout)
    # Only load from auto_read_settings.json if no layout file path exists (backward compatibility)
    if os.path.exists(temp_path) and app.areas and not (last_layout_path and os.path.exists(last_layout_path)):
        try:
            # Basic file validation
            file_size = os.path.getsize(temp_path)
            if file_size > 1024 * 1024:  # 1MB limit for auto-read settings
                print("Warning: Auto-read settings file is too large, skipping load")
            else:
                with open(temp_path, 'r', encoding='utf-8') as f:
                    all_settings = json.load(f)
                
                # Check if this is the new format (with 'areas' key) or old format
                if 'areas' in all_settings and isinstance(all_settings['areas'], dict):
                    # New format: load all Auto Read areas
                    areas_dict = all_settings['areas']
                    stop_read_on_select = all_settings.get('stop_read_on_select', False)
                    
                    # Set interrupt on new scan setting
                    app.interrupt_on_new_scan_var.set(stop_read_on_select)
                    
                    # Load settings for each Auto Read area
                    for area_name, settings in areas_dict.items():
                        # Find the matching area in the UI
                        matching_area = None
                        for area in app.areas:
                            if len(area) >= 9:
                                area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var, freeze_screen_var = area[:9]
                            else:
                                area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var = area[:8]
                                freeze_screen_var = None
                            if area_name_var.get() == area_name:
                                matching_area = (area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var, freeze_screen_var)
                                break
                        
                        if matching_area:
                            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var, freeze_screen_var = matching_area[:9] if len(matching_area) >= 9 else matching_area[:8] + (None,)
                            
                            # Load basic settings
                            preprocess_var.set(settings.get('preprocess', False))
                            speed_var.set(settings.get('speed', '100'))
                            psm_var.set(settings.get('psm', '3 (Default - Fully auto, no OSD)'))
                            
                            # Load voice
                            saved_voice = settings.get('voice', 'Select Voice')
                            if saved_voice != 'Select Voice':
                                display_name = 'Select Voice'
                                full_voice_name = None
                                
                                # Check if saved_voice is a full name (matches GetDescription)
                                for i, voice in enumerate(app.voices, 1):
                                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == saved_voice:
                                        full_voice_name = saved_voice
                                        full_name = voice.GetDescription()
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
                                        break
                                
                                if full_voice_name:
                                    voice_var.set(display_name)
                                    voice_var._full_name = full_voice_name
                            
                            # Load hotkey
                            if settings.get('hotkey'):
                                hotkey_button.hotkey = settings['hotkey']
                                display_name = settings['hotkey'].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/') if settings['hotkey'].startswith('num_') else settings['hotkey'].replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                                hotkey_button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                                app.setup_hotkey(hotkey_button, area_frame)
                            
                            # Load processing settings
                            processing_settings = settings.get('processing', {})
                            if processing_settings:
                                app.processing_settings[area_name] = {
                                    'brightness': processing_settings.get('brightness', 1.0),
                                    'contrast': processing_settings.get('contrast', 1.0),
                                    'saturation': processing_settings.get('saturation', 1.0),
                                    'sharpness': processing_settings.get('sharpness', 1.0),
                                    'blur': processing_settings.get('blur', 0.0),
                                    'hue': processing_settings.get('hue', 0.0),
                                    'exposure': processing_settings.get('exposure', 1.0),
                                    'threshold': processing_settings.get('threshold', 128),
                                    'threshold_enabled': processing_settings.get('threshold_enabled', False),
                                    'preprocess': settings.get('preprocess', False)
                                }
                                
                                # Update UI widgets if they exist
                                if hasattr(app, 'processing_settings_widgets'):
                                    widgets = app.processing_settings_widgets.get(area_name, {})
                                    if 'brightness' in widgets:
                                        widgets['brightness'].set(processing_settings.get('brightness', 1.0))
                                    if 'contrast' in widgets:
                                        widgets['contrast'].set(processing_settings.get('contrast', 1.0))
                                    if 'saturation' in widgets:
                                        widgets['saturation'].set(processing_settings.get('saturation', 1.0))
                                    if 'sharpness' in widgets:
                                        widgets['sharpness'].set(processing_settings.get('sharpness', 1.0))
                                    if 'blur' in widgets:
                                        widgets['blur'].set(processing_settings.get('blur', 0.0))
                                    if 'hue' in widgets:
                                        widgets['hue'].set(processing_settings.get('hue', 0.0))
                                    if 'exposure' in widgets:
                                        widgets['exposure'].set(processing_settings.get('exposure', 1.0))
                                    if 'threshold' in widgets:
                                        widgets['threshold'].set(processing_settings.get('threshold', 128))
                                    if 'threshold_enabled' in widgets:
                                        widgets['threshold_enabled'].set(processing_settings.get('threshold_enabled', False))
                    
                    print("Loaded Auto Read settings successfully")
                else:
                    # Old format (backward compatibility): load just the first "Auto Read" area
                    settings = all_settings
                    if app.areas:
                        if len(app.areas[0]) >= 9:
                            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var, freeze_screen_var = app.areas[0][:9]
                        else:
                            area_frame, hotkey_button, set_area_button, area_name_var, preprocess_var, voice_var, speed_var, psm_var = app.areas[0][:8]
                            freeze_screen_var = None
                        
                        # Only load if this is the first "Auto Read" area (Auto Read 1, or "Auto Read" for backward compatibility)
                        area_name_check = area_name_var.get()
                        if area_name_check == "Auto Read 1" or area_name_check == "Auto Read":
                            preprocess_var.set(settings.get('preprocess', False))
                            speed_var.set(settings.get('speed', '100'))
                            psm_var.set(settings.get('psm', '3 (Default - Fully auto, no OSD)'))
                            
                            # Load voice (same logic as above)
                            saved_voice = settings.get('voice', 'Select Voice')
                            if saved_voice != 'Select Voice':
                                display_name = 'Select Voice'
                                full_voice_name = None
                                for i, voice in enumerate(app.voices, 1):
                                    if hasattr(voice, 'GetDescription') and voice.GetDescription() == saved_voice:
                                        full_voice_name = saved_voice
                                        full_name = voice.GetDescription()
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
                                        break
                                
                                if full_voice_name:
                                    voice_var.set(display_name)
                                    voice_var._full_name = full_voice_name
                            
                            if settings.get('hotkey'):
                                hotkey_button.hotkey = settings['hotkey']
                                display_name = settings['hotkey'].replace('num_', 'num:').replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/') if settings['hotkey'].startswith('num_') else settings['hotkey'].replace('multiply', '*').replace('add', '+').replace('subtract', '-').replace('divide', '/')
                                hotkey_button.config(text=f"Set Hotkey: [ {display_name.upper()} ]")
                                app.setup_hotkey(hotkey_button, area_frame)
                            
                            processing_settings = settings.get('processing', {})
                            if processing_settings:
                                app.processing_settings['Auto Read'] = {
                                    'brightness': processing_settings.get('brightness', 1.0),
                                    'contrast': processing_settings.get('contrast', 1.0),
                                    'saturation': processing_settings.get('saturation', 1.0),
                                    'sharpness': processing_settings.get('sharpness', 1.0),
                                    'blur': processing_settings.get('blur', 0.0),
                                    'hue': processing_settings.get('hue', 0.0),
                                    'exposure': processing_settings.get('exposure', 1.0),
                                    'threshold': processing_settings.get('threshold', 128),
                                    'threshold_enabled': processing_settings.get('threshold_enabled', False),
                                    'preprocess': settings.get('preprocess', False)
                                }
                                
                                # Update UI widgets if they exist
                                if hasattr(app, 'processing_settings_widgets'):
                                    widgets = app.processing_settings_widgets.get('Auto Read', {})
                                    if 'brightness' in widgets:
                                        widgets['brightness'].set(processing_settings.get('brightness', 1.0))
                                    if 'contrast' in widgets:
                                        widgets['contrast'].set(processing_settings.get('contrast', 1.0))
                                    if 'saturation' in widgets:
                                        widgets['saturation'].set(processing_settings.get('saturation', 1.0))
                                    if 'sharpness' in widgets:
                                        widgets['sharpness'].set(processing_settings.get('sharpness', 1.0))
                                    if 'blur' in widgets:
                                        widgets['blur'].set(processing_settings.get('blur', 0.0))
                                    if 'hue' in widgets:
                                        widgets['hue'].set(processing_settings.get('hue', 0.0))
                                    if 'exposure' in widgets:
                                        widgets['exposure'].set(processing_settings.get('exposure', 1.0))
                                    if 'threshold' in widgets:
                                        widgets['threshold'].set(processing_settings.get('threshold', 128))
                                    if 'threshold_enabled' in widgets:
                                        widgets['threshold_enabled'].set(processing_settings.get('threshold_enabled', False))
        except Exception as e:
            print(f"Warning: Failed to load Auto Read settings: {e}")
    
    # Start the main event loop
    root.mainloop()


if __name__ == "__main__":
    main()

