"""
Shared helpers for sizing and remembering secondary window geometry.

Windows are created with fixed pixel geometry strings (e.g. "700x300"), but the
app runs as PROCESS_SYSTEM_DPI_AWARE (see main.py), so Tk receives raw physical
pixels with no OS-level scaling applied. At high display scaling (e.g. 300% on
a 4K screen) those fixed pixel sizes render as a small fraction of their
intended on-screen size. apply_window_geometry() scales the requested default
size by the actual DPI factor so a window opens at a sensible size the first
time, and remembers whatever size the user resizes it to afterward so it
doesn't need to be resized again on the next open.
"""
import json
import os

from .constants import APP_DOCUMENTS_DIR
from .screen_capture import get_dpi_scale

_SIZES_PATH = os.path.join(APP_DOCUMENTS_DIR, 'window_sizes.json')
_cache = None


def _load_sizes():
    global _cache
    if _cache is None:
        _cache = {}
        try:
            if os.path.exists(_SIZES_PATH):
                with open(_SIZES_PATH, 'r', encoding='utf-8') as f:
                    _cache = json.load(f)
        except Exception:
            _cache = {}
    return _cache


def _save_sizes():
    try:
        os.makedirs(APP_DOCUMENTS_DIR, exist_ok=True)
        with open(_SIZES_PATH, 'w', encoding='utf-8') as f:
            json.dump(_cache, f, indent=2)
    except Exception:
        pass


def apply_window_geometry(window, window_key, default_width, default_height, center=True, screen_margin=80, parent=None):
    """Size and position `window`, remembering the user's last resize under `window_key`.

    On first open (or if nothing was saved yet for `window_key`), the default
    size is scaled by the current display's DPI factor. After that, whatever
    size the user resizes the window to is remembered and reused next time.
    The final size is always clamped to fit the current screen, in case a
    saved size came from a different monitor/scaling setup.

    If `parent` is given, the window is centered over that widget (matching
    dialogs that were originally centered on the main window) instead of the
    full screen - useful on large displays where the main window doesn't fill
    the screen, so a screen-centered popup would appear far from it.
    """
    sizes = _load_sizes()
    saved = sizes.get(window_key)
    current_scale = get_dpi_scale()

    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    if saved and 'width' in saved and 'height' in saved:
        width, height = saved['width'], saved['height']
        # Sizes saved before this field existed have no recorded scale;
        # assume 100% rather than silently skipping the rescale, since an
        # unscaled (pre-DPI-fix) save is effectively what "1.0" means here.
        saved_scale = saved.get('dpi_scale', 1.0)
        if saved_scale and abs(saved_scale - current_scale) > 0.01:
            # The display's DPI scale has changed since this size was saved
            # (new monitor, changed Windows scaling setting, etc.) - a raw
            # pixel size from the old scale is the wrong size for fonts/
            # widgets rendered at the new one, so rescale it proportionally
            # instead of reusing it verbatim.
            ratio = current_scale / saved_scale
            width = int(width * ratio)
            height = int(height * ratio)
            _source = f"remembered size, rescaled for DPI change ({saved_scale:.2f}x -> {current_scale:.2f}x)"
        else:
            _source = f"remembered size (from a previous resize)"
    else:
        width = int(default_width * current_scale)
        height = int(default_height * current_scale)
        _source = f"DPI-scaled default ({default_width}x{default_height} x {current_scale:.2f})"

    width = max(200, min(width, screen_width - 40))
    height = max(150, min(height, screen_height - screen_margin))
    print(f"[WindowGeometry] '{window_key}': {width}x{height} ({_source})")

    if center:
        if parent is not None:
            parent.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (width // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (height // 2)
        else:
            x = (screen_width // 2) - (width // 2)
            y = (screen_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
    else:
        window.geometry(f"{width}x{height}")

    # Persist size changes (debounced) as the user resizes the window
    state = {'timer': None}

    def _persist():
        state['timer'] = None
        try:
            # A window that's been withdraw()'n (e.g. hidden right after
            # creation, or toggled closed) reports a bogus placeholder size,
            # not its real last-shown size - don't let that clobber the
            # remembered size.
            if not window.winfo_viewable():
                return
            new_size = {
                'width': window.winfo_width(),
                'height': window.winfo_height(),
                'dpi_scale': get_dpi_scale(),
            }
            sizes[window_key] = new_size
            _save_sizes()
            print(f"[WindowGeometry] '{window_key}': saved resized size {new_size['width']}x{new_size['height']} @ {new_size['dpi_scale']:.2f}x")
        except Exception:
            pass

    def on_configure(event):
        if event.widget is not window:
            return
        if not window.winfo_viewable():
            return
        if state['timer']:
            window.after_cancel(state['timer'])
        state['timer'] = window.after(500, _persist)

    window.bind('<Configure>', on_configure, add='+')

    return width, height
