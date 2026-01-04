"""
Utility functions for GameTextReader
"""
import keyboard
import threading


class InputManager:
    """Centralized input management for all hotkey handlers.
    
    This class provides a thread-safe way to enable/disable all hotkey handlers
    without needing to unhook and rehook them. All handlers should check
    InputManager.is_allowed() before executing.
    """
    enabled = True
    _lock = threading.Lock()
    
    @classmethod
    def allow(cls):
        """Enable all hotkey handlers"""
        with cls._lock:
            cls.enabled = True
    
    @classmethod
    def block(cls):
        """Disable all hotkey handlers"""
        with cls._lock:
            cls.enabled = False
    
    @classmethod
    def is_allowed(cls):
        """Check if input is currently allowed (thread-safe)"""
        with cls._lock:
            return cls.enabled


def get_current_keyboard_layout():
    """Stub function - always returns None for simplicity"""
    return None


def normalize_key_name(key_name, scan_code=None):
    """Stub function - returns key_name as-is for simplicity"""
    return key_name


def is_special_character(key_name):
    """Check if a key name contains special characters that may cause issues"""
    if not key_name:
        return False
    
    # Check for Nordic/Special characters that commonly cause issues
    special_chars = ['å', 'ä', 'ö', '¨', '´', '`', '~', '^', '°', '§', '±', 'µ', '¶', '·', '¸', '¹', '²', '³']
    
    # Check for any special characters in the key name
    for char in special_chars:
        if char in key_name:
            return True
    
    # Check for other potentially problematic characters
    if any(ord(char) > 127 for char in key_name):  # Non-ASCII characters
        return True
    
    return False


def suggest_alternative_key(special_char):
    """Suggest alternative keys for special characters"""
    alternatives = {
        'å': 'a',
        'ä': 'a', 
        'ö': 'o',
        '¨': 'u',
        '´': "'",
        '`': "'",
        '~': '~',
        '^': '^',
        '°': 'o',
        '§': 's',
        '±': '=',
        'µ': 'u',
        '¶': 'p',
        '·': '.',
        '¸': ',',
        '¹': '1',
        '²': '2',
        '³': '3'
    }
    
    return alternatives.get(special_char, None)


def detect_ctrl_keys():
    """
    Detect which Ctrl keys are currently pressed using scan code detection.
    Returns a tuple of (left_ctrl_pressed, right_ctrl_pressed).
    This function provides more reliable left/right distinction than keyboard.is_pressed().
    """
    left_ctrl_pressed = False
    right_ctrl_pressed = False
    
    try:
        # Check if any Ctrl key is pressed first
        if keyboard.is_pressed('ctrl'):
            # Use scan code to determine which one
            # Check if _listener exists and has pressed_events before accessing
            if hasattr(keyboard, '_listener') and hasattr(keyboard._listener, 'pressed_events'):
                for event in keyboard._listener.pressed_events:
                    if hasattr(event, 'scan_code'):
                        if event.scan_code == 29:  # Left Ctrl
                            left_ctrl_pressed = True
                        elif event.scan_code == 157:  # Right Ctrl
                            right_ctrl_pressed = True
            
            # Fallback: if scan code detection fails, assume left
            if not left_ctrl_pressed and not right_ctrl_pressed:
                left_ctrl_pressed = True
    except Exception:
        # Fallback to basic detection
        if keyboard.is_pressed('ctrl'):
            left_ctrl_pressed = True
    
    return left_ctrl_pressed, right_ctrl_pressed


def _ensure_uwp_available():
    """Ensure UWP TTS is available"""
    global UWP_TTS_AVAILABLE
    try:
        from winsdk.windows.media.speechsynthesis import SpeechSynthesizer as _SS  # noqa: F401
        from winsdk.windows.storage.streams import DataReader as _DR  # noqa: F401
        UWP_TTS_AVAILABLE = True
    except Exception:
        try:
            import importlib
            importlib.import_module('winsdk')
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer as _SS  # noqa: F401
            from winsdk.windows.storage.streams import DataReader as _DR  # noqa: F401
            UWP_TTS_AVAILABLE = True
        except Exception:
            UWP_TTS_AVAILABLE = False
    return UWP_TTS_AVAILABLE


# Initialize UWP availability
try:
    import importlib
    UWP_TTS_AVAILABLE = False
    try:
        from winsdk.windows.media.speechsynthesis import SpeechSynthesizer
        from winsdk.windows.storage.streams import DataReader
        UWP_TTS_AVAILABLE = True
    except Exception:
        try:
            importlib.import_module('winsdk')
            from winsdk.windows.media.speechsynthesis import SpeechSynthesizer
            from winsdk.windows.storage.streams import DataReader
            UWP_TTS_AVAILABLE = True
        except Exception:
            UWP_TTS_AVAILABLE = False
except Exception:
    UWP_TTS_AVAILABLE = False

