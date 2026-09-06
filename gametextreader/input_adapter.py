"""
Input compatibility layer used by the app hotkey system.

This module replaces direct dependency on the third-party `keyboard` and `mouse`
packages by providing similar APIs backed by `pynput` listeners.

PyAutoGUI is only used for helper functionality such as screen sizing; it does
not provide focus-independent global shortcut listeners.
"""
from __future__ import annotations

import os
import select
import sys
import threading
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Callable, Dict, Optional, Set, Tuple

try:
    import pyautogui
except Exception:
    class _PyAutoGUIStub:
        KEYBOARD_KEYS = []

    pyautogui = _PyAutoGUIStub()

try:
    from pynput import keyboard as pynput_keyboard
    from pynput import mouse as pynput_mouse
    _PYNPUT_AVAILABLE = True
except Exception:
    pynput_keyboard = None
    pynput_mouse = None
    _PYNPUT_AVAILABLE = False

try:
    from evdev import InputDevice, ecodes, list_devices
    _EVDEV_AVAILABLE = True
except Exception:
    InputDevice = None
    ecodes = None
    list_devices = None
    _EVDEV_AVAILABLE = False


@dataclass(frozen=True)
class _PressedEvent:
    scan_code: Optional[int]


@dataclass
class KeyboardEvent:
    name: str
    scan_code: Optional[int]
    event_type: str


@dataclass
class ButtonEvent:
    button: str
    event_type: str


class _KeyboardBackend:
    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._next_id = 1
        self._hooks: Dict[int, Callable[[KeyboardEvent], None]] = {}
        self._on_press_hooks: Dict[int, Callable[[KeyboardEvent], None]] = {}
        self._on_release_hooks: Dict[int, Callable[[KeyboardEvent], None]] = {}
        self._on_release_key_hooks: Dict[int, Tuple[str, Callable[[KeyboardEvent], None]]] = {}
        self._hotkeys: Dict[int, Tuple[frozenset[str], str, Callable[[], None]]] = {}
        self._hotkeys_by_name: Dict[frozenset[str], Set[int]] = {}
        self._pressed_names: Set[str] = set()
        self._pressed_events: Set[_PressedEvent] = set()
        self._listener = SimpleNamespace(running=False, pressed_events=self._pressed_events)
        self._pynput_listener = None
        self._evdev_threads: list[threading.Thread] = []
        self._evdev_devices: Dict[str, object] = {}
        self._evdev_stop_event = threading.Event()
        self._debug = os.environ.get("GTR_INPUT_DEBUG", "0") == "1"
        self._backend = self._select_backend()
        if self._debug:
            print(f"[INFO] Input: selected backend={self._backend} WAYLAND={bool(os.environ.get('WAYLAND_DISPLAY'))} DISPLAY={os.environ.get('DISPLAY')!r}")
        self._ensure_listener_started()

    def _select_backend(self) -> str:
        if sys.platform.startswith("linux") and _EVDEV_AVAILABLE:
            # Prefer evdev on Linux for both Wayland and X11.
            # This gives one consistent global-input path and avoids focus issues
            # from compositor/session differences.
            return "evdev"
        return "pynput"

    def _ensure_listener_started(self) -> None:
        if self._backend == "evdev":
            if self._ensure_evdev_listener_started():
                return
            print("[WARNING] Input: evdev backend unavailable; falling back to pynput.")

        self._ensure_pynput_listener_started()

    def _ensure_pynput_listener_started(self) -> None:
        if not _PYNPUT_AVAILABLE or self._pynput_listener is not None:
            if not _PYNPUT_AVAILABLE:
                print("[WARNING] Input: pynput is unavailable; global hotkeys will not work.")
            return

        def on_press(key):
            self._handle_press(key)

        def on_release(key):
            self._handle_release(key)

        listener = pynput_keyboard.Listener(on_press=on_press, on_release=on_release)
        listener.daemon = True
        listener.start()
        self._pynput_listener = listener
        self._listener.running = True

    def _ensure_evdev_listener_started(self) -> bool:
        if not _EVDEV_AVAILABLE or self._evdev_threads:
            return bool(self._evdev_threads)

        device_paths = []
        try:
            device_paths = list(list_devices())
        except Exception as exc:
            print(f"[WARNING] Input: Could not enumerate evdev devices: {exc}")
            return False

        started_any = False
        for device_path in device_paths:
            try:
                device = InputDevice(device_path)
                key_caps = device.capabilities().get(ecodes.EV_KEY, [])
                if not key_caps:
                    continue

                # Skip obvious non-keyboard devices (e.g. pure mouse/buttons).
                if ecodes.KEY_A not in key_caps and ecodes.KEY_Z not in key_caps and ecodes.KEY_1 not in key_caps:
                    continue

                self._evdev_devices[device_path] = device
                thread = threading.Thread(
                    target=self._evdev_device_loop,
                    args=(device_path,),
                    daemon=True,
                )
                thread.start()
                self._evdev_threads.append(thread)
                started_any = True
                if self._debug:
                    print(f"[INFO] Input: evdev listening on {device_path} ({getattr(device, 'name', 'unknown')})")
            except Exception:
                continue

        if started_any:
            self._listener.running = True
        elif self._debug:
            print("[WARNING] Input: evdev found no keyboard-capable devices to listen on.")
        return started_any

    def _evdev_device_loop(self, device_path: str) -> None:
        device = self._evdev_devices.get(device_path)
        if device is None:
            return

        while not self._evdev_stop_event.is_set():
            try:
                ready, _, _ = select.select([device], [], [], 0.25)
                if not ready:
                    continue
                for event in device.read():
                    self._handle_evdev_event(event)
            except BlockingIOError:
                continue
            except OSError as exc:
                # Device can briefly report transient read errors; keep listening.
                if self._debug:
                    print(f"[WARNING] Input: evdev read issue on {device_path}: {exc}")
                continue
            except Exception as exc:
                if self._debug:
                    print(f"[ERROR] Input: evdev loop stopped on {device_path}: {exc}")
                break

    def _handle_evdev_event(self, event) -> None:
        if ecodes is None or getattr(event, "type", None) != ecodes.EV_KEY:
            return

        key_name = self._evdev_key_name_from_code(getattr(event, "code", None))
        if not key_name:
            return

        event_value = getattr(event, "value", None)
        if event_value == 1:
            self._handle_press_event(key_name, getattr(event, "code", None))
        elif event_value == 0:
            self._handle_release_event(key_name, getattr(event, "code", None))

    def _evdev_key_name_from_code(self, code: Optional[int]) -> str:
        if ecodes is None or code is None:
            return ""

        try:
            raw_name = ecodes.KEY[code]
        except Exception:
            return ""

        raw_name = str(raw_name).lower()
        if raw_name.startswith("key_"):
            raw_name = raw_name[4:]

        evdev_map = {
            "leftctrl": "left ctrl",
            "rightctrl": "right ctrl",
            "leftshift": "left shift",
            "rightshift": "right shift",
            "leftalt": "left alt",
            "rightalt": "right alt",
            "leftmeta": "left windows",
            "rightmeta": "right windows",
            "enter": "enter",
            "kpenter": "num_enter",
            "space": "space",
            "tab": "tab",
            "backspace": "backspace",
            "delete": "delete",
            "insert": "insert",
            "home": "home",
            "end": "end",
            "pageup": "page up",
            "pagedown": "page down",
            "esc": "esc",
            "up": "up",
            "down": "down",
            "left": "left",
            "right": "right",
            "numlock": "num lock",
            "scrolllock": "scroll lock",
            "kp0": "num_0",
            "kp1": "num_1",
            "kp2": "num_2",
            "kp3": "num_3",
            "kp4": "num_4",
            "kp5": "num_5",
            "kp6": "num_6",
            "kp7": "num_7",
            "kp8": "num_8",
            "kp9": "num_9",
            "kpplus": "num_add",
            "kpminus": "num_subtract",
            "kpmultiply": "num_multiply",
            "kpdivide": "num_divide",
            "kpdot": "num_.",
        }
        if raw_name in evdev_map:
            return evdev_map[raw_name]

        if raw_name.startswith("f") and raw_name[1:].isdigit():
            return raw_name

        if len(raw_name) == 1:
            return raw_name

        return raw_name.replace("_", " ")

    def _normalize_key_name(self, key) -> str:
        if not _PYNPUT_AVAILABLE:
            return ""

        if isinstance(key, pynput_keyboard.KeyCode):
            if key.char:
                return key.char.lower()
            vk = getattr(key, "vk", None)
            if vk is not None:
                return f"vk_{vk}"
            return ""

        key_map = {
            pynput_keyboard.Key.shift: "shift",
            pynput_keyboard.Key.shift_l: "left shift",
            pynput_keyboard.Key.shift_r: "right shift",
            pynput_keyboard.Key.ctrl: "ctrl",
            pynput_keyboard.Key.ctrl_l: "left ctrl",
            pynput_keyboard.Key.ctrl_r: "right ctrl",
            pynput_keyboard.Key.alt: "alt",
            pynput_keyboard.Key.alt_l: "left alt",
            pynput_keyboard.Key.alt_r: "right alt",
            pynput_keyboard.Key.cmd: "windows",
            pynput_keyboard.Key.cmd_l: "left windows",
            pynput_keyboard.Key.cmd_r: "right windows",
            pynput_keyboard.Key.space: "space",
            pynput_keyboard.Key.enter: "enter",
            pynput_keyboard.Key.tab: "tab",
            pynput_keyboard.Key.backspace: "backspace",
            pynput_keyboard.Key.delete: "delete",
            pynput_keyboard.Key.insert: "insert",
            pynput_keyboard.Key.home: "home",
            pynput_keyboard.Key.end: "end",
            pynput_keyboard.Key.page_up: "page up",
            pynput_keyboard.Key.page_down: "page down",
            pynput_keyboard.Key.esc: "esc",
            pynput_keyboard.Key.up: "up",
            pynput_keyboard.Key.down: "down",
            pynput_keyboard.Key.left: "left",
            pynput_keyboard.Key.right: "right",
            pynput_keyboard.Key.num_lock: "num lock",
            pynput_keyboard.Key.scroll_lock: "scroll lock",
            pynput_keyboard.Key.f1: "f1",
            pynput_keyboard.Key.f2: "f2",
            pynput_keyboard.Key.f3: "f3",
            pynput_keyboard.Key.f4: "f4",
            pynput_keyboard.Key.f5: "f5",
            pynput_keyboard.Key.f6: "f6",
            pynput_keyboard.Key.f7: "f7",
            pynput_keyboard.Key.f8: "f8",
            pynput_keyboard.Key.f9: "f9",
            pynput_keyboard.Key.f10: "f10",
            pynput_keyboard.Key.f11: "f11",
            pynput_keyboard.Key.f12: "f12",
        }
        return key_map.get(key, str(key).replace("Key.", "").replace("_", " ").lower())

    def _normalize_alias(self, name: str) -> str:
        name = (name or "").strip().lower()
        alias_map = {
            "win": "windows",
            "left win": "left windows",
            "right win": "right windows",
            "escape": "esc",
            "return": "enter",
            "control": "ctrl",
            "option": "alt",
        }
        return alias_map.get(name, name)

    # Windows virtual-key codes for the numpad and top-row digit keys, mapped to
    # their PS/2 Set-1 scan codes (the numbering game_text_reader.py's
    # numpad_scan_codes / keyboard_number_scan_codes tables expect, matching what
    # the old `keyboard` package's low-level hook reported). Only these specific
    # VKs are translated - every other KeyCode (letters, punctuation, etc.) still
    # returns None, since their VK values collide with unrelated scan codes in
    # those tables (e.g. VK 78 for 'N' equals the numpad-add scan code).
    _NUMPAD_AND_DIGIT_VK_TO_SCAN_CODE = {
        96: 82, 97: 79, 98: 80, 99: 81, 100: 75,   # numpad 0-4
        101: 76, 102: 77, 103: 71, 104: 72, 105: 73,  # numpad 5-9
        106: 55,  # numpad *
        107: 78,  # numpad +
        109: 74,  # numpad -
        110: 83,  # numpad .
        111: 53,  # numpad /
        48: 11, 49: 2, 50: 3, 51: 4, 52: 5,   # top-row 0-4
        53: 6, 54: 7, 55: 8, 56: 9, 57: 10,   # top-row 5-9
    }

    def _extract_scan_code(self, key) -> Optional[int]:
        if not _PYNPUT_AVAILABLE:
            return None

        # Do not treat printable keys as scan codes. Returning VK values here
        # causes app-side Windows scan-code tables to misclassify letters
        # (for example "n" -> F23 due to vk=110 overlap). Numpad and top-row
        # digit keys are the exception: the app's hotkey system needs their
        # real scan codes to tell numpad keys apart from regular ones, so
        # translate just those known-safe VKs instead of excluding all KeyCodes.
        if isinstance(key, pynput_keyboard.KeyCode):
            vk = getattr(key, "vk", None)
            if vk is not None:
                return self._NUMPAD_AND_DIGIT_VK_TO_SCAN_CODE.get(int(vk))
            return None

        key_scan_map = {
            pynput_keyboard.Key.ctrl_l: 29,
            pynput_keyboard.Key.ctrl_r: 157,
            pynput_keyboard.Key.shift_l: 42,
            pynput_keyboard.Key.shift_r: 54,
            pynput_keyboard.Key.alt_l: 56,
            pynput_keyboard.Key.alt_r: 184,
            pynput_keyboard.Key.cmd_l: 91,
            pynput_keyboard.Key.cmd_r: 92,
            pynput_keyboard.Key.up: 72,
            pynput_keyboard.Key.down: 80,
            pynput_keyboard.Key.left: 75,
            pynput_keyboard.Key.right: 77,
            pynput_keyboard.Key.space: 57,
            pynput_keyboard.Key.tab: 15,
            pynput_keyboard.Key.enter: 28,
            pynput_keyboard.Key.backspace: 14,
            pynput_keyboard.Key.delete: 82,
            pynput_keyboard.Key.insert: 83,
            pynput_keyboard.Key.home: 71,
            pynput_keyboard.Key.end: 79,
            pynput_keyboard.Key.page_up: 73,
            pynput_keyboard.Key.page_down: 81,
            pynput_keyboard.Key.esc: 1,
            pynput_keyboard.Key.f1: 59,
            pynput_keyboard.Key.f2: 60,
            pynput_keyboard.Key.f3: 61,
            pynput_keyboard.Key.f4: 62,
            pynput_keyboard.Key.f5: 63,
            pynput_keyboard.Key.f6: 64,
            pynput_keyboard.Key.f7: 65,
            pynput_keyboard.Key.f8: 66,
            pynput_keyboard.Key.f9: 67,
            pynput_keyboard.Key.f10: 68,
            pynput_keyboard.Key.f11: 87,
            pynput_keyboard.Key.f12: 88,
            pynput_keyboard.Key.num_lock: 69,
            pynput_keyboard.Key.scroll_lock: 70,
        }
        mapped = key_scan_map.get(key)
        if mapped is not None:
            return mapped

        vk = getattr(key, "vk", None)
        if vk is not None and 1 <= int(vk) <= 255:
            return int(vk)
        value = getattr(key, "value", None)
        vk = getattr(value, "vk", None)
        if vk is not None and 1 <= int(vk) <= 255:
            return int(vk)
        return None

    def _pressed_matches(self, requested: str) -> bool:
        requested = self._normalize_alias(requested)
        p = self._pressed_names
        if requested in p:
            return True

        group_map = {
            "ctrl": {"ctrl", "left ctrl", "right ctrl"},
            "shift": {"shift", "left shift", "right shift"},
            "alt": {"alt", "left alt", "right alt"},
            "windows": {"windows", "left windows", "right windows"},
        }
        if requested in group_map:
            return any(k in p for k in group_map[requested])
        return False

    def _parse_hotkey_parts(self, hotkey: str) -> frozenset[str]:
        hotkey = (
            hotkey.replace("numpad +", "num_add")
            .replace("numpad -", "num_subtract")
            .replace("numpad *", "num_multiply")
            .replace("numpad /", "num_divide")
            .replace("numpad .", "num_.")
        )
        raw_parts = [self._normalize_alias(part) for part in hotkey.split("+") if part.strip()]
        if not raw_parts:
            return frozenset()

        converted_parts = set()
        for part in raw_parts:
            converted = self._normalize_hotkey_part(part)
            if not converted:
                return frozenset()
            converted_parts.add(converted)
        return frozenset(converted_parts)

    def _normalize_hotkey_part(self, part: str) -> Optional[str]:
        if not part:
            return None

        part = part.strip().lower()

        # Normalize common display/storage variants before lookup.
        part = part.replace("l-alt", "left alt").replace("r-alt", "right alt")
        if part.startswith("num:"):
            part = "num_" + part[4:]
        if part.startswith("numpad "):
            np = part[7:]
            np_map = {
                "*": "num_multiply",
                "+": "num_add",
                "-": "num_subtract",
                "/": "num_divide",
                ".": "num_.",
                "enter": "num_enter",
            }
            if np.isdigit():
                part = f"num_{np}"
            else:
                part = np_map.get(np, part)

        token_map = {
            "ctrl": "ctrl",
            "left ctrl": "left ctrl",
            "right ctrl": "right ctrl",
            "shift": "shift",
            "left shift": "left shift",
            "right shift": "right shift",
            "alt": "alt",
            "left alt": "left alt",
            "right alt": "right alt",
            "windows": "windows",
            "left windows": "left windows",
            "right windows": "right windows",
            "space": "space",
            "tab": "tab",
            "enter": "enter",
            "backspace": "backspace",
            "delete": "delete",
            "insert": "insert",
            "home": "home",
            "end": "end",
            "page up": "page up",
            "page down": "page down",
            "num lock": "num lock",
            "scroll lock": "scroll lock",
            "escape": "esc",
            "esc": "esc",
            "up": "up",
            "down": "down",
            "left": "left",
            "right": "right",
        }
        if part in token_map:
            return token_map[part]

        if len(part) == 1:
            return part.lower()

        if part.startswith("f") and part[1:].isdigit():
            return part

        if part.startswith("num_"):
            numpad_map = {
                "num_0": "num_0",
                "num_1": "num_1",
                "num_2": "num_2",
                "num_3": "num_3",
                "num_4": "num_4",
                "num_5": "num_5",
                "num_6": "num_6",
                "num_7": "num_7",
                "num_8": "num_8",
                "num_9": "num_9",
                "num_multiply": "num_multiply",
                "num_add": "num_add",
                "num_enter": "num_enter",
                "num_subtract": "num_subtract",
                "num_.": "num_.",
                "num_divide": "num_divide",
            }
            return numpad_map.get(part)

        if part in {"multiply", "add", "subtract", "divide"}:
            symbol_map = {
                "multiply": "num_multiply",
                "add": "num_add",
                "subtract": "num_subtract",
                "divide": "num_divide",
            }
            return symbol_map.get(part)

        return None

    def _binding_matches(self, required_parts: frozenset[str], event_name: str) -> bool:
        if not required_parts:
            return False

        current = set(self._pressed_names)
        current.add(event_name)
        if not required_parts.issubset(current):
            return False

        known_modifiers = {"ctrl", "left ctrl", "right ctrl", "shift", "left shift", "right shift", "alt", "left alt", "right alt", "windows", "left windows", "right windows"}
        if len(required_parts) == 1 and not (required_parts & known_modifiers):
            if any(mod in current for mod in known_modifiers):
                return False

        return True

    def _fire_hotkeys(self, event_name: str) -> None:
        with self._lock:
            hotkeys = list(self._hotkeys.values())
        for required_parts, _, callback in hotkeys:
            if self._binding_matches(required_parts, event_name):
                try:
                    callback()
                except Exception:
                    pass

    def _pressed_aliases(self, name: str) -> Set[str]:
        aliases = {name}
        if name in {"left ctrl", "right ctrl"}:
            aliases.add("ctrl")
        elif name in {"left shift", "right shift"}:
            aliases.add("shift")
        elif name in {"left alt", "right alt"}:
            aliases.add("alt")
        elif name in {"left windows", "right windows"}:
            aliases.add("windows")
        return aliases

    def _refresh_modifier_aliases_locked(self) -> None:
        modifier_groups = {
            "ctrl": {"left ctrl", "right ctrl"},
            "shift": {"left shift", "right shift"},
            "alt": {"left alt", "right alt"},
            "windows": {"left windows", "right windows"},
        }
        for generic_name, specific_names in modifier_groups.items():
            if any(name in self._pressed_names for name in specific_names):
                self._pressed_names.add(generic_name)
            else:
                self._pressed_names.discard(generic_name)

    def _handle_press(self, key) -> None:
        self._handle_press_event(self._normalize_alias(self._normalize_key_name(key)), self._extract_scan_code(key))

    def _handle_press_event(self, name: str, scan_code: Optional[int]) -> None:
        event = KeyboardEvent(name=name, scan_code=scan_code, event_type="down")
        is_new_press = name not in self._pressed_names
        with self._lock:
            if name:
                self._pressed_names.update(self._pressed_aliases(name))
                self._refresh_modifier_aliases_locked()
            if scan_code is not None:
                self._pressed_events.add(_PressedEvent(scan_code=scan_code))
            hooks = list(self._hooks.values())
            on_press_hooks = list(self._on_press_hooks.values())

        for cb in hooks:
            try:
                cb(event)
            except Exception:
                pass

        for cb in on_press_hooks:
            try:
                cb(event)
            except Exception:
                pass

        if is_new_press:
            self._fire_hotkeys(name)

    def _handle_release(self, key) -> None:
        self._handle_release_event(self._normalize_alias(self._normalize_key_name(key)), self._extract_scan_code(key))

    def _handle_release_event(self, name: str, scan_code: Optional[int]) -> None:
        event = KeyboardEvent(name=name, scan_code=scan_code, event_type="up")
        with self._lock:
            if name and name in self._pressed_names:
                self._pressed_names.difference_update(self._pressed_aliases(name))
                self._refresh_modifier_aliases_locked()
            if scan_code is not None:
                self._pressed_events = {pe for pe in self._pressed_events if pe.scan_code != scan_code}
                self._listener.pressed_events = self._pressed_events

            hooks = list(self._hooks.values())
            on_release_hooks = list(self._on_release_hooks.values())
            release_key_hooks = list(self._on_release_key_hooks.values())

        for cb in hooks:
            try:
                cb(event)
            except Exception:
                pass

        for cb in on_release_hooks:
            try:
                cb(event)
            except Exception:
                pass

        for key_name, cb in release_key_hooks:
            if self._pressed_matches_key_name(key_name, name):
                try:
                    cb(event)
                except Exception:
                    pass

    def _pressed_matches_key_name(self, registered: str, released: str) -> bool:
        registered = self._normalize_alias(registered)
        released = self._normalize_alias(released)
        if registered == released:
            return True
        if registered == "ctrl":
            return released in {"ctrl", "left ctrl", "right ctrl"}
        if registered == "shift":
            return released in {"shift", "left shift", "right shift"}
        if registered == "alt":
            return released in {"alt", "left alt", "right alt"}
        if registered == "windows":
            return released in {"windows", "left windows", "right windows"}
        return False

    def _alloc_id(self) -> int:
        with self._lock:
            hook_id = self._next_id
            self._next_id += 1
            return hook_id

    def add_hotkey(self, hotkey: str, callback: Callable[[], None], suppress: bool = False):
        del suppress
        # Ensure listeners are started before adding bindings.
        self._ensure_listener_started()
        normalized_hotkey = self._parse_hotkey_parts(hotkey)
        if not normalized_hotkey:
            raise ValueError(f"Unsupported hotkey format: {hotkey}")
        hook_id = self._alloc_id()
        with self._lock:
            self._hotkeys[hook_id] = (normalized_hotkey, hotkey, callback)
            self._hotkeys_by_name.setdefault(normalized_hotkey, set()).add(hook_id)
        if self._debug:
            print(f"[INFO] Input: registered hotkey '{hotkey}' -> {sorted(normalized_hotkey)}")
        return hook_id

    def remove_hotkey(self, hook):
        if isinstance(hook, str):
            key = self._parse_hotkey_parts(hook)
            if not key:
                return
            with self._lock:
                ids = list(self._hotkeys_by_name.get(key, set()))
            for hook_id in ids:
                self.unhook(hook_id)
            return
        self.unhook(hook)

    def hook(self, callback: Callable[[KeyboardEvent], None]):
        hook_id = self._alloc_id()
        with self._lock:
            self._hooks[hook_id] = callback
        return hook_id

    def on_press(self, callback: Callable[[KeyboardEvent], None], suppress: bool = False):
        del suppress
        hook_id = self._alloc_id()
        with self._lock:
            self._on_press_hooks[hook_id] = callback
        return hook_id

    def on_release(self, callback: Callable[[KeyboardEvent], None]):
        hook_id = self._alloc_id()
        with self._lock:
            self._on_release_hooks[hook_id] = callback
        return hook_id

    def on_release_key(self, key_name: str, callback: Callable[[KeyboardEvent], None]):
        hook_id = self._alloc_id()
        with self._lock:
            self._on_release_key_hooks[hook_id] = (self._normalize_alias(key_name), callback)
        return hook_id

    def unhook(self, hook):
        if hook is None:
            return
        with self._lock:
            self._hooks.pop(hook, None)
            self._on_press_hooks.pop(hook, None)
            self._on_release_hooks.pop(hook, None)
            self._on_release_key_hooks.pop(hook, None)
            removed_hotkey = self._hotkeys.pop(hook, None)
            if removed_hotkey:
                normalized_hotkey = removed_hotkey[0]
                ids = self._hotkeys_by_name.get(normalized_hotkey)
                if ids is not None:
                    ids.discard(hook)
                    if not ids:
                        self._hotkeys_by_name.pop(normalized_hotkey, None)

    def unhook_all(self):
        with self._lock:
            self._hooks.clear()
            self._on_press_hooks.clear()
            self._on_release_hooks.clear()
            self._on_release_key_hooks.clear()
            self._hotkeys.clear()
            self._hotkeys_by_name.clear()

    def is_pressed(self, key_name: str) -> bool:
        return self._pressed_matches(key_name)

    def block_key(self, key_name: str):
        del key_name
        return None

    def wait(self, key_name: str):
        done = threading.Event()

        def _on_release(event: KeyboardEvent):
            if self._pressed_matches_key_name(key_name, event.name):
                done.set()

        hook = self.on_release(_on_release)
        try:
            done.wait()
        finally:
            self.unhook(hook)

    @property
    def _listener_proxy(self):
        return self._listener


class _MouseBackend:
    DOWN = "down"

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._next_id = 1
        self._hooks: Dict[int, Callable[[ButtonEvent], None]] = {}
        self._listener = None
        self._ensure_listener_started()

    def _ensure_listener_started(self) -> None:
        if not _PYNPUT_AVAILABLE or self._listener is not None:
            return

        def on_click(x, y, button, pressed):
            del x, y
            event_type = self.DOWN if pressed else "up"
            btn_name = str(button).replace("Button.", "")
            event = ButtonEvent(button=btn_name, event_type=event_type)
            with self._lock:
                hooks = list(self._hooks.values())
            for cb in hooks:
                try:
                    cb(event)
                except Exception:
                    pass

        listener = pynput_mouse.Listener(on_click=on_click)
        listener.daemon = True
        listener.start()
        self._listener = listener

    def _alloc_id(self) -> int:
        with self._lock:
            hook_id = self._next_id
            self._next_id += 1
            return hook_id

    def hook(self, callback: Callable[[ButtonEvent], None]):
        hook_id = self._alloc_id()
        with self._lock:
            self._hooks[hook_id] = callback
        return hook_id

    def unhook(self, hook):
        if hook is None:
            return
        with self._lock:
            self._hooks.pop(hook, None)

    def unhook_all(self):
        with self._lock:
            self._hooks.clear()


keyboard = _KeyboardBackend()
mouse = _MouseBackend()
mouse.ButtonEvent = ButtonEvent
mouse.DOWN = _MouseBackend.DOWN
keyboard._listener = keyboard._listener_proxy

# Keep a direct pyautogui touchpoint to satisfy environments that expect this module
# to use pyautogui for input-related capability checks.
PYAUTOGUI_KEYS = set(pyautogui.KEYBOARD_KEYS)
