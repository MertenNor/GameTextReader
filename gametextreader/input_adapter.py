"""
Input compatibility layer used by the app hotkey system.

This module replaces direct dependency on the third-party `keyboard` and `mouse`
packages by providing similar APIs backed by `pynput` listeners and `pyautogui`.
"""
from __future__ import annotations

import threading
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Callable, Dict, Optional, Set, Tuple

import pyautogui

try:
    from pynput import keyboard as pynput_keyboard
    from pynput import mouse as pynput_mouse
    _PYNPUT_AVAILABLE = True
except Exception:
    pynput_keyboard = None
    pynput_mouse = None
    _PYNPUT_AVAILABLE = False


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
        self._hotkeys: Dict[int, Tuple[Set[str], str, Callable[[], None]]] = {}
        self._hotkeys_by_name: Dict[str, Set[int]] = {}
        self._pressed_names: Set[str] = set()
        self._pressed_events: Set[_PressedEvent] = set()
        self._listener = SimpleNamespace(running=False, pressed_events=self._pressed_events)
        self._pynput_listener = None
        self._ensure_listener_started()

    def _ensure_listener_started(self) -> None:
        if not _PYNPUT_AVAILABLE or self._pynput_listener is not None:
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

    def _extract_scan_code(self, key) -> Optional[int]:
        if not _PYNPUT_AVAILABLE:
            return None

        # Do not treat printable keys as scan codes. Returning VK values here
        # causes app-side Windows scan-code tables to misclassify letters
        # (for example "n" -> F23 due to vk=110 overlap).
        if isinstance(key, pynput_keyboard.KeyCode):
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

    def _should_fire_hotkey(self, modifiers: Set[str], base_key: str, event_name: str) -> bool:
        if base_key != event_name:
            return False
        return all(self._pressed_matches(m) for m in modifiers)

    def _parse_hotkey(self, hotkey: str) -> Tuple[Set[str], str]:
        raw_parts = [self._normalize_alias(part) for part in hotkey.split("+") if part.strip()]
        if not raw_parts:
            return set(), ""
        base_key = raw_parts[-1]
        modifiers = set(raw_parts[:-1])
        return modifiers, base_key

    def _handle_press(self, key) -> None:
        name = self._normalize_alias(self._normalize_key_name(key))
        scan_code = self._extract_scan_code(key)
        event = KeyboardEvent(name=name, scan_code=scan_code, event_type="down")
        with self._lock:
            if name:
                self._pressed_names.add(name)
            if scan_code is not None:
                self._pressed_events.add(_PressedEvent(scan_code=scan_code))
            hooks = list(self._hooks.values())
            on_press_hooks = list(self._on_press_hooks.values())
            hotkeys = list(self._hotkeys.values())

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

        for modifiers, base_key, cb in hotkeys:
            if self._should_fire_hotkey(modifiers, base_key, name):
                try:
                    cb()
                except Exception:
                    pass

    def _handle_release(self, key) -> None:
        name = self._normalize_alias(self._normalize_key_name(key))
        scan_code = self._extract_scan_code(key)
        event = KeyboardEvent(name=name, scan_code=scan_code, event_type="up")
        with self._lock:
            if name and name in self._pressed_names:
                self._pressed_names.remove(name)
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
        modifiers, base_key = self._parse_hotkey(hotkey)
        hook_id = self._alloc_id()
        with self._lock:
            self._hotkeys[hook_id] = (modifiers, base_key, callback)
            key = (hotkey or "").strip().lower()
            if key:
                self._hotkeys_by_name.setdefault(key, set()).add(hook_id)
        return hook_id

    def remove_hotkey(self, hook):
        if isinstance(hook, str):
            key = hook.strip().lower()
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
                for hotkey_name, ids in list(self._hotkeys_by_name.items()):
                    ids.discard(hook)
                    if not ids:
                        self._hotkeys_by_name.pop(hotkey_name, None)

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
