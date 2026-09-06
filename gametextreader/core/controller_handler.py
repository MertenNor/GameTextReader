"""
Controller input handler for game controller support
"""
import ctypes
import threading
import time

# Controller support
try:
    import inputs
    CONTROLLER_AVAILABLE = True
except ImportError:
    CONTROLLER_AVAILABLE = False


class ControllerHandler:
    """Handles controller input detection for hotkey assignment"""
    
    def __init__(self):
        self.running = False
        self.controller_thread = None
        self.controller_available = CONTROLLER_AVAILABLE
        self.last_button_press = None
        self.button_press_event = threading.Event()
        self._lock = threading.Lock()  # Lock for thread-safe access to game_reader
        
    def start_monitoring(self):
        """Start monitoring controller inputs in a separate thread"""
        if not self.controller_available:
            return False
            
        try:
            self.running = True
            self.controller_thread = threading.Thread(target=self._monitor_controller)
            self.controller_thread.daemon = True
            self.controller_thread.start()
            return True
        except Exception as e:
            print(f"Error starting controller monitoring: {e}")
            return False
            
    def stop_monitoring(self):
        """Stop monitoring controller inputs"""
        self.running = False
        if self.controller_thread and self.controller_thread.is_alive():
            # Wait longer for thread to finish, with multiple attempts
            for _ in range(5):  # Try up to 5 times
                self.controller_thread.join(timeout=0.5)
                if not self.controller_thread.is_alive():
                    break
                # If still alive, try to interrupt by setting event
                self.button_press_event.set()
            
    def wait_for_button_press(self, timeout=10):
        """Wait for a controller button press and return the button name"""
        if not self.controller_available:
            return None
            
        self.last_button_press = None
        self.button_press_event.clear()
        
        # Start monitoring if not already running
        if not self.running:
            self.start_monitoring()
            
        # Wait for button press
        if self.button_press_event.wait(timeout):
            button_name = self.last_button_press
            # Clear state immediately after detection so button can be used again right away
            self.last_button_press = None
            self.button_press_event.clear()
            return button_name
        return None
        
    def _monitor_controller(self):
        """Monitor controller events in a loop"""
        while self.running:
            try:
                # Check running flag before blocking call
                if not self.running:
                    break
                
                # Use timeout to prevent indefinite blocking
                # Note: inputs.get_gamepad() doesn't support timeout directly,
                # but we can check running flag between iterations
                try:
                    events = inputs.get_gamepad()
                except inputs.UnpluggedError:
                    # Controller disconnected - wait and retry
                    if self.running:
                        time.sleep(1)
                    continue
                
                # Check running flag again after potentially blocking call
                if not self.running:
                    break
                
                for event in events:
                    if not self.running:
                        break
                    
                    button_name = None
                    
                    # Handle Key events (regular buttons)
                    if event.ev_type == 'Key' and event.state == 1:  # Button press
                        button_name = self._get_button_name(event.code)
                    
                    # Handle Absolute events (D-Pad only - no analog sticks)
                    elif event.ev_type == 'Absolute':
                        button_name = self._get_absolute_button_name(event.code, event.state)
                    
                    # If we got a button name, process it
                    if button_name:
                        self.last_button_press = button_name
                        self.button_press_event.set()
                        
                        # Trigger any active controller hotkeys
                        self._trigger_controller_hotkeys(button_name)
                        
                        # Notify the main class about the button press (thread-safe)
                        with self._lock:
                            if hasattr(self, 'game_reader') and self.game_reader:
                                try:
                                    self.game_reader._check_controller_hotkeys(button_name)
                                except Exception as e:
                                    print(f"Error checking controller hotkeys: {e}")
                        break
            except inputs.UnpluggedError:
                # Controller disconnected
                if self.running:
                    time.sleep(1)
            except Exception as e:
                print(f"Controller error: {e}")
                if self.running:
                    time.sleep(1)
                
    def _get_button_name(self, code):
        """Convert controller button code to readable name"""
        # Use generic button numbers that work for all controller types
        button_mapping = {
            # Face buttons - these are universal across controllers
            'BTN_SOUTH': 'Btn 1',      # A on Xbox, Cross on PlayStation, A on Nintendo
            'BTN_EAST': 'Btn 2',       # B on Xbox, Circle on PlayStation, B on Nintendo  
            'BTN_NORTH': 'Btn 3',      # Y on Xbox, Triangle on PlayStation, Y on Nintendo
            'BTN_WEST': 'Btn 4',       # X on Xbox, Square on PlayStation, X on Nintendo
            
            # Shoulder buttons
            'BTN_TL': 'Btn 5',         # LB on Xbox, L1 on PlayStation, L on Nintendo
            'BTN_TR': 'Btn 6',         # RB on Xbox, R1 on PlayStation, R on Nintendo
            
            # Stick buttons
            'BTN_THUMBL': 'Btn 7',     # LS on Xbox, L3 on PlayStation, Left Stick on Nintendo
            'BTN_THUMBR': 'Btn 8',     # RS on Xbox, R3 on PlayStation, Right Stick on Nintendo
            
            # Menu buttons
            'BTN_START': 'Btn 9',      # START on Xbox, OPTIONS on PlayStation, + on Nintendo
            'BTN_SELECT': 'Btn 10',    # SELECT on Xbox, SHARE on PlayStation, - on Nintendo
            'BTN_MODE': 'Btn 11',      # HOME on Xbox, PS Button on PlayStation, HOME on Nintendo
            
            # D-Pad buttons (digital only - no analog stick)
            'BTN_DPAD_UP': 'DPAD_UP',
            'BTN_DPAD_DOWN': 'DPAD_DOWN',
            'BTN_DPAD_LEFT': 'DPAD_LEFT',
            'BTN_DPAD_RIGHT': 'DPAD_RIGHT',
            
            # Additional D-Pad codes that some controllers use
            'BTN_DPAD_UP_ALT': 'DPAD_UP',
            'BTN_DPAD_DOWN_ALT': 'DPAD_DOWN',
            'BTN_DPAD_LEFT_ALT': 'DPAD_LEFT',
            'BTN_DPAD_RIGHT_ALT': 'DPAD_RIGHT',
            
            # Some controllers use different naming conventions
            'BTN_HAT_UP': 'DPAD_UP',
            'BTN_HAT_DOWN': 'DPAD_DOWN',
            'BTN_HAT_LEFT': 'DPAD_LEFT',
            'BTN_HAT_RIGHT': 'DPAD_RIGHT'
        }
        return button_mapping.get(code, f"Btn_{code}")
    
    def _get_absolute_button_name(self, code, state):
        """Convert absolute controller events (D-Pad only) to button names"""
        # Only handle true D-Pad events (HAT codes) - ignore analog stick movements
        # D-Pad events - many controllers use ABS_HAT0X and ABS_HAT0Y
        if code == 'ABS_HAT0X':
            if state == -1:  # Left
                return 'DPAD_LEFT'
            elif state == 1:  # Right
                return 'DPAD_RIGHT'
        elif code == 'ABS_HAT0Y':
            if state == -1:  # Up
                return 'DPAD_UP'
            elif state == 1:  # Down
                return 'DPAD_DOWN'
        
        # Alternative D-Pad codes that some controllers use
        elif code == 'ABS_HAT1X':
            if state == -1:  # Left
                return 'DPAD_LEFT'
            elif state == 1:  # Right
                return 'DPAD_RIGHT'
        elif code == 'ABS_HAT1Y':
            if state == -1:  # Up
                return 'DPAD_UP'
            elif state == 1:  # Down
                return 'DPAD_DOWN'
        
        # Return None for analog stick movements and other absolute events
        # This prevents accidental hotkey triggers from stick movements
        return None
    
    def _trigger_controller_hotkeys(self, button_name):
        """Trigger any active controller hotkeys for the given button"""
        try:
            # This method will be called by the main class to check for hotkeys
            # The actual hotkey checking is done in the main class
            pass
        except Exception as e:
            print(f"Error triggering controller hotkeys: {e}")
    
    def list_input_devices(self):
        """List all input devices (keyboard, mouse, and game controllers)"""
        devices = []
        
        # Always add keyboard and mouse (they're always present)
        devices.append("- Keyboard")
        devices.append("- Mouse")
        
        # Try to list game controllers using Windows API first (more reliable names)
        controller_names = set()
        
        # Method 1: Use Windows joystick API (joyGetDevCapsW)
        try:
            joyGetNumDevs = ctypes.windll.winmm.joyGetNumDevs
            joyGetDevCapsW = ctypes.windll.winmm.joyGetDevCapsW
            
            class JOYCAPS(ctypes.Structure):
                _fields_ = [
                    ("wMid", ctypes.c_ushort),
                    ("wPid", ctypes.c_ushort),
                    ("szPname", ctypes.c_wchar * 260),
                    ("wXmin", ctypes.c_uint),
                    ("wXmax", ctypes.c_uint),
                    ("wYmin", ctypes.c_uint),
                    ("wYmax", ctypes.c_uint),
                    ("wZmin", ctypes.c_uint),
                    ("wZmax", ctypes.c_uint),
                    ("wNumButtons", ctypes.c_uint),
                    ("wPeriodMin", ctypes.c_uint),
                    ("wPeriodMax", ctypes.c_uint),
                ]
            
            num_slots = joyGetNumDevs()
            for i in range(num_slots):
                try:
                    caps = JOYCAPS()
                    if joyGetDevCapsW(i, ctypes.byref(caps), ctypes.sizeof(JOYCAPS)) == 0:
                        name = caps.szPname.strip()
                        if name:
                            controller_names.add(name)
                except Exception:
                    pass
        except Exception:
            pass
        
        # Method 2: Try inputs library as fallback
        if self.controller_available:
            try:
                device_manager = inputs.devices
                
                # Try to get gamepads
                gamepads = []
                try:
                    if hasattr(device_manager, 'gamepads'):
                        gamepads = list(device_manager.gamepads)
                except Exception:
                    pass
                
                # Try to get joysticks
                joysticks = []
                try:
                    if hasattr(device_manager, 'joysticks'):
                        joysticks = list(device_manager.joysticks)
                except Exception:
                    pass
                
                # Extract names from gamepads and joysticks
                for gp in gamepads + joysticks:
                    try:
                        if hasattr(gp, 'get_char_name'):
                            name = gp.get_char_name()
                        elif hasattr(gp, 'name'):
                            name = gp.name
                        else:
                            name = str(gp)
                        
                        if name and name.strip():
                            # Clean up the name
                            name = name.strip()
                            # If it's a generic name, try to make it more readable
                            if 'USB' in name.upper() or 'GAMEPAD' in name.upper() or 'JOYSTICK' in name.upper():
                                controller_names.add(name)
                            elif name not in ['', 'None']:
                                controller_names.add(name)
                    except Exception:
                        pass
            except Exception:
                pass
        
        # Add controllers to the list
        for controller_name in sorted(controller_names):
            devices.append(f"- {controller_name}")
        
        return devices

