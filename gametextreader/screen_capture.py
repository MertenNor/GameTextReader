"""
Screen capture functions for capturing game text areas
"""
import ctypes
from PIL import Image
import win32api
import win32con
import win32gui
import win32ui


def get_primary_monitor_info():
    """
    Get the primary monitor's position and dimensions using EnumDisplayMonitors.
    This is more reliable than MonitorFromPoint in multi-monitor setups.
    Returns: (x, y, width, height) tuple.
    """
    try:
        # Define RECT structure
        class RECT(ctypes.Structure):
            _fields_ = [
                ('left', ctypes.c_long),
                ('top', ctypes.c_long),
                ('right', ctypes.c_long),
                ('bottom', ctypes.c_long)
            ]
        
        # Define MONITORINFO structure
        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ('cbSize', ctypes.c_uint32),
                ('rcMonitor', RECT),
                ('rcWork', RECT),
                ('dwFlags', ctypes.c_uint32)
            ]
        
        # MONITORINFOF_PRIMARY = 0x00000001
        MONITORINFOF_PRIMARY = 0x00000001
        
        # Store the primary monitor info
        primary_monitor_info = None
        
        def monitor_enum_proc(hmonitor, hdc, lprcMonitor, lParam):
            """Callback function for EnumDisplayMonitors"""
            nonlocal primary_monitor_info
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            
            # Get monitor info
            if ctypes.windll.user32.GetMonitorInfoW(hmonitor, ctypes.byref(mi)):
                # Check if this is the primary monitor
                if mi.dwFlags & MONITORINFOF_PRIMARY:
                    rect = mi.rcMonitor
                    primary_monitor_info = (
                        rect.left,
                        rect.top,
                        rect.right - rect.left,
                        rect.bottom - rect.top
                    )
            return 1  # Continue enumeration
        
        # Define the callback type
        MonitorEnumProc = ctypes.WINFUNCTYPE(
            ctypes.c_int,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.POINTER(RECT),
            ctypes.c_ulong
        )
        
        # Enumerate all monitors
        callback = MonitorEnumProc(monitor_enum_proc)
        ctypes.windll.user32.EnumDisplayMonitors(None, None, callback, 0)
        
        if primary_monitor_info:
            return primary_monitor_info
        else:
            # Fallback: use GetSystemMetrics
            print("Warning: Could not find primary monitor via EnumDisplayMonitors, using fallback")
            return (0, 0, 
                   win32api.GetSystemMetrics(win32con.SM_CXSCREEN),
                   win32api.GetSystemMetrics(win32con.SM_CYSCREEN))
    
    except Exception as e:
        print(f"Error detecting primary monitor: {e}, using fallback")
        # Fallback: use GetSystemMetrics
        return (0, 0,
               win32api.GetSystemMetrics(win32con.SM_CXSCREEN),
               win32api.GetSystemMetrics(win32con.SM_CYSCREEN))


def capture_screen_area(x1, y1, x2, y2, use_printwindow=False, target_hwnd=None):
    """
    Capture screen area across multiple monitors using win32api.
    
    Args:
        x1, y1, x2, y2: Screen coordinates for the area to capture
        use_printwindow: If True, try to use PrintWindow API (better for fullscreen apps)
        target_hwnd: Window handle to capture from (for PrintWindow mode)
    """
    # Get virtual screen bounds
    min_x = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)  # Leftmost x (can be negative)
    min_y = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)  # Topmost y (can be negative)
    total_width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
    total_height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
    max_x = min_x + total_width
    max_y = min_y + total_height

    # Clamp coordinates to virtual screen bounds
    x1 = max(min_x, min(max_x, x1))
    y1 = max(min_y, min(max_y, y1))
    x2 = max(min_x, min(max_x, x2))
    y2 = max(min_y, min(max_y, y2))

    # Ensure valid area (swap if necessary and check size)
    x1, x2 = min(x1, x2), max(x1, x2)
    y1, y2 = min(y1, y2), max(y1, y2)
    width = x2 - x1
    height = y2 - y1
    if width <= 0 or height <= 0:
        return Image.new('RGB', (1, 1))  # Return a blank 1x1 image for invalid areas

    # Try PrintWindow method first if requested (better for fullscreen apps)
    if use_printwindow and target_hwnd:
        try:
            # Get window rectangle
            window_rect = win32gui.GetWindowRect(target_hwnd)
            window_width = window_rect[2] - window_rect[0]
            window_height = window_rect[3] - window_rect[1]
            
            if window_width > 0 and window_height > 0:
                # Create a device context for the window
                hwindc = win32gui.GetWindowDC(target_hwnd)
                if hwindc:
                    memdc = None
                    bmp = None
                    try:
                        srcdc = win32ui.CreateDCFromHandle(hwindc)
                        memdc = srcdc.CreateCompatibleDC()
                        
                        # Create bitmap for the window
                        bmp = win32ui.CreateBitmap()
                        bmp.CreateCompatibleBitmap(srcdc, window_width, window_height)
                        memdc.SelectObject(bmp)
                        
                        # Use PrintWindow to capture the window's content
                        # PW_RENDERFULLCONTENT = 0x00000002 (captures even if window is occluded)
                        PW_RENDERFULLCONTENT = 0x00000002
                        result = ctypes.windll.user32.PrintWindow(target_hwnd, memdc.GetHandle(), PW_RENDERFULLCONTENT)
                        
                        if result:
                            # Convert bitmap to PIL Image
                            bmpinfo = bmp.GetInfo()
                            bmpstr = bmp.GetBitmapBits(True)
                            full_img = Image.frombuffer(
                                'RGB',
                                (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                                bmpstr, 'raw', 'BGRX', 0, 1
                            )
                            
                            # Calculate crop coordinates relative to window
                            # Window position on screen
                            win_x = window_rect[0]
                            win_y = window_rect[1]
                            
                            # Convert screen coordinates to window-relative coordinates
                            crop_x1 = max(0, x1 - win_x)
                            crop_y1 = max(0, y1 - win_y)
                            crop_x2 = min(window_width, x2 - win_x)
                            crop_y2 = min(window_height, y2 - win_y)
                            
                            # Crop to the requested area
                            if crop_x2 > crop_x1 and crop_y2 > crop_y1:
                                img = full_img.crop((crop_x1, crop_y1, crop_x2, crop_y2))
                                
                                # Clean up before returning
                                try:
                                    if memdc:
                                        memdc.DeleteDC()
                                    if bmp:
                                        win32gui.DeleteObject(bmp.GetHandle())
                                except Exception:
                                    pass
                                
                                print("Fullscreen mode: Captured using PrintWindow API")
                                return img
                    except Exception as e:
                        print(f"PrintWindow capture error: {e}")
                    finally:
                        # Always clean up resources
                        try:
                            if memdc:
                                memdc.DeleteDC()
                            if bmp:
                                win32gui.DeleteObject(bmp.GetHandle())
                        except Exception:
                            pass
                        try:
                            win32gui.ReleaseDC(target_hwnd, hwindc)
                        except Exception:
                            pass
        except Exception as e:
            print(f"Error in PrintWindow capture: {e}")
            # Fall through to normal BitBlt method

    # Normal BitBlt method (fallback or default)
    # Get DC from entire virtual screen
    hwin = win32gui.GetDesktopWindow()
    hwindc = win32gui.GetWindowDC(hwin)
    memdc = None
    bmp = None
    try:
        srcdc = win32ui.CreateDCFromHandle(hwindc)
        memdc = srcdc.CreateCompatibleDC()

        # Create bitmap for capture area
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(srcdc, width, height)
        memdc.SelectObject(bmp)

        # Copy screen into bitmap
        memdc.BitBlt((0, 0), (width, height), srcdc, (x1, y1), win32con.SRCCOPY)

        # Convert bitmap to PIL Image
        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )

        return img
    finally:
        # Always clean up resources
        try:
            if memdc:
                memdc.DeleteDC()
            if bmp:
                win32gui.DeleteObject(bmp.GetHandle())
            win32gui.ReleaseDC(hwin, hwindc)
        except Exception:
            pass

