# browser_commands.py
# Safe, cross-browser (Edge/Chrome) tab automation
import pyautogui
import psutil
import win32gui
import win32con
import time
import json

# ------------------------------------------------------------------
# Low-level helpers
# ------------------------------------------------------------------
def _is_browser_active():
    """Return True if Chrome or Edge is the foreground window."""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return False
    title = win32gui.GetWindowText(hwnd)
    return any(name in title.lower() for name in ("chrome", "edge", "msedge"))

def _send_hotkey_safe(*keys):
    """Send keys only if a browser is active; swallow errors."""
    try:
        if not _is_browser_active():
            return False
        pyautogui.hotkey(*keys)
        return True
    except Exception:
        return False

# ------------------------------------------------------------------
# Command implementations
# ------------------------------------------------------------------
def previous_tab():        return _send_hotkey_safe('ctrl', 'shift', 'tab')
def next_tab():            return _send_hotkey_safe('ctrl', 'tab')
def close_tab():           return _send_hotkey_safe('ctrl', 'w')
def refresh():             return _send_hotkey_safe('f5')
def zoom_in():             return _send_hotkey_safe('ctrl', '+')
def zoom_out():            return _send_hotkey_safe('ctrl', '-')
def bookmark_tab():        return _send_hotkey_safe('ctrl', 'd')
def open_incognito():      return _send_hotkey_safe('ctrl', 'shift', 'n')

def switch_tab(n):
    """Jump to tab 1-9."""
    if 1 <= n <= 9:
        return _send_hotkey_safe('ctrl', str(n))
    return False

def search(query):
    import webbrowser
    try:
        if not _is_browser_active():
            # No browser active, open a new one with Google search
            search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(search_url)
            return True
        else:
            # Browser is active, use Ctrl+L to focus address bar
            pyautogui.hotkey('ctrl', 'l')
            time.sleep(0.05)
            pyautogui.write(query)
            pyautogui.press('enter')
            return True
    except Exception:
        return False

def clear_browsing_data():
    """Open Ctrl+Shift+Delete dialog and accept defaults."""
    try:
        if not _is_browser_active():
            return False
        pyautogui.hotkey('ctrl', 'shift', 'delete')
        time.sleep(0.4)
        pyautogui.press('enter')
        return True
    except Exception:
        return False