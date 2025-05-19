import platform
from typing import Optional
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pythoncom
import wmi
import pyautogui
import time
import os
from datetime import datetime
import win32gui
import win32con
import win32process
import win32api
import win32com.client

class OSManagement:
    def __init__(self, speech):
        self.speech = speech
        self.system = platform.system()
        self.volume_interface = None
        self.brightness_interface = None
        self.window_handles = []
        self.current_window_index = 0

        if self.system != "Windows":
            print("This application only supports Windows.")
            self.speech.speak("This application only supports Windows.")
            return

        try:
            pythoncom.CoInitialize()
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume_interface = interface.QueryInterface(IAudioEndpointVolume)
        except Exception as e:
            print(f"Error initializing Windows audio interface: {e}")
            self.speech.speak("Error initializing audio control. Using hotkey fallback.")

        try:
            self.brightness_interface = wmi.WMI(namespace="wmi")
        except Exception as e:
            print(f"Error initializing Windows brightness interface: {e}")
            self.speech.speak("Error initializing brightness control. Using hotkey fallback.")
            self.brightness_interface = None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1

        self._update_window_handles()

    def _update_window_handles(self):
        """Enumerate all visible windows and store their handles, excluding the desktop."""
        self.window_handles = []
        def enum_windows_callback(hwnd, _):
            class_name = win32gui.GetClassName(hwnd)
            window_text = win32gui.GetWindowText(hwnd)
            if (win32gui.IsWindowVisible(hwnd) and window_text and 
                class_name not in ["Progman", "WorkerW"] and "Desktop" not in window_text):
                self.window_handles.append(hwnd)
        win32gui.EnumWindows(enum_windows_callback, None)
        if self.current_window_index >= len(self.window_handles):
            self.current_window_index = 0

    def _restore_and_focus(self, hwnd):
        """Restore a minimized window and force focus on it."""
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)

            target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
            current_thread_id = win32api.GetCurrentThreadId()

            if target_thread_id != current_thread_id:
                win32process.AttachThreadInput(current_thread_id, target_thread_id, True)

            win32gui.SetForegroundWindow(hwnd)

            if target_thread_id != current_thread_id:
                win32process.AttachThreadInput(current_thread_id, target_thread_id, False)

            return True
        except Exception as e:
            print(f"Error restoring and focusing window: {e}")
            return False

    def get_active_explorer_path(self) -> Optional[str]:
        """Get the path of the active File Explorer window."""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return None

            class_name = win32gui.GetClassName(hwnd)
            if class_name not in ["CabinetWClass", "ExplorerWClass"]:
                return None

            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                if window.hwnd == hwnd:
                    path = window.LocationURL
                    if path:
                        path = path.replace("file:///", "").replace("/", "\\")
                        path = os.path.normpath(path)
                        if os.path.isdir(path):
                            return path
            return None
        except Exception as e:
            print(f"Error getting active File Explorer path: {e}")
            self.speech.speak("Error detecting active folder.")
            return None

    def get_open_explorer_paths(self) -> list:
        """Get paths of all open File Explorer windows (active or minimized)."""
        paths = []
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                try:
                    path = window.LocationURL
                    if path:
                        path = path.replace("file:///", "").replace("/", "\\")
                        path = os.path.normpath(path)
                        if os.path.isdir(path) and path not in paths:
                            paths.append(path)
                except:
                    continue
            return paths
        except Exception as e:
            print(f"Error getting open File Explorer paths: {e}")
            self.speech.speak("Error detecting open folders.")
            return []

    def volume_up(self) -> bool:
        """Increase the system volume by 10%."""
        if self.volume_interface:
            try:
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = min(1.0, current_volume + 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume increased to {int(new_volume * 100)}%")
                self.speech.speak("Volume increased")
                return True
            except Exception as e:
                print(f"Error increasing volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        try:
            pyautogui.press('volumeup')
            print("Volume increased via hotkey")
            self.speech.speak("Volume increased")
            return True
        except Exception as e:
            print(f"Error with hotkey volume up: {e}")
            self.speech.speak("Error increasing volume")
            return False

    def volume_down(self) -> bool:
        """Decrease the system volume by 10%."""
        if self.volume_interface:
            try:
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = max(0.0, current_volume - 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume decreased to {int(new_volume * 100)}%")
                self.speech.speak("Volume decreased")
                return True
            except Exception as e:
                print(f"Error decreasing volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        try:
            pyautogui.press('volumedown')
            print("Volume decreased via hotkey")
            self.speech.speak("Volume decreased")
            return True
        except Exception as e:
            print(f"Error with hotkey volume down: {e}")
            self.speech.speak("Error decreasing volume")
            return False

    def mute_toggle(self) -> bool:
        """Toggle mute/unmute for the system volume."""
        if self.volume_interface:
            try:
                is_muted = self.volume_interface.GetMute()
                self.volume_interface.SetMute(not is_muted, None)
                status = "muted" if not is_muted else "unmuted"
                print(f"Volume {status}")
                self.speech.speak(f"Volume {status}")
                return True
            except Exception as e:
                print(f"Error toggling mute with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with mute control. Trying hotkey.")
        
        try:
            pyautogui.press('volumemute')
            status = "toggled"
            print(f"Volume mute {status} via hotkey")
            self.speech.speak("Volume mute toggled")
            return True
        except Exception as e:
            print(f"Error with hotkey mute toggle: {e}")
            self.speech.speak("Error toggling mute")
            return False

    def maximize_volume(self) -> bool:
        """Set the system volume to maximum (100%)."""
        if self.volume_interface:
            try:
                self.volume_interface.SetMasterVolumeLevelScalar(1.0, None)
                print("Volume set to maximum (100%)")
                self.speech.speak("Volume maximized")
                return True
            except Exception as e:
                print(f"Error maximizing volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        try:
            for _ in range(10):
                pyautogui.press('volumeup')
                time.sleep(0.05)
            print("Volume maximized via hotkey")
            self.speech.speak("Volume maximized")
            return True
        except Exception as e:
            print(f"Error with hotkey maximize volume: {e}")
            self.speech.speak("Error maximizing volume")
            return False

    def set_volume(self, level: int) -> bool:
        """Set the system volume to a specific level (0–100)."""
        if not 0 <= level <= 100:
            print(f"Volume level must be between 0 and 100, got {level}")
            self.speech.speak("Volume level must be between 0 and 100")
            return False

        if self.volume_interface:
            try:
                volume_scalar = level / 100.0
                self.volume_interface.SetMasterVolumeLevelScalar(volume_scalar, None)
                print(f"Volume set to {level}%")
                self.speech.speak(f"Volume set to {level} percent")
                return True
            except Exception as e:
                print(f"Error setting volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        try:
            pyautogui.press('volumemute')
            for _ in range(10):
                pyautogui.press('volumedown')
                time.sleep(0.05)
            pyautogui.press('volumemute')
            presses = int(level // 5)
            for _ in range(presses):
                pyautogui.press('volumeup')
                time.sleep(0.05)
            print(f"Volume set to approximately {level}% via hotkey")
            self.speech.speak(f"Volume set to {level} percent")
            return True
        except Exception as e:
            print(f"Error with hotkey set volume: {e}")
            self.speech.speak("Error setting volume")
            return False

    def brightness_up(self) -> bool:
        """Increase the screen brightness by 10%."""
        if self.brightness_interface:
            try:
                monitor = self.brightness_interface.WmiMonitorBrightnessMethods()[0]
                current_brightness = self.brightness_interface.WmiMonitorBrightness()[0].CurrentBrightness
                new_brightness = min(100, current_brightness + 10)
                monitor.WmiSetBrightness(new_brightness, 0)
                print(f"Brightness increased to {new_brightness}%")
                self.speech.speak("Brightness increased")
                return True
            except Exception as e:
                print(f"Error increasing brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        try:
            pyautogui.hotkey('fn', 'f6')
            print("Brightness increased via hotkey")
            self.speech.speak("Brightness increased")
            return True
        except Exception as e:
            print(f"Error with hotkey brightness up: {e}. Check if Fn+F6 is correct for your device.")
            self.speech.speak("Error increasing brightness. Check your brightness hotkey.")
            return False

    def brightness_down(self) -> bool:
        """Decrease the screen brightness by 10%."""
        if self.brightness_interface:
            try:
                monitor = self.brightness_interface.WmiMonitorBrightnessMethods()[0]
                current_brightness = self.brightness_interface.WmiMonitorBrightness()[0].CurrentBrightness
                new_brightness = max(0, current_brightness - 10)
                monitor.WmiSetBrightness(new_brightness, 0)
                print(f"Brightness decreased to {new_brightness}%")
                self.speech.speak("Brightness decreased")
                return True
            except Exception as e:
                print(f"Error decreasing brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        try:
            pyautogui.hotkey('fn', 'f5')
            print("Brightness decreased via hotkey")
            self.speech.speak("Brightness decreased")
            return True
        except Exception as e:
            print(f"Error with hotkey brightness down: {e}. Check if Fn+F5 is correct for your device.")
            self.speech.speak("Error decreasing brightness. Check your brightness hotkey.")
            return False

    def maximize_brightness(self) -> bool:
        """Set the screen brightness to maximum (100%)."""
        if self.brightness_interface:
            try:
                monitor = self.brightness_interface.WmiMonitorBrightnessMethods()[0]
                monitor.WmiSetBrightness(100, 0)
                print("Brightness set to maximum (100%)")
                self.speech.speak("Brightness maximized")
                return True
            except Exception as e:
                print(f"Error maximizing brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        try:
            for _ in range(10):
                pyautogui.hotkey('fn', 'f6')
                time.sleep(0.05)
            print("Brightness maximized via hotkey")
            self.speech.speak("Brightness maximized")
            return True
        except Exception as e:
            print(f"Error with hotkey maximize brightness: {e}. Check if Fn+F6 is correct for your device.")
            self.speech.speak("Error maximizing brightness. Check your brightness hotkey.")
            return False

    def set_brightness(self, level: int) -> bool:
        """Set the screen brightness to a specific level (0–100)."""
        if not 0 <= level <= 100:
            print(f"Brightness level must be between 0 and 100, got {level}")
            self.speech.speak("Brightness level must be between 0 and 100")
            return False

        if self.brightness_interface:
            try:
                monitor = self.brightness_interface.WmiMonitorBrightnessMethods()[0]
                monitor.WmiSetBrightness(level, 0)
                print(f"Brightness set to {level}%")
                self.speech.speak(f"Brightness set to {level} percent")
                return True
            except Exception as e:
                print(f"Error setting brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        try:
            for _ in range(10):
                pyautogui.hotkey('fn', 'f5')
                time.sleep(0.05)
            presses = int(level // 5)
            for _ in range(presses):
                pyautogui.hotkey('fn', 'f6')
                time.sleep(0.05)
            print(f"Brightness set to approximately {level}% via hotkey")
            self.speech.speak(f"Brightness set to {level} percent")
            return True
        except Exception as e:
            print(f"Error with hotkey set brightness: {e}. Check if Fn+F5/F6 are correct for your device.")
            self.speech.speak("Error setting brightness. Check your brightness hotkey.")
            return False

    def switch_window(self) -> bool:
        """Switch to the next open window using Alt+Tab simulation."""
        try:
            # Update the list of window handles to ensure it's current
            self._update_window_handles()

            if not self.window_handles:
                print("No windows available to switch to.")
                self.speech.speak("No windows available to switch to.")
                return False

            # Increment the window index to select the next window
            self.current_window_index = (self.current_window_index + 1) % len(self.window_handles)

            # Simulate Alt+Tab to cycle to the desired window
            pyautogui.keyUp('alt')  # Ensure Alt is not stuck
            pyautogui.keyUp('tab')  # Ensure Tab is not stuck
            time.sleep(0.1)  # Small delay to ensure keys are released

            pyautogui.keyDown('alt')
            time.sleep(0.2)  # Hold Alt briefly to open the window switcher

            # Press Tab the required number of times to reach the target window
            for _ in range(self.current_window_index):
                pyautogui.press('tab')
                time.sleep(0.1)  # Small delay between Tab presses for reliability

            pyautogui.keyUp('alt')  # Release Alt to select the window
            time.sleep(0.5)  # Allow the window switch to complete

            # Verify the active window
            active_hwnd = win32gui.GetForegroundWindow()
            target_hwnd = self.window_handles[self.current_window_index]
            active_title = win32gui.GetWindowText(active_hwnd)

            if active_hwnd == target_hwnd:
                print(f"Switched to window: {active_title} (Index: {self.current_window_index + 1}/{len(self.window_handles)})")
                self.speech.speak("Window switched")
                return True
            else:
                print(f"Failed to switch to expected window. Current window: {active_title}. Retrying with direct focus.")
                self.speech.speak("Error switching window. Retrying.")

                # Fallback: Try direct focus if Alt+Tab didn't work
                win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)
                if self._restore_and_focus(target_hwnd):
                    window_title = win32gui.GetWindowText(target_hwnd)
                    print(f"Retry successful. Switched to window: {window_title} (Index: {self.current_window_index + 1}/{len(self.window_handles)})")
                    self.speech.speak("Window switched")
                    return True
                else:
                    print("Retry failed. Could not switch to the target window.")
                    self.speech.speak("Error switching window.")
                    return False

        except Exception as e:
            print(f"Error switching window: {e}")
            self.speech.speak("Error switching window")
            return False

    def minimize_all_windows(self) -> bool:
        """Minimize all windows to show the desktop."""
        try:
            pyautogui.hotkey('win', 'd')
            print("All windows minimized")
            self.speech.speak("All windows minimized")
            return True
        except Exception as e:
            print(f"Error minimizing all windows: {e}")
            self.speech.speak("Error minimizing windows")
            return False

    def restore_all_windows(self) -> bool:
        """Restore all minimized windows."""
        try:
            pyautogui.hotkey('win', 'shift', 'm')
            print("All windows restored")
            self.speech.speak("Windows restored")
            return True
        except Exception as e:
            print(f"Error restoring windows: {e}")
            self.speech.speak("Error restoring windows")
            return False

    def maximize_current_window(self) -> bool:
        """Maximize the current window."""
        try:
            pyautogui.hotkey('win', 'up')
            print("Current window maximized")
            self.speech.speak("Window maximized")
            return True
        except Exception as e:
            print(f"Error maximizing window: {e}")
            self.speech.speak("Error maximizing window")
            return False

    def minimize_current_window(self) -> bool:
        """Minimize the current window."""
        try:
            pyautogui.hotkey('win', 'down')
            print("Current window minimized")
            self.speech.speak("Window minimized")
            return True
        except Exception as e:
            print(f"Error minimizing window: {e}")
            self.speech.speak("Error minimizing window")
            return False

    def close_current_window(self) -> bool:
        """Close the current window or application."""
        try:
            pyautogui.hotkey('alt', 'f4')
            print("Current window closed")
            self.speech.speak("Window closed")
            return True
        except Exception as e:
            print(f"Error closing window: {e}")
            self.speech.speak("Error closing window")
            return False

    def move_window_left(self) -> bool:
        """Move the current window to the left half of the screen."""
        try:
            pyautogui.hotkey('win', 'left')
            print("Window moved to left")
            self.speech.speak("Window moved left")
            return True
        except Exception as e:
            print(f"Error moving window left: {e}")
            self.speech.speak("Error moving window left")
            return False

    def move_window_right(self) -> bool:
        """Move the current window to the right half of the screen."""
        try:
            pyautogui.hotkey('win', 'right')
            print("Window moved to right")
            self.speech.speak("Window moved right")
            return True
        except Exception as e:
            print(f"Error moving window right: {e}")
            self.speech.speak("Error moving window right")
            return False

    def take_screenshot(self) -> bool:
        """Take a screenshot and save it to the Desktop."""
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            desktop_path = os.path.expanduser("~/Desktop")
            screenshot_path = os.path.join(desktop_path, f"screenshot_{timestamp}.png")
            
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            print(f"Screenshot saved to {screenshot_path}")
            self.speech.speak("Screenshot taken and saved to Desktop")
            return True
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            self.speech.speak("Error taking screenshot")
            return False

    def run_application(self, app_name: str) -> bool:
        """Launch an application by searching in the Windows Start menu."""
        if not app_name:
            print("No application name provided.")
            self.speech.speak("Please provide an application name.")
            return False

        try:
            # Clean the app_name to remove 'run' prefix and unwanted characters
            app_name = app_name.strip().lower()
            if app_name.startswith("run "):
                app_name = app_name[4:].strip()  # Remove 'run ' prefix
            elif app_name == "run":
                print("No application name provided after 'run'.")
                self.speech.speak("Please provide an application name after run.")
                return False

            if not app_name:
                print("Application name is empty after cleaning.")
                self.speech.speak("No valid application name provided.")
                return False

            # Simulate Windows key press to open Start menu
            pyautogui.press('win')
            time.sleep(0.5)  # Wait for Start menu to open

            # Type the application name
            pyautogui.write(app_name, interval=0.05)
            time.sleep(1)  # Wait for search results to populate

            # Press Enter to launch the first matching application
            pyautogui.press('enter')
            time.sleep(0.5)  # Allow time for the application to start

            print(f"Attempted to launch application: {app_name} via Start menu search")
            self.speech.speak(f"Opening {app_name}")
            return True

        except Exception as e:
            print(f"Error launching application '{app_name}': {e}")
            self.speech.speak(f"Error opening {app_name}. Please try again.")
            return False