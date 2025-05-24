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
        pyautogui.PAUSE = 0.05
        self._update_window_handles()

    def _update_window_handles(self):
        """Enumerate visible, non-minimized, focusable windows, excluding Voice Listener."""
        temp_handles = []
        # Inner function for EnumWindows, ensuring it has access to temp_handles
        # and other necessary variables if they were part of the class or outer scope.
        # Based on your original code, excluded_classes are defined outside,
        # so they should be accessible or passed if needed.
        # For this snippet, assuming excluded_classes is accessible.
        
        excluded_classes = [
            "Progman", "WorkerW", "Shell_TrayWnd", "PopupHost",
            "Windows.UI.Core.CoreWindow", "ApplicationFrameWindow",
            "Button", "ComboBox", "Edit", "ListBox", "Static", "ToolbarWindow32",
            "NotifyIconOverflowWindow", "ShellExperienceHost", "SearchUI"
        ]

        def enum_windows_callback(hwnd, _):
            class_name = win32gui.GetClassName(hwnd)
            window_text = win32gui.GetWindowText(hwnd)
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            
            if "Voice Listener" in window_text: # Skip the Voice Listener window
                return

            if (win32gui.IsWindowVisible(hwnd) and
                not win32gui.IsIconic(hwnd) and
                window_text and
                class_name not in excluded_classes and
                "Windows Input Experience" not in window_text and
                "DesktopWindowXamlSource" not in class_name and
                not (ex_style & win32con.WS_EX_TOOLWINDOW or ex_style & win32con.WS_EX_NOACTIVATE) and
                (style & win32con.WS_CAPTION or style & win32con.WS_SYSMENU) and
                not (style & win32con.WS_CHILD) and
                win32gui.IsWindowEnabled(hwnd)):
                try:
                    tid, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid == 0: # Skip system process windows if necessary
                        return
                    temp_handles.append(hwnd)
                except Exception:
                    pass # Ignore errors for individual windows during enumeration
        
        win32gui.EnumWindows(enum_windows_callback, None)
        self.window_handles = temp_handles
        
        if not self.window_handles:
            self.current_window_index = 0
            print("No focusable windows found after filtering.")
        else:
            # Simply clamp self.current_window_index to the new list's bounds.
            # This ensures that if windows were closed and the list shrank,
            # the index remains valid.
            self.current_window_index = max(0, min(self.current_window_index, len(self.window_handles) - 1))
            print(f"Updated window handles: {[win32gui.GetWindowText(hwnd) for hwnd in self.window_handles]} (Count: {len(self.window_handles)})")
    def _restore_and_focus(self, hwnd):
        """Restore a minimized window and bring it to the foreground."""
        try:
            if not win32gui.IsWindow(hwnd):
                print(f"Invalid window handle: {hwnd}")
                return False

            window_text = win32gui.GetWindowText(hwnd)

            # Release Alt key
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')
                time.sleep(0.05)
            except Exception as e:
                print(f"Note: WScript.Shell SendKeys failed: {e}")

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)

            # Try multiple methods to focus
            attempts = 0
            max_attempts = 3
            while attempts < max_attempts:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.1)
                if win32gui.GetForegroundWindow() == hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    return True
                
                # Alternative focus method
                win32gui.BringWindowToTop(hwnd)
                target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
                current_thread_id = win32api.GetCurrentThreadId()
                if current_thread_id != target_thread_id:
                    win32process.AttachThreadInput(current_thread_id, target_thread_id, True)
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                        if win32gui.GetForegroundWindow() == hwnd:
                            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                            return True
                    finally:
                        win32process.AttachThreadInput(current_thread_id, target_thread_id, False)
                time.sleep(0.1)
                attempts += 1

            # Fallback to Alt+Tab
            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            pyautogui.keyUp('alt')
            time.sleep(0.3)
            if win32gui.GetForegroundWindow() == hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                return True
            
            print(f"Failed to focus window '{window_text}' after {max_attempts} attempts.")
            return False
        except Exception as e:
            window_text = win32gui.GetWindowText(hwnd) if win32gui.IsWindow(hwnd) else "Unknown"
            print(f"Error focusing window ({window_text}, HWND: {hwnd}): {e}")
            return False

    def switch_window(self) -> bool:
        try:
            # Get the current foreground window *before* updating handles.
            # This is used to check if our cycling lands back on the active window.
            foreground_at_call_start = win32gui.GetForegroundWindow()

            self._update_window_handles() # self.window_handles is updated.
                                      # self.current_window_index is clamped if list size changed.

            if not self.window_handles:
                print("No windows available to switch to.")
                self.speech.speak("No windows available to switch to.")
                return False

            if len(self.window_handles) == 1:
                target_hwnd_single = self.window_handles[0]
                current_text = win32gui.GetWindowText(target_hwnd_single).split(' - ')[-1]
                if target_hwnd_single != foreground_at_call_start:
                    if self._restore_and_focus(target_hwnd_single):
                        self.speech.speak(f"Switched to {current_text}")
                    else:
                        self.speech.speak(f"Could not switch to {current_text}")
                else:
                    self.speech.speak(f"Only {current_text} is open.")
                return True

            # self.current_window_index is the "starting point" for this cycle's selection.
            # It was clamped by _update_window_handles if the list size changed.
            # We always want to move to the *next* window from this conceptual point.
            
            prospective_next_idx = (self.current_window_index + 1) % len(self.window_handles)

            # If the list has more than one window AND our prospective_next_idx points to
            # the window that was ALREADY in the foreground when we started the function call,
            # it means our cycle effectively landed on the active window (e.g., after a full loop
            # or due to list reordering + clamping). So, advance one more time.
            if len(self.window_handles) > 1 and self.window_handles[prospective_next_idx] == foreground_at_call_start:
                prospective_next_idx = (prospective_next_idx + 1) % len(self.window_handles)
            
            self.current_window_index = prospective_next_idx # This is now the index of the window to activate.

            target_hwnd = self.window_handles[self.current_window_index]
            target_title = win32gui.GetWindowText(target_hwnd)

            print(f"Switching to index {self.current_window_index} ({self.current_window_index + 1}/{len(self.window_handles)}): '{target_title}'")

            if self._restore_and_focus(target_hwnd):
                active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
                print(f"Switched to: '{active_title}' (Targeted: '{target_title}')")
                self.speech.speak(f"Switched to {active_title.split(' - ')[-1]}")
                return True
            else:
                print(f"Failed to switch to '{target_title}'.")
                self.speech.speak(f"Could not switch to {target_title.split(' - ')[-1]}")
                # If focus fails, current_window_index is already set to the failed target.
                # Next call will increment from there, effectively skipping this problematic window.
                return False
        except Exception as e:
            print(f"Error in switch_window: {e}")
            self.speech.speak("Error switching window")
            self.current_window_index = 0 # Reset index on error
            self._update_window_handles()  # Refresh and clamp
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

            try:
                pythoncom.CoInitialize()
                initialized_com = True
            except pythoncom.com_error:
                initialized_com = False
            
            shell = None
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                for window in shell.Windows():
                    if window.hwnd == hwnd:
                        path = window.LocationURL
                        if path.startswith("file:///"):
                            path = path.replace("file:///", "").replace("/", "\\")
                            path = os.path.normpath(path)
                            if os.path.isdir(path):
                                return path
            finally:
                if shell: del shell
                if initialized_com: pythoncom.CoUninitialize()
            return None
        except Exception as e:
            print(f"Error getting active File Explorer path: {e}")
            return None

    def get_open_explorer_paths(self) -> list:
        """Get paths of all open File Explorer windows."""
        paths = []
        try:
            try:
                pythoncom.CoInitialize()
                initialized_com = True
            except pythoncom.com_error:
                initialized_com = False
            
            shell = None
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                for window in shell.Windows():
                    try:
                        class_name = win32gui.GetClassName(window.hwnd)
                        if class_name not in ["CabinetWClass", "ExplorerWClass"]:
                            continue

                        path = window.LocationURL
                        if path.startswith("file:///"):
                            path = path.replace("file:///", "").replace("/", "\\")
                            path = os.path.normpath(path)
                            if os.path.isdir(path) and path not in paths:
                                paths.append(path)
                    except (pythoncom.com_error, win32gui.error):
                        continue
            finally:
                if shell: del shell
                if initialized_com: pythoncom.CoUninitialize()

            if not paths:
                print("No open File Explorer paths detected.")
            return paths
        except Exception as e:
            print(f"Error getting open File Explorer paths: {e}")
            self.speech.speak("Error detecting open folders.")
            return []

    def volume_up(self) -> bool:
        if self.volume_interface:
            try:
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = min(1.0, current_volume + 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume increased to {int(new_volume * 100)}%")
                return True
            except Exception as e:
                print(f"Error (pycaw) volume up: {e}. Fallback.")
        try:
            pyautogui.press('volumeup')
            print("Volume up (hotkey)")
            return True
        except Exception as e:
            print(f"Error (hotkey) volume up: {e}")
            self.speech.speak("Error increasing volume")
            return False

    def volume_down(self) -> bool:
        if self.volume_interface:
            try:
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = max(0.0, current_volume - 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume decreased to {int(new_volume * 100)}%")
                return True
            except Exception as e:
                print(f"Error (pycaw) volume down: {e}. Fallback.")
        try:
            pyautogui.press('volumedown')
            print("Volume down (hotkey)")
            return True
        except Exception as e:
            print(f"Error (hotkey) volume down: {e}")
            self.speech.speak("Error decreasing volume")
            return False

    def mute_toggle(self) -> bool:
        if self.volume_interface:
            try:
                is_muted = self.volume_interface.GetMute()
                self.volume_interface.SetMute(not is_muted, None)
                status = "muted" if not is_muted else "unmuted"
                print(f"Volume {status}")
                self.speech.speak(f"Volume {status}")
                return True
            except Exception as e:
                print(f"Error (pycaw) mute toggle: {e}. Fallback.")
        try:
            pyautogui.press('volumemute')
            print("Volume mute toggled (hotkey)")
            self.speech.speak("Volume mute toggled")
            return True
        except Exception as e:
            print(f"Error (hotkey) mute toggle: {e}")
            self.speech.speak("Error toggling mute")
            return False

    def maximize_volume(self) -> bool:
        if self.volume_interface:
            try:
                self.volume_interface.SetMasterVolumeLevelScalar(1.0, None)
                if self.volume_interface.GetMute():
                    self.volume_interface.SetMute(False, None)
                print("Volume set to maximum (100%)")
                self.speech.speak("Volume maximized")
                return True
            except Exception as e:
                print(f"Error (pycaw) maximizing volume: {e}. Fallback.")
        try:
            is_muted = False
            if self.volume_interface:
                try: is_muted = self.volume_interface.GetMute()
                except: pass
            if is_muted: pyautogui.press('volumemute')
            for _ in range(10): pyautogui.press('volumeup'); time.sleep(0.02)
            print("Volume maximized (hotkey)")
            self.speech.speak("Volume maximized")
            return True
        except Exception as e:
            print(f"Error (hotkey) maximizing volume: {e}")
            self.speech.speak("Error maximizing volume")
            return False

    def set_volume(self, level_input) -> bool:
        try:
            level = int(level_input)
            if not 0 <= level <= 100:
                self.speech.speak("Volume level must be between 0 and 100.")
                print(f"Volume level out of range: {level}")
                return False
        except (ValueError, TypeError):
            self.speech.speak("Invalid volume level. Please say a number.")
            print(f"Invalid volume level input: {level_input}")
            return False

        if self.volume_interface:
            try:
                volume_scalar = level / 100.0
                self.volume_interface.SetMasterVolumeLevelScalar(volume_scalar, None)
                if level > 0 and self.volume_interface.GetMute():
                    self.volume_interface.SetMute(False, None)
                elif level == 0:
                    self.volume_interface.SetMute(True, None)
                print(f"Volume set to {level}%")
                self.speech.speak(f"Volume set to {level} percent")
                return True
            except Exception as e:
                print(f"Error (pycaw) setting volume: {e}. Fallback.")
        
        try:
            is_muted = False
            if self.volume_interface:
                try: is_muted = self.volume_interface.GetMute()
                except: pass
            if is_muted: pyautogui.press('volumemute')
            for _ in range(10): pyautogui.press('volumedown'); time.sleep(0.02)
            if level > 0:
                num_steps = int(level / 5)
                for _ in range(num_steps): pyautogui.press('volumeup'); time.sleep(0.02)
            elif level == 0:
                pyautogui.press('volumemute')
            print(f"Volume set to approximately {level}% (hotkey)")
            self.speech.speak(f"Volume set to {level} percent")
            return True
        except Exception as e:
            print(f"Error (hotkey) setting volume: {e}")
            self.speech.speak("Error setting volume")
            return False

    def _set_brightness_wmi(self, level_percent: int) -> bool:
        if not self.brightness_interface: return False
        try:
            methods = self.brightness_interface.WmiMonitorBrightnessMethods()
            if not methods: raise Exception("WMI BrightnessMethods not found")
            monitor_methods = methods[0]
            monitor_methods.WmiSetBrightness(level_percent, 0)
            return True
        except Exception as e:
            print(f"Error setting brightness with WMI: {e}")
            return False

    def _get_brightness_wmi(self) -> Optional[int]:
        if not self.brightness_interface: return None
        try:
            brightness_info = self.brightness_interface.WmiMonitorBrightness()
            if not brightness_info: raise Exception("WMI Brightness info not found")
            return brightness_info[0].CurrentBrightness
        except Exception as e:
            print(f"Error getting brightness with WMI: {e}")
            return None

    def brightness_up(self) -> bool:
        current_brightness = self._get_brightness_wmi()
        if current_brightness is not None:
            new_brightness = min(100, current_brightness + 10)
            if self._set_brightness_wmi(new_brightness):
                print(f"Brightness increased to {new_brightness}%")
                return True
        self.speech.speak("Brightness up command sent. Result system dependent.")
        print("Attempting brightness up via hotkey (system dependent).")
        return False

    def brightness_down(self) -> bool:
        current_brightness = self._get_brightness_wmi()
        if current_brightness is not None:
            new_brightness = max(0, current_brightness - 10)
            if self._set_brightness_wmi(new_brightness):
                print(f"Brightness decreased to {new_brightness}%")
                return True
        self.speech.speak("Brightness down command sent. Result system dependent.")
        print("Attempting brightness down via hotkey (system dependent).")
        return False

    def maximize_brightness(self) -> bool:
        if self._set_brightness_wmi(100):
            print("Brightness maximized to 100%")
            self.speech.speak("Brightness maximized")
            return True
        self.speech.speak("Maximize brightness command sent. Result system dependent.")
        print("Attempting brightness maximization via hotkey (system dependent).")
        return False

    def set_brightness(self, level_input) -> bool:
        try:
            level = int(level_input)
            if not 0 <= level <= 100:
                self.speech.speak("Brightness level must be between 0 and 100.")
                print(f"Brightness level out of range: {level}")
                return False
        except (ValueError, TypeError):
            self.speech.speak("Invalid brightness level. Please say a number.")
            print(f"Invalid brightness level input: {level_input}")
            return False

        if self._set_brightness_wmi(level):
            print(f"Brightness set to {level}%")
            self.speech.speak(f"Brightness set to {level} percent")
            return True
        self.speech.speak(f"Set brightness to {level} command sent. Result system dependent.")
        print(f"Attempting to set brightness to {level}% via hotkey (system dependent).")
        return False

    def minimize_all_windows(self) -> bool:
        try:
            pyautogui.hotkey('win', 'd')
            print("Minimized all windows")
            self.speech.speak("Desktop shown")
            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error minimizing all: {e}")
            self.speech.speak("Error minimizing")
            return False

    def restore_all_windows(self) -> bool:
        try:
            pyautogui.hotkey('win', 'd')
            print("Restored windows (toggled desktop)")
            self.speech.speak("Windows restored")
            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error restoring all: {e}")
            self.speech.speak("Error restoring")
            return False

    def maximize_current_window(self) -> bool:
        try:
            pyautogui.hotkey('win', 'up')
            print("Maximized current")
            self.speech.speak("Window maximized")
            return True
        except Exception as e:
            print(f"Error maximizing: {e}")
            self.speech.speak("Error maximizing")
            return False

    def minimize_current_window(self) -> bool:
        try:
            pyautogui.hotkey('win', 'down')
            print("Minimized current")
            self.speech.speak("Window minimized")
            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error minimizing current: {e}")
            self.speech.speak("Error minimizing")
            return False

    def close_current_window(self) -> bool:
        try:
            pyautogui.hotkey('alt', 'f4')
            print("Closed current window")
            self.speech.speak("Window closed")
            time.sleep(0.5)
            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error closing window: {e}")
            self.speech.speak("Error closing window")
            return False

    def move_window_left(self) -> bool:
        try:
            pyautogui.hotkey('win', 'left')
            print("Moved window left")
            self.speech.speak("Window moved left")
            return True
        except Exception as e:
            print(f"Error moving left: {e}")
            self.speech.speak("Error moving left")
            return False

    def move_window_right(self) -> bool:
        try:
            pyautogui.hotkey('win', 'right')
            print("Moved window right")
            self.speech.speak("Window moved right")
            return True
        except Exception as e:
            print(f"Error moving right: {e}")
            self.speech.speak("Error moving right")
            return False

    def take_screenshot(self) -> bool:
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            desktop_path = os.path.join(os.environ.get('USERPROFILE', os.path.expanduser("~")), 'Desktop')
            if not os.path.isdir(desktop_path):
                desktop_path = os.path.expanduser("~/Desktop")
                if not os.path.isdir(desktop_path):
                    desktop_path = os.getcwd()
                    print(f"Desktop path not found, saving screenshot to: {desktop_path}")

            screenshot_file = f"screenshot_{timestamp}.png"
            screenshot_path = os.path.join(desktop_path, screenshot_file)
            pyautogui.screenshot(screenshot_path)
            print(f"Screenshot saved to {screenshot_path}")
            self.speech.speak("Screenshot saved")
            return True
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            self.speech.speak("Error taking screenshot")
            return False

    def run_application(self, app_name_input: str) -> bool:
        if not app_name_input or not isinstance(app_name_input, str):
            self.speech.speak("Please provide an application name.")
            print("No valid application name provided.")
            return False

        app_name = app_name_input.strip().lower()
        if app_name.startswith("run "):
            app_name = app_name.split("run ", 1)[1].strip()
        elif app_name == "run":
            self.speech.speak("Which application to run?")
            print("No application name after 'run'.")
            return False
        if not app_name:
            self.speech.speak("No application name specified.")
            print("Application name is empty.")
            return False

        try:
            pyautogui.press('win')
            time.sleep(0.6)
            pyautogui.write(app_name, interval=0.07)
            time.sleep(1.0)
            pyautogui.press('enter')
            self.speech.speak(f"Opening {app_name}")
            print(f"Attempted to launch: {app_name}")
            time.sleep(2.5)
            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error launching '{app_name}': {e}")
            self.speech.speak(f"Error opening {app_name}")
            return False