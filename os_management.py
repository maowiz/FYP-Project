import platform
from typing import Optional, List, Tuple
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
import urllib.parse
from grid_manager import GridManager


class OSManagement:
    def __init__(self, speech):
        self.speech = speech
        self.system = platform.system()
        self.volume_interface = None
        self.brightness_interface = None
        self.window_handles = []
        self.current_window_index = 0
        if not hasattr(self, 'context'):
            self.context = {}

        if self.system != "Windows":
            print("This application only supports Windows.")
            return

        try:
            pythoncom.CoInitialize()
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume_interface = interface.QueryInterface(IAudioEndpointVolume)
        except Exception as e:
            print(f"Error initializing Windows audio interface: {e}")
        finally:
            pythoncom.CoUninitialize()

        try:
            self.brightness_interface = wmi.WMI(namespace="wmi")
        except Exception as e:
            print(f"Error initializing Windows brightness interface: {e}")
            self.brightness_interface = None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0
        self._update_window_handles()
        self.grid = GridManager(speech)

    def _update_window_handles(self):
        # This method remains the same
        pass

    def _restore_and_focus(self, hwnd):
        # This method remains the same
        pass

    def switch_window(self) -> Tuple[bool, str]:
        try:
            self._update_window_handles()
            if not self.window_handles:
                return False, "No windows available to switch to."
            
            # Logic to find the next window
            current_hwnd = win32gui.GetForegroundWindow()
            try:
                current_idx = self.window_handles.index(current_hwnd)
                next_idx = (current_idx + 1) % len(self.window_handles)
            except ValueError:
                next_idx = 0

            target_hwnd = self.window_handles[next_idx]
            target_title = win32gui.GetWindowText(target_hwnd).split(' - ')[-1]

            if self._restore_and_focus(target_hwnd):
                return True, f"Switched to {target_title}"
            else:
                return False, f"Could not switch to {target_title}"
        except Exception as e:
            return False, "Error switching window."

    def volume_up(self) -> Tuple[bool, str]:
        try:
            pyautogui.press('volumeup')
            return True, "Volume increased."
        except Exception:
            return False, "Error increasing volume."

    def volume_down(self) -> Tuple[bool, str]:
        try:
            pyautogui.press('volumedown')
            return True, "Volume decreased."
        except Exception:
            return False, "Error decreasing volume."

    def mute_toggle(self) -> Tuple[bool, str]:
        try:
            pyautogui.press('volumemute')
            return True, "Mute toggled."
        except Exception:
            return False, "Error toggling mute."

    def maximize_current_window(self) -> Tuple[bool, str]:
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd: return False, "No active window found."
            if win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMAXIMIZED:
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                return True, "Window restored."
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                return True, "Window maximized."
        except Exception:
            return False, "Error maximizing window."

    def minimize_current_window(self) -> Tuple[bool, str]:
        try:
            pyautogui.hotkey('win', 'down')
            return True, "Window minimized."
        except Exception:
            return False, "Error minimizing window."

    def close_current_window(self) -> Tuple[bool, str]:
        try:
            pyautogui.hotkey('alt', 'f4')
            return True, "Window closed."
        except Exception:
            return False, "Error closing window."

    def take_screenshot(self) -> Tuple[bool, str]:
        try:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
            if not os.path.isdir(desktop_path): os.makedirs(desktop_path)
            path = os.path.join(desktop_path, f"screenshot_{timestamp}.png")
            pyautogui.screenshot(path)
            return True, "Screenshot saved to desktop."
        except Exception:
            return False, "Error taking screenshot."

    def run_application(self, app_name: str) -> Tuple[bool, str]:
        if not app_name: return False, "Please specify an application."
        try:
            pyautogui.press('win')
            time.sleep(0.5)
            pyautogui.write(app_name, interval=0.05)
            time.sleep(0.5)
            pyautogui.press('enter')
            return True, f"Opening {app_name}."
        except Exception:
            return False, f"Could not open {app_name}."
            
    def change_wallpaper(self) -> Tuple[bool, str]:
        try:
            # This is a complex UI automation, we'll keep it simple
            pyautogui.hotkey('win', 'd') # Show desktop
            time.sleep(0.5)
            pyautogui.rightClick(100, 100) # Right click on a safe spot
            time.sleep(0.5)
            pyautogui.press('n') # 'Next desktop background'
            time.sleep(0.2)
            pyautogui.press('esc') # Close context menu
            return True, "Wallpaper changed."
        except Exception:
            return False, "Could not change wallpaper."
            
    def minimize_all_windows(self) -> Tuple[bool, str]:
        try:
            pyautogui.hotkey('win', 'd')
            return True, "Showing desktop."
        except Exception:
            return False, "Could not show desktop."

import platform
from typing import Optional, List
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
import urllib.parse
from grid_manager import GridManager

class OSManagement:
    def __init__(self, speech):
        self.speech = speech
        self.system = platform.system() 
        self.volume_interface = None
        self.brightness_interface = None
        self.window_handles = []
        self.current_window_index = 0
        if not hasattr(self, 'context'):
            self.context = {}

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
        finally:
            pythoncom.CoUninitialize()

        try:
            self.brightness_interface = wmi.WMI(namespace="wmi")
        except Exception as e:
            print(f"Error initializing Windows brightness interface: {e}")
            self.speech.speak("Error initializing brightness control. Using hotkey fallback.")
            self.brightness_interface = None

        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0
        self._update_window_handles()
        # Grid manager
        self.grid = GridManager(speech)

    def _update_window_handles(self):
        """Enumerate focusable application windows (including minimized), excluding Voice Listener."""
        temp_handles = []
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

            if "Voice Listener" in window_text:
                return

            if (win32gui.IsWindowVisible(hwnd) and
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
                    if pid == 0:
                        return
                    temp_handles.append(hwnd)
                except Exception:
                    pass

        win32gui.EnumWindows(enum_windows_callback, None)
        self.window_handles = temp_handles

        if not self.window_handles:
            self.current_window_index = 0
            print("No focusable windows found after filtering.")
        else:
            self.current_window_index = max(0, min(self.current_window_index, len(self.window_handles) - 1))
            print(f"Updated window handles: {[win32gui.GetWindowText(hwnd) for hwnd in self.window_handles]} (Count: {len(self.window_handles)})")

    def _restore_and_focus(self, hwnd):
        """Restore a minimized window and bring it to the foreground."""
        try:
            if not win32gui.IsWindow(hwnd):
                print(f"Invalid window handle: {hwnd}")
                return False

            window_text = win32gui.GetWindowText(hwnd)

            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')
                time.sleep(0.05)
            except Exception as e:
                print(f"Note: WScript.Shell SendKeys failed: {e}")

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.2)

            attempts = 0
            max_attempts = 3
            while attempts < max_attempts:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.1)
                if win32gui.GetForegroundWindow() == hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    return True

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
            foreground_at_call_start = win32gui.GetForegroundWindow()
            self._update_window_handles()

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

            prospective_next_idx = (self.current_window_index + 1) % len(self.window_handles)
            if len(self.window_handles) > 1 and self.window_handles[prospective_next_idx] == foreground_at_call_start:
                prospective_next_idx = (prospective_next_idx + 1) % len(self.window_handles)

            self.current_window_index = prospective_next_idx
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
                return False
        except Exception as e:
            print(f"Error in switch_window: {e}")
            self.speech.speak("Error switching window")
            self.current_window_index = 0
            self._update_window_handles()
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

            pythoncom.CoInitialize()
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                for window in shell.Windows():
                    if window.hwnd == hwnd:
                        path_url = window.LocationURL
                        if path_url.startswith("file:///"):
                            path = path_url[8:].replace("/", "\\")
                            path = urllib.parse.unquote(path)
                            if os.path.isdir(path):
                                return path
                        elif path_url == "" and "This PC" in window.LocationName:
                            return None  # This PC is not a specific directory
                        break
            finally:
                pythoncom.CoUninitialize()
            return None
        except Exception as e:
            print(f"Error getting active File Explorer path: {e}")
            return None

    def get_open_explorer_paths(self) -> List[str]:
        """Get paths of all open File Explorer windows, including minimized ones."""
        paths = []
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()

            for window in windows:
                try:
                    if "explorer.exe" in window.FullName.lower():
                        path_url = window.LocationURL
                        if path_url.startswith("file:///"):
                            path = path_url[8:].replace("/", "\\")
                            path = urllib.parse.unquote(path)
                            if os.path.isdir(path) and path not in paths:
                                paths.append(path)
                        # Skip "This PC" as it’s not a valid directory path
                except (AttributeError, Exception):
                    continue
            if not paths:
                print("No open File Explorer paths detected.")
            return paths
        except Exception as e:
            print(f"Error getting open File Explorer paths: {e}")
            self.speech.speak("Error detecting open folders.")
            return []
        finally:
            pythoncom.CoUninitialize()

    def get_active_window_title(self) -> Optional[str]:
        """Returns the title of the active window."""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                title = win32gui.GetWindowText(hwnd)
                return title
            return None
        except Exception as e:
            print(f"Error getting active window title: {e}")
            return None

    def volume_up(self) -> bool:
        if self.volume_interface:
            try:
                pythoncom.CoInitialize()
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = min(1.0, current_volume + 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume increased to {int(new_volume * 100)}%")
                return True
            except Exception as e:
                print(f"Error (pycaw) volume up: {e}. Fallback.")
            finally:
                pythoncom.CoUninitialize()
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
                pythoncom.CoInitialize()
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = max(0.0, current_volume - 0.1)
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume decreased to {int(new_volume * 100)}%")
                return True
            except Exception as e:
                print(f"Error (pycaw) volume down: {e}. Fallback.")
            finally:
                pythoncom.CoUninitialize()
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
                pythoncom.CoInitialize()
                is_muted = self.volume_interface.GetMute()
                self.volume_interface.SetMute(not is_muted, None)
                status = "muted" if not is_muted else "unmuted"
                print(f"Volume {status}")
                self.speech.speak(f"Volume {status}")
                return True
            except Exception as e:
                print(f"Error (pycaw) mute toggle: {e}. Fallback.")
            finally:
                pythoncom.CoUninitialize()
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
                pythoncom.CoInitialize()
                self.volume_interface.SetMasterVolumeLevelScalar(1.0, None)
                if self.volume_interface.GetMute():
                    self.volume_interface.SetMute(False, None)
                print("Volume set to maximum (100%)")
                self.speech.speak("Volume maximized")
                return True
            except Exception as e:
                print(f"Error (pycaw) maximizing volume: {e}. Fallback.")
            finally:
                pythoncom.CoUninitialize()
        try:
            is_muted = False
            if self.volume_interface:
                try:
                    pythoncom.CoInitialize()
                    is_muted = self.volume_interface.GetMute()
                except:
                    pass
                finally:
                    pythoncom.CoUninitialize()
            if is_muted:
                pyautogui.press('volumemute')
            for _ in range(10):
                pyautogui.press('volumeup')
                time.sleep(0.02)
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
                pythoncom.CoInitialize()
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
            finally:
                pythoncom.CoUninitialize()

        try:
            is_muted = False
            if self.volume_interface:
                try:
                    pythoncom.CoInitialize()
                    is_muted = self.volume_interface.GetMute()
                except:
                    pass
                finally:
                    pythoncom.CoUninitialize()
            if is_muted:
                pyautogui.press('volumemute')
            for _ in range(10):
                pyautogui.press('volumedown')
                time.sleep(0.02)
            if level > 0:
                num_steps = int(level / 5)
                for _ in range(num_steps):
                    pyautogui.press('volumeup')
                    time.sleep(0.02)
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
        if not self.brightness_interface:
            return False
        try:
            methods = self.brightness_interface.WmiMonitorBrightnessMethods()
            if not methods:
                raise Exception("WMI BrightnessMethods not found")
            monitor_methods = methods[0]
            monitor_methods.WmiSetBrightness(level_percent, 0)
            return True
        except Exception as e:
            print(f"Error setting brightness with WMI: {e}")
            return False

    def _get_brightness_wmi(self) -> Optional[int]:
        if not self.brightness_interface:
            return None
        try:
            brightness_info = self.brightness_interface.WmiMonitorBrightness()
            if not brightness_info:
                raise Exception("WMI Brightness info not found")
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
        """Restore all minimized application windows.

        Strategy:
        - Enumerate top-level app windows (including minimized)
        - For each that is minimized (IsIconic), call ShowWindow(SW_RESTORE)
        - Fallback to Win+Shift+M if none restored or enumeration fails
        """
        try:
            restored_count = 0
            excluded_classes = [
                "Progman", "WorkerW", "Shell_TrayWnd", "PopupHost",
                "Windows.UI.Core.CoreWindow", "ApplicationFrameWindow",
                "Button", "ComboBox", "Edit", "ListBox", "Static", "ToolbarWindow32",
                "NotifyIconOverflowWindow", "ShellExperienceHost", "SearchUI"
            ]

            def enum_cb(hwnd, _):
                nonlocal restored_count
                try:
                    class_name = win32gui.GetClassName(hwnd)
                    window_text = win32gui.GetWindowText(hwnd)
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)

                    if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowEnabled(hwnd):
                        return
                    if not window_text or "Voice Listener" in window_text:
                        return
                    if class_name in excluded_classes:
                        return
                    if "Windows Input Experience" in window_text:
                        return
                    if "DesktopWindowXamlSource" in class_name:
                        return
                    if (ex_style & win32con.WS_EX_TOOLWINDOW) or (ex_style & win32con.WS_EX_NOACTIVATE):
                        return
                    if (style & win32con.WS_CHILD):
                        return
                    if not ((style & win32con.WS_CAPTION) or (style & win32con.WS_SYSMENU)):
                        return

                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        restored_count += 1
                except Exception:
                    # Ignore per-window errors
                    pass

            win32gui.EnumWindows(enum_cb, None)

            if restored_count > 0:
                print(f"Restored {restored_count} minimized window(s)")
                self.speech.speak(f"Restored {restored_count} windows")
            else:
                # Fallback: Win+Shift+M restores minimized windows on many systems
                try:
                    pyautogui.hotkey('win', 'shift', 'm')
                    print("Sent Win+Shift+M to restore minimized windows")
                    self.speech.speak("Windows restored")
                except Exception as hotkey_err:
                    print(f"Fallback restore hotkey failed: {hotkey_err}")
                    self.speech.speak("No minimized windows found")

            self._update_window_handles()
            return True
        except Exception as e:
            print(f"Error restoring all: {e}")
            self.speech.speak("Error restoring")
            return False

    def maximize_current_window(self) -> bool:
        """Adaptive Maximize: browser → toggle F11; editors/docs → maximize + zoom; others → toggle maximize."""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                self.speech.speak("No active window to maximize")
                return False

            title = (win32gui.GetWindowText(hwnd) or "").lower()
            class_name = (win32gui.GetClassName(hwnd) or "").lower()

            def is_browser():
                keys = ["chrome", "edge", "brave", "opera", "firefox"]
                return any(k in title for k in keys) or any(k in class_name for k in keys)

            def is_editor_or_doc():
                keys = [
                    "visual studio code", " code", "notepad", "notepad++", "word", "excel",
                    "powerpoint", "sublime", "pycharm", "intellij", "rider", "atom",
                    "writer", "libreoffice", "onenote"
                ]
                return any(k in title for k in keys)

            # Case 1: Browsers → use true fullscreen toggle
            if is_browser():
                pyautogui.press('f11')
                print("Toggled browser full screen (F11)")
                self.speech.speak("Toggled full screen")
                return True

            # Helper to toggle maximize
            def _is_window_maximized(target_hwnd) -> bool:
                try:
                    placement = win32gui.GetWindowPlacement(target_hwnd)
                    # placement is a tuple: (flags, showCmd, ptMinPos, ptMaxPos, rcNormalPosition)
                    show_cmd = placement[1] if isinstance(placement, (list, tuple)) and len(placement) > 1 else None
                    return show_cmd == win32con.SW_SHOWMAXIMIZED
                except Exception:
                    return False

            def toggle_maximize():
                try:
                    if win32gui.IsIconic(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.05)
                    if _is_window_maximized(hwnd):
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        print("Restored window from maximized")
                        return "restored"
                    else:
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                        print("Maximized window")
                        return "maximized"
                except Exception as e:
                    print(f"Toggle maximize error: {e}")
                    return None

            # Case 2: Editors/Docs → maximize + zoom in; toggle restores and zooms out
            if is_editor_or_doc():
                state = toggle_maximize()
                if state == "maximized":
                    try:
                        pyautogui.hotkey('ctrl', '+')
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', '+')
                    except Exception:
                        pass
                    self.speech.speak("Maximized and zoomed in")
                    return True
                elif state == "restored":
                    try:
                        pyautogui.hotkey('ctrl', '-')
                        time.sleep(0.05)
                        pyautogui.hotkey('ctrl', '-')
                    except Exception:
                        pass
                    self.speech.speak("Restored and zoomed out")
                    return True
                else:
                    self.speech.speak("Could not change window state")
                    return False

            # Case 3: Default → toggle maximize
            state = toggle_maximize()
            if state == "maximized":
                self.speech.speak("Window maximized")
                return True
            elif state == "restored":
                self.speech.speak("Window restored")
                return True
            else:
                self.speech.speak("Error maximizing")
                return False
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