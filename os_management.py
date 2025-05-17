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

class OSManagement:
    def __init__(self, speech):
        self.speech = speech
        self.system = platform.system()
        self.volume_interface = None
        self.brightness_interface = None

        if self.system != "Windows":
            print("This application only supports Windows.")
            self.speech.speak("This application only supports Windows.")
            return

        # Initialize audio interface
        try:
            pythoncom.CoInitialize()
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self.volume_interface = interface.QueryInterface(IAudioEndpointVolume)
        except Exception as e:
            print(f"Error initializing Windows audio interface: {e}")
            self.speech.speak("Error initializing audio control. Using hotkey fallback.")

        # Initialize brightness interface
        try:
            self.brightness_interface = wmi.WMI(namespace="wmi")
        except Exception as e:
            print(f"Error initializing Windows brightness interface: {e}")
            self.speech.speak("Error initializing brightness control. Using hotkey fallback.")
            self.brightness_interface = None

        # Configure pyautogui for hotkey control
        pyautogui.FAILSAFE = True  # Move mouse to upper-left corner to abort
        pyautogui.PAUSE = 0.1  # Small pause between key presses

    def volume_up(self) -> bool:
        """Increase the system volume by 10%."""
        if self.volume_interface:
            try:
                current_volume = self.volume_interface.GetMasterVolumeLevelScalar()
                new_volume = min(1.0, current_volume + 0.1)  # Increase by 10%, cap at 100%
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume increased to {int(new_volume * 100)}%")
                self.speech.speak("Volume increased")
                return True
            except Exception as e:
                print(f"Error increasing volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        # Hotkey fallback
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
                new_volume = max(0.0, current_volume - 0.1)  # Decrease by 10%, floor at 0%
                self.volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
                print(f"Volume decreased to {int(new_volume * 100)}%")
                self.speech.speak("Volume decreased")
                return True
            except Exception as e:
                print(f"Error decreasing volume with pycaw: {e}. Falling back to hotkey.")
                self.speech.speak("Error with volume control. Trying hotkey.")
        
        # Hotkey fallback
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
        
        # Hotkey fallback
        try:
            pyautogui.press('volumemute')
            status = "toggled"  # Exact mute state is harder to determine with hotkeys
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
        
        # Hotkey fallback (simulate multiple volume up presses)
        try:
            for _ in range(10):  # Press volume up multiple times to approximate max
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
        
        # Hotkey fallback (approximate by resetting to 0 then increasing)
        try:
            # Mute and press volume down to approximate 0%
            pyautogui.press('volumemute')
            for _ in range(10):
                pyautogui.press('volumedown')
                time.sleep(0.05)
            # Unmute and press volume up to reach desired level
            pyautogui.press('volumemute')
            presses = int(level // 5)  # Roughly 5% per press
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
                new_brightness = min(100, current_brightness + 10)  # Increase by 10%, cap at 100%
                monitor.WmiSetBrightness(new_brightness, 0)
                print(f"Brightness increased to {new_brightness}%")
                self.speech.speak("Brightness increased")
                return True
            except Exception as e:
                print(f"Error increasing brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        # Hotkey fallback (assumes Fn + F6 or similar, may need customization)
        # Note: Brightness keys vary by laptop (e.g., Fn+F5/F6 on Dell, Fn+F7/F8 on Lenovo)
        try:
            pyautogui.hotkey('fn', 'f6')  # Example hotkey, adjust based on device
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
                new_brightness = max(0, current_brightness - 10)  # Decrease by 10%, floor at 0%
                monitor.WmiSetBrightness(new_brightness, 0)
                print(f"Brightness decreased to {new_brightness}%")
                self.speech.speak("Brightness decreased")
                return True
            except Exception as e:
                print(f"Error decreasing brightness with wmi: {e}. Falling back to hotkey.")
                self.speech.speak("Error with brightness control. Trying hotkey.")
        
        # Hotkey fallback (assumes Fn + F5 or similar, may need customization)
        try:
            pyautogui.hotkey('fn', 'f5')  # Example hotkey, adjust based on device
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
        
        # Hotkey fallback (simulate multiple brightness up presses)
        try:
            for _ in range(10):  # Press brightness up multiple times to approximate max
                pyautogui.hotkey('fn', 'f6')  # Adjust based on device
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
        
        # Hotkey fallback (approximate by resetting to 0 then increasing)
        try:
            # Press brightness down to approximate 0%
            for _ in range(10):
                pyautogui.hotkey('fn', 'f5')  # Adjust based on device
                time.sleep(0.05)
            # Press brightness up to reach desired level
            presses = int(level // 5)  # Roughly 5% per press
            for _ in range(presses):
                pyautogui.hotkey('fn', 'f6')  # Adjust based on device
                time.sleep(0.05)
            print(f"Brightness set to approximately {level}% via hotkey")
            self.speech.speak(f"Brightness set to {level} percent")
            return True
        except Exception as e:
            print(f"Error with hotkey set brightness: {e}. Check if Fn+F5/F6 are correct for your device.")
            self.speech.speak("Error setting brightness. Check your brightness hotkey.")
            return False

    def switch_window(self) -> bool:
        """Switch to the next open window."""
        try:
            pyautogui.hotkey('alt', 'tab')
            print("Switched to next window")
            self.speech.speak("Window switched")
            return True
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
            # Windows+Down minimizes if the window is maximized or restored
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
            # Generate timestamped filename
            timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
            desktop_path = os.path.expanduser("~/Desktop")
            screenshot_path = os.path.join(desktop_path, f"screenshot_{timestamp}.png")
            
            # Take and save screenshot
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            print(f"Screenshot saved to {screenshot_path}")
            self.speech.speak("Screenshot taken and saved to Desktop")
            return True
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            self.speech.speak("Error taking screenshot")
            return False