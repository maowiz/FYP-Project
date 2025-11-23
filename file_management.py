import os
import shutil
import subprocess
import pythoncom
import win32com.client
import win32gui
from typing import Optional, Tuple, List
import time
import urllib.parse
import pyttsx3

class FileManager:
    def __init__(self, speech, os_manager, voice_recognizer=None):
        self.speech = speech
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        self.command_handler = None
        self.context = {
            "working_directory": None,
            "last_created_folder": None,
            "last_opened_item": None
        }

    def _get_active_explorer_path(self) -> Optional[str]:
        """
        🎯 SMART DETECTION: Get the current File Explorer path if user is in Explorer.
        Returns:
            - Path string if in File Explorer
            - None if not in File Explorer (e.g., browser, other apps)
        """
        try:
            pythoncom.CoInitialize()
            hwnd = win32gui.GetForegroundWindow()
            
            if not hwnd:
                return None
            
            # Check if the active window is File Explorer
            class_name = win32gui.GetClassName(hwnd)
            
            if class_name not in ("CabinetWClass", "ExplorerWClass"):
                return None
            
            # Get the path from the active Explorer window
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                for window in shell.Windows():
                    if window.HWND == hwnd:
                        url = window.LocationURL
                        if url and url.startswith("file:///"):
                            # Convert file:/// URL to Windows path
                            path = urllib.parse.unquote(url[8:].replace("/", "\\"))
                            
                            # Verify it's a valid directory (not special folders)
                            if os.path.isdir(path):
                                return path
                            else:
                                return None
                        else:
                            return None
            except Exception as e:
                print(f"DEBUG: Error getting Explorer path: {e}")
                return None
            
            return None
            
        except Exception as e:
            print(f"DEBUG: Error in _get_active_explorer_path: {e}")
            return None
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def _get_desktop_path(self) -> str:
        """Get the Desktop path using multiple fallback methods."""
        # Method 1: Standard expanduser
        desktop = os.path.expanduser("~/Desktop")
        if os.path.isdir(desktop):
            return desktop
        
        # Method 2: Environment variable
        userprofile = os.environ.get('USERPROFILE', '')
        if userprofile:
            desktop = os.path.join(userprofile, 'Desktop')
            if os.path.isdir(desktop):
                return desktop
        
        # Method 3: Common path (fallback)
        desktop = r"C:\Users\{}\Desktop".format(os.getlogin())
        if os.path.isdir(desktop):
            return desktop
        
        # Final fallback
        return os.path.expanduser("~/Desktop")

    def _get_smart_target_directory(self) -> Tuple[str, str]:
        """
        🧠 INTELLIGENT DIRECTORY DETECTION
        Returns: (target_path, location_description)
        - If in File Explorer: Use current Explorer location
        - Otherwise: Use Desktop as fallback
        """
        # Try to get active Explorer path
        explorer_path = self._get_active_explorer_path()
        
        if explorer_path:
            # User is in File Explorer - use that location
            location_name = os.path.basename(explorer_path) or "root of drive"
            return (explorer_path, location_name)
        else:
            # User is NOT in File Explorer - fallback to Desktop
            desktop_path = self._get_desktop_path()
            return (desktop_path, "Desktop")

    def open_my_computer(self) -> Tuple[bool, str]:
        """Opens 'This PC' (My Computer) in File Explorer."""
        try:
            subprocess.run(['explorer', 'shell:MyComputerFolder'], shell=True)
            print("Opened This PC")
            self.context["working_directory"] = None  # My Computer is not a specific directory
            return True, "This PC opened"
        except Exception as e:
            print(f"Error opening This PC: {e}")
            return False, "Error opening This PC"

    def open_disk(self, disk_letter: str) -> Tuple[bool, str]:
        """Opens a specific disk (e.g., 'E:\\')."""
        disk_letter = disk_letter.strip().upper()
        disk_path = f"{disk_letter}:\\"
        if not os.path.exists(disk_path):
            print(f"Disk {disk_letter} does not exist")
            return False, f"Disk {disk_letter} not found"
        try:
            subprocess.run(['explorer', disk_path], shell=True)
            print(f"Opened disk {disk_letter}")
            self.context["working_directory"] = disk_path  # Update context
            return True, f"Disk {disk_letter} opened"
        except Exception as e:
            print(f"Error opening disk {disk_letter}: {e}")
            return False, f"Error opening disk {disk_letter}"

    def go_back(self) -> Tuple[bool, str]:
        """Navigates to the parent directory using Alt+Up if in File Explorer."""
        try:
            import pyautogui
            
            # Check if in File Explorer first
            explorer_path = self._get_active_explorer_path()
            if not explorer_path:
                return False, "Not in a File Explorer window"
            
            # Send Alt+Up to go to parent directory
            pyautogui.hotkey('alt', 'up')
            print("Navigated to parent folder using Alt+Up")
            return True, "Navigated to parent folder"
            
        except Exception as e:
            print(f"Error navigating back: {e}")
            return False, "Error going back"

    def create_folder(self, folder_name: Optional[str] = None) -> Tuple[bool, str]:
        """
        🚀 SMART FOLDER CREATION
        Returns: (success, message)
        """
        if not folder_name:
            folder_name = self._prompt_for_name("What do you want to name the new folder?")
            if not folder_name:
                return False, "Folder creation cancelled"

        try:
            # 🧠 Get smart target directory
            target_dir, location_name = self._get_smart_target_directory()
            folder_path = os.path.join(target_dir, folder_name)
            
            print(f"DEBUG: Creating folder '{folder_name}' in {target_dir}")

            # Check if folder already exists
            if os.path.exists(folder_path):
                print(f"Folder '{folder_name}' already exists in {location_name}")
                
                # Update context anyway
                self.context["working_directory"] = target_dir
                self.context["last_created_folder"] = (target_dir, folder_name)
                return True, f"Folder {folder_name} already exists in {location_name}"

            # Create the folder
            os.makedirs(folder_path, exist_ok=True)
            print(f"✅ Created folder '{folder_name}' in {target_dir}")
            
            # Update context IMMEDIATELY
            self.context["working_directory"] = target_dir
            self.context["last_created_folder"] = (target_dir, folder_name)
            self.context["last_opened_item"] = (target_dir, folder_name)
            
            # Open the folder in Explorer
            try:
                subprocess.run(['explorer', folder_path], shell=True, timeout=2)
                print(f"Opened folder: {folder_path}")
            except Exception as e:
                print(f"Could not open folder: {e}")
            
            # Provide location-specific feedback
            if location_name == "Desktop":
                return True, f"Folder {folder_name} created on Desktop"
            else:
                return True, f"Folder {folder_name} created in {location_name}"

        except Exception as e:
            print(f"Error creating folder '{folder_name}': {e}")
            return False, f"Error creating folder {folder_name}"

    def open_folder(self, folder_name: str) -> Tuple[bool, str]:
        """Opens a folder with the given name in the current location."""
        if not folder_name:
            folder_name = self._prompt_for_name("What folder would you like to open?")
            if not folder_name:
                return False, "Open operation cancelled"

        try:
            # Get smart target directory
            target_dir, location_name = self._get_smart_target_directory()
            folder_path = os.path.join(target_dir, folder_name)
            
            # Normalize path for Windows
            folder_path = os.path.normpath(folder_path)
            
            print(f"DEBUG: Opening folder '{folder_name}' from {target_dir}")
            print(f"DEBUG: Full path: {folder_path}")
            print(f"DEBUG: Folder exists: {os.path.exists(folder_path)}")
            print(f"DEBUG: Is directory: {os.path.isdir(folder_path) if os.path.exists(folder_path) else 'N/A'}")

            # Check if folder exists
            if not os.path.exists(folder_path):
                print(f"Folder '{folder_name}' not found in {location_name}")
                return False, f"Folder {folder_name} not found in {location_name}"
            
            # Check if it's actually a directory
            if not os.path.isdir(folder_path):
                print(f"'{folder_name}' exists but is not a folder in {location_name}")
                return False, f"{folder_name} is not a folder in {location_name}"

            # Open the folder (Explorer exit codes are unreliable, so we don't check them)
            subprocess.run(['explorer', folder_path], shell=True)
            
            print(f"Opened folder: {folder_path}")
            
            # Update context
            self.context["working_directory"] = folder_path
            self.context["last_opened_item"] = (folder_path, folder_name)
            
            return True, f"Folder {folder_name} opened from {location_name}"

        except Exception as e:
            print(f"Error opening folder '{folder_name}': {e}")
            import traceback
            traceback.print_exc()
            return False, f"Error opening folder {folder_name}"

    def delete_folder(self, folder_name: str) -> Tuple[bool, str]:
        """Deletes a folder with the given name in the current location."""
        if not folder_name:
            folder_name = self._prompt_for_name("What folder would you like to delete?")
            if not folder_name:
                return False, "Deletion cancelled"

        try:
            # Get smart target directory
            target_dir, location_name = self._get_smart_target_directory()
            folder_path = os.path.join(target_dir, folder_name)
            
            print(f"DEBUG: Deleting folder '{folder_name}' from {target_dir}")

            if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
                print(f"Folder '{folder_name}' not found in {location_name}")
                return False, f"Folder {folder_name} not found in {location_name}"

            # Delete the folder
            shutil.rmtree(folder_path)
            print(f"Deleted folder: {folder_path}")
            
            # Update context
            self.context["working_directory"] = target_dir
            self.context["last_created_folder"] = None
            self.context["last_opened_item"] = (target_dir, location_name)
            
            return True, f"Folder {folder_name} deleted from {location_name}"

        except Exception as e:
            print(f"Error deleting folder '{folder_name}': {e}")
            return False, f"Error deleting folder {folder_name}"

    def rename_folder(self, old_name: str, new_name: Optional[str] = None) -> Tuple[bool, str]:
        """Renames a folder from old_name to new_name in the current location."""
        if not old_name:
            old_name = self._prompt_for_name("What is the current name of the folder?")
            if not old_name:
                return False, "Rename operation cancelled"

        if not new_name:
            new_name = self._prompt_for_name(f"What would you like to rename {old_name} to?")
            if not new_name:
                return False, "Rename operation cancelled"

        try:
            # Get smart target directory
            target_dir, location_name = self._get_smart_target_directory()
            old_path = os.path.join(target_dir, old_name)
            new_path = os.path.join(target_dir, new_name)

            print(f"DEBUG: Renaming '{old_name}' to '{new_name}' in {target_dir}")

            if not os.path.exists(old_path) or not os.path.isdir(old_path):
                print(f"Folder '{old_name}' not found in {location_name}")
                return False, f"Folder {old_name} not found in {location_name}"

            if os.path.exists(new_path):
                print(f"Folder '{new_name}' already exists in {location_name}")
                return False, f"Folder {new_name} already exists in {location_name}"

            # Rename the folder
            os.rename(old_path, new_path)
            print(f"Renamed folder '{old_name}' to '{new_name}'")
            
            # Update context
            self.context["working_directory"] = target_dir
            self.context["last_created_folder"] = (target_dir, new_name)
            self.context["last_opened_item"] = (target_dir, new_name)
            
            return True, f"Folder {old_name} renamed to {new_name} in {location_name}"

        except Exception as e:
            print(f"Error renaming folder '{old_name}' to '{new_name}': {e}")
            return False, f"Error renaming folder {old_name} to {new_name}"

    def _prompt_for_name(self, prompt_message: str = "What name would you like to use?") -> Optional[str]:
        """Prompts for a name via voice input, supporting spaces."""
        self.speech.speak(prompt_message)
        wait_time = 0
        max_wait = 5
        while wait_time < max_wait and self.voice_recognizer:
            try:
                name = self.voice_recognizer.get_transcription()
                if name:
                    name = name.strip()
                    # Validate name for illegal characters
                    illegal_chars = r'[\\/:*?"<>|]'
                    import re
                    if re.search(illegal_chars, name):
                        self.speech.speak("The name contains invalid characters. Please try again.")
                        continue
                    return name
                self.speech.speak("The name cannot be empty. Please try again.")
                time.sleep(0.5)
                wait_time += 0.5
            except Exception as e:
                print(f"Error getting name: {e}")
        self.speech.speak("No name provided")
        return None