import os
import shutil
import subprocess
import pythoncom
import win32com.client
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

    def _get_selectable_locations(self) -> List[Tuple[str, str]]:
        """Identifies open (even minimized) File Explorer windows and the Desktop.
        Returns a list of tuples: (display_name, path)."""
        pythoncom.CoInitialize()
        locations = []
        
        # ONLY use Desktop - don't detect other File Explorer windows
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        locations.append(("Desktop", desktop_path))

        # Force Desktop as the working directory when available
        self.context["working_directory"] = desktop_path
        
        # Debug: Print available locations
        print(f"DEBUG: Only using Desktop location: {desktop_path}")
        
        return locations

    def _prompt_for_location_choice(self, pre_prompt: str = "Please choose a location:") -> Optional[str]:
        """Lists available locations and prompts the user to choose via voice.
        Returns the chosen path or None."""
        locations = self._get_selectable_locations()
        if not locations:
            self.speech.speak("No open folders or Desktop found to select from.")
            return None

        self.speech.speak(pre_prompt)
        for i, (display_name, path) in enumerate(locations, 1):
            self.speech.speak(f"Say {i} for {display_name}")

        wait_time = 0
        max_wait = 5  # Wait up to 5 seconds
        while wait_time < max_wait and self.voice_recognizer:
            try:
                response = self.voice_recognizer.get_transcription()
                if response and response.isdigit() and 1 <= int(response) <= len(locations):
                    selected_path = locations[int(response) - 1][1]
                    display_name = locations[int(response) - 1][0]
                    self.speech.speak(f"Selected {display_name}")
                    return selected_path
                time.sleep(0.5)
                wait_time += 0.5
            except Exception as e:
                print(f"Error getting transcription: {e}")

        self.speech.speak("No selection made. Using Desktop.")
        # Force the working directory to Desktop to prevent switching to other locations
        self._force_desktop_working_directory()
        return os.path.expanduser("~/Desktop")

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
        self.speech.speak("No name provided.")
        return None

    def _get_target_directory(self) -> str:
        """Determine the target directory for file operations."""
        print(f"DEBUG: _get_target_directory called. Context: {self.context}")
        
        # ALWAYS use Desktop for folder creation - ignore active explorer windows
        desktop_path = os.path.expanduser("~/Desktop")
        
        # Force Desktop as the working directory
        self.context["working_directory"] = desktop_path
        
        print(f"DEBUG: Forcing Desktop as target directory: {desktop_path}")
        self.speech.speak("Using Desktop")
        return desktop_path
    
    def _force_desktop_working_directory(self):
        """Force the working directory to be Desktop to prevent switching to other locations."""
        desktop_path = os.path.expanduser("~/Desktop")
        if os.path.isdir(desktop_path):
            # Reset all context to Desktop
            self.context["working_directory"] = desktop_path
            self.context["last_created_folder"] = None
            self.context["last_opened_item"] = (desktop_path, "Desktop")
            print(f"DEBUG: Forced working directory to Desktop: {desktop_path}")
            print(f"DEBUG: Context reset: {self.context}")
            return desktop_path
        return None
    
    def _reset_to_desktop(self):
        """Completely reset the system to work only with Desktop."""
        desktop_path = os.path.expanduser("~/Desktop")
        self.context = {
            "working_directory": desktop_path,
            "last_created_folder": None,
            "last_opened_item": (desktop_path, "Desktop")
        }
        print(f"DEBUG: System reset to Desktop: {self.context}")
        return desktop_path
    
    def _force_desktop_only(self):
        """Force the system to only work with Desktop and ignore all other locations."""
        desktop_path = os.path.expanduser("~/Desktop")
        
        # Override any OS manager calls to always return Desktop
        def mock_get_active_explorer_path():
            print("DEBUG: Mock get_active_explorer_path called - returning Desktop")
            return desktop_path
        
        def mock_get_active_window_title():
            print("DEBUG: Mock get_active_window_title called - returning Desktop")
            return "Desktop"
        
        # Store original methods
        if hasattr(self.os_manager, '_original_get_active_explorer_path'):
            pass
        else:
            self.os_manager._original_get_active_explorer_path = self.os_manager.get_active_explorer_path
            self.os_manager._original_get_active_window_title = self.os_manager.get_active_window_title
        
        # Override methods
        self.os_manager.get_active_explorer_path = mock_get_active_explorer_path
        self.os_manager.get_active_window_title = mock_get_active_window_title
        
        print(f"DEBUG: OS manager methods overridden to always return Desktop")
        return desktop_path

    def open_my_computer(self) -> bool:
        """Opens 'This PC' (My Computer) in File Explorer."""
        try:
            subprocess.run(['explorer', 'shell:MyComputerFolder'], shell=True)
            print("Opened This PC")
            self.speech.speak("This PC opened")
            self.context["working_directory"] = None  # My Computer is not a specific directory
            return True
        except Exception as e:
            print(f"Error opening This PC: {e}")
            self.speech.speak("Error opening This PC")
            return False

    def open_disk(self, disk_letter: str) -> bool:
        """Opens a specific disk (e.g., 'E:\\')."""
        disk_letter = disk_letter.strip().upper()
        disk_path = f"{disk_letter}:\\"
        if not os.path.exists(disk_path):
            print(f"Disk {disk_letter} does not exist")
            self.speech.speak(f"Disk {disk_letter} not found")
            return False
        try:
            subprocess.run(['explorer', disk_path], shell=True)
            print(f"Opened disk {disk_letter}")
            self.speech.speak(f"Disk {disk_letter} opened")
            self.context["working_directory"] = disk_path  # Update context
            return True
        except Exception as e:
            print(f"Error opening disk {disk_letter}: {e}")
            self.speech.speak(f"Error opening disk {disk_letter}")
            return False

    def go_back(self) -> bool:
        """Navigates to the parent directory of the current working directory, mimicking Alt + Up Arrow."""
        # Get the active window title to check if it's a File Explorer window
        active_title = self.os_manager.get_active_window_title()
        if not active_title or "File Explorer" not in active_title:
            print(f"Not in a File Explorer window: {active_title}")
            self.speech.speak("You must be in a File Explorer window to go back.")
            return False

        # Get the current path from the active File Explorer window
        current_path = self.os_manager.get_active_explorer_path()
        if not current_path or not os.path.isdir(current_path):
            print(f"No valid current directory found: {current_path}")
            self.speech.speak("No valid directory to go back from. Please open a folder or disk first.")
            return False

        self.context["working_directory"] = current_path  # Update context with current path

        parent_dir = os.path.dirname(current_path)
        if parent_dir == current_path:  # Root directory
            print(f"At root directory {current_path}, opening This PC")
            self.speech.speak("You are at the root of the disk. Opening This PC.")
            return self.open_my_computer()

        try:
            subprocess.run(['explorer', parent_dir], shell=True)
            print(f"Navigated to parent directory: {parent_dir}")
            self.speech.speak(f"Opened {os.path.basename(parent_dir)}")
            self.context["working_directory"] = parent_dir
            return True
        except Exception as e:
            print(f"Error navigating to parent directory: {e}")
            self.speech.speak("Error going back")
            return False

    def create_folder(self, folder_name: Optional[str] = None) -> bool:
        """Creates a folder and opens it immediately."""
        if not folder_name:
            folder_name = self._prompt_for_name("What do you want to name the new folder?")
            if not folder_name:
                self.speech.speak("Folder creation cancelled.")
                return False

        try:
            # COMPLETELY BYPASS EVERYTHING - ONLY DESKTOP
            print("DEBUG: === COMPLETE BYPASS MODE ===")
            
            # Get Desktop path using multiple methods to ensure it's correct
            desktop_path = os.path.expanduser("~/Desktop")
            if not os.path.exists(desktop_path):
                desktop_path = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
            if not os.path.exists(desktop_path):
                desktop_path = "C:\\Users\\SC\\Desktop"
            
            print(f"DEBUG: Final Desktop path: {desktop_path}")
            
            # Create folder path
            folder_path = os.path.join(desktop_path, folder_name)
            print(f"DEBUG: Creating folder at: {folder_path}")

            # Check if folder exists
            if os.path.exists(folder_path):
                print(f"DEBUG: Folder already exists")
                self.speech.speak(f"Folder {folder_name} already exists")
                return True

            # Create the folder
            os.makedirs(folder_path)
            print(f"DEBUG: Folder created successfully")
            self.speech.speak(f"Folder {folder_name} created in Desktop")
            
            # COMPLETELY IGNORE ALL CONTEXT AND OS MANAGER
            # Try to force Desktop using multiple methods
            try:
                print("DEBUG: Attempting to force Desktop open")
                
                # Method 1: Use shell:Desktop special folder
                try:
                    print("DEBUG: Method 1 - shell:Desktop")
                    subprocess.run(['explorer', 'shell:Desktop'], shell=True)
                    print("DEBUG: Desktop opened using shell:Desktop")
                    self.speech.speak("Desktop opened to show the new folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 1 failed: {e}")
                
                # Method 2: Use %USERPROFILE%\Desktop
                try:
                    print("DEBUG: Method 2 - %USERPROFILE%\\Desktop")
                    user_profile = os.environ.get('USERPROFILE', '')
                    desktop_cmd = f'explorer "{user_profile}\\Desktop"'
                    print(f"DEBUG: Command: {desktop_cmd}")
                    subprocess.run(desktop_cmd, shell=True)
                    print("DEBUG: Desktop opened using %USERPROFILE%")
                    self.speech.speak("Desktop opened to show the new folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 2 failed: {e}")
                
                # Method 3: Use Windows Run dialog equivalent
                try:
                    print("DEBUG: Method 3 - Windows Run")
                    run_cmd = f'cmd /c start explorer "{desktop_path}"'
                    print(f"DEBUG: Command: {run_cmd}")
                    subprocess.run(run_cmd, shell=True)
                    print("DEBUG: Desktop opened using cmd start")
                    self.speech.speak("Desktop opened to show the new folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 3 failed: {e}")
                
                # Method 4: Use PowerShell
                try:
                    print("DEBUG: Method 4 - PowerShell")
                    ps_cmd = f'powershell -command "Start-Process explorer -ArgumentList \'{desktop_path}\'"'
                    print(f"DEBUG: Command: {ps_cmd}")
                    subprocess.run(ps_cmd, shell=True)
                    print("DEBUG: Desktop opened using PowerShell")
                    self.speech.speak("Desktop opened to show the new folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 4 failed: {e}")
                
                # If all methods fail, just tell the user
                print("DEBUG: All opening methods failed")
                self.speech.speak("Folder created on Desktop but could not open Desktop view")
                return True
                
            except Exception as e:
                print(f"DEBUG: Critical error in opening methods: {e}")
                self.speech.speak("Folder created but could not open Desktop")
                return True
        except Exception as e:
            print(f"DEBUG: Critical error in create_folder: {e}")
            self.speech.speak(f"Error creating folder {folder_name}")
            return False

    def open_folder(self, folder_name: str) -> bool:
        """Opens a folder with the given name in the target directory."""
        if not folder_name:
            folder_name = self._prompt_for_name("What folder would you like to open?")
            if not folder_name:
                self.speech.speak("Open operation cancelled.")
                return False

        try:
            # Use Desktop directly - same as create/rename/delete methods
            desktop_path = os.path.expanduser("~/Desktop")
            folder_path = os.path.join(desktop_path, folder_name)
            
            print(f"DEBUG: Attempting to open folder: {folder_path}")

            if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
                print(f"Folder '{folder_name}' does not exist on Desktop")
                self.speech.speak(f"Folder {folder_name} not found on Desktop")
                return False

                        # Bypass Windows File Explorer completely - use alternative method
            try:
                print(f"DEBUG: Bypassing Windows File Explorer for: {folder_path}")
                
                # Method 1: Try to use Windows Run dialog equivalent
                try:
                    run_cmd = f'cmd /c start "" "{folder_path}"'
                    print(f"DEBUG: Method 1 - start command: {run_cmd}")
                    subprocess.run(run_cmd, shell=True)
                    print("DEBUG: Folder opened using start command")
                    self.speech.speak(f"Folder {folder_name} opened from Desktop")
                    
                    # Update context to the opened folder
                    self.context["working_directory"] = folder_path
                    self.context["last_opened_item"] = (folder_path, folder_name)
                    return True
                    
                except Exception as e:
                    print(f"DEBUG: Method 1 failed: {e}")
                
                # Method 2: Use Windows shell command directly
                try:
                    shell_cmd = f'explorer /select,"{folder_path}"'
                    print(f"DEBUG: Method 2 - shell command: {shell_cmd}")
                    subprocess.run(shell_cmd, shell=True)
                    print("DEBUG: Folder opened using shell command")
                    self.speech.speak(f"Folder {folder_name} opened from Desktop")
                    
                    # Update context to the opened folder
                    self.context["working_directory"] = folder_path
                    self.context["last_opened_item"] = (folder_path, folder_name)
                    return True
                    
                except Exception as e:
                    print(f"DEBUG: Method 2 failed: {e}")
                
                # Method 3: Last resort - use os.startfile
                try:
                    os.startfile(folder_path)
                    print("DEBUG: Folder opened using os.startfile")
                    self.speech.speak(f"Folder {folder_name} opened from Desktop")
                    
                    # Update context to the opened folder
                    self.context["working_directory"] = folder_path
                    self.context["last_opened_item"] = (folder_path, folder_name)
                    return True
                    
                except Exception as e:
                    print(f"DEBUG: Method 3 failed: {e}")
                
                # If all methods fail, just tell the user
                print("DEBUG: All methods failed")
                self.speech.speak(f"Could not open folder {folder_name} - Windows issue")
                return False
                
            except Exception as e:
                print(f"DEBUG: Critical error in folder opening: {e}")
                self.speech.speak(f"Error opening folder {folder_name}")
                return False
        except Exception as e:
            print(f"Error opening folder '{folder_name}': {e}")
            self.speech.speak(f"Error opening folder {folder_name}")
            return False

    def _prompt_for_confirmation(self, prompt_message: str = "Please say yes or no") -> bool:
        """Prompts for yes/no confirmation via voice input."""
        self.speech.speak(prompt_message)
        wait_time = 0
        max_wait = 5
        while wait_time < max_wait and self.voice_recognizer:
            try:
                response = self.voice_recognizer.get_transcription()
                if response:
                    response = response.lower().strip()
                    if response in ['yes', 'y', 'yeah', 'sure', 'okay', 'ok']:
                        return True
                    elif response in ['no', 'n', 'nope', 'cancel', 'stop']:
                        return False
                    else:
                        self.speech.speak("Please say yes or no")
                time.sleep(0.5)
                wait_time += 0.5
            except Exception as e:
                print(f"Error getting confirmation: {e}")
        self.speech.speak("No confirmation received. Cancelling operation.")
        return False

    def delete_folder(self, folder_name: str) -> bool:
        """Deletes a folder with the given name in the target directory."""
        if not folder_name:
            folder_name = self._prompt_for_name("What folder would you like to delete?")
            if not folder_name:
                self.speech.speak("Deletion cancelled.")
                return False

        try:
            # Use Desktop directly for deletion
            desktop_path = os.path.expanduser("~/Desktop")
            folder_path = os.path.join(desktop_path, folder_name)
            
            print(f"DEBUG: Attempting to delete folder: {folder_path}")

            if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
                print(f"Folder '{folder_name}' does not exist on Desktop")
                self.speech.speak(f"Folder {folder_name} not found on Desktop")
                return False

            # Delete the folder immediately without confirmation
            shutil.rmtree(folder_path)
            print(f"Deleted folder '{folder_name}' from Desktop")
            self.speech.speak(f"Folder {folder_name} deleted from Desktop")
            
            # Update context to Desktop
            self.context["working_directory"] = desktop_path
            self.context["last_created_folder"] = None
            self.context["last_opened_item"] = (desktop_path, "Desktop")
            
            # Open Desktop to show the deletion result
            try:
                print("DEBUG: Opening Desktop after deletion")
                
                # Method 1: Use shell:Desktop special folder
                try:
                    print("DEBUG: Method 1 - shell:Desktop")
                    subprocess.run(['explorer', 'shell:Desktop'], shell=True)
                    print("DEBUG: Desktop opened using shell:Desktop")
                    self.speech.speak("Desktop opened to show the deletion result")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 1 failed: {e}")
                
                # Method 2: Use %USERPROFILE%\Desktop
                try:
                    print("DEBUG: Method 2 - %USERPROFILE%\\Desktop")
                    user_profile = os.environ.get('USERPROFILE', '')
                    desktop_cmd = f'explorer "{user_profile}\\Desktop"'
                    print(f"DEBUG: Command: {desktop_cmd}")
                    subprocess.run(desktop_cmd, shell=True)
                    print("DEBUG: Desktop opened using %USERPROFILE%")
                    self.speech.speak("Desktop opened to show the deletion result")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 2 failed: {e}")
                
                # Method 3: Use Windows Run dialog equivalent
                try:
                    print("DEBUG: Method 3 - Windows Run")
                    run_cmd = f'cmd /c start explorer "{desktop_path}"'
                    print(f"DEBUG: Command: {run_cmd}")
                    subprocess.run(run_cmd, shell=True)
                    print("DEBUG: Desktop opened using cmd start")
                    self.speech.speak("Desktop opened to show the deletion result")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 3 failed: {e}")
                
                # Method 4: Use PowerShell
                try:
                    print("DEBUG: Method 4 - PowerShell")
                    ps_cmd = f'powershell -command "Start-Process explorer -ArgumentList \'{desktop_path}\'"'
                    print(f"DEBUG: Command: {ps_cmd}")
                    subprocess.run(ps_cmd, shell=True)
                    print("DEBUG: Desktop opened using PowerShell")
                    self.speech.speak("Desktop opened using PowerShell")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 4 failed: {e}")
                
                # If all methods fail, just tell the user
                print("DEBUG: All opening methods failed")
                self.speech.speak("Folder deleted but could not open Desktop view")
                return True
                
            except Exception as e:
                print(f"DEBUG: Critical error in opening methods: {e}")
                self.speech.speak("Folder deleted but could not open Desktop")
            return True
                
        except Exception as e:
            print(f"Error deleting folder '{folder_name}': {e}")
            self.speech.speak(f"Error deleting folder {folder_name}")
            return False

    def rename_folder(self, old_name: str, new_name: Optional[str] = None) -> bool:
        """Renames a folder from old_name to new_name in the target directory."""
        if not old_name:
            old_name = self._prompt_for_name("What is the current name of the folder?")
            if not old_name:
                self.speech.speak("Rename operation cancelled.")
                return False

        if not new_name:
            new_name = self._prompt_for_name(f"What would you like to rename {old_name} to?")
            if not new_name:
                self.speech.speak("Rename operation cancelled.")
                return False

        try:
            target_dir = self._get_target_directory()
            old_path = os.path.join(target_dir, old_name)
            new_path = os.path.join(target_dir, new_name)

            if not os.path.exists(old_path) or not os.path.isdir(old_path):
                print(f"Folder '{old_name}' does not exist in {target_dir}")
                self.speech.speak(f"Folder {old_name} not found")
                return False

            if os.path.exists(new_path):
                print(f"Folder '{new_name}' already exists in {target_dir}")
                self.speech.speak(f"Folder {new_name} already exists")
                return False

            os.rename(old_path, new_path)
            print(f"Renamed folder '{old_name}' to '{new_name}' in {target_dir}")
            self.speech.speak(f"Folder {old_name} renamed to {new_name} in {os.path.basename(target_dir)}")
            
            # Update context to Desktop
            desktop_path = os.path.expanduser("~/Desktop")
            self.context["working_directory"] = desktop_path
            self.context["last_created_folder"] = (desktop_path, new_name)
            self.context["last_opened_item"] = (desktop_path, "Desktop")
            
            # Open Desktop to show the renamed folder
            try:
                print("DEBUG: Opening Desktop after rename to show renamed folder")
                
                # Method 1: Use shell:Desktop special folder
                try:
                    print("DEBUG: Method 1 - shell:Desktop")
                    subprocess.run(['explorer', 'shell:Desktop'], shell=True)
                    print("DEBUG: Desktop opened using shell:Desktop")
                    self.speech.speak("Desktop opened to show the renamed folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 1 failed: {e}")
                
                # Method 2: Use %USERPROFILE%\Desktop
                try:
                    print("DEBUG: Method 2 - %USERPROFILE%\\Desktop")
                    user_profile = os.environ.get('USERPROFILE', '')
                    desktop_cmd = f'explorer "{user_profile}\\Desktop"'
                    print(f"DEBUG: Command: {desktop_cmd}")
                    subprocess.run(desktop_cmd, shell=True)
                    print("DEBUG: Desktop opened using %USERPROFILE%")
                    self.speech.speak("Desktop opened to show the renamed folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 2 failed: {e}")
                
                # Method 3: Use Windows Run dialog equivalent
                try:
                    print("DEBUG: Method 3 - Windows Run")
                    run_cmd = f'cmd /c start explorer "{desktop_path}"'
                    print(f"DEBUG: Command: {run_cmd}")
                    subprocess.run(run_cmd, shell=True)
                    print("DEBUG: Desktop opened using cmd start")
                    self.speech.speak("Desktop opened to show the renamed folder")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 3 failed: {e}")
                
                # Method 4: Use PowerShell
                try:
                    print("DEBUG: Method 4 - PowerShell")
                    ps_cmd = f'powershell -command "Start-Process explorer -ArgumentList \'{desktop_path}\'"'
                    print(f"DEBUG: Command: {ps_cmd}")
                    subprocess.run(ps_cmd, shell=True)
                    print("DEBUG: Desktop opened using PowerShell")
                    self.speech.speak("Desktop opened using PowerShell")
                    return True
                except Exception as e:
                    print(f"DEBUG: Method 4 failed: {e}")
                
                # If all methods fail, just tell the user
                print("DEBUG: All opening methods failed")
                self.speech.speak("Folder renamed but could not open Desktop view")
                return True
                
            except Exception as e:
                print(f"DEBUG: Critical error in opening methods: {e}")
                self.speech.speak("Folder renamed but could not open Desktop")
            return True
        except Exception as e:
            print(f"Error renaming folder '{old_name}' to '{new_name}': {e}")
            self.speech.speak(f"Error renaming folder {old_name} to {new_name}")
            return False