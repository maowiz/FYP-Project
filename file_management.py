import os
import shutil
import subprocess
from typing import Optional, Tuple, List
import time  # Added import for time.sleep

class FileManager:
    def __init__(self, speech, os_manager, voice_recognizer=None):
        self.speech = speech
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        self.context = {"working_directory": None}

    def _get_target_directory(self) -> str:
        """Determine the target directory for file operations.
        Prioritizes active File Explorer, then lists open directories for user selection, then Desktop."""
        # Try active File Explorer directory
        active_path = self.os_manager.get_active_explorer_path()
        if active_path and os.path.isdir(active_path):
            print(f"Using active directory: {active_path}")
            self.speech.speak(f"Using directory {os.path.basename(active_path)}")
            return active_path

        # Get list of open File Explorer directories
        open_paths = self.os_manager.get_open_explorer_paths()
        if open_paths:
            print("No active File Explorer directory found. Available directories:")
            self.speech.speak("No active folder found. Available directories are:")
            for i, path in enumerate(open_paths, 1):
                last_folder = os.path.basename(path)
                print(f"  {i}. {last_folder} ({path})")
                self.speech.speak(f"Say {i} for {last_folder}")

            # Wait for user selection via voice
            wait_time = 0
            max_wait = 5  # Wait up to 5 seconds
            selected_path = None
            while wait_time < max_wait and self.voice_recognizer:
                try:
                    response = self.voice_recognizer.get_transcription()
                    if response and response.isdigit() and 1 <= int(response) <= len(open_paths):
                        selected_path = open_paths[int(response) - 1]
                        break
                except Exception as e:
                    print(f"Error getting transcription: {e}")
                time.sleep(0.5)
                wait_time += 0.5

            if selected_path:
                last_folder = os.path.basename(selected_path)
                print(f"Selected directory: {selected_path}")
                self.speech.speak(f"Using directory {last_folder}")
                return selected_path
            else:
                # Default to the first detected path if no selection
                if open_paths:
                    print(f"No selection made. Using first available directory: {open_paths[0]}")
                    self.speech.speak(f"No selection. Using {os.path.basename(open_paths[0])}")
                    return open_paths[0]

        # Fall back to Desktop if no valid directory is found
        desktop_path = os.path.expanduser("~/Desktop")
        print(f"No valid directories found. Using Desktop: {desktop_path}")
        self.speech.speak("Using Desktop")
        return desktop_path

    def create_folder(self, folder_name: str) -> bool:
        """Create a folder with the given name in the target directory."""
        if not folder_name:
            print("No folder name provided.")
            self.speech.speak("Please provide a folder name and try again.")
            return False

        try:
            target_dir = self._get_target_directory()
            folder_path = os.path.join(target_dir, folder_name)

            if os.path.exists(folder_path):
                print(f"Folder '{folder_name}' already exists in {target_dir}.")
                self.speech.speak(f"Folder {folder_name} already exists.")
                return False

            os.makedirs(folder_path)
            print(f"Created folder '{folder_name}' in {target_dir}.")
            self.speech.speak(f"Folder {folder_name} created in {os.path.basename(target_dir)}.")
            self.context["working_directory"] = target_dir
            return True

        except Exception as e:
            print(f"Error creating folder '{folder_name}': {e}")
            self.speech.speak(f"Error creating folder {folder_name}. Please try again.")
            return False

    def delete_folder(self, folder_name: str) -> bool:
        """Delete a folder with the given name in the target directory."""
        if not folder_name:
            print("No folder name provided.")
            self.speech.speak("Please provide a folder name and try again.")
            return False

        try:
            target_dir = self._get_target_directory()
            folder_path = os.path.join(target_dir, folder_name)

            if not os.path.exists(folder_path):
                print(f"Folder '{folder_name}' does not exist in {target_dir}.")
                self.speech.speak(f"Folder {folder_name} not found.")
                return False

            shutil.rmtree(folder_path)
            print(f"Deleted folder '{folder_name}' from {target_dir}.")
            self.speech.speak(f"Folder {folder_name} deleted from {os.path.basename(target_dir)}.")
            self.context["working_directory"] = target_dir
            return True

        except Exception as e:
            print(f"Error deleting folder '{folder_name}': {e}")
            self.speech.speak(f"Error deleting folder {folder_name}. Please try again.")
            return False

    def rename_folder(self, old_name: str, new_name: str) -> bool:
        """Rename a folder from old_name to new_name in the target directory."""
        if not old_name or not new_name:
            print("Old or new folder name not provided.")
            self.speech.speak("Please provide both the old and new folder names and try again.")
            return False

        try:
            target_dir = self._get_target_directory()
            old_path = os.path.join(target_dir, old_name)
            new_path = os.path.join(target_dir, new_name)

            if not os.path.exists(old_path):
                print(f"Folder '{old_name}' does not exist in {target_dir}.")
                self.speech.speak(f"Folder {old_name} not found.")
                return False

            if os.path.exists(new_path):
                print(f"Folder '{new_name}' already exists in {target_dir}.")
                self.speech.speak(f"Folder {new_name} already exists.")
                return False

            os.rename(old_path, new_path)
            print(f"Renamed folder '{old_name}' to '{new_name}' in {target_dir}.")
            self.speech.speak(f"Folder {old_name} renamed to {new_name} in {os.path.basename(target_dir)}.")
            self.context["working_directory"] = target_dir
            return True

        except Exception as e:
            print(f"Error renaming folder '{old_name}' to '{new_name}': {e}")
            self.speech.speak(f"Error renaming folder {old_name} to {new_name}. Please try again.")
            return False

    def open_folder(self, folder_name: str) -> bool:
        """Open a folder with the given name in the target directory."""
        if not folder_name:
            print("No folder name provided.")
            self.speech.speak("Please provide a folder name and try again.")
            return False

        try:
            target_dir = self._get_target_directory()
            folder_path = os.path.join(target_dir, folder_name)

            if not os.path.exists(folder_path):
                print(f"Folder '{folder_name}' does not exist in {target_dir}.")
                self.speech.speak(f"Folder {folder_name} not found.")
                return False

            subprocess.run(['explorer', folder_path], shell=True)
            print(f"Opened folder '{folder_name}' in {target_dir}.")
            self.speech.speak(f"Folder {folder_name} opened from {os.path.basename(target_dir)}.")
            self.context["working_directory"] = folder_path
            return True

        except Exception as e:
            print(f"Error opening folder '{folder_name}': {e}")
            self.speech.speak(f"Error opening folder {folder_name}. Please try again.")
            return False