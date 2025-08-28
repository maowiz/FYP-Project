import os  # Added to resolve NameError for os.path.basename and os.path.join

class FileCommandHandler:
    def __init__(self, file_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.voice_recognizer = voice_recognizer

    # ---------- NEW HELPER ----------
    def _is_in_explorer_folder(self):
        """Return True if the currently-focused window is File Explorer
           and its current location is a real directory (not Desktop or This PC)."""
        try:
            import win32gui
            import pythoncom
            import win32com.client
            import os
            import urllib.parse

            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return False

            if win32gui.GetClassName(hwnd) not in ("CabinetWClass", "ExplorerWClass"):
                return False

            pythoncom.CoInitialize()
            try:
                shell = win32com.client.Dispatch("Shell.Application")
                for w in shell.Windows():
                    if w.hwnd == hwnd:
                        url = w.LocationURL
                        if url.startswith("file:///"):
                            path = urllib.parse.unquote(url[8:].replace("/", "\\"))
                            if os.path.isdir(path):
                                return True
                        break
            finally:
                pythoncom.CoUninitialize()
            return False
        except Exception:
            return False

    def handle_open_my_computer(self, context=None) -> bool:
        """Handle the 'open my computer' command."""
        result = self.file_manager.open_my_computer()
        if result and context is not None:
            context["last_opened_item"] = None
            context["last_created_folder"] = None
            context["working_directory"] = None  # My Computer is not a specific directory
        if not result:
            self.file_manager.speech.speak("Failed to open This PC.")
        return result

    def handle_open_disk(self, disk_letter, context=None) -> bool:
        """Handle the 'open disk <letter>' command."""
        if not disk_letter:
            self.file_manager.speech.speak("Please specify a disk letter.")
            return False
        result = self.file_manager.open_disk(disk_letter)
        if result and context is not None:
            disk_path = f"{disk_letter.upper()}:\\"
            context["last_opened_item"] = (disk_path, disk_path)
            context["last_created_folder"] = None
            context["working_directory"] = disk_path
        if not result:
            self.file_manager.speech.speak(f"Failed to open disk {disk_letter.upper()}.")
        return result

    def handle_go_back(self, context=None) -> bool:
        """Go up one folder **only** if currently inside an Explorer folder."""
        if self._is_in_explorer_folder():
            try:
                import pyautogui
                pyautogui.hotkey('alt', 'up')
                print("Sent Alt+↑ to go up one folder.")
                return True
            except Exception as e:
                print(f"Error sending Alt+Up: {e}")
                return False
        else:
            # Optional: give feedback (or just silently ignore)
            self.file_manager.speech.speak("Not in a folder window.")
            print("Ignoring go-back: not in File Explorer inside a folder.")
            return False

    def handle_create_folder(self, folder_name, context=None) -> bool:
        """Handle the 'create folder <name>' command."""
        if not folder_name and self.voice_recognizer:
            wait_time = 0
            max_wait = 5
            self.file_manager.speech.speak("What do you want to name the new folder?")
            while wait_time < max_wait:
                try:
                    folder_name = self.voice_recognizer.get_transcription()
                    if folder_name:
                        break
                    time.sleep(0.5)
                    wait_time += 0.5
                except Exception as e:
                    print(f"Error getting folder name: {e}")
            if not folder_name:
                self.file_manager.speech.speak("No folder name provided. Creation cancelled.")
                return False

        result = self.file_manager.create_folder(folder_name)
        if result and context is not None:
            target_dir = self.file_manager.context["working_directory"]
            if target_dir and os.path.isdir(target_dir):
                context["last_created_folder"] = (target_dir, folder_name)
                context["last_opened_item"] = (target_dir, folder_name)  # Folder is opened after creation
                context["working_directory"] = os.path.join(target_dir, folder_name)
            else:
                self.file_manager.speech.speak("Unable to update context due to invalid working directory.")
        if not result:
            self.file_manager.speech.speak(f"Failed to create folder {folder_name}.")
        return result

    def handle_open_folder(self, folder_name, context=None) -> bool:
        """Handle the 'open folder <name>' command."""
        if not folder_name and self.voice_recognizer:
            wait_time = 0
            max_wait = 5
            self.file_manager.speech.speak("What folder would you like to open?")
            while wait_time < max_wait:
                try:
                    folder_name = self.voice_recognizer.get_transcription()
                    if folder_name:
                        break
                    time.sleep(0.5)
                    wait_time += 0.5
                except Exception as e:
                    print(f"Error getting folder name: {e}")
            if not folder_name:
                self.file_manager.speak("No folder name provided. Open cancelled.")
                return False

        result = self.file_manager.open_folder(folder_name)
        if result and context is not None:
            target_dir = self.file_manager.context["working_directory"]
            if target_dir and os.path.isdir(target_dir):
                new_path = os.path.join(target_dir, folder_name)
                context["last_opened_item"] = (new_path, folder_name)
                context["last_created_folder"] = None
                context["working_directory"] = new_path  # Update to the opened folder
            else:
                self.file_manager.speech.speak("Unable to update context due to invalid working directory.")
        if not result:
            self.file_manager.speech.speak(f"Failed to open folder {folder_name}.")
        return result

    def handle_delete_folder(self, folder_name, context=None) -> bool:
        """Handle the 'delete folder <name>' command."""
        if not folder_name and self.voice_recognizer:
            wait_time = 0
            max_wait = 5
            self.file_manager.speak("What folder would you like to delete?")
            while wait_time < max_wait:
                try:
                    folder_name = self.voice_recognizer.get_transcription()
                    if folder_name:
                        break
                    time.sleep(0.5)
                    wait_time += 0.5
                except Exception as e:
                    print(f"Error getting folder name: {e}")
            if not folder_name:
                self.file_manager.speak("No folder name provided. Deletion cancelled.")
                return False

        result = self.file_manager.delete_folder(folder_name)
        if result and context is not None:
            target_dir = self.file_manager.context["working_directory"]
            if context.get("last_created_folder") and context["last_created_folder"][1] == folder_name:
                context["last_created_folder"] = None
            if context.get("last_opened_item") and context["last_opened_item"][1] == folder_name:
                context["last_opened_item"] = None
            context["working_directory"] = target_dir if target_dir and os.path.isdir(target_dir) else None
        if not result:
            self.file_manager.speech.speak(f"Failed to delete folder {folder_name}.")
        return result

    def handle_rename_folder(self, old_name, new_name, context=None) -> bool:
        """Handle the 'rename folder <old_name> to <new_name>' command."""
        if not old_name and self.voice_recognizer:
            wait_time = 0
            max_wait = 5
            self.file_manager.speak("What is the current name of the folder?")
            while wait_time < max_wait:
                try:
                    old_name = self.voice_recognizer.get_transcription()
                    if old_name:
                        break
                    time.sleep(0.5)
                    wait_time += 0.5
                except Exception as e:
                    print(f"Error getting old folder name: {e}")
            if not old_name:
                self.file_manager.speak("No folder name provided. Rename cancelled.")
                return False

        if not new_name and self.voice_recognizer:
            wait_time = 0
            max_wait = 5
            self.file_manager.speak(f"What would you like to rename {old_name} to?")
            while wait_time < max_wait:
                try:
                    new_name = self.voice_recognizer.get_transcription()
                    if new_name:
                        break
                    time.sleep(0.5)
                    wait_time += 0.5
                except Exception as e:
                    print(f"Error getting new folder name: {e}")
            if not new_name:
                self.file_manager.speak("No new name provided. Rename cancelled.")
                return False

        result = self.file_manager.rename_folder(old_name, new_name)
        if result and context is not None:
            target_dir = self.file_manager.context["working_directory"]
            if target_dir and os.path.isdir(target_dir):
                if context.get("last_created_folder") and context["last_created_folder"][1] == old_name:
                    context["last_created_folder"] = (target_dir, new_name)
                if context.get("last_opened_item") and context["last_opened_item"][1] == old_name:
                    context["last_opened_item"] = (target_dir, new_name)
                context["working_directory"] = target_dir
            else:
                self.file_manager.speech.speak("Unable to update context due to invalid working directory.")
        if not result:
            self.file_manager.speech.speak(f"Failed to rename folder {old_name} to {new_name}.")
        return result