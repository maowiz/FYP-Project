import os
import shutil
import datetime
import docx
import time

class FileManager:
    def __init__(self, speech, os_manager):
        self.speech = speech
        self.os_manager = os_manager  # Add OSManagement instance for window detection
        self.clipboard = None  # Store folder for copy/cut operations

    class ReturnToMain(Exception):
        """Custom exception to return to main menu."""
        pass

    def select_directory(self, voice_recognizer=None):
        """Determine the working directory based on active File Explorer windows."""
        # Step 1: Check for active File Explorer window
        active_path = self.os_manager.get_active_explorer_path()
        if active_path and os.path.isdir(active_path):
            last_folder = os.path.basename(active_path)
            print(f"Using active directory: {active_path}")
            self.speech.speak(f"Using directory {last_folder}")
            return active_path

        # Step 2: Check for open (including minimized) File Explorer windows
        explorer_paths = self.os_manager.get_open_explorer_paths()
        if explorer_paths:
            print("No active File Explorer window found. Open folders detected:")
            self.speech.speak("No active folder window found. Which folder should I use?")
            for i, path in enumerate(explorer_paths, 1):
                last_folder = os.path.basename(path)
                print(f"  {i}. {last_folder} ({path})")
                self.speech.speak(f"Say {i} for {last_folder}")

            # Wait for user selection
            wait_time = 0
            max_wait = 5  # Wait up to 5 seconds
            selected_path = None
            while wait_time < max_wait:
                if voice_recognizer:
                    response = voice_recognizer.get_transcription()
                    if response:
                        print(f"You said: {response}")
                        try:
                            # Handle numeric input (e.g., "1", "2") or folder name
                            response = response.lower().strip()
                            if response.isdigit() and 1 <= int(response) <= len(explorer_paths):
                                selected_path = explorer_paths[int(response) - 1]
                                break
                            # Check if response matches a folder name
                            for path in explorer_paths:
                                if os.path.basename(path).lower() in response:
                                    selected_path = path
                                    break
                            if selected_path:
                                break
                            print("Invalid selection. Please say a number or folder name.")
                            self.speech.speak("Invalid selection. Please say a number or folder name.")
                        except ValueError:
                            print("Please say a valid number or folder name.")
                            self.speech.speak("Please say a valid number or folder name.")
                time.sleep(0.5)
                wait_time += 0.5

            if selected_path:
                last_folder = os.path.basename(selected_path)
                print(f"Selected directory: {selected_path}")
                self.speech.speak(f"Using directory {last_folder}")
                return selected_path
            else:
                print("No selection made. Defaulting to Desktop.")
                self.speech.speak("No selection made. Defaulting to Desktop.")

        # Step 3: Default to Desktop if no Explorer windows are open
        desktop_path = os.path.expanduser("~/Desktop")
        print(f"No open folders found. Using Desktop: {desktop_path}")
        self.speech.speak("Using Desktop")
        return desktop_path

    def get_folder_or_file_name(self, prompt, voice_recognizer=None):
        """Get folder or file name via voice or text."""
        while True:
            if voice_recognizer:
                print(prompt)
                self.speech.speak(prompt.replace("Enter", "Say").replace(": ", ""))
                name = voice_recognizer.get_transcription()
            else:
                name = input(prompt).strip()

            if name and name.lower() == "quit":
                raise self.ReturnToMain()

            if name:
                return name
            print("No name provided. Please try again.")
            self.speech.speak("No name provided. Please try again.")

    def create_folder(self, directory, folder_name):
        """Create a folder in the specified directory."""
        folder_path = os.path.join(directory, folder_name)
        try:
            os.makedirs(folder_path, exist_ok=True)
            last_folder = os.path.basename(directory)
            print(f"Folder '{folder_name}' created in {directory}")
            self.speech.speak(f"Folder {folder_name} created in {last_folder}")
            return directory, folder_name
        except Exception as e:
            print(f"Error creating folder: {e}")
            self.speech.speak(f"Error creating folder: {e}")
            return False

    def open_folder_or_file(self, directory, name, file_type=None):
        """Open a folder or file."""
        path = os.path.join(directory, name)
        if file_type:
            path = path + file_type
        try:
            if os.path.isdir(path):
                os.startfile(path)
                print(f"Opened folder: {path}")
                self.speech.speak(f"Opened folder {name}")
                return True
            elif os.path.isfile(path):
                os.startfile(path)
                print(f"Opened file: {path}")
                self.speech.speak(f"Opened file {name}")
                return True
            else:
                print(f"'{name}' does not exist in {directory}.")
                self.speech.speak("Not found. Try again.")
                return ("not_found", name)
        except Exception as e:
            print(f"Error opening item: {e}")
            self.speech.speak(f"Error opening item: {e}")
            return False

    def delete_folder(self, directory, folder_name):
        """Delete a folder."""
        folder_path = os.path.join(directory, folder_name)
        try:
            if os.path.isdir(folder_path):
                shutil.rmtree(folder_path)
                last_folder = os.path.basename(directory)
                print(f"Folder '{folder_name}' deleted from {directory}")
                self.speech.speak(f"Folder {folder_name} deleted from {last_folder}")
                return True
            else:
                print(f"Folder '{folder_name}' does not exist in {directory}")
                self.speech.speak(f"Folder {folder_name} does not exist")
                return False
        except Exception as e:
            print(f"Error deleting folder: {e}")
            self.speech.speak(f"Error deleting folder: {e}")
            return False

    def copy_folder(self, directory, folder_name):
        """Copy a folder to clipboard."""
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            self.clipboard = {"type": "copy", "path": folder_path, "name": folder_name}
            last_folder = os.path.basename(directory)
            print(f"Folder '{folder_name}' copied from {directory}")
            self.speech.speak(f"Folder {folder_name} copied from {last_folder}")
            return True
        else:
            print(f"Folder '{folder_name}' does not exist in {directory}")
            self.speech.speak(f"Folder {folder_name} does not exist")
            return False

    def cut_folder(self, directory, folder_name):
        """Cut a folder to clipboard."""
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            self.clipboard = {"type": "cut", "path": folder_path, "name": folder_name}
            last_folder = os.path.basename(directory)
            print(f"Folder '{folder_name}' cut from {directory}")
            self.speech.speak(f"Folder {folder_name} cut from {last_folder}")
            return True
        else:
            print(f"Folder '{folder_name}' does not exist in {directory}")
            self.speech.speak(f"Folder {folder_name} does not exist")
            return False

    def paste_folder(self, directory):
        """Paste a folder from clipboard."""
        if not self.clipboard:
            print("Clipboard is empty. Copy or cut a folder first.")
            self.speech.speak("Clipboard is empty. Copy or cut a folder first.")
            return False
        try:
            dest_path = os.path.join(directory, self.clipboard["name"])
            last_folder = os.path.basename(directory)
            if self.clipboard["type"] == "copy":
                shutil.copytree(self.clipboard["path"], dest_path, dirs_exist_ok=True)
                print(f"Folder '{self.clipboard['name']}' pasted to {directory}")
                self.speech.speak(f"Folder {self.clipboard['name']} pasted to {last_folder}")
            elif self.clipboard["type"] == "cut":
                shutil.move(self.clipboard["path"], dest_path)
                print(f"Folder '{self.clipboard['name']}' moved to {directory}")
                self.speech.speak(f"Folder {self.clipboard['name']} moved to {last_folder}")
                self.clipboard = None
            return True
        except Exception as e:
            print(f"Error pasting folder: {e}")
            self.speech.speak(f"Error pasting folder: {e}")
            return False

    def rename_folder_or_file(self, directory, old_name, new_name):
        """Rename a folder or file."""
        old_path = os.path.join(directory, old_name)
        new_path = os.path.join(directory, new_name)
        try:
            if os.path.exists(old_path):
                os.rename(old_path, new_path)
                last_folder = os.path.basename(directory)
                print(f"Renamed '{old_name}' to '{new_name}' in {directory}")
                self.speech.speak(f"Renamed {old_name} to {new_name} in {last_folder}")
                return True
            else:
                print(f"'{old_name}' does not exist in {directory}")
                self.speech.speak(f"{old_name} does not exist")
                return False
        except Exception as e:
            print(f"Error renaming item: {e}")
            self.speech.speak(f"Error renaming item: {e}")
            return False

    def read_text_file(self, directory, file_name, file_type=None):
        """Read text, Word (.docx), or Markdown (.md) files, handling names without extensions."""
        if file_type and not file_name.lower().endswith(('.txt', '.docx', '.md')):
            file_path = os.path.join(directory, f"{file_name}{file_type}")
        else:
            file_path = os.path.join(directory, file_name)

        if not os.path.isfile(file_path):
            possible_files = [
                os.path.join(directory, f"{file_name}{ext}")
                for ext in ['.txt', '.docx', '.md']
                if os.path.isfile(os.path.join(directory, f"{file_name}{ext}"))
            ]
            if not possible_files:
                last_folder = os.path.basename(directory)
                print(f"'{file_name}' does not exist in {directory}.")
                self.speech.speak(f"Not found in {last_folder}. Try again.")
                return False
            elif len(possible_files) > 1 and not file_type:
                print(f"Multiple files found for '{file_name}': {', '.join(possible_files)}")
                self.speech.speak(f"Multiple files found for {file_name}. Please specify text or Word.")
                return ("ambiguous", possible_files)
            else:
                file_path = possible_files[0]

        extension = os.path.splitext(file_path)[1].lower()
        try:
            last_folder = os.path.basename(directory)
            if extension == ".txt":
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    print(f"Content of {file_name} in {directory}:\n{content}")
                    self.speech.speak(f"Content of {file_name} in {last_folder}")
                    self.speech.speak(content)
            elif extension == ".docx" and "docx" in globals():
                doc = docx.Document(file_path)
                content = "\n".join([para.text for para in doc.paragraphs])
                print(f"Content of {file_name} in {directory}:\n{content}")
                self.speech.speak(f"Content of {file_name} in {last_folder}")
                self.speech.speak(content)
            elif extension == ".md":
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    print(f"Content of {file_name} in {directory}:\n{content}")
                    self.speech.speak(f"Content of {file_name} in {last_folder}")
                    self.speech.speak(content)
            else:
                print(f"Cannot read '{file_name}': Unsupported format (supported: .txt, .docx, .md).")
                self.speech.speak(f"Cannot read {file_name}. Unsupported format. Supported formats are text, docx, and md.")
                return False
            self.speech.speak("File read")
            return True
        except Exception as e:
            print(f"Error reading file: {e}")
            self.speech.speak(f"Error reading file: {e}")
            return False

    def list_contents(self, directory, extension=None):
        """List directory contents, optionally filtering by extension."""
        try:
            items = os.listdir(directory)
            if extension:
                items = [item for item in items if item.lower().endswith(extension.lower())]
            last_folder = os.path.basename(directory)
            if not items:
                print(f"No {'files' if extension else 'items'} found in {directory}")
                self.speech.speak(f"No {'files' if extension else 'items'} found in {last_folder}")
                return
            print(f"{'Files' if extension else 'Contents'} in {directory}:")
            self.speech.speak(f"{'Files' if extension else 'Contents'} in {last_folder}")
            for item in items:
                print(f"  - {item}")
                self.speech.speak(item)
        except Exception as e:
            print(f"Error listing contents: {e}")
            self.speech.speak(f"Error listing contents: {e}")

    def get_properties(self, directory, name):
        """Get properties of a folder or file."""
        path = os.path.join(directory, name)
        try:
            if not os.path.exists(path):
                last_folder = os.path.basename(directory)
                print(f"'{name}' does not exist in {directory}")
                self.speech.speak(f"{name} does not exist in {last_folder}")
                return
            stats = os.stat(path)
            is_dir = os.path.isdir(path)
            size = stats.st_size
            created = datetime.datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d %H:%M:%S")
            modified = datetime.datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            last_folder = os.path.basename(directory)
            properties = (
                f"Properties of '{name}' in {directory}:\n"
                f"  Type: {'Folder' if is_dir else 'File'}\n"
                f"  Size: {size / 1024:.2f} KB\n"
                f"  Created: {created}\n"
                f"  Modified: {modified}"
            )
            print(properties)
            self.speech.speak(f"Properties of {name} in {last_folder}: {'Folder' if is_dir else 'File'}, {size / 1024:.2f} KB, created {created}, modified {modified}")
        except Exception as e:
            print(f"Error getting properties: {e}")
            self.speech.speak(f"Error getting properties: {e}")