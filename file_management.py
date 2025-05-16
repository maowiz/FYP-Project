import os
import subprocess
import platform
import win32gui
import win32com.client
import win32process
import shutil
import time
from pathlib import Path

# Optional for .docx files
try:
    import docx
except ImportError:
    print("python-docx not installed. .docx file reading will be unavailable.")

class FileManager:
    # Custom exception to return to main menu
    class ReturnToMain(Exception):
        pass

    def __init__(self, speech):
        self.speech = speech  # Speech object for voice feedback
        self.speak = speech.speak  # Shorthand for the speak function
        self.clipboard = None  # Simulated clipboard for copy/cut operations

    def get_active_explorer_directory(self):
        """Get the directory of the currently active File Explorer window (Windows only)."""
        try:
            foreground = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(foreground)
            shell = win32com.client.Dispatch("Shell.Application")
            for window in shell.Windows():
                if window.HWND == foreground and ("File Explorer" in window.Name or "Windows Explorer" in window.Name):
                    path = window.Document.Folder.Self.Path
                    print(f"Detected active File Explorer directory: {path}")
                    return path
            print("No active File Explorer window found.")
            return None
        except Exception as e:
            print(f"Error getting active explorer directory: {e}")
            return None

    def get_open_folders_windows(self):
        """Get directories of all open File Explorer windows (Windows only)."""
        try:
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()
            folders = []
            for window in windows:
                if "File Explorer" in window.Name or "Windows Explorer" in window.Name:
                    path = window.Document.Folder.Self.Path
                    folders.append(path)
            return folders
        except Exception as e:
            print(f"Error getting open folders: {e}")
            return []

    def is_desktop_active(self):
        """Check if the desktop is the active window (Windows only)."""
        try:
            foreground = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(foreground)
            window_class = win32gui.GetClassName(foreground)
            _, pid = win32process.GetWindowThreadProcessId(foreground)
            process_name = win32process.GetModuleFileNameEx(win32process.OpenProcess(0x0400, False, pid), 0)
            print(f"Foreground window: title='{window_title}', class='{window_class}', process='{process_name}'")
            if "Program Manager" in window_title and window_class == "Progman":
                return True
            if "Desktop" in window_title and "Explorer.EXE" in process_name and window_class == "WorkerW":
                return True
            return False
        except Exception as e:
            print(f"Error checking desktop activity: {e}")
            return False

    def select_directory(self, voice_recognizer=None):
        """List available directories and return the selected one based on user input or voice command."""
        search_dirs = []
        if platform.system() == "Windows":
            active_dir = self.get_active_explorer_directory()
            if active_dir and active_dir not in search_dirs:
                search_dirs.append(active_dir)
        desktop_path = os.path.expanduser("~/Desktop")
        if desktop_path not in search_dirs:
            search_dirs.append(desktop_path)
        cwd = os.getcwd()
        if cwd not in search_dirs:
            search_dirs.append(cwd)
        if platform.system() == "Windows":
            open_folders = self.get_open_folders_windows()
            search_dirs.extend([f for f in open_folders if f not in search_dirs])
        if not search_dirs:
            search_dirs.append(desktop_path)

        print("Available directories (say directory number or 'quit' to return to main menu):")
        for i, dir_path in enumerate(search_dirs, 1):
            print(f"{i}. {dir_path} ({'one' if i == 1 else 'two' if i == 2 else 'three' if i == 3 else str(i)})")

        # Speak all directory options
        self.speak("Available directories. Say the number to select:")
        
        # Check for voice input while listing directories
        if voice_recognizer:
            for i, dir_path in enumerate(search_dirs, 1):
                dir_name = os.path.basename(dir_path)
                number_word = 'one' if i == 1 else 'two' if i == 2 else 'three' if i == 3 else str(i)
                self.speak(f"Number {number_word}: {dir_name}")
                
                # Check for voice input after each directory announcement
                wait_count = 0
                while wait_count < 2:  # Check for input for 1 second
                    cmd_text = voice_recognizer.get_transcription()
                    if cmd_text:
                        print(f"You said: {cmd_text}")
                        
                        # Check for quit/stop command
                        if "quit" in cmd_text.lower() or "stop" in cmd_text.lower():
                            print("Returning to main menu")
                            self.speak("Returning to main menu")
                            raise self.ReturnToMain()
                        
                        # Check if user said a number
                        number_map = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, 
                                    "1": 1, "2": 2, "3": 3, "4": 4, "5": 5}
                        
                        for word, num in number_map.items():
                            if word in cmd_text.lower():
                                if 1 <= num <= len(search_dirs):
                                    selected_dir = search_dirs[num - 1]
                                    print(f"Selected directory: {selected_dir}")
                                    self.speak(f"Selected directory: {os.path.basename(selected_dir)}")
                                    return selected_dir
                    
                    time.sleep(0.5)
                    wait_count += 1

        # Wait for voice input if voice_recognizer is provided
        if voice_recognizer:
            self.speak("Say the directory number you want to select.")
            wait_time = 0
            while wait_time < 15:  # Wait up to 15 seconds for voice input
                cmd_text = voice_recognizer.get_transcription()
                if cmd_text:
                    print(f"You said: {cmd_text}")
                    
                    # Check for quit command
                    if "quit" in cmd_text.lower():
                        raise self.ReturnToMain()
                    
                    # Try to extract a number from the voice command
                    number_map = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, 
                                 "1": 1, "2": 2, "3": 3, "4": 4, "5": 5}
                    
                    for word, num in number_map.items():
                        if word in cmd_text.lower():
                            if 1 <= num <= len(search_dirs):
                                selected_dir = search_dirs[num - 1]
                                print(f"Selected directory: {selected_dir}")
                                self.speak(f"Selected directory: {os.path.basename(selected_dir)}")
                                return selected_dir
                    
                    # If no valid number found
                    print("I didn't understand. Please say a number between 1 and", len(search_dirs))
                    self.speak(f"I didn't understand. Please say a number between 1 and {len(search_dirs)}")
                    
                time.sleep(0.5)
                wait_time += 0.5
            
            print("No voice input detected. Falling back to text input.")
            self.speak("No voice input detected. Please type your selection.")

        # Fallback to text input if no voice input or voice_recognizer not provided
        while True:
            choice = input(f"Choose a directory (1-{len(search_dirs)} or one/two/three, 'quit' to return to main menu): ").lower().strip()
            if choice == 'quit':
                raise self.ReturnToMain()
            number_map = {"one": 1, "two": 2, "three": 3}
            try:
                if choice in number_map:
                    choice = number_map[choice]
                choice = int(choice)
                if 1 <= choice <= len(search_dirs):
                    selected_dir = search_dirs[choice - 1]
                    print(f"Selected directory: {selected_dir}")
                    self.speak(f"Selected directory: {os.path.basename(selected_dir)}")
                    return selected_dir
                else:
                    print(f"Please enter a number between 1 and {len(search_dirs)}.")
                    self.speak(f"Please enter a number between 1 and {len(search_dirs)}.")
            except ValueError:
                print("Invalid input. Please enter a number (1, 2, 3) or word (one, two, three).")
                self.speak("Invalid input. Please enter a number one, two, three or one, two, three.")

    def get_folder_or_file_name(self, prompt, voice_recognizer=None):
        """Ask for the folder or file name and return it using voice or text input."""
        # Try voice input first if voice_recognizer is provided
        if voice_recognizer:
            print(prompt + " (say 'quit' to return to main menu)")
            self.speak(prompt)
            
            wait_time = 0
            while wait_time < 15:  # Wait up to 15 seconds for voice input
                cmd_text = voice_recognizer.get_transcription()
                if cmd_text:
                    print(f"You said: {cmd_text}")
                    
                    # Check for quit command
                    if "quit" in cmd_text.lower():
                        raise self.ReturnToMain()
                    
                    # Clean up the name (remove common phrases like "call it" or "name it")
                    name = cmd_text.lower()
                    # Remove common phrases that might be misinterpreted
                    for phrase in ["call it", "name it", "folder name", "file name", "name", "called", 
                                  "number", "one", "two", "three", "four", "five", "1", "2", "3", "4", "5"]:
                        name = name.replace(phrase, "").strip()
                    
                    # If after cleaning the name is too short, ask for clarification
                    if len(name) < 2:
                        print("The name is too short. Please provide a longer name.")
                        self.speak("The name is too short. Please provide a longer name.")
                        continue
                    
                    if name:
                        print(f"Using name: {name}")
                        self.speak(f"Using name: {name}")
                        return name
                    else:
                        print("I couldn't understand a name. Please try again.")
                        self.speak("I couldn't understand a name. Please try again.")
                
                time.sleep(0.5)
                wait_time += 0.5
            
            print("No voice input detected. Falling back to text input.")
            self.speak("No voice input detected. Please type the name.")
        
        # Fallback to text input
        while True:
            name = input(prompt + " ('quit' to return to main menu): ").strip()
            if name.lower() == 'quit':
                raise self.ReturnToMain()
            if name:
                return name
            print("Name cannot be empty. Please try again.")
            self.speak("Name cannot be empty. Please try again.")

    def create_folder(self, directory, folder_name):
        """Create a folder in the specified directory with the given name."""
        while True:
            try:
                folder_path = os.path.join(directory, folder_name)
                os.makedirs(folder_path, exist_ok=False)
                print(f"Created folder: {folder_path}")
                self.speak("Folder created")
                break
            except FileExistsError:
                print(f"The folder '{folder_name}' already exists at {folder_path}.")
                self.speak(f"The folder {folder_name} already exists at {folder_path}.")
            except Exception as e:
                print(f"Error creating folder: {e}")
                self.speak(f"Error creating folder: {e}")
            
            retry = input("Do you want to select a different directory? (yes/no, 'quit' to return to main menu): ").lower().strip()
            if retry == 'quit':
                raise ReturnToMain()
            if retry == 'yes':
                return self.select_directory(), folder_name
            elif retry == 'no':
                print("Operation cancelled.")
                self.speak("Operation cancelled.")
                break
            else:
                print("Please enter 'yes' or 'no'.")
                self.speak("Please enter yes or no.")

    def open_folder_or_file(self, directory, name, file_type=None):
        """Open a folder or file in the specified directory."""
        path = os.path.join(directory, name)
        if os.path.isdir(path):
            try:
                print(f"Opening folder: {path}")
                self.speak(f"Opening folder: {path}")
                if platform.system() == "Windows":
                    os.startfile(path)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", path])
                else:
                    subprocess.run(["xdg-open", path])
                self.speak("Folder opened")
                return True
            except Exception as e:
                print(f"Error opening folder: {e}")
                self.speak(f"Error opening folder: {e}")
                return False
        elif os.path.isfile(path):
            if file_type:
                if not name.endswith(file_type):
                    print(f"File '{name}' does not match the specified type '{file_type}'.")
                    self.speak(f"File {name} does not match the specified type {file_type}.")
                    return False
            try:
                print(f"Opening file: {path}")
                self.speak(f"Opening file: {path}")
                if platform.system() == "Windows":
                    os.startfile(path)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", path])
                else:
                    subprocess.run(["xdg-open", path])
                self.speak("File opened")
                return True
            except Exception as e:
                print(f"Error opening file: {e}")
                self.speak(f"Error opening file: {e}")
                return False
        else:
            print(f"'{name}' does not exist in {directory} as a folder or file.")
            self.speak("Not found. Try again.")
            # Return a special tuple to indicate we need to try again
            return ("not_found", directory)

    def delete_folder(self, directory, folder_name):
        """Delete a folder in the specified directory."""
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            try:
                os.rmdir(folder_path)
                print(f"Deleted folder: {folder_path}")
                self.speak("Folder deleted")
                return True
            except OSError as e:
                print(f"Error deleting folder: {e}")
                self.speak(f"Error deleting folder: {e}")
                return False
        else:
            print(f"The folder '{folder_name}' does not exist in {directory}.")
            self.speak(f"The folder {folder_name} does not exist in {directory}.")
            return False

    def copy_folder(self, directory, folder_name):
        """Copy a folder to a simulated clipboard."""
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            self.clipboard = {"action": "copy", "source": folder_path, "name": folder_name}
            print(f"Copied folder '{folder_name}' from {directory} to clipboard.")
            self.speak("Folder copied")
            return True
        else:
            print(f"The folder '{folder_name}' does not exist in {directory}.")
            self.speak(f"The folder {folder_name} does not exist in {directory}.")
            return False

    def cut_folder(self, directory, folder_name):
        """Cut a folder to a simulated clipboard."""
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            self.clipboard = {"action": "cut", "source": folder_path, "name": folder_name}
            print(f"Cut folder '{folder_name}' from {directory} to clipboard.")
            self.speak("Folder cut")
            return True
        else:
            print(f"The folder '{folder_name}' does not exist in {directory}.")
            self.speak(f"The folder {folder_name} does not exist in {directory}.")
            return False

    def paste_folder(self, directory):
        """Paste the copied/cut folder to the specified directory."""
        if not self.clipboard:
            print("No folder in clipboard to paste.")
            self.speak("No folder in clipboard to paste.")
            return False
        source_path = self.clipboard["source"]
        folder_name = self.clipboard["name"]
        dest_path = os.path.join(directory, folder_name)
        try:
            if self.clipboard["action"] == "copy":
                shutil.copytree(source_path, dest_path, dirs_exist_ok=True)
                print(f"Copied folder to {dest_path}")
                self.speak("Folder pasted")
            elif self.clipboard["action"] == "cut":
                shutil.move(source_path, dest_path)
                print(f"Moved folder to {dest_path}")
                self.speak("Folder pasted")
                self.clipboard = None
            return True
        except Exception as e:
            print(f"Error pasting folder: {e}")
            self.speak(f"Error pasting folder: {e}")
            return False

    def rename_folder_or_file(self, directory, old_name, new_name):
        """Rename a folder or file in the specified directory."""
        old_path = os.path.join(directory, old_name)
        new_path = os.path.join(directory, new_name)
        if os.path.exists(old_path):
            try:
                os.rename(old_path, new_path)
                print(f"Renamed '{old_name}' to '{new_name}' in {directory}")
                self.speak("Folder renamed")
                return True
            except Exception as e:
                print(f"Error renaming: {e}")
                self.speak(f"Error renaming: {e}")
                return False
        else:
            print(f"'{old_name}' does not exist in {directory}.")
            self.speak("Not found. Try again.")
            return False

    def read_text_file(self, directory, file_name):
        """Read text, Word (.docx), or Markdown (.md) files."""
        file_path = os.path.join(directory, file_name)
        if not os.path.isfile(file_path):
            print(f"'{file_name}' does not exist in {directory}.")
            self.speak("Not found. Try again.")
            return False
        extension = os.path.splitext(file_path)[1].lower()
        try:
            if extension == ".txt":
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    print(f"Content of {file_name}:\n{content}")
                    self.speak(f"Content of {file_name}")
                    self.speak(content)
            elif extension == ".docx" and "docx" in globals():
                doc = docx.Document(file_path)
                content = "\n".join([para.text for para in doc.paragraphs])
                print(f"Content of {file_name}:\n{content}")
                self.speak(f"Content of {file_name}")
                self.speak(content)
            elif extension == ".md":
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    print(f"Content of {file_name}:\n{content}")
                    self.speak(f"Content of {file_name}")
                    self.speak(content)
            else:
                print(f"Cannot read '{file_name}': Unsupported format (supported: .txt, .docx, .md).")
                self.speak(f"Cannot read {file_name}. Unsupported format. Supported formats are text, docx, and md.")
                return False
            self.speak("File read")
            return True
        except Exception as e:
            print(f"Error reading file: {e}")
            self.speak(f"Error reading file: {e}")
            return False

    def list_contents(self, directory):
        """List all files and folders in the specified directory."""
        try:
            items = os.listdir(directory)
            if not items:
                print(f"No files or folders in {directory}.")
                self.speak(f"No files or folders in {directory}.")
                return
            print(f"Contents of {directory}:")
            self.speak(f"Contents of {directory}")
            for item in items:
                path = os.path.join(directory, item)
                is_dir = " (folder)" if os.path.isdir(path) else ""
                print(f"- {item}{is_dir}")
                self.speak(f"- {item} {is_dir}")
            self.speak("Listing complete")
        except Exception as e:
            print(f"Error listing contents: {e}")
            self.speak(f"Error listing contents: {e}")

    def get_properties(self, directory, name):
        """Get properties of a file or folder in the specified directory."""
        path = os.path.join(directory, name)
        if not os.path.exists(path):
            print(f"'{name}' does not exist in {directory}.")
            self.speak("Not found. Try again.")
            return
        stat = os.stat(path)
        size = stat.st_size / 1024  # Size in KB
        ctime = time.ctime(stat.st_ctime)
        mtime = time.ctime(stat.st_mtime)
        is_dir = "folder" if os.path.isdir(path) else "file"
        print(f"Properties of '{name}' in {directory}:")
        print(f"- Type: {is_dir}")
        print(f"- Size: {size:.2f} KB")
        print(f"- Created: {ctime}")
        print(f"- Last modified: {mtime}")
        self.speak(f"Properties of {name} in {directory}")
        self.speak(f"Type: {is_dir}")
        self.speak(f"Size: {size:.2f} kilobytes")
        self.speak(f"Created: {ctime}")
        self.speak(f"Last modified: {mtime}")
        self.speak("Properties retrieved")