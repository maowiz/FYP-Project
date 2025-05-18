import sys
import os
import time
import re
from fuzzywuzzy import fuzz

class CommandHandler:
    def __init__(self, file_manager, os_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        # Expanded command mappings with synonyms and parameter requirements
        self.commands = {
            "create folder": {"handler": self.handle_create_folder, "params": None},
            "open folder or file": {"handler": self.handle_open_folder_or_file, "params": None},
            "delete folder": {"handler": self.handle_delete_folder, "params": None},
            "copy folder": {"handler": self.handle_copy_folder, "params": None},
            "cut folder": {"handler": self.handle_cut_folder, "params": None},
            "paste folder": {"handler": self.handle_paste_folder, "params": None},
            "rename folder or file": {"handler": self.handle_rename_folder_or_file, "params": None},
            "read text file": {"handler": self.handle_read_text_file, "params": None},
            "list contents": {"handler": self.handle_list_contents, "params": None},
            "get properties": {"handler": self.handle_get_properties, "params": None},
            "list commands": {"handler": self.handle_list_commands, "params": None},
            "exit": {"handler": self.handle_exit, "params": None},
            "increase volume": {"handler": self.handle_volume_up, "params": None},
            "decrease volume": {"handler": self.handle_volume_down, "params": None},
            "mute volume": {"handler": self.handle_mute_toggle, "params": None},
            "unmute volume": {"handler": self.handle_mute_toggle, "params": None},
            "maximize volume": {"handler": self.handle_maximize_volume, "params": None},
            "set volume": {"handler": self.handle_set_volume, "params": "number"},
            "increase brightness": {"handler": self.handle_brightness_up, "params": None},
            "decrease brightness": {"handler": self.handle_brightness_down, "params": None},
            "maximize brightness": {"handler": self.handle_maximize_brightness, "params": None},
            "set brightness": {"handler": self.handle_set_brightness, "params": "number"},
            "switch window": {"handler": self.handle_switch_window, "params": None},
            "minimize all windows": {"handler": self.handle_minimize_all_windows, "params": None},
            "restore windows": {"handler": self.handle_restore_all_windows, "params": None},
            "maximize window": {"handler": self.handle_maximize_current_window, "params": None},
            "minimize window": {"handler": self.handle_minimize_current_window, "params": None},
            "close window": {"handler": self.handle_close_current_window, "params": None},
            "move window left": {"handler": self.handle_move_window_left, "params": None},
            "move window right": {"handler": self.handle_move_window_right, "params": None},
            "take screenshot": {"handler": self.handle_take_screenshot, "params": None}
        }
        # Synonyms for natural language commands
        self.command_synonyms = {
            "create folder": ["make folder", "new folder", "add folder"],
            "open folder or file": ["open folder", "open file", "access folder", "access file"],
            "delete folder": ["remove folder", "delete directory"],
            "copy folder": ["duplicate folder"],
            "cut folder": ["move folder"],
            "paste folder": ["place folder"],
            "rename folder or file": ["rename folder", "rename file", "change name"],
            "read text file": ["read file", "open text", "read document"],
            "list contents": ["show contents", "list files", "display directory"],
            "get properties": ["show properties", "file info", "folder info"],
            "list commands": ["help", "commands", "what can you do"],
            "exit": ["quit", "stop program", "close"],
            "increase volume": ["volume up", "turn volume up", "louder"],
            "decrease volume": ["volume down", "turn volume down", "quieter"],
            "mute volume": ["mute", "silence"],
            "unmute volume": ["unmute"],
            "maximize volume": ["max volume", "full volume"],
            "set volume": ["set volume to", "turn volume to", "adjust volume to"],
            "increase brightness": ["brightness up", "turn brightness up", "brighter"],
            "decrease brightness": ["brightness down", "turn brightness down", "dimmer"],
            "maximize brightness": ["max brightness", "full brightness"],
            "set brightness": ["set brightness to", "turn brightness to", "adjust brightness to"],
            "switch window": ["next window", "change window"],
            "minimize all windows": ["show desktop", "minimize all"],
            "restore windows": ["restore all windows", "bring back windows"],
            "maximize window": ["maximize this window", "full screen"],
            "minimize window": ["minimize this window"],
            "close window": ["close this window", "close app"],
            "move window left": ["snap window left"],
            "move window right": ["snap window right"],
            "take screenshot": ["screenshot", "capture screen"]
        }
        # Context dictionary to store recent actions and file type
        self.context = {
            "last_created_folder": None,  # Stores (directory, folder_name)
            "last_opened_item": None,    # Stores (directory, name)
            "file_type": None            # Stores file type (e.g., '.txt', '.docx')
        }

    def get_command_list(self):
        """Return the list of available commands."""
        return list(self.commands.keys())

    def preprocess_command(self, cmd_text):
        """Preprocess the command text to remove polite phrases and normalize."""
        cmd_text = cmd_text.lower().strip()
        # Remove polite phrases
        polite_phrases = [
            r"can you please\s*", r"please\s*", r"could you\s*", r"would you\s*",
            r"kindly\s*", r"i want to\s*", r"i would like to\s*"
        ]
        for phrase in polite_phrases:
            cmd_text = re.sub(phrase, "", cmd_text)
        # Replace pronouns with context
        if " it " in cmd_text or cmd_text.endswith(" it"):
            if self.context["last_created_folder"]:
                directory, name = self.context["last_created_folder"]
                cmd_text = cmd_text.replace(" it", f" {name}")
            elif self.context["last_opened_item"]:
                directory, name = self.context["last_opened_item"]
                cmd_text = cmd_text.replace(" it", f" {name}")
        return cmd_text

    def extract_parameters(self, cmd_text, param_type):
        """Extract parameters (e.g., numbers) from the command text."""
        if param_type == "number":
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            return int(match.group(1)) if match else None
        return None

    def execute_command(self, cmd_text):
        """Execute commands with hybrid processing (regex, fuzzy matching, synonyms)."""
        cmd_text = self.preprocess_command(cmd_text)
        
        # Check for stop command first (highest priority)
        if "stop" in cmd_text or "cancel" in cmd_text:
            print("Stopping current operation and returning to main menu")
            self.file_manager.speech.speak("Stopping. Ready for new command.")
            return True

        # Split text on conjunctions like "and" or "then" to detect multiple commands
        command_parts = re.split(r'\s+(and|then)\s+', cmd_text)
        command_parts = [part for part in command_parts if part not in ["and", "then"]]
        executed = False

        # Process each command part
        for part in command_parts:
            part = part.strip()
            if not part:
                continue

            # Try exact matches
            for cmd, info in self.commands.items():
                if cmd in part:
                    print(f"Executing command: {cmd}")
                    params = self.extract_parameters(part, info["params"])
                    if info["params"] and params is None:
                        print(f"Missing required parameter for '{cmd}'")
                        self.file_manager.speech.speak(f"Missing required parameter for {cmd}")
                        continue
                    info["handler"](part if info["params"] else None)
                    executed = True
                    break

            # If no exact match, try synonyms
            if not executed:
                for cmd, info in self.commands.items():
                    for synonym in self.command_synonyms.get(cmd, []):
                        if synonym in part:
                            print(f"Executing synonym command: {cmd} (matched: {synonym})")
                            params = self.extract_parameters(part, info["params"])
                            if info["params"] and params is None:
                                print(f"Missing required parameter for '{cmd}'")
                                self.file_manager.speech.speak(f"Missing required parameter for {cmd}")
                                continue
                            info["handler"](part if info["params"] else None)
                            executed = True
                            break
                    if executed:
                        break

            # If no exact or synonym match, try fuzzy matching
            if not executed:
                for cmd, info in self.commands.items():
                    if fuzz.ratio(cmd, part) > 80:  # 80% similarity threshold
                        print(f"Executing fuzzy-matched command: {cmd} (input: {part})")
                        params = self.extract_parameters(part, info["params"])
                        if info["params"] and params is None:
                            print(f"Missing required parameter for '{cmd}'")
                            self.file_manager.speech.speak(f"Missing required parameter for {cmd}")
                            continue
                        info["handler"](part if info["params"] else None)
                        executed = True
                        break
                    # Check synonyms with fuzzy matching
                    for synonym in self.command_synonyms.get(cmd, []):
                        if fuzz.ratio(synonym, part) > 80:
                            print(f"Executing fuzzy-matched synonym command: {cmd} (matched: {synonym}, input: {part})")
                            params = self.extract_parameters(part, info["params"])
                            if info["params"] and params is None:
                                print(f"Missing required parameter for '{cmd}'")
                                self.file_manager.speech.speak(f"Missing required parameter for {cmd}")
                                continue
                            info["handler"](part if info["params"] else None)
                            executed = True
                            break
                    if executed:
                        break

        return executed

    def handle_create_folder(self, cmd_text=None):
        """Handle the 'create folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            while True:
                folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to create: ", self.voice_recognizer)
                result = self.file_manager.create_folder(directory, folder_name)
                if isinstance(result, tuple):
                    directory, folder_name = result
                else:
                    # Update context with created folder
                    self.context["last_created_folder"] = (directory, folder_name)
                    self.context["last_opened_item"] = None
                    break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_open_folder_or_file(self, cmd_text=None):
        """Handle the 'open folder or file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            
            # Ask if user wants to list directory contents
            print("Do you want to list the files and folders in this directory? (yes/no)")
            self.file_manager.speech.speak("Do you want to list the files and folders in this directory? Say yes or no.")
            
            # Wait for response
            wait_time = 0
            while wait_time < 5:  # Wait up to 5 seconds for response
                if self.voice_recognizer:
                    response = self.voice_recognizer.get_transcription()
                    if response:
                        print(f"You said: {response}")
                        if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                            # List directory contents
                            self.file_manager.list_contents(directory)
                            break
                        elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                            print("Continuing without listing directory contents.")
                            self.file_manager.speech.speak("Continuing without listing directory contents.")
                            break
                time.sleep(0.5)
                wait_time += 0.5
            
            # If no response in time
            if wait_time >= 5:
                print("No response. Continuing without listing directory contents.")
                self.file_manager.speech.speak("Continuing without listing directory contents.")
            
            # Continue with file selection in a loop until successful or user quits
            while True:
                name = self.file_manager.get_folder_or_file_name("Enter folder or file name to open: ", self.voice_recognizer)
                
                # Check if it's a file that needs a file type
                if os.path.isfile(os.path.join(directory, name)):
                    while True:
                        file_type = input("Enter file type/extension (e.g., .txt, .docx, .py, .js) or press Enter for any, 'quit' to return to main menu: ").strip()
                        if file_type.lower() == 'quit':
                            raise self.file_manager.ReturnToMain()
                        if not file_type:
                            break
                        if file_type.startswith("."):
                            break
                        print("Please include the dot (e.g., .txt).")
                        self.file_manager.speech.speak("Please include the dot, for example, dot text.")
                    
                    result = self.file_manager.open_folder_or_file(directory, name, file_type)
                    if result == True:
                        # Update context with opened item
                        self.context["last_opened_item"] = (directory, name)
                        self.context["last_created_folder"] = None
                        break
                    elif result == False:
                        break  # Error that's not 'not found'
                else:
                    # Try to open as folder or file without type
                    result = self.file_manager.open_folder_or_file(directory, name)
                    
                    # Check the result
                    if isinstance(result, tuple) and result[0] == "not_found":
                        # Item not found, ask again
                        print("Please try a different name.")
                        self.file_manager.speech.speak("Please try a different name.")
                        # Continue the loop to ask again
                    else:
                        if result == True:
                            # Update context with opened item
                            self.context["last_opened_item"] = (directory, name)
                            self.context["last_created_folder"] = None
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_delete_folder(self, cmd_text=None):
        """Handle the 'delete folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            
            # Ask if user wants to list directory contents
            print("Do you want to list the files and folders in this directory? (yes/no)")
            self.file_manager.speech.speak("Do you want to list the files and folders in this directory? Say yes or no.")
            
            # Wait for response
            wait_time = 0
            while wait_time < 5:  # Wait up to 5 seconds for response
                if self.voice_recognizer:
                    response = self.voice_recognizer.get_transcription()
                    if response:
                        print(f"You said: {response}")
                        if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                            # List directory contents
                            self.file_manager.list_contents(directory)
                            break
                        elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                            print("Continuing without listing directory contents.")
                            self.file_manager.speech.speak("Continuing without listing directory contents.")
                            break
                time.sleep(0.5)
                wait_time += 0.5
            
            # If no response in time
            if wait_time >= 5:
                print("No response. Continuing without listing directory contents.")
                self.file_manager.speech.speak("Continuing without listing directory contents.")
            
            # Loop until successful deletion or user quits
            while True:
                folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to delete: ", self.voice_recognizer)
                result = self.file_manager.delete_folder(directory, folder_name)
                
                # If deletion was successful or there was an error other than 'not found'
                if result == True:
                    # Clear context if deleted item was in context
                    if self.context["last_created_folder"] and self.context["last_created_folder"][1] == folder_name:
                        self.context["last_created_folder"] = None
                    if self.context["last_opened_item"] and self.context["last_opened_item"][1] == folder_name:
                        self.context["last_opened_item"] = None
                    break
                else:
                    # Check if the folder exists
                    folder_path = os.path.join(directory, folder_name)
                    if not os.path.isdir(folder_path):
                        # Folder not found, ask again
                        print("Please try a different folder name.")
                        self.file_manager.speech.speak("Please try a different folder name.")
                        # Continue the loop to ask again
                    else:
                        # Some other error occurred
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_copy_folder(self, cmd_text=None):
        """Handle the 'copy folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to copy: ", self.voice_recognizer)
            result = self.file_manager.copy_folder(directory, folder_name)
            if result:
                # Update context with copied folder
                self.context["last_opened_item"] = (directory, folder_name)
                self.context["last_created_folder"] = None
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_cut_folder(self, cmd_text=None):
        """Handle the 'cut folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to cut: ", self.voice_recognizer)
            result = self.file_manager.cut_folder(directory, folder_name)
            if result:
                # Update context with cut folder
                self.context["last_opened_item"] = (directory, folder_name)
                self.context["last_created_folder"] = None
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_paste_folder(self, cmd_text=None):
        """Handle the 'paste folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            result = self.file_manager.paste_folder(directory)
            if result and self.file_manager.clipboard:
                # Update context with pasted folder
                folder_name = self.file_manager.clipboard["name"]
                self.context["last_created_folder"] = (directory, folder_name)
                self.context["last_opened_item"] = None
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_rename_folder_or_file(self, cmd_text=None):
        """Handle the 'rename folder or file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            old_name = self.file_manager.get_folder_or_file_name("Enter current folder/file name to rename: ", self.voice_recognizer)
            new_name = self.file_manager.get_folder_or_file_name("Enter new folder/file name: ", self.voice_recognizer)
            result = self.file_manager.rename_folder_or_file(directory, old_name, new_name)
            if result:
                # Update context if renamed item was in context
                if self.context["last_created_folder"] and self.context["last_created_folder"][1] == old_name:
                    self.context["last_created_folder"] = (directory, new_name)
                if self.context["last_opened_item"] and self.context["last_opened_item"][1] == old_name:
                    self.context["last_opened_item"] = (directory, new_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_read_text_file(self, cmd_text=None):
        """Handle the 'read text file' command with file type specification."""
        try:
            # Prompt for file type (text or Word)
            print("Is this a text file or Word file? (say 'text' or 'Word')")
            self.file_manager.speech.speak("Is this a text file or Word file? Say text or Word.")
            wait_time = 0
            file_type = None
            while wait_time < 5:  # Wait up to 5 seconds for response
                if self.voice_recognizer:
                    response = self.voice_recognizer.get_transcription()
                    if response:
                        print(f"You said: {response}")
                        if "text" in response.lower():
                            file_type = ".txt"
                            break
                        elif "word" in response.lower():
                            file_type = ".docx"
                            break
                        else:
                            print("Please say 'text' or 'Word'.")
                            self.file_manager.speech.speak("Please say text or Word.")
                            wait_time = 0  # Reset wait time on invalid response
                    time.sleep(0.5)
                    wait_time += 0.5
            
            if not file_type:
                print("No valid file type provided. Defaulting to text file.")
                self.file_manager.speech.speak("No valid file type provided. Defaulting to text file.")
                file_type = ".txt"

            # Store file type in context
            self.context["file_type"] = file_type

            directory = self.file_manager.select_directory(self.voice_recognizer)
            
            # Ask if user wants to list directory contents (filtered by file type)
            print(f"Do you want to list the {file_type} files in this directory? (yes/no)")
            self.file_manager.speech.speak(f"Do you want to list the {file_type} files in this directory? Say yes or no.")
            
            # Wait for response
            wait_time = 0
            while wait_time < 5:  # Wait up to 5 seconds for response
                if self.voice_recognizer:
                    response = self.voice_recognizer.get_transcription()
                    if response:
                        print(f"You said: {response}")
                        if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                            # List directory contents filtered by file type
                            self.file_manager.list_contents(directory, extension=file_type)
                            break
                        elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                            print("Continuing without listing directory contents.")
                            self.file_manager.speech.speak("Continuing without listing directory contents.")
                            break
                time.sleep(0.5)
                wait_time += 0.5
            
            # If no response in time
            if wait_time >= 5:
                print("No response. Continuing without listing directory contents.")
                self.file_manager.speech.speak("Continuing without listing directory contents.")
            
            # Continue with file selection in a loop until successful or user quits
            while True:
                name = self.file_manager.get_folder_or_file_name("Enter file name to read: ", self.voice_recognizer)
                
                # Try to read the file with the specified file type
                result = self.file_manager.read_text_file(directory, name, file_type)
                
                # Handle results
                if result == True:
                    # Update context with read file
                    self.context["last_opened_item"] = (directory, name)
                    self.context["last_created_folder"] = None
                    break
                elif isinstance(result, tuple) and result[0] == "ambiguous":
                    # Multiple files found, prompt for clarification
                    possible_files = result[1]
                    print(f"Multiple files found: {', '.join(possible_files)}")
                    self.file_manager.speech.speak(f"Multiple files found. Please say text or Word to clarify.")
                    wait_time = 0
                    while wait_time < 5:
                        if self.voice_recognizer:
                            response = self.voice_recognizer.get_transcription()
                            if response:
                                print(f"You said: {response}")
                                if "text" in response.lower():
                                    file_type = ".txt"
                                    break
                                elif "word" in response.lower():
                                    file_type = ".docx"
                                    break
                                else:
                                    print("Please say 'text' or 'Word'.")
                                    self.file_manager.speech.speak("Please say text or Word.")
                                    wait_time = 0
                            time.sleep(0.5)
                            wait_time += 0.5
                    
                    if not file_type:
                        print("No valid file type provided. Please try again.")
                        self.file_manager.speech.speak("No valid file type provided. Please try again.")
                        continue
                    
                    # Update context and retry with clarified file type
                    self.context["file_type"] = file_type
                    result = self.file_manager.read_text_file(directory, name, file_type)
                    if result == True:
                        self.context["last_opened_item"] = (directory, name)
                        self.context["last_created_folder"] = None
                        break
                    elif result == False:
                        continue  # Error, ask for another name
                else:
                    # File not found or other error, ask again
                    continue
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_list_contents(self, cmd_text=None):
        """Handle the 'list contents' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            self.file_manager.list_contents(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_get_properties(self, cmd_text=None):
        """Handle the 'get properties' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to get properties: ", self.voice_recognizer)
            self.file_manager.get_properties(directory, name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_list_commands(self, cmd_text=None):
        """Handle the 'list commands' command - lists all available commands."""
        print("Available commands are:")
        self.file_manager.speech.speak("Available commands are:")
        for cmd in self.get_command_list():
            print(f"  • {cmd}")
            self.file_manager.speech.speak(cmd)
        print("You can say 'exit' to quit the program.")
        self.file_manager.speech.speak("You can say exit to quit the program.")

    def handle_exit(self, cmd_text=None):
        """Handle the 'exit' command."""
        print("Exiting the program...")
        self.file_manager.speech.speak("Exiting the program")
        sys.exit(0)

    def handle_volume_up(self, cmd_text=None):
        """Handle the 'increase volume' command."""
        self.os_manager.volume_up()

    def handle_volume_down(self, cmd_text=None):
        """Handle the 'decrease volume' command."""
        self.os_manager.volume_down()

    def handle_mute_toggle(self, cmd_text=None):
        """Handle the 'mute/unmute volume' command."""
        self.os_manager.mute_toggle()

    def handle_maximize_volume(self, cmd_text=None):
        """Handle the 'maximize volume' command."""
        self.os_manager.maximize_volume()

    def handle_set_volume(self, cmd_text):
        """Handle the 'set volume' command."""
        try:
            level = self.extract_parameters(cmd_text, "number")
            if level is not None:
                self.os_manager.set_volume(level)
            else:
                print("No valid volume level found. Please say a number between 0 and 100.")
                self.os_manager.speech.speak("No valid volume level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting volume: {e}")
            self.os_manager.speech.speak("Error setting volume.")

    def handle_brightness_up(self, cmd_text=None):
        """Handle the 'increase brightness' command."""
        self.os_manager.brightness_up()

    def handle_brightness_down(self, cmd_text=None):
        """Handle the 'decrease brightness' command."""
        self.os_manager.brightness_down()

    def handle_maximize_brightness(self, cmd_text=None):
        """Handle the 'maximize brightness' command."""
        self.os_manager.maximize_brightness()

    def handle_set_brightness(self, cmd_text):
        """Handle the 'set brightness' command."""
        try:
            level = self.extract_parameters(cmd_text, "number")
            if level is not None:
                self.os_manager.set_brightness(level)
            else:
                print("No valid brightness level found. Please say a number between 0 and 100.")
                self.os_manager.speech.speak("No valid brightness level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting brightness: {e}")
            self.os_manager.speech.speak("Error setting brightness.")

    def handle_switch_window(self, cmd_text=None):
        """Handle the 'switch window' command."""
        self.os_manager.switch_window()

    def handle_minimize_all_windows(self, cmd_text=None):
        """Handle the 'minimize all windows' command."""
        self.os_manager.minimize_all_windows()

    def handle_restore_all_windows(self, cmd_text=None):
        """Handle the 'restore windows' command."""
        self.os_manager.restore_all_windows()

    def handle_maximize_current_window(self, cmd_text=None):
        """Handle the 'maximize window' command."""
        self.os_manager.maximize_current_window()

    def handle_minimize_current_window(self, cmd_text=None):
        """Handle the 'minimize window' command."""
        self.os_manager.minimize_current_window()

    def handle_close_current_window(self, cmd_text=None):
        """Handle the 'close window' command."""
        self.os_manager.close_current_window()

    def handle_move_window_left(self, cmd_text=None):
        """Handle the 'move window left' command."""
        self.os_manager.move_window_left()

    def handle_move_window_right(self, cmd_text=None):
        """Handle the 'move window right' command."""
        self.os_manager.move_window_right()

    def handle_take_screenshot(self, cmd_text=None):
        """Handle the 'take screenshot' command."""
        self.os_manager.take_screenshot()