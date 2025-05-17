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
        self.commands = {
            "create folder": self.handle_create_folder,
            "open folder or file": self.handle_open_folder_or_file,
            "delete folder": self.handle_delete_folder,
            "copy folder": self.handle_copy_folder,
            "cut folder": self.handle_cut_folder,
            "paste folder": self.handle_paste_folder,
            "rename folder or file": self.handle_rename_folder_or_file,
            "read text file": self.handle_read_text_file,
            "list contents": self.handle_list_contents,
            "get properties": self.handle_get_properties,
            "list commands": self.handle_list_commands,
            "exit": self.handle_exit,
            "increase volume": self.handle_volume_up,
            "decrease volume": self.handle_volume_down,
            "mute volume": self.handle_mute_toggle,
            "unmute volume": self.handle_mute_toggle,
            "maximize volume": self.handle_maximize_volume,
            "set volume": self.handle_set_volume,
            "increase brightness": self.handle_brightness_up,
            "decrease brightness": self.handle_brightness_down,
            "maximize brightness": self.handle_maximize_brightness,
            "set brightness": self.handle_set_brightness,
            "switch window": self.handle_switch_window,
            "minimize all windows": self.handle_minimize_all_windows,
            "restore windows": self.handle_restore_all_windows,
            "maximize window": self.handle_maximize_current_window,
            "minimize window": self.handle_minimize_current_window,
            "close window": self.handle_close_current_window,
            "move window left": self.handle_move_window_left,
            "move window right": self.handle_move_window_right,
            "take screenshot": self.handle_take_screenshot
        }
        # Context dictionary to store recent actions
        self.context = {
            "last_created_folder": None,  # Stores (directory, folder_name)
            "last_opened_item": None,    # Stores (directory, name)
        }

    def get_command_list(self):
        """Return the list of available commands."""
        return list(self.commands.keys())

    def execute_command(self, cmd_text):
        """Execute commands with fuzzy matching, multiple commands, and context awareness."""
        cmd_text = cmd_text.lower()
        
        # Check for stop command first (highest priority)
        if "stop" in cmd_text:
            print("Stopping current operation and returning to main menu")
            self.file_manager.speak("Stopping. Ready for new command.")
            return True

        # Handle pronouns like "it" by replacing with last created/opened item
        if " it " in cmd_text or cmd_text.endswith(" it"):
            if self.context["last_created_folder"]:
                directory, name = self.context["last_created_folder"]
                cmd_text = cmd_text.replace(" it", f" {name}")
            elif self.context["last_opened_item"]:
                directory, name = self.context["last_opened_item"]
                cmd_text = cmd_text.replace(" it", f" {name}")

        # Split text on conjunctions like "and" to detect multiple commands
        command_parts = re.split(r'\s+and\s+', cmd_text)
        executed = False

        # Process each command part
        for part in command_parts:
            part = part.strip()
            if not part:
                continue

            # Try exact matches
            for cmd, handler in self.commands.items():
                if cmd in part:
                    print(f"Executing command: {cmd}")
                    handler(part if cmd in ["set volume", "set brightness"] else None)
                    executed = True
                    break  # Move to next command part after executing one

            # If no exact match, try fuzzy matching
            if not executed:
                for cmd, handler in self.commands.items():
                    if fuzz.ratio(cmd, part) > 80:  # 80% similarity threshold
                        print(f"Executing fuzzy-matched command: {cmd} (input: {part})")
                        handler(part if cmd in ["set volume", "set brightness"] else None)
                        executed = True
                        break

            # If no exact or fuzzy match, try partial matching
            if not executed:
                partial_matches = {
                    "open folder": "open folder or file",
                    "open file": "open folder or file",
                    "rename folder": "rename folder or file",
                    "rename file": "rename folder or file",
                    "help": "list commands",
                    "commands": "list commands",
                    "what can you do": "list commands",
                    "turn volume up": "increase volume",
                    "turn volume down": "decrease volume",
                    "mute": "mute volume",
                    "unmute": "unmute volume",
                    "max volume": "maximize volume",
                    "turn brightness up": "increase brightness",
                    "turn brightness down": "decrease brightness",
                    "max brightness": "maximize brightness",
                    "next window": "switch window",
                    "show desktop": "minimize all windows",
                    "restore all windows": "restore windows",
                    "maximize this window": "maximize window",
                    "minimize this window": "minimize window",
                    "close this app": "close window",
                    "screenshot": "take screenshot"
                }
                
                for partial, full_cmd in partial_matches.items():
                    if partial in part:
                        print(f"Recognized partial command '{partial}' as '{full_cmd}'")
                        self.commands[full_cmd](part if full_cmd in ["set volume", "set brightness"] else None)
                        executed = True
                        break
                    # Try fuzzy matching for partial commands
                    if fuzz.ratio(partial, part) > 80:
                        print(f"Recognized fuzzy-matched partial command '{partial}' as '{full_cmd}' (input: {part})")
                        self.commands[full_cmd](part if full_cmd in ["set volume", "set brightness"] else None)
                        executed = True
                        break

        # Return True if at least one command was executed
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
            self.file_manager.speak("Returning to main menu.")

    def handle_open_folder_or_file(self, cmd_text=None):
        """Handle the 'open folder or file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            
            # Ask if user wants to list directory contents
            print("Do you want to list the files and folders in this directory? (yes/no)")
            self.file_manager.speak("Do you want to list the files and folders in this directory? Say yes or no.")
            
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
                            self.file_manager.speak("Continuing without listing directory contents.")
                            break
                time.sleep(0.5)
                wait_time += 0.5
            
            # If no response in time
            if wait_time >= 5:
                print("No response. Continuing without listing directory contents.")
                self.file_manager.speak("Continuing without listing directory contents.")
            
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
                        self.file_manager.speak("Please include the dot, for example, dot text.")
                    
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
                        self.file_manager.speak("Please try a different name.")
                        # Continue the loop to ask again
                    else:
                        if result == True:
                            # Update context with opened item
                            self.context["last_opened_item"] = (directory, name)
                            self.context["last_created_folder"] = None
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_delete_folder(self, cmd_text=None):
        """Handle the 'delete folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            
            # Ask if user wants to list directory contents
            print("Do you want to list the files and folders in this directory? (yes/no)")
            self.file_manager.speak("Do you want to list the files and folders in this directory? Say yes or no.")
            
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
                            self.file_manager.speak("Continuing without listing directory contents.")
                            break
                time.sleep(0.5)
                wait_time += 0.5
            
            # If no response in time
            if wait_time >= 5:
                print("No response. Continuing without listing directory contents.")
                self.file_manager.speak("Continuing without listing directory contents.")
            
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
                        self.file_manager.speak("Please try a different folder name.")
                        # Continue the loop to ask again
                    else:
                        # Some other error occurred
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

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
            self.file_manager.speak("Returning to main menu.")

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
            self.file_manager.speak("Returning to main menu.")

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
            self.file_manager.speak("Returning to main menu.")

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
            self.file_manager.speak("Returning to main menu.")

    def handle_read_text_file(self, cmd_text=None):
        """Handle the 'read text file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            file_name = self.file_manager.get_folder_or_file_name("Enter file name to read: ", self.voice_recognizer)
            result = self.file_manager.read_text_file(directory, file_name)
            if result:
                # Update context with read file
                self.context["last_opened_item"] = (directory, file_name)
                self.context["last_created_folder"] = None
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_list_contents(self, cmd_text=None):
        """Handle the 'list contents' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            self.file_manager.list_contents(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_get_properties(self, cmd_text=None):
        """Handle the 'get properties' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to get properties: ", self.voice_recognizer)
            self.file_manager.get_properties(directory, name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_list_commands(self, cmd_text=None):
        """Handle the 'list commands' command - lists all available commands."""
        print("Available commands are:")
        self.file_manager.speak("Available commands are:")
        for cmd in self.get_command_list():
            print(f"  • {cmd}")
            self.file_manager.speak(cmd)
        print("You can say 'exit' to quit the program.")
        self.file_manager.speak("You can say exit to quit the program.")

    def handle_exit(self, cmd_text=None):
        """Handle the 'exit' command."""
        print("Exiting the program...")
        self.file_manager.speak("Exiting the program")
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
            # Extract number from command (e.g., "set volume to 50" -> 50)
            match = re.search(r'\d+', cmd_text)
            if match:
                level = int(match.group())
                self.os_manager.set_volume(level)
            else:
                print("No valid volume level found. Please say a number between 0 and 100.")
                self.file_manager.speak("No valid volume level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting volume: {e}")
            self.file_manager.speak("Error setting volume.")

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
            # Extract number from command (e.g., "set brightness to 60" -> 60)
            match = re.search(r'\d+', cmd_text)
            if match:
                level = int(match.group())
                self.os_manager.set_brightness(level)
            else:
                print("No valid brightness level found. Please say a number between 0 and 100.")
                self.file_manager.speak("No valid brightness level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting brightness: {e}")
            self.file_manager.speak("Error setting brightness.")

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