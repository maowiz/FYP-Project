import sys
import os
import time

class CommandHandler:
    def __init__(self, file_manager, voice_recognizer=None):
        self.file_manager = file_manager
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
            "exit": self.handle_exit
        }

    def get_command_list(self):
        """Return the list of available commands."""
        return list(self.commands.keys())

    def execute_command(self, cmd_text):
        """Execute the command based on the transcribed text.
        Now with improved partial command matching."""
        cmd_text = cmd_text.lower()
        
        # Check for stop command first (highest priority)
        if "stop" in cmd_text.lower():
            print("Stopping current operation and returning to main menu")
            self.file_manager.speak("Stopping. Ready for new command.")
            return True
            
        # First try exact matches
        for cmd, handler in self.commands.items():
            if cmd in cmd_text:
                print(f"Executing command: {cmd}")
                handler()
                return True
        
        # If no exact match, try partial matching for common commands
        partial_matches = {
            "open folder": "open folder or file",
            "open file": "open folder or file",
            "rename folder": "rename folder or file",
            "rename file": "rename folder or file",
            "help": "list commands",
            "commands": "list commands",
            "what can you do": "list commands"
        }
        
        for partial, full_cmd in partial_matches.items():
            if partial in cmd_text:
                print(f"Recognized partial command '{partial}' as '{full_cmd}'")
                self.commands[full_cmd]()
                return True
                
        return False

    def handle_create_folder(self):
        """Handle the 'create folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            while True:
                folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to create: ", self.voice_recognizer)
                result = self.file_manager.create_folder(directory, folder_name)
                if isinstance(result, tuple):
                    directory, folder_name = result
                else:
                    break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_open_folder_or_file(self):
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
                    if result == True or result == False:
                        break  # Successfully opened or error that's not 'not found'
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
                        # Either success or some other error
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_delete_folder(self):
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
                    # Successfully deleted
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

    def handle_copy_folder(self):
        """Handle the 'copy folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to copy: ", self.voice_recognizer)
            self.file_manager.copy_folder(directory, folder_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_cut_folder(self):
        """Handle the 'cut folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to cut: ", self.voice_recognizer)
            self.file_manager.cut_folder(directory, folder_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_paste_folder(self):
        """Handle the 'paste folder' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            self.file_manager.paste_folder(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_rename_folder_or_file(self):
        """Handle the 'rename folder or file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            old_name = self.file_manager.get_folder_or_file_name("Enter current folder/file name to rename: ", self.voice_recognizer)
            new_name = self.file_manager.get_folder_or_file_name("Enter new folder/file name: ", self.voice_recognizer)
            self.file_manager.rename_folder_or_file(directory, old_name, new_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_read_text_file(self):
        """Handle the 'read text file' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            file_name = self.file_manager.get_folder_or_file_name("Enter file name to read: ", self.voice_recognizer)
            self.file_manager.read_text_file(directory, file_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_list_contents(self):
        """Handle the 'list contents' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            self.file_manager.list_contents(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_get_properties(self):
        """Handle the 'get properties' command."""
        try:
            directory = self.file_manager.select_directory(self.voice_recognizer)
            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to get properties: ", self.voice_recognizer)
            self.file_manager.get_properties(directory, name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_list_commands(self):
        """Handle the 'list commands' command - lists all available commands."""
        print("Available commands are:")
        self.file_manager.speak("Available commands are:")
        for cmd in self.get_command_list():
            print(f"  • {cmd}")
            self.file_manager.speak(cmd)
        print("You can say 'exit' to quit the program.")
        self.file_manager.speak("You can say exit to quit the program.")

    def handle_exit(self):
        """Handle the 'exit' command."""
        print("Exiting the program...")
        self.file_manager.speak("Exiting the program")
        sys.exit(0)