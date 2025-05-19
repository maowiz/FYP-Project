import os
import time

class FileCommandHandler:
    def __init__(self, file_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.voice_recognizer = voice_recognizer

    def handle_create_folder(self, cmd_text=None, context=None):
        """Handle the 'create folder' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to create: ", self.voice_recognizer)
            result = self.file_manager.create_folder(directory, folder_name)
            if isinstance(result, tuple):
                directory, folder_name = result
                # Update context with created folder
                if context:
                    context["last_created_folder"] = (directory, folder_name)
                    context["last_opened_item"] = None
                    context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_open_folder_or_file(self, cmd_text=None, context=None):
        """Handle the 'open folder or file' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            # Ask if user wants to list directory contents (skip if directory was auto-selected)
            if not (context and context.get("working_directory")):
                print("Do you want to list the files and folders in this directory? (yes/no)")
                self.file_manager.speech.speak("Do you want to list the files and folders in this directory? Say yes or no.")
                
                # Wait for response
                wait_time = 0
                while wait_time < 5:  # Wait up to 5 seconds
                    if self.voice_recognizer:
                        response = self.voice_recognizer.get_transcription()
                        if response:
                            print(f"You said: {response}")
                            if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                                self.file_manager.list_contents(directory)
                                break
                            elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                                print("Continuing without listing directory contents.")
                                self.file_manager.speech.speak("Continuing without listing directory contents.")
                                break
                    time.sleep(0.5)
                    wait_time += 0.5
                
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
                        if not file_type or file_type.startswith("."):
                            break
                        print("Please include the dot (e.g., .txt).")
                        self.file_manager.speech.speak("Please include the dot, for example, dot text.")
                    
                    result = self.file_manager.open_folder_or_file(directory, name, file_type)
                    if result == True:
                        # Update context with opened item
                        if context:
                            context["last_opened_item"] = (directory, name)
                            context["last_created_folder"] = None
                            context["working_directory"] = directory
                        break
                    elif result == False:
                        break
                else:
                    # Try to open as folder or file without type
                    result = self.file_manager.open_folder_or_file(directory, name)
                    
                    # Check the result
                    if isinstance(result, tuple) and result[0] == "not_found":
                        print("Please try a different name.")
                        self.file_manager.speech.speak("Please try a different name.")
                    else:
                        if result == True:
                            # Update context with opened item
                            if context:
                                context["last_opened_item"] = (directory, name)
                                context["last_created_folder"] = None
                                context["working_directory"] = directory
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_delete_folder(self, cmd_text=None, context=None):
        """Handle the 'delete folder' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            # Ask if user wants to list directory contents (skip if directory was auto-selected)
            if not (context and context.get("working_directory")):
                print("Do you want to list the files and folders in this directory? (yes/no)")
                self.file_manager.speech.speak("Do you want to list the files and folders in this directory? Say yes or no.")
                
                # Wait for response
                wait_time = 0
                while wait_time < 5:  # Wait up to 5 seconds
                    if self.voice_recognizer:
                        response = self.voice_recognizer.get_transcription()
                        if response:
                            print(f"You said: {response}")
                            if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                                self.file_manager.list_contents(directory)
                                break
                            elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                                print("Continuing without listing directory contents.")
                                self.file_manager.speech.speak("Continuing without listing directory contents.")
                                break
                    time.sleep(0.5)
                    wait_time += 0.5
                
                if wait_time >= 5:
                    print("No response. Continuing without listing directory contents.")
                    self.file_manager.speech.speak("Continuing without listing directory contents.")

            # Loop until successful deletion or user quits
            while True:
                folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to delete: ", self.voice_recognizer)
                result = self.file_manager.delete_folder(directory, folder_name)
                
                # If deletion was successful
                if result == True:
                    # Clear context if deleted item was in context
                    if context and context.get("last_created_folder") and context["last_created_folder"][1] == folder_name:
                        context["last_created_folder"] = None
                    if context and context.get("last_opened_item") and context["last_opened_item"][1] == folder_name:
                        context["last_opened_item"] = None
                    context["working_directory"] = directory
                    break
                else:
                    # Check if the folder exists
                    folder_path = os.path.join(directory, folder_name)
                    if not os.path.isdir(folder_path):
                        print("Please try a different folder name.")
                        self.file_manager.speech.speak("Please try a different folder name.")
                    else:
                        break
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_copy_folder(self, cmd_text=None, context=None):
        """Handle the 'copy folder' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to copy: ", self.voice_recognizer)
            result = self.file_manager.copy_folder(directory, folder_name)
            if result:
                # Update context with copied folder
                if context:
                    context["last_opened_item"] = (directory, folder_name)
                    context["last_created_folder"] = None
                    context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_cut_folder(self, cmd_text=None, context=None):
        """Handle the 'cut folder' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to cut: ", self.voice_recognizer)
            result = self.file_manager.cut_folder(directory, folder_name)
            if result:
                # Update context with cut folder
                if context:
                    context["last_opened_item"] = (directory, folder_name)
                    context["last_created_folder"] = None
                    context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_paste_folder(self, cmd_text=None, context=None):
        """Handle the 'paste folder' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            result = self.file_manager.paste_folder(directory)
            if result and self.file_manager.clipboard:
                # Update context with pasted folder
                folder_name = self.file_manager.clipboard["name"]
                if context:
                    context["last_created_folder"] = (directory, folder_name)
                    context["last_opened_item"] = None
                    context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_rename_folder_or_file(self, cmd_text=None, context=None):
        """Handle the 'rename folder or file' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            old_name = self.file_manager.get_folder_or_file_name("Enter current folder/file name to rename: ", self.voice_recognizer)
            new_name = self.file_manager.get_folder_or_file_name("Enter new folder/file name: ", self.voice_recognizer)
            result = self.file_manager.rename_folder_or_file(directory, old_name, new_name)
            if result:
                # Update context if renamed item was in context
                if context and context.get("last_created_folder") and context["last_created_folder"][1] == old_name:
                    context["last_created_folder"] = (directory, new_name)
                if context and context.get("last_opened_item") and context["last_opened_item"][1] == old_name:
                    context["last_opened_item"] = (directory, new_name)
                context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_read_text_file(self, cmd_text=None, context=None):
        """Handle the 'read text file' command with file type specification."""
        try:
            # Prompt for file type (text or Word)
            print("Is this a text file or Word file? (say 'text' or 'Word')")
            self.file_manager.speech.speak("Is this a text file or Word file? Say text or Word.")
            wait_time = 0
            file_type = None
            while wait_time < 5:  # Wait up to 5 seconds
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
                print("No valid file type provided. Defaulting to text file.")
                self.file_manager.speech.speak("No valid file type provided. Defaulting to text file.")
                file_type = ".txt"

            # Store file type in context
            if context:
                context["file_type"] = file_type

            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            # Ask if user wants to list directory contents (skip if directory was auto-selected)
            if not (context and context.get("working_directory")):
                print(f"Do you want to list the {file_type} files in this directory? (yes/no)")
                self.file_manager.speech.speak(f"Do you want to list the {file_type} files in this directory? Say yes or no.")
                
                # Wait for response
                wait_time = 0
                while wait_time < 5:
                    if self.voice_recognizer:
                        response = self.voice_recognizer.get_transcription()
                        if response:
                            print(f"You said: {response}")
                            if "yes" in response.lower() or "yeah" in response.lower() or "sure" in response.lower():
                                self.file_manager.list_contents(directory, extension=file_type)
                                break
                            elif "no" in response.lower() or "nope" in response.lower() or "stop" in response.lower():
                                print("Continuing without listing directory contents.")
                                self.file_manager.speech.speak("Continuing without listing directory contents.")
                                break
                        time.sleep(0.5)
                        wait_time += 0.5
                
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
                    if context:
                        context["last_opened_item"] = (directory, name)
                        context["last_created_folder"] = None
                        context["working_directory"] = directory
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
                        if context:
                            context["file_type"] = file_type
                        result = self.file_manager.read_text_file(directory, name, file_type)
                        if result == True:
                            if context:
                                context["last_opened_item"] = (directory, name)
                                context["last_created_folder"] = None
                                context["working_directory"] = directory
                            break
                        elif result == False:
                            continue
                else:
                    continue
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_list_contents(self, cmd_text=None, context=None):
        """Handle the 'list contents' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            self.file_manager.list_contents(directory)
            if context:
                context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")

    def handle_get_properties(self, cmd_text=None, context=None):
        """Handle the 'get properties' command."""
        try:
            # Use working_directory from context if available and valid
            directory = None
            if context and context.get("working_directory") and os.path.isdir(context["working_directory"]):
                directory = context["working_directory"]
                last_folder = os.path.basename(directory)
                print(f"Using current directory: {directory}")
                self.file_manager.speech.speak(f"Using directory {last_folder}")
            else:
                directory = self.file_manager.select_directory(self.voice_recognizer)

            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to get properties: ", self.voice_recognizer)
            self.file_manager.get_properties(directory, name)
            if context:
                context["working_directory"] = directory
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speech.speak("Returning to main menu.")