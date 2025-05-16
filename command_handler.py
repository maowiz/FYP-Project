import sys

class CommandHandler:
    def __init__(self, file_manager):
        self.file_manager = file_manager
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
            "exit": self.handle_exit
        }

    def get_command_list(self):
        """Return the list of available commands."""
        return list(self.commands.keys())

    def execute_command(self, cmd_text):
        """Execute the command based on the transcribed text."""
        for cmd, handler in self.commands.items():
            if cmd in cmd_text:
                print(f"Executing command: {cmd}")
                handler()
                return True
        return False

    def handle_create_folder(self):
        """Handle the 'create folder' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            while True:
                folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to create: ")
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
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to open: ")
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
                self.file_manager.open_folder_or_file(directory, name, file_type)
            else:
                self.file_manager.open_folder_or_file(directory, name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_delete_folder(self):
        """Handle the 'delete folder' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to delete: ")
            self.file_manager.delete_folder(directory, folder_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_copy_folder(self):
        """Handle the 'copy folder' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to copy: ")
            self.file_manager.copy_folder(directory, folder_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_cut_folder(self):
        """Handle the 'cut folder' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            folder_name = self.file_manager.get_folder_or_file_name("Enter folder name to cut: ")
            self.file_manager.cut_folder(directory, folder_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_paste_folder(self):
        """Handle the 'paste folder' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            self.file_manager.paste_folder(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_rename_folder_or_file(self):
        """Handle the 'rename folder or file' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            old_name = self.file_manager.get_folder_or_file_name("Enter current folder/file name to rename: ")
            new_name = self.file_manager.get_folder_or_file_name("Enter new folder/file name: ")
            self.file_manager.rename_folder_or_file(directory, old_name, new_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_read_text_file(self):
        """Handle the 'read text file' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            file_name = self.file_manager.get_folder_or_file_name("Enter file name to read: ")
            self.file_manager.read_text_file(directory, file_name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_list_contents(self):
        """Handle the 'list contents' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            self.file_manager.list_contents(directory)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_get_properties(self):
        """Handle the 'get properties' command."""
        try:
            directory = self.file_manager.select_directory(self.file_manager.speak.stop_speaking)
            name = self.file_manager.get_folder_or_file_name("Enter folder or file name to get properties: ")
            self.file_manager.get_properties(directory, name)
        except self.file_manager.ReturnToMain:
            print("Returning to main menu.")
            self.file_manager.speak("Returning to main menu.")

    def handle_exit(self):
        """Handle the 'exit' command."""
        sys.exit("Exiting program.")