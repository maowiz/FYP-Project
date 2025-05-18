import sys

class GeneralCommandHandler:
    def __init__(self, file_manager, command_handler=None):
        self.file_manager = file_manager
        # CommandHandler instance to access command list
        self.command_handler = command_handler

    def handle_list_commands(self, cmd_text=None):
        """Handle the 'list commands' command - lists all available commands."""
        print("Available commands are:")
        self.file_manager.speech.speak("Available commands are:")
        if self.command_handler:
            for cmd in self.command_handler.get_command_list():
                print(f"  • {cmd}")
                self.file_manager.speech.speak(cmd)
        print("You can say 'exit' to quit the program.")
        self.file_manager.speech.speak("You can say exit to quit the program.")

    def handle_exit(self, cmd_text=None):
        """Handle the 'exit' command."""
        print("Exiting the program...")
        self.file_manager.speech.speak("Exiting the program")
        sys.exit(0)