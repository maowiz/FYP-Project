class FileCommandHandler:
    def __init__(self, file_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.voice_recognizer = voice_recognizer

    def handle_create_folder(self, folder_name, context=None):
        """Handle the 'create folder <name>' command."""
        result = self.file_manager.create_folder(folder_name)
        if result and context:
            context["last_created_folder"] = (self.file_manager.context["working_directory"], folder_name)
            context["last_opened_item"] = None

    def handle_open_folder(self, folder_name, context=None):
        """Handle the 'open folder <name>' command."""
        result = self.file_manager.open_folder(folder_name)
        if result and context:
            context["last_opened_item"] = (self.file_manager.context["working_directory"], folder_name)
            context["last_created_folder"] = None

    def handle_delete_folder(self, folder_name, context=None):
        """Handle the 'delete folder <name>' command."""
        result = self.file_manager.delete_folder(folder_name)
        if result and context:
            if context.get("last_created_folder") and context["last_created_folder"][1] == folder_name:
                context["last_created_folder"] = None
            if context.get("last_opened_item") and context["last_opened_item"][1] == folder_name:
                context["last_opened_item"] = None

    def handle_rename_folder(self, old_name, new_name, context=None):
        """Handle the 'rename folder <old_name> to <new_name>' command."""
        result = self.file_manager.rename_folder(old_name, new_name)
        if result and context:
            if context.get("last_created_folder") and context["last_created_folder"][1] == old_name:
                context["last_created_folder"] = (self.file_manager.context["working_directory"], new_name)
            if context.get("last_opened_item") and context["last_opened_item"][1] == old_name:
                context["last_opened_item"] = (self.file_manager.context["working_directory"], new_name) 