import os

class FileCommandHandler:
    def __init__(self, file_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.voice_recognizer = voice_recognizer

    def handle_open_my_computer(self, context=None):
        """Handle the 'open my computer' command."""
        success, message = self.file_manager.open_my_computer()
        if success and context is not None:
            context["last_opened_item"] = None
            context["last_created_folder"] = None
            context["working_directory"] = None  # My Computer is not a specific directory
        return message

    def handle_open_disk(self, disk_letter, context=None):
        """Handle the 'open disk <letter>' command."""
        if not disk_letter:
            return "Please specify a disk letter"
        success, message = self.file_manager.open_disk(disk_letter)
        if success and context is not None:
            disk_path = f"{disk_letter.upper()}:\\"
            context["last_opened_item"] = (disk_path, disk_path)
            context["last_created_folder"] = None
            context["working_directory"] = disk_path
        return message

    def handle_go_back(self, context=None):
        """Go back to parent directory if in File Explorer."""
        success, message = self.file_manager.go_back()
        return message

    def handle_create_folder(self, folder_name, context=None):
        """Handle the 'create folder <name>' command."""
        success, message = self.file_manager.create_folder(folder_name)
        
        if success and context is not None:
            # Sync context from file_manager
            fm_context = self.file_manager.context
            if fm_context.get("working_directory"):
                context["working_directory"] = fm_context["working_directory"]
                context["last_created_folder"] = fm_context.get("last_created_folder")
                context["last_opened_item"] = fm_context.get("last_opened_item")
                print(f"DEBUG: Context synced successfully: {context}")
        
        return message

    def handle_open_folder(self, folder_name, context=None):
        """Handle the 'open folder <name>' command."""
        success, message = self.file_manager.open_folder(folder_name)
        
        if success and context is not None:
            # Sync context from file_manager
            fm_context = self.file_manager.context
            if fm_context.get("working_directory"):
                context["working_directory"] = fm_context["working_directory"]
                context["last_opened_item"] = fm_context.get("last_opened_item")
                context["last_created_folder"] = None
                print(f"DEBUG: Context synced successfully: {context}")
        
        return message

    def handle_delete_folder(self, folder_name, context=None):
        """Handle the 'delete folder <name>' command."""
        success, message = self.file_manager.delete_folder(folder_name)
        
        if success and context is not None:
            # Sync context from file_manager
            fm_context = self.file_manager.context
            if fm_context.get("working_directory"):
                context["working_directory"] = fm_context["working_directory"]
                context["last_opened_item"] = fm_context.get("last_opened_item")
                
                # Clear references to deleted folder
                if context.get("last_created_folder") and context["last_created_folder"][1] == folder_name:
                    context["last_created_folder"] = None
                    
                print(f"DEBUG: Context synced successfully: {context}")
        
        return message

    def handle_rename_folder(self, old_name, new_name, context=None):
        """Handle the 'rename folder <old_name> to <new_name>' command."""
        success, message = self.file_manager.rename_folder(old_name, new_name)
        
        if success and context is not None:
            # Sync context from file_manager
            fm_context = self.file_manager.context
            if fm_context.get("working_directory"):
                context["working_directory"] = fm_context["working_directory"]
                context["last_created_folder"] = fm_context.get("last_created_folder")
                context["last_opened_item"] = fm_context.get("last_opened_item")
                print(f"DEBUG: Context synced successfully: {context}")
        
        return message

    def handle_open_file(self, file_name, context=None):
        """Handle the 'open file <name>' command.
        If the provided name lacks an extension, the method will search the target directory
        for a file with the same base name and any extension, then open the first match.
        """
        if not file_name:
            return "Please specify a file name"
        
        try:
            # Get smart target directory
            target_dir, location_name = self.file_manager._get_smart_target_directory()
            
            # Determine if an extension was supplied
            base_name, ext = os.path.splitext(file_name)
            if ext:
                candidate_name = file_name
            else:
                # Search for a file with the same base name and any extension
                candidates = [f for f in os.listdir(target_dir)
                              if os.path.isfile(os.path.join(target_dir, f))
                              and os.path.splitext(f)[0].lower() == base_name.lower()]
                if not candidates:
                    return f"File {file_name} not found in {location_name}"
                candidate_name = candidates[0]
            
            file_path = os.path.join(target_dir, candidate_name)
            if not os.path.exists(file_path) or not os.path.isfile(file_path):
                return f"File {candidate_name} not found in {location_name}"
            
            # Open the file
            os.startfile(file_path)
            
            if context is not None:
                context["working_directory"] = target_dir
                context["last_opened_item"] = (target_dir, candidate_name)
            
            return f"File {candidate_name} opened from {location_name}"
            
        except Exception as e:
            print(f"Error opening file '{file_name}': {e}")
            return f"Error opening file {file_name}"