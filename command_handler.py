import re
from fuzzywuzzy import fuzz
from file_command_handler import FileCommandHandler
from os_command_handler import OSCommandHandler
from general_command_handler import GeneralCommandHandler

class CommandHandler:
    """
    Initializes the CommandHandler with a FileManager and an OSManager.
    Handles commands like 'create folder <name>', 'delete folder <name>',
    'rename folder <old_name> to <new_name>', and 'open folder <name>'.
    Uses context to track the current working directory.
    """
    # Command definitions
    COMMANDS = {
        "create folder": {
            "handler": "handle_create_folder",
            "handler_module": "file",
            "params": "folder_name"
        },
        "open folder": {
            "handler": "handle_open_folder",
            "handler_module": "file",
            "params": "folder_name"
        },
        "delete folder": {
            "handler": "handle_delete_folder",
            "handler_module": "file",
            "params": "folder_name"
        },
        "rename folder": {
            "handler": "handle_rename_folder",
            "handler_module": "file",
            "params": "old_name_new_name"
        },
        "list commands": {
            "handler": "handle_list_commands",
            "handler_module": "general",
            "params": None
        },
        "exit": {
            "handler": "handle_exit",
            "handler_module": "general",
            "params": None
        },
        "increase volume": {
            "handler": "handle_volume_up",
            "handler_module": "os",
            "params": None
        },
        "decrease volume": {
            "handler": "handle_volume_down",
            "handler_module": "os",
            "params": None
        },
        "mute volume": {
            "handler": "handle_mute_toggle",
            "handler_module": "os",
            "params": None
        },
        "unmute volume": {
            "handler": "handle_mute_toggle",
            "handler_module": "os",
            "params": None
        },
        "maximize volume": {
            "handler": "handle_maximize_volume",
            "handler_module": "os",
            "params": None
        },
        "set volume": {
            "handler": "handle_set_volume",
            "handler_module": "os",
            "params": "number"
        },
        "increase brightness": {
            "handler": "handle_brightness_up",
            "handler_module": "os",
            "params": None
        },
        "decrease brightness": {
            "handler": "handle_brightness_down",
            "handler_module": "os",
            "params": None
        },
        "maximize brightness": {
            "handler": "handle_maximize_brightness",
            "handler_module": "os",
            "params": None
        },
        "set brightness": {
            "handler": "handle_set_brightness",
            "handler_module": "os",
            "params": "number"
        },
        "switch window": {
            "handler": "handle_switch_window",
            "handler_module": "os",
            "params": None
        },
        "minimize all windows": {
            "handler": "handle_minimize_all_windows",
            "handler_module": "os",
            "params": None
        },
        "restore windows": {
            "handler": "handle_restore_all_windows",
            "handler_module": "os",
            "params": None
        },
        "maximize window": {
            "handler": "handle_maximize_current_window",
            "handler_module": "os",
            "params": None
        },
        "minimize window": {
            "handler": "handle_minimize_current_window",
            "handler_module": "os",
            "params": None
        },
        "close window": {
            "handler": "handle_close_current_window",
            "handler_module": "os",
            "params": None 
        },
        "move window left": {
            "handler": "handle_move_window_left",
            "handler_module": "os",
            "params": None
        },
        "move window right": {
            "handler": "handle_move_window_right",
            "handler_module": "os",
            "params": None
        },
        "take screenshot": {
            "handler": "handle_take_screenshot",
            "handler_module": "os",
            "params": None
        },
        "run application": {
            "handler": "handle_run_application",
            "handler_module": "os",
            "params": "app_name"
        }
    }

    # Synonyms for natural language commands
    COMMAND_SYNONYMS = {
        "create folder": ["make folder", "new folder", "add folder"],
        "open folder": ["access folder", "go to folder"],
        "delete folder": ["remove folder", "delete directory"],
        "rename folder": ["change folder name"],
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
        "switch window": ["next window", "change window", "switch"],
        "minimize all windows": ["show desktop", "minimize all"],
        "restore windows": ["restore all windows", "bring back windows"],
        "maximize window": ["maximize this window", "full screen"],
        "minimize window": ["minimize this window"],
        "close window": ["close this window", "close app"],
        "move window left": ["snap window left"],
        "move window right": ["snap window right"],
        "take screenshot": ["screenshot", "capture screen"],
        "run application": ["run", "open", "launch", "start"]
    }

    def __init__(self, file_manager, os_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        # Initialize specific command handlers
        self.file_handler = FileCommandHandler(file_manager, voice_recognizer)
        self.os_handler = OSCommandHandler(os_manager)
        self.general_handler = GeneralCommandHandler(file_manager, self)
        # Command context for tracking recent actions
        self.context = {
            "last_created_folder": None,
            "last_opened_item": None,
            "working_directory": None
        }

    def get_command_list(self):
        """Return the list of available commands."""
        return list(self.COMMANDS.keys())

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
        """Extract parameters from the command text."""
        if param_type == "number":
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            return match.group(1) if match else None  # Return as string
        elif param_type == "app_name":
            match = re.search(r'(?:run|open|launch|start)\s+(.+?)(?:\s|$)', cmd_text, re.IGNORECASE)
            return match.group(1).strip() if match else None
        elif param_type == "folder_name":
            match = re.search(r'(?:create|open|delete|folder)\s+(.+?)(?:\s|$)', cmd_text, re.IGNORECASE)
            return match.group(1).strip() if match else None
        elif param_type == "old_name_new_name":
            match = re.search(r'rename\s+folder\s+(.+?)\s+to\s+(.+?)(?:\s|$)', cmd_text, re.IGNORECASE)
            return (match.group(1).strip(), match.group(2).strip()) if match else (None, None)
        return None

    def execute_command(self, cmd_text):
        """Execute commands by delegating to appropriate handlers."""
        cmd_text = self.preprocess_command(cmd_text)
        
        # Check for stop command first
        if "stop" in cmd_text or "cancel" in cmd_text:
            print("Stopping current operation and returning to main menu")
            self.file_manager.speech.speak("Stopping. Ready for new command.")
            return True

        # Split text on conjunctions
        command_parts = re.split(r'\s+(and|then)\s+', cmd_text)
        command_parts = [part for part in command_parts if part not in ["and", "then"]]
        executed = False

        for part in command_parts:
            part = part.strip()
            if not part:
                continue

            # Try exact matches
            for cmd, info in self.COMMANDS.items():
                if cmd in part:
                    print(f"Executing command: {cmd}")
                    executed = self._execute_handler(info, part, cmd)
                    break

            # Try synonyms
            if not executed:
                for cmd, info in self.COMMANDS.items():
                    for synonym in self.COMMAND_SYNONYMS.get(cmd, []):
                        if synonym in part:
                            print(f"Executing synonym command: {cmd} (matched: {synonym})")
                            executed = self._execute_handler(info, part, cmd)
                            break
                    if executed:
                        break

            # Try fuzzy matching
            if not executed:
                for cmd, info in self.COMMANDS.items():
                    if fuzz.ratio(cmd, part) > 80:
                        print(f"Executing fuzzy-matched command: {cmd} (input: {part})")
                        executed = self._execute_handler(info, part, cmd)
                        break
                    for synonym in self.COMMAND_SYNONYMS.get(cmd, []):
                        if fuzz.ratio(synonym, part) > 80:
                            print(f"Executing fuzzy-matched synonym command: {cmd} (matched: {synonym}, input: {part})")
                            executed = self._execute_handler(info, part, cmd)
                            break
                    if executed:
                        break

        return executed

    def _execute_handler(self, info, cmd_text, cmd_name=None):
        """Execute the handler for a matched command."""
        params = self.extract_parameters(cmd_text, info.get("params"))
        if info["params"] and params is None:
            command_display = cmd_name if cmd_name else info["handler"]
            print(f"Missing required parameter for '{command_display}'")
            self.file_manager.speech.speak(f"Missing required parameter for {command_display}")
            return False
        handler_module = info["handler_module"]
        handler_name = info["handler"]
        if handler_module == "file":
            handler_func = getattr(self.file_handler, handler_name)
            if info["params"] == "folder_name":
                handler_func(params, self.context)
            elif info["params"] == "old_name_new_name":
                old_name, new_name = params
                handler_func(old_name, new_name, self.context)
        elif handler_module == "os":
            handler_func = getattr(self.os_handler, handler_name)
            handler_func(params if info["params"] else None)
        elif handler_module == "general":
            handler_func = getattr(self.general_handler, handler_name)
            handler_func(cmd_text if info["params"] else None)
        return True