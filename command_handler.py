import re
from fuzzywuzzy import fuzz
from file_command_handler import FileCommandHandler
from os_command_handler import OSCommandHandler
from general_command_handler import GeneralCommandHandler

class CommandHandler:
    """
    Initializes the CommandHandler with a FileManager and an OSManager.

    The FileManager provides methods for performing file and folder operations,
    while the OSManager provides methods for performing system-level operations.

    The voice_recognizer is an optional parameter that can be used to enable
    voice recognition. If provided, the CommandHandler will use this to
    recognize commands.

    The CommandHandler also stores a context dictionary that keeps track of
    recent actions and file types. This context can be used by the
    FileManager and OSManager to infer the user's intent when a command is
    given with incomplete information.

    For example, if the user says "create folder", the FileManager can use
    the context to determine which directory to create the folder in.

    The CommandHandler also stores a dictionary of command synonyms, which
    are used to recognize natural language commands. For example, the
    command "create folder" can also be recognized as "make folder",
    "new folder", or "add folder".
    """
    # Command definitions
    COMMANDS = {
        "create folder": {
            "handler": "handle_create_folder",
            "handler_module": "file",
            "params": None
        },
        "open folder or file": {
            "handler": "handle_open_folder_or_file",
            "handler_module": "file",
            "params": None
        },
        "delete folder": {
            "handler": "handle_delete_folder",
            "handler_module": "file",
            "params": None
        },
        "copy folder": {
            "handler": "handle_copy_folder",
            "handler_module": "file",
            "params": None
        },
        "cut folder": {
            "handler": "handle_cut_folder",
            "handler_module": "file",
            "params": None
        },
        "paste folder": {
            "handler": "handle_paste_folder",
            "handler_module": "file",
            "params": None
        },
        "rename folder or file": {
            "handler": "handle_rename_folder_or_file",
            "handler_module": "file",
            "params": None
        },
        "read text file": {
            "handler": "handle_read_text_file",
            "handler_module": "file",
            "params": None
        },
        "list contents": {
            "handler": "handle_list_contents",
            "handler_module": "file",
            "params": None
        },
        "get properties": {
            "handler": "handle_get_properties",
            "handler_module": "file",
            "params": None
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
        }
    }

    # Synonyms for natural language commands
    COMMAND_SYNONYMS = {
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

    def __init__(self, file_manager, os_manager, voice_recognizer=None):
        self.file_manager = file_manager
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        # Initialize specific command handlers
        self.file_handler = FileCommandHandler(file_manager, voice_recognizer)
        self.os_handler = OSCommandHandler(os_manager)
        self.general_handler = GeneralCommandHandler(file_manager)
        # Command context for tracking recent actions
        self.context = {
            "last_created_folder": None,  # Stores (directory, folder_name)
            "last_opened_item": None,    # Stores (directory, name)
            "file_type": None            # Stores file type (e.g., '.txt', '.docx')
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
        """Extract parameters (e.g., numbers) from the command text."""
        if param_type == "number":
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            return int(match.group(1)) if match else None
        return None

    def execute_command(self, cmd_text):
        """Execute commands by delegating to appropriate handlers."""
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
            for cmd, info in self.COMMANDS.items():
                if cmd in part:
                    print(f"Executing command: {cmd}")
                    self._execute_handler(info, part)
                    executed = True
                    break

            # If no exact match, try synonyms
            if not executed:
                for cmd, info in self.COMMANDS.items():
                    for synonym in self.COMMAND_SYNONYMS.get(cmd, []):
                        if synonym in part:
                            print(f"Executing synonym command: {cmd} (matched: {synonym})")
                            self._execute_handler(info, part)
                            executed = True
                            break
                    if executed:
                        break

            # If no exact or synonym match, try fuzzy matching
            if not executed:
                for cmd, info in self.COMMANDS.items():
                    if fuzz.ratio(cmd, part) > 80:  # 80% similarity threshold
                        print(f"Executing fuzzy-matched command: {cmd} (input: {part})")
                        self._execute_handler(info, part)
                        executed = True
                        break
                    # Check synonyms with fuzzy matching
                    for synonym in self.COMMAND_SYNONYMS.get(cmd, []):
                        if fuzz.ratio(synonym, part) > 80:
                            print(f"Executing fuzzy-matched synonym command: {cmd} (matched: {synonym}, input: {part})")
                            self._execute_handler(info, part)
                            executed = True
                            break
                    if executed:
                        break

        return executed

    def _execute_handler(self, info, cmd_text):
        """Execute the handler for a matched command."""
        params = self.extract_parameters(cmd_text, info.get("params"))
        if info["params"] and params is None:
            print(f"Missing required parameter for '{info.get('command', 'unknown command')}'")
            self.file_manager.speech.speak(f"Missing required parameter for {info.get('command', 'unknown command')}")
            return
        # Delegate to the appropriate handler
        handler_module = info["handler_module"]
        handler_name = info["handler"]
        if handler_module == "file":
            handler_func = getattr(self.file_handler, handler_name)
            handler_func(cmd_text if info["params"] else None, self.context)
        elif handler_module == "os":
            handler_func = getattr(self.os_handler, handler_name)
            handler_func(cmd_text if info["params"] else None)
        elif handler_module == "general":
            handler_func = getattr(self.general_handler, handler_name)
            handler_func(cmd_text if info["params"] else None)