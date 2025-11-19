import re
from fuzzywuzzy import fuzz
from file_command_handler import FileCommandHandler
from os_command_handler import OSCommandHandler
from general_command_handler import GeneralCommandHandler
import time
import threading
import tkinter as tk
import ctypes
import win32con
import webbrowser
import os
import pickle
import logging
import subprocess
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Vision mode keywords for offline priority
VISION_MODE_KEYWORDS = [
    "go vision", 
    "vision mode", 
    "activate vision", 
    "hand control",
    "gesture control",
    "visual input"
]

class CommandHandler:
    """
    Initializes the CommandHandler with a FileManager and an OSManager.
    Handles commands like 'create folder <name>', 'delete folder <name>',
    'rename folder <old_name> to <new_name>', 'open folder <name>',
    'open my computer', 'open disk <letter>', and 'go back'.
    Uses context to track the current working directory.
    """
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
        "open my computer": {
            "handler": "handle_open_my_computer",
            "handler_module": "file",
            "params": None
        },
        "open disk": {
            "handler": "handle_open_disk",
            "handler_module": "file",
            "params": "disk_letter"
        },
        "go back": {
            "handler": "handle_go_back",
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
        "switch tab": {
            "handler": "handle_switch_tab",
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
        "show grid": {
            "handler": "handle_show_grid",
            "handler_module": "os",
            "params": None
        },
        "hide grid": {
            "handler": "handle_hide_grid",
            "handler_module": "os",
            "params": None
        },
        "click cell": {
            "handler": "handle_click_cell",
            "handler_module": "os",
            "params": "number"
        },
        "double click cell": {
            "handler": "handle_double_click_cell",
            "handler_module": "os",
            "params": "number"
        },
        "right click cell": {
            "handler": "handle_right_click_cell",
            "handler_module": "os",
            "params": "number"
        },
        "drag from": {
            "handler": "handle_drag_from",
            "handler_module": "os",
            "params": "number"
        },
        "drop on": {
            "handler": "handle_drop_on",
            "handler_module": "os",
            "params": "number"
        },
        "zoom cell": {
            "handler": "handle_zoom_cell",
            "handler_module": "os",
            "params": "number"
        },
        "exit zoom": {
            "handler": "handle_exit_zoom",
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
        },
        "add numbers": {
            "handler": "handle_add_numbers",
            "handler_module": "general",
            "params": "numbers"
        },
        "subtract numbers": {
            "handler": "handle_subtract_numbers",
            "handler_module": "general",
            "params": "numbers"
        },
        "multiply numbers": {
            "handler": "handle_multiply_numbers",
            "handler_module": "general",
            "params": "numbers"
        },
        "divide numbers": {
            "handler": "handle_divide_numbers",
            "handler_module": "general",
            "params": "numbers"
        },
        "show system info": {
            "handler": "handle_system_info",
            "handler_module": "general",
            "params": None
        },
        "tell time": {
            "handler": "handle_tell_time",
            "handler_module": "general",
            "params": None
        },
        "tell date": {
            "handler": "handle_tell_date",
            "handler_module": "general",
            "params": None
        },
        "tell day": {
            "handler": "handle_tell_day",
            "handler_module": "general",
            "params": None
        },
        "tell weather": {
            "handler": "handle_tell_weather",
            "handler_module": "general",
            "params": "city_name"
        },
        "tell joke": {
            "handler": "handle_tell_joke",
            "handler_module": "general",
            "params": None
        },
        "check internet": {
            "handler": "handle_check_internet",
            "handler_module": "general",
            "params": None
        }, 
        "check internet speed": {
            "handler": "handle_check_internet_speed",
            "handler_module": "general",
            "params": None
        },
        "check bmi": {
            "handler": "handle_check_bmi",
            "handler_module": "general",
            "params": "bmi_data"
        },
        "take a photo": {
            "handler": "handle_take_photo",
            "handler_module": "general",
            "params": None
        },
        "set wallpaper to": {
            "handler": "handle_set_wallpaper",
            "handler_module": "general",
            "params": "image_file"
        },
        "go to desktop": {
            "handler": "handle_go_to_desktop",
            "handler_module": "os",
            "params": None
        },
        "change wallpaper": {
            "handler": "handle_change_wallpaper",
            "handler_module": "os",
            "params": None
        },
        "countdown": {
            "handler": "handle_countdown",
            "handler_module": "general",
            "params": "seconds"
        }, 
        "empty recycle bin": {
          "handler": "handle_empty_recycle_bin",
          "handler_module": "os",
          "params": None
        },
        "find my phone": {
            "handler": "handle_find_phone",
            "handler_module": "general",
            "params": None
        },
        "spell": {
            "handler": "handle_spell",
            "handler_module": "general",
            "params": "text"
        }, 
        "previous tab": {
            "handler": "handle_previous_tab",
            "handler_module": "os",
            "params": None
        },
        "next tab": {
            "handler": "handle_next_tab",   
            "handler_module": "os",
            "params": None
        },
        "close tab": {
            "handler": "handle_close_tab",
            "handler_module": "os",
            "params": None
        },
        "refresh": {
            "handler": "handle_refresh",
            "handler_module": "os", 
            "params": None
        },
        "zoom in": {
            "handler": "handle_zoom_in",
            "handler_module": "os",
            "params": None  
        },
        "zoom out": {
            "handler": "handle_zoom_out",
            "handler_module": "os",
            "params": None
        },  
        "set grid size": {
            "handler": "handle_set_grid_size",
            "handler_module": "os",
            "params": "number"
        },
        "bookmark tab": {
            "handler": "handle_bookmark_tab",
            "handler_module": "os",
            "params": None
        },
        "open incognito": { 
            "handler": "handle_open_incognito",
            "handler_module": "os",
            "params": None
        },
        "switch tab": {
            "handler": "handle_switch_tab",     
            "handler_module": "os",
            "params": "number"
        },
        "search": {
            "handler": "handle_search",
            "handler_module": "os",
            "params": "query"
        },
        "clear browsing data": {
            "handler": "handle_clear_browsing_data",
            "handler_module": "os",
            "params": None
        },
        "scroll up": {
            "handler": "handle_scroll_up",
            "handler_module": "os",
            "params": None
        },
        "scroll down": {
            "handler": "handle_scroll_down",
            "handler_module": "os",
            "params": None
        },
        "scroll left": {
            "handler": "handle_scroll_left",
            "handler_module": "os",
            "params": None
        },
        "scroll right": {
            "handler": "handle_scroll_right",
            "handler_module": "os",
            "params": None
        },
        "stop scrolling": {
            "handler": "handle_stop_scrolling",
            "handler_module": "os",
            "params": None
        },
        "open": {
            "handler": "handle_open_generic",
            "handler_module": "os",
            "params": "open_target"
        },
        "play on youtube": {
            "handler": "handle_play_on_youtube",
            "handler_module": "os",
            "params": "query"
        },
        "check disk space": {
            "handler": "handle_check_disk_space",
            "handler_module": "general",
            "params": None
        },
        
        "read most recent email": {
            "handler": "handle_read_most_recent_email",
            "handler_module": "general",
            "params": None
        },
        "read oldest email": {
            "handler": "handle_read_oldest_email",
            "handler_module": "general",
            "params": None
        },
        "read nth most recent email": {
            "handler": "handle_read_nth_most_recent_email",
            "handler_module": "general",
            "params": "nth_email"
        },
        "read nth oldest email": {
            "handler": "handle_read_nth_oldest_email",
            "handler_module": "general",
            "params": "nth_email"
        },
        "copy": {
            "handler": "handle_copy",
            "handler_module": "os",
            "params": None
        },
        "paste": {
            "handler": "handle_paste",
            "handler_module": "os",
            "params": None
        },
        "read clipboard": {
            "handler": "handle_read_clipboard",
            "handler_module": "os",
            "params": None
        },
        "select all": {
            "handler": "handle_select_all",
            "handler_module": "os",
            "params": None
        },
        "open word": {
            "handler": "handle_open_word",
            "handler_module": "os",
            "params": None
        },
        "write essay": {
            "handler": "handle_write_essay",
            "handler_module": "general",
            "params": "topic"
        },
        "save file": {
            "handler": "handle_save_file",
            "handler_module": "os",
            "params": "filename"
        },
        "remove this": {
            "handler": "handle_remove_selection",
            "handler_module": "os",
            "params": None
        },
        "undo": {
            "handler": "handle_undo_action",
            "handler_module": "os",
            "params": None
        },
        "redo": {
            "handler": "handle_redo_action",
            "handler_module": "os",
            "params": None
        },
        "send to chatgpt": {
            "handler": "handle_send_to_chatgpt",
            "handler_module": "general",
            "params": "query"
        },
        "start dictation": {
            "handler": "handle_noop",  # This command is handled in main.py
            "handler_module": "general",
            "params": None
        },
        "stop dictation": {
            "handler": "handle_noop",  # This command is handled in main.py
            "handler_module": "general",
            "params": None
        },
        "lock computer": {
            "handler": "handle_lock_computer",
            "handler_module": "os",
            "params": None
        },
        "take a note": {
            "handler": "handle_noop",  # This command is handled in main.py
            "handler_module": "general",
            "params": None
        },
        "read last note": {
            "handler": "handle_read_last_note",
            "handler_module": "general",
            "params": None
        },
        "show all notes": {
            "handler": "handle_show_all_notes",
            "handler_module": "general",
            "params": None
        },
        "summarize clipboard": {
            "handler": "handle_summarize_clipboard",
            "handler_module": "general",
            "params": None
        },
    }

    # Synonyms for natural language commands
    COMMAND_SYNONYMS = {
        "create folder": ["make folder", "new folder", "add folder"],
        "open folder": ["access folder", "go to folder"],
        "delete folder": ["remove folder", "delete directory"],
        "rename folder": ["change folder name"],
        "open my computer": ["open this pc", "this pc", "my computer"],
        "open disk": [
            "open drive", "access disk", "access drive", "open disc", "access disc",
            "open local disk", "access local disk"
        ],
        "go back": ["back", "return", "go up"],
        "list commands": ["help", "commands", "what can you do"],
        "exit": [
            "quit", "stop program", "bye", "good bye", "goodbye",
            "shut down", "terminate", "kill program", "shutdown assistant"
        ],
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
        "switch tab": ["next tab", "change tab", "switch tab", "switch to next tab"],
        "switch window": ["next window", "change window", "switch"],
        "minimize all windows": ["show desktop", "minimize all"],
        "restore windows": ["restore all windows", "bring back windows"],
        "show grid": ["show grade", "display grid", "open grid", "grid on", "show mouse grid"],
        "hide grid": ["close grid", "grid off", "remove grid"],
        "click cell": ["click", "left click"],
        "double click cell": ["double click", "double-click"],
        "right click cell": ["right click", "right-click"],
        "drag from": ["drag", "start drag"],
        "drop on": ["drop", "to"],
        "zoom cell": ["zoom cell", "zoom into cell"],
        "exit zoom": ["exit grid zoom", "go back", "back from zoom"],
        "maximize window": ["maximize this window", "full screen"],
        "minimize window": ["minimize this window"],
        "close window": [
            "close app", "close application", "close program",
            "close current window", "close current app", "close active window",
            "close focused window", "close this window", "close that window", "kill window", "kill app"
        ],
        "move window left": ["snap window left"],
        "move window right": ["snap window right"],
        "take screenshot": ["screenshot", "capture screen"],
        "run application": ["run", "launch", "start"],
        "add numbers": ["plus", "add", "sum"],
        "subtract numbers": ["minus", "subtract", "difference"],
        "multiply numbers": ["times", "multiply", "product"],
        "divide numbers": ["divide", "division", "quotient"],
        "show system info": ["system info", "check ram", "tell me battery status"],
        "tell time": ["what time is it", "current time", "time now"],
        "tell date": ["tell me the date", "what is the date", "current date", "today's date"],
        "tell day": ["what day is it", "current day", "today is"],
        "tell weather": ["weather in", "what's the weather in", "weather forecast for", "weather", "tell weather"],
        "tell joke": ["tell me a joke", "joke", "make me laugh"],
        "check internet": ["is internet working", "check connection"],
        "check internet speed": ["test internet speed", "internet speed", "check my internet speed"],
        "check bmi": ["calculate bmi", "what is my bmi", "bmi check", "bmi"],
        "take a photo": ["capture photo", "take picture", "snap a photo"],
        "set wallpaper to": ["change wallpaper to", "update wallpaper to"],
        "go to desktop": [
            "show desktop", "take me to desktop", "open desktop",
            "minimize all", "desktop please"
        ],
        "change wallpaper": [
            "next wallpaper", "next background", "change background",
            "switch wallpaper", "next slide"
        ],
        "countdown": ["start countdown", "timer", "count down", "set timer"],
        "empty recycle bin": ["clear trash", "delete trash", "empty trash", "clear bin"],
        "find my phone": ["ring my phone", "locate my phone", "where is my phone"],
        "spell": ["phonetic", "phonetically", "spelling"],
        "previous tab": ["back tab", "prev tab", "go back"],
        "next tab": ["forward tab", "next browser tab"],
        "close tab": ["close current tab", "close browser tab", "close this tab", "close active tab"],
        "refresh": ["reload tab", "refresh tab"],
        "zoom in": ["increase zoom", "zoom bigger", "zooming"],
        "zoom out": ["decrease zoom", "zoom smaller", "zoom out", "zoomed out"],
        "set grid size": ["zoom", "grid"],
        "bookmark tab": ["save tab", "bookmark current tab"],
        "open incognito": ["private tab", "incognito window"],
        "switch tab": ["go to tab", "jump to tab"],
        "search": ["google", "look up", "find"],
        "clear browsing data": ["clear history", "delete cache", "clear browser data"],
        "scroll up": ["scroll upward", "move up", "page up"],
        "scroll down": ["scroll downward", "move down", "page down"],
        "scroll left": ["move left", "pan left"],
        "scroll right": ["move right", "pan right"],
        "stop scrolling": ["cancel scrolling", "halt scrolling", "stop scroll"],
        "open": ["launch", "start", "go to"],
        "play on youtube": ["youtube", "play youtube", "play video", "play song", "play music on youtube", "search youtube"],
        "check disk space": [
            "disk space", "show disk space", "free space", "storage info", "check storage", "disk usage", "drive space", "space left on disk"
        ],
        
        "read most recent email": ["read latest email", "read newest email", "read top email"],
        "read oldest email": ["read first email", "read very first email", "read bottom email"],
        "read nth most recent email": ["read {n}th most recent email", "read {n} most recent email", "read {n}th latest email", "read {n} latest email"],
        "read nth oldest email": ["read {n}th oldest email", "read {n} oldest email", "read {n}th first email", "read {n} first email"],
        "copy": ["copy that", "copy this"],
        "paste": ["paste that", "paste this", "paste here"],
        "read clipboard": ["what's on the clipboard", "read my clipboard", "tell me what's on the clipboard"],
        "select all": ["select everything"],
        "open word": ["launch word", "start word", "microsoft word"],
        "write essay": ["write an essay on", "write about", "compose an essay on"],
        "save file": ["save this file", "save it", "save the file"],
        "remove this": ["delete this", "remove selection", "delete selection", "clear selection"],
        "undo": ["undo that", "control z", "that was a mistake", "go back"],
        "redo": ["redo that", "control y"],
        "send to chatgpt": ["chatgpt", "ask chatgpt", "tell chatgpt", "on chatgpt"],
        "start dictation": ["start dictation mode", "begin dictation", "dictation on"],
        "stop dictation": ["stop dictation mode", "end dictation", "dictation off"],
        "take a note": ["add a note", "new note", "write a note", "note this down"],
        "read last note": ["what was my last note", "read the last note"],
        "show all notes": ["open my notes", "show notes", "view all notes"],
        "lock computer": ["lock my computer", "lock the screen", "lock screen", "lock pc"],
        "summarize clipboard": ["summarize this", "summarize the clipboard", "give me a summary", "summarise", "summary"],
    }

    ORDINAL_WORDS = {
        "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
        "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
        "eleventh": 11, "twelfth": 12, "thirteenth": 13, "fourteenth": 14, "fifteenth": 15,
        "sixteenth": 16, "seventeenth": 17, "eighteenth": 18, "nineteenth": 19, "twentieth": 20
    }

    def __init__(self, file_manager, os_manager, voice_recognizer=None, speech=None, is_online=True):
        self.file_manager = file_manager
        self.os_manager = os_manager
        self.voice_recognizer = voice_recognizer
        self.speech = speech
        self.hybrid_processor = None # Will be set from main.py
        # Initialize specific command handlers
        self.file_handler = FileCommandHandler(file_manager, voice_recognizer)
        self.os_handler = OSCommandHandler(os_manager)
        self.general_handler = GeneralCommandHandler(file_manager, self, speech)
        # Command context for tracking recent actions
        self.context = {
            "last_created_folder": None,
            "last_opened_item": None,
            "working_directory": None,
            "last_read_email": None,
        }
        # Gmail API setup (conditional)
        self.gmail_service = None
        if is_online:
            try:
                self._setup_gmail_api()
            except Exception as e:
                print(f"Failed to setup Gmail API even in online mode: {e}")
                self.gmail_service = None # Ensure it's None on failure
        else:
            print("Offline mode: Gmail API and email commands are disabled.")
            # Remove email commands if offline to prevent errors
            email_commands = [
                "read most recent email",
                "read oldest email",
                "read nth most recent email",
                "read nth oldest email",
            ]
            for cmd in email_commands:
                if cmd in self.COMMANDS:
                    del self.COMMANDS[cmd]


    def _setup_gmail_api(self):
        print("Setting up Gmail API...")
        SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        creds = None
        
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        credentials_path = os.path.join(script_dir, 'credentials.json')
        token_path = os.path.join(script_dir, 'token.pickle')
        
        if os.path.exists(token_path):
            print("Found token.pickle")
            with open(token_path, 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            print("No valid creds, starting OAuth flow")
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token_path, 'wb') as token:
                pickle.dump(creds, token)
        self.gmail_service = build('gmail', 'v1', credentials=creds)
        print("Gmail API setup complete.")

    def ordinal_to_index(self, ordinal):
        # Supports both words and numbers ("first"/"1st", etc.)
        ordinals = {
            'first': 0, '1st': 0, 'one': 0, '1': 0,
            'second': 1, '2nd': 1, 'two': 1, '2': 1,
            'third': 2, '3rd': 2, 'three': 2, '3': 2,
            'fourth': 3, '4th': 3, 'four': 3, '4': 3,
            'fifth': 4, '5th': 4, 'five': 4, '5': 4,
            'sixth': 5, '6th': 5, 'six': 5, '6': 5,
            'seventh': 6, '7th': 6, 'seven': 6, '7': 6,
            'eighth': 7, '8th': 7, 'eight': 7, '8': 7,
            'ninth': 8, '9th': 8, 'nine': 8, '9': 8,
            'tenth': 9, '10th': 9, 'ten': 9, '10': 9
        }
        return ordinals.get(ordinal.lower(), None)

    def handle_read_nth_email(self, ordinal):
        print(f"handle_read_nth_email called with ordinal: {ordinal}")
        idx = self.ordinal_to_index(ordinal)
        print(f"Resolved ordinal to index: {idx}")
        if idx is None:
            self.file_manager.speech.speak("I didn't understand which email you want to read.")
            print("Ordinal not understood.")
            return
        try:
            service = self.gmail_service
            print("Fetching messages from Gmail...")
            results = service.users().messages().list(userId='me', maxResults=idx+1).execute()
            messages = results.get('messages', [])
            print(f"Fetched {len(messages)} messages.")
            if not messages or len(messages) <= idx:
                self.file_manager.speech.speak("There aren't that many emails in your inbox.")
                print("Not enough emails.")
                return
            msg_id = messages[idx]['id']
            print(f"Fetching message with id: {msg_id}")
            msg_data = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
            headers = msg_data['payload'].get('headers', [])
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '(No Subject)')
            from_ = next((h['value'] for h in headers if h['name'] == 'From'), '(Unknown Sender)')
            snippet = msg_data.get('snippet', '')
            self.context['last_read_email'] = {'id': msg_id, 'subject': subject}
            self.file_manager.speech.speak(f"Email from {from_}, subject: {subject}. {snippet}")
            print(f"Read email: From: {from_}, Subject: {subject}, Snippet: {snippet}")
            url = f"https://mail.google.com/mail/u/0/#inbox/{msg_id}"
            webbrowser.open(url)
        except Exception as e:
            self.file_manager.speech.speak(f"Failed to read email: {e}")
            print(f"Exception in handle_read_nth_email: {e}")

    

    def handle_open_that_email(self):
        email = self.context.get('last_read_email')
        if not email:
            self.file_manager.speech.speak("No email has been read yet.")
            return
        url = f"https://mail.google.com/mail/u/0/#inbox/{email['id']}"
        webbrowser.open(url)
        self.file_manager.speech.speak(f"Opening the last read email: {email['subject']}")

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
        # Handle number-only inputs by prepending the last command
        if re.match(r'^\s*\d+\s+\d+\s*$', cmd_text) and self.context.get("last_command"):
            cmd_text = f"{self.context['last_command']} {cmd_text}"
            print(f"Prepended last command: {cmd_text}")
        # Replace pronouns with context
        if " it " in cmd_text or cmd_text.endswith(" it"):
            if self.context["last_created_folder"]:
                directory, name = self.context["last_created_folder"]
                cmd_text = cmd_text.replace(" it", f" {name}")
            elif self.context["last_opened_item"]:
                directory, name = self.context["last_opened_item"]
                cmd_text = cmd_text.replace(" it", f" {name}")
        print(f"Preprocessed command: {cmd_text}")
        return cmd_text

    def find_command(self, cmd_text):
        # cmd_text is already preprocessed by execute_command
        print(f"Preprocessed command for matching: '{cmd_text}'")
        # Normalize the first word to 'read' if it is a close fuzzy match
        words = cmd_text.split()
        if words and fuzz.ratio(words[0], 'read') > 80:
            print(f"Normalizing first word '{words[0]}' to 'read'")
            words[0] = 'read'
            cmd_text = ' '.join(words)
        print(f"Command after normalization: '{cmd_text}'")
        cmd_name = None
        params = None
        
        # --- PATTERN MATCHING FIRST (HIGHEST PRIORITY) ---
        # Browser zoom commands (must come before grid size pattern)
        m_zoom_in = re.match(r'^zoom\s+in$', cmd_text, re.IGNORECASE)
        if m_zoom_in:
            print("Pattern matched 'zoom in'")
            cmd_name = "zoom in"
            params = None
            return cmd_name, params

        m_zoom_out = re.match(r'^zoom\s+out$', cmd_text, re.IGNORECASE)
        if m_zoom_out:
            print("Pattern matched 'zoom out'")
            cmd_name = "zoom out"
            params = None
            return cmd_name, params

        # Handle "zooming" as "zoom in"
        m_zooming = re.match(r'^zooming$', cmd_text, re.IGNORECASE)
        if m_zooming:
            print("Pattern matched 'zooming' as 'zoom in'")
            cmd_name = "zoom in"
            params = None
            return cmd_name, params

        # Handle "zoom bigger" as "zoom in"
        m_zoom_bigger = re.match(r'^zoom\s+bigger$', cmd_text, re.IGNORECASE)
        if m_zoom_bigger:
            print("Pattern matched 'zoom bigger' as 'zoom in'")
            cmd_name = "zoom in"
            params = None
            return cmd_name, params

        # Handle "zoom smaller" as "zoom out"
        m_zoom_smaller = re.match(r'^zoom\s+smaller$', cmd_text, re.IGNORECASE)
        if m_zoom_smaller:
            print("Pattern matched 'zoom smaller' as 'zoom out'")
            cmd_name = "zoom out"
            params = None
            return cmd_name, params

        # Handle "zoomed" as "zoom in"
        m_zoomed = re.match(r'^zoomed$', cmd_text, re.IGNORECASE)
        if m_zoomed:
            print("Pattern matched 'zoomed' as 'zoom in'")
            cmd_name = "zoom in"
            params = None
            return cmd_name, params

        # Scroll commands pattern matching
        m_scroll_up = re.match(r'^scroll\s+up$', cmd_text, re.IGNORECASE)
        if m_scroll_up:
            print("Pattern matched 'scroll up'")
            cmd_name = "scroll up"
            params = None
            return cmd_name, params

        m_scroll_down = re.match(r'^scroll\s+down$', cmd_text, re.IGNORECASE)
        if m_scroll_down:
            print("Pattern matched 'scroll down'")
            cmd_name = "scroll down"
            params = None
            return cmd_name, params

        m_scroll_left = re.match(r'^scroll\s+left$', cmd_text, re.IGNORECASE)
        if m_scroll_left:
            print("Pattern matched 'scroll left'")
            cmd_name = "scroll left"
            params = None
            return cmd_name, params

        m_scroll_right = re.match(r'^scroll\s+right$', cmd_text, re.IGNORECASE)
        if m_scroll_right:
            print("Pattern matched 'scroll right'")
            cmd_name = "scroll right"
            params = None
            return cmd_name, params

        m_stop_scrolling = re.match(r'^stop\s+scrolling$', cmd_text, re.IGNORECASE)
        if m_stop_scrolling:
            print("Pattern matched 'stop scrolling'")
            cmd_name = "stop scrolling"
            params = None
            return cmd_name, params

        # Browser tab commands pattern matching
        m_close_tab = re.match(r'^close\s+tab$', cmd_text, re.IGNORECASE)
        if m_close_tab:
            print("Pattern matched 'close tab'")
            cmd_name = "close tab"
            params = None
            return cmd_name, params

        m_next_tab = re.match(r'^next\s+tab$', cmd_text, re.IGNORECASE)
        if m_next_tab:
            print("Pattern matched 'next tab'")
            cmd_name = "next tab"
            params = None
            return cmd_name, params

        m_previous_tab = re.match(r'^previous\s+tab$', cmd_text, re.IGNORECASE)
        if m_previous_tab:
            print("Pattern matched 'previous tab'")
            cmd_name = "previous tab"
            params = None
            return cmd_name, params

        # Prefer specific disk open phrases over generic 'open'
        m_open_disk = re.match(r'^(?:open|access|go to)\s+(?:disk|disc|drive)\s+([a-zA-Z]):?(?:\b|$)', cmd_text, re.IGNORECASE)
        if m_open_disk:
            print("Pattern matched 'open disk'")
            cmd_name = "open disk"
            params = None
            return cmd_name, params

        # Also support reversed phrasing like 'open C drive' or 'go to D disk'
        m_open_disk_rev = re.match(r'^(?:open|access|go to)\s+([a-zA-Z])\s*(?:drive|disk|disc)(?:\b|$)', cmd_text, re.IGNORECASE)
        if m_open_disk_rev:
            print("Pattern matched 'open <letter> drive'")
            cmd_name = "open disk"
            params = None
            return cmd_name, params

        # Pattern match for grid sizing: 'zoom 15', 'grid 10', 'zoom fifteen' (but not 'zoom in' or 'zoom out')
        m_zoom_size = re.match(r'^(?:zoom|grid)\s+([a-z0-9 -]+)$', cmd_text, re.IGNORECASE)
        if m_zoom_size and cmd_text.lower() not in ['zoom in', 'zoom out']:
            print("Pattern matched 'set grid size'")
            cmd_name = "set grid size"
            params = None  # will be extracted later by extract_parameters("number")
            return cmd_name, params

        # Pattern match for 'read [the] {ordinal|number} (most recent|oldest|recent) [email]'
        match = re.match(r'read\s+(?:the\s+)?(\d+|[a-z]+)(?:st|nd|rd|th)?\s+(most recent|oldest|recent)(?:\s+email)?$', cmd_text, re.IGNORECASE)
        if match:
            idx_raw = match.group(1)
            which = match.group(2)
            print(f"Pattern match: idx_raw='{idx_raw}', which='{which}'")
            # Treat 'recent' as 'most recent'
            if which == 'recent':
                which = 'most recent'
            if idx_raw.isdigit():
                idx = int(idx_raw) - 1
            else:
                idx = self.ORDINAL_WORDS.get(idx_raw.lower(), None)
                if idx is not None:
                    idx -= 1
            cmd_name = f"read nth {which} email"
            params = (idx, which)
            print(f"Pattern matched '{cmd_name}' with index: {params[0]}")
            return cmd_name, params
        
        # Pattern match for 'read [the] (most recent|oldest|recent) [email]'
        match = re.match(r'read\s+(?:the\s+)?(most recent|oldest|recent)(?:\s+email)?$', cmd_text, re.IGNORECASE)
        if match:
            which = match.group(1)
            print(f"Pattern match: which='{which}'")
            if which == 'recent':
                which = 'most recent'
            cmd_name = f"read {which} email"
            params = None
            print(f"Pattern matched '{cmd_name}'")
            return cmd_name, params
        # --- END PATTERN MATCHING ---
        
        # --- FIX: More flexible matching for commands with parameters ---
        # This helps catch commands like "on chat gpt write a poem" where the trigger has variations.
        # This block is now placed before the direct/synonym matching.
        text_lower = cmd_text.lower().strip()
        # Define commands that often start with a trigger phrase.

        # --- FIX: Prioritize suffix commands to override prefixes ---
        # If a command ends with a target ("on gpt" or "on word"), identify it first.
        if any(text_lower.endswith(" " + trigger) for trigger in ["gpt", "chat gpt", "chatgpt", "on gpt", "on chat gpt", "on chatgpt"]):
            cmd_name = "send to chatgpt"
        elif any(text_lower.endswith(" " + trigger) for trigger in ["word", "on word", "in word"]):
            cmd_name = "write essay"
        
        prefix_commands = {
            "send to chatgpt": ["ask chatgpt", "tell chatgpt", "on chatgpt", "chatgpt", "chat gpt", "on chat gpt"],
            "write essay": ["write an essay on", "write about", "compose an essay on", "write a"],
            "search": ["search for", "google", "look up", "find"],
            "play on youtube": ["play on youtube", "youtube", "play video", "play song", "play music on youtube"],
            "open": ["open", "launch", "start", "go to"],
        }

        # Only check for prefixes if a more specific suffix command wasn't already found
        if cmd_name is None:
            # Sort by length to match longer triggers first (e.g., "play on youtube" before "youtube")
            for command, triggers in prefix_commands.items():
                # Sort triggers by length, longest first, to avoid partial matches (e.g., "chat gpt" before "gpt")
                triggers.sort(key=len, reverse=True)
                for trigger in triggers:
                    if text_lower.startswith(trigger + ' '):
                        cmd_name = command
                        break # Found a prefix match
                if cmd_name:
                    break # Stop after finding the first matching command group
        if cmd_name is None:
            for command in self.COMMANDS:
                if f" {cmd_text} ".startswith(f" {command} "): # Use word boundaries
                    cmd_name = command
                    print(f"Direct match found: '{command}' for input '{cmd_text}'")
                    break
        # Synonym match
        if not cmd_name:
            for command, synonyms in self.COMMAND_SYNONYMS.items():
                for synonym in synonyms:
                    if f" {cmd_text} ".startswith(f" {synonym} "): # Use word boundaries
                        cmd_name = command
                        print(f"Synonym match found: '{command}' via synonym '{synonym}' for input '{cmd_text}'")
                        break
                if cmd_name:
                    break
        # Fuzzy matching for commands and synonyms
        if not cmd_name:
            for cmd, info in self.COMMANDS.items():
                handler_name = info["handler"]
                if fuzz.ratio(cmd, cmd_text) > 80:
                    print(f"Executing fuzzy-matched command: {cmd} (input: {cmd_text})")
                    cmd_name = cmd
                    break
                for synonym in self.COMMAND_SYNONYMS.get(cmd, []):
                    if fuzz.ratio(synonym, cmd_text) > 80:
                        print(f"Executing fuzzy-matched command: {cmd} (matched: {synonym}, input: {cmd_text})")
                        cmd_name = cmd
                        break
                if cmd_name:
                    break
        print(f"Attempted command match: '{cmd_text}', Matched command: '{cmd_name}'")  # Debug print
        if cmd_name and cmd_name in self.COMMANDS:
            info = self.COMMANDS[cmd_name]
            if info["params"]:
                params = self.extract_parameters(cmd_text, info["params"])
            return cmd_name, params
        print(f"No command matched for: {cmd_text}")
        return None, None

    def extract_parameters(self, cmd_text, param_type):
        """Extract parameters from the command text."""
        cmd_text = cmd_text.lower().strip()
        print(f"Extracting parameters for '{cmd_text}' with type '{param_type}'")
        if param_type == "open_target":
            # Extracts the target after 'open', 'launch', 'start', or 'go to'
            match = re.search(r'(?:open|launch|start|go to)\s+(.+)', cmd_text)
            if match:
                param = match.group(1).strip()
                print(f"Extracted open_target: {param}")
                return param
            print("No open_target extracted")
            return None
        if param_type == "number":
            # 1) Explicit pattern for specific commands (e.g., switch tab N)
            match = re.search(r'switch tab (\d+)', cmd_text.lower())
            if match:
                param = match.group(1)
                print(f"Extracted number: {param}")
                return param
            # 2) Plain digits in the command
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            if match:
                param = match.group(1)
                print(f"Extracted number: {param}")
                return param
            # 3) Spoken numbers (zero–ninety nine) and ordinals (first–twentieth)
            words_map = {
                'zero':0,'one':1,'two':2,'three':3,'four':4,'five':5,'six':6,'seven':7,'eight':8,'nine':9,
                'ten':10,'eleven':11,'twelve':12,'thirteen':13,'fourteen':14,'fifteen':15,'sixteen':16,
                'seventeen':17,'eighteen':18,'nineteen':19,'twenty':20,'thirty':30,'forty':40,'fifty':50,
                'sixty':60,'seventy':70,'eighty':80,'ninety':90
            }
            # Try ordinal words via class map (1-based)
            try:
                for ord_word, val in self.ORDINAL_WORDS.items():
                    if re.search(fr"\b{ord_word}\b", cmd_text, re.IGNORECASE):
                        print(f"Extracted number (ordinal word): {val}")
                        return str(val)
            except Exception:
                pass
            # Try direct word match
            for w, v in words_map.items():
                if re.search(fr"\b{w}\b", cmd_text, re.IGNORECASE):
                    print(f"Extracted number (word): {v}")
                    return str(v)
            # Try hyphenated or spaced tens + ones (e.g., 'twenty one')
            m = re.search(r"\b(twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)[ -](one|two|three|four|five|six|seven|eight|nine)\b", cmd_text, re.IGNORECASE)
            if m:
                val = words_map[m.group(1).lower()] + words_map[m.group(2).lower()]
                print(f"Extracted number (compound words): {val}")
                return str(val)
            print("No number extracted")
            return None  # Return as string
        elif param_type == "app_name":
            # Only match if not a disk or folder command
            if not any(keyword in cmd_text for keyword in ["disk", "disc", "drive", "folder"]):
                match = re.search(r'(?:run|launch|start)\s+(.+?)(?:\s|$)', cmd_text, re.IGNORECASE)
                param = match.group(1).strip() if match else None
                print(f"Extracted app_name: {param}")
                return param
            print("No app_name extracted (disk/folder command detected)")
            return None
        elif param_type == "folder_name":
            # Capture full folder name, including spaces, until end of string
            patterns = [
                r'(?:create|open|delete)\s+folder\s+(.+)$',  # e.g., "open folder my work"
                r'(?:access|go to)\s+folder\s+(.+)$'  # Synonyms
            ]
            for pattern in patterns:
                match = re.search(pattern, cmd_text, re.IGNORECASE)
                if match:
                    param = match.group(1).strip()
                    print(f"Extracted folder_name: {param}")
                    return param
            print("No folder_name extracted")
            return None
        elif param_type == "old_name_new_name":
            # Capture old and new names, including spaces, until 'to' and end
            match = re.search(r'rename\s+folder\s+(.+?)\s+to\s+(.+)$', cmd_text, re.IGNORECASE)
            if match:
                param = (match.group(1).strip(), match.group(2).strip())
                print(f"Extracted old_name_new_name: {param}")
                return param
            print("No old_name_new_name extracted")
            return (None, None)
        elif param_type == "disk_letter":
            # Capture disk letter from phrases like 'open/go to/access disk|disc|drive C:' or 'open C drive'
            match = re.search(r'(?:open|access|go to)\s+(?:disk|disc|drive)\s+([a-zA-Z]):?(?:\s|$)', cmd_text, re.IGNORECASE)
            if match:
                param = match.group(1).strip().upper()
                print(f"Extracted disk_letter: {param}")
                return param
            match = re.search(r'(?:open|access|go to)\s+([a-zA-Z])\s*(?:drive|disk|disc)(?:\s|$)', cmd_text, re.IGNORECASE)
            if match:
                param = match.group(1).strip().upper()
                print(f"Extracted disk_letter (reversed): {param}")
                return param
            print("No disk_letter extracted")
            return None
        elif param_type == "numbers":
            # Extract numbers from the command (e.g., "2 by 2" or "4 and 3")
            numbers = re.findall(r'\b\d+\b', cmd_text)
            if len(numbers) >= 2:
                param = numbers  # Return list of numbers as strings
                print(f"Extracted numbers: {param}")
                return param
            print("Not enough numbers extracted")
            return None
        elif param_type == "city_name":
            # Extract city name after "tell weather in", "weather in", or synonyms
            patterns = [
                r'(?:tell\s+weather\s+in|weather\s+in|what\'s\s+the\s+weather\s+in|weather\s+forecast\s+for|weather)\s+(.+)$',
                r'tell\s+weather\s+(.+?)(?:\s+weather|$)'
            ]
            for pattern in patterns:
                match = re.search(pattern, cmd_text, re.IGNORECASE)
                if match:
                    param = match.group(1).strip()
                    print(f"Extracted city_name: {param}")
                    return param
            print("No city_name extracted")
            return None
        elif param_type == "bmi_data":
            # Support multiple formats: "i am 70 kg 1.75 m", "check bmi in 20 kg 1.2 m", "bmi 65 kg 1.8 m"
            weight_match = re.search(r'(\d+\.?\d*)\s*kg', cmd_text)
            height_match = re.search(r'(\d+\.?\d*)\s*(?:m|metre|meter)', cmd_text)
            if weight_match and height_match:
                weight = float(weight_match.group(1))
                height = float(height_match.group(1))
                param = (weight, height)
                print(f"Extracted bmi_data: weight {weight} kg, height {height} m")
                return param
            print("No valid weight or height extracted for BMI")
            return None
        elif param_type == "image_file":
            match = re.search(r'set\s+wallpaper\s+to\s+(.+?)(?:\s|$)', cmd_text.lower())
            if match:
               param = match.group(1).strip()
               print(f"Extracted image_file: {param}")
               return param
            print("No image file extracted")
            return None
        elif param_type == "seconds":
            match = re.search(r'countdown (\d+)', cmd_text.lower())
            if match:
                param = match.group(1)
                print(f"Extracted seconds: {param}")
                return param
            print("No seconds extracted")
            return None
        elif param_type == "text":
            match = re.search(r'spell (.+)', cmd_text.lower())
            if match:
                param = match.group(1).strip()
                print(f"Extracted text: {param}")
                return param
            print("No text extracted")
            return None
        elif param_type == "text":
            match = re.search(r'spell (.+)', cmd_text.lower())
            if match:
                param = match.group(1).strip()
                print(f"Extracted text: {param}")
                return param
            print("No text extracted")
            return None
        elif param_type == "number":
            # For switch tab, extract the tab number if provided
            match = re.search(r'switch tab (\d+)', cmd_text.lower())
            if match:
                param = match.group(1)
                print(f"Extracted number: {param}")
                return param
            print("No number extracted")
            return None
        elif param_type == "query":
           
            match = re.search(r'(?:play|search|play video|play song|play music)\s+(.+?)\s+(?:on\s+youtube|youtube)$', cmd_text.lower())
            if match:
                param = match.group(1).strip()
                print(f"Extracted query: {param}")
                return param
            # Handles: 'play on youtube <query>'
            match = re.search(r'(?:play (?:on )?youtube|search youtube|play video|play song|play music on youtube)\s+(.+)', cmd_text.lower())
            if match:
                param = match.group(1).strip()
                print(f"Extracted query: {param}")
                return param
            # fallback for "search"
            match = re.search(r'search (.+)', cmd_text.lower())
            if match:
                param = match.group(1).strip()
                print(f"Extracted query: {param}")
                return param
            
            # --- FIX: More flexible ChatGPT query extraction ---
            # This handles cases like "ask chatgpt hello" and "hello chatgpt".
            # It also supports "chat gpt" with a space and just "gpt" as a suffix.
            triggers = ["ask chatgpt", "tell chatgpt", "on chatgpt", "send to chatgpt", "chat gpt", "on gpt", "chatgpt", "gpt"]
            text_lower = cmd_text.lower().strip()
            
            # Find the longest matching trigger to avoid partial matches
            triggers.sort(key=len, reverse=True)
            for trigger in triggers:
                # Handle prefix: "ask chatgpt write a poem"
                if text_lower.startswith(trigger + " "):
                    return text_lower[len(trigger):].strip()
                # Handle suffix: "write a poem on chat gpt"
                if text_lower.endswith(" " + trigger):
                    return text_lower[:-len(trigger)].strip()

            print("No query extracted")
            return None
        elif param_type == "topic":
            # --- FIX: Handle flexible "on word" commands first ---
            # If the command ends with a "word" trigger, the topic is everything before it.
            text_lower = cmd_text.lower().strip()
            word_triggers = [" on word", " in word"]
            for trigger in word_triggers:
                if text_lower.endswith(trigger):
                    topic = text_lower[:-len(trigger)].strip()
                    # The topic itself might be a prompt, like "write a fees application"
                    print(f"Extracted topic for Word (suffix match): '{topic}'")
                    return topic

            # Fallback for more structured "write essay on..." commands
            match = re.search(r'(?:write|compose|type)(?:.*?)(?:on|about)\s+(.+?)(?:\s+on\s+word)?$', cmd_text, re.IGNORECASE)
            if match:
                param = match.group(1).strip()
                print(f"Extracted topic: {param}")
                return param
            print("No topic extracted")
            return None
        elif param_type == "filename":
            # Extracts the filename after 'save file as', or uses the topic from context.
            match = re.search(r'(?:save file as|save this as|save as)\s+(.+)', cmd_text, re.IGNORECASE)
            if match:
                param = match.group(1).strip()
                print(f"Extracted filename: {param}")
                return param
            print("No explicit filename extracted, will use context or default.")
            return None # No explicit filename, handler will use context.
        elif param_type == "nth_email":
            # Match 'read [the] {number|ordinal} (most recent|oldest) [email]' or ordinal words
            match = re.search(r'read\s+(?:the\s+)?(\d+|[a-z]+)(?:st|nd|rd|th)?\s+(most recent|oldest)(?:\s+email)?$', cmd_text, re.IGNORECASE)
            if match:
                idx_raw = match.group(1)
                which = match.group(2)
                print(f"Extracted nth_email: idx_raw='{idx_raw}', which='{which}'")
                if idx_raw.isdigit():
                    idx = int(idx_raw) - 1
                else:
                    idx = self.ORDINAL_WORDS.get(idx_raw.lower(), None)
                    if idx is not None:
                        idx -= 1  # Convert to 0-based index
                print(f"Extracted nth_email index: {idx}, which: {which}")
                return (idx, which)
            match = re.search(r'read\s+(?:the\s+)?(most recent|oldest)(?:\s+email)?$', cmd_text, re.IGNORECASE)
            if match:
                which = match.group(1)
                print(f"Extracted which: {which}")
                return (0, which)
            print("No nth_email index extracted")
            return None
        print("No parameters extracted (unknown param_type)")
        return None
        

    def execute_command(self, cmd_text):
        """Execute commands by delegating to appropriate handlers. Returns response text or True/False."""
        cmd_text = self.preprocess_command(cmd_text)
        print(f"Processing command: {cmd_text}")

        # --- Vision Mode Command (High Priority Offline Recognition) ---
        if any(keyword in cmd_text.lower() for keyword in VISION_MODE_KEYWORDS):
            # Handles the seamless, threaded transition to the virtual mouse mode.
            # This is the final implementation.
            
            # --- Sci-Fi Transition Lines ---
            transition_speech = (
                "Engaging ocular-input matrix. "
                "Hand-off to gesture control is now active. "
                "Your movements will command the system."
            )
            return_speech = (
                "Gesture control deactivated. "
                "Re-engaging auditory command processor. "
                "Welcome back."
            )
            
            # --- Threading Logic for a Smooth Hand-off ---
            
            def launch_vision_mode():
                """This function runs the virtual mouse and waits for it to close."""
                logging.info("Virtual mouse thread started. Launching process...")
                virtual_mouse_script = "vm_gpt11.py"
                try:
                    # Use sys.executable to ensure we run with the same Python environment.
                    # The subprocess will wait here until vm_gpt11.py finishes.
                    subprocess.run([sys.executable, virtual_mouse_script], check=True, capture_output=True, text=True)
                except FileNotFoundError:
                    logging.error(f"FATAL: Could not find the script '{virtual_mouse_script}'. Make sure it's in the main project directory.")
                    self.speech.speak("Critical error: Vision mode script not found.")
                except subprocess.CalledProcessError as e:
                    logging.error(f"The vision mode process exited with an error: {e.stderr}")
                    self.speech.speak("An error occurred within vision mode.")
                except Exception as e:
                    logging.error(f"An unexpected error occurred while launching vision mode: {e}", exc_info=True)
                    self.speech.speak("A critical error occurred while starting vision mode.")

            # 1. Pause listening before doing anything else.
            self.voice_recognizer.pause_listening()
            logging.info("Voice recognition paused for vision mode hand-off.")

            # 2. Create and start the background thread for the virtual mouse.
            vision_thread = threading.Thread(target=launch_vision_mode, daemon=True)
            vision_thread.start()

            # 3. Immediately speak the transition message while the mouse loads.
            self.speech.speak(transition_speech)

            # 4. Wait for the virtual mouse process to complete.
            vision_thread.join()
            
            # 5. The user has given the thumbs-up and the script has exited.
            #    Resume listening and speak the return message.
            logging.info("Vision mode process has finished. Resuming standard operations.")
            self.voice_recognizer.resume_listening()
            
            return return_speech

        # --- Google Search Command ---
        search_match = re.match(r"(search for|google)\s+(.+)", cmd_text, re.IGNORECASE)
        if search_match:
            query = search_match.group(2)
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(url)
            return f"Searching Google for {query}"
        # --- End Google Search Command ---
        
        # Only catch standalone "stop" or "cancel" commands, not specific ones like "stop scrolling"
        if cmd_text.strip() in ["stop", "cancel"]:
            print("Stopping current operation and returning to main menu")
            return "Stopping. Ready for new command."

        command_parts = re.split(r'\s+(and|then)\s+', cmd_text)
        command_parts = [part for part in command_parts if part not in ["and", "then"]]
        response_texts = []
        executed = False

        for part in command_parts:
            part = part.strip()
            if not part:
                continue

            # Use find_command for all command matching (including pattern matching)
            cmd_name, params = self.find_command(part)
            if cmd_name and cmd_name in self.COMMANDS:
                info = self.COMMANDS[cmd_name]
                result = self._execute_handler(info, part, cmd_name)
                if isinstance(result, str):
                    response_texts.append(result)
                    executed = True
                elif result:
                    executed = True
            else:
                print(f"No command matched for: {part}")
                response_texts.append("I didn't understand that command. Please try again.")

        # Return aggregated response text if any, otherwise return execution status
        if response_texts:
            return " ".join(response_texts)
        return executed

    def _execute_handler(self, info, cmd_text, cmd_name=None):
        """Execute the handler for a matched command. Returns handler result (string or boolean)."""
        # Use params that were already extracted in find_command
        params = None
        if info["params"]:
            # For email commands, the params were already extracted in find_command
            if "email" in cmd_name:
                # The params are already available from find_command
                params = None  # We'll handle this in the specific handlers
            else:
                params = self.extract_parameters(cmd_text, info.get("params"))
        
        if info["params"] and params is None and "email" not in cmd_name:
            # Special cases: these commands can work without parameters (they have fallbacks)
            if cmd_name in ["switch tab", "save file"]:
                pass  # Allow execution without parameters
            else:
                command_display = cmd_name if cmd_name else info["handler"]
                print(f"Missing required parameter for '{command_display}'")
                # Tailored prompt per expected param type
                param_type = info.get("params")
                if param_type == "number":
                    prompt = f"Missing required parameter for {command_display}. Please provide a number."
                elif param_type == "folder_name":
                    prompt = f"Missing required parameter for {command_display}. Please provide a folder name."
                elif param_type == "city_name":
                    prompt = f"Missing required parameter for {command_display}. Please provide a city name."
                else:
                    prompt = f"Missing required parameter for {command_display}."
                return prompt  # Return error message as string
        
        handler_module = info["handler_module"]
        handler_name = info["handler"]
        result = None
        
        if handler_module == "file":
            handler_func = getattr(self.file_handler, handler_name)
            if info["params"] == "folder_name":
                result = handler_func(params, self.context)
            elif info["params"] == "old_name_new_name":
                old_name, new_name = params
                result = handler_func(old_name, new_name, self.context)
            elif info["params"] == "disk_letter":
                result = handler_func(params, self.context)
            else:
                result = handler_func(self.context)  # For commands like open_my_computer, go_back
        elif handler_module == "os":
            handler_func = getattr(self.os_handler, handler_name)
            result = handler_func(params if info["params"] else None)
        elif handler_module == "general":
            handler_func = getattr(self.general_handler, handler_name)
            # For email commands, pass the cmd_text so the handler can extract params
            if "email" in cmd_name:
                result = handler_func(cmd_text)
            else:
                # Pass both params and the full text to handlers that might need context
                if cmd_name == "write essay":
                    result = handler_func(params, cmd_text)
                else:
                    result = handler_func(params if info["params"] else None)
        
        # Return the handler result (could be string, boolean, or None)
        # If result is None or True, return a generic success message
        if result is not None and isinstance(result, str):
            return result
        elif result:
            return "Command executed successfully."
        else:
            return "Command execution failed."

    def handle_read_most_recent_email(self):
        print("handle_read_most_recent_email called")
        self.handle_read_nth_email_index(0, reverse=False)

    def handle_read_oldest_email(self):
        print("handle_read_oldest_email called")
        self.handle_read_nth_email_index(0, reverse=True)

    def handle_read_nth_most_recent_email(self, cmd_text):
        print(f"handle_read_nth_most_recent_email called with cmd_text: {cmd_text}")
        # Extract parameters from the command text
        params = self.extract_parameters(cmd_text, "nth_email")
        if params and isinstance(params, tuple):
            idx, _ = params
            print(f"Extracted index: {idx}")
            self.handle_read_nth_email_index(idx, reverse=False)
        else:
            print("Failed to extract nth_email parameters")
            self.file_manager.speech.speak("I didn't understand which email you want to read.")

    def handle_read_nth_oldest_email(self, cmd_text):
        print(f"handle_read_nth_oldest_email called with cmd_text: {cmd_text}")
        # Extract parameters from the command text
        params = self.extract_parameters(cmd_text, "nth_email")
        if params and isinstance(params, tuple):
            idx, _ = params
            print(f"Extracted index: {idx}")
            self.handle_read_nth_email_index(idx, reverse=True)
        else:
            print("Failed to extract nth_email parameters")
            self.file_manager.speech.speak("I didn't understand which email you want to read.")

    def handle_read_nth_email_index(self, idx, reverse=False):
        print(f"handle_read_nth_email_index called with idx: {idx}, reverse: {reverse}")
        if idx is None or not isinstance(idx, int) or idx < 0:
            self.file_manager.speech.speak("I didn't understand which email you want to read.")
            print("Invalid index for nth email.")
            return
        try:
            service = self.gmail_service
            print("Fetching messages from Gmail...")
            # Fetch up to 500 emails to allow for large N
            results = service.users().messages().list(userId='me', maxResults=max(idx+1, 500)).execute()
            messages = results.get('messages', [])
            print(f"Fetched {len(messages)} messages.")
            if not messages or len(messages) <= idx:
                self.file_manager.speech.speak("There aren't that many emails in your inbox.")
                print("Not enough emails.")
                return
            if reverse:
                messages = list(reversed(messages))
            msg_id = messages[idx]['id']
            print(f"Fetching message with id: {msg_id}")
            msg_data = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
            headers = msg_data['payload'].get('headers', [])
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '(No Subject)')
            from_ = next((h['value'] for h in headers if h['name'] == 'From'), '(Unknown Sender)')
            snippet = msg_data.get('snippet', '')
            self.context['last_read_email'] = {'id': msg_id, 'subject': subject}
            self.file_manager.speech.speak(f"Email from {from_}, subject: {subject}. {snippet}")
            print(f"Read email: From: {from_}, Subject: {subject}, Snippet: {snippet}")
            url = f"https://mail.google.com/mail/u/0/#inbox/{msg_id}"
            webbrowser.open(url)
        except Exception as e:
            self.file_manager.speech.speak(f"Failed to read email: {e}")
            print(f"Exception in handle_read_nth_email_index: {e}")
