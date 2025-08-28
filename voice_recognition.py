# voice_recognition.py

import threading
import queue
import time
import json
import socket
import re
import logging
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass

# Third-party libraries
import speech_recognition as sr
import numpy as np
import sounddevice as sd


# Project-specific imports
# No longer need FileManager for offline commands

# Attempt to import Faster Whisper for offline support
try:
    from faster_whisper import WhisperModel
    OFFLINE_AVAILABLE = True
except ImportError:
    OFFLINE_AVAILABLE = False
    logging.warning("Faster Whisper not installed. Offline mode will not be available.")

# --- Configuration ---
ASSEMBLYAI_KEY = "d5a709c0d7c74944b75b91904b86405a"  # Your AssemblyAI API Key

# --- Offline Engine Configuration (from user's script) ---
@dataclass
class OfflineConfig:
    """Configuration for the JARVIS STT Engine"""
    model_name: str = "base.en"
    sample_rate: int = 16000
    chunk_duration_seconds: float = 3.0
    silence_threshold: float = 0.02  # Energy threshold for simple VAD
    cpu_threads: int = 4
    compute_type: str = "int8"

# --- Dynamic Command Parser (from user's script) ---
class CommandParser:
    """Parses transcribed text against a strict set of commands defined by GBNF-style grammar rules."""
    def __init__(self):
        # This is the comprehensive command list from the user's script
        self.commands = {
            'file_folder': {
                'create_folder': {'patterns': [r'create\s+folder\s+([\w\s]+)', r'make\s+folder\s+([\w\s]+)', r'new\s+folder\s+([\w\s]+)'], 'params': ['folder_name']},
                'open_folder': {'patterns': [r'open\s+folder\s+([\w\s]+)', r'access\s+folder\s+([\w\s]+)', r'go\s+to\s+folder\s+([\w\s]+)'], 'params': ['folder_name']},
                'delete_folder': {'patterns': [r'delete\s+folder\s+([\w\s]+)', r'remove\s+folder\s+([\w\s]+)', r'delete\s+directory\s+([\w\s]+)'], 'params': ['folder_name']},
                'rename_folder': {'patterns': [r'rename\s+folder\s+([\w\s]+)\s+to\s+([\w\s]+)', r'change\s+folder\s+name\s+([\w\s]+)\s+to\s+([\w\s]+)'], 'params': ['old_name', 'new_name']},
                'open_my_computer': {'patterns': [r'open\s+my\s+computer', r'open\s+this\s+pc', r'this\s+pc', r'my\s+computer'], 'params': []},
                'open_disk': {'patterns': [r'open\s+disk\s+([a-zA-Z])', r'open\s+drive\s+([a-zA-Z])', r'access\s+disk\s+([a-zA-Z])', r'open\s+local\s+disk\s+([a-zA-Z])'], 'params': ['drive_letter']},
                'go_back': {'patterns': [r'go\s+back', r'back', r'return', r'go\s+up'], 'params': []},
            },
            'audio': {
                'increase_volume': {'patterns': [r'increase\s+volume', r'volume\s+up', r'turn\s+volume\s+up', r'louder'], 'params': []},
                'decrease_volume': {'patterns': [r'decrease\s+volume', r'volume\s+down', r'turn\s+volume\s+down', r'quieter'], 'params': []},
                'mute_volume': {'patterns': [r'mute\s+volume', r'mute', r'silence'], 'params': []},
                'unmute_volume': {'patterns': [r'unmute\s+volume', r'unmute'], 'params': []},
                'maximize_volume': {'patterns': [r'maximize\s+volume', r'max\s+volume', r'full\s+volume'], 'params': []},
                'set_volume': {'patterns': [r'set\s+volume\s+(?:to\s+)?(\d+)', r'turn\s+volume\s+to\s+(\d+)', r'adjust\s+volume\s+to\s+(\d+)'], 'params': ['level']},
            },
            'brightness': {
                'increase_brightness': {'patterns': [r'increase\s+brightness', r'brightness\s+up', r'turn\s+brightness\s+up', r'brighter'], 'params': []},
                'decrease_brightness': {'patterns': [r'decrease\s+brightness', r'brightness\s+down', r'turn\s+brightness\s+down', r'dimmer'], 'params': []},
                'maximize_brightness': {'patterns': [r'maximize\s+brightness', r'max\s+brightness', r'full\s+brightness'], 'params': []},
                'set_brightness': {'patterns': [r'set\s+brightness\s+(?:to\s+)?(\d+)', r'turn\s+brightness\s+to\s+(\d+)', r'adjust\s+brightness\s+to\s+(\d+)'], 'params': ['level']},
            },
            'browser': {
                'next_tab': {'patterns': [r'next\s+tab', r'forward\s+tab', r'next\s+browser\s+tab'], 'params': []},
                'previous_tab': {'patterns': [r'previous\s+tab', r'back\s+tab', r'prev\s+tab'], 'params': []},
                'switch_tab': {'patterns': [r'switch\s+tab\s+(\d+)', r'go\s+to\s+tab\s+(\d+)', r'jump\s+to\s+tab\s+(\d+)'], 'params': ['tab_number']},
                'close_tab': {'patterns': [r'close\s+tab', r'close\s+current\s+tab', r'close\s+browser\s+tab', r'close\s+this\s+tab'], 'params': []},
                'refresh': {'patterns': [r'refresh', r'reload\s+tab', r'refresh\s+tab'], 'params': []},
                'zoom_in': {'patterns': [r'zoom\s+in', r'increase\s+zoom', r'zoom\s+bigger'], 'params': []},
                'zoom_out': {'patterns': [r'zoom\s+out', r'decrease\s+zoom', r'zoom\s+smaller'], 'params': []},
                'search': {'patterns': [r'search\s+(?:for\s+)?(.*)', r'look\s+up\s+(.*)', r'find\s+(.*)', r'google\s+(.*)'], 'params': ['query']},
                'youtube': {'patterns': [r'play\s+on\s+youtube\s+(.*)', r'youtube\s+(.*)', r'play\s+video\s+(.*)', r'search\s+youtube\s+(.*)'], 'params': ['query']},
            },
            'window': {
                'switch_window': {'patterns': [r'switch\s+window', r'next\s+window', r'change\s+window', r'switch'], 'params': []},
                'maximize_window': {'patterns': [r'maximize\s+window', r'maximize\s+this\s+window', r'full\s+screen'], 'params': []},
                'minimize_window': {'patterns': [r'minimize\s+window', r'minimize\s+this\s+window'], 'params': []},
                'close_window': {'patterns': [r'close\s+window', r'close\s+app', r'close\s+application', r'close\s+program', r'close\s+current\s+window', r'close\s+this\s+window'], 'params': []},
                'move_window_left': {'patterns': [r'move\s+window\s+left', r'snap\s+window\s+left'], 'params': []},
                'move_window_right': {'patterns': [r'move\s+window\s+right', r'snap\s+window\s+right'], 'params': []},
                'minimize_all': {'patterns': [r'minimize\s+all\s+windows', r'show\s+desktop', r'minimize\s+all'], 'params': []},
            },
            'grid': {
                'show_grid': {'patterns': [r'show\s+grid', r'show\s+grade', r'display\s+grid', r'open\s+grid', r'grid\s+on'], 'params': []},
                'hide_grid': {'patterns': [r'hide\s+grid', r'close\s+grid', r'grid\s+off', r'remove\s+grid'], 'params': []},
                'set_grid_size': {'patterns': [r'set\s+grid\s+size\s+(\d+)', r'grid\s+(\d+)'], 'params': ['size']},
                'click_cell': {'patterns': [r'click\s+cell\s+(\d+)', r'click\s+(\d+)', r'left\s+click\s+(\d+)'], 'params': ['cell_number']},
                'double_click_cell': {'patterns': [r'double\s+click\s+(?:cell\s+)?(\d+)', r'double-click\s+(\d+)'], 'params': ['cell_number']},
                'right_click_cell': {'patterns': [r'right\s+click\s+(?:cell\s+)?(\d+)', r'right-click\s+(\d+)'], 'params': ['cell_number']},
                'drag_from': {'patterns': [r'drag\s+from\s+(\d+)', r'drag\s+(\d+)', r'start\s+drag\s+(\d+)'], 'params': ['cell_number']},
                'drop_on': {'patterns': [r'drop\s+on\s+(\d+)', r'drop\s+(\d+)', r'to\s+(\d+)'], 'params': ['cell_number']},
                'zoom_cell': {'patterns': [r'zoom\s+(?:cell\s+)?(\d+)', r'zoom\s+into\s+cell\s+(\d+)'], 'params': ['cell_number']},
                'exit_zoom': {'patterns': [r'exit\s+zoom', r'exit\s+grid\s+zoom', r'back\s+from\s+zoom'], 'params': []},
            },
            'scrolling': {
                'scroll_up': {'patterns': [r'scroll\s+up', r'scroll\s+upward', r'move\s+up', r'page\s+up'], 'params': []},
                'scroll_down': {'patterns': [r'scroll\s+down', r'scroll\s+downward', r'move\s+down', r'page\s+down'], 'params': []},
                'scroll_left': {'patterns': [r'scroll\s+left', r'move\s+left', r'pan\s+left'], 'params': []},
                'scroll_right': {'patterns': [r'scroll\s+right', r'move\s+right', r'pan\s+right'], 'params': []},
                'stop_scrolling': {'patterns': [r'stop\s+scrolling', r'cancel\s+scrolling', r'halt\s+scrolling', r'stop\s+scroll'], 'params': []},
            },
            'utility': {
                'screenshot': {'patterns': [r'take\s+screenshot', r'screenshot', r'capture\s+screen'], 'params': []},
                'photo': {'patterns': [r'take\s+a\s+photo', r'capture\s+photo', r'take\s+picture', r'snap\s+a\s+photo'], 'params': []},
                'desktop': {'patterns': [r'go\s+to\s+desktop', r'show\s+desktop', r'take\s+me\s+to\s+desktop', r'open\s+desktop'], 'params': []},
                'change_wallpaper': {'patterns': [r'change\s+wallpaper', r'next\s+wallpaper', r'next\s+background', r'change\s+background'], 'params': []},
                'countdown': {'patterns': [r'countdown\s+(\d+)', r'start\s+countdown\s+(\d+)', r'timer\s+(\d+)', r'set\s+timer\s+(\d+)'], 'params': ['seconds']},
                'tell_time': {'patterns': [r'tell\s+time', r'what\s+time\s+is\s+it', r'current\s+time', r'time\s+now'], 'params': []},
                'tell_date': {'patterns': [r'tell\s+date', r'tell\s+me\s+the\s+date', r'what\s+is\s+the\s+date', r'current\s+date'], 'params': []},
                'tell_weather': {'patterns': [r'tell\s+weather\s+(?:in\s+)?(.*)', r'weather\s+(?:in\s+)?(.*)', r'weather\s+forecast\s+(?:for\s+)?(.*)'], 'params': ['city']},
            },
            'system': {
                'exit': {'patterns': [r'exit', r'quit', r'stop\s+program', r'bye', r'goodbye', r'shut\s+down', r'terminate'], 'params': []},
                'stop': {'patterns': [r'stop', r'cancel'], 'params': []},
                'help': {'patterns': [r'list\s+commands', r'help', r'commands', r'what\s+can\s+you\s+do'], 'params': []},
            }
        }
        
        self.number_words = {
            'zero': '0', 'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
            'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10',
            'eleven': '11', 'twelve': '12', 'thirteen': '13', 'fourteen': '14', 'fifteen': '15',
            'sixteen': '16', 'seventeen': '17', 'eighteen': '18', 'nineteen': '19', 'twenty': '20',
            'thirty': '30', 'forty': '40', 'fifty': '50', 'sixty': '60', 'seventy': '70',
            'eighty': '80', 'ninety': '90', 'hundred': '100',
            'first': '1', 'second': '2', 'third': '3', 'fourth': '4', 'fifth': '5',
            'sixth': '6', 'seventh': '7', 'eighth': '8', 'ninth': '9', 'tenth': '10'
        }

    def _preprocess_text(self, text: str) -> str:
        text_lower = text.lower().strip()
        # Replace number words with digits
        for word, digit in self.number_words.items():
            text_lower = re.sub(rf'\b{word}\b', digit, text_lower)
        # Remove common filler words
        text_lower = re.sub(r'\b(please|could you|can you|would you)\b', '', text_lower)
        # Normalize whitespace
        text_lower = ' '.join(text_lower.split())
        return text_lower

    def parse_command(self, text: str) -> Optional[Dict]:
        processed_text = self._preprocess_text(text)
        
        best_match = None
        best_score = 0
        
        for category, commands in self.commands.items():
            for cmd_name, cmd_info in commands.items():
                for pattern in cmd_info['patterns']:
                    match = re.search(pattern, processed_text)
                    if match:
                        match_span = match.end() - match.start()
                        score = match_span / len(processed_text) if processed_text else 0
                        
                        if score > best_score:
                            best_score = score
                            best_match = {
                                'command': cmd_name,
                                'original_text': text,
                                'parameters': match.groups()
                            }
        
        if best_match and best_score > 0.3:
            return best_match
        return None

    def generate_command_prompt(self) -> str:
        """Generate a comprehensive prompt containing all key command words to bias Whisper."""
        primary_commands = [
            "create folder", "open folder", "delete folder", "rename folder",
            "open my computer", "open disk", "go back",
            "increase volume", "decrease volume", "mute", "unmute", "set volume",
            "increase brightness", "decrease brightness", "set brightness",
            "next tab", "previous tab", "close tab", "refresh", "zoom in", "zoom out",
            "search", "youtube", "play video",
            "switch window", "maximize window", "minimize window", "close window",
            "show grid", "hide grid", "click cell", "double click", "right click",
            "drag from", "drop on", "zoom cell",
            "scroll up", "scroll down", "scroll left", "scroll right",
            "take screenshot", "take photo", "go to desktop", "change wallpaper",
            "countdown", "tell time", "tell date", "tell weather",
            "exit", "stop", "help"
        ]
        
        secondary_keywords = [
            "folder", "file", "directory", "computer", "disk", "drive",
            "volume", "sound", "audio", "mute", "louder", "quieter",
            "brightness", "screen", "display", "brighter", "dimmer",
            "tab", "browser", "window", "chrome", "edge", "refresh", "reload",
            "grid", "cell", "click", "mouse", "drag", "drop",
            "scroll", "up", "down", "left", "right",
            "screenshot", "photo", "picture", "desktop", "wallpaper",
            "time", "date", "weather", "countdown", "timer",
            "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"
        ]
        
        prompt_parts = primary_commands + secondary_keywords
        prompt = "Commands: " + ", ".join(prompt_parts[:50])
        return prompt

# --- Offline STT Engine (from user's script) ---
class OfflineSTT:
    """Main JARVIS STT Engine (continuous listening)."""
    def __init__(self, config: Optional[OfflineConfig] = None):
        if not OFFLINE_AVAILABLE:
            raise RuntimeError("Cannot initialize OfflineSTT: Faster Whisper is not installed.")
            
        self.config = config or OfflineConfig()
        self.parser = CommandParser() # The new parser is self-contained
        self.audio_queue = queue.Queue()
        self.transcription_queue = queue.Queue() # For parsed commands
        self.is_running = False
        self.is_paused = True
        self.model = None
        self.command_prompt = self.parser.generate_command_prompt()
        self._processing_thread: Optional[threading.Thread] = None
        
        self.recent_commands = []
        self.max_recent_commands = 5

    def start(self):
        logging.info("Starting Offline STT Engine...")
        self._load_model()
        self.is_running = True
        self.is_paused = False

        self._processing_thread = threading.Thread(target=self._process_audio_loop, daemon=True)
        self._processing_thread.start()

        threading.Thread(target=self._audio_stream_loop, daemon=True).start()

    def stop(self):
        if self.is_running:
            logging.info("Stopping Offline STT Engine...")
            self.is_running = False

    def pause(self):
        self.is_paused = True
        with self.audio_queue.mutex:
            self.audio_queue.queue.clear()

    def resume(self):
        self.is_paused = False

    def _load_model(self):
        logging.info(f"Loading Whisper model: {self.config.model_name}")
        self.model = WhisperModel(
            self.config.model_name, 
            device="cpu",
            compute_type=self.config.compute_type, 
            cpu_threads=self.config.cpu_threads
        )
        logging.info("Model loaded. Using command prompt for biasing.")

    def _audio_stream_loop(self):
        try:
            with sd.InputStream(
                samplerate=self.config.sample_rate,
                channels=1,
                dtype=np.float32,
                callback=self._audio_callback
            ):
                while self.is_running:
                    time.sleep(0.1)
        except Exception as e:
            logging.error(f"Audio stream failed: {e}. Offline engine will stop.")
            self.is_running = False

    def _audio_callback(self, indata, frames, time_info, status):
        if status:
            logging.warning(f"Audio callback status: {status}")
        if self.is_running and not self.is_paused:
            self.audio_queue.put(indata[:, 0].copy())

    def _get_contextual_prompt(self) -> str:
        """Generate a context-aware prompt based on recent commands."""
        base_prompt = self.command_prompt
        if self.recent_commands:
            recent_str = "Recent: " + ", ".join(self.recent_commands[-3:])
            return f"{recent_str}. {base_prompt}"
        return base_prompt

    def _process_audio_loop(self):
        audio_buffer = []
        while self.is_running:
            try:
                if self.is_paused:
                    time.sleep(0.1)
                    continue

                chunk = self.audio_queue.get(timeout=1.0)
                audio_buffer.append(chunk)
                
                buffer_duration = sum(len(c) for c in audio_buffer) / self.config.sample_rate
                if buffer_duration < self.config.chunk_duration_seconds:
                    continue
                
                full_audio = np.concatenate(audio_buffer)
                audio_buffer = []
                
                audio_energy = np.sqrt(np.mean(full_audio**2))
                if audio_energy < self.config.silence_threshold:
                    continue
                
                contextual_prompt = self._get_contextual_prompt()
                
                segments, info = self.model.transcribe(
                    full_audio, 
                    beam_size=3,
                    best_of=3,
                    temperature=0.0,
                    vad_filter=True,
                    vad_parameters=dict(
                        threshold=0.5,
                        min_speech_duration_ms=250,
                        max_speech_duration_s=10
                    ),
                    language="en",
                    initial_prompt=contextual_prompt,
                    suppress_blank=True,
                    suppress_tokens=[-1],
                    condition_on_previous_text=False
                )
                
                transcribed_text = " ".join(seg.text.strip() for seg in segments)
                
                if not transcribed_text or len(transcribed_text) < 2:
                    continue
                
                self._handle_transcription(transcribed_text)
                
            except queue.Empty:
                continue
            except Exception as e:
                logging.error(f"Audio processing error: {e}")

    def _handle_transcription(self, text: str):
        logging.info(f"[Offline Raw]: {text}")
        cmd = self.parser.parse_command(text)
        if cmd:
            logging.info(f"==> [Offline Command Parsed]: {cmd['command']} | Params: {cmd['parameters']}")
            self.transcription_queue.put(cmd['original_text']) # Put original text for CommandHandler
            
            # Update recent commands for context
            self.recent_commands.append(cmd['command'])
            if len(self.recent_commands) > self.max_recent_commands:
                self.recent_commands.pop(0)
        else:
            logging.info(f"[Offline Ignored]: No valid command parsed from '{text.lower()}'")

    def get_transcription(self) -> Optional[str]:
        try:
            return self.transcription_queue.get_nowait()
        except queue.Empty:
            return None

# --- Online STT Engine (Original) ---
class OnlineSTT:
    """Handles online speech recognition using Google Web Speech API."""
    def __init__(self, transcription_queue: queue.Queue):
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone(sample_rate=16000)
        self.transcription_queue = transcription_queue
        self._stop_listening = None
        self.is_paused = True

    def start(self):
        with self.microphone as source:
            logging.info("Calibrating for ambient noise (online)...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
        
        logging.info("Starting online background listening...")
        self._stop_listening = self.recognizer.listen_in_background(
            self.microphone, self._online_callback
        )
        self.is_paused = False

    def stop(self):
        if self._stop_listening:
            self._stop_listening(wait_for_stop=False)
            logging.info("Online background listening stopped.")

    def pause(self):
        self.is_paused = True

    def resume(self):
        self.is_paused = False

    def _online_callback(self, recognizer, audio):
        if self.is_paused:
            return
        try:
            text = recognizer.recognize_google(audio)
            if text:
                logging.info(f"[Online Transcribed]: '{text}'")
                self.transcription_queue.put(text)
        except sr.UnknownValueError:
            pass # Ignore silence
        except sr.RequestError as e:
            logging.error(f"Google API request failed; {e}")


# --- Network Monitor (Original) ---
class NetworkMonitor:
    """Monitors internet connectivity and triggers callbacks on status change."""
    def __init__(self, check_interval=10):
        self.is_online = self._check_internet()
        self._check_interval = check_interval
        self._callbacks = []
        self._stop_event = threading.Event()
        self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._monitor_thread.start()

    def _check_internet(self):
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            return True
        except OSError:
            return False

    def _monitor_loop(self):
        while not self._stop_event.is_set():
            new_status = self._check_internet()
            if new_status != self.is_online:
                logging.info(f"Network status changed to: {'ONLINE' if new_status else 'OFFLINE'}")
                self.is_online = new_status
                for callback in self._callbacks:
                    callback(self.is_online)
            time.sleep(self._check_interval)

    def register_callback(self, callback):
        self._callbacks.append(callback)

    def stop(self):
        self._stop_event.set()


# --- Main Hybrid Voice Recognizer (Adapted) ---
class HybridVoiceRecognizer:
    """
    Orchestrates between online and offline engines based on network availability.
    """
    def __init__(self):
        self.transcription_queue = queue.Queue()
        self.network_monitor = NetworkMonitor()
        
        self.online_engine: Optional[OnlineSTT] = None
        self.offline_engine: Optional[OfflineSTT] = None
        self.current_mode: Optional[str] = None

        self.network_monitor.register_callback(self._on_network_change)
        self._initialize_engine()
        logging.info(f"Hybrid Voice Recognizer initialized in {self.current_mode} mode.")

    def _initialize_engine(self):
        """Starts the appropriate engine based on the current network status."""
        if self.network_monitor.is_online:
            self._start_online_engine()
        elif OFFLINE_AVAILABLE:
            self._start_offline_engine()
        else:
            logging.error("No STT engine available! Internet is offline and Faster Whisper is not installed.")
            self.current_mode = "UNAVAILABLE"

    def _start_online_engine(self):
        if self.current_mode == "ONLINE":
            return
        logging.info("Switching to ONLINE recognition engine...")
        if self.offline_engine:
            self.offline_engine.stop()
            self.offline_engine = None
        
        self.online_engine = OnlineSTT(self.transcription_queue)
        self.online_engine.start()
        self.current_mode = "ONLINE"

    def _start_offline_engine(self):
        if self.current_mode == "OFFLINE":
            return
        logging.info("Switching to OFFLINE recognition engine...")
        if self.online_engine:
            self.online_engine.stop()
            self.online_engine = None
        
        # The new OfflineSTT is self-contained and doesn't need command_mappings
        self.offline_engine = OfflineSTT(OfflineConfig())
        self.offline_engine.start()
        self.current_mode = "OFFLINE"

    def _on_network_change(self, is_online: bool):
        """Callback for network status changes."""
        if is_online:
            self._start_online_engine()
        else:
            if OFFLINE_AVAILABLE:
                self._start_offline_engine()
            else:
                logging.warning("Lost internet connection, and no offline engine is available.")
                self.current_mode = "UNAVAILABLE"

    # --- Public API for main.py ---
    def start_listening(self):
        logging.info("Voice recognition system is active.")
        self.resume_listening()

    def stop_listening(self):
        logging.info("Stopping all voice recognition engines.")
        if self.online_engine:
            self.online_engine.stop()
        if self.offline_engine:
            self.offline_engine.stop()
        self.network_monitor.stop()

    def pause_listening(self):
        logging.info("Voice recognition paused.")
        if self.online_engine:
            self.online_engine.pause()
        if self.offline_engine:
            self.offline_engine.pause()
    
    def resume_listening(self):
        logging.info("Voice recognition resumed.")
        if self.online_engine:
            self.online_engine.resume()
        if self.offline_engine:
            self.offline_engine.resume()

    def get_transcription(self) -> Optional[str]:
        if self.current_mode == "ONLINE":
            try:
                return self.transcription_queue.get_nowait()
            except queue.Empty:
                return None
        elif self.current_mode == "OFFLINE" and self.offline_engine:
            return self.offline_engine.get_transcription()
        
        return None

# --- Backward Compatibility ---
VoiceRecognizer = HybridVoiceRecognizer
