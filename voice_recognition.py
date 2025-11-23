# voice_recognition.py

import threading
import queue
import time
import json
import socket
import re
import logging
import os
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass

# Third-party libraries
import speech_recognition as sr
import numpy as np
import sounddevice as sd
import torch  # For prompt biasing tensor operations


# Project-specific imports
# No longer need FileManager for offline commands

# Attempt to import OpenVINO Whisper for offline support
try:
    from optimum.intel.openvino import OVModelForSpeechSeq2Seq
    from transformers import AutoProcessor, pipeline
    OPENVINO_AVAILABLE = True
except ImportError:
    OPENVINO_AVAILABLE = False
    logging.warning("Optimum-Intel not installed. OpenVINO-based offline mode will not be available.")

# Unified flag: offline engine availability
OFFLINE_AVAILABLE = OPENVINO_AVAILABLE

# Special marker used to signal that online STT should fail over to offline
SWITCH_TO_OFFLINE_SIGNAL = "__STT_SWITCH_OFFLINE__"

# --- Configuration ---
ASSEMBLYAI_KEY = "d5a709c0d7c74944b75b91904b86405a"  # Your AssemblyAI API Key

# --- OpenVINO Engine Configuration ---
@dataclass
class OpenVINOConfig:
    """Configuration for the OpenVINO Whisper STT Engine"""
    model_path: str = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models", "distil_small_openvino")
    sample_rate: int = 16000
    # Slightly longer chunks to capture phrases like "set volume 50" without cutting off the number
    chunk_duration_seconds: float = 1.5
    # Lowered slightly to catch quieter speech
    silence_threshold: float = 0.015  # Energy threshold for simple VAD
    # VAD parameters
    vad_onset: float = 0.5
    vad_offset: float = 0.35

# --- Offline Vocabulary for Vosk (strict command bias) ---
OFFLINE_COMMAND_PHRASES: List[str] = [
    "create folder", "open folder", "delete folder", "rename folder",
    "open my computer", "open this pc", "go back", "open disk", "access drive",
    "show desktop", "minimize all windows", "restore windows",
    "switch window", "maximize window", "minimize window", "close window",
    "move window left", "move window right", "snap window", "go to desktop",
    "new tab", "close tab", "switch tab", "next tab", "previous tab",
    "refresh page", "reload", "zoom in", "zoom out", "search google",
    "open youtube", "play video", "pause video", "search for",
    "increase volume", "decrease volume", "set volume", "mute volume", "unmute",
    "maximize volume", "increase brightness", "decrease brightness", "set brightness",
    "take screenshot", "take photo", "open camera",
    "tell time", "tell date", "tell day", "tell weather", "tell joke",
    "countdown", "set timer",
    "open calculator", "open notepad", "open word", "run application",
    "read clipboard", "summarize clipboard", "copy", "paste", "select all", "undo", "redo",
    "show grid", "hide grid", "click cell", "double click cell", "right click cell",
    "drag from", "drop on", "zoom cell", "exit zoom", "set grid size",
    "stop listening", "start listening", "exit system", "shutdown", "restart",
    "hello", "hey assistant", "wake up", "sleep", "cancel", "stop",
    "start dictation", "stop dictation", "take a note", "start dictation mode", "stop dictation mode",
    # LLM-style triggers to bias Vosk towards conversational queries when offline
    "can you tell me", "tell me about", "can you explain", "explain",
    "what is", "who is", "why is", "how to", "how do",
    "write a", "write an", "translate", "summarize", "let's talk", "assistant",
]

OFFLINE_NUMBER_WORDS: List[str] = [
    "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
    "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen",
    "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred",
    "0", "10", "20", "30", "40", "50", "60", "70", "80", "90", "100",
]

# Default Vosk model location inside the project (user-specific absolute path avoided)
VOSK_MODEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models", "vosk-model-en-in-0.5")

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
            # Expanded numeric vocabulary to improve commands like "set volume 50"
            "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
            "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety", "hundred",
            "0", "10", "20", "30", "40", "50", "60", "70", "80", "90", "100",
        ]
        
        prompt_parts = primary_commands + secondary_keywords
        # Slightly larger prompt window so the extra numeric keywords are included
        prompt = "Commands: " + ", ".join(prompt_parts[:80])
        return prompt

# --- OpenVINO Whisper Offline STT Engine ---
class OpenVINOWhisperSTT:
    """Offline STT using OpenVINO-optimized Whisper for fast, hardware-accelerated inference."""
    
    def __init__(self, config: Optional[OpenVINOConfig] = None):
        if not OPENVINO_AVAILABLE:
            raise RuntimeError("Cannot initialize OpenVINOWhisperSTT: optimum-intel is not installed.")
        
        self.config = config or OpenVINOConfig()
        
        # Verify model path exists
        if not os.path.isdir(self.config.model_path):
            raise RuntimeError(f"OpenVINO model directory not found at '{self.config.model_path}'")
        
        self.parser = CommandParser()
        self.audio_queue = queue.Queue()
        self.transcription_queue = queue.Queue()
        self.is_running = False
        self.is_paused = True
        self.model = None
        self.processor = None
        self.command_prompt = self.parser.generate_command_prompt()
        self._processing_thread: Optional[threading.Thread] = None
        
        self.recent_commands = []
        self.max_recent_commands = 5
        
        # Mode for switching between command and dictation
        self.mode = "COMMAND"
        
        logging.info(f"OpenVINO Whisper STT initialized with model at: {self.config.model_path}")
    
    def start(self):
        logging.info("Starting OpenVINO Whisper STT Engine...")
        self._load_model()
        self.is_running = True
        self.is_paused = False
        
        self._processing_thread = threading.Thread(target=self._process_audio_loop, daemon=True)
        self._processing_thread.start()
        
        threading.Thread(target=self._audio_stream_loop, daemon=True).start()
        logging.info("✅ OpenVINO Whisper STT Engine started successfully")
    
    def stop(self):
        if self.is_running:
            logging.info("Stopping OpenVINO Whisper STT Engine...")
            self.is_running = False
    
    def pause(self):
        logging.info("OpenVINO Whisper STT pause requested; clearing audio queue.")
        self.is_paused = True
        with self.audio_queue.mutex:
            self.audio_queue.queue.clear()
    
    def resume(self):
        logging.info("OpenVINO Whisper STT resume requested.")
        self.is_paused = False
    
    def set_mode(self, mode: str):
        """Switch between 'COMMAND' and 'DICTATION' modes."""
        mode_upper = (mode or "").upper()
        if mode_upper == "DICTATION":
            if self.mode != "DICTATION":
                logging.info("OpenVINO Whisper STT switched to DICTATION mode.")
            self.mode = "DICTATION"
        else:
            if self.mode != "COMMAND":
                logging.info("OpenVINO Whisper STT switched to COMMAND mode.")
            self.mode = "COMMAND"
    
    def _load_model(self):
        logging.info(f"Loading OpenVINO Whisper model from: {self.config.model_path}")
        try:
            # Load OpenVINO-optimized model
            self.model = OVModelForSpeechSeq2Seq.from_pretrained(
                self.config.model_path,
                compile=True
            )
            
            # Load processor (tokenizer + feature extractor)
            self.processor = AutoProcessor.from_pretrained(self.config.model_path)
            
            logging.info("✅ OpenVINO Whisper model loaded successfully with hardware acceleration")
        except Exception as e:
            logging.error(f"Failed to load OpenVINO Whisper model: {e}")
            raise
    
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
            logging.error(f"Audio stream failed: {e}. OpenVINO engine will stop.")
            self.is_running = False
    
    def _audio_callback(self, indata, frames, time_info, status):
        if status:
            logging.warning(f"Audio callback status: {status}")
        if self.is_running and not self.is_paused:
            self.audio_queue.put(indata[:, 0].copy())
    
    def _get_contextual_prompt(self) -> str:
        """Generate a context-aware prompt based on recent commands.
        In offline mode we avoid a static command list to prevent hallucination.
        """
        # Only include recent commands if any; otherwise return empty string
        if self.recent_commands:
            recent_str = "Recent: " + ", ".join(self.recent_commands[-3:])
            return recent_str
        return ""
    
    def _process_audio_loop(self):
        """Process audio from queue using VAD to detect complete utterances."""
        logging.info("OpenVINO audio processing loop started.")
        
        speech_buffer = []
        silence_chunks = 0
        max_silence_chunks = 15  # ~1.5 seconds of silence to end utterance
        min_speech_chunks = 3    # Minimum chunks before considering it speech
        
        while self.is_running:
            try:
                if self.is_paused:
                    time.sleep(0.1)
                    continue
                
                # Get audio chunk from queue
                audio_chunk = self.audio_queue.get(timeout=0.5)
                
                # Calculate energy for VAD
                audio_energy = np.abs(audio_chunk).mean()
                
                # Speech detection
                if audio_energy > self.config.silence_threshold:
                    # Speech detected
                    speech_buffer.append(audio_chunk)
                    silence_chunks = 0
                else:
                    # Silence detected
                    if len(speech_buffer) > 0:
                        # We have speech in buffer, this is silence after speech
                        silence_chunks += 1
                        speech_buffer.append(audio_chunk)  # Include some silence
                        
                        # Check if we have enough silence to consider utterance complete
                        if silence_chunks >= max_silence_chunks and len(speech_buffer) >= min_speech_chunks:
                            # Transcribe the complete utterance
                            full_audio = np.concatenate(speech_buffer)
                            buffer_duration = len(full_audio) / self.config.sample_rate
                            
                            logging.info(
                                "OpenVINO STT: complete utterance detected duration=%.2fs chunks=%d",
                                buffer_duration,
                                len(speech_buffer)
                            )
                            
                            # Process the complete utterance
                            self._transcribe_audio(full_audio)
                            
                            # Reset buffer
                            speech_buffer = []
                            silence_chunks = 0
                
            except queue.Empty:
                # Check if we have buffered speech that's been waiting too long
                if len(speech_buffer) >= min_speech_chunks:
                    silence_chunks += 1
                    if silence_chunks >= max_silence_chunks:
                        full_audio = np.concatenate(speech_buffer)
                        buffer_duration = len(full_audio) / self.config.sample_rate
                        
                        logging.info(
                            "OpenVINO STT: timeout utterance duration=%.2fs",
                            buffer_duration
                        )
                        
                        self._transcribe_audio(full_audio)
                        speech_buffer = []
                        silence_chunks = 0
                continue
            except Exception as e:
                logging.error(f"Audio processing error in OpenVINO STT: {e}")
    
    def _transcribe_audio(self, audio_data):
        """Transcribe a complete audio utterance."""
        try:
            # Prepare audio for model
            inputs = self.processor(
                audio_data,
                sampling_rate=self.config.sample_rate,
                return_tensors="pt"
            )
            
            # Generate transcription using OpenVINO model
            # Re-enabling prompt biasing with robust shape handling
            contextual_prompt = self._get_contextual_prompt()
            
            # Get base prompt IDs
            prompt_ids = self.processor.get_decoder_prompt_ids(
                task="transcribe",
                language="en"
            )
            
            # Convert to tensor if it's a list
            if not isinstance(prompt_ids, torch.Tensor):
                prompt_ids = torch.tensor([prompt_ids]) if isinstance(prompt_ids, list) else torch.tensor([[prompt_ids]])
            
            # Ensure prompt_ids is 2D: [batch_size, seq_len]
            if prompt_ids.dim() == 1:
                prompt_ids = prompt_ids.unsqueeze(0)
            elif prompt_ids.dim() == 3:
                prompt_ids = prompt_ids.squeeze(0)
            
            # FIX: Ensure batch dimension is 1 to match prompt_tokens
            if prompt_ids.shape[0] > 1:
                prompt_ids = prompt_ids[0:1]
            
            # Tokenize the contextual prompt for biasing
            if contextual_prompt:
                prompt_tokens = self.processor.tokenizer(
                    contextual_prompt,
                    add_special_tokens=False,
                    return_tensors="pt"
                )["input_ids"]
                
                # Ensure prompt_tokens is also 2D: [batch_size, seq_len]
                if prompt_tokens.dim() == 1:
                    prompt_tokens = prompt_tokens.unsqueeze(0)
                elif prompt_tokens.dim() == 3:
                    prompt_tokens = prompt_tokens.squeeze(0)
                
                # Concatenate if shapes match in dim 0 (batch size)
                if prompt_ids.shape[0] == prompt_tokens.shape[0]:
                    full_prompt_ids = torch.cat([prompt_ids, prompt_tokens], dim=1)
                else:
                    logging.warning(f"Shape mismatch in biasing: ids={prompt_ids.shape}, tokens={prompt_tokens.shape}")
                    full_prompt_ids = prompt_ids
            else:
                full_prompt_ids = prompt_ids

            # Generate transcription using OpenVINO model
            # Provide attention_mask if available to avoid warnings
            generate_kwargs = {
                "max_new_tokens": 128,
                "attention_mask": inputs.get("attention_mask", None),
            }
            predicted_ids = self.model.generate(inputs["input_features"], **generate_kwargs)

            
            # Decode transcription
            decoded_list = self.processor.batch_decode(
                predicted_ids,
                skip_special_tokens=True
            )
            
            if not decoded_list:
                return

            transcribed_text = decoded_list[0].strip()
            
            # Filter common Whisper hallucinations/noise
            hallucinations = ["you", "so", "sure", "thanks", "thank you", "subtitles by", "watching"]
            clean_text = transcribed_text.lower().strip(".,!?")
            
            if clean_text in hallucinations:
                logging.debug(f"Ignored hallucination: '{transcribed_text}'")
                return
            
            # Filter prompt regurgitation
            if clean_text.startswith("commands:") or "create folder, open folder" in clean_text:
                logging.warning(f"Ignored prompt regurgitation: '{transcribed_text}'")
                return

            if transcribed_text and len(transcribed_text) >= 2:
                self._handle_transcription(transcribed_text)
                
        except Exception as e:
            logging.error(f"Transcription error in OpenVINO STT: {e}")
    
    def _handle_transcription(self, text: str):
        logging.info(f"[OpenVINO Raw]: {text}")
        
        if self.mode == "COMMAND":
            # Parse as command
            cmd = self.parser.parse_command(text)
            if cmd:
                logging.info(f"==> [OpenVINO Command Parsed]: {cmd['command']} | Params: {cmd['parameters']}")
                self.transcription_queue.put(cmd['original_text'])
                
                # Update recent commands for context
                self.recent_commands.append(cmd['command'])
                if len(self.recent_commands) > self.max_recent_commands:
                    self.recent_commands.pop(0)
            else:
                logging.info(f"[OpenVINO Ignored]: No valid command parsed from '{text.lower()}'")
        else:
            # Dictation mode: pass through all transcriptions
            logging.info(f"[OpenVINO Dictation]: {text}")
            self.transcription_queue.put(text)
    
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

        # --- ADD THESE SETTINGS ---
        # Increase pause threshold (seconds of non-speaking audio before a phrase is considered complete)
        self.recognizer.pause_threshold = 1.2  # Default is ~0.8. Increase to allow slower speech.

        # Prevent cutting off too early in noisy environments
        self.recognizer.non_speaking_duration = 0.8

        # Dynamic energy adjustment (helps if mic volume changes)
        self.recognizer.dynamic_energy_threshold = True

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
            # Signal to the hybrid controller that we should switch to offline mode
            try:
                self.transcription_queue.put(SWITCH_TO_OFFLINE_SIGNAL)
            except Exception:
                pass


# --- Network Monitor (Original) ---
class NetworkMonitor:
    """Monitors internet connectivity and triggers callbacks on status change."""
    def __init__(self, check_interval=10):
        logging.info("Initializing Network Monitor...")
        self.is_online = self._check_internet()
        logging.info(f"🌐 Network Status on Startup: {'ONLINE' if self.is_online else 'OFFLINE'}")
        if self.is_online:
            logging.info("📡 Google Voice-to-Text will be activated")
        else:
            logging.info("📴 OpenVINO offline mode will be activated")
        self._check_interval = check_interval
        self._callbacks = []
        self._stop_event = threading.Event()
        self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._monitor_thread.start()

    def _check_internet(self):
        try:
            logging.debug("Checking internet connectivity to 8.8.8.8:53...")
            socket.create_connection(("8.8.8.8", 53), timeout=3)
            logging.info("✅ Internet connectivity detected - Google Voice-to-Text will be used")
            return True
        except OSError as e:
            logging.warning(f"❌ Internet connectivity check failed: {e}")
            logging.warning("Falling back to OpenVINO offline mode")
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
        # Offline engine is OpenVINO Whisper
        self.offline_engine: Optional[OpenVINOWhisperSTT] = None
        self.current_mode: Optional[str] = None

        # Pre-build offline engine so it is ready in RAM when needed
        self._build_offline_engine()

        self.network_monitor.register_callback(self._on_network_change)
        self._initialize_engine()
        logging.info(f"Hybrid Voice Recognizer initialized in {self.current_mode} mode.")

    def _build_offline_engine(self):
        """Instantiate the OpenVINO Whisper offline engine."""
        if OPENVINO_AVAILABLE:
            try:
                self.offline_engine = OpenVINOWhisperSTT(OpenVINOConfig())
                logging.info("✅ OpenVINO Whisper offline STT engine created successfully.")
                return
            except Exception as e:
                logging.error(f"Failed to initialize OpenVINO offline engine: {e}")
                self.offline_engine = None
        else:
            logging.warning("OpenVINO not available. Offline mode will not be available.")

    def _initialize_engine(self):
        """Starts the appropriate engine based on the current network status."""
        if self.network_monitor.is_online:
            self._start_online_engine()
        elif OFFLINE_AVAILABLE:
            self._start_offline_engine()
        else:
            logging.error("No STT engine available! Internet is offline and no offline engine could be initialized.")
            self.current_mode = "UNAVAILABLE"

    def _start_online_engine(self):
        if self.current_mode == "ONLINE":
            return
        logging.info("Switching to ONLINE recognition engine...")
        if self.offline_engine:
            # Stop offline engine but keep the instance so it remains preloaded in RAM.
            self.offline_engine.stop()
        
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

        if not self.offline_engine:
            self._build_offline_engine()

        if self.offline_engine:
            self.offline_engine.start()
            self.current_mode = "OFFLINE"
        else:
            logging.error("Failed to start offline engine; no offline STT available.")

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
        logging.info(
            "Voice recognition paused. mode=%s offline_running=%s offline_paused=%s",
            self.current_mode,
            getattr(self.offline_engine, "is_running", None),
            getattr(self.offline_engine, "is_paused", None),
        )
        if self.online_engine:
            self.online_engine.pause()
        if self.offline_engine:
            self.offline_engine.pause()
    
    def resume_listening(self):
        logging.info(
            "Voice recognition resumed. mode=%s offline_running=%s offline_paused=%s",
            self.current_mode,
            getattr(self.offline_engine, "is_running", None),
            getattr(self.offline_engine, "is_paused", None),
        )
        if self.online_engine:
            self.online_engine.resume()
        if self.offline_engine:
            self.offline_engine.resume()

    def set_mode(self, mode: str):
        """Switch offline engine between COMMAND and DICTATION modes when applicable."""
        if self.current_mode == "OFFLINE" and hasattr(self.offline_engine, "set_mode"):
            try:
                self.offline_engine.set_mode(mode)
            except Exception as e:
                logging.error(f"Failed to set offline STT mode to {mode}: {e}")

    def get_transcription(self) -> Optional[str]:
        if self.current_mode == "ONLINE":
            try:
                text = self.transcription_queue.get_nowait()
            except queue.Empty:
                return None
            # Handle signal from OnlineSTT requesting a failover to offline.
            if text == SWITCH_TO_OFFLINE_SIGNAL:
                logging.warning("Online STT requested switch to OFFLINE mode.")
                self._start_offline_engine()
                return None
            return text
        elif self.current_mode == "OFFLINE" and self.offline_engine:
            text = self.offline_engine.get_transcription()
            if text is not None:
                logging.debug("HybridVoiceRecognizer: received offline transcription '%s'", text)
            return text
        
        return None

# --- Backward Compatibility ---
VoiceRecognizer = HybridVoiceRecognizer
