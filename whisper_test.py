#!/usr/bin/env python3
"""
J.A.R.V.I.S Speech-to-Text System
Continuous listening mode (no wake word).
"""

import re
import time
import threading
import queue
import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel
from typing import Dict, Optional
from dataclasses import dataclass
import logging

# --- Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class JARVISConfig:
    """Configuration for the JARVIS STT Engine"""
    model_name: str = "base.en"
    sample_rate: int = 16000
    chunk_duration_seconds: float = 3.0
    silence_threshold: float = 0.02
    cpu_threads: int = 4
    compute_type: str = "int8"

class GBNFCommandParser:
    """Parses transcribed text against a strict set of commands defined by GBNF-style grammar rules."""
    def __init__(self):
        self.commands = {
            'system': {
                'open_my_computer': {'patterns': [r'open\s+my\s+computer', r'open\s+this\s+pc'], 'params': []},
                'go_back': {'patterns': [r'go\s+back', r'back'], 'params': []},
            },
            'window': {
                'close_window': {'patterns': [r'close\s+window', r'close\s+app'], 'params': []},
                'switch_window': {'patterns': [r'switch\s+window', r'next\s+window'], 'params': []},
            },
            'grid': {
                'show_grid': {'patterns': [r'show\s+grid', r'show\s+grade'], 'params': []},
                'hide_grid': {'patterns': [r'hide\s+grid', r'close\s+grid'], 'params': []},
                'click_cell': {'patterns': [r'click\s+cell\s+(\d+)', r'click\s+(\d+)'], 'params': ['cell_number']},
            },
        }
        self.number_words = {
            'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
            'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10',
            'fifteen': '15', 'thirty': '30'
        }

    def _preprocess_text(self, text: str) -> str:
        text_lower = text.lower().strip()
        for word, digit in self.number_words.items():
            text_lower = re.sub(rf'\b{word}\b', digit, text_lower)
        return text_lower

    def parse_command(self, text: str) -> Optional[Dict]:
        processed_text = self._preprocess_text(text)
        for category, commands in self.commands.items():
            for cmd_name, cmd_info in commands.items():
                for pattern in cmd_info['patterns']:
                    match = re.search(pattern, processed_text)
                    if match:
                        params = dict(zip(cmd_info['params'], match.groups()))
                        return {
                            'category': category,
                            'command': cmd_name,
                            'original_text': text,
                            'parameters': params,
                            'timestamp': time.time()
                        }
        return None

    def generate_command_prompt(self) -> str:
        prompt_phrases = []
        for _, commands in self.commands.items():
            for _, cmd_info in commands.items():
                phrase = cmd_info['patterns'][0].split(r'\s+')[0]
                if '(' not in phrase:
                    prompt_phrases.append(phrase)
        prompt_phrases.extend(["create", "open", "delete", "rename", "volume", "brightness", "show", "hide", "click", "countdown"])
        return ", ".join(sorted(list(set(prompt_phrases)))) + "."

class JARVIS_STT:
    """Main JARVIS STT Engine (continuous listening)."""
    def __init__(self, config: Optional[JARVISConfig] = None):
        self.config = config or JARVISConfig()
        self.parser = GBNFCommandParser()
        self.audio_queue = queue.Queue()
        self.command_queue = queue.Queue()
        self.is_running = False
        self.model = None
        self.command_prompt = self.parser.generate_command_prompt()

    def _load_model(self):
        logger.info(f"Loading Whisper model: {self.config.model_name}")
        self.model = WhisperModel(
            self.config.model_name, device="cpu",
            compute_type=self.config.compute_type, cpu_threads=self.config.cpu_threads
        )
        logger.info("Model loaded. Using command prompt for biasing.")

    def _audio_callback(self, indata, frames, time_info, status):
        if status:
            logger.warning(f"Audio status: {status}")
        if self.is_running:
            self.audio_queue.put(indata[:, 0].copy())

    def _process_audio(self):
        logger.info("Audio processing thread started.")
        audio_buffer = []
        while self.is_running:
            try:
                chunk = self.audio_queue.get(timeout=1.0)
                audio_buffer.append(chunk)
                buffer_duration = sum(len(c) for c in audio_buffer) / self.config.sample_rate
                if buffer_duration < self.config.chunk_duration_seconds:
                    continue
                full_audio = np.concatenate(audio_buffer)
                audio_buffer = []
                segments, _ = self.model.transcribe(
                    full_audio, beam_size=1, temperature=0,
                    vad_filter=True, language="en",
                    initial_prompt=self.command_prompt
                )
                transcribed_text = " ".join(seg.text.strip() for seg in segments)
                if not transcribed_text:
                    continue
                logger.info(f"[Transcribed]: {transcribed_text}")
                self._handle_transcription(transcribed_text)
            except queue.Empty:
                continue
            except Exception as e:
                logger.error(f"Processing error: {e}")

    def _handle_transcription(self, text: str):
        text_lower = text.lower()
        self._parse_and_queue_command(text_lower, text)

    def _parse_and_queue_command(self, text_to_parse: str, original_text: str):
        command = self.parser.parse_command(text_to_parse)
        if command:
            logger.info(f"==> Command Parsed: {command['command']} | Params: {command['parameters']}")
            command['original_text'] = original_text
            self.command_queue.put(command)
        else:
            logger.warning(f"No valid command parsed from: '{text_to_parse}'")

    def start(self):
        logger.info("Starting J.A.R.V.I.S. STT Engine...")
        self._load_model()
        self.is_running = True
        processing_thread = threading.Thread(target=self._process_audio)
        processing_thread.daemon = True
        processing_thread.start()
        try:
            with sd.InputStream(
                samplerate=self.config.sample_rate, channels=1,
                dtype=np.float32, callback=self._audio_callback
            ):
                logger.info("Listening (continuous mode). Press Ctrl+C to stop")
                while self.is_running:
                    try:
                        cmd = self.command_queue.get_nowait()
                        logger.info(f"[COMMAND QUEUE] ==> Got command: {cmd}")
                    except queue.Empty:
                        pass
                    time.sleep(0.1)
        except KeyboardInterrupt:
            self.stop()

    def stop(self):
        if self.is_running:
            logger.info("Stopping J.A.R.V.I.S....")
            self.is_running = False

def main():
    jarvis = JARVIS_STT()
    jarvis.start()

if __name__ == "__main__":
    main()
