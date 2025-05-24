import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel
import collections
import sys
import threading
import queue
import torch  # For Silero VAD
import socket  # For checking internet connectivity
from flask import Flask, request, jsonify, render_template_string
import subprocess
import time
import re

# --- Configuration ---
# Voice Recognition Mode
USE_ONLINE_RECOGNITION = True   # Toggle this to True for online (Flask) mode, False for offline (Faster Whisper)

# Whisper Model Settings (Offline)
MODEL_SIZE = "medium"  # Options: "tiny", "base", "small", "medium", "large-v1", "large-v2", "large-v3"
DEVICE = "cpu"  # "cpu" or "cuda"
COMPUTE_TYPE = "int8"  # For CPU: "int8"; for CUDA: "float16" or "int8_float16"

# Audio Stream Settings (Offline)
SAMPLE_RATE = 16000
STREAM_BLOCK_SIZE_MS = 32
SILENCE_THRESHOLD_S = 1.0
SPEECH_PAD_S = 0.25
MAX_UTTERANCE_S = 20
VAD_THRESHOLD = 0.4

# Flask Settings (Online)
FLASK_HOST = "127.0.0.1"
FLASK_PORT = 5000

# HTML for Flask-based Web Speech API
HTML_CODE = """
<!DOCTYPE html>
<html>
<head>
  <title>Voice Listener</title>
</head>
<body>
<h2>Listening to your voice...</h2>
<script>
  const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
  recognition.continuous = true;
  recognition.lang = 'en-US';

  recognition.onresult = function(event) {
    const text = event.results[event.resultIndex][0].transcript.trim();
    fetch("/speech", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text })
    });
  };

  recognition.start();
</script>
</body>
</html>
"""

class VoiceRecognizer:
    def __init__(self):
        self.is_speaking = False
        self.silence_frames_count = 0
        self.current_utterance_audio = collections.deque()
        self.padding_buffer = collections.deque(maxlen=int(SPEECH_PAD_S * SAMPLE_RATE))
        self.last_transcription = ""
        self.new_transcription_available = False
        self.audio_stream = None
        self.flask_app = None
        self.flask_thread = None
        self.transcription_queue = queue.Queue()

        # Check internet connectivity
        self.is_online = self._check_internet()

        # Decide which recognition mode to use based on toggle and connectivity
        self.use_online = USE_ONLINE_RECOGNITION and self.is_online
        if self.use_online:
            print("Initializing online recognition (Web Speech API via Flask)...")
            self._init_flask()
        else:
            print("Initializing offline recognition (Faster Whisper)...")
            self._init_whisper()

        # Start transcription worker for both modes
        self.transcription_thread = threading.Thread(target=self._transcription_worker, daemon=True)
        self.transcription_thread.start()

    def _check_internet(self):
        """Check if internet is available by attempting to connect to a reliable host."""
        try:
            socket.create_connection(("www.google.com", 80), timeout=2)
            return True
        except OSError:
            print("No internet connection detected. Falling back to offline mode.")
            return False

    def _init_flask(self):
        """Initialize Flask server for online recognition."""
        self.flask_app = Flask(__name__)

        @self.flask_app.route("/")
        def index():
            return render_template_string(HTML_CODE)

        @self.flask_app.route("/speech", methods=["POST"])
        def speech():
            data = request.get_json()
            text = data.get("text", "").strip().lower()
            if text:
                # Clean punctuation similar to offline mode
                text = re.sub(r'[!,?;:"\[\]{}()\-_=+*/\\|<>]', '', text)
                self.transcription_queue.put(text)
                print(f"You said: {text}")
            return jsonify({"status": "ok"})

        # Start Flask in a separate thread
        self.flask_thread = threading.Thread(target=self._run_flask, daemon=True)
        self.flask_thread.start()
        time.sleep(2)  # Allow server to start

        # Open Chrome minimized and off-screen
        try:
            subprocess.Popen(
                f'start /min chrome --window-position=-32000,-32000 --window-size=1,1 http://{FLASK_HOST}:{FLASK_PORT}',
                shell=True
            )
            print("Started Chrome for Web Speech API.")
        except Exception as e:
            print(f"Error starting Chrome for Web Speech API: {e}")
            self.use_online = False
            self._init_whisper()  # Fallback to offline mode

    def _run_flask(self):
        """Run the Flask server."""
        self.flask_app.run(host=FLASK_HOST, port=FLASK_PORT, debug=False, use_reloader=False)

    def _init_whisper(self):
        """Initialize Faster Whisper and Silero VAD for offline recognition."""
        self.stream_block_size_frames = int(STREAM_BLOCK_SIZE_MS / 1000 * SAMPLE_RATE)
        self.silence_duration_frames_limit = int(SILENCE_THRESHOLD_S * SAMPLE_RATE)
        self.max_utterance_frames_limit = int(MAX_UTTERANCE_S * SAMPLE_RATE)

        # Initialize Faster-Whisper
        print(f"Loading Faster-Whisper model: {MODEL_SIZE} (Device: {DEVICE}, Compute: {COMPUTE_TYPE})...")
        try:
            self.whisper_model = WhisperModel(MODEL_SIZE, device=DEVICE, compute_type=COMPUTE_TYPE)
            print("Whisper model loaded successfully.")
        except Exception as e:
            print(f"Error loading Faster-Whisper model: {e}")
            sys.exit(1)

        # Initialize Silero VAD
        print("Loading Silero VAD model from local directory...")
        try:
            local_vad_path = r"C:\Users\LENOVO\Desktop\FYP\final\silero-vad-master\silero-vad-master"
            self.vad_model, self.vad_utils = torch.hub.load(
                repo_or_dir=local_vad_path,
                model='silero_vad',
                source='local',
                force_reload=False,
                onnx=True
            )
            (self.vad_get_speech_timestamps, _, _, _, _) = self.vad_utils
            print("Silero VAD model loaded successfully.")
        except Exception as e:
            print(f"Error loading Silero VAD model: {e}")
            sys.exit(1)

    def _transcription_worker(self):
        """Process transcriptions from the queue (handles both audio and text inputs)."""
        while True:
            try:
                data = self.transcription_queue.get(timeout=1)
                if isinstance(data, str):
                    # Online mode: text directly from Web Speech API
                    self.last_transcription = data
                    self.new_transcription_available = True
                elif isinstance(data, np.ndarray):
                    # Offline mode: audio data to transcribe
                    transcribed_text = self._execute_transcription(data)
                    if transcribed_text:
                        self.last_transcription = transcribed_text
                        self.new_transcription_available = True
                self.transcription_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                print(f"Transcription worker error: {e}")

    def _execute_transcription(self, audio_data_np):
        """Transcribe audio data using Faster Whisper (offline mode)."""
        if audio_data_np is None or audio_data_np.size < int(0.1 * SAMPLE_RATE):
            return ""
        try:
            segments, _ = self.whisper_model.transcribe(audio_data_np, beam_size=5, language="en")
            full_text = "".join(segment.text for segment in segments).strip().lower()
            full_text = re.sub(r'[!,?;:"\[\]{}()\-_=+*/\\|<>]', '', full_text)
            return full_text
        except Exception as e:
            print(f"Transcription error: {e}")
            return ""

    def _trigger_transcription(self):
        """Prepare and queue audio for transcription (offline mode)."""
        if not self.current_utterance_audio:
            self.is_speaking = False
            self.silence_frames_count = 0
            self.current_utterance_audio.clear()
            return

        audio_to_process = np.array(list(self.current_utterance_audio)).astype(np.float32)
        self.current_utterance_audio.clear()
        self.is_speaking = False
        self.silence_frames_count = 0

        if len(audio_to_process) < int(0.2 * SAMPLE_RATE):
            return
        if len(audio_to_process) > self.max_utterance_frames_limit:
            print(f"Warning: Utterance exceeded max length ({MAX_UTTERANCE_S}s), truncating.")
            audio_to_process = audio_to_process[:self.max_utterance_frames_limit]

        self.transcription_queue.put(audio_to_process)

    def audio_processing_callback(self, indata, frames, time_info, status):
        """Process audio chunks with Silero VAD (offline mode)."""
        if status:
            print(f"Audio callback status: {status}")

        current_chunk_np = np.frombuffer(indata, dtype=np.float32).copy()
        self.padding_buffer.extend(current_chunk_np)

        try:
            audio_chunk_tensor = torch.from_numpy(current_chunk_np)
            speech_prob = self.vad_model(audio_chunk_tensor, SAMPLE_RATE).item()
        except Exception as e:
            print(f"VAD processing error: {e}")
            speech_prob = 0.0

        is_current_chunk_speech = speech_prob > VAD_THRESHOLD

        if is_current_chunk_speech:
            if not self.is_speaking:
                self.is_speaking = True
                self.current_utterance_audio.clear()
                self.current_utterance_audio.extend(list(self.padding_buffer))
            self.current_utterance_audio.extend(current_chunk_np)
            self.silence_frames_count = 0
            if len(self.current_utterance_audio) > self.max_utterance_frames_limit:
                self._trigger_transcription()
        else:
            if self.is_speaking:
                self.current_utterance_audio.extend(current_chunk_np)
                self.silence_frames_count += frames
                if self.silence_frames_count >= self.silence_duration_frames_limit:
                    self._trigger_transcription()

    def start_listening(self):
        """Start the audio input stream (offline) or ensure Flask is running (online)."""
        if self.use_online:
            if self.flask_app and self.flask_thread and self.flask_thread.is_alive():
                print("Flask server is already running for online recognition.")
            else:
                print("Restarting Flask server...")
                self._init_flask()
        else:
            if self.audio_stream and self.audio_stream.active:
                print("Audio stream is already active.")
                return
            print(f"Starting audio stream with blocksize {self.stream_block_size_frames} frames ({STREAM_BLOCK_SIZE_MS}ms)...")
            try:
                self.audio_stream = sd.InputStream(
                    samplerate=SAMPLE_RATE,
                    blocksize=self.stream_block_size_frames,
                    channels=1,
                    dtype='float32',
                    callback=self.audio_processing_callback,
                    latency="low"
                )
                self.audio_stream.start()
                print("Listening... Press Ctrl+C to stop.")
            except Exception as e:
                print(f"Error starting audio stream: {e}")
                self.audio_stream = None
                sys.exit(1)

    def stop_listening(self):
        """Stop the audio stream (offline) or Flask server (online)."""
        if self.use_online:
            print("Online mode: Cannot stop Flask server programmatically in this implementation.")
            # Note: Stopping Flask requires external intervention or process termination
        else:
            if self.audio_stream:
                if self.audio_stream.active:
                    print("Stopping audio stream...")
                    self.audio_stream.stop()
                    self.audio_stream.close()
                    print("Audio stream stopped.")
                self.audio_stream = None
            self.current_utterance_audio.clear()
            self.padding_buffer.clear()
            self.is_speaking = False
            self.silence_frames_count = 0

    def get_transcription(self):
        """Retrieve the latest transcription (works for both modes)."""
        if self.new_transcription_available:
            self.new_transcription_available = False
            return self.last_transcription
        return None

    @staticmethod
    def get_settings():
        """Return the voice recognition settings."""
        settings = {
            "Mode": "Online (Web Speech API)" if USE_ONLINE_RECOGNITION else "Offline (Faster Whisper)",
            "Flask_Host": FLASK_HOST,
            "Flask_Port": FLASK_PORT
        }
        if not USE_ONLINE_RECOGNITION:
            settings.update({
                "Model_Size": MODEL_SIZE,
                "Device": DEVICE,
                "Compute_Type": COMPUTE_TYPE,
                "Sample_Rate_Hz": SAMPLE_RATE,
                "Stream_Block_Size_ms": STREAM_BLOCK_SIZE_MS,
                "Silence_Threshold_s": SILENCE_THRESHOLD_S,
                "Speech_Pad_s": SPEECH_PAD_S,
                "Max_Utterance_s": MAX_UTTERANCE_S,
                "VAD_Confidence_Threshold": VAD_THRESHOLD
            })
        return settings

if __name__ == '__main__':
    recognizer = VoiceRecognizer()
    recognizer.start_listening()

    try:
        while True:
            transcription = recognizer.get_transcription()
            if transcription:
                pass  # Transcription handled by CommandHandler in main application
            sd.sleep(100) if not USE_ONLINE_RECOGNITION else time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nStopping voice recognition...")
    finally:
        recognizer.stop_listening()
        print("Program terminated.")