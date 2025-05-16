import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel
import collections
import sys

# --- Configuration for Voice Recognition ---
MODEL_SIZE = "tiny"  # "tiny", "base" are fastest. Use "base" if "tiny" is not accurate enough.
DEVICE = "cpu"       # "cpu" or "cuda" (if you have a compatible NVIDIA GPU)
COMPUTE_TYPE = "float32"  # Use "float32" for CPU compatibility

SAMPLE_RATE = 16000  # Whisper models are trained on 16kHz audio
AUDIO_CHUNK_DURATION_S = 0.5  # Process audio in chunks of this duration (seconds)
SILENCE_THRESHOLD_S = 1.0     # Seconds of silence to consider an utterance ended
VAD_ENERGY_THRESHOLD = 0.005  # Energy threshold for VAD. Adjust based on your microphone and noise.
MAX_AUDIO_BUFFER_S = 10       # Max duration of audio to keep in buffer (seconds)

class VoiceRecognizer:
    def __init__(self):
        # Global state variables for voice recognition
        self.g_is_speaking = False
        self.g_silence_frames_count = 0
        self.g_audio_buffer = collections.deque(maxlen=int(MAX_AUDIO_BUFFER_S * SAMPLE_RATE))
        self.g_last_transcription = ""
        self.g_new_transcription_available = False

        # Initialize Faster-Whisper model
        print(f"Loading Faster-Whisper model: {MODEL_SIZE} (Device: {DEVICE}, Compute: {COMPUTE_TYPE})...")
        try:
            self.model = WhisperModel(MODEL_SIZE, device=DEVICE, compute_type=COMPUTE_TYPE)
            print("Model loaded successfully.")
        except Exception as e:
            print(f"Error loading model: {e}")
            print("Please ensure model files are accessible and device/compute_type are valid.")
            sys.exit(1)

    def transcribe_audio_data(self, audio_data_np):
        """Transcribes a NumPy array of audio data."""
        if audio_data_np is None or audio_data_np.size < int(0.1 * SAMPLE_RATE):  # Min audio length check
            return ""
        try:
            segments, _ = self.model.transcribe(audio_data_np, beam_size=1, language="en")
            full_text = "".join(segment.text for segment in segments).strip().lower()
            return full_text
        except Exception as e:
            print(f"Transcription error: {e}")
            return ""

    def audio_processing_callback(self, indata, frames, time_info, status):
        """Callback function for sounddevice to process audio chunks."""
        if status:
            print(f"Audio callback status: {status}", flush=True)

        current_chunk = np.frombuffer(indata, dtype=np.float32).copy()
        
        # Simple energy-based VAD
        energy = np.sum(current_chunk**2) / len(current_chunk)

        if self.g_is_speaking:
            self.g_audio_buffer.extend(current_chunk)
            
            if energy < VAD_ENERGY_THRESHOLD:
                self.g_silence_frames_count += frames
                if self.g_silence_frames_count >= int(SILENCE_THRESHOLD_S * SAMPLE_RATE):
                    audio_to_transcribe = np.array(self.g_audio_buffer).astype(np.float32)
                    self.g_audio_buffer.clear()
                    self.g_is_speaking = False
                    self.g_silence_frames_count = 0
                    
                    transcribed_text = self.transcribe_audio_data(audio_to_transcribe)
                    if transcribed_text:
                        print(f"You said: {transcribed_text}", flush=True)
                        self.g_last_transcription = transcribed_text
                        self.g_new_transcription_available = True  # Signal new transcription
            else:
                self.g_silence_frames_count = 0
        else:
            if energy > VAD_ENERGY_THRESHOLD:
                print("\nSpeech detected!", flush=True)
                self.g_is_speaking = True
                self.g_silence_frames_count = 0
                self.g_audio_buffer.clear()
                self.g_audio_buffer.extend(current_chunk)

    def start_listening(self):
        """Start the audio input stream and return the block size."""
        block_size = int(AUDIO_CHUNK_DURATION_S * SAMPLE_RATE)
        stream = sd.InputStream(
            samplerate=SAMPLE_RATE,
            blocksize=block_size,
            channels=1,
            dtype='float32',
            callback=self.audio_processing_callback,
            latency="high"
        )
        stream.start()
        return stream

    def get_transcription(self):
        """Retrieve the latest transcription if available."""
        if self.g_new_transcription_available:
            self.g_new_transcription_available = False
            return self.g_last_transcription
        return None

    @staticmethod
    def get_settings():
        """Return the voice recognition settings for display."""
        return {
            "Chunk_duration": AUDIO_CHUNK_DURATION_S,
            "Silence_threshold": SILENCE_THRESHOLD_S,
            "VAD_energy": VAD_ENERGY_THRESHOLD
        }