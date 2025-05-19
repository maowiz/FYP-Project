import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel
import collections
import sys
import threading
import queue
import torch # For Silero VAD

# --- Configuration ---
# Whisper Model Settings
MODEL_SIZE = "small"  # Options: "tiny", "base", "small", "medium", "large-v1", "large-v2", "large-v3".
                     # "small" or "medium" offer a good balance of speed and accuracy if "base" isn't enough.
                     # Larger models are more accurate but slower.
DEVICE = "cpu"       # "cpu" or "cuda" (if you have an NVIDIA GPU and CUDA toolkit installed).
                     # Using "cuda" significantly speeds up transcription.
COMPUTE_TYPE = "float32" # Default for CPU.
                         # For CPU, "int8" can offer more speed but may reduce accuracy (requires ctranslate2>=3.16.0).
                         # For "cuda" DEVICE:
                         #   "float16" is recommended for a good speedup with minimal accuracy loss.
                         #   "int8_float16" or "int8" can be faster but with a higher chance of accuracy reduction.

# Audio Stream Settings
SAMPLE_RATE = 16000  # Whisper models are trained on 16kHz audio. Silero VAD also works well with this.
STREAM_BLOCK_SIZE_MS = 32  # Block size for sounddevice stream callback in milliseconds.
                           # Silero VAD works with chunks like 256, 512, 768, 1024, 1536 samples.
                           # 32ms * 16kHz = 512 samples. 48ms * 16kHz = 768 samples.
                           # This value (e.g. 32, 48, 64) will be converted to frames.
SILENCE_THRESHOLD_S = 1.0   # Seconds of VAD-detected silence to consider an utterance ended.
SPEECH_PAD_S = 0.25         # Seconds of padding to add to the beginning of detected speech.
MAX_UTTERANCE_S = 20        # Max duration of a single utterance to transcribe (seconds).
                            # Prevents excessively long audio segments.

# VAD Settings (Silero)
VAD_THRESHOLD = 0.4 # Confidence threshold for Silero VAD (0.0 to 1.0). Higher means more confident speech detection.
                    # Adjust this based on your microphone and environment.

class VoiceRecognizer:
    def __init__(self):
        self.is_speaking = False
        self.silence_frames_count = 0
        self.current_utterance_audio = collections.deque()
        # Buffer for recent audio, used to provide padding before speech starts
        self.padding_buffer = collections.deque(maxlen=int(SPEECH_PAD_S * SAMPLE_RATE))

        self.last_transcription = ""
        self.new_transcription_available = False

        self.stream_block_size_frames = int(STREAM_BLOCK_SIZE_MS / 1000 * SAMPLE_RATE)
        self.silence_duration_frames_limit = int(SILENCE_THRESHOLD_S * SAMPLE_RATE)
        self.max_utterance_frames_limit = int(MAX_UTTERANCE_S * SAMPLE_RATE)

        # Initialize Faster-Whisper model
        print(f"Loading Faster-Whisper model: {MODEL_SIZE} (Device: {DEVICE}, Compute: {COMPUTE_TYPE})...")
        try:
            self.whisper_model = WhisperModel(MODEL_SIZE, device=DEVICE, compute_type=COMPUTE_TYPE)
            print("Whisper model loaded successfully.")
        except Exception as e:
            print(f"Error loading Faster-Whisper model: {e}")
            print("Please ensure model files are accessible and device/compute_type are valid.")
            print("If using GPU, ensure CUDA and cuDNN are correctly installed and compatible.")
            print("If using int8 on CPU, ensure ctranslate2 version is >= 3.16.0.")
            sys.exit(1)

        # Initialize Silero VAD
        print("Loading Silero VAD model from local directory...")
        try:
            # Use the local path to the silero-vad directory
            local_vad_path = r"C:\Users\LENOVO\Desktop\FYP\final\silero-vad-master\silero-vad-master"
            self.vad_model, self.vad_utils = torch.hub.load(repo_or_dir=local_vad_path,
                                                             model='silero_vad',
                                                             source='local',  # Specify that we're using a local directory
                                                             force_reload=False,  # Set to False to avoid unnecessary reloads
                                                             onnx=True)  # Use ONNX for potentially faster CPU inference
            (self.vad_get_speech_timestamps, _, _, _, _) = self.vad_utils  # Unpack utils
            print("Silero VAD model loaded successfully from local directory.")
        except Exception as e:
            print(f"Error loading Silero VAD model: {e}")
            print("Ensure the local path contains the necessary model files (e.g., silero_vad.onnx or silero_vad.jit).")
            print("If issues persist, try setting onnx=False or verify the directory structure.")
            sys.exit(1)

        # Transcription queue and worker thread
        self.transcription_queue = queue.Queue()
        self.transcription_thread = threading.Thread(target=self._transcription_worker, daemon=True)
        self.transcription_thread.start()
        self.audio_stream = None

    def _transcription_worker(self):
        """Continuously processes audio data from the queue for transcription."""
        while True:
            try:
                audio_data_np = self.transcription_queue.get(timeout=1) # Wait for 1 sec then check again
                if audio_data_np is None: # Signal to exit (optional)
                    continue

                # print(f"Worker: Transcribing {len(audio_data_np)/SAMPLE_RATE:.2f}s of audio...")
                transcribed_text = self._execute_transcription(audio_data_np)
                if transcribed_text:
                    print(f"You said: {transcribed_text}", flush=True)
                    self.last_transcription = transcribed_text
                    self.new_transcription_available = True
                self.transcription_queue.task_done()
            except queue.Empty:
                continue # No data, continue waiting
            except Exception as e:
                print(f"Transcription worker error: {e}", flush=True)

    def _execute_transcription(self, audio_data_np):
        """Transcribes a NumPy array of audio data using Whisper."""
        if audio_data_np is None or audio_data_np.size < int(0.1 * SAMPLE_RATE):  # Min audio length
            return ""
        try:
            # Note: beam_size can be adjusted. Higher values might increase accuracy but are slower.
            # language="en" is set to improve accuracy and speed if only English is expected.
            segments, _ = self.whisper_model.transcribe(audio_data_np, beam_size=5, language="en")
            full_text = "".join(segment.text for segment in segments).strip().lower()

            # Original punctuation cleaning. Consider if this is too aggressive for your needs.
            # For some applications (e.g., feeding to an LLM), more natural punctuation is better.
            import re
            # This regex removes most punctuation except periods.
            # You might want to preserve apostrophes (e.g., "it's") or hyphens.
            full_text = re.sub(r'[!,?;:"\[\]{}()\-_=+*/\\|<>]', '', full_text)
            # Example: keep apostrophes and hyphens within words:
            # full_text = re.sub(r"[^\w\s.'-]|(?<!\w)[.-](?!\w)|(?<!\w)[.-]|[.-](?!\w)", "", full_text)


            return full_text
        except Exception as e:
            print(f"Transcription error: {e}", flush=True)
            return ""

    def _trigger_transcription(self):
        """Prepares the current utterance audio and puts it on the transcription queue."""
        if not self.current_utterance_audio:
            self.is_speaking = False
            self.silence_frames_count = 0
            self.current_utterance_audio.clear() # Should be already clear
            return

        audio_to_process = np.array(list(self.current_utterance_audio)).astype(np.float32)
        self.current_utterance_audio.clear() # Clear for the next utterance
        self.is_speaking = False # Reset speaking flag
        self.silence_frames_count = 0

        # Final checks on the audio segment before queuing
        if len(audio_to_process) < int(0.2 * SAMPLE_RATE): # Min length (e.g., 200ms)
            # print("Utterance too short, discarding.", flush=True)
            return
        if len(audio_to_process) > self.max_utterance_frames_limit:
            print(f"Warning: Utterance exceeded max length ({MAX_UTTERANCE_S}s), truncating.", flush=True)
            audio_to_process = audio_to_process[:self.max_utterance_frames_limit]

        self.transcription_queue.put(audio_to_process)


    def audio_processing_callback(self, indata, frames, time_info, status):
        """Callback function for sounddevice to process audio chunks with Silero VAD."""
        if status:
            print(f"Audio callback status: {status}", flush=True)

        if frames != self.stream_block_size_frames:
            # This might happen if the host can't keep up or blocksize is not a power of 2.
            # print(f"Warning: Expected {self.stream_block_size_frames} frames, got {frames}", flush=True)
            # For simplicity, we'll process what we got, but VAD might be less optimal.
            pass


        current_chunk_np = np.frombuffer(indata, dtype=np.float32).copy()

        # Update padding buffer continuously
        self.padding_buffer.extend(current_chunk_np)

        # Perform VAD on the current chunk
        # Silero VAD expects a torch tensor.
        try:
            audio_chunk_tensor = torch.from_numpy(current_chunk_np)
            # The VAD model gives speech probability for the current chunk.
            speech_prob = self.vad_model(audio_chunk_tensor, SAMPLE_RATE).item()
        except Exception as e:
            print(f"VAD processing error: {e}", flush=True)
            speech_prob = 0.0 # Assume no speech on error

        is_current_chunk_speech = speech_prob > VAD_THRESHOLD

        if is_current_chunk_speech:
            if not self.is_speaking:
                # print("\nSpeech detected!", flush=True)
                self.is_speaking = True
                self.current_utterance_audio.clear() # Start a new utterance
                # Prepend the audio from the padding_buffer to capture the start of speech
                self.current_utterance_audio.extend(list(self.padding_buffer))

            self.current_utterance_audio.extend(current_chunk_np)
            self.silence_frames_count = 0 # Reset silence counter upon detecting speech

            # Check if the current utterance exceeds the maximum allowed length
            if len(self.current_utterance_audio) > self.max_utterance_frames_limit:
                # print("Max utterance length reached during speech, processing current audio.", flush=True)
                self._trigger_transcription() # This also resets is_speaking
                # The current chunk (which is speech) might start a *new* utterance if _trigger_transcription
                # clears current_utterance_audio and this chunk is then added again.
                # The _trigger_transcription clears current_utterance_audio.
                # The current_chunk_np that *caused* the overflow was already added.
                # So, the next chunk will be the start of potentially new audio.
        else: # No speech detected in the current chunk
            if self.is_speaking:
                # Speech was ongoing, but this chunk is silence. It might be a pause or end of utterance.
                self.current_utterance_audio.extend(current_chunk_np) # Add the silence chunk
                self.silence_frames_count += frames
                if self.silence_frames_count >= self.silence_duration_frames_limit:
                    # print("Silence threshold reached, processing utterance.", flush=True)
                    self._trigger_transcription()
            # else: Not speaking and current chunk is silence. padding_buffer is updated. Nothing else to do.

    def start_listening(self):
        """Start the audio input stream."""
        if self.audio_stream and self.audio_stream.active:
            print("Audio stream is already active.")
            return

        print(f"Starting audio stream with blocksize {self.stream_block_size_frames} frames ({STREAM_BLOCK_SIZE_MS}ms)...")
        try:
            self.audio_stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                blocksize=self.stream_block_size_frames, # Use the calculated frame size
                channels=1,
                dtype='float32',
                callback=self.audio_processing_callback,
                latency="low"
            )
            self.audio_stream.start()
            print("Listening... Press Ctrl+C to stop.")
        except Exception as e:
            print(f"Error starting audio stream: {e}")
            self.audio_stream = None # Ensure stream is None if it failed to start
            sys.exit(1)

    def stop_listening(self):
        """Stop the audio input stream."""
        if self.audio_stream:
            if self.audio_stream.active:
                print("Stopping audio stream...")
                self.audio_stream.stop()
                self.audio_stream.close()
                print("Audio stream stopped.")
            else:
                print("Audio stream was not active.")
            self.audio_stream = None
        else:
            print("No active audio stream to stop.")

        # Optionally, clear buffers and reset state when stopping explicitly
        self.current_utterance_audio.clear()
        self.padding_buffer.clear()
        self.is_speaking = False
        self.silence_frames_count = 0
        # You might want to signal the transcription worker to finish any pending tasks
        # or add a sentinel value to its queue if you plan to join the thread.
        # Since it's a daemon thread, it will exit when the main program exits.

    def get_transcription(self):
        """Retrieve the latest transcription if available."""
        if self.new_transcription_available:
            self.new_transcription_available = False
            return self.last_transcription
        return None

    @staticmethod
    def get_settings():
        """Return the voice recognition settings for display."""
        return {
            "Model_Size": MODEL_SIZE,
            "Device": DEVICE,
            "Compute_Type": COMPUTE_TYPE,
            "Sample_Rate_Hz": SAMPLE_RATE,
            "Stream_Block_Size_ms": STREAM_BLOCK_SIZE_MS,
            "Silence_Threshold_s": SILENCE_THRESHOLD_S,
            "Speech_Pad_s": SPEECH_PAD_S,
            "Max_Utterance_s": MAX_UTTERANCE_S,
            "VAD_Confidence_Threshold": VAD_THRESHOLD
        }

if __name__ == '__main__':
    recognizer = VoiceRecognizer()
    recognizer.start_listening()

    try:
        while True:
            # Example of how to use get_transcription in a non-blocking way
            transcription = recognizer.get_transcription()
            if transcription:
                # Do something with the transcription in your main application logic
                # print(f"Main Loop - New Transcription: {transcription}")
                pass
            sd.sleep(100) # Check for new transcription every 100ms
    except KeyboardInterrupt:
        print("\nStopping voice recognition...")
    finally:
        recognizer.stop_listening()
        print("Program terminated.")