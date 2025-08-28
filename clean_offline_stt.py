# clean_offline_stt.py

import queue
import threading
import numpy as np
import sounddevice as sd
from faster_whisper import WhisperModel
from logger import info, warn

class CleanOfflineSTT:
    """
    A clean, continuous listening offline Speech-to-Text engine
    based on the proven components of whisper_test.py.
    """
    def __init__(self, command_list):
        # --- Core Settings from whisper_test.py ---
        self.model_name = "base.en"
        self.sample_rate = 16000
        self.chunk_duration_seconds = 3.0
        self.silence_threshold = 0.02
        self.cpu_threads = 4
        self.compute_type = "int8"
        self.command_list = command_list
        self.command_prompt = self._generate_command_prompt()

        # --- Queues and State Management ---
        self.audio_queue = queue.Queue()
        self.transcription_queue = queue.Queue()
        self.is_listening = threading.Event()
        self._stop_event = threading.Event()
        self._processing_thread = None
        self._stream = None

        # --- Load Whisper Model ---
        info("Loading offline Whisper model...")
        try:
            self.model = WhisperModel(
                self.model_name,
                device="cpu",
                compute_type=self.compute_type,
                cpu_threads=self.cpu_threads
            )
            info("Offline Whisper model loaded successfully.")
        except Exception as e:
            warn(f"Could not load Whisper model: {e}")
            raise

    def _generate_command_prompt(self):
        """Generates a command prompt to improve transcription accuracy."""
        return ", ".join(self.command_list) + "."

    def _audio_callback(self, indata, frames, time, status):
        """This is called by sounddevice for each audio block."""
        if status:
            warn(f"Sounddevice status: {status}")
        if self.is_listening.is_set():
            self.audio_queue.put(indata.copy())

    def _process_audio(self):
        """Continuously processes audio from the queue."""
        accumulated_audio = np.array([], dtype=np.float32)

        while not self._stop_event.is_set():
            try:
                # Get all available audio from the queue
                while not self.audio_queue.empty():
                    audio_chunk = self.audio_queue.get_nowait()
                    accumulated_audio = np.concatenate((accumulated_audio, audio_chunk))

                # Process only if we have enough audio data
                required_samples = int(self.chunk_duration_seconds * self.sample_rate)
                if len(accumulated_audio) < required_samples:
                    time.sleep(0.1)  # Wait for more audio
                    continue
                
                # Use the most recent audio chunk
                processing_audio = accumulated_audio[-required_samples:]
                # Keep a small overlap for the next round
                overlap_samples = int(0.5 * self.sample_rate)
                accumulated_audio = accumulated_audio[-overlap_samples:]

                # --- VAD Filtering (from whisper_test.py) ---
                if np.max(np.abs(processing_audio)) < self.silence_threshold:
                    continue

                # --- Transcription (from whisper_test.py) ---
                segments, _ = self.model.transcribe(
                    processing_audio,
                    beam_size=1,
                    temperature=0,
                    vad_filter=True,
                    language="en",
                    initial_prompt=self.command_prompt
                )

                transcription = "".join(segment.text for segment in segments).strip()

                if transcription:
                    info(f"Offline STT transcribed: '{transcription}'")
                    self.transcription_queue.put(transcription)

            except queue.Empty:
                continue
            except Exception as e:
                warn(f"Error in audio processing thread: {e}")


    def start(self):
        """Starts the audio stream and processing thread."""
        if self._stream is None:
            self._stream = sd.InputStream(
                samplerate=self.sample_rate,
                channels=1,
                dtype='float32',
                callback=self._audio_callback
            )
            self._stream.start()
            info("Sounddevice stream started.")

        if self._processing_thread is None:
            self._stop_event.clear()
            self._processing_thread = threading.Thread(target=self._process_audio)
            self._processing_thread.daemon = True
            self._processing_thread.start()
            info("Audio processing thread started.")

        self.resume()

    def stop(self):
        """Stops the audio stream and processing thread."""
        self.pause()
        self._stop_event.set()

        if self._stream:
            self._stream.stop()
            self._stream.close()
            self._stream = None
            info("Sounddevice stream stopped.")

        if self._processing_thread:
            self._processing_thread.join(timeout=2)
            self._processing_thread = None
            info("Audio processing thread stopped.")

    def pause(self):
        """Pauses listening."""
        self.is_listening.clear()

    def resume(self):
        """Resumes listening."""
        self.is_listening.set()

    def get_transcription(self):
        """Gets a transcription from the queue if available."""
        try:
            return self.transcription_queue.get_nowait()
        except queue.Empty:
            return None