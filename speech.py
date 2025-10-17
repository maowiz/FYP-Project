import subprocess
import json
from typing import List, Dict, Optional

try:
    import pyttsx3
except ImportError:
    pyttsx3 = None

from assistant_state import set_speaking


class Speech:
    """Windows-native TTS using PowerShell (System.Speech) to avoid pyttsx3 issues."""

    def __init__(self):
        print("Initializing Windows PowerShell TTS...")
        self.can_speak_flag: bool = False
        self.voice_name: Optional[str] = None
        # System.Speech rate is -10..10; pick a clear default
        self.rate_steps: int = 2
        # System.Speech volume is 0..100
        self.volume_percent: int = 100
        self._current_ps_proc: Optional[subprocess.Popen] = None

        # --- Fallback Engine ---
        self.use_fallback = False
        self.fallback_engine = None
        if pyttsx3:
            try:
                # Explicitly initialize with the SAPI5 driver for Windows for better compatibility.
                self.fallback_engine = pyttsx3.init(driverName='sapi5')
                print("[TTS] pyttsx3 fallback engine initialized with SAPI5 driver.")

                # --- Diagnostics and Robustness ---
                voices = self.fallback_engine.getProperty('voices')
                if voices:
                    print(f"[TTS] Found {len(voices)} voices for pyttsx3.")
                    # Log the first few voices for debugging purposes.
                    for i, voice in enumerate(voices[:2]):
                        print(f"  - Voice {i}: {voice.name} ({voice.id})")
                    # Explicitly set the first available voice to avoid issues with a bad default.
                    self.fallback_engine.setProperty('voice', voices[0].id)
                    print(f"[TTS] Set pyttsx3 voice to: {voices[0].name}")
                else:
                    print("[TTS] WARNING: No voices found for pyttsx3. Fallback TTS may not work.")
            except Exception as e:
                print(f"[TTS] FATAL: Could not initialize pyttsx3 fallback engine: {e}")
                self.fallback_engine = None # Ensure it's None on failure
        else:
            print("[TTS] Warning: pyttsx3 not installed. No fallback TTS is available.")
        print("PowerShell TTS initialized.")

    def _build_ps_command(self) -> List[str]:
        voice = self.voice_name or ""
        escaped_voice = voice.replace("'", "''")
        script = f"""
        $rate={self.rate_steps};
        $vol={self.volume_percent};
        $voice='{escaped_voice}';
        Add-Type -AssemblyName System.Speech;
        $sp = New-Object System.Speech.Synthesis.SpeechSynthesizer;
        try {{ $sp.SetOutputToDefaultAudioDevice() }} catch {{ Write-Error 'No default audio device found.'; exit 1; }};
        $sp.Rate = $rate; $sp.Volume = $vol;
        if ($voice -ne '') {{ try {{ $sp.SelectVoice($voice) }} catch {{ }} }}
        $text = [Console]::In.ReadToEnd();
        $sp.Speak($text);
        """
        return [
            "powershell",
            "-NoProfile",
            "-Command",
            script,
        ]


    def speak(self, text: str) -> None:
        if not self.can_speak_flag:
            print(f"[TTS] Blocked: can_speak_flag is False. Text: '{text}'")
            return

        # --- Fallback Logic ---
        if self.use_fallback:
            if self.fallback_engine:
                try:
                    print(f"[TTS] Speaking via pyttsx3 (fallback): '{text}'")
                    set_speaking(True)
                    self.fallback_engine.say(text)
                    self.fallback_engine.runAndWait()
                    set_speaking(False)
                except Exception as e:
                    print(f"[TTS] Error in pyttsx3 fallback: {e}")
            return # End of fallback path

        if self._current_ps_proc and self._current_ps_proc.poll() is None:
            print("[TTS] Engine is already busy. Command ignored.")
            return
        try:
            print(f"[TTS] Speaking via PowerShell: '{text}'")
            cmd = self._build_ps_command()
            self._current_ps_proc = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            set_speaking(True)
            stdout, stderr = self._current_ps_proc.communicate(input=text)
            if self._current_ps_proc.returncode != 0:
                err_msg = stderr.strip()
                print(f"[TTS] PowerShell error (code {self._current_ps_proc.returncode}): {err_msg}")
                # If we get the specific audio device error, switch to the fallback
                if "AudioException" in err_msg or "0x20" in err_msg:
                    print("[TTS] FATAL: Audio device error detected. Switching to pyttsx3 fallback engine for future calls.")
                    self.use_fallback = True
                    # Retry the same speech command with the newly activated fallback engine.
                    print(f"[TTS] Retrying with fallback: '{text}'")
                    self.speak(text)
            self._current_ps_proc = None
            set_speaking(False)
        except Exception as e:
            print(f"[TTS] Error in PowerShell TTS: {e}")
            self._current_ps_proc = None
            set_speaking(False)

    def stop_speaking(self) -> None:
        self.can_speak_flag = False
        if self._current_ps_proc and self._current_ps_proc.poll() is None:
            try:
                self._current_ps_proc.terminate()
            except Exception:
                pass
        self._current_ps_proc = None
        set_speaking(False)

    def start_speaking(self) -> None:
        print("start_speaking() called. Setting can_speak_flag to True.")
        self.can_speak_flag = True

    def set_voice(self, voice_index: int = 0) -> bool:
        try:
            voices = self.get_available_voices()
            if voices and 0 <= voice_index < len(voices):
                self.voice_name = voices[voice_index]["name"]
                return True
            return False
        except Exception as e:
            print(f"Error setting voice: {e}")
            return False

    def set_rate(self, rate: int = 2) -> bool:
        try:
            # System.Speech expects -10..10
            self.rate_steps = max(-10, min(10, int(rate)))
            return True
        except Exception as e:
            print(f"Error setting rate: {e}")
            return False

    def set_volume(self, volume: float = 1.0) -> bool:
        try:
            self.volume_percent = max(0, min(100, int(volume * 100)))
            return True
        except Exception as e:
            print(f"Error setting volume: {e}")
            return False

    def get_available_voices(self) -> List[Dict[str, str]]:
        try:
            script = (
                "Add-Type -AssemblyName System.Speech; "
                "$sp = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                "($sp.GetInstalledVoices()).VoiceInfo | Select-Object Name, Culture, Gender, Description | ConvertTo-Json -Compress"
            )
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                check=False,
            )
            if proc.returncode != 0:
                return []
            data = json.loads(proc.stdout or "[]")
            if isinstance(data, dict):
                data = [data]
            result: List[Dict[str, str]] = []
            for idx, item in enumerate(data):
                result.append({
                    "index": idx,
                    "name": item.get("Name", ""),
                    "culture": item.get("Culture", ""),
                    "gender": str(item.get("Gender", "")),
                    "description": item.get("Description", ""),
                })
            return result
        except Exception:
            return []