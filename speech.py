import subprocess
import json
from typing import List, Dict, Optional
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
        print("PowerShell TTS initialized.")

    def _build_ps_command(self) -> List[str]:
        voice = self.voice_name or ""
        script = (
            f"$rate={self.rate_steps}; $vol={self.volume_percent}; $voice='{voice.replace("'", "''")}';\n"
            "Add-Type -AssemblyName System.Speech;\n"
            "$sp = New-Object System.Speech.Synthesis.SpeechSynthesizer;\n"
            "$sp.Rate = $rate; $sp.Volume = $vol;\n"
            "if ($voice -ne '') { try { $sp.SelectVoice($voice) } catch { } }\n"
            "$text = [Console]::In.ReadToEnd();\n"
            "$sp.Speak($text);\n"
        )
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
                print(f"[TTS] PowerShell error (code {self._current_ps_proc.returncode}): {stderr.strip()}")
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