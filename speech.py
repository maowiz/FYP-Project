import pyttsx3

class Speech:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 200)  # Default rate for clearer speech
        self.is_speaking = False

    def speak(self, text):
        """Convert text to speech with interruption control."""
        if not self.is_speaking:
            return
        try:
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            print(f"Error in speaking: {e}")

    def stop_speaking(self):
        """Stop the speech engine."""
        self.is_speaking = False
        self.engine.stop()

    def start_speaking(self):
        """Enable speaking."""
        self.is_speaking = True