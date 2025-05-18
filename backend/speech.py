import pyttsx3
import time

class Speech:
    def __init__(self):
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 225)  # Slightly increased rate for faster speech
        self.is_speaking = False
        
        # Get available voices
        self.voices = self.engine.getProperty('voices')
        # Default to the first voice (usually male voice)
        self.engine.setProperty('voice', self.voices[0].id)

    def speak(self, text):
        """Convert text to speech with interruption control."""
        if not self.is_speaking:
            return
        try:
            # Break long text into shorter sentences to allow for easier interruption
            sentences = text.split('.')
            for sentence in sentences:
                if sentence.strip():
                    if not self.is_speaking:
                        return  # Check if speaking was stopped
                    self.engine.say(sentence.strip() + '.')
                    self.engine.runAndWait()
                    # Small pause between sentences to allow for interruption
                    time.sleep(0.1)
        except Exception as e:
            print(f"Error in speaking: {e}")

    def stop_speaking(self):
        """Stop the speech engine."""
        self.is_speaking = False
        self.engine.stop()

    def start_speaking(self):
        """Enable speaking."""
        self.is_speaking = True
        
    def set_voice(self, voice_index=0):
        """Change the voice of the speech engine.
        
        Args:
            voice_index (int): Index of the voice to use (0 for male, 1 for female typically)
        """
        try:
            if 0 <= voice_index < len(self.voices):
                self.engine.setProperty('voice', self.voices[voice_index].id)
                return True
            return False
        except Exception as e:
            print(f"Error setting voice: {e}")
            return False
    
    def set_rate(self, rate=225):
        """Set the speech rate.
        
        Args:
            rate (int): Speech rate (higher is faster, default is 225)
        """
        try:
            self.engine.setProperty('rate', rate)
            return True
        except Exception as e:
            print(f"Error setting rate: {e}")
            return False
            
    def get_available_voices(self):
        """Get information about available voices.
        
        Returns:
            list: List of voice information dictionaries
        """
        voice_info = []
        for i, voice in enumerate(self.voices):
            voice_info.append({
                'index': i,
                'id': voice.id,
                'name': voice.name
            })
        return voice_info