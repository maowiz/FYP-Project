import time
import sys
from voice_recognition import VoiceRecognizer
from speech import Speech
from file_management import FileManager
from command_handler import CommandHandler

def main():
    """Main function to run the voice-controlled file management assistant."""
    # Initialize modules
    voice_recognizer = VoiceRecognizer()
    speech = Speech()
    file_manager = FileManager(speech)
    command_handler = CommandHandler(file_manager, voice_recognizer)
    
    # Flag to track if we've already asked about listing commands
    asked_about_commands = False

    # Start speaking
    speech.start_speaking()

    # Display initial message and settings
    settings = voice_recognizer.get_settings()
    print("\n--- Real-time Voice File Management Assistant (Ctrl+C to stop) ---")
    print(f"Settings: Chunk_duration={settings['Chunk_duration']}s, Silence_threshold={settings['Silence_threshold']}s, VAD_energy={settings['VAD_energy']}")
    print("Available commands: " + ", ".join(command_handler.get_command_list()))
    print("Speak clearly. The system will try to transcribe after you pause.")
    speech.speak("Say a command")

    # Start listening for voice input
    stream = voice_recognizer.start_listening()

    try:
        while True:
            time.sleep(0.1)  # Keep the main thread alive
            # Check for new transcriptions
            cmd_text = voice_recognizer.get_transcription()
            if cmd_text:
                # Execute the command if recognized
                if not command_handler.execute_command(cmd_text):
                    print("Unknown command.")
                    
                    # Track how many times we've had unknown commands
                    if not hasattr(main, 'unknown_command_count'):
                        main.unknown_command_count = 0
                    main.unknown_command_count += 1
                    
                    # First time: Speak the error and suggest listing commands
                    if main.unknown_command_count == 1:
                        speech.speak("Unknown command.")
                        print("You can say 'list commands' to hear all available commands.")
                        speech.speak("You can say list commands to hear all available commands.")
                    # Second time: Just a brief notification
                    elif main.unknown_command_count == 2:
                        speech.speak("Unknown command. Try saying list commands.")
                    # After that: Just make a sound to indicate error
                    else:
                        # Play a short error sound (beep)
                        print("\a")  # Terminal bell sound
                        speech.speak("Unknown command")
                        
                    # Reset the counter after a while
                    if main.unknown_command_count > 5:
                        main.unknown_command_count = 0
                
                # Prompt for the next command
                print("\nSay another command:")
                speech.speak("Say another command")

    except KeyboardInterrupt:
        print("\nProgram terminated by user.")
        speech.speak("Program terminated by user.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        speech.speak(f"An unexpected error occurred: {e}")
    finally:
        stream.stop()
        stream.close()
        print("Goodbye!")
        speech.speak("Goodbye")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nProgram terminated by user.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")