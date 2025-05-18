import time
import sys
from voice_recognition import VoiceRecognizer
from speech import Speech
from file_management import FileManager
from os_management import OSManagement
from command_handler import CommandHandler
from auth.face_auth import FaceAuthenticator  # Import FaceAuthenticator

def main():
    """Main function to run the voice-controlled file and system management assistant."""
    # Initialize modules
    speech = Speech()
    voice_recognizer = VoiceRecognizer()
    file_manager = FileManager(speech)
    os_manager = OSManagement(speech)
    command_handler = CommandHandler(file_manager, os_manager, voice_recognizer)
    
    # Initialize FaceAuthenticator with the Speech object
    try:
        face_auth = FaceAuthenticator(speech=speech)
    except Exception as e:
        print(f"Failed to initialize FaceAuthenticator: {e}")
        speech.speak(f"Failed to initialize face authentication: {e}")
        sys.exit(1)

    # Perform face authentication
    print("Please authenticate using face recognition...")
    speech.speak("Please look at the camera to authenticate.")
    recognized_name, status = face_auth.verify_user()
    
    if status != "success":
        print(f"Authentication failed: {status}")
        speech.speak(f"Authentication failed: {status}. Exiting program.")
        sys.exit(1)
    
    print(f"Authentication successful! Welcome, {recognized_name}!")
    speech.speak(f"Authentication successful! Welcome, {recognized_name}!")

    # Flag to track if we've already asked about listing commands
    asked_about_commands = False

    # Start speaking
    speech.start_speaking()

    # Display initial message and settings
    settings = voice_recognizer.get_settings()
    print("\n--- Real-time Voice File and System Management Assistant (Ctrl+C to stop) ---")
    print(f"Settings: Chunk_duration={settings['Chunk_duration']}s, Silence_threshold={settings['Silence_threshold']}s, VAD_energy={settings['VAD_energy']}")
    print("Available commands: " + ", ".join(command_handler.get_command_list()))
    print("Speak clearly. The system will try to transcribe after you pause.")
    print("You can use 'it' to refer to the last created or opened item (e.g., 'open it' after 'create folder').")
    speech.speak("Say a command")

    # Start listening for voice input
    stream = voice_recognizer.start_listening()

    try:
        while True:
            time.sleep(0.1)  # Keep the main thread alive
            # Check for new transcriptions
            cmd_text = voice_recognizer.get_transcription()
            if cmd_text:
                # Execute the command(s) and check if any were recognized
                if not command_handler.execute_command(cmd_text):
                    print("Unknown command(s).")
                    # Track how many times we've had unknown commands
                    if not hasattr(main, 'unknown_command_count'):
                        main.unknown_command_count = 0
                    main.unknown_command_count += 1
                    
                    # First time: Speak the error and suggest listing commands
                    if main.unknown_command_count == 1:
                        speech.speak("Unknown command.")
                        print("You can say 'list commands' to hear all available commands or use 'it' for the last item.")
                        speech.speak("You can say list commands to hear all available commands or use 'it' for the last item.")
                    # Second time: Just a brief notification
                    elif main.unknown_command_count == 2:
                        speech.speak("Unknown command. Try saying list commands or use 'it'.")
                    # After that: Just make a sound to indicate error
                    else:
                        print("\a")  # Terminal bell sound
                        speech.speak("Unknown command")
                        
                    # Reset the counter after a while
                    if main.unknown_command_count > 5:
                        main.unknown_command_count = 0
                
                # Prompt for the next command only after processing all parts
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