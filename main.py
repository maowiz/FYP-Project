import time
from speech import Speech
from file_management import FileManager
from os_management import OSManagement
from voice_recognition import VoiceRecognizer
from auth.face_auth import FaceAuthenticator  # Correct import
from command_handler import CommandHandler
from general_command_handler import GeneralCommandHandler

def main():
    print("Main function started.") # New print
    # Initialize components
    speech = Speech()
    speech.start_speaking()  # Enable speech output
    file_manager = FileManager(speech)
    os_manager = OSManagement(speech)
    
    print("Initializing VoiceRecognizer...") # New print
    voice_recognizer = VoiceRecognizer()
    print("Calling start_listening...") # New print
    # Store the audio stream to keep it active
    audio_stream = voice_recognizer.start_listening()  # Start listening for voice commands
    print("start_listening finished.") # New print

    print("Initializing FaceAuthenticator...") # New print
    face_auth = FaceAuthenticator()
    print("FaceAuthenticator initialized.") # New print

    # Perform face authentication
    print("Initiating face authentication...")
    file_manager.speech.speak("Initiating face authentication. Please look at the camera.")
    
    authenticated = False
    max_attempts = 3
    attempt = 1
    
    while attempt <= max_attempts:
        print(f"Attempt {attempt} of {max_attempts}...")
        recognized_name, status = face_auth.authenticate()
        if recognized_name:
            print(f"Face authentication successful! Welcome, {recognized_name}!")
            file_manager.speech.speak(f"Face authentication successful! Welcome, {recognized_name}!")
            authenticated = True
            break
        else:
            print(f"Face authentication failed. Status: {status}")
            file_manager.speech.speak("Face authentication failed. Please try again.")
            attempt += 1
            time.sleep(2)
    
    if not authenticated:
        print("Maximum authentication attempts reached. Exiting program.")
        file_manager.speech.speak("Maximum authentication attempts reached. Exiting program.")
        return

    # Initialize CommandHandler after authentication
    command_handler = CommandHandler(file_manager, os_manager, voice_recognizer)
    
    # Initialize GeneralCommandHandler
    command_handler.general_handler = GeneralCommandHandler(file_manager, command_handler)

    # Main command loop
    print("Starting voice command system. Say 'stop' to exit.")
    file_manager.speech.speak("Starting voice command system. I am ready to assist you. Say stop to exit.")
    
    try:
        while True:
            transcription = voice_recognizer.get_transcription()
            if transcription:
                print(f"You said: {transcription}")
                executed = command_handler.execute_command(transcription)
                if not executed:
                    print("Command not recognized. Say 'list commands' for available commands.")
                    file_manager.speech.speak("Command not recognized. Say list commands for available commands.")
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\nProgram interrupted by user.")
        file_manager.speech.speak("Program interrupted by user. Goodbye.")
    except Exception as e:
        print(f"An error occurred: {e}")
        file_manager.speech.speak("An error occurred. Please try again or restart the program.")

if __name__ == "__main__":
    main()