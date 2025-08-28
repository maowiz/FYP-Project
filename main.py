# main.py (Updated with all functionality preserved)

import time
import sys
import traceback
import logging
from speech import Speech
from voice_recognition import VoiceRecognizer
from file_management import FileManager
from os_management import OSManagement
from command_handler import CommandHandler
from assistant_state import set_speaking

# Use the logger for cleaner output instead of print()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    """Professional voice assistant with proper state management."""
    logging.info("Initializing Professional Voice Assistant...")

    # --- Initialize core components ---
    try:
        speech = Speech()
        speech.start_speaking()
        logging.info("Speech module initialized")

        # The new HybridVoiceRecognizer is initialized here
        voice_recognizer = VoiceRecognizer()
        logging.info("VoiceRecognizer initialized")

        os_manager = OSManagement(speech)
        logging.info("OSManagement initialized")

        # Correctly pass all dependencies
        file_manager = FileManager(speech, os_manager, voice_recognizer)
        logging.info("FileManager initialized")

        command_handler = CommandHandler(
            file_manager,
            os_manager,
            voice_recognizer,
            speech,
            is_online=voice_recognizer.network_monitor.is_online
        )
        logging.info("CommandHandler initialized")
        file_manager.command_handler = command_handler # Link back for circular dependency

        # Start the background listening threads
        voice_recognizer.start_listening()
        logging.info("Voice recognition started")

        # ✅ FUNCTIONALITY PRESERVED: Connect the grid's pause/resume hooks
        if hasattr(os_manager, 'grid') and os_manager.grid:
            os_manager.grid.set_pause_resume(voice_recognizer.pause_listening, voice_recognizer.resume_listening)
            logging.info("Grid pause/resume functionality connected.")

    except Exception as e:
        # Use logging for better error tracking
        logging.critical(f"A fatal error occurred during initialization: {e}", exc_info=True)
        if 'speech' in locals() and speech.can_speak():
            speech.speak("A critical error occurred during initialization. Shutting down.")
        sys.exit(1)

    # --- Initial Greeting and Main Loop ---
    speech.speak("Professional voice assistant ready. Say a command.")
    
    # The recognizer starts active now, no need to manually resume here.
    # voice_recognizer.resume_listening() # This is handled by start_listening() in the new recognizer

    try:
        while True:
            # STATE 1: LISTENING (Get transcription from the background thread)
            transcription = voice_recognizer.get_transcription()

            if transcription:
                # STATE 2: PROCESSING
                logging.info(f"Heard: '{transcription}'")
                voice_recognizer.pause_listening() # Pause to prevent feedback loops
                set_speaking(True)

                response_text = None
                try:
                    # ✅ FIX: Using the correct method name
                    response_text = command_handler.execute_command(transcription)
                except Exception as e:
                    logging.error(f"Command execution error: {e}", exc_info=True)
                    response_text = "An error occurred while processing the command."

                # STATE 3: SPEAKING
                if response_text and isinstance(response_text, str):
                    logging.info(f"Speaking: '{response_text}'")
                    speech.speak(response_text)
                    # A slightly longer pause can help prevent cutting off speech
                    time.sleep(1.0)

                # STATE 4: RETURN TO LISTENING
                set_speaking(False)
                voice_recognizer.resume_listening()

            else:
                # Efficiently wait without pinning the CPU
                time.sleep(0.05)

    except KeyboardInterrupt:
        logging.warning("Shutdown initiated by user.")
        speech.speak("Shutting down professional voice assistant.")
    except Exception as e:
        logging.critical(f"Critical error in main loop: {e}", exc_info=True)
        if 'speech' in locals() and speech.can_speak():
            speech.speak("A critical error occurred. Shutting down.")
    finally:
        logging.info("Terminating assistant processes...")
        if 'voice_recognizer' in locals():
            voice_recognizer.stop_listening()
        if 'speech' in locals():
            speech.stop_speaking()
        logging.info("Professional voice assistant terminated.")
        sys.exit(0)

if __name__ == "__main__":
    main()