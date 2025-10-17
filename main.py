# main.py (Updated with Hybrid AI Core)

import time 
import sys
import traceback            
import traceback            
import logging               
import os     
from speech import Speech
import pyautogui
from voice_recognition import VoiceRecognizer
from file_management import FileManager
from os_management import OSManagement          
from command_handler import CommandHandler
from assistant_state import set_speaking

# --- NEW IMPORT ---
# This brings in the new intelligent brain for your assistant.
from hybrid_processor import HybridCommandProcessor
# --------------------

# Use the logger for cleaner output instead of print()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    """Professional voice assistant with a hybrid AI core and proper state management."""
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

        # --- MODIFICATION: Instantiate both handlers ---
        # 1. The original, fast, fuzzy command handler. It will now be used as a tool.
        original_command_handler = CommandHandler(
            file_manager,
            os_manager,
            voice_recognizer,
            speech,
            is_online=voice_recognizer.network_monitor.is_online
        )
        logging.info("Original CommandHandler initialized.")

        # Define a configuration dictionary to centralize settings.
        # This is the most robust way to fix the model_path issue.
        script_dir = os.path.dirname(os.path.abspath(__file__))
        model_filename = "qwen2.5-0.5b-instruct-q4_k_m.gguf"
        model_path = os.path.join(script_dir, model_filename)

        config = {
            'enable_llm': True,
            'model_path': model_path,
            'confidence_threshold': 0.6,
            'llm_timeout': 10
        }

        # 2. The new hybrid processor that uses the original handler and the LLM.
        # This is now the main brain of the assistant.
        hybrid_processor = HybridCommandProcessor(original_command_handler, config=config)
        original_command_handler.hybrid_processor = hybrid_processor # Link back for translation
        logging.info("HybridCommandProcessor initialized.")



        # Link back for circular dependency (this remains unchanged)
        file_manager.command_handler = original_command_handler

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
        if 'speech' in locals() and speech.can_speak_flag:
            speech.speak("A critical error occurred during initialization. Shutting down.")
        sys.exit(1)

    # --- Initial Greeting and Main Loop ---
    speech.speak("Professional voice assistant ready. Say a command.")
    
    # The recognizer starts active now, no need to manually resume here.
    # voice_recognizer.resume_listening() # This is handled by start_listening() in the new recognizer

    dictation_mode = False
    note_taking_mode = False

    try:
        while True:
            # STATE 1: LISTENING (Get transcription from the background thread)
            transcription = voice_recognizer.get_transcription()

            if transcription:
                transcription_lower = transcription.lower().strip()

                # --- NOTE-TAKING MODE LOGIC ---
                if transcription_lower in ["take a note", "add a note", "new note", "write a note", "note this down"]:
                    if not note_taking_mode:
                        note_taking_mode = True
                        speech.speak("What's the note?")
                    continue  # Skip the rest of the loop, wait for the note content

                if note_taking_mode:
                    try:
                        with open("notes.txt", "a", encoding="utf-8") as f:
                            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                            f.write(f"{timestamp} - Note: {transcription}\n")
                        logging.info(f"Note taken: '{transcription}'")
                        speech.speak("Note taken.")
                    except Exception as e:
                        logging.error(f"Failed to write note: {e}")
                        speech.speak("Sorry, I couldn't save that note.")
                    finally:
                        note_taking_mode = False
                    continue # Immediately go back to listening

                # --- DICTATION MODE LOGIC ---
                # Check for commands to enter/exit dictation mode first.
                if transcription_lower in ["start dictation", "start dictation mode", "begin dictation", "dictation on"]:
                    if not dictation_mode:
                        dictation_mode = True
                        speech.speak("Dictation mode started.")
                    continue # Skip the rest of the loop

                if transcription_lower in ["stop dictation", "stop dictation mode", "end dictation", "dictation off"]:
                    if dictation_mode:
                        dictation_mode = False
                        speech.speak("Dictation mode stopped.")
                    continue # Skip the rest of the loop

                # If in dictation mode, type the transcription and bypass command processing.
                if dictation_mode:
                    logging.info(f"Dictating: '{transcription}'")
                    pyautogui.write(transcription + ' ')
                    continue # Immediately go back to listening
                # --- END DICTATION MODE LOGIC ---

                # STATE 2: PROCESSING
                logging.info(f"Heard: '{transcription}'")
                voice_recognizer.pause_listening() # Pause to prevent feedback loops
                set_speaking(True)

                response_text = None
                try:
                    # --- THE ONLY CHANGE IN THE MAIN LOOP ---
                    # OLD WAY: response_text = original_command_handler.execute_command(transcription)
                    # NEW WAY: The hybrid processor now handles all incoming text.
                    response_text = hybrid_processor.process(transcription)
                    # ------------------------------------------
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
        if 'speech' in locals() and speech.can_speak_flag:
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
