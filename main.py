# main.py (Updated with Hybrid AI Core)
from __future__ import annotations

import time 
import sys
import traceback            
import traceback            
import logging               
import os     
import pyautogui
from assistant_state import set_speaking

# --- NEW IMPORTS ---
import asyncio
import websockets
import threading
import queue
import json
from assistant_state import is_speaking  # Make sure to import is_speaking
from typing import Optional, Any, Union, cast, TYPE_CHECKING

if TYPE_CHECKING:
    from speech import Speech
    from voice_recognition import VoiceRecognizer
    from file_management import FileManager
    from os_management import OSManagement
    from command_handler import CommandHandler
# ------------------

# --- NEW IMPORT ---
# This brings in the new intelligent brain for your assistant.
from hybrid_processor import HybridCommandProcessor
# --------------------

# Use the logger for cleaner output instead of print()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- NEW: Queues for UI Communication ---
# Queue for messages going from Python -> UI
ui_message_queue = queue.Queue()
# Queue for messages going from UI -> Python
python_command_queue = queue.Queue()
# Cache historical init_progress messages so late UI clients can sync state
progress_history: list[dict[str, Any]] = []


def push_progress(percent: float, message: str, module: str | None = None, status: str | None = None, system_ready: bool = False) -> None:
    """Send a structured init_progress packet to the UI."""
    payload: dict[str, Any] = {
        "type": "init_progress",
        "percent": percent,
        "message": message,
    }
    if module:
        payload["module"] = module
    if status:
        payload["status"] = status
    if system_ready:
        payload["systemReady"] = True
    ui_message_queue.put(payload)

# --- NEW: WebSocket Server Logic ---
connected_clients = set()
ws_loop = None  # Global variable to store the WebSocket event loop
dictation_mode = False  # Global variable for dictation mode

async def broadcast_message(message_json):
    """Sends a JSON message to all connected UI clients."""
    if connected_clients:
        # Create a list of tasks to send to all clients
        tasks = [client.send(message_json) for client in connected_clients]
        # Wait for all tasks to complete
        await asyncio.gather(*tasks, return_exceptions=True)

async def websocket_handler(websocket):
    """Handles WebSocket connections."""
    connected_clients.add(websocket)
    logging.info("UI connected.")
    try:
        # Send current state on connect
        current_state = "dictation" if dictation_mode else "listening"
        await websocket.send(json.dumps({"type": "state", "phase": current_state}))

        # Replay cached init progress so late clients immediately catch up
        if progress_history:
            for past_message in progress_history:
                try:
                    await websocket.send(json.dumps(past_message))
                except Exception as e:
                    logging.error(f"Failed to replay init progress: {e}")
                    break

        async for message in websocket:
            # Messages from UI -> Python
            try:
                data = json.loads(message)
                logging.info(f"Received from UI: {data}")
                # Add to command queue for main.py to process
                python_command_queue.put(data)
            except json.JSONDecodeError:
                logging.warning(f"Received invalid JSON from UI: {message}")
            
    except websockets.exceptions.ConnectionClosed:
        logging.info("UI disconnected.")
    except Exception as e:
        logging.error(f"Error in websocket_handler: {e}")
    finally:
        connected_clients.discard(websocket)

# This is an async function that will contain all our server logic
async def main_server_task():
    """Runs all async tasks for the server."""
    # Create the task that polls the queue and sends updates
    asyncio.create_task(send_ui_updates())
    
    # Start the server and run forever
    server = await websockets.serve(websocket_handler, "127.0.0.1", 8765)
    logging.info("WebSocket server started on ws://127.0.0.1:8765")
    
    # Notify that server is ready
    try:
        ui_message_queue.put({"type": "server_status", "status": "ready"})
    except Exception as e:
        logging.error(f"Error sending server ready message: {e}")
        
    await server.wait_closed()

def start_websocket_server():
    """Starts the WebSocket server in a separate thread."""
    global ws_loop
    ws_loop = asyncio.new_event_loop()
    asyncio.set_event_loop(ws_loop)

    try:
        # Run the main async task until it completes (which is forever)
        ws_loop.run_until_complete(main_server_task())
    except asyncio.CancelledError:
        logging.info("WebSocket server task was cancelled.")
    except Exception as e:
        logging.error(f"WebSocket server thread failed: {e}", exc_info=True)
    finally:
        logging.info("WebSocket server shutting down.")
        if ws_loop and ws_loop.is_running():
            ws_loop.call_soon_threadsafe(ws_loop.stop)
        if ws_loop:
            ws_loop.close()

async def send_ui_updates():
    """Routinely checks the UI queue and broadcasts messages."""
    while True:
        try:
            # Get message from the thread-safe queue
            message_dict = ui_message_queue.get_nowait()

            if message_dict.get("type") == "init_progress":
                # Store a copy so new clients can be replayed the sequence later
                progress_history.append(dict(message_dict))
                # Keep history from growing unbounded (init sequence is short)
                if len(progress_history) > 100:
                    progress_history.pop(0)

            message_json = json.dumps(message_dict)
            # Await the broadcast (which is async)
            await broadcast_message(message_json)
        except queue.Empty:
            # No messages? sleep for a tiny bit to prevent a busy-loop
            await asyncio.sleep(0.05)
        except Exception as e:
            logging.error(f"Error in send_ui_updates: {e}")

# This function will replace the existing set_speaking
def new_set_speaking(is_speaking_bool):
    """Updates the global state and broadcasts it to the UI."""
    global global_is_speaking # Assuming you have a way to set this
    global_is_speaking = is_speaking_bool
    phase = "speaking" if is_speaking_bool else "listening" # Simplified
    try:
        ui_message_queue.put({"type": "state", "phase": phase})
    except Exception as e:
        logging.error(f"Error sending state update: {e}")

# --- END NEW WebSocket Server Logic ---

def main():
    """Professional voice assistant with a hybrid AI core and proper state management."""
    logging.info("Initializing Professional Voice Assistant...")
    push_progress(2, "Booting AI Core services...")

    # Initialize speech variable to avoid unbound variable error
    speech: Optional[Speech] = None
    voice_recognizer: Optional[VoiceRecognizer] = None
    original_command_handler: Optional[CommandHandler] = None
    file_manager: Optional[FileManager] = None
    hybrid_processor: Optional[HybridCommandProcessor] = None
    
    # --- NEW: Replace original set_speaking ---
    # We need to find where assistant_state.set_speaking is used and replace it.
    # In main.py, you import and use set_speaking.
    # And REPLACE it with:
    # from assistant_state import set_speaking as original_set_speaking
    
    # Then, right at the top of main():
    # This is a bit of a hack, but it overrides the function everywhere
    # Find this line:
    # from assistant_state import set_speaking
    # And REPLACE it with:
    # from assistant_state import set_speaking as original_set_speaking
    
    # Let's globally override it.
    import assistant_state
    assistant_state.set_speaking = new_set_speaking 
    
    # Also, find the line in main():
    # set_speaking(True) 
    # And change it to:
    # new_set_speaking(True)
    
    # And the line:
    # set_speaking(False)
    # And change it to:
    # new_set_speaking(False)

    # --- NEW: Start the WebSocket Server Thread ---
    ws_thread = threading.Thread(target=start_websocket_server, daemon=True)
    ws_thread.start()
    push_progress(4, "WebSocket bridge online. Awaiting backend startup...")

    # Heavy imports live inside main so we can show UI feedback while Python loads packages
    push_progress(8, "Loading AI dependencies (ctranslate2, comtypes, TTS)...")
    from speech import Speech
    from voice_recognition import VoiceRecognizer
    from file_management import FileManager
    from os_management import OSManagement
    from command_handler import CommandHandler
    push_progress(12, "Core dependencies loaded. Initializing subsystems...")

    # ... back in main() ...
    
    # (Your existing init code: speech, voice_recognizer, os_manager, etc.)
    # --- Initialize core components ---
    try:
        push_progress(18, "Engaging speech synthesis pipeline...", "speech", "loading")
        speech = Speech()
        speech.start_speaking()
        logging.info("Speech module initialized")
        push_progress(28, "Speech System online.", "speech", "done")

        push_progress(32, "Activating voice engine...", "voiceEngine", "loading")
        # The new HybridVoiceRecognizer is initialized here
        voice_recognizer = VoiceRecognizer()
        logging.info("VoiceRecognizer initialized")
        push_progress(48, "Voice Engine online.", "voiceEngine", "done")

        os_manager = OSManagement(speech)
        logging.info("OSManagement initialized")

        # Correctly pass all dependencies
        file_manager = FileManager(speech, os_manager, voice_recognizer)
        logging.info("FileManager initialized")
        push_progress(56, "System interfaces initialized (OS + File Manager).")

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

        push_progress(65, "Loading knowledge core + local LLM...", "knowledgeCore", "loading")
        # 2. The new hybrid processor that uses the original handler and the LLM.
        # This is now the main brain of the assistant.
        hybrid_processor = HybridCommandProcessor(original_command_handler, config=config)
        if original_command_handler:
            # Cast to avoid type checker issues
            original_command_handler.hybrid_processor = cast(Any, hybrid_processor) # Link back for translation
        logging.info("HybridCommandProcessor initialized.")
        push_progress(88, "Knowledge Core online. Finalizing subsystems...", "knowledgeCore", "done")



        # Link back for circular dependency (this remains unchanged)
        if file_manager:
            # Cast to avoid type checker issues
            file_manager.command_handler = cast(Any, original_command_handler)

        # Start the background listening threads
        if voice_recognizer:
            voice_recognizer.start_listening()
        logging.info("Voice recognition started")

        logging.info("Core initialization complete.")
        push_progress(94, "Running final diagnostics...")

        # ✅ FUNCTIONALITY PRESERVED: Connect the grid's pause/resume hooks
        if hasattr(os_manager, 'grid') and os_manager.grid:
            if voice_recognizer:
                os_manager.grid.set_pause_resume(voice_recognizer.pause_listening, voice_recognizer.resume_listening)
            logging.info("Grid pause/resume functionality connected.")

    except Exception as e:
        # Use logging for better error tracking
        logging.critical(f"A fatal error occurred during initialization: {e}", exc_info=True)
        push_progress(100, "Initialization failed. Check backend logs.", system_ready=False)
        if speech:
            speech.speak("A critical error occurred during initialization. Shutting down.")
        sys.exit(1)

    # --- Initial Greeting and Main Loop ---
    if speech:
        speech.speak("Professional voice assistant ready. Say a command.")
    push_progress(100, "System Ready", system_ready=True)
    
    # --- NEW: Send initial state to UI ---
    # We need to wait a moment for the server to be ready
    time.sleep(2) # Increased delay to ensure server is ready
    try:
        ui_message_queue.put({"type": "state", "phase": "listening"})
    except Exception as e:
        logging.error(f"Error sending initial state: {e}")

    global dictation_mode
    dictation_mode = False
    note_taking_mode = False
    transcription: Optional[str] = None

    try:
        while True:
            # --- NEW: Check for commands from the UI ---
            try:
                ui_command = python_command_queue.get_nowait()
                if ui_command.get("type") == "run_command":
                    transcription = ui_command.get("text", "")
                    # We'll process this command just like speech
                
                if ui_command.get("type") == "toggle_dictation":
                    # This is a special command to toggle dictation
                    if not dictation_mode:
                         dictation_mode = True
                         if speech:
                             speech.speak("Dictation mode started.")
                         try:
                             ui_message_queue.put({"type": "state", "phase": "dictation"})
                         except Exception as e:
                             logging.error(f"Error sending dictation state: {e}")
                    else:
                         dictation_mode = False
                         if speech:
                             speech.speak("Dictation mode stopped.")
                         try:
                             ui_message_queue.put({"type": "state", "phase": "listening"})
                         except Exception as e:
                             logging.error(f"Error sending listening state: {e}")
                    continue # Skip the rest of the loop
            except queue.Empty:
                transcription = None # No command from UI
            
            # STATE 1: LISTENING
            if not transcription: # Only listen if no command from UI
                if voice_recognizer:
                    transcription = voice_recognizer.get_transcription()

            if transcription:
                transcription_lower = transcription.lower().strip()

                # --- NOTE-TAKING MODE LOGIC ---
                if transcription_lower in ["take a note", "add a note", "new note", "write a note", "note this down"]:
                    if not note_taking_mode:
                        note_taking_mode = True
                        if speech:
                            speech.speak("What's the note?")
                        try:
                            ui_message_queue.put({"type": "state", "phase": "note_taking"})
                        except Exception as e:
                            logging.error(f"Error sending note taking state: {e}")
                    continue  

                if note_taking_mode:
                    try:
                        with open("notes.txt", "a", encoding="utf-8") as f:
                            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                            f.write(f"{timestamp} - Note: {transcription}\n")
                        logging.info(f"Note taken: '{transcription}'")
                        if speech:
                            speech.speak("Note taken.")
                    except Exception as e:
                        logging.error(f"Failed to write note: {e}")
                        if speech:
                            speech.speak("Sorry, I couldn't save that note.")
                    finally:
                        note_taking_mode = False
                    try:
                        ui_message_queue.put({"type": "state", "phase": "listening"})
                    except Exception as e:
                        logging.error(f"Error sending listening state: {e}")
                    continue 

                # --- DICTATION MODE LOGIC ---
                # Check for commands to enter/exit dictation mode first.
                if transcription_lower in ["start dictation", "start dictation mode", "begin dictation", "dictation on"]:
                    if not dictation_mode:
                        dictation_mode = True
                        if speech:
                            speech.speak("Dictation mode started.")
                        try:
                            ui_message_queue.put({"type": "state", "phase": "dictation"})
                        except Exception as e:
                            logging.error(f"Error sending dictation state: {e}")
                    continue 

                if transcription_lower in ["stop dictation", "stop dictation mode", "end dictation", "dictation off"]:
                    if dictation_mode:
                        dictation_mode = False
                        if speech:
                            speech.speak("Dictation mode stopped.")
                        try:
                            ui_message_queue.put({"type": "state", "phase": "listening"})
                        except Exception as e:
                            logging.error(f"Error sending listening state: {e}")
                    continue 

                # If in dictation mode, type the transcription and bypass command processing.
                if dictation_mode:
                    logging.info(f"Dictating: '{transcription}'")
                    pyautogui.write(transcription + ' ')
                    try:
                        ui_message_queue.put({"type": "partial_transcript", "text": transcription})
                    except Exception as e:
                        logging.error(f"Error sending partial transcript: {e}")
                    continue 
                # --- END DICTATION MODE LOGIC ---

                # STATE 2: PROCESSING
                logging.info(f"Heard: '{transcription}'")
                try:
                    ui_message_queue.put({"type": "final_transcript", "text": transcription})
                    ui_message_queue.put({"type": "state", "phase": "processing"})
                except Exception as e:
                    logging.error(f"Error sending processing state: {e}")
                
                if voice_recognizer:
                    voice_recognizer.pause_listening() 
                new_set_speaking(True) # This will send "speaking" state

                response_text: Optional[str] = None
                try:
                    # --- THE ONLY CHANGE IN THE MAIN LOOP ---
                    # OLD WAY: response_text = original_command_handler.execute_command(transcription)
                    # NEW WAY: The hybrid processor now handles all incoming text.
                    if hybrid_processor:
                        response_text = hybrid_processor.process(transcription)
                    # ------------------------------------------
                except Exception as e:
                    logging.error(f"Command execution error: {e}", exc_info=True)
                    response_text = "An error occurred while processing the command."

                # STATE 3: SPEAKING
                if response_text and isinstance(response_text, str):
                    logging.info(f"Speaking: '{response_text}'")
                    try:
                        ui_message_queue.put({"type": "assistant_response", "text": response_text})
                    except Exception as e:
                        logging.error(f"Error sending assistant response: {e}")
                    if speech:
                        speech.speak(response_text)
                    # A slightly longer pause can help prevent cutting off speech
                    time.sleep(1.0)

                # STATE 4: RETURN TO LISTENING
                new_set_speaking(False) # This will send "listening" state
                if dictation_mode:
                     try:
                         ui_message_queue.put({"type": "state", "phase": "dictation"})
                     except Exception as e:
                         logging.error(f"Error sending dictation state: {e}")
                else:
                     try:
                         ui_message_queue.put({"type": "state", "phase": "listening"})
                     except Exception as e:
                         logging.error(f"Error sending listening state: {e}")
                
                if voice_recognizer:
                    voice_recognizer.resume_listening()

            else:
                # Efficiently wait without pinning the CPU
                time.sleep(0.05)

    except KeyboardInterrupt:
        logging.warning("Shutdown initiated by user.")
        if speech:
            speech.speak("Shutting down professional voice assistant.")
    except Exception as e:
        logging.critical(f"Critical error in main loop: {e}", exc_info=True)
        if speech:
            speech.speak("A critical error occurred. Shutting down.")
    finally:
        logging.info("Terminating assistant processes...")
        if 'ws_loop' in globals() and ws_loop:
            ws_loop.call_soon_threadsafe(ws_loop.stop) # Stop the server loop
        if voice_recognizer:
            voice_recognizer.stop_listening()
        if speech:
            speech.stop_speaking()
        logging.info("Professional voice assistant terminated.")
        sys.exit(0)

if __name__ == "__main__":
    main()