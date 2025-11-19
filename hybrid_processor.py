# hybrid_processor.py - IMPROVED VERSION
import logging
import time
from intent_classifier import IntentClassifier

class HybridCommandProcessor:
    def __init__(self, existing_command_handler, config=None):
        """Initialize with configuration options"""
        self.config = config or {}
        self.command_handler = existing_command_handler
        
        # Configuration
        self.confidence_threshold = self.config.get('confidence_threshold', 0.6)
        self.enable_llm = self.config.get('enable_llm', True)
        self.max_llm_timeout = self.config.get('llm_timeout', 10)
        
        # Load components
        try:
            print("Loading Hybrid Processor...")
            self.intent_classifier = IntentClassifier(existing_command_handler)
            
            if self.enable_llm:
                try:
                    from optimized_llm_handler import OptimizedLLMHandler
                    model_path = self.config.get('model_path')
                    # Explicitly handle the case where model_path is not in the config.
                    # This prevents passing `model_path=None` to the handler, making
                    # the call robust even if the handler doesn't correctly check for None.
                    if model_path:
                        self.llm_handler = OptimizedLLMHandler(model_path=model_path)
                    else:
                        self.llm_handler = OptimizedLLMHandler()
                except ImportError as e:
                    logging.warning(f"LLM dependencies not found: {e}. The LLM is disabled.")
                    logging.warning("To enable the LLM, please run: pip install llama-cpp-python")
                    self.llm_handler = None
            else:
                self.llm_handler = None
                
            # Cache for common queries
            self.response_cache = {}
            self.cache_size = 50
            
            print("Hybrid Processor ready.")
            
        except Exception as e:
            logging.error(f"Failed to initialize Hybrid Processor: {e}")
            raise

    def process(self, text: str):
        """Process with caching and fallback"""
        # Define action commands that should execute every time, not be cached
        action_commands = [
            "switch window", "change wallpaper", "show desktop", "minimize all windows",
            "restore windows", "maximize window", "minimize window", "close window",
            "move window left", "move window right", "take screenshot", "go to desktop",
            "empty recycle bin", "scroll up", "scroll down", "scroll left", "scroll right",
            "stop scrolling", "copy", "paste", "select all", "remove", "undo", "redo",
            "open word", "save file", "volume up", "volume down", "mute", "maximize volume",
            "set volume", "brightness up", "brightness down", "maximize brightness",
            "set brightness", "show grid", "hide grid", "click cell", "double click cell",
            "right click cell", "drag from", "drop on", "zoom cell", "exit zoom",
            "set grid size", "open my computer", "open disk", "create folder",
            "open folder", "delete folder", "rename folder", "go back", "run application"
        ]
        
        # Check cache first, but skip action commands that need to execute every time
        is_action_command = any(text.lower().strip().startswith(cmd) for cmd in action_commands)
        if text.lower() in self.response_cache and not is_action_command:
            logging.info("Using cached response")
            return self.response_cache[text.lower()]
        
        try:
            # Classify intent
            intent, confidence, use_llm = self.intent_classifier.classify(
                text, 
                confidence_threshold=self.confidence_threshold
            )

            # --- NEW: Special Handling for "write essay" ---
            # Before treating it as a general query, check if it's the essay command.
            # This ensures the typing action is triggered instead of just speaking the result.            
            text_lower = text.lower()
            is_essay_command = any(text_lower.startswith(syn) for syn in ["write an essay on", "write about", "compose an essay on"])
            is_type_on_word_command = ("write" in text_lower and " on word" in text_lower) or ("type" in text_lower and " on word" in text_lower)
            # --- NEW: Special Handling for "save file" ---
            is_save_command = any(keyword in text_lower for keyword in ["save file", "save this", "save it"])
            # --- NEW: Special Handling for "remove this" ---
            is_remove_command = any(keyword in text_lower for keyword in ["remove this", "delete this", "remove selection", "delete selection", "clear selection"])
            # --- FIX: Expanded Special Handling for "chatgpt" to include suffix-only cases ---
            is_chatgpt_command = any(trigger in text_lower for trigger in ["chatgpt", "chat gpt", "ask chatgpt", "tell chatgpt"]) or any(text_lower.endswith(" " + suffix) for suffix in ["on chatgpt", "on gpt", "on chat gpt", "gpt"])
            # --- NEW: Special Handling for "summarize" ---
            is_summarize_command = any(keyword in text_lower for keyword in ["summarize", "summarise", "summary"])


            # If it's a special command that can be misclassified as a general query, handle it directly.
            if is_save_command or is_remove_command or is_chatgpt_command or is_summarize_command or (intent == 'general_query' and (is_essay_command or is_type_on_word_command)):
                if is_save_command:
                    logging.info("Special case: 'save file' command detected. Executing directly.")
                elif is_remove_command:
                    logging.info("Special case: 'remove selection' command detected. Executing directly.")
                elif is_chatgpt_command:
                    logging.info("Special case: 'chatgpt' command detected. Executing directly.")
                elif is_summarize_command:
                    logging.info("Special case: 'summarize' command detected. Executing directly.")
                else:
                    logging.info("Special case: 'write essay' command detected. Executing directly.")
                response = self.command_handler.execute_command(text)
                self._cache_response(text.lower(), response)
                return response
            
            logging.info(f"Intent: {intent}, Confidence: {confidence:.2%}, Use LLM: {use_llm}")
            
            # Route request
            if not use_llm or not self.llm_handler:
                # Use fuzzy command handler
                response = self.command_handler.execute_command(text)
            else:
                # Use LLM with timeout
                start_time = time.time()
                
                if intent == 'command':
                    response = self._llm_command_interpretation(text)
                else:
                    response = self._llm_conversation(text)
                
                # Check timeout
                if time.time() - start_time > self.max_llm_timeout:
                    logging.warning("LLM timeout - falling back to command handler")
                    response = self.command_handler.execute_command(text)
            
            # --- FIX: Only cache string responses (conversations), not action results or dynamic commands ---
            # Define commands whose output changes over time and should not be cached.
            dynamic_commands = [
                "read last note", "tell time", "tell date", "tell day", 
                "show system info", "check disk space", "tell weather",
                "check internet speed", "check bmi", "read clipboard",
                "summarize clipboard"
            ]
            
            # Define action commands that should execute every time, not be cached
            action_commands = [
                "switch window", "change wallpaper", "show desktop", "minimize all windows",
                "restore windows", "maximize window", "minimize window", "close window",
                "move window left", "move window right", "take screenshot", "go to desktop",
                "empty recycle bin", "scroll up", "scroll down", "scroll left", "scroll right",
                "stop scrolling", "copy", "paste", "select all", "remove", "undo", "redo",
                "open word", "save file", "volume up", "volume down", "mute", "maximize volume",
                "set volume", "brightness up", "brightness down", "maximize brightness",
                "set brightness", "show grid", "hide grid", "click cell", "double click cell",
                "right click cell", "drag from", "drop on", "zoom cell", "exit zoom",
                "set grid size", "open my computer", "open disk", "create folder",
                "open folder", "delete folder", "rename folder", "go back", "run application"
            ]
            
            is_dynamic_command = any(text.lower().strip().startswith(cmd) for cmd in dynamic_commands)
            is_action_command = any(text.lower().strip().startswith(cmd) for cmd in action_commands)

            # Only cache conversational responses, not action commands or dynamic commands
            if isinstance(response, str) and not is_dynamic_command and not is_action_command:
                self._cache_response(text.lower(), response)
                
            return response
            
        except Exception as e:
            logging.error(f"Processing error: {e}")
            # Fallback to command handler
            return self.command_handler.execute_command(text)
    
    def _cache_response(self, text, response):
        """Simple LRU cache implementation for string responses."""
        if len(self.response_cache) >= self.cache_size:
            # Remove oldest entry
            oldest = next(iter(self.response_cache))
            del self.response_cache[oldest]
        self.response_cache[text] = response
    
    def _llm_command_interpretation(self, text: str):
        """Enhanced command interpretation with validation"""
        # Get command descriptions for better context
        command_descriptions = self._get_command_descriptions()
        
        try:
            for response, is_command in self.llm_handler.process_fast(
                text, 
                command_descriptions
            ):
                if is_command and "CMD:" in response:
                    # Extract and validate command
                    command = response.split("CMD:", 1)[1].strip()
                    
                    # Validate command exists
                    if self._validate_command(command):
                        logging.info(f"LLM interpreted: '{command}'")
                        return self.command_handler.execute_command(command)
                    else:
                        logging.warning(f"Invalid command from LLM: '{command}'")
                        
        except Exception as e:
            logging.error(f"LLM command interpretation failed: {e}")
        
        return "I couldn't understand that command. Please try rephrasing."
    
    def _validate_command(self, command):
        """Validate that command is safe and exists"""
        # Remove any potentially dangerous characters
        safe_command = command.lower().strip()
        
        # Check if command exists in handler
        for known_cmd in self.command_handler.COMMANDS.keys():
            if safe_command.startswith(known_cmd):
                return True
        
        # Check synonyms
        for synonym in self.command_handler.COMMAND_SYNONYMS.keys():
            if safe_command.startswith(synonym):
                return True
                
        return False
    
    def _get_command_descriptions(self):
        """Get commands with descriptions for LLM context"""
        descriptions = []
        
        # Sample of important commands with descriptions
        command_info = {
            "create folder": "Creates a new folder",
            "open folder": "Opens a specified folder",
            "delete folder": "Deletes a folder",
            "increase volume": "Increases system volume",
            "take screenshot": "Captures the screen",
            "play on youtube": "Plays media on YouTube",
            "tell weather": "Gets weather information",
            "check internet speed": "Tests internet connection speed"
        }
        
        for cmd, desc in command_info.items():
            if cmd in self.command_handler.COMMANDS:
                descriptions.append(f"{cmd}: {desc}")
        
        return descriptions[:20]  # Limit to avoid token overflow
    
    def _llm_conversation(self, text: str):
        """Handle general conversation with error recovery"""
        try:
            # This loop will only run once since process_fast yields a single final result
            for response, is_command in self.llm_handler.process_fast(text):
                # --- NEW: Robustness Check ---
                # If the LLM identified a command even when the classifier didn't,
                # we trust the LLM and execute it.
                if is_command and "CMD:" in response:
                    command = response.split("CMD:", 1)[1].strip()
                    if self._validate_command(command):
                        logging.info(f"LLM overrode classifier and interpreted command: '{command}'")
                        return self.command_handler.execute_command(command)
                    else:
                        logging.warning(f"LLM suggested an invalid command during conversation: '{command}'")
                
                # If it's not a command, proceed with the conversational response
                full_response = response.strip()
                
                # Ensure response isn't too long for TTS
                if len(full_response) > 500:
                    full_response = full_response[:497] + "..."
                
                return full_response or "I'm not sure how to respond to that."
            
        except Exception as e:
            logging.error(f"LLM conversation failed: {e}")
            return "I'm having trouble processing that request right now."