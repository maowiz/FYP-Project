import sys
import requests
import speedtest
import cv2
import os
from datetime import datetime
import ctypes
import re
import time
import threading
import pyautogui
import pyperclip
import logging
import tkinter as tk
import subprocess

NOTES_FILE = "notes.txt"

class GeneralCommandHandler:
    def __init__(self, file_manager, command_handler=None, speech=None):
        self.file_manager = file_manager
        self.command_handler = command_handler
        # Prefer explicitly injected speech; fall back to file_manager.speech
        self.speech = speech if speech is not None else getattr(file_manager, "speech", None)

    def handle_check_internet(self, cmd_text=None):
        """Handle the 'check internet' command. Returns a status string."""
        try:
            import socket
            socket.create_connection(("www.google.com", 80), timeout=2)
            print("Internet is connected.")
            return "Internet connection is active."
        except OSError as e:
            print(f"No internet connection: {e}")
            return "No internet connection detected."
        except Exception as e:
            print(f"Error checking internet: {e}")
            return "Error checking internet connection."
    
    def handle_tell_joke(self, cmd_text=None):
        """Handle the 'tell joke' command. Returns a joke string."""
        jokes = [
            "Why don’t skeletons fight each other? Because they don’t have the guts!",
            "What do you call fake spaghetti? An impasta!",
            "Why did the scarecrow win an award? Because he was outstanding in his field!",
            "Why don’t scientists trust atoms? Because they make up everything!",
            "What do you call a bear with no teeth? A gummy bear!",
            "Why don’t eggs tell jokes? They’d crack up!",
            "What did one wall say to the other wall? I'll meet you at the corner!",
            "Why did the bicycle fall over? It was two tired!",
            "What do you call cheese that isn’t yours? Nacho cheese!",
            "Why did the tomato turn red? Because it saw the salad dressing!",
            "What do you call a fish with no eyes? A fsh!",
            "Why did the banana go to the doctor? It wasn’t peeling well!",
            "How do you organize a space party? You planet!",
            "Why don’t skeletons play music in church? Because they have no organs!",
            "What do you call a lazy kangaroo? A pouch potato!",
            "Why did the coffee file a police report? It got mugged!",
        ]
        import random
        joke = random.choice(jokes)
        print(f"Selected joke: {joke}")
        return joke
    def handle_tell_time(self, cmd_text=None):
        """Handle the 'tell time' command. Returns current time string."""
        current_time = datetime.now().strftime("%I:%M %p")
        print(f"Current time: {current_time}")
        return f"The current time is {current_time}."

    def handle_tell_date(self, cmd_text=None):
        """Handle the 'tell date' command. Returns current date string."""
        current_date = datetime.now().strftime("%B %d, %Y")
        print(f"Current date: {current_date}")
        return f"Today's date is {current_date}."

    def handle_tell_day(self, cmd_text=None):
        """Handle the 'tell day' command. Returns current day string."""
        current_day = datetime.now().strftime("%A")
        print(f"Current day: {current_day}")
        return f"Today is {current_day}."

    def handle_system_info(self, cmd_text=None):
        """Handle the 'show system info' command. Returns a summary string."""
        import psutil
        import platform
        try:
            ram = psutil.virtual_memory()
            cpu = psutil.cpu_percent(interval=1)
            battery = psutil.sensors_battery()
            system = platform.system()
            ram_usage = f"RAM usage: {ram.percent}% ({ram.used / (1024 ** 3):.2f} GB used of {ram.total / (1024 ** 3):.2f} GB)"
            cpu_usage = f"CPU usage: {cpu}%"
            battery_status = f"Battery: {battery.percent}% (Plugged in: {battery.power_plugged})" if battery else "No battery detected"
            info = f"System: {system}. {ram_usage}. {cpu_usage}. {battery_status}."
            print(info)
            return info
        except Exception as e:
            print(f"Error retrieving system info: {e}")
            return "Error retrieving system information."

    def handle_add_numbers(self, cmd_text):
        """Handle the 'add numbers' command. Returns a result string or prompt."""
        numbers = self._extract_numbers(cmd_text)
        if len(numbers) >= 2:
            result = sum(numbers)
            print(f"Addition result: {result}")
            return f"The sum is {result}."
        else:
            print("Insufficient numbers for addition.")
            return "Please provide at least two numbers to add."

    def handle_subtract_numbers(self, cmd_text):
        """Handle the 'subtract numbers' command. Returns a result string or prompt."""
        numbers = self._extract_numbers(cmd_text)
        if len(numbers) >= 2:
            result = numbers[0] - sum(numbers[1:])
            print(f"Subtraction result: {result}")
            return f"The difference is {result}."
        else:
            print("Insufficient numbers for subtraction.")
            return "Please provide at least two numbers to subtract."

    def handle_multiply_numbers(self, cmd_text):
        """Handle the 'multiply numbers' command. Returns a result string or prompt."""
        numbers = self._extract_numbers(cmd_text)
        if len(numbers) >= 2:
            result = 1
            for num in numbers:
                result *= num
            print(f"Multiplication result: {result}")
            return f"The product is {result}."
        else:
            print("Insufficient numbers for multiplication.")
            return "Please provide at least two numbers to multiply."

    def handle_divide_numbers(self, cmd_text):
        """Handle the 'divide numbers' command. Returns a result string or prompt."""
        numbers = self._extract_numbers(cmd_text)
        if len(numbers) >= 2:
            try:
                result = numbers[0]
                for num in numbers[1:]:
                    if num == 0:
                        print("Division by zero attempted.")
                        return "Cannot divide by zero."
                    result /= num
                print(f"Division result: {result}")
                return f"The quotient is {result}."
            except Exception as e:
                print(f"Division error: {e}")
                return "Error performing division."
        else:
            print("Insufficient numbers for division.")
            return "Please provide at least two numbers to divide."

    def _extract_numbers(self, cmd_text):
        """Extract numbers from command text."""
        import re
        numbers = re.findall(r'\b\d+\b', cmd_text)
        return [int(num) for num in numbers]

    def handle_list_commands(self, cmd_text=None):
        """Return a single string listing available commands."""
        if not self.command_handler:
            return "Cannot list commands."
        commands = self.command_handler.get_command_list()
        return "Here are the available commands: " + ", ".join(commands)

    def handle_exit(self, cmd_text=None):
        """Handle the 'exit' command: speak and terminate."""
        print("Exiting the program...")
        if self.speech:
            self.speech.speak("Exiting the program.")
        elif hasattr(self.file_manager, "speech") and self.file_manager.speech:
            self.file_manager.speech.speak("Exiting the program.")
        sys.exit(0)

    def handle_tell_weather(self, cmd_text):
        """Handle the 'tell weather' command. Returns a weather report string."""
        import re
        # Extract city name from cmd_text (already handled by extract_parameters)
        city_match = re.search(r'(?:tell\s+weather\s+in|weather\s+in|what\'s\s+the\s+weather\s+in|weather\s+forecast\s+for|weather)\s+(.+)$', cmd_text, re.IGNORECASE)
        city = city_match.group(1).strip() if city_match else None
        if not city:
            print("No city name extracted from command.")
            return "Please specify a city for the weather."

        # OpenWeatherMap API setup
        API_KEY = "3fdfbf6d393a7b2a64310166ddef0dbc"
        BASE_URL = "http://api.openweathermap.org/data/2.5/weather?"
        COMPLETE_URL = f"{BASE_URL}q={city}&appid={API_KEY}&units=metric"

        try:
            response = requests.get(COMPLETE_URL)
            data = response.json()

            if data["cod"] != 200:
                if data.get("message") == "Invalid API key":
                    print(f"API error for {city}: Invalid API key. Please see https://openweathermap.org/faq#error401 for more info.")
                    return "The weather API key is invalid. Please update it."
                else:
                    print(f"API error for {city}: {data['message']}")
                    return f"Could not find weather for {city}."

            main = data["main"]
            weather = data["weather"][0]
            temp = main["temp"]
            feels_like = main["feels_like"]
            humidity = main["humidity"]
            description = weather["description"]

            weather_report = (
                f"In {city}, it's {description} with a temperature of {temp} degrees Celsius. "
                f"It feels like {feels_like} degrees with {humidity}% humidity."
            )
            print(f"Weather report for {city}: {weather_report}")
            return weather_report

        except requests.exceptions.RequestException as e:
            print(f"Request error: {e}")
            return "Sorry, I couldn't fetch the weather right now."
        except KeyError as e:
            print(f"Key error: {e}")
            return "Error parsing weather data."
        except Exception as e:
            print(f"Unexpected error: {e}")
            return "An unexpected error occurred while fetching weather."

    def handle_check_internet_speed(self, cmd_text=None):
        """Handle the 'check internet speed' command. Returns a result string."""
        try:
            import speedtest
            print("Starting internet speed test...")
            st = speedtest.Speedtest()
            st.get_best_server()  # Select the best server for accurate results
            download_speed = st.download() / 1_000_000  # Convert bits/s to Mbps
            upload_speed = st.upload() / 1_000_000      # Convert bits/s to Mbps
            response = f"Your internet speed is: Download {download_speed:.2f} Mbps, Upload {upload_speed:.2f} Mbps."
            print(f"Internet speed test result: {response}")
            return response
        except Exception as e:
            print(f"Error checking internet speed: {e}")
            error_message = "Failed to check internet speed. Please check your internet connection."
            return error_message
    
    def handle_take_photo(self, cmd_text=None):
        """Handle the 'take a photo' command. Returns a success/error string."""
        try:
            import cv2
            import os
            from datetime import datetime

            # Initialize the webcam
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                return "Could not access the webcam. Please ensure it’s connected."

            # Capture a single frame
            ret, frame = cap.read()
            if not ret:
                cap.release()
                return "Failed to capture photo."

            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"photo_{timestamp}.jpg"
            save_path = os.path.join(os.getcwd(), filename)

            # Save the image
            cv2.imwrite(save_path, frame)
            cap.release()
            print(f"Photo saved at: {save_path}")
            return f"Photo saved as {filename}."
        except Exception as e:
            error_message = "Error taking photo. Please check your webcam."
            print(f"Error taking photo: {e}")
            if 'cap' in locals():
                cap.release()
            return error_message

    def handle_check_bmi(self, cmd_text=None):
        """Handle the 'check bmi' command. Returns a result or prompt string."""
        try:
            # Extract weight and height from command text
            import re
            weight_match = re.search(r'(\d+\.?\d*)\s*kg', cmd_text.lower())
            height_match = re.search(r'(\d+\.?\d*)\s*(?:m|metre|meter)', cmd_text.lower())
            if not weight_match or not height_match:
                return "Please provide weight in kg and height in meters, e.g., 'check bmi 70 kg 1.7 m'."
            weight = float(weight_match.group(1))
            height = float(height_match.group(1))

            # Calculate BMI
            bmi = weight / (height ** 2)
            category = ""
            if bmi < 18.5:
                category = "Underweight"
            elif 18.5 <= bmi < 25:
                category = "Normal weight"
            elif 25 <= bmi < 30:
                category = "Overweight"
            else:
                category = "Obese"

            # Prepare and speak the response
            response = f"Your BMI is {bmi:.1f}, which is categorized as {category}."
            print(f"BMI calculation result: {response}")
            return response
        except Exception as e:
            error_message = "Error calculating BMI. Please check your input format."
            print(f"Error calculating BMI: {e}")
            return error_message
    def handle_set_wallpaper(self, cmd_text=None):
        """Handle the 'set wallpaper to image.jpg' command. Returns a result string."""
        try:
            # Extract the filename from the command
            import re
            match = re.search(r'set\s+wallpaper\s+to\s+(.+?)(?:\s|$)', cmd_text.lower())
            if not match:
                return "Please specify an image file, e.g., 'set wallpaper to image.jpg'."
            filename = match.group(1).strip()
            file_path = os.path.abspath(filename)  # Convert to absolute path

            # Check if the file exists
            if not os.path.exists(file_path):
                return f"File {filename} not found. Please ensure the image exists."

            # Set the wallpaper using Windows API
            SPI_SETDESKWALLPAPER = 20
            ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, file_path, 0)
            print(f"Wallpaper changed to: {file_path}")
            return f"Wallpaper set to {filename}."
        except Exception as e:
            error_message = "Error setting wallpaper. Please check the file path."
            print(f"Error setting wallpaper: {e}")
            return error_message
    
    def handle_check_disk_space(self, cmd_text=None):
        """Handle the 'check disk space' command for C: and D: drives. Returns a summary string."""
        import shutil
        drives = ['C', 'D']
        messages = []
        for drive in drives:
            path = f"{drive}:\\"
            try:
                total, used, free = shutil.disk_usage(path)
                total_gb = total / (1024 ** 3)
                free_gb = free / (1024 ** 3)
                used_gb = used / (1024 ** 3)
                msg = f"Drive {drive}: {free_gb:.2f} GB free out of {total_gb:.2f} GB."
                messages.append(msg)
                print(msg)
            except Exception as e:
                error_msg = f"Could not check disk space for drive {drive}."
                messages.append(error_msg)
                print(f"{error_msg} Error: {e}")
        return " ".join(messages) if messages else "No drives found."
    
    
    

    def handle_find_phone(self, _=None):
        """Ring the linked Android via Google Find My Device. Returns a status string."""
        import webbrowser
        url = "https://www.google.com/android/find"
        try:
            webbrowser.open(url)
            return "Opening Google Find My Device."
        except Exception as e:
            print("Find phone error:", e)
            return "Could not open Find My Device."

    
    
    def handle_countdown(self, cmd_text=None):
        """Handle 'countdown <seconds>' command. Returns a start message string."""
        import re
        if not cmd_text:
            return "Please say countdown followed by seconds, e.g., countdown 30."
        m = re.search(r'countdown (\d+)', cmd_text.lower())
        if not m:
            return "Please say countdown followed by seconds, e.g., countdown 30."
        secs = int(m.group(1))
        if secs <= 0 or secs > 3600:
            return "Please choose 1–3600 seconds."

        def _countdown_win():
            root = tk.Tk()
            root.title("Countdown")
            root.attributes('-topmost', True)
            root.geometry("200x100+500+300")
            root.overrideredirect(True)
            label = tk.Label(root, font=('Arial', 36), fg='white', bg='black')
            label.pack(fill='both', expand=True)
            for i in range(secs, 0, -1):
                label.config(text=str(i))
                root.update()
                time.sleep(1)
            root.destroy()
            # Allow end-of-countdown notification immediately
            if self.speech:
                self.speech.speak("Countdown finished.")

        threading.Thread(target=_countdown_win, daemon=True).start()
        return f"Starting countdown of {secs} seconds."
    
    def handle_spell(self, cmd_text=None):
        """Spell out text phonetically. Returns the spelled string."""
        import re
        print(f"handle_spell called with cmd_text: {cmd_text}")  # Debug log
        if not cmd_text:
            print("No cmd_text provided, returning False")
            return "Please say spell followed by the word."
        m = re.search(r'spell (.+)', cmd_text.lower())
        if not m:
            print("No match found in cmd_text, returning False")
            return "Please say spell followed by the word."
        word = m.group(1).strip()
        spelled = " ".join(word.upper())
        print(f"Spelling out: {spelled}")  # Debug log
        return spelled
    
    def handle_write_essay(self, topic, cmd_text=""):
        """Generates an essay on a given topic using the LLM and types it out."""
        if not topic:
            return "Please specify a topic for the essay."

        if not self.command_handler or not self.command_handler.hybrid_processor:
            return "The essay writing function is not available right now."

        # --- NEW: Check if the command includes "on word" ---
        write_on_word = " on word" in cmd_text.lower()

        if write_on_word:
            try:
                import pygetwindow as gw
                word_windows = gw.getWindowsWithTitle('Word')
                if not word_windows:
                    self.speech.speak("Microsoft Word is not open. I'll open it first.")
                    # Use the existing OS handler to open and maximize Word
                    if hasattr(self.command_handler, 'os_handler'):
                        self.command_handler.os_handler.handle_open_word()
                        time.sleep(4) # Give Word extra time to load before typing
                else:
                    # If Word is open, make sure it's focused
                    word_windows[0].activate()
            except (ImportError, Exception) as e:
                print(f"Could not check for or open Word due to an error: {e}")


        # Store the topic in context for the 'save file' command
        if hasattr(self.file_manager.os_manager, 'context'):
            self.file_manager.os_manager.context['last_essay_topic'] = topic

        self.speech.speak(f"Okay, writing a short essay about {topic}. Please wait a moment.")

        # Construct a prompt for the LLM
        prompt = f"Write a 2 to 3 paragraph essay about {topic}."
        
        # --- FIX: Call the LLM directly to avoid recursion ---
        # The hybrid_processor has the llm_handler instance.
        if not self.command_handler.hybrid_processor.llm_handler:
            return "The Language Model is not available for writing essays."
            
        essay_text = self.command_handler.hybrid_processor.llm_handler.generate_essay(prompt)
        # ----------------------------------------------------

        if not essay_text or "I'm sorry" in essay_text or "I couldn't" in essay_text:
             self.speech.speak("I'm sorry, I couldn't generate an essay on that topic.")
             return None
        
        # --- NEW: Conditional output ---
        if write_on_word:
            # Type the essay into the active window (Word)
            self.speech.speak("Here is the essay.")
            time.sleep(1) # Give user time to focus the desired window
            pyautogui.write(essay_text, interval=0.02)
            return None # No verbal response needed after typing
        else:
            # Print to terminal and speak the result
            print(f"\n--- Essay on {topic} ---\n{essay_text}\n---------------------\n")
            return essay_text # Return the text to be spoken
    
    def handle_send_to_chatgpt(self, query):
        """Opens ChatGPT, focuses the window, and types the query."""
        if not query:
            self.speech.speak("Please tell me what to send to ChatGPT.")
            return None

        self.speech.speak(f"Sending to ChatGPT: {query}")

        try:
            import pygetwindow as gw
            import webbrowser

            chatgpt_url = "https://chat.openai.com/"
            logging.info("Attempting to find ChatGPT window...")
            # Find ChatGPT window
            chat_windows = gw.getWindowsWithTitle('ChatGPT')
            
            if not chat_windows:
                self.speech.speak("ChatGPT is not open. I'll open it now.")
                logging.info("ChatGPT window not found. Opening new browser tab.")
                webbrowser.open(chatgpt_url)
                time.sleep(5) # Give the browser time to load the page
                chat_windows = gw.getWindowsWithTitle('ChatGPT')

            if chat_windows:
                chat_window = chat_windows[0]
                # --- FIX: More robust window activation ---
                logging.info(f"Found ChatGPT window: {chat_window.title}. Activating...")
                try:
                    chat_window.activate()
                except Exception as e:
                    # If activate() fails, try a more forceful method
                    logging.warning(f"Standard window activation failed with error: {e}. Trying fallback method.")
                    chat_window.minimize()
                    time.sleep(0.2)
                    chat_window.restore()

                time.sleep(1) # Wait for the window to be active
                logging.info("Typing query into ChatGPT window.")
                pyautogui.write(query, interval=0.03)
                pyautogui.press('enter')
                logging.info("Query sent to ChatGPT.")
            else:
                logging.error("Failed to find ChatGPT window even after attempting to open it.")
                self.speech.speak("I couldn't find or open the ChatGPT window.")

        except ImportError:
            logging.error("A required library (pygetwindow or webbrowser) is not installed.")
            self.speech.speak("A required library for browser interaction is missing.")
        except Exception as e:
            # This will now catch any other unexpected error during the process.
            logging.error(f"An unexpected error occurred in handle_send_to_chatgpt: {e}", exc_info=True)
            self.speech.speak("I had trouble sending your message to ChatGPT.")

    def handle_noop(self, _=None):
        """A handler that does nothing, for commands handled by the main loop."""
        return True
    
    def handle_read_last_note(self, _=None):
        """Reads the last entry from the notes file."""
        try:
            if not os.path.exists(NOTES_FILE):
                return "You haven't taken any notes yet."
            
            with open(NOTES_FILE, 'r', encoding='utf-8') as f:
                notes = f.readlines()
            
            if not notes:
                return "Your notes file is empty."
            
            last_note = notes[-1].strip()
            # The note might start with a timestamp, so we clean it for reading.
            if ' - Note: ' in last_note:
                last_note = last_note.split(' - Note: ', 1)[1]

            return f"Your last note was: {last_note}"
        except Exception as e:
            logging.error(f"Error reading last note: {e}")
            return "I had trouble reading your last note."

    def handle_show_all_notes(self, _=None):
        """Opens the notes file in the default text editor."""
        try:
            if not os.path.exists(NOTES_FILE):
                return "You haven't taken any notes yet. I'll create a new notes file for you."
            
            os.startfile(NOTES_FILE)
            return "Opening your notes."
        except Exception as e:
            logging.error(f"Error opening notes file: {e}")
            return "I couldn't open your notes file."

    def _chunk_text(self, text, llm_handler, max_chunk_tokens=250):
        """Splits text into chunks that are under the token limit."""
        chunks = []
        current_chunk = []
        # Split by paragraphs first, then sentences, to keep context together.
        paragraphs = text.split('\n\n')
        for para in paragraphs:
            sentences = re.split(r'(?<=[.!?])\s+', para)
            for sentence in sentences:
                if not sentence:
                    continue
                # Check if adding the next sentence would exceed the chunk size
                if llm_handler.count_tokens(" ".join(current_chunk + [sentence])) > max_chunk_tokens:
                    # If the current chunk is not empty, finalize it
                    if current_chunk:
                        chunks.append(" ".join(current_chunk))
                        current_chunk = []
                current_chunk.append(sentence)
        
        # Add the last remaining chunk
        if current_chunk:
            chunks.append(" ".join(current_chunk))
            
        return chunks

    def handle_summarize_clipboard(self, _=None):
        """Summarizes the text currently on the clipboard using the LLM."""
        try:
            text_to_summarize = pyperclip.paste()
            if not text_to_summarize or not text_to_summarize.strip():
                return "The clipboard is empty. Please copy some text to summarize."
            
            if not self.command_handler.hybrid_processor.llm_handler:
                return "The Language Model is not available for summarization."
            
            llm_handler = self.command_handler.hybrid_processor.llm_handler
            
            # The prompt itself takes ~70 tokens, and we want ~150 for the output.
            # So, the safe input size is 512 - 70 - 150 = 292. We'll use 250 as a safe limit.
            safe_token_limit = 250
            current_text = text_to_summarize

            # Check if the initial text is long
            if llm_handler.count_tokens(current_text) > safe_token_limit:
                self.speech.speak("The text is long. I will summarize it in parts. This may take a moment.")
            else:
                self.speech.speak("Summarizing the text from your clipboard. Please wait.")

            # --- RECURSIVE REDUCTION LOOP ---
            # Keep summarizing until the text is short enough for a final pass.
            while llm_handler.count_tokens(current_text) > safe_token_limit:
                logging.info(f"Text is too long ({llm_handler.count_tokens(current_text)} tokens). Chunking and reducing...")
                chunks = self._chunk_text(current_text, llm_handler, max_chunk_tokens=safe_token_limit)
                
                intermediate_summaries = []
                for i, chunk in enumerate(chunks):
                    logging.info(f"Summarizing chunk {i+1}/{len(chunks)}...")
                    chunk_summary = llm_handler.generate_summary(chunk)
                    intermediate_summaries.append(chunk_summary)
                # The combined summaries become the input for the next loop iteration
                current_text = "\n".join(intermediate_summaries)
            
            # Perform the final summarization on the now-manageable text
            logging.info("Performing final summarization...")
            summary = llm_handler.generate_summary(current_text)

            print(f"\n--- Summary ---\n{summary}\n-----------------\n")
            return f"Here is the summary: {summary}"

        except Exception as e:
            logging.error(f"Error summarizing clipboard: {e}")
            return "I had trouble summarizing the text from the clipboard."

    def handle_read_most_recent_email(self, cmd_text=None):
        if hasattr(self.file_manager, 'command_handler') and hasattr(self.file_manager.command_handler, 'handle_read_most_recent_email'):
            return self.file_manager.command_handler.handle_read_most_recent_email()
        return "Email reading is not available."

    def handle_read_oldest_email(self, cmd_text=None):
        if hasattr(self.file_manager, 'command_handler') and hasattr(self.file_manager.command_handler, 'handle_read_oldest_email'):
            return self.file_manager.command_handler.handle_read_oldest_email()
        return "Email reading is not available."

    def handle_read_nth_most_recent_email(self, param):
        if hasattr(self.file_manager, 'command_handler') and hasattr(self.file_manager.command_handler, 'handle_read_nth_most_recent_email'):
            return self.file_manager.command_handler.handle_read_nth_most_recent_email(param)
        return "Email reading is not available."

    def handle_read_nth_oldest_email(self, param):
        if hasattr(self.file_manager, 'command_handler') and hasattr(self.file_manager.command_handler, 'handle_read_nth_oldest_email'):
            return self.file_manager.command_handler.handle_read_nth_oldest_email(param)
        return "Email reading is not available."
    

