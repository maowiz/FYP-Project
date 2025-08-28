import sys
import requests
import speedtest
import cv2
import os
from datetime import datetime
import ctypes
import time
import threading
import tkinter as tk

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
    
    

