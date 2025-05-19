import re

class OSCommandHandler:
    def __init__(self, os_manager):
        self.os_manager = os_manager

    def handle_volume_up(self, cmd_text=None):
        """Handle the 'increase volume' command."""
        self.os_manager.volume_up()

    def handle_volume_down(self, cmd_text=None):
        """Handle the 'decrease volume' command."""
        self.os_manager.volume_down()

    def handle_mute_toggle(self, cmd_text=None):
        """Handle the 'mute/unmute volume' command."""
        self.os_manager.mute_toggle()

    def handle_maximize_volume(self, cmd_text=None):
        """Handle the 'maximize volume' command."""
        self.os_manager.maximize_volume()

    def handle_set_volume(self, cmd_text):
        """Handle the 'set volume' command."""
        try:
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            level = int(match.group(1)) if match else None
            if level is not None:
                self.os_manager.set_volume(level)
            else:
                print("No valid volume level found. Please say a number between 0 and 100.")
                self.os_manager.speech.speak("No valid volume level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting volume: {e}")
            self.os_manager.speech.speak("Error setting volume.")

    def handle_brightness_up(self, cmd_text=None):
        """Handle the 'increase brightness' command."""
        self.os_manager.brightness_up()

    def handle_brightness_down(self, cmd_text=None):
        """Handle the 'decrease brightness' command."""
        self.os_manager.brightness_down()

    def handle_maximize_brightness(self, cmd_text=None):
        """Handle the 'maximize brightness' command."""
        self.os_manager.maximize_brightness()

    def handle_set_brightness(self, cmd_text):
        """Handle the 'set brightness' command."""
        try:
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            level = int(match.group(1)) if match else None
            if level is not None:
                self.os_manager.set_brightness(level)
            else:
                print("No valid brightness level found. Please say a number between 0 and 100.")
                self.os_manager.speech.speak("No valid brightness level found. Please say a number between 0 and 100.")
        except Exception as e:
            print(f"Error setting brightness: {e}")
            self.os_manager.speech.speak("Error setting brightness.")

    def handle_switch_window(self, cmd_text=None):
        """Handle the 'switch window' command."""
        self.os_manager.switch_window()

    def handle_minimize_all_windows(self, cmd_text=None):
        """Handle the 'minimize all windows' command."""
        self.os_manager.minimize_all_windows()

    def handle_restore_all_windows(self, cmd_text=None):
        """Handle the 'restore windows' command."""
        self.os_manager.restore_all_windows()

    def handle_maximize_current_window(self, cmd_text=None):
        """Handle the 'maximize window' command."""
        self.os_manager.maximize_current_window()

    def handle_minimize_current_window(self, cmd_text=None):
        """Handle the 'minimize window' command."""
        self.os_manager.minimize_current_window()

    def handle_close_current_window(self, cmd_text=None):
        """Handle the 'close window' command."""
        self.os_manager.close_current_window()

    def handle_move_window_left(self, cmd_text=None):
        """Handle the 'move window left' command."""
        self.os_manager.move_window_left()

    def handle_move_window_right(self, cmd_text=None):
        """Handle the 'move window right' command."""
        self.os_manager.move_window_right()

    def handle_take_screenshot(self, cmd_text=None):
        """Handle the 'take screenshot' command."""
        self.os_manager.take_screenshot()

    def handle_run_application(self, app_name):
        """Handle the 'run application' command."""
        if app_name:
            self.os_manager.run_application(app_name)
        else:
            print("No application name provided. Please say the name of the application to run.")
            self.os_manager.speech.speak("No application name provided. Please say the name of the application to run.")