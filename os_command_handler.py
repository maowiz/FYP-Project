import re
import os
import time
import threading
import tkinter as tk
import ctypes
import win32con
import webbrowser
import subprocess
import pyperclip
import sys
from browser_commands import (
    previous_tab, next_tab, close_tab, refresh, zoom_in, zoom_out,
    bookmark_tab, open_incognito, switch_tab, search, clear_browsing_data
)


class OSCommandHandler:
    def __init__(self, os_manager):
        self.os_manager = os_manager
        self._scrolling = False
        self._scroll_thread = None
        if not hasattr(self.os_manager, 'context'):
            self.os_manager.context = {}
    
    def is_scrolling(self):
        """Check if scrolling is currently active"""
        return self._scrolling and self._scroll_thread and self._scroll_thread.is_alive()

    def handle_volume_up(self, cmd_text=None):
        """Handle the 'increase volume' command."""
        success, message = self.os_manager.volume_up()
        return message

    def handle_volume_down(self, cmd_text=None):
        """Handle the 'decrease volume' command."""
        success, message = self.os_manager.volume_down()
        return message

    def handle_mute_toggle(self, cmd_text=None):
        """Handle the 'mute/unmute volume' command."""
        success, message = self.os_manager.mute_toggle()
        return message

    def handle_maximize_volume(self, cmd_text=None):
        """Handle the 'maximize volume' command."""
        success, message = self.os_manager.maximize_volume()
        return message

    def handle_set_volume(self, cmd_text):
        """Handle the 'set volume' command."""
        try:
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            level = int(match.group(1)) if match else None
            if level is not None:
                success, message = self.os_manager.set_volume(str(level))
                return message
            else:
                return "No valid volume level found. Please say a number between 0 and 100."
        except Exception as e:
            print(f"Error setting volume: {e}")
            return "Error setting volume."

    def handle_brightness_up(self, cmd_text=None):
        """Handle the 'increase brightness' command."""
        success, message = self.os_manager.brightness_up()
        return message

    def handle_brightness_down(self, cmd_text=None):
        """Handle the 'decrease brightness' command."""
        success, message = self.os_manager.brightness_down()
        return message

    def handle_maximize_brightness(self, cmd_text=None):
        """Handle the 'maximize brightness' command."""
        success, message = self.os_manager.maximize_brightness()
        return message

    def handle_set_brightness(self, cmd_text):
        """Handle the 'set brightness' command."""
        try:
            match = re.search(r'\b(\d{1,3})\b', cmd_text)
            level = int(match.group(1)) if match else None
            if level is not None:
                success, message = self.os_manager.set_brightness(str(level))
                return message
            else:
                return "No valid brightness level found. Please say a number between 0 and 100."
        except Exception as e:
            print(f"Error setting brightness: {e}")
            return "Error setting brightness."

    def handle_switch_window(self, cmd_text=None):
        """Handle the 'switch window' command."""
        success, message = self.os_manager.switch_window()
        return message

    # ----- Aura Grid handlers -----
    def handle_show_grid(self, cmd_text=None):
        # Always show default grid; no variants required
        success, message = self.os_manager.grid.show_grid(density=None, pinned=False)
        return message

    def handle_hide_grid(self, cmd_text=None):
        success, message = self.os_manager.grid.hide_grid()
        return message

    def handle_click_cell(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid cell number."
        success, message = self.os_manager.grid.click_cell(n, button="left")
        return message

    def handle_double_click_cell(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid cell number."
        success, message = self.os_manager.grid.double_click_cell(n)
        return message

    def handle_right_click_cell(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid cell number."
        success, message = self.os_manager.grid.click_cell(n, button="right")
        return message

    def handle_drag_from(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid start cell number."
        success, message = self.os_manager.grid.start_drag(n)
        return message

    def handle_drop_on(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid target cell number."
        success, message = self.os_manager.grid.drop_on(n)
        return message

    def handle_zoom_cell(self, cmd_text):
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a valid zoom cell number."
        success, message = self.os_manager.grid.zoom_cell(n)
        return message

    def handle_exit_zoom(self, cmd_text=None):
        success, message = self.os_manager.grid.exit_zoom()
        return message

    def handle_set_grid_size(self, cmd_text):
        """Handle commands like 'zoom 15' or 'grid 10' to set N x N grid divisions."""
        n = self._extract_number(cmd_text)
        if n is None:
            return "Please say a grid number, for example zoom 15."
        success, message = self.os_manager.grid.set_grid_divisions(n)
        return message

    def _extract_number(self, text):
        """
        Extract numbers from text with proper compound number handling.
        FIXED: Now prioritizes "twenty five" -> 25 instead of extracting just "five" -> 5
        """
        if not text:
            return None
        
        # 1. Check for digits first (highest priority)
        m = re.search(r"\b(\d{1,4})\b", text)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                return None
        
        # 2. Word-to-number mapping
        words = {
            'zero':0, 'one':1, 'two':2, 'three':3, 'four':4, 'five':5, 
            'six':6, 'seven':7, 'eight':8, 'nine':9, 'ten':10,
            'eleven':11, 'twelve':12, 'thirteen':13, 'fourteen':14, 'fifteen':15,
            'sixteen':16, 'seventeen':17, 'eighteen':18, 'nineteen':19,
            'twenty':20, 'thirty':30, 'forty':40, 'fifty':50,
            'sixty':60, 'seventy':70, 'eighty':80, 'ninety':90
        }
        
        text_l = text.lower()
        
        # 3. CRITICAL FIX: Try compound numbers FIRST (e.g., "twenty five" = 25)
        # This must come before single word extraction to prevent "twenty five" -> 5
        compound_pattern = r'\b(twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)[\s-]+(one|two|three|four|five|six|seven|eight|nine)\b'
        m_compound = re.search(compound_pattern, text_l)
        if m_compound:
            result = words[m_compound.group(1)] + words[m_compound.group(2)]
            print(f"Extracted number (compound): {result} from '{text}'")
            return result
        
        # 4. Try single word numbers (but search larger numbers FIRST to avoid partial matches)
        # This prevents "twenty" from being found in "twenty five" after compound check fails
        for w, val in sorted(words.items(), key=lambda x: -x[1]):  # Sort by value descending
            if re.search(fr'\b{w}\b', text_l):
                print(f"Extracted number (word): {val} from '{text}'")
                return val
        
        return None


    def handle_minimize_all_windows(self, cmd_text=None):
        """Handle the 'minimize all windows' command."""
        success, message = self.os_manager.minimize_all_windows()
        return message

    def handle_restore_all_windows(self, cmd_text=None):
        """Handle the 'restore windows' command."""
        success, message = self.os_manager.restore_all_windows()
        return message

    def handle_maximize_current_window(self, cmd_text=None):
        """Handle the 'maximize window' command."""
        success, message = self.os_manager.maximize_current_window()
        return message

    def handle_minimize_current_window(self, cmd_text=None):
        """Handle the 'minimize window' command."""
        success, message = self.os_manager.minimize_current_window()
        return message

    def handle_close_current_window(self, cmd_text=None):
        """Handle the 'close window' command."""
        success, message = self.os_manager.close_current_window()
        return message

    def handle_move_window_left(self, cmd_text=None):
        """Handle the 'move window left' command."""
        success, message = self.os_manager.move_window_left()
        return message

    def handle_move_window_right(self, cmd_text=None):
        """Handle the 'move window right' command."""
        success, message = self.os_manager.move_window_right()
        return message

    def handle_take_screenshot(self, cmd_text=None):
        """Handle the 'take screenshot' command."""
        success, message = self.os_manager.take_screenshot()
        return message

    def handle_run_application(self, app_name):
        """Handle the 'run application' command."""
        if app_name:
            success, message = self.os_manager.run_application(app_name)
            return message
        else:
            print("No application name provided. Please say the name of the application to run.")
            return "No application name provided. Please say the name of the application to run."

    def handle_go_to_desktop(self, _=None):
        """Press Win + D to show desktop and stay there."""
        try:
            import pyautogui
            pyautogui.hotkey('win', 'd')  # Only once!
            print("Pressed Win+D → Desktop shown")
            return "Desktop is now showing."
        except Exception as e:
            print(f"Error showing desktop: {e}")
            return "Failed to show desktop."

    def handle_change_wallpaper(self, _=None):
        """Switch to next wallpaper using reliable COM method."""
        success, message = self.os_manager.next_wallpaper()
        return message

    def handle_empty_recycle_bin(self, _=None):
        """Empty the Recycle Bin silently."""
        try:
            # Use SHERB_NOCONFIRMATION flag with ctypes
            SHERB_NOCONFIRMATION = 0x00000001
            result = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, SHERB_NOCONFIRMATION)
            if result == 0:
                return "Recycle bin has been emptied."
            else:
                return "Failed to empty recycle bin."
        except Exception as e:
            print("Empty recycle-bin error:", e)
            return "Error emptying recycle bin."

    def handle_scroll_up(self, _=None):
        return self._start_scrolling(direction='up')

    def handle_scroll_down(self, _=None):
        return self._start_scrolling(direction='down')

    def handle_scroll_left(self, _=None):
        return self._start_scrolling(direction='left')

    def handle_scroll_right(self, _=None):
        return self._start_scrolling(direction='right')

    def handle_stop_scrolling(self, _=None):
        print("handle_stop_scrolling called")  # Debug print
        if self._scrolling:
            self._scrolling = False
            # Wait for the scroll thread to finish
            if self._scroll_thread and self._scroll_thread.is_alive():
                self._scroll_thread.join(timeout=2)
            print("Scroll thread stopped")  # Debug print
            return "Scrolling has been stopped."
        else:
            print("No scrolling was active")  # Debug print
            return "No scrolling is currently active."

    def _start_scrolling(self, direction):
        import pyautogui
        if self._scrolling:
            self._scrolling = False
            if self._scroll_thread:
                self._scroll_thread.join(timeout=1)
        self._scrolling = True
        def scroll_loop():
            print(f"Starting to scroll {direction}")  # Debug print
            # Don't speak immediately to avoid interrupting the scroll
            time.sleep(0.5)  # Small delay before starting
            while self._scrolling:
                try:
                    if direction == 'up':
                        pyautogui.scroll(100)
                    elif direction == 'down':
                        pyautogui.scroll(-100)
                    elif direction == 'left':
                        pyautogui.keyDown('shift')
                        pyautogui.scroll(-100)
                        pyautogui.keyUp('shift')
                    elif direction == 'right':
                        pyautogui.keyDown('shift')
                        pyautogui.scroll(100)
                        pyautogui.keyUp('shift')
                    else:
                        print(f"Unknown scroll direction: {direction}")
                        break
                    time.sleep(0.1)
                except Exception as e:
                    print(f"Error in scroll loop: {e}")
                    break
            print(f"Stopped scrolling {direction}")  # Debug print
        self._scroll_thread = threading.Thread(target=scroll_loop, daemon=True)
        self._scroll_thread.start()
        # Speak after starting the thread
        return f"Scrolling {direction}."

    def handle_previous_tab(self, _):  return "Previous tab" if previous_tab() else "Failed to switch tab"
    def handle_next_tab(self, _):      return "Next tab" if next_tab() else "Failed to switch tab"
    def handle_close_tab(self, _):     return "Tab closed" if close_tab() else "Failed to close tab"
    def handle_refresh(self, _):       return "Refreshed" if refresh() else "Failed to refresh"
    def handle_zoom_in(self, _):       return "Zoomed in" if zoom_in() else "Failed to zoom in"
    def handle_zoom_out(self, _):      return "Zoomed out" if zoom_out() else "Failed to zoom out"
    def handle_bookmark_tab(self, _):  return "Tab bookmarked" if bookmark_tab() else "Failed to bookmark tab"
    def handle_open_incognito(self, _): return "Incognito window opened" if open_incognito() else "Failed to open incognito"
    def handle_switch_tab(self, n):
        if n is None:
            # No number provided, switch to next tab
            return "Next tab" if next_tab() else "Failed to switch tab"
        else:
            # Number provided, switch to specific tab
            return f"Switched to tab {n}" if switch_tab(int(n)) else f"Failed to switch to tab {n}"
            
    def handle_search(self, q):
        result = search(q)
        if result:
            return f"Searching for {q}"
        else:
            return f"Failed to search for {q}"
            
    def handle_clear_browsing_data(self, _): return "Browsing data clearing dialog opened" if clear_browsing_data() else "Failed to open clearing dialog"

    def handle_open_generic(self, open_target):
        """Open a website or application by name (e.g., 'open google', 'open youtube')."""
        if not open_target:
            return "Please specify what you want to open."
            
        # Map common names to URLs
        url_map = {
            'google': 'https://www.google.com',
            'youtube': 'https://www.youtube.com',
            'gmail': 'https://mail.google.com',
            'facebook': 'https://www.facebook.com',
            'twitter': 'https://twitter.com',
            'github': 'https://github.com',
            'reddit': 'https://www.reddit.com',
            'whatsapp': 'https://web.whatsapp.com',
            'instagram': 'https://www.instagram.com',
            'amazon': 'https://www.amazon.com',
            'netflix': 'https://www.netflix.com',
            'stackoverflow': 'https://stackoverflow.com',
            'bing': 'https://www.bing.com',
            'yahoo': 'https://www.yahoo.com',
            'wikipedia': 'https://www.wikipedia.org',
        }
        
        key = open_target.lower().strip()
        
        if key in url_map:
            url = url_map[key]
            try:
                # Try to open with webbrowser first
                success = webbrowser.open(url)
                if not success:
                    # Fallback: try to open with system command
                    if sys.platform.startswith('win'):
                        subprocess.run(['start', url], shell=True)
                    elif sys.platform.startswith('darwin'):  # macOS
                        subprocess.run(['open', url])
                    else:  # Linux
                        subprocess.run(['xdg-open', url])
                
                print(f"Attempted to open: {url}")
                
                # Set context for YouTube
                if hasattr(self.os_manager, 'context'):
                    self.os_manager.context['youtube_open'] = (key == 'youtube')
                
                return f"Opening {key}."
                    
            except Exception as e:
                print(f"Error opening {key}: {e}")
                return f"Error opening {key}. Please check your browser settings."
                
        elif key.startswith('http') or key.startswith('www.'):
            try:
                webbrowser.open(open_target)
                return f"Opening {open_target}."
            except Exception as e:
                print(f"Error opening URL {open_target}: {e}")
                return f"Error opening {open_target}."
        else:
            # Try to open as an application
            try:
                success, message = self.os_manager.run_application(open_target)
                if hasattr(self.os_manager, 'context'):
                    self.os_manager.context['youtube_open'] = False
                return message
            except Exception as e:
                print(f"Error opening application {open_target}: {e}")
                return f"Error opening {open_target} application."

    def handle_play_on_youtube(self, query):
        """Handle the 'play on youtube' command by searching and playing music/video on YouTube."""
        if not query:
            print("No query provided for YouTube search.")
            return "Please specify what you want to play on YouTube."
        # Construct YouTube search URL
        search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
        webbrowser.open(search_url)
        print(f"Opened YouTube search for: {query}")
        return f"Searching YouTube for {query}."

    def handle_copy(self, cmd_text=None):
        """Handle the 'copy' command by sending Ctrl+C."""
        try:
            import pyautogui
            pyautogui.hotkey('ctrl', 'c')
            return "Content has been copied to clipboard."
        except Exception as e:
            print(f"Error copying: {e}")
            return "Failed to copy content."

    def handle_paste(self, cmd_text=None):
        """Handle the 'paste' command by sending Ctrl+V."""
        try:
            import pyautogui
            pyautogui.hotkey('ctrl', 'v')
            return "Content has been pasted."
        except Exception as e:
            print(f"Error pasting: {e}")
            return "Failed to paste content."

    def handle_read_clipboard(self, cmd_text=None):
        """Reads the current content of the clipboard."""
        try:
            content = pyperclip.paste()
            if content:
                # To prevent reading out very long text, we'll truncate it for speech.
                spoken_content = (content[:150] + '...') if len(content) > 150 else content
                print(f"Clipboard content: {content}")
                return f"The clipboard says: {spoken_content}"
            else:
                return "The clipboard is empty."
        except Exception as e:
            print(f"Error reading clipboard: {e}")
            return "Failed to read clipboard."

    def handle_select_all(self, cmd_text=None):
        """Handle the 'select all' command by sending Ctrl+A."""
        try:
            import pyautogui
            pyautogui.hotkey('ctrl', 'a')
            return "All content has been selected."
        except Exception as e:
            print(f"Error selecting all: {e}")
            return "Failed to select all content."

    def handle_open_word(self, cmd_text=None):
        """Handle the 'open word' command."""
        try:
            import pyautogui
            import time
            import pygetwindow as gw

            success, message = self.os_manager.run_application("word")
            time.sleep(3) # Wait for Word to launch
            word_windows = gw.getWindowsWithTitle('Word')
            if word_windows:
                word_window = word_windows[0]
                word_window.maximize()
                return "Microsoft Word has been opened and maximized."
            # The run_application method already provides feedback
            return message
        except Exception as e:
            print(f"Error opening Word: {e}")
            return "Failed to open Microsoft Word."

    def handle_save_file(self, filename=None):
        """Saves the current active file to the desktop with a given name."""
        try:
            import pyautogui
            import time
            # If no filename is provided, use the context from the essay command or a default.
            if not filename:
                topic = self.os_manager.context.get("last_essay_topic", "document")
                # Sanitize topic to be a valid filename
                filename = "".join([c for c in topic if c.isalpha() or c.isdigit() or c==' ']).rstrip()
                filename = f"{filename.replace(' ', '_')}.txt"

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            full_path = os.path.join(desktop_path, filename)

            pyautogui.hotkey('ctrl', 's')
            time.sleep(1) # Wait for the save dialog to appear
            pyautogui.write(full_path, interval=0.05)
            time.sleep(0.5)
            pyautogui.press('enter')
            
            # --- NEW: Open the desktop to show the saved file ---
            time.sleep(1) # Wait a moment before showing the desktop
            self.handle_go_to_desktop()
            return f"File has been saved as {filename} on your desktop."
        except Exception as e:
            print(f"Error saving file: {e}")
            return "Failed to save the file."

    def handle_remove_selection(self, cmd_text=None):
        """Handle the 'remove this' command by pressing the Delete key."""
        try:
            import pyautogui
            pyautogui.press('delete')
            return "Selection has been removed."
        except Exception as e:
            print(f"Error removing selection: {e}")
            return "Failed to remove selection."

    def handle_undo_action(self, cmd_text=None):
        """Handle the 'undo' command by sending Ctrl+Z."""
        try:
            import pyautogui
            pyautogui.hotkey('ctrl', 'z')
            return "Action has been undone."
        except Exception as e:
            print(f"Error performing undo: {e}")
            return "Failed to undo action."
    
    def handle_redo_action(self, cmd_text=None):
        """Handle the 'redo' command by sending Ctrl+Y."""
        try:
            import pyautogui
            pyautogui.hotkey('ctrl', 'y')
            return "Action has been redone."
        except Exception as e:
            print(f"Error performing redo: {e}")
            return "Failed to redo action."