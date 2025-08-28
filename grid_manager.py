import threading
import queue
import time
from typing import Dict, Tuple, Optional, List

import pyautogui

try:
    import tkinter as tk
except Exception:  # Headless safeguard
    tk = None

try:
    import win32api
    import win32con
    import win32gui
except Exception:
    win32api = None
    win32con = None
    win32gui = None


CellIndex = int
Point = Tuple[int, int]
Rect = Tuple[int, int, int, int]


class GridManager:
    """Translucent, topmost grid overlay across monitors with click/drag helpers."""

    def __init__(self, speech, default_density: str = "medium"):
        self.speech = speech
        self.default_density = default_density

        # UI state
        self._tk_thread: Optional[threading.Thread] = None
        self._root: Optional[tk.Tk] = None
        self._canvas: Optional[tk.Canvas] = None
        self._cmd_q: "queue.Queue[Tuple[str, tuple, dict]]" = queue.Queue()
        self._visible: bool = False
        self._pinned: bool = False
        self._zoom_stack: List[Rect] = []

        # Geometry
        self._virtual_rect: Rect = (0, 0, pyautogui.size().width, pyautogui.size().height)
        self._cell_centers: Dict[CellIndex, Point] = {}
        self._cell_rects: Dict[CellIndex, Rect] = {}
        self._cell_size_px: int = 100

        # External callbacks (set by main): pause/resume listening
        self._pause_cb = None
        self._resume_cb = None

        # Drag state
        self._drag_start_cell: Optional[CellIndex] = None

    def set_pause_resume(self, pause_cb, resume_cb) -> None:
        self._pause_cb = pause_cb
        self._resume_cb = resume_cb

    # ---------- Public API ----------
    def show_grid(self, density: Optional[str] = None, pinned: bool = False, on_window_rect: Optional[Rect] = None) -> None:
        self._pinned = pinned
        if density:
            self._cell_size_px = self._density_to_cell_size(density)
        else:
            # Default to ~15 divisions across the smaller screen axis
            rect = self._zoom_stack[-1] if self._zoom_stack else self._get_virtual_screen_rect()
            vx, vy, vr, vb = rect
            width = max(1, vr - vx)
            height = max(1, vb - vy)
            step = max(30, min(width // 15, height // 15))
            self._cell_size_px = int(step)
        if on_window_rect:
            self._zoom_stack = [on_window_rect]
        else:
            self._zoom_stack = []

        self._ensure_thread()
        self._enqueue("_cmd_show", {})
        if self._pause_cb:
            self._pause_cb()

    def hide_grid(self) -> None:
        self._enqueue("_cmd_hide", {})

    def set_grid_divisions(self, divisions: int) -> None:
        """Dynamically set grid to N x N divisions over the current view.

        Chooses a cell size so that both width and height are divided into
        approximately 'divisions' steps, using the smaller axis to keep cells square.
        """
        try:
            if divisions is None or not isinstance(divisions, int):
                self.speech.speak("Please say a valid grid number")
                return
            # Clamp to sensible range
            divisions = max(2, min(50, divisions))
            draw_rect = self._zoom_stack[-1] if self._zoom_stack else self._get_virtual_screen_rect()
            vx, vy, vr, vb = draw_rect
            width = max(1, vr - vx)
            height = max(1, vb - vy)
            # Choose size so we get ~divisions across the smaller dimension
            step = max(30, min(width // divisions, height // divisions))
            self._cell_size_px = int(step)
            self._enqueue("_cmd_redraw", {})
        except Exception:
            self.speech.speak("Could not set grid size")

    def click_cell(self, cell: CellIndex, button: str = "left") -> bool:
        def _do_click():
            if button == "right":
                pyautogui.click(button="right")
            else:
                pyautogui.click(button="left" if button == "left" else button)
        return self._perform_mouse_action(cell, _do_click, keep_mouse_at_target=True)

    def double_click_cell(self, cell: CellIndex) -> bool:
        def _do_double():
            pyautogui.doubleClick(interval=0.1)
        return self._perform_mouse_action(cell, _do_double, keep_mouse_at_target=True)

    def _perform_mouse_action(self, cell: CellIndex, action_func, keep_mouse_at_target: bool = False) -> bool:
        """Perform a mouse action at the specified cell. If grid is not pinned, briefly hide it to guarantee click-through."""
        pt = self._cell_centers.get(cell)
        if not pt:
            print(f"Unknown cell number: {cell}")
            return False

        try:
            # If overlay is visible and not pinned, briefly hide to guarantee OS click-through
            should_temp_hide = self._visible and (not self._pinned)
            if should_temp_hide:
                self.hide_grid()
                time.sleep(0.12)

            # Move to cell center and perform action
            pyautogui.moveTo(pt[0], pt[1])
            print(f"Clicking at cell {cell} coordinates: {pt}")
            action_func()

            # Give visual feedback only if still visible
            if not should_temp_hide:
                self._flash_cell(cell, color="#00ff00")

            print(f"Successfully clicked cell {cell}")
            return True
        except Exception as e:
            print(f"Mouse action error at cell {cell}: {e}")
            return False

    def start_drag(self, cell: CellIndex) -> bool:
        pt = self._cell_centers.get(cell)
        if not pt:
            self.speech.speak("Unknown cell number")
            return False
        try:
            pyautogui.moveTo(pt[0], pt[1])
            pyautogui.mouseDown()
            self._drag_start_cell = cell
            self._flash_cell(cell, color="#ffc107")
            return True
        except Exception:
            self.speech.speak("Drag start failed")
            return False

    def drop_on(self, cell: CellIndex, duration: float = 0.4) -> bool:
        if self._drag_start_cell is None:
            self.speech.speak("No drag in progress")
            return False
        pt = self._cell_centers.get(cell)
        if not pt:
            self.speech.speak("Unknown target cell")
            return False
        try:
            pyautogui.dragTo(pt[0], pt[1], duration=duration)
            pyautogui.mouseUp()
            self._drag_start_cell = None
            return True
        except Exception:
            self.speech.speak("Drag failed")
            return False

    def zoom_cell(self, cell: CellIndex) -> bool:
        rect = self._cell_rects.get(cell)
        if not rect:
            self.speech.speak("Unknown cell number")
            return False
        self._zoom_stack.append(rect)
        self._enqueue("_cmd_redraw", {})
        return True

    def exit_zoom(self) -> bool:
        if not self._zoom_stack:
            self.speech.speak("Not zoomed")
            return False
        self._zoom_stack.pop()
        self._enqueue("_cmd_redraw", {})
        return True

    # ---------- Internal UI Thread ----------
    def _ensure_thread(self) -> None:
        if self._tk_thread and self._tk_thread.is_alive():
            return
        if tk is None:
            self.speech.speak("Grid overlay not available")
            return

        self._tk_thread = threading.Thread(target=self._run_tk, daemon=True)
        self._tk_thread.start()
        # Allow thread to initialize
        time.sleep(0.05)

    def _run_tk(self) -> None:
        self._root = tk.Tk()
        self._root.withdraw()
        self._root.overrideredirect(True)
        self._root.attributes("-topmost", True)
        # Configure layered window + click-through; also set transparent colorkey
        try:
            if win32gui is not None and win32con is not None:
                hwnd = self._root.winfo_id()
                ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                ex_style |= win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT
                win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
                # Make pure black fully transparent so only lines/text are visible
                try:
                    if win32api is not None:
                        win32gui.SetLayeredWindowAttributes(
                            hwnd,
                            win32api.RGB(0, 0, 0),  # transparent color key
                            255,                    # full opacity for non-key colors
                            win32con.LWA_COLORKEY | win32con.LWA_ALPHA
                        )
                except Exception:
                    pass
        except Exception:
            pass

        # Cover virtual screen
        self._virtual_rect = self._get_virtual_screen_rect()
        vx, vy, vr, vb = self._virtual_rect
        width = vr - vx
        height = vb - vy
        # Geometry with correct sign formatting for negative coordinates
        geom = f"{width}x{height}"
        geom += (f"+{vx}" if vx >= 0 else f"{vx}")
        geom += (f"+{vy}" if vy >= 0 else f"{vy}")
        self._root.geometry(geom)
        self._root.update_idletasks()
        # Force exact position on virtual desktop using Win32 to support negative coords
        try:
            if win32gui is not None and win32con is not None:
                hwnd = self._root.winfo_id()
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, vx, vy, width, height, win32con.SWP_NOACTIVATE)
        except Exception:
            pass
        # Use solid background; window alpha controls transparency
        self._canvas = tk.Canvas(self._root, width=width, height=height, highlightthickness=0, bg="#000000")
        self._canvas.pack(fill=tk.BOTH, expand=True)

        # Poll command queue
        self._root.after(10, self._process_queue)
        self._root.mainloop()

    def _process_queue(self):
        try:
            while True:
                name, args, kwargs = self._cmd_q.get_nowait()
                getattr(self, name)(*args, **kwargs)
        except queue.Empty:
            pass
        if self._root:
            self._root.after(16, self._process_queue)

    def _enqueue(self, name: str, kwargs: dict) -> None:
        self._cmd_q.put((name, tuple(), kwargs))

    # ---------- UI commands executed on Tk thread ----------
    def _cmd_show(self):
        if not self._root:
            return
        self._root.deiconify()
        self._visible = True
        self._redraw_grid()

    def _cmd_hide(self):
        if not self._root:
            return
        self._visible = False
        self._root.withdraw()
        if self._resume_cb:
            # Small delay so ASR does not capture tail events
            time.sleep(0.15)
            self._resume_cb()

    def _cmd_redraw(self):
        self._redraw_grid()

    # ---------- Drawing ----------
    def _redraw_grid(self):
        if not self._canvas:
            return
        self._canvas.delete("all")
        # Compute drawing rect
        draw_rect = self._zoom_stack[-1] if self._zoom_stack else self._get_virtual_screen_rect()
        vx, vy, vr, vb = draw_rect
        width = vr - vx
        height = vb - vy
        # Background veil
        self._canvas.create_rectangle(0, 0, self._canvas.winfo_width(), self._canvas.winfo_height(), fill="#000000", outline="")

        self._cell_centers.clear()
        self._cell_rects.clear()

        cell = 1
        step = self._cell_size_px
        for y in range(vy, vb, step):
            for x in range(vx, vr, step):
                x2 = min(x + step, vr)
                y2 = min(y + step, vb)
                cx = x + (x2 - x) // 2
                cy = y + (y2 - y) // 2
                # Translate to canvas coordinates
                gx = x - self._virtual_rect[0]
                gy = y - self._virtual_rect[1]
                gx2 = x2 - self._virtual_rect[0]
                gy2 = y2 - self._virtual_rect[1]
                self._canvas.create_rectangle(gx, gy, gx2, gy2, outline="#00ffff", width=2)
                if step <= 150 or (self._zoom_stack):
                    self._canvas.create_text(gx + 8, gy + 8, anchor="nw", text=str(cell), fill="#00ffaa", font=("Segoe UI", 10))
                self._cell_centers[cell] = (cx, cy)
                self._cell_rects[cell] = (x, y, x2, y2)
                cell += 1

    def _flash_cell(self, cell: CellIndex, color: str = "#28a745"):
        if not self._visible or not self._canvas:
            return
        rect = self._cell_rects.get(cell)
        if not rect:
            return
        x, y, x2, y2 = rect
        gx = x - self._virtual_rect[0]
        gy = y - self._virtual_rect[1]
        gx2 = x2 - self._virtual_rect[0]
        gy2 = y2 - self._virtual_rect[1]
        item = None
        try:
            item = self._canvas.create_rectangle(gx, gy, gx2, gy2, outline=color, width=2)
            self._canvas.update_idletasks()
            time.sleep(0.1)
        finally:
            if item:
                self._canvas.delete(item)

    def _auto_hide_if_needed(self):
        if not self._pinned:
            self.hide_grid()

    # ---------- Helpers ----------
    def _density_to_cell_size(self, density: str) -> int:
        density = (density or "").lower()
        if density.startswith("coarse"):
            return 180
        if density.startswith("fine"):
            return 75
        return 120

    def _get_virtual_screen_rect(self) -> Rect:
        if win32api is None:
            size = pyautogui.size()
            return (0, 0, size.width, size.height)
        left = 10**9
        top = 10**9
        right = -10**9
        bottom = -10**9
        for m in win32api.EnumDisplayMonitors():
            monitor_info = win32api.GetMonitorInfo(m[0])
            (l, t, r, b) = monitor_info['Monitor']
            left = min(left, l)
            top = min(top, t)
            right = max(right, r)
            bottom = max(bottom, b)
        self._virtual_rect = (left, top, right, bottom)
        return self._virtual_rect
