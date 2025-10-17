# vm_gpt11_fist.py — Virtual Mouse (with robust Fist/Punch handoff)

import cv2
import mediapipe as mp
import pyautogui
import time
import math
import platform
from enum import Enum

# Fast + safe pyautogui
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0

DEBUG = False  # flip to True for occasional prints


class GestureType(Enum):
    NONE = "none"
    MOVE = "move"
    LEFT_CLICK = "left_click"
    DOUBLE_CLICK = "double_click"
    RIGHT_CLICK = "right_click"
    DRAG_START = "drag_start"
    DRAGGING = "dragging"
    DRAG_END = "drag_end"
    SCROLLING = "scrolling"
    OSK_TOGGLE = "osk_toggle"  # Pinky-only: toggle Windows On-Screen Keyboard
    HANDOFF = "handoff"        # Fist/Punch → exit to voice assistant


class GestureState(Enum):
    IDLE = "idle"
    MOVING = "moving"
    PINCHING = "pinching"
    DRAGGING = "dragging"
    SCROLLING = "scrolling"


class PinchFinger(Enum):
    NONE = 0
    INDEX = 1
    MIDDLE = 2


# ---------- One Euro Filter (v2 params + perf_counter) ----------
class OneEuroFilter:
    def __init__(self, te=1.0/60, min_cutoff=2.5, beta=0.03, d_cutoff=1.0):
        self.te = te
        self.min_cutoff = min_cutoff
        self.beta = beta
        self.d_cutoff = d_cutoff
        self.x_prev = None
        self.dx_prev = 0.0
        self.t_prev = None

    def smoothing_factor(self, cutoff, te):
        tau = 1.0 / (2 * math.pi * cutoff)
        return 1.0 / (1.0 + tau / te)

    def exponential_smoothing(self, x, x_prev, a):
        return a * x + (1 - a) * x_prev

    def filter(self, x, t=None):
        if t is None:
            t = time.perf_counter()
        if self.x_prev is None:
            self.x_prev = x
            self.t_prev = t
            return x
        te = t - (self.t_prev or t)
        if te <= 0:
            return self.x_prev

        dx = (x - self.x_prev) / te
        a_d = self.smoothing_factor(self.d_cutoff, te)
        dx_hat = self.exponential_smoothing(dx, self.dx_prev, a_d)

        cutoff = self.min_cutoff + self.beta * abs(dx_hat)
        a = self.smoothing_factor(cutoff, te)
        x_hat = self.exponential_smoothing(x, self.x_prev, a)

        self.x_prev = x_hat
        self.dx_prev = dx_hat
        self.t_prev = t
        return x_hat


# ---------- Strict Gesture Recognizer (index–thumb = LC, middle–thumb = RC, pinky-only = OSK) ----------
class StrictGestureRecognizer:
    def __init__(self, frame_w: int, frame_h: int):
        self.frame_w = frame_w
        self.frame_h = frame_h

        # Pinch thresholds (pixels) — shared by both index and middle pinch
        self.PINCH_ENTER_THRESHOLD_PX = 30
        self.PINCH_EXIT_THRESHOLD_PX = 40

        # Click timing
        self.CLICK_MAX_DURATION = 1.0
        self.DOUBLE_CLICK_WINDOW = 0.8

        # Scroll gating (harder to start so slow pinches don't scroll)
        self.SCROLL_HOLD_THRESHOLD = 1.2       # must hold pinch this long before scroll considered
        self.SCROLL_MOVEMENT_THRESHOLD = 0.035 # need real vertical movement (normalized)
        self.SCROLL_DEADZONE = 0.02            # ignore tiny jitter (normalized)
        self.SCROLL_ARM_FRAMES = 3             # require consecutive frames of movement
        self._scroll_arm_count = 0

        # Click forgiveness (slow pinch with tiny movement still counts as click)
        self.CLICK_FORGIVENESS_DURATION = 1.6
        self.CLICK_FORGIVENESS_MOVEMENT = 0.025

        # State
        self.current_state = GestureState.IDLE

        # Active pinch state
        self.pinch_detected = False
        self.pinch_finger = PinchFinger.NONE
        self.pinch_start_time = 0.0
        self.pinch_peak_vertical_delta = 0.0

        # Scrolling state
        self.scroll_active = False
        self.scroll_start_y = None
        self.scroll_finger = PinchFinger.NONE

        # Click bookkeeping
        self.last_click_time = 0.0         # for double click (left)
        self.last_right_click_time = 0.0   # for RC debounce
        self.RC_DEBOUNCE = 0.45

        # Movement suppression around middle–thumb right-click
        self.ignore_move_until = 0.0

        # Pinky-only OSK toggle latch/debounce
        self.PINKY_MIN_FRAMES = 2  # require a couple frames to avoid false positives
        self.pinky_hold_frames = 0
        self.pinky_latch = False

    # ---------- utils ----------
    def pixel_distance(self, p1, p2) -> float:
        dx = (p1.x - p2.x) * self.frame_w
        dy = (p1.y - p2.y) * self.frame_h
        return math.hypot(dx, dy)

    def is_finger_extended(self, lm, tip_idx, pip_idx):
        # tip above PIP joint means extended (y is smaller upwards in image space)
        return lm[tip_idx].y < lm[pip_idx].y

    def is_index_only_extended(self, lm):
        idx = self.is_finger_extended(lm, 8, 6)
        mid = self.is_finger_extended(lm, 12, 10)
        ring = self.is_finger_extended(lm, 16, 14)
        pinky = self.is_finger_extended(lm, 20, 18)
        return idx and not (mid or ring or pinky)

    def is_pinky_only_extended(self, lm):
        pinky = self.is_finger_extended(lm, 20, 18)
        idx   = self.is_finger_extended(lm, 8, 6)
        mid   = self.is_finger_extended(lm, 12, 10)
        ring  = self.is_finger_extended(lm, 16, 14)
        # Thumb doesn't matter; we only care about pinky-only among the 4 fingers
        return pinky and not (idx or mid or ring)

    def tip_id_for(self, finger: PinchFinger) -> int:
        if finger == PinchFinger.INDEX:
            return 8
        elif finger == PinchFinger.MIDDLE:
            return 12
        return 8  # default

    def pinch_distance_for(self, lm, finger: PinchFinger) -> float:
        thumb_tip = lm[4]
        tip = lm[self.tip_id_for(finger)]
        return self.pixel_distance(thumb_tip, tip)

    def tip_y_for(self, lm, finger: PinchFinger) -> float:
        return lm[self.tip_id_for(finger)].y

    # ---------- public helpers ----------
    def get_scroll_delta(self, lm) -> int:
        if self.scroll_start_y is None or self.scroll_finger == PinchFinger.NONE:
            return 0
        current_y = self.tip_y_for(lm, self.scroll_finger)
        delta_y = self.scroll_start_y - current_y
        if abs(delta_y) < self.SCROLL_DEADZONE:
            return 0
        # 0.02 normalized ~ one "step" of 120
        return int(delta_y * 50 * 120)

    # ---------- main FSM ----------
    def process_gesture(self, lm, now: float) -> GestureType:
        # If already scrolling: continue as long as the active pinch persists
        if self.current_state == GestureState.SCROLLING:
            if self.pinch_detected and self.pinch_finger != PinchFinger.NONE:
                d_active = self.pinch_distance_for(lm, self.pinch_finger)
                if d_active < self.PINCH_EXIT_THRESHOLD_PX:
                    return GestureType.SCROLLING
            # Pinch released -> stop scrolling
            self.current_state = GestureState.IDLE
            self.scroll_active = False
            self.scroll_start_y = None
            self.scroll_finger = PinchFinger.NONE
            self._scroll_arm_count = 0
            return GestureType.NONE

        # If a pinch is active, track it
        if self.pinch_detected and self.pinch_finger != PinchFinger.NONE:
            d_active = self.pinch_distance_for(lm, self.pinch_finger)
            pinch_now = d_active < self.PINCH_EXIT_THRESHOLD_PX

            # Update peak movement for click forgiveness
            if self.scroll_start_y is not None:
                dy = abs(self.tip_y_for(lm, self.pinch_finger) - self.scroll_start_y)
                if dy > self.pinch_peak_vertical_delta:
                    self.pinch_peak_vertical_delta = dy

            if not pinch_now:
                # Pinch just ended
                duration = now - self.pinch_start_time
                peak = self.pinch_peak_vertical_delta

                # Reset pinch state
                active_finger = self.pinch_finger
                self.pinch_detected = False
                self.pinch_finger = PinchFinger.NONE
                self.scroll_start_y = None
                self._scroll_arm_count = 0
                self.pinch_peak_vertical_delta = 0.0

                # If we never transitioned to scrolling, interpret as click
                if self.current_state != GestureState.SCROLLING:
                    if active_finger == PinchFinger.INDEX:
                        if duration <= self.CLICK_MAX_DURATION:
                            if now - self.last_click_time < self.DOUBLE_CLICK_WINDOW:
                                self.last_click_time = 0.0
                                return GestureType.DOUBLE_CLICK
                            else:
                                self.last_click_time = now
                                return GestureType.LEFT_CLICK
                        elif duration <= self.CLICK_FORGIVENESS_DURATION and peak < self.CLICK_FORGIVENESS_MOVEMENT:
                            self.last_click_time = now
                            return GestureType.LEFT_CLICK

                    elif active_finger == PinchFinger.MIDDLE:
                        if (now - self.last_right_click_time) > self.RC_DEBOUNCE:
                            if duration <= self.CLICK_MAX_DURATION:
                                self.last_right_click_time = now
                                self.ignore_move_until = now + 0.35  # ignore index-move briefly
                                return GestureType.RIGHT_CLICK
                            elif duration <= self.CLICK_FORGIVENESS_DURATION and peak < self.CLICK_FORGIVENESS_MOVEMENT:
                                self.last_right_click_time = now
                                self.ignore_move_until = now + 0.35
                                return GestureType.RIGHT_CLICK

                # Clean up any scrolling state if it existed
                if self.current_state == GestureState.SCROLLING:
                    self.current_state = GestureState.IDLE
                    self.scroll_active = False
                    self.scroll_finger = PinchFinger.NONE

                return GestureType.NONE

            else:
                # Pinch continues — check for scroll arm
                if (now - self.pinch_start_time) > self.SCROLL_HOLD_THRESHOLD and self.scroll_start_y is not None:
                    movement = abs(self.tip_y_for(lm, self.pinch_finger) - self.scroll_start_y)
                    if movement > self.SCROLL_MOVEMENT_THRESHOLD:
                        self._scroll_arm_count += 1
                        if self._scroll_arm_count >= self.SCROLL_ARM_FRAMES:
                            self.current_state = GestureState.SCROLLING
                            self.scroll_active = True
                            self.scroll_finger = self.pinch_finger
                            return GestureType.SCROLLING
                    else:
                        self._scroll_arm_count = 0

                # while pinching but not scrolling — no gesture to emit
                return GestureType.NONE

        # No pinch currently active — check if a new pinch starts
        # Compute distances for both possible pinches
        d_index = self.pinch_distance_for(lm, PinchFinger.INDEX)
        d_middle = self.pinch_distance_for(lm, PinchFinger.MIDDLE)

        # Prefer middle pinch if both are close (explicit RC gesture)
        if d_middle < self.PINCH_ENTER_THRESHOLD_PX or d_index < self.PINCH_ENTER_THRESHOLD_PX:
            if d_middle < self.PINCH_ENTER_THRESHOLD_PX and d_middle <= d_index:
                # Start middle–thumb pinch (will map to right-click logic)
                self.pinch_detected = True
                self.pinch_finger = PinchFinger.MIDDLE
                self.pinch_start_time = now
                self.scroll_start_y = self.tip_y_for(lm, PinchFinger.MIDDLE)
                self.pinch_peak_vertical_delta = 0.0
                self._scroll_arm_count = 0
                self.ignore_move_until = now + 0.20  # ignore index MOVE while initiating RC pinch
                return GestureType.NONE
            else:
                # Start index–thumb pinch (left-click path)
                self.pinch_detected = True
                self.pinch_finger = PinchFinger.INDEX
                self.pinch_start_time = now
                self.scroll_start_y = self.tip_y_for(lm, PinchFinger.INDEX)
                self.pinch_peak_vertical_delta = 0.0
                self._scroll_arm_count = 0
                return GestureType.NONE

        # Pinky-only => Toggle Windows On-Screen Keyboard (Win+Ctrl+O), single fire per hold
        if self.is_pinky_only_extended(lm):
            self.pinky_hold_frames += 1
            if not self.pinky_latch and self.pinky_hold_frames >= self.PINKY_MIN_FRAMES:
                self.pinky_latch = True
                # Suppress cursor move briefly to avoid jitter right after toggle
                self.ignore_move_until = now + 0.25
                return GestureType.OSK_TOGGLE
        else:
            # Reset latch when pinky released (or other fingers raised)
            self.pinky_hold_frames = 0
            self.pinky_latch = False

        # Movement (continuous while index-only), but respect RC/OSK suppression window
        if now >= self.ignore_move_until and self.is_index_only_extended(lm):
            self.current_state = GestureState.MOVING
            return GestureType.MOVE

        # reset to idle when not moving
        if self.current_state == GestureState.MOVING:
            self.current_state = GestureState.IDLE

        return GestureType.NONE


# ---------- Three-Finger Drag (extension-based, debounced) ----------
class ThreeFingerDragController:
    def __init__(self):
        # Debounce for quick, deliberate start/end
        self.MIN_FRAMES_TO_START = 2  # ~66ms @ 30 FPS
        self.MIN_FRAMES_TO_END = 2    # ~66ms @ 30 FPS

        # State
        self.is_dragging = False
        self.extension_detected = False
        self.group_confidence_frames = 0
        self.release_confidence_frames = 0

        self.drag_start_pos = None
        self.last_valid_pos = None

    def is_three_finger_extended(self, lm) -> bool:
        # index, middle, ring: tip above PIP; thumb/pinky don't matter
        try:
            index_extended = lm[8].y  < lm[6].y
            middle_extended = lm[12].y < lm[10].y
            ring_extended = lm[16].y   < lm[14].y
            return index_extended and middle_extended and ring_extended
        except Exception:
            return False

    def start_drag(self, pos):
        self.is_dragging = True
        self.drag_start_pos = pos
        self.last_valid_pos = pos

    def end_drag(self):
        self.is_dragging = False

    def reset_counters(self):
        self.group_confidence_frames = 0
        self.release_confidence_frames = 0

    def process_drag(self, lm, cursor_pos):
        # Hand lost → immediate drop
        if lm is None:
            if self.is_dragging:
                self.end_drag()
                return 'drag_end', self.last_valid_pos or cursor_pos
            self.extension_detected = False
            self.reset_counters()
            return 'none', None

        extension_now = self.is_three_finger_extended(lm)

        if extension_now:
            # Extension continues/starts
            self.release_confidence_frames = 0
            if not self.extension_detected:
                # Extension just started — arm the drag (no hold time)
                self.extension_detected = True
                self.group_confidence_frames = 1
                return 'none', None
            else:
                self.group_confidence_frames += 1
                if not self.is_dragging and self.group_confidence_frames >= self.MIN_FRAMES_TO_START:
                    self.start_drag(cursor_pos)
                    return 'drag_start', cursor_pos
                elif self.is_dragging:
                    self.last_valid_pos = cursor_pos
                    return 'dragging', cursor_pos
                else:
                    return 'none', None

        else:
            # Extension broken
            if self.extension_detected:
                self.extension_detected = False
                self.release_confidence_frames = 1
            elif self.is_dragging:
                self.release_confidence_frames += 1
            else:
                return 'none', None

            if self.is_dragging and self.release_confidence_frames >= self.MIN_FRAMES_TO_END:
                self.end_drag()
                return 'drag_end', self.last_valid_pos or cursor_pos
            elif self.is_dragging:
                # brief grace while broken
                return 'dragging', cursor_pos
            else:
                return 'none', None


# ---------- Fist/Punch Detector (cluster+curl based, debounced, entry-grace) ----------
class PunchFistDetector:
    """
    Detects a closed fist ("punch") robustly:
      - All four fingers curled (PIP angle small OR tip near MCP).
      - Thumb tucked (bent and near index MCP), not pinching index tip.
      - Fingertips cluster in a small radius (to reject random articulations).
    Distances are normalized by palm width (|5 - 17|) to be scale-invariant.
    """
    def __init__(self, frames_to_fire=5, require_curled=4):
        self.frames_to_fire = frames_to_fire
        self.require_curled = require_curled
        self.counter = 0

        # Tunables (normalized by palm width)
        self.PIP_CURL_MAX_DEG = 150          # finger considered curled if PIP angle < this
        self.TIP_MCP_DIST_MAX = 0.45         # tip must be close to mcp when curled
        self.TIP_CLUSTER_MAX = 0.55          # max radius from tip centroid (tight cluster)
        self.TIP_CLUSTER_AVG = 0.35          # optional average distance bound

        # Thumb tuck checks
        self.THUMB_BENT_MAX_DEG = 150        # thumb bent at IP if angle < this
        self.THUMB_TO_INDEX_MCP_MAX = 0.55   # thumb tip near index MCP (tucked)
        self.THUMB_TO_INDEX_TIP_MIN = 0.25   # avoid pinch: not too close to index tip

    def reset(self):
        self.counter = 0

    def dist2d(self, a, b):
        dx = a.x - b.x
        dy = a.y - b.y
        return math.hypot(dx, dy)

    def angle_deg(self, a, b, c):
        # angle at b formed by a-b-c
        bax = a.x - b.x; bay = a.y - b.y
        bcx = c.x - b.x; bcy = c.y - b.y
        dot = bax*bcx + bay*bcy
        na = math.hypot(bax, bay)
        nb = math.hypot(bcx, bcy)
        if na == 0 or nb == 0:
            return 180.0
        cosv = max(-1.0, min(1.0, dot / (na * nb)))
        return math.degrees(math.acos(cosv))

    def is_finger_curled(self, lm, tip_idx, pip_idx, mcp_idx, palm_w):
        # Curl via PIP angle plus fallback "collapsed" geometry
        wrist = lm[0]
        tip = lm[tip_idx]; pip = lm[pip_idx]; mcp = lm[mcp_idx]
        angle = self.angle_deg(mcp, pip, tip)
        tip_mcp = self.dist2d(tip, mcp)
        d_tip_w = self.dist2d(tip, wrist)
        d_pip_w = self.dist2d(pip, wrist)
        return (
            angle < self.PIP_CURL_MAX_DEG or
            tip_mcp < self.TIP_MCP_DIST_MAX * palm_w or
            d_tip_w < d_pip_w + 0.02 * palm_w
        )

    def thumb_tucked(self, lm, palm_w):
        # Thumb joints: 2(MCP), 3(IP), 4(tip)
        # Bent thumb and tip tucked near index MCP (5), not near index tip (8)
        mcp = lm[2]; ip = lm[3]; tip = lm[4]
        idx_mcp = lm[5]; idx_tip = lm[8]
        angle_ip = self.angle_deg(mcp, ip, tip)
        bent = angle_ip < self.THUMB_BENT_MAX_DEG
        near_index_base = self.dist2d(tip, idx_mcp) < self.THUMB_TO_INDEX_MCP_MAX * palm_w
        not_pinching = self.dist2d(tip, idx_tip) > self.THUMB_TO_INDEX_TIP_MIN * palm_w
        return bent and near_index_base and not_pinching

    def tips_clustered(self, lm, palm_w):
        tips = [lm[8], lm[12], lm[16], lm[20]]
        cx = sum(p.x for p in tips) / 4.0
        cy = sum(p.y for p in tips) / 4.0
        dists = [math.hypot(p.x - cx, p.y - cy) for p in tips]
        maxd = max(dists)
        avgd = sum(dists) / 4.0
        return (maxd < self.TIP_CLUSTER_MAX * palm_w) and (avgd < self.TIP_CLUSTER_AVG * palm_w)

    def fist_now(self, lm):
        # Use palm width for normalization
        palm_w = self.dist2d(lm[5], lm[17]) + 1e-6

        curled = 0
        if self.is_finger_curled(lm, 8, 6, 5, palm_w):  curled += 1   # index
        if self.is_finger_curled(lm, 12, 10, 9, palm_w): curled += 1  # middle
        if self.is_finger_curled(lm, 16, 14, 13, palm_w): curled += 1 # ring
        if self.is_finger_curled(lm, 20, 18, 17, palm_w): curled += 1 # pinky

        thumb_ok = self.thumb_tucked(lm, palm_w)
        cluster_ok = self.tips_clustered(lm, palm_w)

        # Strict enough to avoid false fires on first sight of hand
        return (curled >= self.require_curled) and thumb_ok and cluster_ok

    def process(self, lm, now=None):
        if lm is None:
            self.reset()
            return False
        if self.fist_now(lm):
            self.counter += 1
        else:
            self.counter = 0
        return self.counter >= self.frames_to_fire


# ---------- Hybrid Virtual Mouse ----------
class VirtualMouseHybrid:
    def __init__(self):
        self.mp_hands = mp.solutions.hands
        self.hands = self.mp_hands.Hands(
            static_image_mode=False,
            max_num_hands=1,
            min_detection_confidence=0.6,
            min_tracking_confidence=0.5,
            model_complexity=0
        )

        self.screen_w, self.screen_h = pyautogui.size()

        # Slightly tuned for stability without sluggishness
        self.fx = OneEuroFilter(min_cutoff=2.5, beta=0.03)
        self.fy = OneEuroFilter(min_cutoff=2.5, beta=0.03)

        self.recognizer = None
        self.dragger = ThreeFingerDragController()

        self.last_x = self.screen_w // 2
        self.last_y = self.screen_h // 2
        self.micro_deadzone_px = 3

        self.overlay_enabled = False
        self.mirror = True
        self.show_fps = True
        self.current_fps = 0
        self._fps_t = time.perf_counter()
        self._fps_n = 0

        self._last_landmarks = None

        # Action feedback display
        self.action_feedback = ""
        self.action_feedback_time = 0
        self.action_feedback_duration = 1.0  # Show feedback for 1 second

        # Platform
        self.is_windows = platform.system().lower() == "windows"

        # Fist/Punch handoff detector + hand-entry grace period (prevents instant fire)
        self.punch = PunchFistDetector(frames_to_fire=5, require_curled=4)
        self.HAND_ENTRY_GRACE = 0.45  # seconds to ignore handoff after first sight of a hand
        self._hand_was_present = False
        self._hand_first_seen_t = 0.0

        self.exit_requested = False

    def map_to_screen(self, nx, ny):
        # map 0.1..0.9 → 0..1 so edges are reachable
        x = (nx - 0.1) / 0.8
        y = (ny - 0.1) / 0.8
        x = max(0.0, min(1.0, x))
        y = max(0.0, min(1.0, y))
        sx = int(x * self.screen_w)
        sy = int(y * self.screen_h)
        return max(0, min(self.screen_w - 1, sx)), max(0, min(self.screen_h - 1, sy))

    def apply_pointer_dynamics(self, x, y):
        dx = x - self.last_x
        dy = y - self.last_y
        if abs(dx) < self.micro_deadzone_px and abs(dy) < self.micro_deadzone_px:
            # light smoothing on micro-movements
            sx = int(self.last_x * 0.7 + x * 0.3)
            sy = int(self.last_y * 0.7 + y * 0.3)
            return sx, sy
        return x, y

    def process_frame(self, frame):
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        t_now = time.perf_counter()
        results = self.hands.process(rgb)

        if results.multi_hand_landmarks:
            lm = results.multi_hand_landmarks[0].landmark
            self._last_landmarks = lm

            # Track hand entry to apply handoff grace window
            if not self._hand_was_present:
                self._hand_was_present = True
                self._hand_first_seen_t = t_now
                self.punch.reset()  # clear any stale counters on new hand

            # initialize recognizer with actual dims on first frame
            if self.recognizer is None:
                h, w = frame.shape[:2]
                self.recognizer = StrictGestureRecognizer(frame_w=w, frame_h=h)
                if DEBUG:
                    print(f"Recognizer dims set to {w}x{h}")

            # Fist/Punch (debounced) → exit to voice assistant (checked FIRST)
            allow_handoff = (t_now - self._hand_first_seen_t) >= self.HAND_ENTRY_GRACE
            if allow_handoff and self.punch.process(lm, t_now):
                return GestureType.HANDOFF, None, lm

            # index fingertip control point + filtered motion (for cursor mapping)
            nx, ny = lm[8].x, lm[8].y
            fx = self.fx.filter(nx, t_now)
            fy = self.fy.filter(ny, t_now)
            sx, sy = self.map_to_screen(fx, fy)

            # 1) Drag has priority (controller)
            drag_type, drag_pos = self.dragger.process_drag(lm, (sx, sy))
            if drag_type == 'drag_start':
                return GestureType.DRAG_START, drag_pos, lm
            elif drag_type == 'dragging':
                return GestureType.DRAGGING, drag_pos, lm
            elif drag_type == 'drag_end':
                return GestureType.DRAG_END, drag_pos, lm

            # 2) Other gestures (recognizer)
            g = self.recognizer.process_gesture(lm, t_now)

            # position for actions:
            if g == GestureType.MOVE:
                return g, (sx, sy), lm
            elif g in (GestureType.LEFT_CLICK, GestureType.RIGHT_CLICK, GestureType.DOUBLE_CLICK):
                # click where the cursor is (avoid jitter)
                return g, (self.last_x, self.last_y), lm
            elif g == GestureType.SCROLLING:
                # position doesn't matter for scroll
                return g, (self.last_x, self.last_y), lm
            elif g == GestureType.OSK_TOGGLE:
                return g, None, lm
            else:
                return GestureType.NONE, None, lm

        else:
            # no hand → reset fist debounce and entry tracker
            self.punch.reset()
            self._hand_was_present = False
            self._last_landmarks = None

        return GestureType.NONE, None, None

    def show_action_feedback(self, action_text):
        self.action_feedback = action_text
        self.action_feedback_time = time.perf_counter()

    def _press_hotkey_win_ctrl_o(self):
        # Using both hotkey and manual press for robustness on some Windows setups
        try:
            # Try standard hotkey call
            pyautogui.hotkey('winleft', 'ctrl', 'o')
            return True
        except Exception:
            pass
        try:
            # Alternate order
            pyautogui.hotkey('ctrl', 'winleft', 'o')
            return True
        except Exception:
            pass

        # Manual fallback
        try:
            pyautogui.keyDown('winleft')
            pyautogui.keyDown('ctrl')
            pyautogui.press('o')
            pyautogui.keyUp('ctrl')
            pyautogui.keyUp('winleft')
            return True
        except Exception:
            return False

    def toggle_osk(self):
        if not self.is_windows:
            if DEBUG:
                print("OSK toggle ignored: Non-Windows platform.")
            return
        ok = self._press_hotkey_win_ctrl_o()
        if not ok and DEBUG:
            print("OSK toggle hotkey failed.")

    def execute(self, gesture, pos, lm):
        if gesture == GestureType.NONE:
            return

        if gesture == GestureType.MOVE and pos:
            x, y = pos
            mx, my = self.apply_pointer_dynamics(x, y)
            # skip tiny moves to reduce OS event spam
            if abs(mx - self.last_x) >= 1 or abs(my - self.last_y) >= 1:
                pyautogui.moveTo(mx, my, duration=0)
                self.last_x, self.last_y = mx, my

        elif gesture == GestureType.LEFT_CLICK:
            pyautogui.click(self.last_x, self.last_y)
            self.show_action_feedback("LC")

        elif gesture == GestureType.DOUBLE_CLICK:
            pyautogui.doubleClick(self.last_x, self.last_y)
            self.show_action_feedback("DC")

        elif gesture == GestureType.RIGHT_CLICK:
            pyautogui.rightClick(self.last_x, self.last_y)
            self.show_action_feedback("RC")

        elif gesture == GestureType.DRAG_START and pos:
            x, y = pos
            pyautogui.mouseDown(x, y)
            self.last_x, self.last_y = x, y
            self.show_action_feedback("D")

        elif gesture == GestureType.DRAGGING and pos:
            x, y = pos
            pyautogui.moveTo(x, y, duration=0)
            self.last_x, self.last_y = x, y
            # Continue showing "D" while dragging
            if time.perf_counter() - self.action_feedback_time > 0.5:
                self.show_action_feedback("D")

        elif gesture == GestureType.DRAG_END:
            pyautogui.mouseUp()
            self.show_action_feedback("")

        elif gesture == GestureType.SCROLLING:
            if self._last_landmarks and self.recognizer:
                delta = self.recognizer.get_scroll_delta(self._last_landmarks)
                lines = int(delta / 120)
                if lines != 0:
                    pyautogui.scroll(lines)
                    self.show_action_feedback("S")

        elif gesture == GestureType.OSK_TOGGLE:
            self.toggle_osk()
            self.show_action_feedback("KB")

        elif gesture == GestureType.HANDOFF:
            # Trigger a graceful shutdown from run() after we draw feedback
            self.show_action_feedback("VA")
            self.exit_requested = True

    def draw_overlay(self, frame, gesture):
        h, w = frame.shape[:2]

        # Always draw action feedback (even when overlay is off)
        current_time = time.perf_counter()
        if self.action_feedback and (current_time - self.action_feedback_time) < self.action_feedback_duration:
            # Draw large action feedback in the center-top area
            font_scale = 3.0
            thickness = 5
            text_size = cv2.getTextSize(self.action_feedback, cv2.FONT_HERSHEY_SIMPLEX, font_scale, thickness)[0]

            # Position: center horizontally, top area vertically
            text_x = (w - text_size[0]) // 2
            text_y = 80

            # Draw background rectangle for better visibility
            padding = 20
            bg_x1 = text_x - padding
            bg_y1 = text_y - text_size[1] - padding
            bg_x2 = text_x + text_size[0] + padding
            bg_y2 = text_y + padding

            # Semi-transparent background
            overlay = frame.copy()
            cv2.rectangle(overlay, (bg_x1, bg_y1), (bg_x2, bg_y2), (0, 0, 0), -1)
            cv2.addWeighted(overlay, 0.7, frame, 0.3, 0, frame)

            # Draw text with white color
            cv2.putText(frame, self.action_feedback, (text_x, text_y),
                        cv2.FONT_HERSHEY_SIMPLEX, font_scale, (255, 255, 255), thickness, cv2.LINE_AA)

        # Draw regular overlay if enabled
        if self.overlay_enabled:
            cv2.rectangle(frame, (0, 0), (w, 28), (0, 0, 0), -1)
            txt = f"FPS: {self.current_fps} | Gesture: {gesture.value}"
            cv2.putText(frame, txt, (8, 20), cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 255, 0), 1, cv2.LINE_AA)

        return frame

    def configure_camera(self):
        # Try MSMF → DSHOW → ANY
        for backend, name in [(cv2.CAP_MSMF, "MSMF"), (cv2.CAP_DSHOW, "DSHOW"), (cv2.CAP_ANY, "ANY")]:
            cap = cv2.VideoCapture(0, backend)
            if cap.isOpened():
                # set props BEFORE reading to cut latency
                cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 360)
                cap.set(cv2.CAP_PROP_FPS, 30)
                cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
                # try MJPG for lower CPU if supported
                cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))

                ok, test = cap.read()
                if ok:
                    if DEBUG:
                        print(f"Camera ready via {name}")
                    return cap
                cap.release()
        return None

    def run(self):
        cap = self.configure_camera()
        if cap is None:
            print("Error: Could not open camera.")
            return

        win = "Vision Control (Make a Fist to Return to Assistant)"
        cv2.namedWindow(win, cv2.WINDOW_NORMAL)
        cv2.resizeWindow(win, 640, 400)

        print("VIRTUAL_MOUSE_STARTED")
        print("Tip: Three-finger extension (index+middle+ring) = Drag & Drop")
        print("Actions: LC=Left Click (index–thumb), RC=Right Click (middle–thumb), DC=Double Click, D=Drag, S=Scroll, KB=Toggle Keyboard (pinky-only)")
        print("Make a fist (punch) and hold briefly to return to the voice assistant.")

        exit_reason = "USER_QUIT"

        while True:
            ok, frame = cap.read()
            if not ok:
                continue

            if self.mirror:
                frame = cv2.flip(frame, 1)

            gesture, pos, lm = self.process_frame(frame)

            if gesture == GestureType.HANDOFF:
                # Ensure any drag is released before leaving
                if self.dragger.is_dragging:
                    try:
                        pyautogui.mouseUp()
                    except Exception:
                        pass
                self.execute(gesture, pos, lm)  # sets feedback + exit_requested
                frame = self.draw_overlay(frame, gesture)
                cv2.imshow(win, frame)
                cv2.waitKey(350)  # brief visual confirmation
                exit_reason = "HANDOFF"
                break

            self.execute(gesture, pos, lm)

            # FPS (cheap)
            self._fps_n += 1
            t_now = time.perf_counter()
            if self.show_fps and (t_now - self._fps_t) >= 1.0:
                self.current_fps = self._fps_n
                self._fps_n = 0
                self._fps_t = t_now

            # Always draw overlay (for action feedback)
            frame = self.draw_overlay(frame, gesture)
            cv2.imshow(win, frame)

            k = cv2.waitKey(1) & 0xFF
            if k in (ord('q'), ord('Q'), 27):  # Q or ESC
                # If dragging, release to avoid stuck button
                if self.dragger.is_dragging:
                    try:
                        pyautogui.mouseUp()
                    except Exception:
                        pass
                exit_reason = "USER_QUIT"
                break
            elif k in (ord('v'), ord('V')):
                self.overlay_enabled = not self.overlay_enabled
            elif k in (ord('m'), ord('M')):
                self.mirror = not self.mirror
            elif k in (ord('f'), ord('F')):
                self.show_fps = not self.show_fps

        cap.release()
        cv2.destroyAllWindows()
        self.hands.close()

        # Simple signal so the parent assistant can react if it wants to
        print(exit_reason)


if __name__ == "__main__":
    try:
        vm = VirtualMouseHybrid()
        vm.run()
    except Exception as e:
        print(f"Error: {e}")