# =========================================================
# VISUAL INSPECTION AUTOMATION SYSTEM
# =========================================================
# FEATURES
# =========================================================
# 1. MAIN SIGNAL SUPPORT
# 2. CALLING ON SUPPORT
# 3. SHUNT SUPPORT
# 4. SIGNAL CONFIG SAVE/LOAD
# 5. TRACK VISUAL INSPECTION
# 6. VISUAL INSPECTION PASS / FAIL REPORT
# 7. TRACK DROP SUPPORT
# 8. AUTO SIGNAL TYPE DETECTION
# =========================================================
#
# REQUIRED MODULES
#
# pip install pyautogui openpyxl pywinauto
# pip install pynput pillow uiautomation pywin32
#
# =========================================================
import os

import tkinter as tk
import keyboard
import sys
import openpyxl
from pynput import keyboard as pynput_keyboard
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import ttk
from tkinter import messagebox
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border
from openpyxl.styles import Side
from PIL import ImageGrab

from openpyxl.utils import get_column_letter

from openpyxl.drawing.image import Image

import threading
import os


import win32api
import win32con


from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

import uiautomation as auto
import pyautogui
import time

# =========================================================
# MULTI-MONITOR SCREEN CAPTURE
# =========================================================
# pyautogui.screenshot() only captures the PRIMARY monitor.
# If the simulator is on the second screen, its coordinates
# (e.g. X > 1920) fall outside that image and pixel reads fail.
# These helpers grab the whole virtual desktop (all monitors)
# and read pixels using the real screen coordinates, so the
# same code works with one screen or two.

class _VirtualScreenshot:
    """Full virtual-desktop grab; getpixel() takes SCREEN coordinates."""

    def __init__(self):
        self.left = win32api.GetSystemMetrics(76)    # SM_XVIRTUALSCREEN
        self.top = win32api.GetSystemMetrics(77)     # SM_YVIRTUALSCREEN
        self.image = ImageGrab.grab(all_screens=True).convert("RGB")

    def getpixel(self, xy):
        x, y = xy
        return self.image.getpixel((int(x) - self.left, int(y) - self.top))

    def save(self, *args, **kwargs):
        return self.image.save(*args, **kwargs)


def grab_all_screens():
    """Screenshot of ALL monitors (single or dual screen)."""
    return _VirtualScreenshot()


def grab_screen_region(x, y, w, h):
    """Region screenshot in SCREEN coordinates, on any monitor."""
    return ImageGrab.grab(
        bbox=(int(x), int(y), int(x) + int(w), int(y) + int(h)),
        all_screens=True
    ).convert("RGB")

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "VISUAL_INSPECTION_ON_TRACK_REPORT.xlsx"

COORD_FILE = os.environ.get(
    "EDRC_TRACK_COORDS",
    os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "track_coordinates.xlsx"
    )
)

# =========================================================
# GLOBALS
# =========================================================

signals = {}

lock_routes_data = []

routes_data = []

track_points = {}

# Track coordinate list imported from Excel for capture
track_capture_list = []
track_capture_index = 0
track_undo_stack = []

running = False

capture_queue = []

capture_module = None

capture_index = 0

capture_mode = None

current_signal = None

capture_master_list = []

current_step = 0

last_capture_time = 0

undo_stack = []

current_signal = None

config_file_path = ""

# =========================================================
# LEGACY CAPTURE COMPATIBILITY
# =========================================================
# The current signal wizard uses master_save_point(), but the
# older create_setup() function still calls capture_point().
# Keep this small compatibility layer so the file remains
# runnable without changing the Main/Calling-On automation.
capture_waiting = False
captured_point = None

# =========================================================
# CAPTURE OVERLAY
# =========================================================

capture_overlay = None

overlay_progress_label = None

overlay_signal_label = None

overlay_step_label = None

overlay_hint_label = None

overlay_button_frame = None

# =========================================================
# PAUSE CONTROL
# =========================================================

paused = False

pause_event = threading.Event()

pause_event.set()

# =========================================================
# COLORS
# =========================================================

BG = "#0f172a"

BTN = "#2563eb"

GREEN = "#22c55e"

RED = "#ef4444"

# =========================================================
# CAPTURE GUIDE COLORS
# =========================================================

TYPE_LABELS = {
    "MAIN": "MAIN SIGNAL",
    "CALLING_ON": "CALLING-ON SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "POINT": "POINT MACHINE",
    "CH": "EMERGENCY CRANK HANDLE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CALLING_ON": "#f97316",
    "SHUNT": "#a855f7",
    "POINT": "#0891b2",
    "CH": "#16a34a",
    "LC": "#db2777"
}

# =========================================================
# REPORT DATA STORAGE
# =========================================================

report_rows = []

report_saved = False
current_test = ""

def create_capture_overlay():

    global capture_overlay, overlay_progress_label, overlay_signal_label
    global overlay_step_label, overlay_hint_label, overlay_button_frame

    capture_overlay = tk.Toplevel(root)
    capture_overlay.title("Coordinate Capture Guide")
    capture_overlay.configure(bg="#0f172a")
    capture_overlay.resizable(False, False)
    capture_overlay.attributes("-topmost", True)

    screen_w = capture_overlay.winfo_screenwidth()
    capture_overlay.geometry(f"420x320+{screen_w - 440}+40")

    capture_overlay.protocol("WM_DELETE_WINDOW", cancel_master_capture)

    # SPACE/BACKSPACE/ENTER are handled globally via pynput (dispatch_save_point
    # / dispatch_undo). If this window or a button inside it ever ends up with
    # keyboard focus, Tk's own default "space/enter activates focused button"
    # behavior could double-fire alongside that. Swallowing the events here
    # guarantees they never do.
    capture_overlay.bind("<space>", lambda e: "break")
    capture_overlay.bind("<Return>", lambda e: "break")

    tk.Label(
        capture_overlay, text="COORDINATE CAPTURE",
        font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="#64748b"
    ).pack(pady=(16, 0))

    overlay_progress_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 10), bg="#0f172a", fg="#94a3b8"
    )
    overlay_progress_label.pack(pady=(2, 10))

    overlay_signal_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 18, "bold"), bg="#0f172a", fg="white"
    )
    overlay_signal_label.pack()

    overlay_step_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 15, "bold"), bg="#0f172a", fg="#22c55e"
    )
    overlay_step_label.pack(pady=(6, 8))

    overlay_hint_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 10), bg="#0f172a", fg="#cbd5e1",
        wraplength=380, justify="center"
    )
    overlay_hint_label.pack(pady=(0, 10))

    overlay_button_frame = tk.Frame(capture_overlay, bg="#0f172a")
    overlay_button_frame.pack(pady=(0, 6))

    tk.Button(
        capture_overlay, text="UNDO LAST CLICK", command=undo_last_capture,
        bg="#f59e0b", fg="#1a1a1a", activebackground="#fbbf24",
        activeforeground="#1a1a1a", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=22, height=1, takefocus=0
    ).pack(pady=(10, 2))

    tk.Label(
        capture_overlay, text="(or press BACKSPACE)",
        font=("Segoe UI", 8), bg="#0f172a", fg="#475569"
    ).pack()

    tk.Button(
        capture_overlay, text="CANCEL CAPTURE", command=cancel_master_capture,
        bg="#0f172a", fg="#64748b", activebackground="#0f172a",
        activeforeground="#ef4444", relief="flat", cursor="hand2",
        font=("Segoe UI", 9, "underline"), takefocus=0
    ).pack(side="bottom", pady=(0, 10))

def destroy_capture_overlay():

    global capture_overlay

    if capture_overlay is not None:

        try:
            capture_overlay.destroy()
        except:
            pass

        capture_overlay = None


def create_track_capture_overlay():
    """Create the same dark TL-style popup used for signal capture,
    but dedicated to sequential track-coordinate capture."""
    global capture_overlay, overlay_progress_label, overlay_signal_label
    global overlay_step_label, overlay_hint_label, overlay_button_frame

    capture_overlay = tk.Toplevel(root)
    capture_overlay.title("Coordinate Capture Guide")
    capture_overlay.configure(bg="#0f172a")
    capture_overlay.resizable(False, False)
    capture_overlay.attributes("-topmost", True)

    screen_w = capture_overlay.winfo_screenwidth()
    capture_overlay.geometry(f"420x320+{screen_w - 440}+40")

    capture_overlay.protocol("WM_DELETE_WINDOW", cancel_track_capture)

    capture_overlay.bind("<space>", lambda e: "break")
    capture_overlay.bind("<Return>", lambda e: "break")

    tk.Label(
        capture_overlay, text="COORDINATE CAPTURE",
        font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="#64748b"
    ).pack(pady=(16, 0))

    overlay_progress_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 10), bg="#0f172a", fg="#94a3b8"
    )
    overlay_progress_label.pack(pady=(2, 10))

    overlay_signal_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 18, "bold"), bg="#0f172a", fg="#a855f7"
    )
    overlay_signal_label.pack()

    overlay_step_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 15, "bold"), bg="#0f172a", fg="#22c55e"
    )
    overlay_step_label.pack(pady=(6, 8))

    overlay_hint_label = tk.Label(
        capture_overlay, text="",
        font=("Segoe UI", 10), bg="#0f172a", fg="#cbd5e1",
        wraplength=380, justify="center"
    )
    overlay_hint_label.pack(pady=(0, 10))

    overlay_button_frame = tk.Frame(capture_overlay, bg="#0f172a")
    overlay_button_frame.pack(pady=(0, 6))

    tk.Button(
        capture_overlay, text="UNDO LAST CLICK", command=undo_last_track_capture,
        bg="#f59e0b", fg="#1a1a1a", activebackground="#fbbf24",
        activeforeground="#1a1a1a", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=22, height=1, takefocus=0
    ).pack(pady=(10, 2))

    tk.Label(
        capture_overlay, text="(or press BACKSPACE)",
        font=("Segoe UI", 8), bg="#0f172a", fg="#475569"
    ).pack()

    tk.Button(
        capture_overlay, text="CANCEL CAPTURE", command=cancel_track_capture,
        bg="#0f172a", fg="#64748b", activebackground="#0f172a",
        activeforeground="#ef4444", relief="flat", cursor="hand2",
        font=("Segoe UI", 9, "underline"), takefocus=0
    ).pack(side="bottom", pady=(0, 10))


def update_track_capture_overlay():
    if capture_overlay is None or not track_capture_list:
        return

    clear_overlay_buttons()

    overlay_progress_label.config(
        text=f"TRACK {track_capture_index + 1} of {len(track_capture_list)}"
    )
    overlay_signal_label.config(
        text=track_capture_list[track_capture_index],
        fg="#a855f7"
    )
    overlay_step_label.config(text="CAPTURE COORDINATE")
    overlay_hint_label.config(
        text="Move the mouse onto the track.\n\n"
            "Press SPACE to save.\n"
            "Press BACKSPACE to undo."
    )

    capture_overlay.lift()


def cancel_track_capture():
    global capture_module
    capture_module = None
    destroy_capture_overlay()
    root.deiconify()
    log("TRACK CAPTURE CANCELLED")


def undo_last_track_capture():
    global track_capture_index

    if not track_undo_stack:
        log("Nothing to undo")
        return

    action = track_undo_stack.pop()
    track = action["track"]

    if action["old_value"] is None:
        track_points.pop(track, None)
    else:
        track_points[track] = action["old_value"]

    track_capture_index = action["capture_index"]
    update_track_capture_overlay()
    log(f"Last track capture undone : {track}")


def finish_track_capture():
    global capture_module
    capture_module = None
    destroy_capture_overlay()
    root.deiconify()
    log("ALL TRACK COORDINATES CAPTURED")
    log(f"TRACKS CAPTURED : {len(track_capture_list)}")


def capture_track_point(x, y):
    global track_capture_index

    if capture_module != "TRACK" or track_capture_index >= len(track_capture_list):
        return

    track = track_capture_list[track_capture_index]
    old_value = track_points.get(track)

    track_undo_stack.append({
        "track": track,
        "old_value": list(old_value) if isinstance(old_value, (list, tuple)) and len(old_value) == 2 else None,
        "capture_index": track_capture_index
    })

    track_points[track] = [x, y]
    log(f"TRACK SAVED: {track} at ({x}, {y})")

    try:
        save_coordinate(track, x, y)
    except Exception as e:
        log(f"COORDINATE FILE SAVE ERROR : {e}")

    track_capture_index += 1

    if track_capture_index >= len(track_capture_list):
        finish_track_capture()
    else:
        update_track_capture_overlay()


def clear_overlay_buttons():

    if overlay_button_frame is None:
        return

    for widget in overlay_button_frame.winfo_children():
        widget.destroy()


def update_capture_overlay(signal_type, signal, step_text, hint_text):

    if capture_overlay is None:
        return

    clear_overlay_buttons()

    overlay_progress_label.config(
        text=TYPE_LABELS.get(signal_type, "")
    )

    overlay_signal_label.config(
        text=signal,
        fg=TYPE_COLORS.get(signal_type, "white")
    )

    overlay_step_label.config(text=step_text)

    overlay_hint_label.config(text=hint_text)

    capture_overlay.lift()

def show_aspect_buttons(signal):

    if capture_overlay is None:
        return

    clear_overlay_buttons()

    overlay_progress_label.config(
        text=f"Signal {capture_index + 1} of {len(capture_master_list)}   |   {TYPE_LABELS['MAIN']}"
    )
    overlay_signal_label.config(text=signal, fg=TYPE_COLORS["MAIN"])
    overlay_step_label.config(text="HOW MANY ASPECTS?")
    overlay_hint_label.config(text="Click 2 / 3 / 4 below, or press the 2, 3, or 4 key on your keyboard.")

    for n in (2, 3, 4):

        tk.Button(
            overlay_button_frame, text=str(n), width=6, height=1,
            font=("Segoe UI", 12, "bold"),
            bg="#2563eb", fg="white", activebackground="#2563eb",
            activeforeground="white", relief="flat", cursor="hand2",
            command=lambda n=n: set_aspects_and_continue(signal, n),
            takefocus=0
        ).pack(side="left", padx=6)

    capture_overlay.lift()

def push_undo(clear_fn):
    """Records how to reverse the point about to be captured. Call this
    BEFORE mutating any data or advancing record_stage/capture_index, so
    it snapshots the state as it was just before this capture."""

    undo_stack.append({
        "capture_index": capture_index,
        "record_stage": record_stage,
        "clear": clear_fn
    })

def undo_last_capture():

    global capture_index
    global record_stage

    if not undo_stack:

        log("Nothing to undo")

        return

    action = undo_stack.pop()

    action["clear"]()

    capture_index = action["capture_index"]

    record_stage = action["record_stage"]

    master_next_capture()

    log("Last capture undone")

def cancel_master_capture():

    destroy_capture_overlay()

    root.deiconify()

    log("CAPTURE CANCELLED")

def build_capture_master_list():

    capture_list = []

    # MAIN SIGNALS
    for signal, info in signals.items():

        if info["type"] == "MAIN":
            capture_list.append(("MAIN", signal))

    # CALLING ON
    for signal, info in signals.items():

        if info["type"] == "CALLING_ON":
            capture_list.append(("CALLING_ON", signal))

    # SHUNT
    for signal, info in signals.items():

        if info["type"] == "SHUNT":
            capture_list.append(("SHUNT", signal))

    return capture_list

def set_initial_stage_for_current():

    global record_stage

    sig_type, signal = capture_master_list[capture_index]

    if sig_type == "MAIN":

        record_stage = "coordinate"

    elif sig_type == "SHUNT":

        record_stage = "menu"

    elif sig_type == "CALLING_ON":

        record_stage = "menu"

    elif sig_type == "POINT":

        # POINT capture starts from MENU.
        record_stage = "menu"

    elif sig_type == "CH":

        # CRANK HANDLE capture starts from MENU.
        record_stage = "menu"

    elif sig_type == "LC":

        # LC GATE capture starts from MENU.
        record_stage = "menu"

def advance_to_next_signal():

    global capture_index

    capture_index += 1

    if capture_index < len(capture_master_list):
        set_initial_stage_for_current()

    master_next_capture()

def build_signals_from_toc():

    global signals

    signals = {}

    added = set()

    for row in routes_data:

        signal = str(row["signal"]).strip().upper()

        if signal in added:
            continue

        added.add(signal)

        if signal.startswith("SH"):

            signals[signal] = {
                "type": "SHUNT",
                "menu": None,
                "state_indicator": None,
                "route_init": None,
                "initial_snapshot": None,
                "C1": None,
                "C2": None,
                "C3": None
            }

        elif signal.endswith("C"):

            signals[signal] = {
                "type": "CALLING_ON",
                "menu": None,
                "YELLOW": None,
                "ROUTE_INIT": None
            }

        else:

            signals[signal] = {
                "type": "MAIN",
                "aspects": None,
                "menu": None,
                "RED": None,
                "YELLOW": None,
                "DOUBLE_YELLOW": None,
                "GREEN": None,
                "ROUTE_INDICATOR": None
            }

    log(f"{len(signals)} signals loaded from TOC.")

def start_master_capture():

    global capture_module
    global capture_master_list
    global capture_index

    capture_master_list = build_capture_master_list()

    if not capture_master_list:

        log("No Signals To Capture")

        return

    capture_index = 0

    capture_module = "MASTER"

    undo_stack.clear()

    set_initial_stage_for_current()

    create_capture_overlay()

    root.iconify()

    log("Recording Signal Coordinates")

    log(f"Total Signals : {len(capture_master_list)}")

    master_next_capture()

def master_next_capture():

    global capture_module

    if capture_index >= len(capture_master_list):
        log("ALL SIGNAL COORDINATES CAPTURED")

        capture_module = None

        destroy_capture_overlay()

        root.deiconify()

        save_config()

        messagebox.showinfo(
            "Completed",
            "Signal configuration saved successfully."
        )

        return

    sig_type, signal = capture_master_list[capture_index]

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if sig_type == "MAIN":

        if record_stage == "coordinate":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : SIGNAL MENU",
                "Move the mouse onto the signal menu.\n\nPress SPACE."
            )

        elif record_stage == "ask_aspects":

            show_aspect_buttons(signal)

        elif record_stage == "RED":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : RED ASPECT",
                "Move the mouse onto the RED lamp.\n\nPress SPACE."
            )

        elif record_stage == "YELLOW":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : YELLOW ASPECT",
                "Move the mouse onto the YELLOW lamp.\n\nPress SPACE."
            )

        elif record_stage == "DOUBLE_YELLOW":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : DOUBLE YELLOW",
                "Move the mouse onto the DOUBLE YELLOW lamp.\n\nPress SPACE."
            )

        elif record_stage == "GREEN":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : GREEN ASPECT",
                "Move the mouse onto the GREEN lamp.\n\nPress SPACE."
            )

        elif record_stage == "ROUTE_INDICATOR":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK : ROUTE INDICATOR",
                "Move the mouse onto the route initiation indicator.\n\nPress SPACE."
            )

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif sig_type == "SHUNT":

        if record_stage == "menu":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK : MENU",
                "Move the mouse onto the menu.\n\nPress SPACE."
            )

        elif record_stage == "C1" or record_stage == "indicator":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK : ASPECT INDICATOR",
                "Move the mouse onto the aspect/state indicator, then press SPACE."
            )

        elif record_stage == "C2" or record_stage == "C3" or record_stage == "route_init":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

        elif record_stage == "route_init":
            update_capture_overlay(
                "SHUNT", signal, "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # CALLING ON SIGNAL
    # =====================================================

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK : MENU",
                "Move the mouse onto the signal menu.\n\nPress SPACE."
            )

        elif record_stage == "YELLOW":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK : YELLOW",
                "Move the mouse onto the yellow lamp.\n\nPress SPACE."
            )

        elif record_stage == "ROUTE_INIT":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator.\n\nPress SPACE."
            )

    # =====================================================
    # POINT - MENU -> NORMAL -> REVERSE -> FREE
    # =====================================================

    elif sig_type == "POINT":
        if record_stage == "menu":
            update_capture_overlay(
                "POINT", signal, "CLICK : MENU",
                "Move the mouse onto this point's menu button, then press SPACE."
            )
        elif record_stage == "normal":
            update_capture_overlay(
                "POINT", signal, "CLICK : NORMAL ASPECT",
                "Move the mouse onto the NORMAL indication, then press SPACE."
            )
        elif record_stage == "reverse":
            update_capture_overlay(
                "POINT", signal, "CLICK : REVERSE ASPECT",
                "Move the mouse onto the REVERSE indication, then press SPACE."
            )
        elif record_stage == "free":
            update_capture_overlay(
                "POINT", signal, "CLICK : FREE ASPECT",
                "Move the mouse onto the FREE / OUT-OF-CORRESPONDENCE indication, then press SPACE."
            )

    # =====================================================
    # CRANK HANDLE - MENU -> IN -> OUT -> ECH -> FREE
    # =====================================================

    elif sig_type == "CH":
        if record_stage == "menu":
            update_capture_overlay(
                "CH", signal, "CLICK : MENU",
                "Move the mouse onto this crank handle's menu button, then press SPACE."
            )
        elif record_stage == "IN":
            update_capture_overlay(
                "CH", signal, "CLICK : IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )
        elif record_stage == "OUT":
            update_capture_overlay(
                "CH", signal, "CLICK : OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )
        elif record_stage == "ECH":
            update_capture_overlay(
                "CH", signal, "CLICK : ECH ASPECT",
                "Move the mouse onto the ECH indicator, then press SPACE."
            )
        elif record_stage == "FREE":
            update_capture_overlay(
                "CH", signal, "CLICK : FREE ASPECT",
                "Move the mouse onto the FREE indicator, then press SPACE."
            )

    # =====================================================
    # LC GATE - MENU -> IN -> OUT
    # =====================================================

    elif sig_type == "LC":
        if record_stage == "menu":
            update_capture_overlay(
                "LC", signal, "CLICK : MENU",
                "Move the mouse onto this LC gate's menu button, then press SPACE."
            )
        elif record_stage == "IN":
            update_capture_overlay(
                "LC", signal, "CLICK : IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )
        elif record_stage == "OUT":
            update_capture_overlay(
                "LC", signal, "CLICK : OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )

def master_save_point(x, y):

    global record_stage

    if capture_index >= len(capture_master_list):
        return

    sig_type, signal = capture_master_list[capture_index]

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if sig_type == "MAIN":

        if record_stage == "coordinate":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved")

            record_stage = "ask_aspects"

            master_next_capture()

            return

        if record_stage == "ask_aspects":
            return

        if record_stage == "RED":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("RED", None)
            )

            signals[signal]["RED"] = [x, y]

            log(f"{signal} RED Saved")

            if current_aspects == 2:
                record_stage = "GREEN"

            else:
                record_stage = "YELLOW"

            master_next_capture()

            return

        if record_stage == "YELLOW":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("YELLOW", None)
            )

            signals[signal]["YELLOW"] = [x, y]

            log(f"{signal} YELLOW Saved")

            if current_aspects == 4:
                record_stage = "DOUBLE_YELLOW"

            else:
                record_stage = "GREEN"

            master_next_capture()

            return

        if record_stage == "DOUBLE_YELLOW":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("DOUBLE_YELLOW", None)
            )

            signals[signal]["DOUBLE_YELLOW"] = [x, y]

            log(f"{signal} DOUBLE YELLOW Saved")

            record_stage = "GREEN"

            master_next_capture()

            return

        if record_stage == "GREEN":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("GREEN", None)
            )

            signals[signal]["GREEN"] = [x, y]

            log(f"{signal} GREEN Saved")

            record_stage = "ROUTE_INDICATOR"

            master_next_capture()

            return

        if record_stage == "ROUTE_INDICATOR":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("ROUTE_INDICATOR", None)
            )

            signals[signal]["ROUTE_INDICATOR"] = [x, y]

            log(f"{signal} ROUTE INDICATOR Saved")

            advance_to_next_signal()

            return

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif sig_type == "SHUNT":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved")

            record_stage = "C1"

            master_next_capture()

            return

        if record_stage == "C1":
            # Backward-compatible entry point: old configs may still use C1.
            push_undo(lambda s=signal: signals[s].__setitem__("state_indicator", None))
            signals[signal]["state_indicator"] = [x, y]
            signals[signal]["C1"] = [x, y]
            log(f"{signal} ASPECT INDICATOR Saved")
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "indicator":
            push_undo(lambda s=signal: signals[s].__setitem__("state_indicator", None))
            signals[signal]["state_indicator"] = [x, y]
            signals[signal]["C1"] = [x, y]
            log(f"{signal} ASPECT INDICATOR Saved")
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":
            push_undo(lambda s=signal: signals[s].__setitem__("route_init", None))
            signals[signal]["route_init"] = [x, y]
            signals[signal]["C2"] = [x, y]
            signals[signal]["C3"] = [x, y]
            log(f"{signal} ROUTE INITIATION Saved")
            advance_to_next_signal()
            return

        # Legacy C2/C3 stages, if an older state is restored.
        if record_stage == "C2":
            signals[signal]["C2"] = [x, y]
            signals[signal]["route_init"] = [x, y]
            record_stage = "C3"
            master_next_capture()
            return

        if record_stage == "C3":
            signals[signal]["C3"] = [x, y]
            signals[signal]["route_init"] = [x, y]
            advance_to_next_signal()
            return

    # =====================================================
    # CALLING ON SIGNAL
    # =====================================================

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved")

            record_stage = "YELLOW"

            master_next_capture()

            return

        if record_stage == "YELLOW":
            push_undo(lambda s=signal: signals[s].__setitem__("YELLOW", None))
            signals[signal]["YELLOW"] = [x, y]
            log(f"{signal} YELLOW Saved")
            record_stage = "ROUTE_INIT"
            master_next_capture()
            return

        if record_stage == "ROUTE_INIT":
            push_undo(lambda s=signal: signals[s].__setitem__("ROUTE_INIT", None))
            signals[signal]["ROUTE_INIT"] = [x, y]
            log(f"{signal} ROUTE INITIATION Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # POINT - MENU -> NORMAL -> REVERSE -> FREE
    # =====================================================

    elif sig_type == "POINT":
        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            # Keep coordinate as a compatibility alias for older automation.
            signals[signal]["coordinate"] = [x, y]
            log(f"{signal} POINT MENU Saved ({x},{y})")
            record_stage = "normal"
            master_next_capture()
            return
        if record_stage == "normal":
            push_undo(lambda s=signal: signals[s].__setitem__("normal", None))
            signals[signal]["normal"] = [x, y]
            log(f"{signal} POINT NORMAL Saved ({x},{y})")
            record_stage = "reverse"
            master_next_capture()
            return
        if record_stage == "reverse":
            push_undo(lambda s=signal: signals[s].__setitem__("reverse", None))
            signals[signal]["reverse"] = [x, y]
            log(f"{signal} POINT REVERSE Saved ({x},{y})")
            record_stage = "free"
            master_next_capture()
            return
        if record_stage == "free":
            push_undo(lambda s=signal: signals[s].__setitem__("free", None))
            signals[signal]["free"] = [x, y]
            log(f"{signal} POINT FREE Saved ({x},{y})")
            advance_to_next_signal()
            return

    # =====================================================
    # CRANK HANDLE - MENU -> IN -> OUT -> ECH -> FREE
    # =====================================================

    elif sig_type == "CH":
        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            signals[signal]["coordinate"] = [x, y]
            log(f"{signal} CH MENU Saved ({x},{y})")
            record_stage = "IN"
            master_next_capture()
            return
        if record_stage == "IN":
            push_undo(lambda s=signal: signals[s].__setitem__("IN", None))
            signals[signal]["IN"] = [x, y]
            log(f"{signal} CH IN Saved ({x},{y})")
            record_stage = "OUT"
            master_next_capture()
            return
        if record_stage == "OUT":
            push_undo(lambda s=signal: signals[s].__setitem__("OUT", None))
            signals[signal]["OUT"] = [x, y]
            log(f"{signal} CH OUT Saved ({x},{y})")
            record_stage = "ECH"
            master_next_capture()
            return
        if record_stage == "ECH":
            push_undo(lambda s=signal: signals[s].__setitem__("ECH", None))
            signals[signal]["ECH"] = [x, y]
            log(f"{signal} CH ECH Saved ({x},{y})")
            record_stage = "FREE"
            master_next_capture()
            return
        if record_stage == "FREE":
            push_undo(lambda s=signal: signals[s].__setitem__("FREE", None))
            signals[signal]["FREE"] = [x, y]
            log(f"{signal} CH FREE Saved ({x},{y})")
            advance_to_next_signal()
            return

    # =====================================================
    # LC GATE - MENU -> IN -> OUT
    # =====================================================

    elif sig_type == "LC":
        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            signals[signal]["coordinate"] = [x, y]
            log(f"{signal} LC MENU Saved ({x},{y})")
            record_stage = "IN"
            master_next_capture()
            return
        if record_stage == "IN":
            push_undo(lambda s=signal: signals[s].__setitem__("IN", None))
            signals[signal]["IN"] = [x, y]
            log(f"{signal} LC IN Saved ({x},{y})")
            record_stage = "OUT"
            master_next_capture()
            return
        if record_stage == "OUT":
            push_undo(lambda s=signal: signals[s].__setitem__("OUT", None))
            signals[signal]["OUT"] = [x, y]
            log(f"{signal} LC OUT Saved ({x},{y})")
            advance_to_next_signal()
            return


def set_aspects_and_continue(signal, aspects):

    global current_aspects
    global record_stage

    current_aspects = aspects

    signals[signal]["aspects"] = aspects

    record_stage = "RED"

    master_next_capture()

def center_dialog(dlg, width=380, height=280):
    """Centers a Toplevel on screen - used for the count/ID prompts, which
    appear while the main window is still visible (not minimized), unlike
    the capture overlay which is positioned top-right once capture starts."""

    screen_w = dlg.winfo_screenwidth()
    screen_h = dlg.winfo_screenheight()
    dlg.geometry(f"{width}x{height}+{(screen_w - width) // 2}+{(screen_h - height) // 2}")


def styled_ask_count(label):
    """Dark-themed replacement for simpledialog.askinteger, styled to match
    the capture overlay (same colors/fonts/button look). Blocks until the
    operator confirms or skips; returns 0 for skip."""

    result = {"value": 0}

    dlg = tk.Toplevel(root)
    dlg.title("How Many?")
    dlg.configure(bg="#0f172a")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.transient(root)
    dlg.grab_set()

    tk.Label(
        dlg, text="HOW MANY?",
        font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="#64748b"
    ).pack(pady=(20, 4), padx=40)

    tk.Label(
        dlg, text=label,
        font=("Segoe UI", 16, "bold"), bg="#0f172a", fg="white",
        wraplength=340, justify="center"
    ).pack(pady=(0, 16), padx=20)

    entry_var = tk.StringVar(value="0")

    entry = tk.Entry(
        dlg, textvariable=entry_var, font=("Segoe UI", 20, "bold"),
        justify="center", width=6, bg="#1e293b", fg="white",
        insertbackground="white", relief="flat"
    )
    entry.pack(pady=(0, 6), ipady=6)
    entry.focus_set()
    entry.select_range(0, tk.END)

    tk.Label(
        dlg, text="Enter 0 to skip this category",
        font=("Segoe UI", 9), bg="#0f172a", fg="#475569"
    ).pack(pady=(0, 14))

    def confirm():
        try:
            result["value"] = max(0, int(entry_var.get().strip() or 0))
        except ValueError:
            result["value"] = 0
        dlg.destroy()

    def skip():
        result["value"] = 0
        dlg.destroy()

    btn_frame = tk.Frame(dlg, bg="#0f172a")
    btn_frame.pack(pady=(0, 20))

    tk.Button(
        btn_frame, text="CONFIRM", command=confirm,
        bg="#2563eb", fg="white", activebackground="#2563eb",
        activeforeground="white", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=12, height=1, takefocus=0
    ).pack(side="left", padx=6)

    tk.Button(
        btn_frame, text="SKIP (0)", command=skip,
        bg="#0f172a", fg="#64748b", activebackground="#0f172a",
        activeforeground="#ef4444", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "underline"), takefocus=0
    ).pack(side="left", padx=6)

    dlg.bind("<Return>", lambda e: confirm())

    center_dialog(dlg, 380, 260)

    dlg.wait_window()

    return result["value"]


def styled_ask_name(label, index, total):
    """Dark-themed replacement for simpledialog.askstring, styled to match
    the capture overlay. Blocks until the operator confirms."""

    result = {"value": None}

    dlg = tk.Toplevel(root)
    dlg.title("Signal / Element ID")
    dlg.configure(bg="#0f172a")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.transient(root)
    dlg.grab_set()

    tk.Label(
        dlg, text=f"{label.upper()}   |   {index + 1} OF {total}",
        font=("Segoe UI", 11, "bold"), bg="#0f172a", fg="#64748b"
    ).pack(pady=(20, 4), padx=40)

    tk.Label(
        dlg, text="ENTER ID",
        font=("Segoe UI", 16, "bold"), bg="#0f172a", fg="white"
    ).pack(pady=(0, 12))

    entry_var = tk.StringVar()

    entry = tk.Entry(
        dlg, textvariable=entry_var, font=("Segoe UI", 16, "bold"),
        justify="center", width=18, bg="#1e293b", fg="white",
        insertbackground="white", relief="flat"
    )
    entry.pack(pady=(0, 6), ipady=6, padx=20)
    entry.focus_set()

    tk.Label(
        dlg, text="Leave blank for an auto-generated ID",
        font=("Segoe UI", 9), bg="#0f172a", fg="#475569"
    ).pack(pady=(0, 14))

    def confirm():
        result["value"] = entry_var.get()
        dlg.destroy()

    tk.Button(
        dlg, text="CONFIRM", command=confirm,
        bg="#2563eb", fg="white", activebackground="#2563eb",
        activeforeground="white", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=14, height=1, takefocus=0
    ).pack(pady=(0, 20))

    dlg.bind("<Return>", lambda e: confirm())

    center_dialog(dlg, 400, 240)

    dlg.wait_window()

    return result["value"]

def prompt_count(label):
    """Asks how many of this element type exist. 0 (or Skip) skips the
    category entirely and moves on to the next one."""

    return styled_ask_count(f"How many {label} are in this yard?")


def prompt_name(label, index, total, existing_names):
    """Asks for the ID/name of one element. Falls back to an auto-generated
    name if left blank, and re-prompts on a duplicate ID."""

    while True:

        name = styled_ask_name(label, index, total)

        if name is None or not name.strip():
            name = f"{label.upper().replace(' ', '_')}_{index + 1}"
            log(f"No ID entered - using default: {name}")
        else:
            name = name.strip().upper()

        if name in existing_names:
            messagebox.showwarning(
                "Duplicate ID",
                f"'{name}' is already used. Please enter a different ID."
            )
            continue

        return name


def build_quantity_capture_list():

    """
    Ask operator how many MAIN / SHUNT / CALLING-ON signals exist.

    This is independent of the TOC workbook.
    """

    global signals

    signals = {}

    combined = []

    existing = set()

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    count = prompt_count("MAIN Signals")

    for i in range(count):

        name = prompt_name(
            "MAIN Signal",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {

            "type": "MAIN",

            "aspects": None,

            "menu": None,

            "RED": None,

            "YELLOW": None,

            "DOUBLE_YELLOW": None,

            "GREEN": None,

            "ROUTE_INDICATOR": None

        }

        combined.append(("MAIN", name))

    # =====================================================
    # SHUNT SIGNALS
    # =====================================================

    count = prompt_count("SHUNT Signals")

    for i in range(count):

        name = prompt_name(
            "SHUNT Signal",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {

            "type": "SHUNT",

            "menu": None,

            "state_indicator": None,

            "route_init": None,

            "initial_snapshot": None,

            "C1": None,

            "C2": None,

            "C3": None

        }

        combined.append(("SHUNT", name))

    # =====================================================
    # CALLING ON SIGNALS
    # =====================================================

    count = prompt_count("CALLING-ON Signals")

    for i in range(count):

        name = prompt_name(
            "CALLING-ON Signal",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {

            "type": "CALLING_ON",

            "menu": None,

            "YELLOW": None,

            "ROUTE_INIT": None

        }

        combined.append(("CALLING_ON", name))

    # =====================================================
    # POINTS
    # =====================================================

    count = prompt_count("POINTS")

    for i in range(count):
        name = prompt_name(
            "POINT",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {
            "type": "POINT",
            "menu": None,
            "normal": None,
            "reverse": None,
            "free": None,
            "coordinate": None
        }

        combined.append(("POINT", name))

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    count = prompt_count("CRANK HANDLES")

    for i in range(count):
        name = prompt_name(
            "CRANK HANDLE",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {
            "type": "CH",
            "menu": None,
            "IN": None,
            "OUT": None,
            "ECH": None,
            "FREE": None,
            "coordinate": None
        }

        combined.append(("CH", name))

    # =====================================================
    # LC GATES
    # =====================================================

    count = prompt_count("LC GATES")

    for i in range(count):
        name = prompt_name(
            "LC GATE",
            i,
            count,
            existing
        )

        existing.add(name)

        signals[name] = {
            "type": "LC",
            "menu": None,
            "IN": None,
            "OUT": None,
            "coordinate": None
        }

        combined.append(("LC", name))

    return combined

def start_quantity_capture():

    global capture_module
    global capture_master_list
    global capture_index
    global signals

    # -----------------------------------------------------
    # Ask before clearing existing signal coordinates
    # -----------------------------------------------------

    if signals:

        proceed = messagebox.askyesno(

            "Start New Recording",

            "This will clear all existing signal coordinates.\n\n"

            "Do you want to continue?"

        )

        if not proceed:

            log("Signal recording cancelled")

            return

    # -----------------------------------------------------
    # Fresh start
    # -----------------------------------------------------

    signals = {}

    undo_stack.clear()

    capture_master_list = build_quantity_capture_list()

    if not capture_master_list:

        log("No signals entered")

        return

    capture_index = 0

    capture_module = "MASTER"

    set_initial_stage_for_current()

    create_capture_overlay()

    root.iconify()

    log("Recording Signal Coordinates")

    log(f"Total Signals : {len(capture_master_list)}")

    master_next_capture()

# =========================================================================
# =========================================================================
#  SHARED CAPTURE DISPATCHER
# =========================================================================
# =========================================================================

last_space_time = 0
last_backspace_time = 0
last_digit_time = 0

def dispatch_save_point(x, y):

    global route_capture_index

    if capture_module == "MASTER":
        master_save_point(x, y)

    elif capture_module == "TRACK":
        capture_track_point(x, y)

    elif capture_module == "LEGACY":
        global capture_waiting, captured_point
        captured_point = [x, y]
        capture_waiting = False

def dispatch_undo():

    if capture_module == "MASTER":
        undo_last_capture()

    elif capture_module == "TRACK":
        undo_last_track_capture()

    elif capture_module == "ROUTE":
        pass

def on_press(key):

    global last_space_time, last_backspace_time, last_digit_time

    # ---------------- SPACE ----------------
    if key == pynput_keyboard.Key.space:

        now = time.time()

        if now - last_space_time < 0.4:
            log("SPACE ignored (pressed too soon after previous capture)")
            return

        last_space_time = now

        if capture_module not in ("MASTER", "ROUTE", "TRACK", "LEGACY"):
            return

        x, y = win32api.GetCursorPos()

        root.after(0, lambda: dispatch_save_point(x, y))

    # ---------------- BACKSPACE ----------------
    elif key == pynput_keyboard.Key.backspace:

        now = time.time()

        if now - last_backspace_time < 0.4:
            return

        last_backspace_time = now

        if capture_module not in ("MASTER", "TRACK"):
            return

        root.after(0, dispatch_undo)

    # ---------------- NUMBER KEYS ----------------
    else:

        char = getattr(key, "char", None)

        if char in ("2", "3", "4") and capture_module == "MASTER":

            if capture_index < len(capture_master_list):

                sig_type, signal = capture_master_list[capture_index]

                if sig_type == "MAIN" and record_stage == "ask_aspects":

                    now = time.time()

                    if now - last_digit_time < 0.4:
                        return

                    last_digit_time = now

                    root.after(0, lambda n=int(char): set_aspects_and_continue(signal, n))


# =========================================================
# LOG
# =========================================================

def log(msg):

    timestamp = time.strftime("%H:%M:%S")

    # =====================================================
    # WRITE TO CONSOLE / console.log
    # =====================================================

    print(
        f"[{timestamp}] {msg}",
        flush=True
    )

    # =====================================================
    # WRITE TO GUI LOG
    # =====================================================

    def write():

        try:

            log_text.config(
                state="normal"
            )

            log_text.insert(
                tk.END,
                f"[{timestamp}] {msg}\n"
            )

            log_text.see(
                tk.END
            )

            log_text.config(
                state="disabled"
            )

        except Exception:
            pass

    try:

        root.after(
            0,
            write
        )

    except Exception:
        pass

# =========================================================
# CLICK
# =========================================================

def click(point):

    if point is None:

        return False

    x, y = point

    win32api.SetCursorPos((x, y))

    time.sleep(0.4)

    pause_event.wait()

    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTDOWN,
        0,
        0
    )

    time.sleep(0.1)

    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTUP,
        0,
        0
    )

    time.sleep(0.4)

    return True

# =========================================================
# CAPTURE POINT (LEGACY COMPATIBILITY)
# =========================================================

def capture_point():
    """Wait for one global SPACE capture and return [x, y]."""
    global capture_module, capture_waiting, captured_point

    previous_module = capture_module
    capture_module = "LEGACY"
    capture_waiting = True
    captured_point = None

    try:
        while capture_waiting:
            root.update()
            time.sleep(0.05)

        if captured_point is None:
            return None

        return list(captured_point)
    finally:
        capture_module = previous_module


# =========================================================
# CAPTURE TRACK COORDINATES
# =========================================================

def capture_track_points():
    """Compatibility wrapper for the track capture button."""
    start_track_capture()

# =========================================================
# CREATE SIGNAL SETUP
# =========================================================

def create_setup():

    global signals

    if not lock_routes_data:

        messagebox.showwarning(
            "WARNING",
            "Please load TOC/RCC file first."
        )

        return

    signals = {}

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    total_main = simpledialog.askinteger(
        "MAIN SIGNALS",
        "How Many Main Signals?"
    )

    if total_main is None:
        return

    for i in range(total_main):

        # =================================================
        # NAME
        # =================================================

        name = simpledialog.askstring(
            "MAIN SIGNAL",
            f"Enter Main Signal Name {i+1}"
        )

        if not name:
            return

        # =================================================
        # ASPECTS
        # =================================================

        aspects = simpledialog.askinteger(
            "ASPECTS",
            f"{name}\n\nEnter Aspects (2/3/4)"
        )

        if not aspects:
            return

        signals[name] = {

            "type": "MAIN",

            "aspects": aspects,

            "menu": None,

            "RED": None,

            "YELLOW": None,

            "DOUBLE_YELLOW": None,

            "GREEN": None
        }

        # =================================================
        # MENU
        # =================================================

        root.iconify()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture MENU\n\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        # =================================================
        # RED
        # =================================================

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture RED\n\nPress SPACE"
        )

        signals[name]["RED"] = capture_point()

        # =================================================
        # YELLOW
        # =================================================

        if aspects >= 3:

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture YELLOW\n\nPress SPACE"
            )

            signals[name]["YELLOW"] = capture_point()

        # =================================================
        # DOUBLE YELLOW
        # =================================================

        if aspects == 4:

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture DOUBLE YELLOW\n\nPress SPACE"
            )

            signals[name]["DOUBLE_YELLOW"] = capture_point()

        # =================================================
        # GREEN
        # =================================================

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture GREEN\n\nPress SPACE"
        )

        signals[name]["GREEN"] = capture_point()

        root.deiconify()

    # =====================================================
    # CALLING ON
    # =====================================================

    total_calling = simpledialog.askinteger(
        "CALLING ON",
        "How Many Calling ON Signals?"
    )

    if total_calling is None:
        return

    for i in range(total_calling):

        name = simpledialog.askstring(
            "CALLING ON",
            f"Enter Calling ON Signal Name {i+1}"
        )

        if not name:
            return

        signals[name] = {

            "type": "CALLING_ON",

            "menu": None,

            "YELLOW": None
        }

        root.iconify()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture MENU\n\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture YELLOW\n\nPress SPACE"
        )

        signals[name]["YELLOW"] = capture_point()

        root.deiconify()

    # =====================================================
    # SHUNT
    # =====================================================

    total_shunt = simpledialog.askinteger(
        "SHUNT",
        "How Many Shunt Signals?"
    )

    if total_shunt is None:
        return

    for i in range(total_shunt):

        name = simpledialog.askstring(
            "SHUNT",
            f"Enter Shunt Signal Name {i+1}"
        )

        if not name:
            return

        signals[name] = {

            "type": "SHUNT",

            "menu": None,
        }

        root.iconify()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture MENU\n\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        root.deiconify()

    log("ALL SIGNALS RECORDED")


# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():

    wb = Workbook()

    # =====================================================
    # MAIN SHEET
    # =====================================================

    ws_main = wb.active

    ws_main.title = "MAIN_SIGNALS"

    ws_main.append([

        "Signal",

        "Aspects",

        "Menu_X",

        "Menu_Y",

        "Red_X",

        "Red_Y",

        "Yellow_X",

        "Yellow_Y",

        "DY_X",

        "DY_Y",

        "Green_X",

        "Green_Y"

    ])

    # =====================================================
    # CALLING ON
    # =====================================================

    ws_call = wb.create_sheet(
        "CALLING_ON"
    )

    ws_call.append([

        "Signal",

        "Menu_X",

        "Menu_Y",

        "Yellow_X",

        "Yellow_Y",

        "Route_Init_X",

        "Route_Init_Y"

    ])

    # =====================================================
    # SHUNT
    # =====================================================

    ws_shunt = wb.create_sheet(
        "SHUNT"
    )

    ws_shunt.append([

        "Signal",

        "Menu_X",

        "Menu_Y",

        "Route_Init_X",

        "Route_Init_Y"

    ])

    # =====================================================
    # WRITE DATA
    # =====================================================

    for sig, info in signals.items():

        typ = info["type"]

        # =================================================
        # MAIN
        # =================================================

        if typ == "MAIN":

            ws_main.append([

                sig,

                info["aspects"],

                info["menu"][0],
                info["menu"][1],

                info["RED"][0],
                info["RED"][1],

                (info.get("YELLOW") or [None, None])[0],
                (info.get("YELLOW") or [None, None])[1],

                (info.get("DOUBLE_YELLOW") or [None, None])[0],
                (info.get("DOUBLE_YELLOW") or [None, None])[1],

                info["GREEN"][0],
                info["GREEN"][1]

            ])

        # =================================================
        # CALLING ON
        # =================================================

        elif typ == "CALLING_ON":

            ws_call.append([

                sig,

                info["menu"][0],
                info["menu"][1],

                info["YELLOW"][0],
                info["YELLOW"][1],

                (info.get("ROUTE_INIT") or [None, None])[0],
                (info.get("ROUTE_INIT") or [None, None])[1]

            ])

        # =================================================
        # SHUNT
        # =================================================

        elif typ == "SHUNT":

            ws_shunt.append([

                sig,

                info["menu"][0],
                info["menu"][1],

                (info.get("route_init") or [None, None])[0],
                (info.get("route_init") or [None, None])[1]

            ])

    # =====================================================
    # SAVE FILE
    # =====================================================

    global config_file_path

    if not config_file_path:

        config_file_path = filedialog.asksaveasfilename(
            title="Save Signal Configuration",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if not config_file_path:
            return

    wb.save(config_file_path)

    log(f"CONFIG SAVED:\n{config_file_path}")

    messagebox.showinfo(

        "CONFIG SAVED",

        f"FILE SAVED SUCCESSFULLY\n\n{config_file_path}"
    )

    # =====================================================
    # OPEN FOLDER
    # =====================================================

    os.startfile(os.getcwd())
# =========================================================
# LOAD CONFIG
# =========================================================

def load_config():

    global signals
    global config_file_path

    if not lock_routes_data:
        messagebox.showwarning(
            "WARNING",
            "Please load TOC/RCC file first."
        )
        return

    file_path = filedialog.askopenfilename(
        title="Select Signal Configuration",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_path:
        return

    def find_sheets(path):
        """Return MAIN/CALLING/SHUNT sheet names for supported config formats."""
        try:
            test_wb = load_workbook(path, read_only=True, data_only=True)
            names = {str(x).strip().upper(): x for x in test_wb.sheetnames}

            main_name = names.get("MAIN") or names.get("MAIN_SIGNALS")
            calling_name = names.get("CAL") or names.get("CALLING_ON")
            shunt_name = names.get("SHUNT")

            test_wb.close()

            if main_name and calling_name and shunt_name:
                return main_name, calling_name, shunt_name
        except Exception:
            pass
        return None

    sheet_names = find_sheets(file_path)

    # If user selected the TOC/RCC workbook instead of the existing
    # signal configuration workbook, find the correct workbook in
    # the same folder automatically.
    if sheet_names is None:
        folder = os.path.dirname(file_path)

        try:
            for filename in os.listdir(folder):
                if not filename.lower().endswith(".xlsx"):
                    continue

                candidate = os.path.join(folder, filename)

                if os.path.abspath(candidate) == os.path.abspath(file_path):
                    continue

                candidate_sheets = find_sheets(candidate)

                if candidate_sheets is not None:
                    file_path = candidate
                    sheet_names = candidate_sheets
                    log(
                        f"SIGNAL CONFIGURATION FOUND : {filename}"
                    )
                    break
        except Exception as e:
            log(f"CONFIG SEARCH ERROR : {e}")

    if sheet_names is None:
        messagebox.showerror(
            "ERROR",
            "Signal configuration workbook not found.\n\n"
            "Required sheets are:\n"
            "MAIN + CAL + SHUNT\n"
            "or\n"
            "MAIN_SIGNALS + CALLING_ON + SHUNT"
        )
        log("SIGNAL CONFIGURATION WORKBOOK NOT FOUND")
        return

    config_file_path = file_path
    signals = {}

    wb = load_workbook(config_file_path, data_only=True)

    main_sheet, calling_sheet, shunt_sheet = sheet_names

    log(f"CONFIG SHEETS FOUND : {wb.sheetnames}")
    log(f"MAIN CONFIG SHEET : {main_sheet}")
    log(f"CALLING-ON CONFIG SHEET : {calling_sheet}")
    log(f"SHUNT CONFIG SHEET : {shunt_sheet}")

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    ws = wb[main_sheet]

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        signal_name = (
            str(row[0])
            .replace(".0", "")
            .strip()
            .upper()
        )

        signals[signal_name] = {
            "type": "MAIN",
            "aspects": row[1] if len(row) > 1 else 1,
            "menu": [
                row[2] if len(row) > 2 else None,
                row[3] if len(row) > 3 else None
            ],
            "RED": [
                row[4] if len(row) > 4 else None,
                row[5] if len(row) > 5 else None
            ],
            "YELLOW": (
                [row[6], row[7]]
                if len(row) > 7 and row[6] is not None
                else None
            ),
            "DOUBLE_YELLOW": (
                [row[8], row[9]]
                if len(row) > 9 and row[8] is not None
                else None
            ),
            "GREEN": [
                row[10] if len(row) > 10 else None,
                row[11] if len(row) > 11 else None
            ]
        }

    # =====================================================
    # CALLING-ON SIGNALS
    # =====================================================

    ws = wb[calling_sheet]

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        signal_name = (
            str(row[0])
            .replace(".0", "")
            .strip()
            .upper()
        )

        signals[signal_name] = {
            "type": "CALLING_ON",
            "menu": [
                row[1] if len(row) > 1 else None,
                row[2] if len(row) > 2 else None
            ],
            "YELLOW": [
                row[3] if len(row) > 3 else None,
                row[4] if len(row) > 4 else None
            ],
            "ROUTE_INIT": (
                [row[5], row[6]]
                if len(row) > 6 and row[5] is not None
                else None
            )
        }

    # =====================================================
    # SHUNT SIGNALS
    # =====================================================

    ws = wb[shunt_sheet]

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        signal_name = (
            str(row[0])
            .replace(".0", "")
            .strip()
            .upper()
        )

        signals[signal_name] = {
            "type": "SHUNT",
            "menu": [
                row[1] if len(row) > 1 else None,
                row[2] if len(row) > 2 else None
            ],
            "route_init": (
                [row[3], row[4]]
                if len(row) > 4 and row[3] is not None
                else None
            ),
            "state_indicator": None,
            "initial_snapshot": None,
            "C1": None,
            "C2": None,
            "C3": None
        }

    wb.close()

    main_count = sum(
        1 for value in signals.values()
        if value.get("type") == "MAIN"
    )
    calling_count = sum(
        1 for value in signals.values()
        if value.get("type") == "CALLING_ON"
    )
    shunt_count = sum(
        1 for value in signals.values()
        if value.get("type") == "SHUNT"
    )

    log(f"MAIN SIGNALS LOADED : {main_count}")
    log(f"CALLING-ON SIGNALS LOADED : {calling_count}")
    log(f"SHUNT SIGNALS LOADED : {shunt_count}")
    log(f"AVAILABLE SIGNALS : {list(signals.keys())}")
    log("CONFIG LOADED SUCCESSFULLY")

def save_new_signal_coordinates():

    global config_file_path

    if not signals:

        messagebox.showwarning(
            "WARNING",
            "No signal coordinates available.\nCreate signals first."
        )
        return

    file_path = filedialog.asksaveasfilename(
        title="Save Signal Configuration",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_path:
        return

    config_file_path = file_path

    save_config()

    log(f"Configuration saved : {config_file_path}")
# =========================================================
# TABLE
# =========================================================



def refresh_table():

    # Keep the existing Track Details table/UI.
    # Add MAIN, CALLING-ON and SHUNT rows to the same table
    # so the loaded TOC is visible completely.

    for item in tree.get_children():
        tree.delete(item)

    row_no = 1

    # =====================================================
    # MAIN ROUTES
    # =====================================================

    for item in lock_routes_data:

        signal = str(
            item.get("signal", "")
        ).strip()

        route = str(
            item.get("route", "")
        ).strip()

        controlled = str(
            item.get("track_data", "")
        ).strip()

        tree.insert(
            "",
            tk.END,
            values=(
                row_no,
                signal,
                route,
                controlled
            )
        )

        row_no += 1

    # =====================================================
    # CALLING-ON ROUTES
    # =====================================================

    for item in calling_on_data:

        signal = str(
            item.get("signal", "")
        ).strip()

        route = str(
            item.get("route", "")
        ).strip()

        controlled = str(
            item.get("track_data", "")
        ).strip()

        tree.insert(
            "",
            tk.END,
            values=(
                row_no,
                signal,
                route,
                controlled
            )
        )

        row_no += 1

    # =====================================================
    # SHUNT ROUTES
    # =====================================================

    for item in shunt_data:

        signal = str(
            item.get("signal", "")
        ).strip()

        route = str(
            item.get("route", "")
        ).strip()

        controlled = str(
            item.get("track_data", "")
        ).strip()

        tree.insert(
            "",
            tk.END,
            values=(
                row_no,
                signal,
                route,
                controlled
            )
        )

        row_no += 1

    # =====================================================
    # STATUS
    # =====================================================

    total = row_no - 1

    log(
        f"TABLE UPDATED : {total} routes "
        f"(MAIN={len(lock_routes_data)}, "
        f"CALLING-ON={len(calling_on_data)}, "
        f"SHUNT={len(shunt_data)})"
    )

def load_lock_routes():

    global lock_routes_data
    global calling_on_data
    global shunt_data
    global track_points

    file = filedialog.askopenfilename(
        title="Select TOC / RCC Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    try:
        wb = load_workbook(file, data_only=True)

        lock_routes_data = []
        calling_on_data = []
        shunt_data = []
        track_points = {}

        # =====================================================
        # FIND TOC SHEET
        # =====================================================
        #
        # Old code required the sheet name "TOC".
        # Your uploaded toc1.xlsx has "Sheet1".
        #
        # Therefore:
        #   1. Prefer TOC
        #   2. Otherwise use the first sheet containing
        #      SIGNAL + ROUTE headers
        # =====================================================

        toc_ws = None

        for ws in wb.worksheets:

            headers = []

            for cell in ws[1]:
                if cell.value is not None:
                    headers.append(
                        str(cell.value).strip().upper()
                    )

            if "SIGNAL" in headers and "ROUTE" in headers:

                if ws.title.strip().upper() == "TOC":
                    toc_ws = ws
                    break

                if toc_ws is None:
                    toc_ws = ws

        # =====================================================
        # OPTIONAL SEPARATE SHEETS
        # =====================================================

        calling_ws = None
        shunt_ws = None

        for ws in wb.worksheets:

            name = ws.title.strip().upper()

            if name == "CALLING_ON_SIGNALS":
                calling_ws = ws

            elif name == "SHUNT_SIGNALS":
                shunt_ws = ws

        if toc_ws is None:

            wb.close()

            messagebox.showerror(
                "ERROR",
                "TOC DATA NOT FOUND.\n\n"
                "The Excel file must contain columns:\n"
                "SIGNAL and ROUTE"
            )

            log("TOC DATA NOT FOUND")
            return

        log(
            f"TOC SHEET : {toc_ws.title}"
        )

        # =====================================================
        # READ HEADERS
        # =====================================================

        header_map = {}

        for col_index, cell in enumerate(
            toc_ws[1],
            start=1
        ):

            if cell.value is None:
                continue

            header = (
                str(cell.value)
                .strip()
                .upper()
                .replace(" ", "_")
            )

            header_map[header] = col_index

        signal_col = header_map.get("SIGNAL")
        route_col = header_map.get("ROUTE")

        controlled_col = (
            header_map.get("CONTROLLED_BY_TRACKS")
            or header_map.get("CONTROLLED_BY_TRACK")
            or header_map.get("TRACKS")
        )

        # Track to drop for CALLING-ON = TOC "Track" column
        approach_col = (
            header_map.get("TRACK")
            or header_map.get("APPROACH_TRACK")
            or header_map.get("APPROACH")
        )

        # -------------------------------------------------
        # SPL_CASE_CALLING_ON sheet (optional):
        # Main Signal | Route | Signal | Route | Track
        # Track = approach track to drop for that
        # CALLING-ON signal / route (e.g. 1C / 1C_A -> 1CXTPR)
        # -------------------------------------------------
        spl_approach = {}

        for ws in wb.worksheets:

            if ws.title.strip().upper() != "SPL_CASE_CALLING_ON":
                continue

            for r in ws.iter_rows(min_row=2, values_only=True):

                if not r or len(r) < 5:
                    continue

                if r[2] is None or r[3] is None or r[4] is None:
                    continue

                key = (
                    str(r[2]).replace(".0", "").strip().upper(),
                    str(r[3]).strip().upper()
                )

                spl_approach[key] = str(r[4]).strip().upper()

            log(f"SPL_CASE_CALLING_ON APPROACH TRACKS : {spl_approach}")

        if signal_col is None or route_col is None:

            wb.close()

            messagebox.showerror(
                "ERROR",
                f"Invalid TOC sheet: {toc_ws.title}\n\n"
                "SIGNAL and ROUTE columns are required."
            )

            return

        # =====================================================
        # HELPER
        # =====================================================

        def cell_value(row, column):

            if column is None:
                return ""

            index = column - 1

            if index >= len(row):
                return ""

            value = row[index]

            if value is None:
                return ""

            return str(value).strip()

        def normalize(value):

            return (
                str(value)
                .replace(".0", "")
                .strip()
                .upper()
            )

        # =====================================================
        # LOAD COMBINED TOC
        # =====================================================

        for row in toc_ws.iter_rows(
            min_row=2,
            values_only=True
        ):

            raw_signal = cell_value(
                row,
                signal_col
            )

            raw_route = cell_value(
                row,
                route_col
            )

            if not raw_signal or not raw_route:
                continue

            signal = normalize(raw_signal)
            route = normalize(raw_route)

            controlled_tracks = normalize(
                cell_value(row, controlled_col)
            )

            approach_track = normalize(
                cell_value(row, approach_col)
            )

            # -------------------------------------------------
            # SHUNT
            # -------------------------------------------------
            #
            # Example from your TOC:
            # 9SH | 9_A | 50BXT,50AXT | ...
            # -------------------------------------------------

            if (
                signal.startswith("SH")
                or signal.endswith("SH")
            ):

                shunt_data.append({

                    "signal": signal,

                    "route": route,

                    "track_data": controlled_tracks

                })

            # -------------------------------------------------
            # CALLING-ON
            # -------------------------------------------------
            #
            # Example from your TOC:
            # 1C | 1C_A | 25XT,... | 1CXTPR | ...
            # -------------------------------------------------

            elif signal.endswith("C"):

                # Track to drop: TOC "Track" column first; only if
                # that cell is empty, use SPL_CASE_CALLING_ON sheet.
                if not approach_track:
                    approach_track = spl_approach.get(
                        (signal, route),
                        ""
                    )

                log(f"CALLING-ON {signal} {route} TRACK : {approach_track}")

                calling_on_data.append({

                    "signal": signal,

                    "route": route,

                    "approach_track": approach_track,

                    "track_data": controlled_tracks

                })

            # -------------------------------------------------
            # MAIN
            # -------------------------------------------------

            else:

                lock_routes_data.append({

                    "signal": signal,

                    "route": route,

                    "track_data": controlled_tracks

                })

            # -------------------------------------------------
            # CREATE TRACK ENTRIES
            # -------------------------------------------------

            if controlled_tracks:

                for track in controlled_tracks.split(","):

                    track = track.strip()

                    if track:
                        track_points.setdefault(
                            track,
                            ()
                        )

            if approach_track:

                for track in approach_track.split(","):

                    track = track.strip()

                    if track:
                        track_points.setdefault(
                            track,
                            ()
                        )

        # =====================================================
        # OPTIONAL SEPARATE CALLING-ON SHEET
        # =====================================================

        if calling_ws is not None:

            for row in calling_ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row:
                    continue

                signal = normalize(
                    cell_value(row, 1)
                )

                route = normalize(
                    cell_value(row, 2)
                )

                if not signal or not route:
                    continue

                approach_track = normalize(
                    cell_value(row, 3)
                )

                controlled_tracks = normalize(
                    cell_value(row, 4)
                )

                calling_on_data.append({

                    "signal": signal,

                    "route": route,

                    "approach_track": approach_track,

                    "track_data": controlled_tracks

                })

        # =====================================================
        # OPTIONAL SEPARATE SHUNT SHEET
        # =====================================================

        if shunt_ws is not None:

            for row in shunt_ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row:
                    continue

                signal = normalize(
                    cell_value(row, 1)
                )

                route = normalize(
                    cell_value(row, 2)
                )

                if not signal or not route:
                    continue

                track_data = normalize(
                    cell_value(row, 3)
                )

                shunt_data.append({

                    "signal": signal,

                    "route": route,

                    "track_data": track_data

                })

        wb.close()

        # =====================================================
        # LOG RESULTS
        # =====================================================

        log(
            f"TOC SHEET USED : {toc_ws.title}"
        )

        log(
            f"MAIN ROUTES : {len(lock_routes_data)}"
        )

        log(
            f"CALLING-ON ROUTES : {len(calling_on_data)}"
        )

        log(
            f"SHUNT ROUTES : {len(shunt_data)}"
        )

        log(
            f"TRACKS FOUND : {len(track_points)}"
        )

        # =====================================================
        # REFRESH UI
        # =====================================================

        refresh_table()

        log("TOC LOADED SUCCESSFULLY")

    except Exception as e:

        try:
            wb.close()
        except Exception:
            pass

        log(
            f"TOC LOAD ERROR : {e}"
        )

        messagebox.showerror(
            "TOC LOAD ERROR",
            f"Unable to load TOC/RCC file.\n\n{e}"
        )


# =========================================================
# MENU ITEM
# =========================================================

def click_menu_item(name):

    try:

        item = auto.MenuItemControl(
            searchDepth=15,
            Name=name
        )

        if item.Exists(3):

            pause_event.wait()

            item.Click(simulateMove=False)

            # Give the route time to be selected
            time.sleep(0.5)

            # Click on an empty area to close the popup menu
            #pyautogui.click(20, 20)

            time.sleep(0.5)

            return True

    except Exception as e:
        log(f"MENU CLICK ERROR : {e}")

    return False

# =========================================================
# AVG COLOR
# =========================================================

def get_avg_color(x, y, screenshot):

    pixels = []

    for dx in range(-2, 3):

        for dy in range(-2, 3):

            pixels.append(
                screenshot.getpixel(
                    (x + dx, y + dy)
                )
            )

    r = sum(p[0] for p in pixels) // len(pixels)

    g = sum(p[1] for p in pixels) // len(pixels)

    b = sum(p[2] for p in pixels) // len(pixels)

    return r, g, b

# =========================================================
# YELLOW DETECT
# =========================================================

def is_yellow(point):

    if point is None:
        return False

    screenshot = grab_all_screens()

    x, y = point

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    return r > 150 and g > 150

# =========================================================
# MAIN SIGNAL CHANGED
# =========================================================

def is_signal_changed(signal):

    if signal not in signals:
        log(f"{signal} NOT IN CONFIG")
        return False  # or treat as PASS safely

    info = signals[signal]

    if info["type"] != "MAIN":
        log(f"{signal} is not MAIN -> skipping RED check")
        return False

    if "RED" not in info or info["RED"] is None:
        log(f"{signal} has no RED point configured")
        return False

    screenshot = grab_all_screens()

    red_point = info["RED"]

    r, g, b = get_avg_color(
        red_point[0],
        red_point[1],
        screenshot
    )

    log(f"{signal} AFTER RED RGB = {r},{g},{b}")

    is_red = (r > 140 and g < 130 and b < 130)

    return not is_red




# =========================================================
# TRACK ACTIVE DETECTION
# =========================================================

def is_track_active(point):

    if point is None:
        log("TRACK POINT IS NONE")
        return False

    x, y = point

    log(
        f"============================================================"
    )

    log(
        f"TRACK CHECK COORDINATE : "
        f"X={x}, Y={y}"
    )

    screenshot = grab_all_screens()

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    log(
        f"TRACK RGB AT ({x},{y}) = "
        f"({r},{g},{b})"
    )

    if r > 140 and g > 140 and b < 120:

        log(
            f"TRACK RESULT : PASS "
            f"(YELLOW DETECTED)"
        )

        return True

    log(
        f"TRACK RESULT : FAIL "
        f"(YELLOW NOT DETECTED)"
    )

    return False

# =========================================================
# SHUNT STATUS
# =========================================================


# =========================================================
# PARSE SIGNAL NAME
# =========================================================

def parse_signal_name(lock_route):

    lr = str(lock_route).strip().upper()

    # ============================================
    # HAS "-"
    # ============================================

    if "-" in lr:

        signal_part = lr.split("-")[0]

    else:

        # LAST LETTER IS ROUTE
        signal_part = lr[:-1]

    return signal_part.strip()
# =========================================================
# DROP TRACK
# =========================================================

# =========================================================
# FIND AND CLICK UI
# =========================================================

def find_and_click(name, control_type=None):

    desktop = Desktop(backend="uia")

    windows = desktop.windows()

    for win in windows:

        try:

            descendants = win.descendants()

            for item in descendants:

                try:

                    text = item.window_text().strip().upper()

                    if text == name.upper():

                        if control_type:

                            current_type = str(
                                item.element_info.control_type
                            )

                            if control_type not in current_type:
                                continue

                        rect = item.rectangle()

                        x = rect.left + 10
                        y = rect.top + 10

                        pyautogui.click(x, y)

                        log(f"Clicked : {name}")

                        return True

                except:
                    pass

        except:
            pass

    return False

# =========================================================
# OPEN BIT CHART
# =========================================================

def open_bit_chart():

    pyautogui.click(500, 500)

    time.sleep(1)

    pyautogui.hotkey("ctrl", "b")

    log("CTRL+B Pressed")

    time.sleep(3)

    found = find_and_click(
        "50051",
        control_type="Text"
    )

    if not found:

        log("50051 Not Found")

        return False

    time.sleep(1)

    ok = find_and_click(
        "OK",
        control_type="Button"
    )

    if not ok:

        log("OK Button Not Found")

        return False

    time.sleep(3)

    return True

# =========================================================
# DROP TRACK
# =========================================================

STATION_BIT_KEY = "50051"


def _is_station_bit_title(title):
    t = str(title or "").upper()
    return STATION_BIT_KEY in t and "INDICATION" in t


def _find_station_bit_window(tries=6):
    """After CTRL+B -> 50051 -> OK the Station Bit window
    'Station 50051: (Indications | Controls)' can open on EITHER
    screen - as a top-level window or inside the Test Panel.
    Search both places, on all monitors."""
    for attempt in range(tries):
        try:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                try:
                    if _is_station_bit_title(win.window_text()):
                        return win
                    for child in win.descendants(control_type="Window"):
                        if _is_station_bit_title(child.window_text()):
                            return child
                except Exception:
                    pass
        except Exception as e:
            log(f"STATION BIT SEARCH ERROR : {e}")
        time.sleep(1)
    return None


def _rect_ok(rect):
    """True if the rectangle is a real on-screen area (any monitor)."""
    try:
        vx = win32api.GetSystemMetrics(76)
        vy = win32api.GetSystemMetrics(77)
        vw = win32api.GetSystemMetrics(78)
        vh = win32api.GetSystemMetrics(79)
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
        return (rect.width() > 0 and rect.height() > 0
                and vx <= cx < vx + vw and vy <= cy < vy + vh)
    except Exception:
        return False


def _find_in_window(win, names, control_type=None):
    """Find a control by text INSIDE the given window only
    (never in the taskbar or any other window)."""
    names = [n.upper() for n in names]
    for nm in names:
        for item in win.descendants():
            try:
                if item.window_text().strip().upper() != nm:
                    continue
                if control_type and control_type not in str(item.element_info.control_type):
                    continue
                return item
            except Exception:
                pass
    return None


def _click_element(item, label):
    """Click an element on whichever monitor it is on.
    Scrolls it into view first; refuses to click a bad rectangle
    (that is what was hitting the taskbar Widgets button)."""
    try:
        item.iface_scroll_item.ScrollIntoView()
        time.sleep(0.3)
    except Exception:
        pass

    rect = item.rectangle()
    log(f"{label} RECT : L={rect.left} T={rect.top} R={rect.right} B={rect.bottom}")

    if not _rect_ok(rect):
        log(f"{label} NOT VISIBLE ON ANY SCREEN - CLICK SKIPPED")
        return False

    x = (rect.left + rect.right) // 2
    y = (rect.top + rect.bottom) // 2

    click((x, y))

    log(f"Clicked : {label} at ({x},{y})")
    return True


def drop_track(track):
    """Drop / restore a track - same sequence as signal_clearance.py
    (CTRL+B -> 50051 -> OK -> track -> TRANSMIT -> CANCEL), but after OK
    the Station Bit window is located on whichever screen it opened,
    brought to the front, and the track / TRANSMIT / CANCEL are looked
    up ONLY inside that window."""
    pause_event.wait()

    track = str(track).strip().upper()

    if not track:
        log("NO APPROACH TRACK GIVEN - DROP SKIPPED")
        return False

    log(f"DROPPING TRACK : {track}")

    # 1. CTRL+B -> 50051 -> OK   (unchanged, same as reference)
    opened = open_bit_chart()

    if not opened:
        return False

    # 2. LOCATE STATION BIT WINDOW ON EITHER SCREEN
    bit_win = _find_station_bit_window()

    if bit_win is None:
        log("STATION BIT WINDOW NOT FOUND ON ANY SCREEN")
        return False

    r = bit_win.rectangle()
    log(f"STATION BIT WINDOW : {bit_win.window_text()} "
        f"at L={r.left} T={r.top} R={r.right} B={r.bottom}")

    # Bring it to the front so the first click selects the bit
    # instead of only activating the window.
    try:
        bit_win.set_focus()
    except Exception:
        pass
    time.sleep(0.5)

    # 3. SELECT THE TRACK BIT (inside the Station Bit window only)
    names = [track]
    if not track.endswith("PR"):
        names.append(track + "PR")

    item = _find_in_window(bit_win, names)

    if item is None:
        log(f"{track} Not Found in Station Bit window")
        return False

    selected = False
    try:
        item.iface_selection_item.Select()
        time.sleep(0.5)
        selected = bool(item.iface_selection_item.CurrentIsSelected)
    except Exception:
        pass

    if not selected:
        if not _click_element(item, track):
            return False
        time.sleep(0.5)
        try:
            selected = bool(item.iface_selection_item.CurrentIsSelected)
        except Exception:
            selected = True

    log(f"{track} SELECTED : {selected}")

    if not selected:
        log(f"{track} NOT SELECTED - TRANSMIT SKIPPED")
        return False

    time.sleep(1)

    # 4. TRANSMIT
    transmit = _find_in_window(bit_win, ["TRANSMIT"], control_type="Button")

    if transmit is None or not _click_element(transmit, "TRANSMIT"):
        log("TRANSMIT Not Found")
        return False

    time.sleep(1)

    # 5. CANCEL
    cancel = _find_in_window(bit_win, ["CANCEL"], control_type="Button")

    if cancel is None or not _click_element(cancel, "CANCEL"):
        log("CANCEL Not Found")
        return False

    log(f"{track} DROP/UP SUCCESS")

    time.sleep(5)

    return True

# =========================================================
# DETECT SIGNAL TYPE
# =========================================================

def detect_signal_type(lock_route):

    lr = lock_route.strip().upper()

    # REMOVE ROUTE PART
    # 25C-J -> 25C
    # 25-J -> 25

    left = lr.split("-")[0]

    # CALLING ON
    # ENDS WITH C

    if left.endswith("C"):

        signal_name = left

        signal_type = "CALLING_ON"

    else:

        signal_name = left

        signal_type = "MAIN"

    return signal_name, signal_type
# =========================================================
# GET TRACK FOR LOCK ROUTE
# =========================================================

def get_track_for_route(lock_route, track_data):

    routes = [

        x.strip().upper()

        for x in lock_route.split(",")

    ]

    tracks = [

        x.strip().upper()

        for x in track_data.split(",")

    ]

    mapping = {}

    track_index = 0

    for route in routes:

        signal_name, signal_type = detect_signal_type(route)

        # ONLY CALLING ON GETS TRACK

        if signal_type == "CALLING_ON":

            if track_index < len(tracks):

                mapping[route] = tracks[track_index]

                track_index += 1

        else:

            mapping[route] = None

    return mapping
def create_report():

    # =====================================================
    # CREATE WORKBOOK
    # =====================================================

    wb = Workbook()

    ws = wb.active
    ws.title = "VISUAL INSPECTION FOR TRACK REPORT"

    # =====================================================
    # FORMATTING
    # =====================================================

    thin = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # Light red ONLY for failed rows
    fail_fill = PatternFill(
        "solid",
        fgColor="F4CCCC"
    )

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title = ws["A1"]

    title.value = "VISUAL INSPECTION FOR TRACK REPORT"

    title.font = Font(
        bold=True,
        size=16
    )

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 30

    # =====================================================
    # GENERATED TIME
    # =====================================================

    ws.merge_cells("A2:D2")

    generated = ws["A2"]

    generated.value = (
        "Generated : "
        + datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    )

    generated.font = Font(
        bold=True,
        size=10
    )

    generated.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # =====================================================
    # TABLE HEADERS
    # =====================================================
    # IMPORTANT:
    # TEST COLUMN REMOVED
    #
    # A = SIGNAL
    # B = ROUTE
    # C = PASS/FAIL
    # D = DATE & TIME
    # =====================================================

    headers = [
        "SIGNAL",
        "ROUTE",
        "PASS/FAIL",
        "DATE & TIME"
    ]

    for col, header in enumerate(headers, start=1):

        cell = ws.cell(
            row=3,
            column=col,
            value=header
        )

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    ws.row_dimensions[3].height = 24

    # =====================================================
    # COLLECT ROUTE RESULTS
    # =====================================================

    route_summary = {}

    for item in report_rows:

        signal = str(
            item.get(
                "SIGNAL",
                ""
            )
        )

        route = str(
            item.get(
                "ROUTE",
                ""
            )
        )

        # -------------------------------------------------
        # Get result
        # Supports RESULT or result
        # -------------------------------------------------

        result = str(
            item.get(
                "RESULT",
                item.get(
                    "result",
                    "FAIL"
                )
            )
        ).upper()

        # -------------------------------------------------
        # Get date/time
        # -------------------------------------------------

        date_time = item.get(
            "DATE & TIME",
            item.get(
                "date_time",
                datetime.now().strftime(
                    "%d-%m-%Y %H:%M:%S"
                )
            )
        )

        key = (
            signal,
            route
        )

        # -------------------------------------------------
        # First occurrence
        # -------------------------------------------------

        if key not in route_summary:

            route_summary[key] = {
                "signal": signal,
                "route": route,
                "result": result,
                "date_time": date_time
            }

        # -------------------------------------------------
        # If any inspection for this route FAILS,
        # complete route is considered FAIL
        # -------------------------------------------------

        if result == "FAIL":

            route_summary[key]["result"] = "FAIL"

            route_summary[key]["date_time"] = date_time

    # =====================================================
    # WRITE ROUTE RESULTS
    # =====================================================

    for key, data in route_summary.items():

        signal = data["signal"]
        route = data["route"]
        result = data["result"]
        date_time = data["date_time"]

        ws.append([
            signal,
            route,
            result,
            date_time
        ])

        row_number = ws.max_row

        # -------------------------------------------------
        # BORDER + ALIGNMENT
        # -------------------------------------------------

        for col in range(1, 5):

            cell = ws.cell(
                row=row_number,
                column=col
            )

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # -------------------------------------------------
        # FAILED ROUTE = RED
        # PASSED ROUTE = NORMAL
        # -------------------------------------------------

        if result == "FAIL":

            for col in range(1, 5):

                ws.cell(
                    row=row_number,
                    column=col
                ).fill = fail_fill

            # Red bold PASS/FAIL
            ws.cell(
                row=row_number,
                column=3
            ).font = Font(
                bold=True,
                color="FF0000"
            )

        else:

            ws.cell(
                row=row_number,
                column=3
            ).font = Font(
                bold=True
            )

    # =====================================================
    # ROUTE SUMMARY
    # =====================================================

    total_routes = len(route_summary)

    initiated_route_signals = []

    non_initiated_route_signals = []

    # -----------------------------------------------------
    # Keep the same order in which routes appeared
    # -----------------------------------------------------

    for key, data in route_summary.items():

        signal = str(
            data["signal"]
        )

        result = str(
            data["result"]
        ).upper()

        if result == "PASS":

            initiated_route_signals.append(
                signal
            )

        else:

            non_initiated_route_signals.append(
                signal
            )

    # =====================================================
    # CONVERT SIGNAL LISTS TO TEXT
    # =====================================================

    if initiated_route_signals:

        initiated_routes_text = ",".join(
            initiated_route_signals
        )

    else:

        initiated_routes_text = "0"

    if non_initiated_route_signals:

        non_initiated_routes_text = ",".join(
            non_initiated_route_signals
        )

    else:

        non_initiated_routes_text = "0"

    # =====================================================
    # SUMMARY START POSITION
    # =====================================================

    summary_start = ws.max_row + 3

    # =====================================================
    # SUMMARY
    # =====================================================

    summary_data = [

        (
            "TOTAL ROUTES",
            total_routes
        ),

        (
            "INITIATED ROUTES",
            initiated_routes_text
        ),

        (
            "NON-INITIATED ROUTES",
            non_initiated_routes_text
        )

    ]

    for index, (label, value) in enumerate(
        summary_data,
        start=summary_start
    ):

        # -------------------------------------------------
        # LABEL
        # -------------------------------------------------

        label_cell = ws.cell(
            row=index,
            column=1,
            value=label
        )

        label_cell.font = Font(
            bold=True
        )

        label_cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        label_cell.border = border

        # -------------------------------------------------
        # VALUE
        # -------------------------------------------------

        value_cell = ws.cell(
            row=index,
            column=2,
            value=value
        )

        value_cell.font = Font(
            bold=True
        )

        value_cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        value_cell.border = border

        # -------------------------------------------------
        # Make summary rows clean
        # -------------------------------------------------

        ws.row_dimensions[index].height = 22

    # =====================================================
    # COLUMN WIDTHS
    # =====================================================

    ws.column_dimensions["A"].width = 18

    ws.column_dimensions["B"].width = 18

    ws.column_dimensions["C"].width = 18

    ws.column_dimensions["D"].width = 24

    # =====================================================
    # FREEZE HEADER
    # =====================================================

    ws.freeze_panes = "A4"

    # =====================================================
    # SAVE REPORT
    # =====================================================

    wb.save(REPORT_FILE)

    log(
        f"REPORT SAVED : {REPORT_FILE}"
    )


def check_track_status(track):

    if track not in track_points:
        log(f"{track} NOT FOUND")
        return "FAIL"

    x, y = map(int, track_points[track])

    log(f"CHECKING {track} AT X={x} Y={y}")

    SEARCH = 8     # Search 20 pixels around the stored point

    img = grab_screen_region(
        x - SEARCH, y - SEARCH,
        SEARCH * 2 + 1,
        SEARCH * 2 + 1
    )

    from datetime import datetime
    img.save(f"debug_{track}_{datetime.now().strftime('%H%M%S')}.png")

    red_found = False

    width = SEARCH * 2 + 1
    height = SEARCH * 2 + 1

    for px in range(width):
        for py in range(height):

            r, g, b = img.getpixel((px, py))

            # Detect any red pixel
            if r > 140 and g < 100 and b < 100:
                red_found = True
                break

        if red_found:
            break

    if red_found:
        log(f"{track} -> RED FOUND")
        return "FAIL"

    log(f"{track} -> NO RED")
    return "PASS"



def check_signal_status(main_signal):

    if main_signal not in signals:
        return "FAIL"

    sig = signals[main_signal]

    x = sig["RED"][0]
    y = sig["RED"][1]

    img = grab_screen_region(x, y, 1, 1)
    r, g, b = img.getpixel((0, 0))

    log(f"{main_signal} SIGNAL RGB = ({r},{g},{b})")

    # Signal RED
    if r > 180 and g < 120 and b < 120:
        return "FAIL"

    return "PASS"


def toggle_track(track_name):
    """Drop / restore a track through the Station Bit window.

    Uses the same method as signal_clearance.py (cal_drop_track):
    CTRL+B -> 50051 -> OK -> track -> TRANSMIT -> CANCEL, where every
    control is searched across ALL desktop windows (pywinauto) instead
    of only inside the "Test Panel" window.  This works whether the
    simulator / Station Bit window is on the first or second screen.
    """

    try:

        track_name = str(track_name).strip().upper()

        log(f"TOGGLING {track_name}")

        ok = drop_track(track_name)

        if ok:
            log(f"{track_name} TOGGLED")
        else:
            log(f"TOGGLE FAILED : {track_name}")

        return ok

    except Exception as e:

        log(f"TOGGLE FAILED : {track_name} : {e}")
        return False



# =========================================================
# SIGNAL CONFIG RESOLVER
# =========================================================

def resolve_signal_config(signal, expected_type=None):
    """Resolve a TOC signal to its existing signal configuration."""
    signal = str(signal).strip().upper()
    candidates = [signal]

    # Shunt TOC often uses 9SH/17SH while the existing
    # signal configuration uses 9/17.
    if expected_type == "SHUNT" and signal.endswith("SH"):
        candidates.append(signal[:-2])

    for candidate in candidates:
        if candidate in signals:
            info = signals[candidate]
            if expected_type is None or str(info.get("type", "")).upper() == expected_type:
                return candidate, info

    return None, None


# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():

    global running
    global report_rows

    report_rows = []

    root.iconify()

    time.sleep(2)

    for row in lock_routes_data:

        pause_event.wait()

        if not running:
            return

        main_signal = str(row["signal"]).replace(".0", "").strip().upper()

        main_route = str(row["route"]).strip().upper()

        track_data = row["track_data"]

        log(f"SETTING {main_signal} -> {main_route}")

        if main_signal not in signals:

            log(f"{main_signal} NOT FOUND IN CONFIG")

            continue

        # =========================================================
        # IMPORTANT:
        # lock_routes_data contains ALL TOC rows (MAIN,
        # CALLING_ON and SHUNT).  The main-route loop must process
        # ONLY MAIN signals.  Otherwise 1C/2C/... are incorrectly
        # treated as normal main routes and the real Calling-On
        # loop is never reached correctly.
        # =========================================================

        signal_type = str(
            signals[main_signal].get("type", "MAIN")
        ).strip().upper()

        if signal_type != "MAIN":

            log(
                f"SKIPPING {main_signal} -> {main_route} "
                f"({signal_type}) - WILL BE TESTED IN ITS OWN LOOP"
            )

            continue

        # ============================================
        # SET ROUTE
        # ============================================

        log("STEP 1 - Before click(menu)")
        click(signals[main_signal]["menu"])

        log("STEP 2 - Before click_menu_item")
        click_menu_item(main_route.replace("-", "_"))

        log("STEP 3 - Route selected")


        # ============================================
        # CHECK TRACKS
        # ============================================

        #pyautogui.click(10, 10)  # click on empty black area
        time.sleep(0.5)

        tracks = [
            t.strip().upper()
            for t in track_data.split(",")
            if t.strip()
        ]

        log(f"Expected Tracks : {tracks}")

        time.sleep(2)
        grab_all_screens().save(f"FULL_SCREEN_{main_route}.png")

        for track in tracks:

            pause_event.wait()

            if not running:
                return

            log(f"CHECKING {track}")

            # ==========================================
            # VERIFY
            # ==========================================

            if track not in track_points:

                log(f"{track} COORDINATE NOT FOUND")

                result = "FAIL"

            else:

                if is_track_active(track_points[track]):
                    result = "PASS"
                else:
                    result = "FAIL"

            report_rows.append({

                "SIGNAL": main_signal,

                "ROUTE": main_route,

                "TRACK": track,

                "RESULT": result,

                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            })

            log(f"{track} -> {result}")


        # ============================================
        # SIGNAL CANCEL
        # ============================================

        click(signals[main_signal]["menu"])

        time.sleep(1)

        click_menu_item("Signal Cancel")

        log(f"{main_signal} SIGNAL CANCEL DONE")

        time.sleep(2)

        # ============================================
        # ROUTE RELEASE
        # ============================================

        click(signals[main_signal]["menu"])

        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{main_signal} ROUTE RELEASE DONE")

        time.sleep(6)

    # =========================================================
    # CALLING-ON ROUTES
    # IMPORTANT: this loop is intentionally separate from the
    # MAIN loop.  It MUST break the approach track BEFORE opening
    # the signal menu and selecting the Calling-On route.
    # =========================================================

    log("============================================================")
    log("CALLING-ON VISUAL INSPECTION FOR TRACK STARTED")
    log(f"CALLING-ON ROUTES : {len(calling_on_data)}")
    log("============================================================")

    for row in calling_on_data:

        pause_event.wait()

        if not running:
            return

        signal = str(row["signal"]).replace(".0", "").strip().upper()
        route = str(row["route"]).strip().upper()
        approach_track = str(row["approach_track"]).strip().upper()
        track_data = row["track_data"]

        if signal not in signals:
            log(f"{signal} NOT FOUND IN CONFIG")
            continue

        log(f"SETTING CALLING ROUTE {signal} -> {route}")

        # ------------------------------------------------------------
        # STEP 1: BREAK APPROACH TRACK
        # ------------------------------------------------------------
        # The track break needs to happen BEFORE the
        # CALLING-ON route is selected.
        log(
            f"STEP 1 - BREAKING APPROACH TRACK : "
            f"{approach_track}"
        )

        toggle_track(approach_track)

        time.sleep(5)

        log(
            f"STEP 1 COMPLETE - "
            f"{approach_track} BREAK COMMAND SENT"
        )

        # ------------------------------------------------------------
        # STEP 2: OPEN CALLING-ON SIGNAL MENU
        # ------------------------------------------------------------
        if not click(signals[signal]["menu"]):
            log(f"{signal} MENU CLICK FAILED")
            toggle_track(approach_track)
            time.sleep(2)
            continue

        time.sleep(1)

        # ------------------------------------------------------------
        # STEP 3: SELECT CALLING-ON ROUTE
        # IMPORTANT: CALLING-ON menu items use the route name as
        # supplied by CALLING_ON_SIGNALS (for example 1C_A).
        # Do NOT change '-' to '_' here.
        # ------------------------------------------------------------
        log(f"SELECTING CALLING-ON ROUTE : {route}")

        route_selected = click_menu_item(route)

        # Fallback for installations whose menu uses '_' instead
        # of '-' in the displayed route name.
        if not route_selected and "-" in route:
            fallback_route = route.replace("-", "_")
            log(f"EXACT ROUTE NOT FOUND - TRYING : {fallback_route}")
            route_selected = click_menu_item(fallback_route)

        if not route_selected:
            log(f"CALLING-ON ROUTE NOT SELECTED : {signal} -> {route}")

            # Restore the approach track before moving to the next
            # calling-on route.
            toggle_track(approach_track)
            time.sleep(2)
            continue

        log(f"CALLING-ON ROUTE SELECTED : {signal} -> {route}")

        # Give the interlocking enough time to establish the route.
        time.sleep(8)

        tracks = [
            t.strip().upper()
            for t in track_data.split(",")
            if t.strip()
        ]

        log(f"Expected Tracks : {tracks}")

        time.sleep(2)
        grab_all_screens().save(f"FULL_SCREEN_{route}.png")

        for track in tracks:

            pause_event.wait()

            if not running:
                return

            log(f"CHECKING {track}")

            # ==========================================
            # VERIFY
            # ==========================================

            if track not in track_points:

                log(f"{track} COORDINATE NOT FOUND")

                result = "FAIL"

            else:

                if is_track_active(track_points[track]):
                    result = "PASS"
                else:
                    result = "FAIL"

            report_rows.append({

                "SIGNAL": signal,

                "ROUTE": route,

                "TRACK": track,

                "RESULT": result,

                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            })

            log(f"{track} -> {result}")

        # ============================================
        # SIGNAL CANCEL
        # ============================================

        click(signals[signal]["menu"])

        time.sleep(1)

        click_menu_item("Signal Cancel")

        log(f"{signal} SIGNAL CANCEL DONE")

        time.sleep(2)

        # ============================================
        # ROUTE RELEASE
        # ============================================

        click(signals[signal]["menu"])

        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{signal} ROUTE RELEASE DONE")

        time.sleep(3)

        toggle_track(approach_track)

        time.sleep(2)


    # =========================================================
    # SHUNT ROUTE TEST
    # =========================================================

    log("============================================================")
    log("SHUNT VISUAL INSPECTION FOR TRACK STARTED")
    log(f"SHUNT ROUTES : {len(shunt_data)}")
    log("============================================================")

    for row in shunt_data:

        pause_event.wait()

        if not running:
            return

        signal = str(row["signal"]).replace(".0", "").strip().upper()
        route = str(row["route"]).strip().upper()
        track_data = row["track_data"]

        # Resolve 9SH -> 9, 17SH -> 17 when required.
        config_signal, signal_info = resolve_signal_config(signal, "SHUNT")

        if config_signal is None:
            log(f"{signal} NOT FOUND IN SHUNT CONFIG")
            report_rows.append({
                "SIGNAL": signal,
                "ROUTE": route,
                "TRACK": "CONFIGURATION",
                "RESULT": "FAIL",
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })
            continue

        log(f"SETTING SHUNT ROUTE {signal} -> {route}")

        if config_signal != signal:
            log(f"SHUNT CONFIG MAPPED : {signal} -> {config_signal}")

        menu_point = signal_info.get("menu")

        if not menu_point:
            log(f"{signal} SHUNT MENU COORDINATE NOT FOUND")
            report_rows.append({
                "SIGNAL": signal,
                "ROUTE": route,
                "TRACK": "MENU COORDINATE",
                "RESULT": "FAIL",
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })
            continue

        # ---------------------------------------------------------
        # SELECT SHUNT ROUTE
        # ---------------------------------------------------------
        click(menu_point)
        time.sleep(1)

        menu_route = route.replace("-", "_")
        route_selected = click_menu_item(menu_route)

        if not route_selected:
            log(f"SHUNT ROUTE NOT SELECTED : {signal} -> {route}")
            report_rows.append({
                "SIGNAL": signal,
                "ROUTE": route,
                "TRACK": "ROUTE SELECTION",
                "RESULT": "FAIL",
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })
            continue

        log(f"SHUNT ROUTE SELECTED : {signal} -> {route}")
        time.sleep(5)

        # ---------------------------------------------------------
        # SHUNT ROUTE INITIATION
        # ---------------------------------------------------------

        route_init_point = signal_info.get("route_init")

        if not route_init_point:
            log(f"{signal} SHUNT ROUTE INITIATION COORDINATE NOT FOUND")
            continue

        log(f"STEP 3 - SHUNT ROUTE INITIATION : {signal}")

        click(route_init_point)

        log(f"{signal} SHUNT ROUTE INITIATION CLICK DONE")

        time.sleep(5)

        tracks = [
            t.strip().upper()
            for t in str(track_data).split(",")
            if t.strip()
        ]

        log(f"Expected Tracks : {tracks}")

        time.sleep(2)
        grab_all_screens().save(f"FULL_SCREEN_{route}.png")

        for track in tracks:

            pause_event.wait()

            if not running:
                return

            log(f"CHECKING {track}")

            if track not in track_points:
                log(f"{track} COORDINATE NOT FOUND")
                result = "FAIL"
            else:
                result = "PASS" if is_track_active(track_points[track]) else "FAIL"

            report_rows.append({
                "SIGNAL": signal,
                "ROUTE": route,
                "TRACK": track,
                "RESULT": result,
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            log(f"{track} -> {result}")

        # ---------------------------------------------------------
        # SIGNAL CANCEL
        # ---------------------------------------------------------
        click(menu_point)
        time.sleep(1)
        click_menu_item("Signal Cancel")
        log(f"{signal} SIGNAL CANCEL DONE")
        time.sleep(2)

        # ---------------------------------------------------------
        # ROUTE RELEASE
        # ---------------------------------------------------------
        click(menu_point)
        time.sleep(1)
        click_menu_item("Route Release")
        log(f"{signal} ROUTE RELEASE DONE")
        time.sleep(4)


    running = False

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    create_report()

    root.deiconify()

    log("AUTOMATION COMPLETED")

    try:
        os.startfile(REPORT_FILE)

    except:
        pass




def create_excel():

    wb = Workbook()
    ws = wb.active
    ws.title = "TRACK COORDINATES"

    ws.merge_cells("A1:C1")
    ws["A1"] = "TRACK COORDINATES"
    ws["A1"].font = Font(
        bold=True,
        size=16,
        color="FFFFFF"
    )
    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="1E293B"
    )
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = ["TRACK NAME", "X COORDINATE", "Y COORDINATE"]
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill(
        "solid",
        fgColor="2563EB"
    )

    thin = Side(style="thin", color="000000")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    wb.save(COORD_FILE)

    log(f"Excel created : {COORD_FILE}")
    log(f"Excel Location : {os.path.abspath(COORD_FILE)}")


def save_coordinate(track, x, y):

    # Create the coordinate workbook if it does not exist.
    if not os.path.exists(COORD_FILE):
        create_excel()

    wb = openpyxl.load_workbook(COORD_FILE)
    ws = wb.active

    # Update an existing track instead of creating duplicate rows.
    existing_row = None

    for row in range(4, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value

        if value is not None and str(value).strip().upper() == track.strip().upper():
            existing_row = row
            break

    if existing_row is None:
        existing_row = ws.max_row + 1

    ws.cell(row=existing_row, column=1).value = track
    ws.cell(row=existing_row, column=2).value = x
    ws.cell(row=existing_row, column=3).value = y

    thin = Side(style="thin", color="000000")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in ws[existing_row]:
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)

        for cell in column:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[column_letter].width = max_length + 5

    wb.save(COORD_FILE)

    log(f"{track} saved successfully")


def load_track_coordinates():
    global track_points

    log("============================================================")
    log("LOADING TRACK COORDINATES")
    log(f"COORD_FILE = {COORD_FILE}")
    log(f"COORD_FILE EXISTS = {os.path.exists(COORD_FILE)}")
    log("============================================================")

    if not os.path.exists(COORD_FILE):
        messagebox.showerror(
            "ERROR",
            "Track Coordinate Excel not found."
        )
        log("TRACK COORDINATE FILE NOT FOUND")
        return

    try:
        wb = openpyxl.load_workbook(
            COORD_FILE,
            data_only=True
        )

        log(
            "WORKBOOK SHEETS = "
            + ", ".join(wb.sheetnames)
        )

        ws = wb.active

        log(f"ACTIVE SHEET = {ws.title}")

        track_points = {}

        # Track-coordinate workbook format:
        # Row 1 = title
        # Row 2 = blank
        # Row 3 = headers
        # Row 4+ = data

        for row in ws.iter_rows(
            min_row=2,
            values_only=True
        ):
            if not row or row[0] is None:
                continue

            track = str(row[0]).strip().upper()

            if row[1] is None or row[2] is None:
                continue

            try:
                x = int(float(row[1]))
                y = int(float(row[2]))
            except (ValueError, TypeError):
                log(
                    f"INVALID COORDINATE : "
                    f"{track} -> X={row[1]}, Y={row[2]}"
                )
                continue

            track_points[track] = [x, y]

            log(
                f"TRACK COORD : "
                f"{track} -> X={x}, Y={y}"
            )

        log(
            f"TOTAL TRACK COORDINATES = "
            f"{len(track_points)}"
        )

        log("TRACK COORDINATES LOADED")
        log("============================================================")

        wb.close()

    except Exception as e:
        log(f"TRACK COORDINATE LOAD ERROR : {e}")

        messagebox.showerror(
            "TRACK COORDINATE ERROR",
            str(e)
        )

def load_existing_track_coordinates():
    global track_points

    file_path = filedialog.askopenfilename(
        title="Select Saved Track Coordinate Excel",
        filetypes=[
            ("Excel Files", "*.xlsx"),
            ("All Files", "*.*")
        ]
    )

    if not file_path:
        return

    try:
        wb = openpyxl.load_workbook(
            file_path,
            data_only=True
        )

        ws = wb.active

        log("============================================================")
        log("LOADING EXISTING TRACK COORDINATES")
        log(f"FILE = {file_path}")
        log(f"SHEET = {ws.title}")
        log("============================================================")

        track_points = {}

        # Row 1 = title
        # Row 2 = blank
        # Row 3 = headers
        # Row 4+ = track data

        for row in ws.iter_rows(
            min_row=2,
            values_only=True
        ):
            if not row or row[0] is None:
                continue

            track = str(row[0]).strip().upper()

            if row[1] is None or row[2] is None:
                continue

            try:
                x = int(float(row[1]))
                y = int(float(row[2]))
            except (ValueError, TypeError):
                log(
                    f"INVALID COORDINATE : "
                    f"{track} -> X={row[1]}, Y={row[2]}"
                )
                continue

            track_points[track] = [x, y]

            log(
                f"TRACK COORD : "
                f"{track} -> X={x}, Y={y}"
            )

        wb.close()

        log(
            f"TOTAL TRACK COORDINATES = "
            f"{len(track_points)}"
        )

        if not track_points:
            messagebox.showwarning(
                "TRACK COORDINATES",
                "No valid track coordinates were found."
            )
            return

        log("EXISTING TRACK COORDINATES LOADED SUCCESSFULLY")

        messagebox.showinfo(
            "TRACK COORDINATES",
            f"{len(track_points)} track coordinates loaded successfully."
        )

    except Exception as e:
        log(f"TRACK COORDINATE LOAD ERROR : {e}")

        messagebox.showerror(
            "TRACK COORDINATE ERROR",
            str(e)
        )

def start_track_capture():

    global capture_module
    global track_capture_index
    global track_undo_stack
    global track_points

    if not track_capture_list:
        messagebox.showwarning(
            "TRACK CAPTURE",
            "Please click 'IMPORT TRACK COORDINATES' first."
        )
        return

    if not os.path.exists(COORD_FILE):
        create_excel()

    # Ensure every imported track is ready for a new coordinate.
    for track in track_capture_list:
        track_points[track] = ()

    track_capture_index = 0
    track_undo_stack.clear()
    capture_module = "TRACK"

    create_track_capture_overlay()

    root.iconify()

    log("TRACK CAPTURE STARTED")
    log(f"TRACKS TO CAPTURE : {len(track_capture_list)}")

    update_track_capture_overlay()


def start_capture():
    # Keep the old function name so the existing button wiring can remain
    # unchanged if needed.
    start_track_capture()


#==========================================
# START
# =========================================================

def start_automation():

    global running

    # ============================================
    # CHECK TRACK COORDINATES
    # ============================================

    if not track_points:
        messagebox.showwarning(
            "TRACK COORDINATES",
            "Please click 'LOAD EXISTING TRACK COORDINATES' first."
        )
        return

    invalid_tracks = [
        track
        for track, point in track_points.items()
        if not isinstance(point, (list, tuple))
           or len(point) != 2
    ]

    if invalid_tracks:
        messagebox.showerror(
            "TRACK COORDINATES",
            "Invalid or missing coordinates for:\n\n"
            + "\n".join(invalid_tracks)
        )
        return

    log("TRACK COORDINATES ALREADY LOADED:")
    log(str(track_points))

    # ============================================
    # CHECK SIGNAL CONFIG
    # ============================================

    if not signals:

        log("LOAD CONFIG FIRST")

        return

    # ============================================
    # CHECK LOCK ROUTES
    # ============================================

    if not lock_routes_data:

        log("LOAD LOCK ROUTES FIRST")

        return

    running = True

    pause_event.set()

    status_label.config(
        text="RUNNING",
        fg="#16a34a"
    )

    threading.Thread(
        target=run_engine,
        daemon=True
    ).start()
# =========================================================
# STOP
# =========================================================

def stop_automation():

    global running
    global paused

    running = False

    paused = False

    pause_event.set()

    status_label.config(
        text="STOPPED",
        fg="#dc2626"
    )

    root.deiconify()

    log("STOPPED")

    # =============================================
    # SAVE REPORT EVEN IF STOPPED
    # =============================================

    if report_rows:

        create_report()

        try:
            os.startfile(REPORT_FILE)

        except:
            pass
# =========================================================
# PAUSE / RESUME
# =========================================================

def toggle_pause():

    global paused

    if not running:
        return

    paused = not paused

    if paused:

        pause_event.clear()

        status_label.config(
            text="PAUSED",
            fg="#f59e0b"
        )

        log("AUTOMATION PAUSED")

    else:

        pause_event.set()

        status_label.config(
            text="RUNNING",
            fg="#16a34a"
        )

        log("AUTOMATION RESUMED")

# =========================================================
# GUI
# =========================================================

root = tk.Tk()

root.title(
    "VISUAL INSPECTION FOR TRACK"
)

root.geometry("1450x900")

root.configure(bg="#e9edf2")

root.state("zoomed")

# =========================================================
# STYLE
# =========================================================

style = ttk.Style()

style.theme_use("clam")

style.configure(
    "Treeview",
    background="white",
    foreground="black",
    rowheight=35,
    fieldbackground="white",
    font=("Segoe UI", 10)
)

style.configure(
    "Treeview.Heading",
    background="#1e293b",
    foreground="white",
    font=("Segoe UI", 11, "bold")
)

style.map(
    "Treeview",
    background=[("selected", "#2563eb")]
)

# =========================================================
# TOP HEADER
# =========================================================

header_frame = tk.Frame(
    root,
    bg="#0f172a",
    height=100
)

header_frame.pack(
    fill="x"
)

header_frame.pack_propagate(False)

# =========================================================
# LEFT LOGO
# =========================================================

left_logo_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

left_logo_frame.pack(
    side="left",
    padx=20
)

try:

    railway_logo = tk.PhotoImage(
        file="indian_railways.png"
    )

    railway_logo = railway_logo.subsample(2, 2)

    tk.Label(
        left_logo_frame,
        image=railway_logo,
        bg="#0f172a"
    ).pack()

except:

    tk.Label(
        left_logo_frame,
        text="INDIAN RAILWAYS",
        font=("Segoe UI", 12, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack()

# =========================================================
# TITLE
# =========================================================

title_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

title_frame.pack(
    side="left",
    expand=True
)

tk.Label(

    title_frame,

    text="VISUAL INSPECTION FOR TRACK",

    font=("Segoe UI", 24, "bold"),

    bg="#0f172a",

    fg="white"

).pack(
    pady=(15, 0)
)

tk.Label(

    title_frame,

    text="",

    font=("Segoe UI", 11),

    bg="#0f172a",

    fg="#cbd5e1"

).pack()

# =========================================================
# RIGHT LOGO
# =========================================================

right_logo_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

right_logo_frame.pack(
    side="right",
    padx=20
)

try:

    company_logo = tk.PhotoImage(
        file="company_logo.png"
    )

    company_logo = company_logo.subsample(1, 1)

    tk.Label(
        right_logo_frame,
        image=company_logo,
        bg="#0f172a"
    ).pack()

except:

    tk.Label(
        right_logo_frame,
        text="COMPANY LOGO",
        font=("Segoe UI", 12, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack()

# =========================================================
# MAIN AREA
# =========================================================

main_frame = tk.Frame(
    root,
    bg="#e9edf2"
)

main_frame.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=15
)

# =========================================================
# LEFT PANEL
# =========================================================

left_panel = tk.Frame(
    main_frame,
    bg="white",
    bd=1,
    relief="solid"
)

left_panel.pack(
    side="left",
    fill="y",
    padx=(0, 10)
)

# =========================================================
# CONTROL TITLE
# =========================================================

tk.Label(

    left_panel,

    text="CONTROL PANEL",

    font=("Segoe UI", 15, "bold"),

    bg="white",

    fg="#0f172a"

).pack(
    pady=20
)

# =========================================================
# BUTTON STYLE
# =========================================================

def create_button(text, command, color):

    return tk.Button(

        left_panel,

        text=text,

        command=command,

        bg=color,

        fg="white",

        activebackground=color,

        activeforeground="white",

        relief="flat",

        cursor="hand2",

        font=("Segoe UI", 11, "bold"),

        width=26,

        height=2
    )

# =========================================================
# SIGNAL CONFIG SECTION
# =========================================================

signal_frame = tk.Frame(
    left_panel,
    bg="white"
)

signal_frame.pack(pady=10)

# =========================================================
# DUMMY MAIN BUTTON
# =========================================================

dummy_btn = tk.Button(

    signal_frame,

    text="CAPTURE SIGNALLING GEARS",

    bg="#2563eb",

    fg="white",

    activebackground="#2563eb",

    activeforeground="white",

    relief="flat",

    font=("Segoe UI", 11, "bold"),

    width=26,

    height=2
)

dummy_btn.pack(pady=(0, 8))

# =========================================================
# SMALL BUTTON FRAME
# =========================================================

small_btn_frame = tk.Frame(
    signal_frame,
    bg="white"
)

small_btn_frame.pack()

# =========================================================
# NEW SIGNAL BUTTON
# =========================================================

new_signal_btn = tk.Button(

    small_btn_frame,

    text="NEW SIGNAL",

    command=start_quantity_capture,

    bg="#ea580c",

    fg="white",

    activebackground="#ea580c",

    activeforeground="white",

    relief="flat",

    cursor="hand2",

    font=("Segoe UI", 9, "bold"),

    width=15,

    height=1
)

new_signal_btn.pack(
    side="left",
    padx=5
)

# =========================================================
# EXISTING SIGNAL BUTTON
# =========================================================

existing_signal_btn = tk.Button(

    small_btn_frame,

    text="EXISTING SIGNAL",

    command=load_config,

    bg="#ea580c",

    fg="white",

    activebackground="#ea580c",

    activeforeground="white",

    relief="flat",

    cursor="hand2",

    font=("Segoe UI", 9, "bold"),

    width=15,

    height=1
)

existing_signal_btn.pack(
    side="left",
    padx=5
)

# =========================================================
# OTHER BUTTONS
# =========================================================

create_button(
    "LOAD TOC",
    load_lock_routes,
    "#2563eb"
).pack(pady=8)



create_button(
    "SAVE NEW SIGNAL COORDINATES",
    save_new_signal_coordinates,
    "#16a34a"
).pack(pady=8)

# ==========================
# CAPTURE TRACK COORDINATES
# ==========================

create_button(
    "CAPTURE TRACK COORDINATES",
    start_track_capture,
    "#9333ea"
).pack(pady=8)

# ==========================
# LOAD EXISTING TRACK COORDINATES
# ==========================

create_button(
    "LOAD EXISTING TRACK COORDINATES",
    load_existing_track_coordinates,
    "#0891b2"
).pack(pady=8)

create_button(
    "START TESTING",
    start_automation,
    "#16a34a"
).pack(pady=20)

create_button(
    "STOP TESTING",
    stop_automation,
    "#dc2626"
).pack(pady=8)

# =========================================================
# STATUS BOX
# =========================================================

status_frame = tk.Frame(
    left_panel,
    bg="#f8fafc",
    bd=1,
    relief="solid"
)

status_frame.pack(
    fill="x",
    padx=15,
    pady=25
)

tk.Label(

    status_frame,

    text="SYSTEM STATUS",

    font=("Segoe UI", 11, "bold"),

    bg="#f8fafc",

    fg="#0f172a"

).pack(
    pady=10
)

status_label = tk.Label(

    status_frame,

    text="READY",

    font=("Segoe UI", 16, "bold"),

    bg="#f8fafc",

    fg="#16a34a"

)

status_label.pack(
    pady=(0, 15)
)

# =========================================================
# RIGHT PANEL
# =========================================================

right_panel = tk.Frame(
    main_frame,
    bg="#e9edf2"
)

right_panel.pack(
    side="left",
    fill="both",
    expand=True
)

# =========================================================
# TABLE TITLE
# =========================================================

table_title = tk.Label(

    right_panel,

    text="TRACK DETAILS",

    font=("Segoe UI", 16, "bold"),

    bg="#e9edf2",

    fg="#0f172a"

)

table_title.pack(
    anchor="w",
    pady=(0, 10)
)

# =========================================================
# TABLE FRAME
# =========================================================

table_frame = tk.Frame(
    right_panel,
    bg="white",
    bd=1,
    relief="solid"
)

table_frame.pack(
    fill="both",
    expand=True
)

# =========================================================
# SCROLLBAR
# =========================================================

tree_scroll = ttk.Scrollbar(
    table_frame
)

tree_scroll.pack(
    side="right",
    fill="y"
)

# =========================================================
# TREEVIEW
# =========================================================

tree = ttk.Treeview(

    table_frame,

    columns=(

        "NO",

        "SIGNAL",

        "ROUTE",

        "LOCK_ROUTE"

    ),

    show="headings",

    yscrollcommand=tree_scroll.set

)

tree_scroll.config(
    command=tree.yview
)

tree.heading("NO", text="NO")

tree.heading("SIGNAL", text="SIGNAL")

tree.heading("ROUTE", text="ROUTE")

tree.heading("LOCK_ROUTE", text="CONTROLLED BY TRACKS")

tree.column("NO", width=80, anchor="center")

tree.column("SIGNAL", width=200, anchor="center")

tree.column("ROUTE", width=220, anchor="center")

tree.column("LOCK_ROUTE", width=500, anchor="w")

tree.pack(
    fill="both",
    expand=True
)

# =========================================================
# LOG TITLE
# =========================================================

log_title = tk.Label(

    right_panel,

    text="LIVE OPERATION LOG",

    font=("Segoe UI", 16, "bold"),

    bg="#e9edf2",

    fg="#0f172a"

)

log_title.pack(
    anchor="w",
    pady=(20, 10)
)

# =========================================================
# LOG FRAME
# =========================================================

log_frame = tk.Frame(
    right_panel,
    bg="white",
    bd=1,
    relief="solid"
)

log_frame.pack(
    fill="both",
    expand=True
)

# =========================================================
# LOG TEXT
# =========================================================

log_text = tk.Text(

    log_frame,

    bg="white",

    fg="black",

    font=("Consolas", 10),

    relief="flat"

)

log_text.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

log_text.config(state="disabled")

# =========================================================
# FOOTER
# =========================================================

footer = tk.Label(

    root,

    text="SPACE KEY = CAPTURE COORDINATES | LOCK ROUTE AUTOMATION SYSTEM",

    bg="#0f172a",

    fg="white",

    font=("Segoe UI", 10)

)

footer.pack(
    fill="x"
)
# =========================================================
# KEYBOARD SHORTCUTS
# =========================================================

keyboard.add_hotkey(
    "p",
    toggle_pause
)
# =========================================================
# RUN
# =========================================================
listener = pynput_keyboard.Listener(on_press=on_press)
listener.start()

root.mainloop()
