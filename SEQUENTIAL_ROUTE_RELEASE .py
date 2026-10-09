# =========================================================
# SEQUENTIAL ROUTE RELEASE AUTOMATION SYSTEM
# =========================================================
# FEATURES
# =========================================================
# 1. HOME SIGNAL ROUTE SETTING
# 2. FIRST-TRACK CONTROL-BIT OPERATION
# 3. SEQUENTIAL TRACK RELEASE
# 4. PASS / FAIL REPORT
# =========================================================
#
# REQUIRED MODULES
#
# pip install pyautogui openpyxl pywinauto
# pip install pynput pillow uiautomation pywin32
#
# =========================================================

import tkinter as tk
import keyboard
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

from openpyxl.utils import get_column_letter

from openpyxl.drawing.image import Image

import threading
import time
import os
import sys
import sys

print("========================================")
print("SEQUENTIAL_ROUTE_RELEASE.PY STARTED")
print("ARGV =", sys.argv)
print("========================================")
# =========================================================
# SHUNT SNAPSHOTS
# =========================================================

SNAPSHOTS_FOLDER = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "snapshots"
)

os.makedirs(
    SNAPSHOTS_FOLDER,
    exist_ok=True
)

import pandas as pd

import pyautogui

import win32api
import win32con
from collections import Counter

import uiautomation as auto

from pywinauto import Desktop
from PIL import Image, ImageTk
from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "SEQUENTIAL_ROUTE_RELEASE_REPORT.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

lock_routes_data = []

toc_header_row = 1

# Rows loaded from the SEQUENTIAL_ROUTE_RELEASE worksheet.
# Required columns: SIGNAL, ROUTE, FIRST_TRACK, RELEASE_SEQUENCE.
# Optional column: EXPECTED_FINAL_ASPECT (defaults to RED).
routes_data = []

track_points = {}

running = False

# True only while Short Train / Long Train / ERR automation is executing.
# Prevents GUI/file-loading functions from modifying runtime state.
automation_active = False

capture_module = None

capture_index = 0

capture_master_list = []

record_stage = None

current_aspects = 2

undo_stack = []

current_signal = None

config_file_path = ""



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
    "SHUNT": "SHUNT SIGNAL"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CALLING_ON": "#f97316",
    "SHUNT": "#a855f7"
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
        capture_overlay, text="\u21b6  UNDO LAST CLICK", command=undo_last_capture,
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

    # =====================================================
    # MAIN
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "MAIN":
            capture_list.append(("MAIN", signal))

    # =====================================================
    # SHUNT
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "SHUNT":
            capture_list.append(("SHUNT", signal))

    # =====================================================
    # CALLING-ON
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "CALLING_ON":
            capture_list.append(("CALLING_ON", signal))

    # =====================================================
    # POINT
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "POINT":
            capture_list.append(("POINT", signal))

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "CH":
            capture_list.append(("CH", signal))

    # =====================================================
    # LC GATE
    # =====================================================

    for signal, info in signals.items():

        if info["type"] == "LC":
            capture_list.append(("LC", signal))

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

        record_stage = "menu"

    elif sig_type == "CH":

        record_stage = "menu"

    elif sig_type == "LC":

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
                "C1": None,
                "C2": None,
                "C3": None
            }

        elif signal.endswith("C"):

            signals[signal] = {
                "type": "CALLING_ON",
                "menu": None,
                "YELLOW": None,
                "ROUTE_INDICATOR": None
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

        capture_module = None

        destroy_capture_overlay()

        root.deiconify()

        log("ALL COORDINATES CAPTURED")

        undo_stack.clear()

        messagebox.showinfo(
            "Capture Complete",
            f"All {len(capture_master_list)} elements captured successfully."
        )
        save_config()
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
                "CLICK: SIGNAL MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "ask_aspects":

            show_aspect_buttons(signal)

        elif record_stage == "RED":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK: RED ASPECT",
                "Move the mouse onto the RED lamp, then press SPACE."
            )

        elif record_stage == "YELLOW":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK: YELLOW ASPECT",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )

        elif record_stage == "DOUBLE_YELLOW":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK: DOUBLE YELLOW ASPECT",
                "Move the mouse onto the DOUBLE YELLOW lamp, then press SPACE."
            )

        elif record_stage == "GREEN":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK: GREEN ASPECT",
                "Move the mouse onto the GREEN lamp, then press SPACE."
            )

        elif record_stage == "ROUTE_INIT":

            update_capture_overlay(
                "MAIN",
                signal,
                "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # CALLING-ON
    # =====================================================

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "yellow":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK: YELLOW LAMP",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )

        elif record_stage == "route_init":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # SHUNT
    # =====================================================

    elif sig_type == "SHUNT":

        if record_stage == "menu":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "indicator":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK: ASPECT INDICATOR",
                "Move the mouse onto the aspect indicator, then press SPACE.\n"
                "This also saves a reference snapshot for comparison."
            )

        elif record_stage == "route_init":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # POINT
    # =====================================================

    elif sig_type == "POINT":

        if record_stage == "menu":

            update_capture_overlay(
                "POINT",
                signal,
                "CLICK: MENU",
                "Move the mouse onto this point's menu button, then press SPACE."
            )

        elif record_stage == "normal":

            update_capture_overlay(
                "POINT",
                signal,
                "CLICK: NORMAL ASPECT",
                "Move the mouse onto the NORMAL aspect indicator, then press SPACE."
            )

        elif record_stage == "reverse":

            update_capture_overlay(
                "POINT",
                signal,
                "CLICK: REVERSE ASPECT",
                "Move the mouse onto the REVERSE aspect indicator, then press SPACE."
            )

        elif record_stage == "free":

            update_capture_overlay(
                "POINT",
                signal,
                "CLICK: FREE ASPECT",
                "Move the mouse onto the FREE (out of correspondence) indicator, then press SPACE."
            )

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif sig_type == "CH":

        if record_stage == "menu":

            update_capture_overlay(
                "CH",
                signal,
                "CLICK: MENU",
                "Move the mouse onto this crank handle's menu button, then press SPACE."
            )

        elif record_stage == "IN":

            update_capture_overlay(
                "CH",
                signal,
                "CLICK: IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )

        elif record_stage == "OUT":

            update_capture_overlay(
                "CH",
                signal,
                "CLICK: OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )

        elif record_stage == "ECH":

            update_capture_overlay(
                "CH",
                signal,
                "CLICK: ECH ASPECT",
                "Move the mouse onto the ECH indicator, then press SPACE."
            )

        elif record_stage == "FREE":

            update_capture_overlay(
                "CH",
                signal,
                "CLICK: FREE ASPECT",
                "Move the mouse onto the FREE indicator, then press SPACE."
            )

    # =====================================================
    # LC GATE
    # =====================================================

    elif sig_type == "LC":

        if record_stage == "menu":

            update_capture_overlay(
                "LC",
                signal,
                "CLICK: MENU",
                "Move the mouse onto this LC gate's menu button, then press SPACE."
            )

        elif record_stage == "in":

            update_capture_overlay(
                "LC",
                signal,
                "CLICK: IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )

        elif record_stage == "out":

            update_capture_overlay(
                "LC",
                signal,
                "CLICK: OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )

def master_save_point(x, y):

    global record_stage

    if capture_index >= len(capture_master_list):
        return

    sig_type, signal = capture_master_list[capture_index]

    # =====================================================
    # MAIN
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

            log(
                f"{signal} - SPACE does nothing here. "
                "Click 2 / 3 / 4 or press 2 / 3 / 4."
            )

            return

        if record_stage == "RED":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("RED", None)
            )

            signals[signal]["RED"] = [x, y]

            log(f"{signal} RED Saved")

            record_stage = (
                "GREEN"
                if current_aspects == 2
                else "YELLOW"
            )

            master_next_capture()
            return

        if record_stage == "YELLOW":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("YELLOW", None)
            )

            signals[signal]["YELLOW"] = [x, y]

            log(f"{signal} YELLOW Saved")

            record_stage = (
                "DOUBLE_YELLOW"
                if current_aspects == 4
                else "GREEN"
            )

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

            record_stage = "ROUTE_INIT"

            master_next_capture()
            return

        if record_stage == "ROUTE_INIT":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("ROUTE_INDICATOR", None)
            )

            signals[signal]["ROUTE_INDICATOR"] = [x, y]

            log(f"{signal} ROUTE INIT Saved")

            advance_to_next_signal()
            return

    # =====================================================
    # CALLING-ON
    # =====================================================

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved")

            record_stage = "yellow"

            master_next_capture()
            return

        if record_stage == "yellow":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("YELLOW", None)
            )

            signals[signal]["YELLOW"] = [x, y]

            log(f"{signal} YELLOW Saved")

            record_stage = "route_init"

            master_next_capture()
            return

        if record_stage == "route_init":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("route_init", None)
            )

            signals[signal]["route_init"] = [x, y]

            log(f"{signal} ROUTE INIT Saved")

            advance_to_next_signal()
            return

    # =====================================================
    # SHUNT
    # =====================================================

    elif sig_type == "SHUNT":

        # -------------------------------------------------
        # MENU
        # -------------------------------------------------

        if record_stage == "menu":
            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved ({x},{y})")

            record_stage = "indicator"

            master_next_capture()

            return

        # -------------------------------------------------
        # INDICATOR + SNAPSHOT
        # -------------------------------------------------

        if record_stage == "indicator":

            screenshot = pyautogui.screenshot()

            region = screenshot.crop(
                (x - 50, y - 50, x + 50, y + 50)
            )

            snapshot_file = os.path.join(
                SNAPSHOTS_FOLDER,
                f"{signal}_initial.png"
            )

            region.save(snapshot_file)

            def undo_shunt(
                    s=signal,
                    f=snapshot_file
            ):

                signals[s]["indicator"] = None
                signals[s]["initial_snapshot"] = None

                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass

            push_undo(undo_shunt)

            signals[signal]["indicator"] = [x, y]

            signals[signal]["initial_snapshot"] = snapshot_file

            log(
                f"{signal} INDICATOR Saved ({x},{y})"
            )

            log(
                f"{signal} Snapshot: {snapshot_file}"
            )

            record_stage = "route_init"

            master_next_capture()

            return

        # -------------------------------------------------
        # ROUTE INITIATION
        # -------------------------------------------------

        if record_stage == "route_init":
            push_undo(
                lambda s=signal:
                signals[s].__setitem__("route_init", None)
            )

            signals[signal]["route_init"] = [x, y]

            log(
                f"{signal} ROUTE INIT Saved ({x},{y})"
            )

            advance_to_next_signal()

            return

    # =====================================================
    # POINT
    # =====================================================

    elif sig_type == "POINT":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} POINT MENU Saved")

            record_stage = "normal"

            master_next_capture()
            return

        if record_stage == "normal":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("normal", None)
            )

            signals[signal]["normal"] = [x, y]

            log(f"{signal} NORMAL Saved")

            record_stage = "reverse"

            master_next_capture()
            return

        if record_stage == "reverse":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("reverse", None)
            )

            signals[signal]["reverse"] = [x, y]

            log(f"{signal} REVERSE Saved")

            record_stage = "free"

            master_next_capture()
            return

        if record_stage == "free":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("free", None)
            )

            signals[signal]["free"] = [x, y]

            log(f"{signal} FREE Saved")

            advance_to_next_signal()
            return

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif sig_type == "CH":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} CH MENU Saved")

            record_stage = "IN"

            master_next_capture()
            return

        if record_stage == "IN":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("IN", None)
            )

            signals[signal]["IN"] = [x, y]

            log(f"{signal} CH IN Saved")

            record_stage = "OUT"

            master_next_capture()
            return

        if record_stage == "OUT":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("OUT", None)
            )

            signals[signal]["OUT"] = [x, y]

            log(f"{signal} CH OUT Saved")

            record_stage = "ECH"

            master_next_capture()
            return

        if record_stage == "ECH":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("ECH", None)
            )

            signals[signal]["ECH"] = [x, y]

            log(f"{signal} CH ECH Saved")

            record_stage = "FREE"

            master_next_capture()
            return

        if record_stage == "FREE":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("FREE", None)
            )

            signals[signal]["FREE"] = [x, y]

            log(f"{signal} CH FREE Saved")

            advance_to_next_signal()
            return

    # =====================================================
    # LC GATE
    # =====================================================

    elif sig_type == "LC":

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} LC MENU Saved")

            record_stage = "in"

            master_next_capture()
            return

        if record_stage == "in":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("in", None)
            )

            signals[signal]["in"] = [x, y]

            log(f"{signal} LC IN Saved")

            record_stage = "out"

            master_next_capture()
            return

        if record_stage == "out":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("out", None)
            )

            signals[signal]["out"] = [x, y]

            log(f"{signal} LC OUT Saved")

            advance_to_next_signal()
            return


def set_aspects_and_continue(signal, aspects):

    global current_aspects
    global record_stage

    push_undo(
        lambda s=signal:
        signals[s].__setitem__("aspects", None)
    )

    signals[signal]["aspects"] = aspects

    current_aspects = aspects

    record_stage = "RED"

    log(f"{signal} Aspects Set To {aspects}")

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
            "indicator": None,
            "initial_snapshot": None,
            "route_init": None
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
            "route_init": None
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
            "free": None
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
            "FREE": None
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
            "in": None,
            "out": None
        }

        combined.append(("LC", name))

    return combined

def start_quantity_capture():

    global capture_module

    if running or automation_active:
        log("NEW SIGNAL BLOCKED : AUTOMATION RUNNING")
        return
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

def dispatch_undo():

    if capture_module == "MASTER":
        undo_last_capture()

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

        if capture_module not in ("MASTER", "ROUTE"):
            return

        x, y = win32api.GetCursorPos()

        root.after(0, lambda: dispatch_save_point(x, y))

    # ---------------- BACKSPACE ----------------
    elif key == pynput_keyboard.Key.backspace:

        now = time.time()

        if now - last_backspace_time < 0.4:
            return

        last_backspace_time = now

        if capture_module != "MASTER":
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

    def write():

        log_text.config(state="normal")

        log_text.insert(
            tk.END,
            f"[{time.strftime('%H:%M:%S')}] {msg}\n"
        )

        log_text.see(tk.END)

        log_text.config(state="disabled")

    root.after(0, write)

# =========================================================
# CLICK
# =========================================================

def click(point):

    if point is None:
        log("CLICK FAILED : POINT IS NONE")
        return False

    if not running or not automation_active:
        log("CLICK BLOCKED : AUTOMATION NOT ACTIVE")
        return False

    pause_event.wait()

    x, y = point

    log(f"EDRC CLICK -> ({x}, {y})")

    pyautogui.moveTo(
        x,
        y,
        duration=0.1
    )

    time.sleep(0.5)

    pyautogui.click()

    return True

# =========================================================
# CREATE SIGNAL SETUP
# =========================================================

def create_setup():

    global signals

    if not routes_data:

        messagebox.showwarning(
            "WARNING",
            "Please load the TOC workbook first."
        )
        return

    start_quantity_capture()


# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():
    global config_file_path

    if config_file_path and os.path.exists(config_file_path):

        wb = load_workbook(config_file_path)


    else:

        wb = Workbook()

    if "MAIN_SIGNALS" in wb.sheetnames:
        del wb["MAIN_SIGNALS"]

    if "CALLING_ON" in wb.sheetnames:
        del wb["CALLING_ON"]

    if "SHUNT" in wb.sheetnames:
        del wb["SHUNT"]

    if "POINTS" in wb.sheetnames:
        del wb["POINTS"]

    if "CRANK_HANDLES" in wb.sheetnames:
        del wb["CRANK_HANDLES"]

    if "LC_GATES" in wb.sheetnames:
        del wb["LC_GATES"]

    # =====================================================
    # MAIN SHEET
    # =====================================================

    ws_main = wb.create_sheet("MAIN_SIGNALS", 0)
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

        "Green_Y",

        "Route_X",

        "Route_Y"

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
        "RouteInit_X",
        "RouteInit_Y"
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
        "Indicator_X",
        "Indicator_Y",
        "RouteInit_X",
        "RouteInit_Y",
        "Snapshot_Path"
    ])

    # =====================================================
    # POINTS
    # =====================================================

    ws_point = wb.create_sheet("POINTS")

    ws_point.append([
        "POINT",
        "Menu_X", "Menu_Y",
        "Normal_X", "Normal_Y",
        "Reverse_X", "Reverse_Y",
        "Free_X", "Free_Y"
    ])

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    ws_ch = wb.create_sheet("CRANK_HANDLES")

    ws_ch.append([
        "CRANK_HANDLE",
        "Menu_X", "Menu_Y",
        "IN_X", "IN_Y",
        "OUT_X", "OUT_Y",
        "ECH_X", "ECH_Y",
        "FREE_X", "FREE_Y"
    ])

    # =====================================================
    # LC GATES
    # =====================================================

    ws_lc = wb.create_sheet("LC_GATES")

    ws_lc.append([
        "LC_GATE",
        "MENU_X", "MENU_Y",
        "IN_X", "IN_Y",
        "OUT_X", "OUT_Y"
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

                (info.get("GREEN") or [None, None])[0],
                (info.get("GREEN") or [None, None])[1],

                (info.get("ROUTE_INDICATOR") or [None, None])[0],
                (info.get("ROUTE_INDICATOR") or [None, None])[1]


            ])

        # =================================================
        # CALLING ON
        # =================================================

        elif typ == "CALLING_ON":

            menu = info.get("menu")

            yellow = info.get("YELLOW")

            route_init = info.get("route_init")

            ws_call.append([

                sig,

                menu[0] if menu else None,

                menu[1] if menu else None,

                yellow[0] if yellow else None,

                yellow[1] if yellow else None,

                route_init[0] if route_init else None,

                route_init[1] if route_init else None

            ])

        # =================================================
        # SHUNT
        # =================================================

        elif typ == "SHUNT":

            menu = info.get("menu")

            indicator = info.get("indicator")

            route_init = info.get("route_init")

            snapshot = info.get("initial_snapshot")

            ws_shunt.append([

                sig,

                menu[0] if menu else None,

                menu[1] if menu else None,

                indicator[0] if indicator else None,

                indicator[1] if indicator else None,

                route_init[0] if route_init else None,

                route_init[1] if route_init else None,

                snapshot if snapshot else ""

            ])

        # =================================================
        # POINT
        # =================================================

        elif typ == "POINT":

            info_menu = info.get("menu")

            info_normal = info.get("normal")

            info_reverse = info.get("reverse")

            info_free = info.get("free")

            ws_point.append([

                sig,

                info_menu[0] if info_menu else None,

                info_menu[1] if info_menu else None,

                info_normal[0] if info_normal else None,

                info_normal[1] if info_normal else None,

                info_reverse[0] if info_reverse else None,

                info_reverse[1] if info_reverse else None,

                info_free[0] if info_free else None,

                info_free[1] if info_free else None

            ])

        # =================================================
        # CRANK HANDLE
        # =================================================

        elif typ == "CH":

            menu = info.get("menu")

            in_point = info.get("IN")

            out_point = info.get("OUT")

            ech = info.get("ECH")

            free = info.get("FREE")

            ws_ch.append([

                sig,

                menu[0] if menu else None,

                menu[1] if menu else None,

                in_point[0] if in_point else None,

                in_point[1] if in_point else None,

                out_point[0] if out_point else None,

                out_point[1] if out_point else None,

                ech[0] if ech else None,

                ech[1] if ech else None,

                free[0] if free else None,

                free[1] if free else None

            ])

        # =================================================
        # LC GATE
        # =================================================

        elif typ == "LC":

            menu = info.get("menu")

            in_point = info.get("in")

            out_point = info.get("out")

            ws_lc.append([

                sig,

                menu[0] if menu else None,

                menu[1] if menu else None,

                in_point[0] if in_point else None,

                in_point[1] if in_point else None,

                out_point[0] if out_point else None,

                out_point[1] if out_point else None

            ])

    # =====================================================
    # SAVE FILE
    # =====================================================

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

def startup_walkaway():

    log("--------------------------------")
    log(f"COMMAND LINE ARGS : {sys.argv}")

    # =====================================================
    # FIRST: CHECK EDRC ENVIRONMENT VARIABLES
    # =====================================================

    toc_file = os.environ.get("EDRC_LIST")
    config_file = os.environ.get("EDRC_COORDS")

    if toc_file and config_file:

        log("EDRC FILES FOUND FROM ENVIRONMENT")
        log(f"TOC    : {toc_file}")
        log(f"CONFIG : {config_file}")

    # =====================================================
    # SECOND: FALLBACK TO COMMAND LINE
    # =====================================================

    elif len(sys.argv) >= 3:

        toc_file = sys.argv[1]
        config_file = sys.argv[2]

        log("FILES FOUND FROM COMMAND LINE")
        log(f"TOC    : {toc_file}")
        log(f"CONFIG : {config_file}")

    # =====================================================
    # THIRD: GUI MODE
    # =====================================================

    else:

        log("NO TOC/CONFIG FILES RECEIVED")
        log("STARTING GUI MODE")
        return

    # =====================================================
    # WALK-AWAY MODE
    # =====================================================

    log("--------------------------------")
    log("WALK-AWAY MODE")
    log(f"TOC    : {toc_file}")
    log(f"CONFIG : {config_file}")

    # =====================================================
    # LOAD TOC
    # =====================================================

    if not load_routes(toc_file):

        log("FATAL : TOC LOAD FAILED")

        root.after(
            500,
            root.destroy
        )

        return

    # =====================================================
    # LOAD SIGNAL CONFIG / COORDINATES
    # =====================================================

    log("--------------------------------")
    log("LOADING SIGNAL CONFIG")
    log(f"CONFIG FILE : {config_file}")

    config_loaded = load_config(config_file)

    log(f"CONFIG LOAD RESULT : {config_loaded}")
    log(f"CONFIG SIGNAL COUNT : {len(signals)}")

    if not config_loaded or not signals:
        log("FATAL : SIGNAL CONFIG LOAD FAILED")
        log("FATAL : NO SIGNALS AVAILABLE")

        root.after(
            500,
            root.destroy
        )

        return

    # =====================================================
    # VERIFY LOADED DATA BEFORE START
    # =====================================================

    if not routes_data:
        log("FATAL : TOC LOADED BUT ROUTES DATA IS EMPTY")

        root.after(
            500,
            root.destroy
        )

        return

    if not signals:
        log("FATAL : CONFIG LOADED BUT SIGNAL DATA IS EMPTY")
        log(f"CONFIG FILE : {config_file_path}")

        root.after(
            500,
            root.destroy
        )

        return

    log("--------------------------------")
    log("EDRC WALK-AWAY MODE")
    log(f"SHARED EDRC TOC : {toc_file}")
    log(f"SIGNAL CONFIG   : {config_file}")
    log(f"FINAL SIGNAL COUNT : {len(signals)}")
    log(f"FINAL ROUTE COUNT  : {len(routes_data)}")
    log("--------------------------------")

    # =====================================================
    # START AUTOMATION ONCE
    # =====================================================

    root.after(
        1000,
        start_automation
    )

# =========================================================
# LOAD CONFIG
# =========================================================

def load_config(file_path=None):

    global signals
    global config_file_path

    # NEVER allow signal configuration to be changed during automation.
    if running or automation_active:
        log("LOAD CONFIG BLOCKED : AUTOMATION RUNNING")
        return False

    # GUI mode: existing behaviour
    if file_path is None:

        file_path = filedialog.askopenfilename(
            title="Select Signal Configuration",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if not file_path:
            return False

    config_file_path = file_path
    signals = {}


    wb = load_workbook(config_file_path)

    # =====================================================
    # FILE FORMAT DETECTION
    # This program's native format uses "MAIN_SIGNALS" /
    # "CALLING_ON" / "SHUNT" / "POINTS" / "CRANK_HANDLES" /
    # "LC_GATES" sheets, accessed by fixed column position. The
    # Universal Yard Coordinate file uses different sheet names
    # ("MAIN" / "CAL" / "SHUNT" / "POINT" / "CH" / "LC") with
    # header-named columns. Reading the wrong one by position
    # would silently mis-map data at best, or crash with a
    # KeyError on the sheet name at worst - so detect which one
    # this actually is before parsing.
    # =====================================================

    if "MAIN_SIGNALS" not in wb.sheetnames and "MAIN" in wb.sheetnames:

        log("UNIVERSAL COORDINATE FORMAT DETECTED (MAIN sheet, no MAIN_SIGNALS)")

        _load_universal_signals_from_workbook(wb)

        wb.close()

        log(f"SIGNAL CONFIG LOADED : {config_file_path}")
        log(f"SIGNALS LOADED : {len(signals)}")

        return bool(signals)

    # =====================================================
    # MAIN
    # =====================================================

    ws = wb["MAIN_SIGNALS"]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        signals[signal_name] = {

            "type": "MAIN",

            "aspects": row[1],

            "menu": [row[2], row[3]],

            "RED": [row[4], row[5]],

            "YELLOW": [row[6], row[7]]
            if row[6] else None,

            "DOUBLE_YELLOW": [row[8], row[9]]
            if row[8] else None,

            "GREEN": [row[10], row[11]]
            if row[10] is not None else None,

            "ROUTE_INDICATOR": [row[12], row[13]]
            if row[12] is not None else None

        }

    # =====================================================
    # CALLING ON
    # =====================================================

    ws = wb["CALLING_ON"]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        signals[signal_name] = {
            "type": "CALLING_ON",

            "menu": (
                [row[1], row[2]]
                if row[1] is not None and row[2] is not None
                else None
            ),

            "YELLOW": (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            ),

            "route_init": (
                [row[5], row[6]]
                if len(row) > 6
                   and row[5] is not None
                   and row[6] is not None
                else None
            )
        }

    # =====================================================
    # SHUNT
    # =====================================================

    ws = wb["SHUNT"]

    for row in ws.iter_rows(
            min_row=2,
            values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        snapshot_path = (
            str(row[7]).strip()
            if len(row) > 7 and row[7]
            else None
        )

        signals[signal_name] = {
            "type": "SHUNT",

            "menu": (
                [row[1], row[2]]
                if row[1] is not None and row[2] is not None
                else None
            ),

            "indicator": (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            ),

            "route_init": (
                [row[5], row[6]]
                if row[5] is not None and row[6] is not None
                else None
            ),

            "initial_snapshot": snapshot_path
        }

    # =====================================================
    # POINTS
    # =====================================================

    if "POINTS" in wb.sheetnames:

        ws = wb["POINTS"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row or not row[0]:
                continue

            name = str(row[0]).strip().upper()

            if name not in signals:
                signals[name] = {
                    "type": "POINT",
                    "menu": None,
                    "normal": None,
                    "reverse": None,
                    "free": None
                }

            signals[name]["menu"] = (
                [row[1], row[2]]
                if row[1] is not None and row[2] is not None
                else None
            )

            signals[name]["normal"] = (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            )

            signals[name]["reverse"] = (
                [row[5], row[6]]
                if row[5] is not None and row[6] is not None
                else None
            )

            signals[name]["free"] = (
                [row[7], row[8]]
                if row[7] is not None and row[8] is not None
                else None
            )

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    if "CRANK_HANDLES" in wb.sheetnames:

        ws = wb["CRANK_HANDLES"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row or not row[0]:
                continue

            name = str(row[0]).strip().upper()

            if name not in signals:
                signals[name] = {
                    "type": "CH",
                    "menu": None,
                    "IN": None,
                    "OUT": None,
                    "ECH": None,
                    "FREE": None
                }

            signals[name]["menu"] = (
                [row[1], row[2]]
                if row[1] is not None and row[2] is not None
                else None
            )

            signals[name]["IN"] = (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            )

            signals[name]["OUT"] = (
                [row[5], row[6]]
                if row[5] is not None and row[6] is not None
                else None
            )

            signals[name]["ECH"] = (
                [row[7], row[8]]
                if row[7] is not None and row[8] is not None
                else None
            )

            signals[name]["FREE"] = (
                [row[9], row[10]]
                if row[9] is not None and row[10] is not None
                else None
            )

    # =====================================================
    # LC GATES
    # =====================================================

    if "LC_GATES" in wb.sheetnames:

        ws = wb["LC_GATES"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row or not row[0]:
                continue

            name = str(row[0]).strip().upper()

            if name not in signals:
                signals[name] = {
                    "type": "LC",
                    "menu": None,
                    "in": None,
                    "out": None
                }

            signals[name]["menu"] = (
                [row[1], row[2]]
                if row[1] is not None and row[2] is not None
                else None
            )

            signals[name]["in"] = (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            )

            signals[name]["out"] = (
                [row[5], row[6]]
                if row[5] is not None and row[6] is not None
                else None
            )
    # =====================================================
    # FINISH CONFIG LOAD
    # =====================================================

    log(f"SIGNAL CONFIG LOADED : {config_file_path}")
    log(f"SIGNALS LOADED : {len(signals)}")

    return bool(signals)


def _load_universal_signals_from_workbook(wb):
    """Populates `signals` from an already-open Universal Yard
    Coordinate workbook, reading MAIN / CAL / SHUNT / POINT / CH / LC
    sheets by HEADER NAME. Shared by load_config()'s auto-detection
    above and load_universal_coordinates() below, so both paths
    behave identically."""

    global signals

    def parse_int(value):
        if value is None:
            return None
        if isinstance(value, str) and value.strip() == "":
            return None
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    def parse_point(x_val, y_val):
        x = parse_int(x_val)
        y = parse_int(y_val)
        if x is None or y is None:
            return None
        return [x, y]

    def header_index_of(ws_u):
        header_row = next(ws_u.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            return {}
        idx = {}
        for i, val in enumerate(header_row):
            if val is not None:
                idx[str(val).strip().upper()] = i
        return idx

    def cell(row, idx):
        return row[idx] if idx is not None and len(row) > idx else None

    loaded_main = loaded_cal = loaded_shunt = 0
    loaded_point = loaded_ch = loaded_lc = 0

    # ---- MAIN ----
    if "MAIN" in wb.sheetnames:
        ws_u = wb["MAIN"]
        hi = header_index_of(ws_u)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).replace(".0", "").strip().upper()
                signals[sig] = {
                    "type": "MAIN",
                    "aspects": parse_int(cell(row, hi.get("ASPECTS"))) or 3,
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "RED": parse_point(cell(row, hi.get("RED_X")), cell(row, hi.get("RED_Y"))),
                    "YELLOW": parse_point(cell(row, hi.get("YELLOW_X")), cell(row, hi.get("YELLOW_Y"))),
                    "DOUBLE_YELLOW": parse_point(cell(row, hi.get("DOUBLE_YELLOW_X")), cell(row, hi.get("DOUBLE_YELLOW_Y"))),
                    "GREEN": parse_point(cell(row, hi.get("GREEN_X")), cell(row, hi.get("GREEN_Y"))),
                    "ROUTE_INDICATOR": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                }
                loaded_main += 1

    # ---- CAL ----
    if "CAL" in wb.sheetnames:
        ws_u = wb["CAL"]
        hi = header_index_of(ws_u)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).replace(".0", "").strip().upper()
                signals[sig] = {
                    "type": "CALLING_ON",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "YELLOW": parse_point(cell(row, hi.get("YELLOW_X")), cell(row, hi.get("YELLOW_Y"))),
                    "route_init": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                }
                loaded_cal += 1

    # ---- SHUNT ----
    if "SHUNT" in wb.sheetnames:
        ws_u = wb["SHUNT"]
        hi = header_index_of(ws_u)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).replace(".0", "").strip().upper()
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "indicator": parse_point(cell(row, hi.get("INDICATOR_X")), cell(row, hi.get("INDICATOR_Y"))),
                    "route_init": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y"))),
                    "initial_snapshot": None
                }
                loaded_shunt += 1

    # ---- POINT ----
    if "POINT" in wb.sheetnames:
        ws_u = wb["POINT"]
        hi = header_index_of(ws_u)
        name_col = hi.get("POINT")
        if name_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, name_col) is None:
                    continue
                name = str(cell(row, name_col)).replace(".0", "").strip().upper()
                signals[name] = {
                    "type": "POINT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "normal": parse_point(cell(row, hi.get("NORMAL_X")), cell(row, hi.get("NORMAL_Y"))),
                    "reverse": parse_point(cell(row, hi.get("REVERSE_X")), cell(row, hi.get("REVERSE_Y"))),
                    "free": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y")))
                }
                loaded_point += 1

    # ---- CH ----
    if "CH" in wb.sheetnames:
        ws_u = wb["CH"]
        hi = header_index_of(ws_u)
        name_col = hi.get("CRANKHANDLE")
        if name_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, name_col) is None:
                    continue
                name = str(cell(row, name_col)).replace(".0", "").strip().upper()
                signals[name] = {
                    "type": "CH",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "IN": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "OUT": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y"))),
                    "ECH": parse_point(cell(row, hi.get("ECH_X")), cell(row, hi.get("ECH_Y"))),
                    "FREE": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y")))
                }
                loaded_ch += 1

    # ---- LC ----
    if "LC" in wb.sheetnames:
        ws_u = wb["LC"]
        hi = header_index_of(ws_u)
        name_col = hi.get("LCGATE")
        if name_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, name_col) is None:
                    continue
                name = str(cell(row, name_col)).strip().upper()
                signals[name] = {
                    "type": "LC",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "in": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "out": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y")))
                }
                loaded_lc += 1

    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_cal} CALLING-ON, {loaded_shunt} SHUNT, "
        f"{loaded_point} POINT, {loaded_ch} CH, {loaded_lc} LC"
    )


def load_universal_coordinates():
    """Explicit button/manual entry point: prompts for a Universal Yard
    Coordinate file and loads it via the same parser load_config() uses
    automatically when it detects the format."""

    global signals
    global config_file_path

    if not routes_data:
        messagebox.showwarning(
            "WARNING",
            "Please load the Sequential Route Release TOC first."
        )
        return

    file_path = filedialog.askopenfilename(
        title="Select Universal Yard Coordinates",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_path:
        return

    config_file_path = file_path
    signals = {}

    wb = load_workbook(file_path)

    _load_universal_signals_from_workbook(wb)

    wb.close()

    log("CONFIG LOADED (UNIVERSAL)")


def save_new_signal_coordinates():

    global config_file_path

    if running or automation_active:
        log("SAVE CONFIG BLOCKED : AUTOMATION RUNNING")
        return

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
# EDRC TOC HEADER SEARCH
# =========================================================

def normalize_toc_header(value):

    if value is None:
        return ""

    text = str(value).strip().upper()

    text = text.replace("_", " ")
    text = text.replace("-", " ")

    return " ".join(text.split())


def find_toc_columns(ws):

    header_aliases = {

        "signal": [
            "SIGNAL"
        ],

        "route": [
            "ROUTE"
        ],

        "first_track": [
            "FIRST_TRACK",
            "FIRST TRACK"
        ],

        "release_sequence": [
            "RELEASE_SEQUENCE",
            "RELEASE SEQUENCE"
        ],

        "back_lock_tracks": [
            "BACK_LOCK_TRACKS",
            "BACK LOCK TRACKS"
        ]
    }

    found = {}

    # Search entire TOC worksheet for headers
    for row in ws.iter_rows():

        for cell in row:

            header = normalize_toc_header(cell.value)

            if not header:
                continue

            for field, aliases in header_aliases.items():

                normalized_aliases = {
                    normalize_toc_header(x)
                    for x in aliases
                }

                if header in normalized_aliases:

                    if field not in found:

                        found[field] = cell.column

                        log(
                            f"TOC HEADER FOUND : "
                            f"{field.upper()} = "
                            f"'{cell.value}' "
                            f"Column {cell.column} "
                            f"({get_column_letter(cell.column)})"
                        )

                    break

    # Required for Sequential
    required = [
        "signal",
        "route",
        "first_track",
        "release_sequence"
    ]

    missing = [
        field
        for field in required
        if field not in found
    ]

    if missing:

        log(
            "MISSING SEQUENTIAL TOC HEADERS : "
            + ", ".join(
                x.upper()
                for x in missing
            )
        )

        return None

    return found

# =========================================================
# LOAD LOCK ROUTE EXCEL
# =========================================================

# =========================================================
# LOAD EDRC TOC FOR SEQUENTIAL ROUTE RELEASE
# =========================================================

def normalize_header(value):
    """
    Normalize Excel headers robustly.

    Handles:
    CALLING_ON_TRACK
    CALLING ON TRACK
    CALLING-ON-TRACK
    CALLING_ON_TRACK with hidden/non-standard spaces
    """

    if value is None:
        return ""

    text = str(value)

    # Remove invisible characters that can come from Excel
    text = (
        text
        .replace("\u00A0", " ")
        .replace("\u200B", "")
        .replace("\u200C", "")
        .replace("\u200D", "")
        .replace("\uFEFF", "")
    )

    # Normalize separators
    text = text.replace("-", "_")
    text = text.replace("/", "_")
    text = text.replace(" ", "_")

    # Remove repeated underscores
    while "__" in text:
        text = text.replace("__", "_")

    return text.strip("_").upper()


def find_header_columns(ws):
    """
    Search the ENTIRE TOC sheet for the required headers.
    Column positions can be anywhere.
    """

    global toc_header_row

    required_headers = [
        "SIGNAL",
        "ROUTE",
        "FIRST_TRACK",
        "RELEASE_SEQUENCE",
        "BACK_LOCK_TRACKS"
    ]

    # -------------------------------------------------
    # SEARCH EVERY CELL IN THE WHOLE WORKSHEET
    # -------------------------------------------------

    found_by_row = {}

    for row in ws.iter_rows():

        row_number = row[0].row

        for cell in row:

            header = normalize_header(cell.value)

            if not header:
                continue

            # Optional: calling-on track column
            # (CALLING_ON_TRACK, or the common TOC "Track" column)
            if header in required_headers or header in (
                "CALLING_ON_TRACK", "TRACK"
            ):

                if row_number not in found_by_row:
                    found_by_row[row_number] = {}

                found_by_row[row_number][header] = cell.column

                log(
                    f"TOC HEADER FOUND : "
                    f"{header} -> "
                    f"ROW {row_number}, "
                    f"COLUMN {get_column_letter(cell.column)}"
                )

    # -------------------------------------------------
    # FIND THE ROW CONTAINING THE MOST HEADERS
    # -------------------------------------------------

    if not found_by_row:

        log("TOC HEADER SEARCH : NO HEADERS FOUND")

        return None

    best_row = max(
        found_by_row,
        key=lambda r: len(found_by_row[r])
    )

    headers = found_by_row[best_row]

    # -------------------------------------------------
    # CHECK REQUIRED HEADERS
    # -------------------------------------------------

    missing = [
        header
        for header in required_headers
        if header not in headers
    ]

    if missing:

        log(
            "MISSING TOC HEADERS : "
            + ", ".join(missing)
        )

        return None

    # -------------------------------------------------
    # SAVE THE HEADER ROW
    # -------------------------------------------------

    toc_header_row = best_row

    log("--------------------------------")
    log(
        f"TOC HEADER ROW FOUND : "
        f"{toc_header_row}"
    )

    log("TOC COLUMN MAPPING:")

    for header in required_headers:

        column = headers[header]

        log(
            f"    {header} -> "
            f"{get_column_letter(column)}"
        )

    return headers


def get_cell_by_header(ws, row_number, headers, header_name):
    """
    Get a cell value by header name.
    """

    normalized = normalize_header(header_name)

    col = headers.get(normalized)

    if col is None:
        return None

    return ws.cell(
        row=row_number,
        column=col
    ).value


def load_routes(file_path=None):

    global routes_data

    # NEVER allow TOC reload while automation is running.
    if running or automation_active:
        log("LOAD TOC BLOCKED : AUTOMATION RUNNING")
        return False
    # IMPORTANT:
    # This function loads ONLY the TOC.
    # It must NEVER modify config_file_path.

    # =====================================================
    # GUI MODE
    # =====================================================

    if file_path is None:

        file_path = filedialog.askopenfilename(
            title="Select TOC Workbook",
            filetypes=[
                ("Excel Files", "*.xlsx")
            ]
        )

        if not file_path:
            return False

    log("--------------------------------")
    log("LOADING TOC")
    log(f"TOC FILE : {file_path}")

    try:

        wb = load_workbook(
            file_path,
            data_only=True
        )

    except Exception as e:

        messagebox.showerror(
            "TOC LOAD ERROR",
            f"Unable to open TOC workbook.\n\n{e}"
        )

        log(f"TOC LOAD ERROR : {e}")

        return False

    # =====================================================
    # FIND TOC SHEET
    # =====================================================

    ws = None

    for sheet_name in wb.sheetnames:

        if normalize_header(sheet_name) == "TOC":

            ws = wb[sheet_name]

            break

    if ws is None:

        messagebox.showerror(
            "TOC ERROR",
            "TOC worksheet was not found."
        )

        log("TOC SHEET NOT FOUND")

        return False

    log(f"TOC SHEET FOUND : {ws.title}")

    # =====================================================
    # FIND COLUMNS BY HEADER NAME
    # =====================================================

    headers = find_header_columns(ws)

    if headers is None:

        log("TOC HEADER SEARCH FAILED - see TOC HEADER FOUND lines above (or lack thereof)")

        messagebox.showerror(
            "TOC HEADER ERROR",
            "Could not find any of the required TOC columns "
            "(SIGNAL, ROUTE, FIRST_TRACK, RELEASE_SEQUENCE, "
            "BACK_LOCK_TRACKS, CALLING_ON_TRACK) anywhere in this sheet."
        )

        return False

    required_headers = [

        "SIGNAL",
        "ROUTE",
        "FIRST_TRACK",
        "RELEASE_SEQUENCE",
        "BACK_LOCK_TRACKS"
    ]

    missing = []

    for header in required_headers:

        if normalize_header(header) not in headers:

            missing.append(header)

    if missing:

        messagebox.showerror(
            "TOC HEADER ERROR",
            "Required TOC columns are missing:\n\n"
            + "\n".join(missing)
        )

        log(
            "MISSING TOC HEADERS : "
            + ", ".join(missing)
        )

        return False

    # =====================================================
    # LOG ACTUAL COLUMN LOCATIONS
    # =====================================================

    log("--------------------------------")
    log("SEQUENTIAL COLUMN MAPPING")

    for header in required_headers:

        col = headers[
            normalize_header(header)
        ]

        log(
            f"{header:20} -> "
            f"{get_column_letter(col)}"
        )

    # =====================================================
    # READ ROUTES
    # =====================================================

    routes_data = []

    for row_number in range(
            toc_header_row + 1,
            ws.max_row + 1
    ):
        signal = get_cell_by_header(
            ws,
            row_number,
            headers,
            "SIGNAL"
        )

        route = get_cell_by_header(
            ws,
            row_number,
            headers,
            "ROUTE"
        )

        first_track = get_cell_by_header(
            ws,
            row_number,
            headers,
            "FIRST_TRACK"
        )

        release_sequence = get_cell_by_header(
            ws,
            row_number,
            headers,
            "RELEASE_SEQUENCE"
        )

        back_lock_tracks = get_cell_by_header(
            ws,
            row_number,
            headers,
            "BACK_LOCK_TRACKS"
        )
        calling_on_track = get_cell_by_header(
            ws,
            row_number,
            headers,
            "CALLING_ON_TRACK"
        )

        # Common TOC has no CALLING_ON_TRACK column - the calling-on
        # track is in the "Track" column (e.g. 1C -> 1CXTPR).
        if calling_on_track is None or not str(calling_on_track).strip():
            calling_on_track = get_cell_by_header(
                ws,
                row_number,
                headers,
                "TRACK"
            )

        # -------------------------------------------------
        # Ignore completely empty rows
        # -------------------------------------------------

        if (
            signal is None
            and route is None
            and first_track is None
            and release_sequence is None
            and back_lock_tracks is None
        ):
            continue

        # -------------------------------------------------
        # Required fields
        # -------------------------------------------------

        # =====================================================
        # REQUIRED FIELDS
        # =====================================================

        if signal is None:
            log(
                f"ROW {row_number} SKIPPED : "
                "SIGNAL missing"
            )
            continue

        signal = (
            str(signal)
            .replace(".0", "")
            .strip()
            .upper()
        )

        # -----------------------------------------------------
        # SIGNAL TYPE
        # -----------------------------------------------------

        if signal.startswith("SH"):
            signal_type = "SHUNT"

        elif signal.endswith("C"):
            signal_type = "CALLING_ON"

        else:
            signal_type = "MAIN"

        # -----------------------------------------------------
        # ROUTE
        # MAIN / SHUNT normally require route.
        # CALLING-ON may have no route.
        # -----------------------------------------------------

        if route is None:
            route = ""

        else:
            route = (
                str(route)
                .replace(".0", "")
                .strip()
                .upper()
            )

        if signal_type != "CALLING_ON" and not route:
            log(
                f"ROW {row_number} SKIPPED : "
                f"{signal_type} SIGNAL HAS NO ROUTE"
            )
            continue

        signal = (
            str(signal)
            .replace(".0", "")
            .strip()
            .upper()
        )

        route = (
            str(route)
            .replace(".0", "")
            .strip()
            .upper()
        )

        first_track = (
            str(first_track).strip().upper()
            if first_track is not None
            else ""
        )

        release_sequence = (
            str(release_sequence).strip().upper()
            if release_sequence is not None
            else ""
        )

        back_lock_tracks = (
            str(back_lock_tracks).strip().upper()
            if back_lock_tracks is not None
            else ""
        )
        calling_on_track = (
            str(calling_on_track).strip().upper()
            if calling_on_track is not None
            else ""
        )

        # =====================================================
        # DETERMINE SIGNAL TYPE
        # =====================================================

        if signal.startswith("SH"):
            signal_type = "SHUNT"

        elif signal.endswith("C"):
            signal_type = "CALLING_ON"

        else:
            signal_type = "MAIN"

        routes_data.append({
            "signal": signal,
            "route": route,
            "signal_type": signal_type,
            "calling_on_track": calling_on_track,
            "first_track": first_track,
            "release_sequence": release_sequence,
            "back_lock_tracks": back_lock_tracks
        })

        log(
            f"Loaded row {row_number} : "
            f"{signal} | "
            f"{route} | "
            f"CALLING_ON_TRACK={calling_on_track} | "
            f"FIRST={first_track} | "
            f"SEQ={release_sequence} | "
            f"BACK={back_lock_tracks}"
        )

    # =====================================================
    # FINISH
    # =====================================================

    refresh_table()

    log("--------------------------------")
    log(
        f"SEQUENTIAL ROUTES LOADED : "
        f"{len(routes_data)}"
    )

    return bool(routes_data)

# =========================================================
# TABLE
# =========================================================

def refresh_table():

    tree.delete(*tree.get_children())

    selected_tab = notebook.tab(notebook.select(), "text")

    count = 1

    for row in routes_data:

        signal = row["signal"].upper()

        # -----------------------------
        # MAIN SIGNALS
        # -----------------------------
        if selected_tab == "MAIN SIGNALS":

            if signal.startswith("SH") or signal.endswith("C"):
                continue

        # -----------------------------
        # CALLING-ON SIGNALS
        # -----------------------------
        elif selected_tab == "CALLING-ON SIGNALS":

            if not signal.endswith("C"):
                continue

        # -----------------------------
        # SHUNT SIGNALS
        # -----------------------------
        elif selected_tab == "SHUNT SIGNALS":

            if not signal.startswith("SH"):
                continue

        tree.insert(
            "",
            "end",
            values=(
                count,
                row["signal"],
                row["route"],
                row["first_track"],
                row["release_sequence"],
                row["back_lock_tracks"]
            )
        )

        count += 1
#refresh_table()

#log("LOCK ROUTES LOADED")
# =========================================================
# MENU ITEM
# =========================================================

def click_menu_item(name):

    try:

        for _ in range(10):      # Wait up to 5 seconds

            item = auto.MenuItemControl(
                searchDepth=15,
                Name=name
            )

            if item.Exists(0.5):

                pause_event.wait()

                item.Click()

                log(f"Clicked : {name}")

                return True

            time.sleep(0.5)

    except Exception as e:

        log(str(e))

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
# CHECK LAMP
# =========================================================

def check_lamp(point, color,screenshot):

    if point is None:
        return False

    x, y = point

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )
    #log(f"{color} -> R={r}, G={g}, B={b}")

    if color == "RED":

        return (
            r > 170 and
            g < 130 and
            b < 130
        )

    elif color == "GREEN":

        return (
            g > 170 and
            r < 150
        )

    elif color == "YELLOW":

        return (
            r > 170 and
            g > 170 and
            b < 130
        )

    return False

def read_lamp_status(signal):

    screenshot = pyautogui.screenshot()

    if signal not in signals:
        return None

    info = signals[signal]

    lamps = {
        "RED": False,
        "YELLOW": False,
        "DOUBLE_YELLOW": False,
        "GREEN": False
    }

    if info.get("RED"):
        lamps["RED"] = check_lamp(info["RED"], "RED", screenshot)

    if info.get("YELLOW"):
        lamps["YELLOW"] = check_lamp(info["YELLOW"], "YELLOW", screenshot)

    if info.get("DOUBLE_YELLOW"):
        lamps["DOUBLE_YELLOW"] = check_lamp(info["DOUBLE_YELLOW"], "YELLOW", screenshot)

    if info.get("GREEN"):
        lamps["GREEN"] = check_lamp(info["GREEN"], "GREEN", screenshot)

    return lamps

# =========================================================
# LAMP CLASSIFIER
# =========================================================

def classify_lamp(samples):

    if not samples:
        return "UNKNOWN"

    # Always ON
    if all(samples):
        return "STABLE_ON"

    # Always OFF
    if not any(samples):
        return "STABLE_OFF"

    # Changed during observation
    return "CHANGING"


# =========================================================
# COUNT STATE CHANGES
# =========================================================

def count_transitions(samples):

    if len(samples) < 2:
        return 0

    count = 0

    for i in range(1, len(samples)):

        if samples[i] != samples[i - 1]:
            count += 1

    return count



# =========================================================
# READ SIGNAL ASPECT
# =========================================================
def get_stable_signal_aspect(signal, observe_time=1.5, sample_interval=0.1):

    red = []
    yellow = []
    dy = []
    green = []

    start = time.time()

    while time.time() - start < observe_time:

        lamps = read_lamp_status(signal)

        if lamps is not None:
            log(f"{signal} RAW LAMPS : {lamps}")

            red.append(lamps["RED"])
            yellow.append(lamps["YELLOW"])
            dy.append(lamps["DOUBLE_YELLOW"])
            green.append(lamps["GREEN"])

        time.sleep(sample_interval)

    log(f"{signal} RED Samples    : {red}")
    log(f"{signal} YELLOW Samples : {yellow}")
    log(f"{signal} DY Samples     : {dy}")
    log(f"{signal} GREEN Samples  : {green}")

    red_state = classify_lamp(red)
    yellow_state = classify_lamp(yellow)
    dy_state = classify_lamp(dy)
    green_state = classify_lamp(green)

    log(f"{signal} RED    = {red_state}")
    log(f"{signal} YELLOW = {yellow_state}")
    log(f"{signal} DY     = {dy_state}")
    log(f"{signal} GREEN  = {green_state}")

    # --------------------------------------------
    # Detect blinking lamps
    # --------------------------------------------

    blinking = []

    if red_state == "CHANGING" and count_transitions(red) >= 2:
        blinking.append("RED")

    if yellow_state == "CHANGING" and count_transitions(yellow) >= 2:
        blinking.append("YELLOW")

    if dy_state == "CHANGING" and count_transitions(dy) >= 2:
        blinking.append("DOUBLE_YELLOW")

    if green_state == "CHANGING" and count_transitions(green) >= 2:
        blinking.append("GREEN")

    log(f"{signal} Blinking Lamps : {blinking}")

    # --------------------------------------------
    # Stable lamps
    # --------------------------------------------

    stable = []

    if red_state == "STABLE_ON":
        stable.append("RED")

    if yellow_state == "STABLE_ON":
        stable.append("YELLOW")

    if dy_state == "STABLE_ON":
        stable.append("DOUBLE_YELLOW")

    if green_state == "STABLE_ON":
        stable.append("GREEN")

    log(f"{signal} Stable Lamps : {stable}")

    # --------------------------------------------
    # Business Rules
    # --------------------------------------------

    # Rule 1 : Stable colour + blinking colour
    #          -> Stable colour wins

    if stable:

        if "RED" in stable:
            return "RED"

        if "DOUBLE_YELLOW" in stable:
            return "DOUBLE_YELLOW"

        if "YELLOW" in stable:
            return "YELLOW"

        if "GREEN" in stable:
            return "GREEN"

    # Rule 2 : Only blinking lamps

    if blinking:
        return "BLANK"

    # Rule 3 : All lamps OFF

    return "BLANK"

# =========================================================
# SELECT CONTROL BIT
# =========================================================

def select_control_bit(control_bit):

    log(f"SETTING CONTROL BIT : {control_bit}")

    if not open_bit_chart():
        return False

    if not find_and_click(control_bit):
        return False

    time.sleep(1)

    if not find_and_click(
        "TRANSMIT",
        control_type="Button"
    ):
        return False

    time.sleep(8)

    if not find_and_click(
            "Cancel",
            control_type="Button"
    ):
        log("Cancel Button Not Found")
        # Don't fail the test because of this

    time.sleep(5)

    return True
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

    screenshot = pyautogui.screenshot()

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
        return False

    screenshot = pyautogui.screenshot()

    x, y = point

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    log(f"TRACK RGB = {r},{g},{b}")

    # Yellow route indication
    if r > 180 and g > 180 and b < 120:
        return True

    return False

# =========================================================
# SHUNT STATUS
# =========================================================

def shunt_fail(signal):

    info = signals[signal]

    c1 = is_yellow(info["C1"])

    c2 = is_yellow(info["C2"])

    c3 = is_yellow(info["C3"])

    # FAIL
    if (not c1) and c2 and c3:

        return True

    return False

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

    time.sleep(0.2)

    pyautogui.hotkey("ctrl", "b")

    log("CTRL+B Pressed")

    time.sleep(0.5)

    found = find_and_click(
        "50051",
        control_type="Text"
    )

    if not found:

        log("50051 Not Found")

        return False

    time.sleep(0.2)

    ok = find_and_click(
        "OK",
        control_type="Button"
    )

    if not ok:

        log("OK Button Not Found")

        return False

    time.sleep(0.2)

    return True

def open_control_bit_window():

    log("OPENING CONTROL BIT WINDOW")

    return open_bit_chart()

def transmit_control_bit(control_bit):

    log(f"SETTING CONTROL BIT : {control_bit}")

    if not find_and_click(control_bit):
        return False

    time.sleep(1)

    if not find_and_click(
        "TRANSMIT",
        control_type="Button"
    ):
        return False

    time.sleep(8)

    return True

def close_control_bit_window():

    if find_and_click(
        "Cancel",
        control_type="Button"
    ):
        log("CONTROL BIT WINDOW CLOSED")
    else:
        log("Cancel Button Not Found")

def err_open_control_bit_window():

    log("OPENING CONTROL BIT WINDOW")

    return open_bit_chart()


def err_transmit_control_bit(control_bit):

    log(f"SETTING CONTROL BIT : {control_bit}")

    if not find_and_click(control_bit):
        return False

    time.sleep(0.1)

    pyautogui.press("enter")

    time.sleep(0.2)

    return True


def err_close_control_bit_window():

    if not find_and_click(
            "Cancel",
            control_type="Button"
    ):
        log("Cancel Button Not Found")

    return True
# =========================================================
# DROP TRACK
# =========================================================

def drop_track(track):
    pause_event.wait()
    log(f"DROPPING TRACK : {track}")

    opened = open_bit_chart()

    if not opened:
        return False

    track_found = find_and_click(track)

    if not track_found:

        log(f"{track} Not Found")

        return False

    time.sleep(1)

    transmit_found = find_and_click(
        "TRANSMIT",
        control_type="Button"
    )

    if not transmit_found:

        log("TRANSMIT Not Found")

        return False

    time.sleep(1)

    cancel_found = find_and_click(
        "CANCEL",
        control_type="Button"
    )

    if not cancel_found:

        log("CANCEL Not Found")

        return False

    log(f"{track} DROP/UP SUCCESS")

    time.sleep(5)

    return True

def save_long_train_report():

    global report_saved

    if report_saved:
        return

    if not report_rows:
        log("NO RESULTS TO SAVE")
        return

    create_report(
        REPORT_FILE,
        "LONG TRAIN"
    )

    report_saved = True

    log("REPORT GENERATED")

def save_short_train_report():
    log("save_short_train_report() called")

    if not report_rows:
        log("No report rows to save")
        return

    create_report(
        REPORT_FILE,
        "SHORT TRAIN"
    )

    log(f"Short Train report saved : {REPORT_FILE}")

def save_err_report():

    global report_saved

    log("save_err_report() called")

    log(f"report_saved = {report_saved}")
    log(f"Rows = {len(report_rows)}")

    if report_saved:
        log("Already saved")
        return

    if not report_rows:
        log("No report rows")
        return

    create_report(
        REPORT_FILE,
        "EMERGENCY ROUTE RELEASE"
    )

    report_saved = True

    log(f"ERR report saved : {REPORT_FILE}")

# =========================================================
# CREATE ADVANCED REPORT
# =========================================================

def create_report(filename, report_name):

    wb = Workbook()

    ws = wb.active

    ws.title = report_name

    if os.path.exists(filename):

        wb = load_workbook(filename)

        if report_name in wb.sheetnames:
            del wb[report_name]

        ws = wb.create_sheet(report_name)

    else:

        wb = Workbook()

        ws = wb.active

        ws.title = report_name

    # =====================================================
    # MAIN TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title_cell = ws["A1"]

    title_cell.value = "SEQUENTIAL ROUTE RELEASE TEST"

    title_cell.font = Font(
        bold=True,
        size=18,
        color="000000"  # black text
    )

    title_cell.fill = PatternFill(fill_type=None)  # no background

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 35

    # =====================================================
    # SUB TITLE
    # =====================================================

    ws.merge_cells("A2:D2")

    subtitle = ws["A2"]

    subtitle.value = report_name.upper()

    subtitle.font = Font(
        bold=True,
        size=14,
        color="1E293B"
    )

    subtitle.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[2].height = 25

    # =====================================================
    # DATE
    # =====================================================

    ws.merge_cells("A3:D3")

    date_cell = ws["A3"]

    date_cell.value = f"Generated : {datetime.now():%d-%m-%Y %H:%M:%S}"

    date_cell.font = Font(
        bold=True,
        size=11
    )

    date_cell.alignment = Alignment(horizontal="center")

    ws.row_dimensions[3].height = 22

    # =====================================================
    # HEADERS
    # =====================================================

    headers = [

        "MAIN SIGNAL",

        "MAIN ROUTE",

        "RESULT",

        "DATE & TIME"

    ]
    ws.append([])
    ws.append(headers)

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

    for cell in ws[5]:
        cell.font = Font(
            bold=True,
            color="000000"  # black text
        )

        cell.fill = PatternFill(fill_type=None)  # no color

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # =====================================================
    # DATA
    # =====================================================

    for row in report_rows:
        ws.append([

            row["MAIN_SIGNAL"],
            row["MAIN_ROUTE"],
            row["RESULT"],
            row["DATE & TIME"]

        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================

    red_fill = PatternFill(
        fill_type="solid",
        fgColor="FF0000"
    )

    for row in ws.iter_rows(min_row=6):

        result = str(row[2].value).strip().upper()

        for cell in row:
            # Default formatting
            cell.font = Font(
                color="000000",
                bold=False
            )

            cell.fill = PatternFill(fill_type=None)

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # FAIL -> Entire row red
        if result == "FAIL":

            for cell in row:
                cell.fill = red_fill
                cell.font = Font(
                    color="FFFFFF",
                    bold=True
                )

    # =====================================================
    # AUTO WIDTH
    # =====================================================

    for column in ws.columns:

        max_length = 0

        column_letter = get_column_letter(column[0].column)

        for cell in column:

            try:

                if len(str(cell.value)) > max_length:

                    max_length = len(str(cell.value))

            except:
                pass

        adjusted = max_length + 5

        ws.column_dimensions[column_letter].width = adjusted

    # =====================================================
    # ROUTE SUMMARY
    # =====================================================

    # Unique routes
    all_routes = set([x["MAIN_ROUTE"] for x in report_rows])

    # PASS routes
    initiated_routes = set([
        x["MAIN_ROUTE"] for x in report_rows
        if x["RESULT"] == "PASS"
    ])

    # FAIL routes
    failed_routes = set([
        x["MAIN_ROUTE"] for x in report_rows
        if x["RESULT"] == "FAIL"
    ])

    # Non-initiated = failed
    non_initiated_routes = failed_routes

    start = ws.max_row + 3

    # TOTAL ROUTES COUNT
    ws[f"A{start}"] = "TOTAL ROUTES"
    ws[f"B{start}"] = len(all_routes)

    # INITIATED ROUTES (PASS)
    ws[f"A{start + 1}"] = "INITIATED ROUTES"
    ws[f"B{start + 1}"] = ", ".join(sorted(initiated_routes)) if initiated_routes else "0"

    # NON-INITIATED ROUTES (FAIL)
    ws[f"A{start + 2}"] = "NON-INITIATED ROUTES"
    ws[f"B{start + 2}"] = ", ".join(sorted(non_initiated_routes)) if non_initiated_routes else "0"

    # Bold labels
    for r in range(start, start + 3):
        ws[f"A{r}"].font = Font(bold=True)

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(filename)

    log(f"REPORT SAVED : {filename}")

# =========================================================
# BRING TEST PANEL TO FRONT
# =========================================================

def bring_test_panel_to_front():

    try:

        desktop = Desktop(backend="uia")

        for win in desktop.windows():

            try:

                title = win.window_text().strip()

                if title.upper() == "TEST PANEL":

                    log("TEST PANEL FOUND")

                    # Restore if minimized
                    try:
                        win.restore()
                    except:
                        pass

                    time.sleep(0.5)

                    # Set focus
                    try:
                        win.set_focus()
                    except:
                        pass

                    time.sleep(0.5)

                    log("TEST PANEL BROUGHT TO FRONT")

                    return True

            except Exception:
                pass

    except Exception as e:

        log(
            f"TEST PANEL FOCUS ERROR : {e}"
        )

    log("TEST PANEL NOT FOUND")

    return False

# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():

    global running

    root.iconify()

    time.sleep(2)

    for row in cascading_routes_data:

        pause_event.wait()

        if not running:
            return

        # -----------------------------------------
        # TOC DATA
        # -----------------------------------------

        main_signal = row["signal"]

        main_route = row["route"]

        cascading_signal = row["cascading_signal"]

        cascading_route = row["cascading_route"]

        extra_signal = row["extra_signal"]

        extra_route = row["extra_route"]

        control_bits = [

            x.strip()

            for x in row["cascading"].split(",")

            if x.strip()

        ]

        log(
            f"MAIN : {main_signal}  ROUTE : {main_route}"
        )

        log(
            f"CASCADING : {cascading_signal}  ROUTE : {cascading_route}"
        )

        # -----------------------------------------
        # CHECK SIGNALS
        # -----------------------------------------

        if main_signal not in signals:
            log(f"{main_signal} NOT FOUND")

            continue

        if cascading_signal not in signals:
            log(f"{cascading_signal} NOT FOUND")

            continue

        # -----------------------------------------
        # SET MAIN ROUTE
        # -----------------------------------------

        click(signals[main_signal]["menu"])

        time.sleep(2)

        ok = click_menu_item(
            main_route.replace("-", "_")
        )

        if not ok:
            log(f"FAILED TO CLICK {main_route}")
            continue

        log(f"{main_signal} {main_route} SET")

        time.sleep(5)

        log("Going to click cascading signal")

        special_case = (
                main_signal == "1" and
                main_route == "1_A1" and
                cascading_signal == "3" and
                cascading_route == "3_K"
        )

        # -----------------------------------------
        # SET CASCADING ROUTE
        # -----------------------------------------

        if not special_case:

            click(signals[cascading_signal]["menu"])

            log("Cascading signal clicked")

            time.sleep(2)

            ok = click_menu_item(
                cascading_route.replace("-", "_")
            )

            if not ok:
                log(f"FAILED TO CLICK {cascading_route}")
                continue

            log(f"{cascading_signal} {cascading_route} SET")

            time.sleep(5)

            log(f"Lamp Status: {read_lamp_status(cascading_signal)}")

            log(">>> Before MAIN stable")
            main_actual = get_stable_signal_aspect(main_signal)
            log(">>> MAIN stable completed")

            log(">>> Before CASCADE stable")
            cascade_actual = get_stable_signal_aspect(cascading_signal)
            log(">>> CASCADE stable completed")

            log(">>> Going to CONTROL BIT LOOP")

            log(f"{main_signal} : {main_actual}")
            log(f"{cascading_signal} : {cascade_actual}")

        else:

            log("SPECIAL CASE : Skipping Cascading Route 3_K")

        # -----------------------------------------
        # SET EXTRA ROUTE
        # -----------------------------------------

        if extra_signal and extra_route:

            click(signals[extra_signal]["menu"])

            time.sleep(2)

            ok = click_menu_item(
                extra_route.replace("-", "_")
            )

            if not ok:
                log(f"FAILED TO CLICK {extra_route}")
                continue

            log(f"{extra_signal} {extra_route} SET")

            time.sleep(5)

            main_actual = get_stable_signal_aspect(main_signal)
            cascade_actual = get_stable_signal_aspect(cascading_signal)

            log(f"{main_signal} : {main_actual}")
            log(f"{cascading_signal} : {cascade_actual}")

        # =================================================
        # CASCADING CONTROL BIT TEST
        # =================================================

        for control_bit in control_bits:

            pause_event.wait()

            if not running:
                return

            log("--------------------------------")
            log(f"CONTROL BIT : {control_bit}")

            # Read current cascade aspect before operating control bit
            ok = select_control_bit(control_bit)

            if not ok:

                result = "FAIL"
                main_actual = "NOT SET"
                cascade_actual = "NOT SET"

            else:

                # Wait for simulator to settle
                log("Waiting for signal to stabilise...")
                time.sleep(5)

                log("Reading stable cascade aspect...")
                cascade_actual = get_stable_signal_aspect(cascading_signal)

                log("Reading stable main aspect...")
                main_actual = get_stable_signal_aspect(main_signal)

                log(f"CASCADE ACTUAL : {cascade_actual}")
                log(f"MAIN ACTUAL : {main_actual}")

                # Your PASS/FAIL logic comes here

                if special_case:

                    # Signal 3 will be blank, which read_signal_aspect() returns as BLANK
                    if main_actual == "RED" and cascade_actual == "BLANK":
                        result = "PASS"
                    else:
                        result = "FAIL"


                else:

                    # Cascading signal is YELLOW

                    if cascade_actual == "YELLOW":

                        # Extra Route condition
                        if extra_signal and extra_route:

                            if main_actual == "YELLOW":

                                result = "PASS"

                            else:

                                result = "FAIL"

                        # Normal Route condition
                        else:

                            if main_actual == "RED":

                                result = "PASS"

                            else:

                                result = "FAIL"

                    # Cascading signal is RED

                    elif cascade_actual == "RED":

                        if main_actual == "YELLOW":

                            result = "PASS"

                        else:

                            result = "FAIL"


                    # Cascading signal is BLANK

                    elif cascade_actual == "BLANK":

                        if main_actual == "RED":

                            result = "PASS"

                        else:

                            result = "FAIL"


                    else:

                        result = "FAIL"

            report_rows.append({

                "MAIN_SIGNAL": main_signal,

                "MAIN_ROUTE": main_route,

                "CASCADING_SIGNAL": cascading_signal,

                "CASCADING_ROUTE": cascading_route,

                "CONTROL_BIT": control_bit,

                "ACTUAL_MAIN": main_actual,

                "ACTUAL_CASCADING": cascade_actual,

                "RESULT": result,

                "DATE & TIME":
                    datetime.now().strftime(
                        "%d-%m-%Y %H:%M:%S"
                    )

            })

            log(f"{control_bit} -> {result}")


        # =============================================
        # CASCADING COMPLETED
        # =============================================

        log(
            f"{main_signal} CASCADING TEST COMPLETED"
        )

        # -----------------------------------------
        # RESTORE ORIGINAL SIGNAL STATE
        # -----------------------------------------

        # -----------------------------------------
        # RESTORE FIRST CONTROL BIT ONLY
        # -----------------------------------------

        log("RESTORING FIRST CONTROL BIT")

        if control_bits:

            first_control_bit = control_bits[0]

            log(f"RESTORE : {first_control_bit}")

            if not select_control_bit(first_control_bit):
                log(f"Failed to restore {first_control_bit}")

            log("Waiting after restore...")
            time.sleep(5)

        main_actual = get_stable_signal_aspect(main_signal)
        cascade_actual = get_stable_signal_aspect(cascading_signal)

        log(f"RESTORED MAIN : {main_actual}")
        log(f"RESTORED CASCADE : {cascade_actual}")

        # =============================================
        # CANCEL CASCADING SIGNAL
        # =============================================

        click(signals[cascading_signal]["menu"])

        time.sleep(1)

        click_menu_item("Signal Cancel")

        log(f"{cascading_signal} SIGNAL CANCEL DONE")

        time.sleep(5)

        click(signals[cascading_signal]["menu"])

        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{cascading_signal} ROUTE RELEASE DONE")

        time.sleep(15)

        # =============================================
        # CANCEL EXTRA SIGNAL
        # =============================================

        if extra_signal and extra_route:
            click(signals[extra_signal]["menu"])

            time.sleep(1)

            click_menu_item("Signal Cancel")

            log(f"{extra_signal} SIGNAL CANCEL DONE")

            time.sleep(5)

            click(signals[extra_signal]["menu"])

            time.sleep(1)

            click_menu_item("Route Release")

            log(f"{extra_signal} ROUTE RELEASE DONE")

            time.sleep(15)

        # =============================================
        # CANCEL MAIN SIGNAL
        # =============================================

        click(signals[main_signal]["menu"])

        time.sleep(1)

        click_menu_item("Signal Cancel")

        log(f"{main_signal} SIGNAL CANCEL DONE")

        time.sleep(5)

        click(signals[main_signal]["menu"])

        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{main_signal} ROUTE RELEASE DONE")

        time.sleep(15)

        log("Checking signals before next route...")

        main_actual = get_stable_signal_aspect(main_signal)
        cascade_actual = get_stable_signal_aspect(cascading_signal)

        log(f"Before Next Route MAIN : {main_actual}")
        log(f"Before Next Route CASCADE : {cascade_actual}")

        time.sleep(5)

    # =========================================================
    # AUTOMATION COMPLETED
    # =========================================================


    running = False

    status_label.config(
        text="STOPPED",
        fg="red"
    )

    log("AUTOMATION COMPLETED")


# =========================================================
# SEQUENTIAL ROUTE RELEASE ENGINE
# =========================================================

def normalize_track_control(value):
    """Map worksheet track names (25XT) to Control Bit names (25XTPR)."""
    control_bit = str(value).strip().upper()
    return f"{control_bit}PR" if control_bit.endswith("XT") else control_bit

def prepare_calling_on_route(row):
    """
    CALLING-ON INITIAL ROUTE SETTING

    Every Calling-On route MUST do:

    1. Open Control Bit
    2. Calling-On Track DOWN
    3. Open Calling-On signal menu
    4. Set Calling-On route
    5. Calling-On Track NORMAL
    6. Close Control Bit

    This is required BEFORE Short Train / Long Train / ERR.
    """

    signal = row["signal"]
    route = row["route"]

    calling_on_track = normalize_track_control(
        row.get("calling_on_track", "")
    )

    if not calling_on_track:
        log(
            f"{signal} : CALLING-ON TRACK MISSING"
        )
        return False

    if signal not in signals:
        log(
            f"{signal} : CALLING-ON SIGNAL NOT FOUND"
        )
        return False

    log("--------------------------------")
    log(
        f"CALLING-ON INITIAL PREPARATION : "
        f"{signal} -> {route}"
    )

    # =================================================
    # STEP 1 : OPEN CONTROL BIT
    # =================================================

    log(
        "CALLING-ON : OPEN CONTROL BIT WINDOW"
    )

    if not open_control_bit_window():

        log(
            "CALLING-ON : CONTROL BIT WINDOW FAILED"
        )

        return False

    # =================================================
    # STEP 2 : CALLING-ON TRACK DOWN
    # =================================================

    log(
        f"CALLING-ON : "
        f"{calling_on_track} DOWN"
    )

    if not transmit_control_bit(
        calling_on_track
    ):

        log(
            f"CALLING-ON : "
            f"{calling_on_track} DOWN FAILED"
        )

        close_control_bit_window()

        return False

    time.sleep(1)

    # =================================================
    # STEP 3 : OPEN CALLING-ON SIGNAL MENU
    # =================================================

    log(
        f"CALLING-ON : OPEN MENU {signal}"
    )

    if not click(
        signals[signal]["menu"]
    ):

        log(
            f"CALLING-ON : "
            f"{signal} MENU CLICK FAILED"
        )

        close_control_bit_window()

        return False

    time.sleep(0.5)

    # =================================================
    # STEP 4 : SET CALLING-ON ROUTE
    # =================================================

    log(
        f"CALLING-ON : SET ROUTE {route}"
    )

    if not click_menu_item(
        route.replace("-", "_")
    ):

        log(
            f"CALLING-ON : "
            f"ROUTE NOT SET : {route}"
        )

        close_control_bit_window()

        return False

    log(
        f"CALLING-ON : "
        f"{route} SET"
    )

    time.sleep(2)

    # =================================================
    # STEP 5 : CALLING-ON TRACK NORMAL
    # =================================================

    log(
        f"CALLING-ON : "
        f"{calling_on_track} NORMAL"
    )

    if not transmit_control_bit(
        calling_on_track
    ):

        log(
            f"CALLING-ON : "
            f"{calling_on_track} NORMAL FAILED"
        )

        close_control_bit_window()

        return False

    time.sleep(1)

    # =================================================
    # STEP 6 : CLOSE CONTROL BIT
    # =================================================

    close_control_bit_window()

    log(
        f"CALLING-ON INITIAL PREPARATION COMPLETE : "
        f"{signal} -> {route}"
    )

    return True

def split_control_bits(value):
    """Return ordered Control Bit names from a comma, >, or newline list."""
    return [
        normalize_track_control(item)
        for item in str(value).replace(">", ",").replace("\n", ",").split(",")
        if item and item.strip()
    ]


def attempt_manual_release(signal):

    click(signals[signal]["menu"])

    time.sleep(0.2)

    if not click_menu_item("Signal Cancel"):
        log("Signal Cancel menu item not found")
        return False

    time.sleep(0.2)

    click(signals[signal]["menu"])

    time.sleep(0.2)

    if not click_menu_item("Route Release"):
        log("Route Release menu item not found")
        return False

    time.sleep(0.2)

    return True

def emergency_route_release(signal, bit_chart_open=False):

    log("STARTING EMERGENCY ROUTE RELEASE")

    # =====================================================
    # STEP 1 : SIGNAL CANCEL + ROUTE RELEASE
    # =====================================================

    if not attempt_manual_release(signal):

        log("MANUAL RELEASE FAILED")

        return False

    # =====================================================
    # STEP 2 : OPEN CONTROL BIT WINDOW
    # =====================================================

    if not bit_chart_open:

        if not err_open_control_bit_window():
            log("CONTROL BIT WINDOW FAILED")
            return False

    # =====================================================
    # STEP 3 : EMKEY_IN (TRANSMIT)
    # =====================================================

    if not err_transmit_control_bit("EMKEY_IN"):

        log("EMKEY_IN TRANSMIT FAILED")

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 4 : EMERGENCY ROUTE RELEASE
    # =====================================================

    for attempt in range(5):

        click(signals[signal]["menu"])

        time.sleep(0.5)

        if click_menu_item("Emergency Route Release"):

            break

        log(f"Retry Emergency Route Release ({attempt + 1}/5)")

    else:

        log("Emergency Route Release NOT FOUND")

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 5 : PASSWORD DIALOG
    # =====================================================

    try:

        log("Waiting for Password dialog...")

        dlg = auto.WindowControl(
            searchDepth=5,
            Name="Password for Emergency Route Release"
        )

        if not dlg.Exists(10):
            log("Password dialog NOT FOUND")

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

        log("Password dialog found")

        edits = dlg.GetChildren()

        edit_boxes = [
            c for c in edits
            if c.ControlTypeName == "EditControl"
        ]

        if len(edit_boxes) < 2:

            log(f"Found only {len(edit_boxes)} edit boxes")

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

        edit_boxes[0].Click()
        edit_boxes[0].SendKeys("ETOE")

        time.sleep(0.2)

        edit_boxes[1].Click()
        edit_boxes[1].SendKeys("ETOE")

        time.sleep(0.2)

        ok = dlg.ButtonControl(Name="OK")

        if ok.Exists(2):

            ok.Click()

            log("Password entered")

        else:

            log("OK button not found")

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

    except Exception as e:

        log(f"Password dialog failed : {e}")

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 6 : EMKEY_IN AGAIN (TRANSMIT)
    # =====================================================

    if not err_transmit_control_bit("EMKEY_IN"):

        log("SECOND EMKEY_IN FAILED")

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 7 : CLOSE CONTROL BIT WINDOW
    # =====================================================

    if not bit_chart_open:
        err_close_control_bit_window()

    log("Waiting for route to release...")

    time.sleep(5)

    log("EMERGENCY ROUTE RELEASE COMPLETED")

    return True

# =========================================================
# SHUNT EMERGENCY KEY SEQUENCE
# =========================================================
# NOTE:
# This function starts AFTER Signal Cancel + Route Release.
# SHUNT ERR already performs those operations for every
# back-lock track before reaching this function.
# =========================================================

def shunt_emergency_key_sequence(signal, bit_chart_open=False):

    log("STARTING SHUNT EMERGENCY ROUTE RELEASE")

    # =====================================================
    # STEP 1 : OPEN CONTROL BIT WINDOW
    # =====================================================

    if not bit_chart_open:

        if not err_open_control_bit_window():

            log(
                "SHUNT CONTROL BIT WINDOW FAILED"
            )

            return False

    # =====================================================
    # STEP 2 : EMKEY_IN
    # =====================================================

    log(
        "SHUNT : SETTING EMKEY_IN"
    )

    if not err_transmit_control_bit("EMKEY_IN"):

        log(
            "SHUNT : EMKEY_IN TRANSMIT FAILED"
        )

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 3 : SHUNT MENU
    # =====================================================

    for attempt in range(5):

        click(
            signals[signal]["menu"]
        )

        time.sleep(0.5)

        if click_menu_item(
            "Emergency Route Release"
        ):

            log(
                "SHUNT : EMERGENCY ROUTE RELEASE "
                "MENU CLICKED"
            )

            break

        log(
            f"SHUNT : RETRY EMERGENCY ROUTE RELEASE "
            f"({attempt + 1}/5)"
        )

    else:

        log(
            "SHUNT : EMERGENCY ROUTE RELEASE "
            "MENU NOT FOUND"
        )

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 4 : PASSWORD DIALOG
    # =====================================================

    try:

        log(
            "SHUNT : WAITING FOR PASSWORD DIALOG..."
        )

        dlg = auto.WindowControl(
            searchDepth=5,
            Name="Password for Emergency Route Release"
        )

        if not dlg.Exists(10):

            log(
                "SHUNT : PASSWORD DIALOG NOT FOUND"
            )

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

        log(
            "SHUNT : PASSWORD DIALOG FOUND"
        )

        edits = dlg.GetChildren()

        edit_boxes = [
            c
            for c in edits
            if c.ControlTypeName == "EditControl"
        ]

        if len(edit_boxes) < 2:

            log(
                f"SHUNT : FOUND ONLY "
                f"{len(edit_boxes)} EDIT BOXES"
            )

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

        # -------------------------------------------------
        # USERNAME
        # -------------------------------------------------

        edit_boxes[0].Click()

        edit_boxes[0].SendKeys(
            "ETOE"
        )

        time.sleep(0.2)

        # -------------------------------------------------
        # PASSWORD
        # -------------------------------------------------

        edit_boxes[1].Click()

        edit_boxes[1].SendKeys(
            "ETOE"
        )

        time.sleep(0.2)

        # -------------------------------------------------
        # OK
        # -------------------------------------------------

        ok = dlg.ButtonControl(
            Name="OK"
        )

        if ok.Exists(2):

            ok.Click()

            log(
                "SHUNT : USERNAME/PASSWORD ENTERED"
            )

        else:

            log(
                "SHUNT : OK BUTTON NOT FOUND"
            )

            if not bit_chart_open:
                err_close_control_bit_window()

            return False

    except Exception as e:

        log(
            f"SHUNT : PASSWORD DIALOG FAILED : {e}"
        )

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 5 : EMKEY_IN AGAIN
    # =====================================================

    log(
        "SHUNT : SETTING EMKEY_IN AGAIN"
    )

    if not err_transmit_control_bit(
        "EMKEY_IN"
    ):

        log(
            "SHUNT : SECOND EMKEY_IN FAILED"
        )

        if not bit_chart_open:
            err_close_control_bit_window()

        return False

    # =====================================================
    # STEP 6 : CLOSE CONTROL BIT WINDOW
    # =====================================================

    if not bit_chart_open:

        err_close_control_bit_window()

    log(
        "SHUNT EMERGENCY ROUTE RELEASE COMPLETED"
    )

    return True

# =========================================================
# CHECK WHETHER ROUTE IS RELEASED
# =========================================================

def is_route_released(signal):

    # =====================================================
    # VALIDATE SIGNAL
    # =====================================================

    if signal not in signals:

        log(f"{signal} NOT FOUND")

        return False

    info = signals[signal]

    # =====================================================
    # SELECT ROUTE STATUS POINT
    #
    # MAIN       -> ROUTE_INDICATOR
    # SHUNT      -> route_init
    # CALLING_ON -> route_init
    # =====================================================

    signal_type = info.get("type", "MAIN")

    if signal_type in ("SHUNT", "CALLING_ON"):

        point = info.get("route_init")

        if point is None:

            log(
                f"{signal} {signal_type} "
                f"ROUTE INIT NOT CONFIGURED"
            )

            return False

    else:

        point = info.get("ROUTE_INDICATOR")

        if point is None:

            log(
                f"{signal} ROUTE INDICATOR "
                f"NOT CONFIGURED"
            )

            return False

    # =====================================================
    # TAKE SCREENSHOT
    # =====================================================

    screenshot = pyautogui.screenshot()

    # =====================================================
    # GET ROUTE STATUS PIXEL
    # =====================================================

    x, y = point

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    log(
        f"{signal} ROUTE STATUS RGB : "
        f"{r},{g},{b}"
    )

    # =====================================================
    # YELLOW = ROUTE STILL LOCKED
    # =====================================================

    if (
        r > 180
        and g > 180
        and b < 120
    ):

        log(
            f"{signal} ROUTE STILL LOCKED"
        )

        return False

    # =====================================================
    # NOT YELLOW = ROUTE RELEASED
    # =====================================================

    log(
        f"{signal} ROUTE RELEASED"
    )

    return True

def append_long_train_result(row, result, actual_aspect, detail):

    report_rows.append({

        "MAIN_SIGNAL": row["signal"],

        "MAIN_ROUTE": row["route"],

        "CASCADING_SIGNAL": "LONG TRAIN",

        "CASCADING_ROUTE": "",

        "CONTROL_BIT": row["release_sequence"],

        "ACTUAL_MAIN": actual_aspect,

        "ACTUAL_CASCADING": detail,

        "RESULT": result,

        "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    })

def append_short_train_result(row, result, actual_aspect, detail):

    report_rows.append({

        "MAIN_SIGNAL": row["signal"],

        "MAIN_ROUTE": row["route"],

        "CASCADING_SIGNAL": "SHORT TRAIN",

        "CASCADING_ROUTE": "",

        "CONTROL_BIT": row["first_track"],

        "ACTUAL_MAIN": actual_aspect,

        "ACTUAL_CASCADING": detail,

        "RESULT": result,

        "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    })

def append_err_result(row, result, actual_aspect, detail):

    report_rows.append({

        "MAIN_SIGNAL": row["signal"],

        "MAIN_ROUTE": row["route"],

        "CASCADING_SIGNAL": "ERR",

        "CASCADING_ROUTE": "",

        "CONTROL_BIT": row["back_lock_tracks"],

        "ACTUAL_MAIN": actual_aspect,

        "ACTUAL_CASCADING": detail,

        "RESULT": result,

        "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    })

def run_long_train_engine():
    """Long Train Control Bit Release Automation"""

    global running

    root.iconify()
    time.sleep(0.5)

    for row in routes_data:
        signal_type = row.get(
            "signal_type",
            "MAIN"
        )

        # MAIN, SHUNT and CALLING-ON all run Long Train.
        # No signal type is skipped here.

        pause_event.wait()

        if not running or not automation_active:
            close_control_bit_window()
            return

        if not signals:
            log("FATAL : SIGNAL CONFIGURATION DISAPPEARED")
            return

        if not routes_data:
            log("FATAL : ROUTE DATA DISAPPEARED")
            return

        signal = row["signal"]
        route = row["route"]

        first_track = normalize_track_control(
            row["first_track"]
        )

        release_sequence = split_control_bits(
            row["release_sequence"]
        )

        back_lock_tracks = split_control_bits(
            row["back_lock_tracks"]
        )

        if signal not in signals:
            log(f"{signal} NOT FOUND")
            continue

        # =================================================
        # NOT APPLICABLE
        # No Sequential Route Release configuration
        # =================================================

        if (
                not first_track
                and not release_sequence
                and not back_lock_tracks
        ):
            log(
                f"{signal} {route} : "
                f"NOT APPLICABLE - NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            append_long_train_result(
                row,
                "NOT APPLICABLE",
                "N/A",
                "NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            continue

        # =================================================
        # INVALID / INCOMPLETE CONFIGURATION
        # =================================================

        if not release_sequence:
            log(
                f"{signal} {route} "
                f"HAS INVALID RELEASE SEQUENCE"
            )

            append_long_train_result(
                row,
                "FAIL",
                "NOT SET",
                "INVALID RELEASE SEQUENCE"
            )

            continue

        log("--------------------------------")
        log(f"LONG TRAIN : {signal} -> {route}")

        # =================================================
        # ROUTE SET
        # =================================================

        if signal_type == "CALLING_ON":

            log(
                "LONG TRAIN : "
                "CALLING-ON INITIAL PREPARATION"
            )

            if not prepare_calling_on_route(row):
                append_long_train_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "CALLING-ON INITIAL PREPARATION FAILED"
                )

                continue

        else:

            click(
                signals[signal]["menu"]
            )

            time.sleep(0.2)

            if not click_menu_item(
                    route.replace("-", "_")
            ):
                append_long_train_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "ROUTE NOT SET"
                )

                continue

            time.sleep(2)

        # -------------------------------------------------
        # OPEN CONTROL BIT WINDOW
        # -------------------------------------------------

        if not open_control_bit_window():

            append_long_train_result(
                row,
                "FAIL",
                "UNKNOWN",
                "CONTROL BIT WINDOW FAILED"
            )

            continue

        sequence_ok = True

        # -------------------------------------------------
        # STEP 1 : ALL CONTROL BITS DOWN
        # -------------------------------------------------

        log("--------------------------------")
        log("CONTROL BITS DOWN")

        for control_bit in release_sequence:

            pause_event.wait()

            if not running:
                append_long_train_result(
                    row,
                    "FAIL",
                    "STOPPED",
                    "AUTOMATION STOPPED"
                )

                close_control_bit_window()
                return

            log(f"{control_bit} DOWN")

            if not transmit_control_bit(control_bit):

                sequence_ok = False
                break

            time.sleep(0.5)

        # -------------------------------------------------
        # STEP 2 : ALL CONTROL BITS NORMAL
        # EXCEPT LAST BIT
        # -------------------------------------------------

        if sequence_ok:

            log("--------------------------------")
            log("CONTROL BITS NORMAL")

            for control_bit in release_sequence[:-1]:

                pause_event.wait()

                if not running:
                    append_long_train_result(
                        row,
                        "FAIL",
                        "STOPPED",
                        "AUTOMATION STOPPED"
                    )

                    close_control_bit_window()
                    return

                log(f"{control_bit} NORMAL")

                if not transmit_control_bit(control_bit):

                    sequence_ok = False
                    break

                time.sleep(0.2)

        if not sequence_ok:

            close_control_bit_window()

            append_long_train_result(
                row,
                "FAIL",
                "UNKNOWN",
                "CONTROL BIT FAILED"
            )

            continue

        # -------------------------------------------------
        # VERIFY ROUTE
        # -------------------------------------------------

        log("Waiting 5 seconds for route release...")

        time.sleep(5)

        route_released = is_route_released(signal)

        if route_released:

            result = "PASS"
            final_aspect = "RELEASED"
            detail = "YELLOW ROUTE DISAPPEARED"

        else:

            result = "FAIL"
            final_aspect = "LOCKED"
            detail = "YELLOW ROUTE STILL PRESENT"

        log(f"{signal} {route} : {result}")

        append_long_train_result(
            row,
            result,
            final_aspect,
            detail
        )

        # -------------------------------------------------
        # LAST CONTROL BIT NORMAL
        # -------------------------------------------------

        last_bit = release_sequence[-1]

        log(f"{last_bit} NORMAL")

        transmit_control_bit(last_bit)

        time.sleep(1)

        # -------------------------------------------------
        # CLOSE CONTROL BIT WINDOW
        # -------------------------------------------------

        close_control_bit_window()

        time.sleep(2)

    log("LONG TRAIN AUTOMATION COMPLETED")

def run_short_train_engine():
    log("Entered run_short_train_engine()")

    global running

    root.iconify()
    time.sleep(0.5)

    for row in routes_data:
        signal_type = row.get(
            "signal_type",
            "MAIN"
        )

        # MAIN, SHUNT and CALLING-ON all run Short Train.
        # No signal type is skipped here.

        pause_event.wait()

        if not running or not automation_active:
            close_control_bit_window()
            return

        if not signals:
            log("FATAL : SIGNAL CONFIGURATION DISAPPEARED")
            return

        if not routes_data:
            log("FATAL : ROUTE DATA DISAPPEARED")
            return

        signal = row["signal"]
        route = row["route"]

        first_track = normalize_track_control(
            row["first_track"]
        )

        release_sequence = split_control_bits(
            row["release_sequence"]
        )

        back_lock_tracks = split_control_bits(
            row["back_lock_tracks"]
        )

        if signal not in signals:
            log(f"{signal} NOT FOUND")
            continue

        # =================================================
        # NOT APPLICABLE
        # No Sequential Route Release configuration
        # =================================================

        if (
                not first_track
                and not release_sequence
                and not back_lock_tracks
        ):
            log(
                f"{signal} {route} : "
                f"NOT APPLICABLE - NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            append_short_train_result(
                row,
                "NOT APPLICABLE",
                "N/A",
                "NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            continue

        # =================================================
        # INVALID / INCOMPLETE CONFIGURATION
        # =================================================

        if (
                not release_sequence
                or (
                signal_type != "CALLING_ON"
                and not first_track
        )
        ):
            log(
                f"{signal} {route} "
                f"HAS INVALID RELEASE SEQUENCE"
            )

            append_short_train_result(
                row,
                "FAIL",
                "NOT SET",
                "INVALID WORKSHEET ROW"
            )

            continue

        log("--------------------------------")
        log(f"SHORT TRAIN : {signal} -> {route}")

        # =================================================
        # CALLING-ON INITIAL PREPARATION
        # =================================================

        if signal_type == "CALLING_ON":

            log(
                "SHORT TRAIN : "
                "CALLING-ON INITIAL PREPARATION"
            )

            if not prepare_calling_on_route(row):
                append_short_train_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "CALLING-ON INITIAL PREPARATION FAILED"
                )

                continue

            log(
                "SHORT TRAIN : "
                "CALLING-ON PREPARATION COMPLETE"
            )

            # IMPORTANT:
            #
            # DO NOT DO:
            # FIRST TRACK DOWN
            # FIRST TRACK NORMAL
            # SIGNAL CANCEL
            # ROUTE RELEASE
            #
            # Calling-On goes DIRECTLY to
            # the Short Train release sequence.

        else:

            # =================================================
            # NORMAL MAIN / SHUNT ROUTE SET
            # =================================================

            click(
                signals[signal]["menu"]
            )

            time.sleep(1)

            if not click_menu_item(
                    route.replace("-", "_")
            ):
                append_short_train_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "ROUTE NOT SET"
                )

                continue

            # =================================================
            # FIRST TRACK DOWN
            # =================================================

            log(
                f"FIRST TRACK DOWN : {first_track}"
            )

            if not select_control_bit(
                    first_track
            ):
                append_short_train_result(
                    row,
                    "FAIL",
                    "UNKNOWN",
                    "FIRST TRACK DOWN FAILED"
                )

                continue

            # =================================================
            # FIRST TRACK NORMAL
            # =================================================

            log(
                f"FIRST TRACK NORMAL : {first_track}"
            )

            if not select_control_bit(
                    first_track
            ):
                append_short_train_result(
                    row,
                    "FAIL",
                    "UNKNOWN",
                    "FIRST TRACK NORMAL FAILED"
                )

                continue

            # =================================================
            # SIGNAL CANCEL / ROUTE RELEASE
            # =================================================

            attempt_manual_release(
                signal
            )

            log(
                "MANUAL RELEASE COMPLETED"
            )

            time.sleep(2)

        log(
            "STARTING RELEASE SEQUENCE"
        )

        if not open_control_bit_window():
            append_short_train_result(
                row,
                "FAIL",
                "UNKNOWN",
                "CONTROL BIT WINDOW FAILED"
            )
            continue

        sequence_ok = True

        previous_track = None

        # --------------------------------------------------
        # Train movement simulation
        # --------------------------------------------------

        for control_bit in release_sequence:

            pause_event.wait()

            if not running:
                append_short_train_result(
                    row,
                    "FAIL",
                    "STOPPED",
                    "AUTOMATION STOPPED"
                )

                close_control_bit_window()

                return

            # Occupy current track
            log(f"{control_bit} DOWN")

            if not transmit_control_bit(control_bit):
                sequence_ok = False
                break

            time.sleep(1)

            # Clear previous track
            if previous_track is not None:

                log(f"{previous_track} NORMAL")

                if not transmit_control_bit(previous_track):
                    sequence_ok = False
                    break

                time.sleep(1)

            previous_track = control_bit

        if sequence_ok and previous_track:

            log(f"{previous_track} NORMAL")

            if not transmit_control_bit(previous_track):
                sequence_ok = False

        close_control_bit_window()

        if not sequence_ok:
            close_control_bit_window()

            append_short_train_result(
                row,
                "FAIL",
                "UNKNOWN",
                "TRACK CONTROL BIT FAILED",
            )

            continue

        # ----------------------------
        # Verify Route Release
        # ----------------------------

        time.sleep(2)

        if not running:
            close_control_bit_window()

            append_short_train_result(
                row,
                "FAIL",
                "STOPPED",
                "AUTOMATION STOPPED"
            )

            return

        log("Waiting 5 seconds for route release...")

        time.sleep(5)

        route_released = is_route_released(signal)

        if route_released:

            result = "PASS"
            final_aspect = "RELEASED"
            detail = "YELLOW ROUTE DISAPPEARED"

        else:

            result = "FAIL"
            final_aspect = "LOCKED"
            detail = "YELLOW ROUTE STILL PRESENT"

        log(f"{signal} {route} : {result}")

        # User may have pressed STOP while verification finished

        append_short_train_result(
            row,
            result,
            final_aspect,
            detail
        )

        if not running:
            return

    log("SHORT TRAIN AUTOMATION COMPLETED")

# =========================================================
# SHUNT EMERGENCY ROUTE RELEASE TEST
# =========================================================

def run_shunt_err_case(row):

    global running

    signal = row["signal"]
    route = row["route"]

    back_lock_tracks = split_control_bits(
        row["back_lock_tracks"]
    )

    # =====================================================
    # START
    # =====================================================

    log("--------------------------------")
    log(
        f"SHUNT ERR : {signal} -> {route}"
    )

    # =====================================================
    # VALIDATION
    # =====================================================

    if signal not in signals:

        log(
            f"SHUNT ERR : {signal} NOT FOUND"
        )

        return

    if not route:

        log(
            f"SHUNT ERR : {signal} ROUTE IS EMPTY"
        )

        append_err_result(
            row,
            "FAIL",
            "NOT SET",
            "SHUNT ROUTE MISSING"
        )

        return

    if not back_lock_tracks:

        log(
            f"SHUNT ERR : {signal} "
            f"HAS NO BACK LOCK TRACKS"
        )

        append_err_result(
            row,
            "FAIL",
            "NOT SET",
            "BACK LOCK TRACKS MISSING"
        )

        return

    # =====================================================
    # STEP 1 : SET ROUTE
    # =====================================================

    log(
        f"SHUNT : SETTING ROUTE : {route}"
    )

    click(
        signals[signal]["menu"]
    )

    time.sleep(0.5)

    if not click_menu_item(
        route.replace("-", "_")
    ):

        log(
            f"SHUNT : ROUTE NOT SET : {route}"
        )

        append_err_result(
            row,
            "FAIL",
            "NOT SET",
            "SHUNT ROUTE NOT SET"
        )

        return

    log(
        f"SHUNT : {route} SET"
    )

    time.sleep(2)

    # =====================================================
    # STEP 2 : OPEN CONTROL BIT WINDOW
    # =====================================================

    if not err_open_control_bit_window():

        log(
            "SHUNT : CONTROL BIT WINDOW FAILED"
        )

        append_err_result(
            row,
            "FAIL",
            "UNKNOWN",
            "CONTROL BIT WINDOW FAILED"
        )

        return

    sequence_ok = True

    # =====================================================
    # STEP 3 :
    #
    # FOR EVERY BACK LOCK TRACK:
    #
    # 1. TRACK DOWN
    # 2. SIGNAL CANCEL
    # 3. ROUTE RELEASE
    # 4. CHECK ROUTE STATUS - TASK ONLY
    # 5. SAME TRACK NORMAL
    # 6. NEXT BACK LOCK TRACK
    # =====================================================

    for index, control_bit in enumerate(
        back_lock_tracks,
        start=1
    ):

        pause_event.wait()

        if not running or not automation_active:

            log(
                "SHUNT : AUTOMATION STOPPED"
            )

            err_close_control_bit_window()

            append_err_result(
                row,
                "FAIL",
                "STOPPED",
                "AUTOMATION STOPPED"
            )

            return

        log("--------------------------------")

        log(
            f"SHUNT BACK LOCK "
            f"{index}/{len(back_lock_tracks)} : "
            f"{control_bit}"
        )

        # =================================================
        # 3.1 BACK LOCK TRACK DOWN
        # =================================================

        log(
            f"SHUNT : {control_bit} DOWN"
        )

        if not err_transmit_control_bit(
            control_bit
        ):

            log(
                f"SHUNT : {control_bit} DOWN FAILED"
            )

            sequence_ok = False
            break

        time.sleep(1)

        log(
            f"SHUNT : {control_bit} DOWN COMPLETED"
        )

        # =================================================
        # 3.2 SIGNAL CANCEL
        # =================================================

        log(
            f"SHUNT : SIGNAL CANCEL : "
            f"{signal}"
        )

        click(
            signals[signal]["menu"]
        )

        time.sleep(0.5)

        if not click_menu_item(
            "Signal Cancel"
        ):

            log(
                "SHUNT : SIGNAL CANCEL NOT FOUND"
            )

            sequence_ok = False
            break

        log(
            "SHUNT : SIGNAL CANCEL DONE"
        )

        time.sleep(0.5)

        # =================================================
        # 3.3 ROUTE RELEASE
        # =================================================

        log(
            f"SHUNT : ROUTE RELEASE : "
            f"{signal}"
        )

        click(
            signals[signal]["menu"]
        )

        time.sleep(0.5)

        if not click_menu_item(
            "Route Release"
        ):

            log(
                "SHUNT : ROUTE RELEASE NOT FOUND"
            )

            sequence_ok = False
            break

        log(
            "SHUNT : ROUTE RELEASE DONE"
        )

        time.sleep(1)

        # =================================================
        # 3.4 CHECK ROUTE STATUS
        #
        # THIS IS ONLY A TASK STEP.
        #
        # DO NOT STORE THIS RESULT IN REPORT.
        # =================================================

        log(
            f"SHUNT : CHECKING ROUTE STATUS "
            f"AFTER {control_bit}"
        )

        route_released = is_route_released(
            signal
        )

        if route_released:

            log(
                f"SHUNT : ROUTE RELEASED "
                f"AFTER {control_bit}"
            )

        else:

            log(
                f"SHUNT : ROUTE STILL LOCKED "
                f"AFTER {control_bit}"
            )

        # =================================================
        # 3.5 SAME BACK LOCK TRACK NORMAL
        #
        # CLICK THE SAME CONTROL BIT AGAIN
        # TO MAKE IT NORMAL.
        # =================================================

        log(
            f"SHUNT : {control_bit} NORMAL"
        )

        if not err_transmit_control_bit(
            control_bit
        ):

            log(
                f"SHUNT : {control_bit} NORMAL FAILED"
            )

            sequence_ok = False
            break

        time.sleep(1)

        log(
            f"SHUNT : {control_bit} NORMAL COMPLETED"
        )

        # =================================================
        # NEXT BACK LOCK TRACK
        # =================================================

        log(
            f"SHUNT : BACK LOCK {control_bit} "
            f"RETURNED TO NORMAL"
        )

    # =====================================================
    # BACK LOCK SEQUENCE FAILED
    # =====================================================

    if not sequence_ok:

        err_close_control_bit_window()

        append_err_result(
            row,
            "FAIL",
            "UNKNOWN",
            "SHUNT BACK LOCK SEQUENCE FAILED"
        )

        return

    # =====================================================
    # STEP 4 : FINAL ROUTE RELEASE CHECK
    #
    # THIS IS THE ONLY CHECK USED FOR REPORT.
    # =====================================================

    log("--------------------------------")

    log(
        "SHUNT : ALL BACK LOCK TRACKS COMPLETED"
    )

    log(
        "SHUNT : WAITING BEFORE FINAL "
        "ROUTE RELEASE CHECK..."
    )

    time.sleep(5)

    log(
        "SHUNT : CHECKING FINAL ROUTE STATUS..."
    )

    route_released = is_route_released(
        signal
    )

    # =====================================================
    # FINAL PASS / FAIL
    #
    # NOT RELEASED -> PASS
    # RELEASED     -> FAIL
    # =====================================================

    if route_released:

        result = "FAIL"

        final_aspect = "RELEASED"

        detail = (
            "SHUNT ROUTE RELEASED "
            "AFTER ALL BACK LOCK TESTS"
        )

        log(
            "SHUNT : FINAL ROUTE RELEASED -> FAIL"
        )

    else:

        result = "PASS"

        final_aspect = "LOCKED"

        detail = (
            "SHUNT ROUTE STILL LOCKED "
            "AFTER ALL BACK LOCK TESTS"
        )

        log(
            "SHUNT : FINAL ROUTE NOT RELEASED -> PASS"
        )

    # =====================================================
    # STEP 5 : STORE FINAL RESULT ONLY
    # =====================================================

    append_err_result(
        row,
        result,
        final_aspect,
        detail
    )

    log(
        f"SHUNT ERR RESULT : "
        f"{signal} -> {result}"
    )

    # =====================================================
    # STEP 6 : SHUNT EMERGENCY ROUTE RELEASE
    #
    # CONTROL BIT WINDOW IS STILL OPEN.
    # =====================================================

    log("--------------------------------")

    log(
        "SHUNT : STARTING EMERGENCY "
        "ROUTE RELEASE SEQUENCE"
    )

    emergency_ok = shunt_emergency_key_sequence(
        signal,
        bit_chart_open=True
    )

    # =====================================================
    # STEP 7 : CLOSE CONTROL BIT WINDOW
    # =====================================================

    err_close_control_bit_window()

    if not emergency_ok:

        log(
            "SHUNT : EMERGENCY ROUTE RELEASE FAILED"
        )

    else:

        log(
            "SHUNT : EMERGENCY ROUTE RELEASE DONE"
        )

    log("--------------------------------")

    log(
        f"SHUNT ERR COMPLETED : "
        f"{signal} -> {result}"
    )

def run_err_engine():

    global running

    root.iconify()
    time.sleep(0.5)

    for row in routes_data:

        pause_event.wait()

        if not running or not automation_active:
            close_control_bit_window()
            return

        if not signals:
            log("FATAL : SIGNAL CONFIGURATION DISAPPEARED")
            return

        if not routes_data:
            log("FATAL : ROUTE DATA DISAPPEARED")
            return

        signal = row["signal"]
        route = row["route"]

        signal_type = row.get(
            "signal_type",
            "MAIN"
        )

        first_track = normalize_track_control(
            row["first_track"]
        )

        release_sequence = split_control_bits(
            row["release_sequence"]
        )

        back_lock_tracks = split_control_bits(
            row["back_lock_tracks"]
        )

        if signal not in signals:
            log(f"{signal} NOT FOUND")
            continue

        # =================================================
        # NOT APPLICABLE
        # No Sequential Route Release configuration
        # =================================================

        if (
                not first_track
                and not release_sequence
                and not back_lock_tracks
        ):
            log(
                f"{signal} {route} : "
                f"NOT APPLICABLE - NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            append_err_result(
                row,
                "NOT APPLICABLE",
                "N/A",
                "NO SEQUENTIAL ROUTE CONFIGURATION"
            )

            continue

        # =================================================
        # SHUNT ERR
        # Keep existing special SHUNT ERR method
        # =================================================

        if (
                signal in signals
                and signals[signal].get("type") == "SHUNT"
        ):
            run_shunt_err_case(row)

            continue

        # =================================================
        # MAIN ERR VALIDATION
        # =================================================

        if not first_track or not back_lock_tracks:
            append_err_result(
                row,
                "FAIL",
                "NOT SET",
                "INVALID WORKSHEET ROW"
            )

            continue

        log("--------------------------------")
        log(f"ERR : {signal} -> {route}")

        # =================================================
        # CALLING-ON INITIAL PREPARATION / NORMAL ROUTE SET
        # =================================================

        if signal_type == "CALLING_ON":

            # -------------------------------------------------
            # CALLING-ON
            #
            # prepare_calling_on_route() does:
            #
            # OPEN CONTROL BIT
            # CALLING-ON TRACK DOWN
            # CALLING-ON ROUTE SET
            # CALLING-ON TRACK NORMAL
            # CLOSE CONTROL BIT
            #
            # Therefore DO NOT SET THE ROUTE AGAIN here.
            # -------------------------------------------------

            log(
                "ERR : CALLING-ON INITIAL PREPARATION"
            )

            if not prepare_calling_on_route(row):
                append_err_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "CALLING-ON INITIAL PREPARATION FAILED"
                )

                continue

            log(
                "ERR : CALLING-ON PREPARATION COMPLETE"
            )

        else:

            # -------------------------------------------------
            # MAIN
            #
            # Main signal sets route normally.
            # -------------------------------------------------

            click(
                signals[signal]["menu"]
            )

            time.sleep(0.2)

            if not click_menu_item(
                    route.replace("-", "_")
            ):
                append_err_result(
                    row,
                    "FAIL",
                    "NOT SET",
                    "ROUTE NOT SET"
                )

                continue

            time.sleep(0.2)

        # ------------------------------------
        # OPEN CONTROL BIT WINDOW (ONLY ONCE)
        # ------------------------------------

        if not err_open_control_bit_window():

            append_err_result(
                row,
                "FAIL",
                "UNKNOWN",
                "CONTROL BIT WINDOW FAILED"
            )

            emergency_route_release(signal)
            continue

        # ------------------------------------
        # RELEASE SEQUENCE
        # ------------------------------------

        sequence_ok = True

        for control_bit in back_lock_tracks:

            pause_event.wait()

            if not running:
                append_err_result(
                    row,
                    "FAIL",
                    "STOPPED",
                    "AUTOMATION STOPPED"
                )

                err_close_control_bit_window()

                return

            log("--------------------------------")
            log(f"TRACK : {control_bit}")

            # Signal Cancel + Route Release
            attempt_manual_release(signal)

            # DOWN
            log(f"{control_bit} DOWN")

            if not err_transmit_control_bit(control_bit):

                append_err_result(
                    row,
                    "FAIL",
                    "UNKNOWN",
                    f"{control_bit} DOWN FAILED"
                )

                sequence_ok = False
                break

            # NORMAL
            log(f"{control_bit} NORMAL")

            if not err_transmit_control_bit(control_bit):

                append_err_result(
                    row,
                    "FAIL",
                    "UNKNOWN",
                    f"{control_bit} NORMAL FAILED"
                )

                sequence_ok = False
                break

            time.sleep(0.2)

        if not sequence_ok:

            err_close_control_bit_window()

            emergency_route_release(signal, bit_chart_open=True)
            continue

        # ------------------------------------
        # FIRST TRACK
        # ------------------------------------

        attempt_manual_release(signal)

        log(f"{first_track} DOWN")

        if not err_transmit_control_bit(first_track):

            append_err_result(
                row,
                "FAIL",
                "UNKNOWN",
                "FIRST TRACK DOWN FAILED"
            )

            err_close_control_bit_window()

            emergency_route_release(signal, bit_chart_open=True)

            continue

        log(f"{first_track} NORMAL")

        if not err_transmit_control_bit(first_track):

            append_err_result(
                row,
                "FAIL",
                "UNKNOWN",
                "FIRST TRACK NORMAL FAILED"
            )

            err_close_control_bit_window()

            emergency_route_release(signal, bit_chart_open=True)

            continue

        # ------------------------------------
        # VERIFY ROUTE
        # ------------------------------------

        log("Waiting 5 seconds for route release...")

        time.sleep(5)

        route_released = is_route_released(signal)

        if route_released:

            result = "FAIL"

            final_aspect = "RELEASED"

            detail = "ROUTE RELEASED"

        else:

            result = "PASS"

            final_aspect = "LOCKED"

            detail = "ROUTE STILL LOCKED"

        log(f"{signal} {route} : {result}")

        append_err_result(
            row,
            result,
            final_aspect,
            detail
        )

        # ------------------------------------
        # CLEANUP
        # ------------------------------------

        try:
            emergency_route_release(signal, bit_chart_open=True)
        finally:
            err_close_control_bit_window()

    log("ERR AUTOMATION COMPLETED")

def initialize_report():

    if os.path.exists(REPORT_FILE):
        try:
            os.remove(REPORT_FILE)
        except PermissionError:
            log("Please close REPORT FILE.")
            return False

    global report_rows

    backup = report_rows.copy()

    report_rows.clear()

    create_report(REPORT_FILE, "SHORT TRAIN")
    create_report(REPORT_FILE, "LONG TRAIN")
    create_report(REPORT_FILE, "EMERGENCY ROUTE RELEASE")

    report_rows = backup

    log("Blank Report Workbook Created")

    return True

def validate_runtime_state():

    log("--------------------------------")
    log("VALIDATING AUTOMATION RUNTIME STATE")

    if not signals:
        log("FATAL : SIGNAL CONFIGURATION IS EMPTY")
        return False

    if not routes_data:
        log("FATAL : ROUTE DATA IS EMPTY")
        return False

    if not config_file_path:
        log("FATAL : CONFIG FILE PATH IS EMPTY")
        return False

    config_name = os.path.basename(
        config_file_path
    ).strip().upper()

    # ---------------------------------------------------------
    # CONFIG FILE VALIDATION
    # Accept any Excel configuration file loaded by the user.
    # ---------------------------------------------------------

    if not os.path.exists(config_file_path):
        log(
            f"FATAL : CONFIG FILE NOT FOUND : "
            f"{config_file_path}"
        )
        return False

    if not config_name.endswith(".XLSX"):
        log(
            f"FATAL : CONFIG FILE IS NOT XLSX : "
            f"{config_file_path}"
        )
        return False

    log(f"CONFIG VERIFIED : {config_file_path}")
    log(f"SIGNALS VERIFIED : {len(signals)}")
    log(f"ROUTES VERIFIED : {len(routes_data)}")
    log("--------------------------------")

    return True
def enable_gui_controls():

    try:
        load_toc_btn.config(state="normal")
        existing_signal_btn.config(state="normal")
        new_signal_btn.config(state="normal")
        save_coordinates_btn.config(state="normal")
        start_testing_btn.config(state="normal")
        stop_testing_btn.config(state="normal")
    except Exception as e:
        log(f"GUI CONTROL ENABLE ERROR : {e}")

# =========================================================
# START
# =========================================================

def start_automation():

    global running
    global report_saved
    global automation_active

    # =====================================================
    # ABSOLUTE FIRST CHECK
    # =====================================================

    if running or automation_active:
        log("START AUTOMATION IGNORED : ALREADY RUNNING")
        return

    log("START_AUTOMATION() ENTERED")
    log(f"START CHECK - signals = {len(signals)}")
    log(f"START CHECK - routes_data = {len(routes_data)}")
    log(f"START CHECK - config_file_path = {config_file_path}")

    # =====================================================
    # PREVENT DOUBLE START
    # =====================================================

    if running:
        log("START AUTOMATION IGNORED : ALREADY RUNNING")
        return

    # =====================================================
    # SIGNAL CONFIG CHECK
    # =====================================================

    if not signals:
        log("FATAL : SIGNAL CONFIG IS EMPTY")
        log(f"CONFIG FILE : {config_file_path}")

        return

    if not routes_data:
        log("FATAL : TOC ROUTES ARE EMPTY")

        return

    # =========================================================
    # PRE-FLIGHT CHECK
    # If none of the TOC's Signal values match a loaded coordinate,
    # every route will fail immediately and the run will look like
    # it finished instantly with nothing actually tested. Catch that
    # here, loudly, BEFORE initialize_report() runs.
    # =========================================================

    toc_signals = {
        str(row.get("signal", "")).strip().upper()
        for row in routes_data
    }
    matched = toc_signals & set(signals.keys())

    if not matched:

        log("=" * 60)
        log("CANNOT START - NO MATCHING SIGNALS")
        log(f"TOC SIGNALS : {sorted(toc_signals)}")
        log(f"LOADED COORDINATE SIGNALS : {sorted(signals.keys())}")
        log(
            "None of the TOC signals match the loaded coordinates. "
            "Check the coordinate file is for the right yard, then "
            "reload it."
        )
        log("=" * 60)

        messagebox.showerror(
            "CANNOT START",
            "None of the TOC signals match the loaded coordinates.\n\n"
            "Check the log for the exact signal lists, then reload "
            "the correct coordinate file before starting."
        )

        return

    if not validate_runtime_state():
        log("AUTOMATION NOT STARTED : INVALID RUNTIME STATE")
        return
    log("--------------------------------")
    log("START AUTOMATION CHECK")
    log(f"SIGNALS LOADED : {len(signals)}")
    log(f"ROUTES LOADED  : {len(routes_data)}")
    log(f"CONFIG FILE    : {config_file_path}")
    log("--------------------------------")

    report_saved = False
    report_rows.clear()

    if not initialize_report():
        return

    running = True
    automation_active = True

    pause_event.set()
    # =====================================================
    # MAKE TEST PANEL ACTIVE
    # =====================================================

    log("================================")
    log("SWITCHING TO TEST PANEL")
    log("================================")

    root.iconify()

    time.sleep(2)

    if not bring_test_panel_to_front():
        log("ERROR: TEST PANEL NOT FOUND")

        running = False
        automation_active = False

        root.after(
            0,
            root.deiconify
        )

        root.after(
            0,
            lambda: status_label.config(
                text="TEST PANEL NOT FOUND",
                fg="#dc2626"
            )
        )

        return

    log("TEST PANEL FOUND")
    log("TEST PANEL BROUGHT TO FRONT")
    log("TEST PANEL READY")

    time.sleep(1)

    # Disable all configuration/start controls during automation.
    root.after(
        0,
        lambda: load_toc_btn.config(state="disabled")
    )

    root.after(
        0,
        lambda: existing_signal_btn.config(state="disabled")
    )

    root.after(
        0,
        lambda: new_signal_btn.config(state="disabled")
    )

    root.after(
        0,
        lambda: save_coordinates_btn.config(state="disabled")
    )

    root.after(
        0,
        lambda: start_testing_btn.config(state="disabled")
    )
    status_label.config(
        text="RUNNING",
        fg="#16a34a"
    )

    def worker():

        global running
        global report_saved
        global current_test

        try:

            # =================================================
            # SHORT TRAIN
            # =================================================

            current_test = "SHORT"

            report_rows.clear()
            report_saved = False

            log("================================")
            log("STARTING SHORT TRAIN")
            log("================================")

            run_short_train_engine()

            log("SHORT TRAIN ENGINE RETURNED")

            if report_rows:

                save_short_train_report()

            else:

                log("SHORT TRAIN PRODUCED NO REPORT ROWS")

            if not running:
                log("AUTOMATION STOPPED DURING SHORT TRAIN")
                automation_active = False
                root.after(
                    0,
                    root.deiconify
                )

                return

            # =================================================
            # LONG TRAIN
            # =================================================

            current_test = "LONG"

            report_rows.clear()
            report_saved = False

            log("================================")
            log("STARTING LONG TRAIN")
            log("================================")

            run_long_train_engine()

            log("LONG TRAIN ENGINE RETURNED")

            if report_rows:

                save_long_train_report()

            else:

                log("LONG TRAIN PRODUCED NO REPORT ROWS")

            if not running:
                log("AUTOMATION STOPPED DURING LONG TRAIN")

                root.after(
                    0,
                    root.deiconify
                )

                return

            # =================================================
            # ERR
            # =================================================

            current_test = "ERR"

            report_rows.clear()
            report_saved = False

            log("================================")
            log("STARTING ERR TEST")
            log("================================")

            run_err_engine()

            log("ERR ENGINE RETURNED")

            if report_rows:

                save_err_report()

            else:

                log("ERR TEST PRODUCED NO REPORT ROWS")

            if not running:
                log("AUTOMATION STOPPED DURING ERR")

                root.after(
                    0,
                    root.deiconify
                )

                return

            # =================================================
            # ALL COMPLETED
            # =================================================

            running = False
            automation_active = False

            root.after(
                0,
                enable_gui_controls
            )

            root.after(
                0,
                root.deiconify
            )

            root.after(
                0,
                lambda: status_label.config(
                    text="COMPLETED",
                    fg="#16a34a"
                )
            )

            root.after(
                0,
                root.deiconify
            )

            log("================================")
            log("ALL TESTS COMPLETED")
            log("================================")

            root.after(
                0,
                lambda: messagebox.showinfo(
                    "Completed",
                    "All tests completed.\n\n"
                    "Report generated successfully."
                )
            )

        except Exception as e:

            import traceback

            error_text = traceback.format_exc()

            print(error_text)

            # =================================================
            # COMPLETE ERROR IN LIVE LOG
            # =================================================

            log("================================")
            log("AUTOMATION CRASHED")
            log(f"CURRENT TEST : {current_test}")
            log(f"ERROR : {e}")

            for line in error_text.splitlines():
                log(line)

            running = False
            automation_active = False

            root.after(
                0,
                enable_gui_controls
            )

            root.after(
                0,
                lambda: status_label.config(
                    text="ERROR",
                    fg="#dc2626"
                )
            )

            root.after(
                0,
                root.deiconify
            )

            root.after(
                0,
                root.deiconify
            )
    threading.Thread(
        target=worker,
        daemon=True
    ).start()

# =========================================================
# STOP
# =========================================================

def stop_automation():

    global running
    global paused
    global automation_active
    global report_saved

    # =====================================================
    # STOP AUTOMATION
    # =====================================================

    running = False
    automation_active = False
    paused = False

    pause_event.set()

    status_label.config(
        text="STOPPED",
        fg="#dc2626"
    )

    root.deiconify()
    root.lift()

    enable_gui_controls()

    log("STOPPED")

    # =====================================================
    # SAVE CURRENT REPORT
    # =====================================================

    try:

        if report_rows:

            log("--------------------------------")
            log(f"STOP : SAVING {current_test} REPORT")
            log(f"STOP : REPORT ROWS = {len(report_rows)}")

            if current_test == "SHORT":

                create_report(
                    REPORT_FILE,
                    "SHORT TRAIN"
                )

                log(
                    f"SHORT TRAIN REPORT SAVED : {REPORT_FILE}"
                )

            elif current_test == "LONG":

                create_report(
                    REPORT_FILE,
                    "LONG TRAIN"
                )

                log(
                    f"LONG TRAIN REPORT SAVED : {REPORT_FILE}"
                )

            elif current_test == "ERR":

                create_report(
                    REPORT_FILE,
                    "EMERGENCY ROUTE RELEASE"
                )

                log(
                    f"ERR REPORT SAVED : {REPORT_FILE}"
                )

            report_saved = True

            log("STOP : REPORT GENERATED")
            log("--------------------------------")

        else:

            log("STOP : NO REPORT ROWS AVAILABLE")

    except Exception as e:

        log(
            f"STOP : REPORT SAVE FAILED : {e}"
        )

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
    "SEQUENTIAL ROUTE RELEASE AUTOMATION"
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
    "TNotebook.Tab",
    font=("Segoe UI", 10, "bold"),
    padding=(20, 8)   # horizontal, vertical
)

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

    img = Image.open(r"Indian_Railway_Logo_2.png")

    # 👇 THIS IS THE MAIN FIX (resize properly)
    img = img.resize((80, 80))   # change size if needed

    railway_logo = ImageTk.PhotoImage(img)

    logo_label = tk.Label(
        left_logo_frame,
        image=railway_logo,
        bg="#0f172a"
    )

    logo_label.image = railway_logo   # IMPORTANT
    logo_label.pack(pady=5)

except Exception as e:

    print("Logo error:", e)

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

    text="SEQUENTIAL ROUTE RELEASE AUTOMATION",

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

load_toc_btn = create_button(
    "LOAD TOC",
    load_routes,
    "#2563eb"
)

load_toc_btn.pack(pady=8)

save_coordinates_btn = create_button(
    "SAVE NEW SIGNAL COORDINATES",
    save_new_signal_coordinates,
    "#16a34a"
)

save_coordinates_btn.pack(pady=8)

load_universal_btn = create_button(
    "LOAD YARD COORDINATES",
    load_universal_coordinates,
    "#7c3aed"
)

load_universal_btn.pack(pady=8)

start_testing_btn = create_button(
    "START TESTING",
    start_automation,
    "#16a34a"
)

start_testing_btn.pack(pady=20)

stop_testing_btn = create_button(
    "STOP TESTING",
    stop_automation,
    "#dc2626"
)

stop_testing_btn.pack(pady=8)

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


notebook = ttk.Notebook(right_panel)
notebook.pack(fill="both", expand=True)

main_tab = tk.Frame(notebook)
calling_tab = tk.Frame(notebook)
shunt_tab = tk.Frame(notebook)

notebook.add(main_tab, text="MAIN SIGNALS")
notebook.add(calling_tab, text="CALLING-ON SIGNALS")
notebook.add(shunt_tab, text="SHUNT SIGNALS")

notebook.bind(
    "<<NotebookTabChanged>>",
    lambda event: refresh_table()
)

# =========================================================
# TABLE FRAME
# =========================================================

table_frame = tk.Frame(
    main_tab,
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

        "FIRST_TRACK",

        "RELEASE_SEQUENCE",

        "BACK_LOCK_TRACKS"

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

tree.heading("FIRST_TRACK", text="FIRST TRACK")

tree.heading("RELEASE_SEQUENCE", text="RELEASE SEQUENCE")

tree.heading("BACK_LOCK_TRACKS", text="BACK LOCK TRACKS")

tree.column("NO", width=70, anchor="center")

tree.column("SIGNAL", width=100, anchor="center")

tree.column("ROUTE", width=100, anchor="center")

tree.column("FIRST_TRACK", width=140, anchor="center")

tree.column("RELEASE_SEQUENCE", width=360, anchor="center")

tree.column("BACK_LOCK_TRACKS", width=250, anchor="center")

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

    text="SPACE KEY = CAPTURE COORDINATES | SEQUENTIAL ROUTE RELEASE AUTOMATION",

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

listener = pynput_keyboard.Listener(on_press=on_press)
listener.daemon = True
listener.start()

# =========================================================
# RUN
# =========================================================
root.after(
    300,
    startup_walkaway
)

root.mainloop()
