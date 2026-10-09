# =========================================================
# UNIVERSAL LOCK ROUTE AUTOMATION SYSTEM
# =========================================================
# FEATURES
# =========================================================
# 1. MAIN SIGNAL SUPPORT
# 2. CALLING ON SUPPORT
# 3. SHUNT SUPPORT
# 4. SIGNAL CONFIG SAVE/LOAD
# 5. LOCK ROUTE AUTOMATION
# 6. PASS / FAIL REPORT
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

import tkinter as tk
import keyboard
from pynput import keyboard as pynput_keyboard
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import ttk
from tkinter import messagebox
from PIL import Image as PILImage
from PIL import ImageTk
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

import pyautogui
import cv2
import numpy as np

import win32api
import win32con
from collections import Counter

import uiautomation as auto

from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

# Base name for generated reports.
# A timestamped filename is created for every completed run.
REPORT_FILE = ""

# =========================================================
# GLOBALS
# =========================================================

signals = {}

lock_routes_data = []

track_points = {}

running = False

capture_queue = []

capture_index = 0

capture_master_list = []

record_stage = None
capture_module = None

current_aspects = 2

undo_stack = []

capture_mode = None

current_signal = None

current_step = 0

last_capture_time = 0

capture_waiting = False

captured_point = None

CONFIG_FILE = ""

config_file_path = ""

# =========================================================
# SHUNT SNAPSHOT / DEBUG FOLDERS
# =========================================================

SNAPSHOTS_FOLDER = "snapshots"
DEBUG_FOLDER = "debug_captures"
os.makedirs(SNAPSHOTS_FOLDER, exist_ok=True)
os.makedirs(DEBUG_FOLDER, exist_ok=True)
orb = cv2.ORB_create(nfeatures=500)

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
    "CH": "CRANK HANDLE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CALLING_ON": "#f97316",
    "SHUNT": "#a855f7",
    "POINT": "#14b8a6",
    "CH": "#eab308",
    "LC": "#f43f5e"
}


# =========================================================
# REPORT DATA STORAGE
# =========================================================

report_rows = []

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
        text=(
            f"Signal {capture_index + 1} of "
            f"{len(capture_master_list)}   |   "
            f"{TYPE_LABELS.get(signal_type, signal_type)}"
        )
    )

    overlay_signal_label.config(
        text=signal,
        fg=TYPE_COLORS.get(signal_type, "white")
    )

    overlay_step_label.config(
        text=step_text
    )

    overlay_hint_label.config(
        text=hint_text
    )

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
    Build the complete NEW-SIGNAL capture list.

    Capture order:
        MAIN -> CALLING-ON -> SHUNT -> POINT -> CH -> LC

    This order is intentionally different from the TL's older quantity
    helper, because the AHEAD capture workflow must finish the Calling-On
    section before moving to SHUNT.
    """

    global signals

    signals = {}
    combined = []
    existing = set()

    # =====================================================
    # 1. MAIN SIGNALS
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
            "aspects": 2,
            "menu": None,
            "RED": None,
            "YELLOW": None,
            "DOUBLE_YELLOW": None,
            "GREEN": None,
            "ROUTE_INDICATOR": None
        }

        combined.append(("MAIN", name))

    # =====================================================
    # 2. CALLING-ON SIGNALS
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
    # 3. SHUNT SIGNALS
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
            "initial_snapshot": None,
            "route_init": None
        }

        combined.append(("SHUNT", name))

    # =====================================================
    # 4. POINTS
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
            # Backward-compatible alias used by older automation.
            "coordinate": None
        }

        combined.append(("POINT", name))

    # =====================================================
    # 5. CRANK HANDLES
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
            # Backward-compatible alias used by older automation.
            "coordinate": None
        }

        combined.append(("CH", name))

    # =====================================================
    # 6. LC GATES
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
            "out": None,
            # Backward-compatible alias used by older automation.
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

    # Enable the global keyboard dispatcher for MASTER coordinate capture.
    capture_module = "MASTER"

    capture_master_list = build_quantity_capture_list()

    if not capture_master_list:

        capture_module = None
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
    global capture_module

    capture_index += 1

    if capture_index < len(capture_master_list):

        set_initial_stage_for_current()

        master_next_capture()

    else:

        capture_module = None

        master_next_capture()

def master_next_capture():

    global capture_module

    if capture_index >= len(capture_master_list):
        log("ALL SIGNAL COORDINATES CAPTURED")

        destroy_capture_overlay()

        root.deiconify()

        capture_module = None

        messagebox.showinfo(
            "Capture Complete",
            "All signal coordinates have been captured.\n\nNow click 'SAVE NEW SIGNAL COORDINATES'."
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
                "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "indicator":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK: ASPECT INDICATOR",
                "Move the mouse onto the aspect indicator, then press SPACE.\n"
                "This also saves the initial reference snapshot."
            )

        elif record_stage == "route_init":

            update_capture_overlay(
                "SHUNT",
                signal,
                "CLICK: ROUTE INITIATION INDICATOR",
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

        elif record_stage == "route_init":

            update_capture_overlay(
                "CALLING_ON",
                signal,
                "CLICK : ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # POINT MACHINE
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
                "Move the mouse onto the FREE / OUT OF CORRESPONDENCE indicator, then press SPACE."
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

            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved ({x},{y})")
            record_stage = "indicator"
            master_next_capture()
            return

        if record_stage == "indicator":

            screenshot = pyautogui.screenshot()
            region = screenshot.crop((x - 50, y - 50, x + 50, y + 50))
            snapshot_file = os.path.abspath(
                os.path.join(SNAPSHOTS_FOLDER, f"{signal}_initial.png")
            )
            region.save(snapshot_file)

            def _undo_shunt_indicator(s=signal, f=snapshot_file):
                signals[s]["state_indicator"] = None
                signals[s]["initial_snapshot"] = None
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass

            push_undo(_undo_shunt_indicator)
            signals[signal]["state_indicator"] = [x, y]
            signals[signal]["initial_snapshot"] = snapshot_file
            log(f"{signal} Indicator Saved ({x},{y})")
            log(f"{signal} Initial Snapshot: {snapshot_file}")
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":

            push_undo(lambda s=signal: signals[s].__setitem__("route_init", None))
            signals[signal]["route_init"] = [x, y]
            log(f"{signal} Route Initiation Indicator Saved ({x},{y})")
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

            log(f"{signal} Route Initiation Indicator Saved ({x},{y})")

            advance_to_next_signal()

            return

    # =====================================================
    # POINT MACHINE
    # =====================================================

    elif sig_type == "POINT":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            # Keep the old generic coordinate alias for compatibility.
            signals[signal]["coordinate"] = [x, y]
            log(f"{signal} POINT MENU Saved ({x},{y})")
            record_stage = "normal"
            master_next_capture()
            return

        if record_stage == "normal":
            push_undo(lambda s=signal: signals[s].__setitem__("normal", None))
            signals[signal]["normal"] = [x, y]
            log(f"{signal} NORMAL Saved ({x},{y})")
            record_stage = "reverse"
            master_next_capture()
            return

        if record_stage == "reverse":
            push_undo(lambda s=signal: signals[s].__setitem__("reverse", None))
            signals[signal]["reverse"] = [x, y]
            log(f"{signal} REVERSE Saved ({x},{y})")
            record_stage = "free"
            master_next_capture()
            return

        if record_stage == "free":
            push_undo(lambda s=signal: signals[s].__setitem__("free", None))
            signals[signal]["free"] = [x, y]
            log(f"{signal} FREE Saved ({x},{y})")
            advance_to_next_signal()
            return

    # =====================================================
    # CRANK HANDLE
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
    # LC GATE
    # =====================================================

    elif sig_type == "LC":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            signals[signal]["coordinate"] = [x, y]
            log(f"{signal} LC MENU Saved ({x},{y})")
            record_stage = "in"
            master_next_capture()
            return

        if record_stage == "in":
            push_undo(lambda s=signal: signals[s].__setitem__("in", None))
            signals[signal]["in"] = [x, y]
            log(f"{signal} LC IN Saved ({x},{y})")
            record_stage = "out"
            master_next_capture()
            return

        if record_stage == "out":
            push_undo(lambda s=signal: signals[s].__setitem__("out", None))
            signals[signal]["out"] = [x, y]
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

def dispatch_save_point(x, y):
    global capture_module

    if capture_module == "MASTER":
        master_save_point(x, y)


def dispatch_undo():

    if capture_module == "MASTER":
        undo_last_capture()

    elif capture_module == "ROUTE":
        pass

def on_press(key):
    global last_space_time, last_backspace_time, last_digit_time

    if key == pynput_keyboard.Key.space:
        now = time.time()
        if now - last_space_time < 0.4:
            return
        last_space_time = now

        if capture_module != "MASTER":
            return

        x, y = win32api.GetCursorPos()
        root.after(0, lambda: dispatch_save_point(x, y))

    elif key == pynput_keyboard.Key.backspace:
        now = time.time()
        if now - last_backspace_time < 0.4:
            return
        last_backspace_time = now

        if capture_module != "MASTER":
            return
        root.after(0, dispatch_undo)

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

#CLICK

def click(point):

    if point is None:
        return False

    x, y = point

    pyautogui.moveTo(x, y, duration=1)

    log(f"Mouse moved to ({x}, {y})")

    time.sleep(5)   # Don't click yet

    pyautogui.click()

    return True
# =========================================================
# CAPTURE POINT
# =========================================================

def capture_point():

    global capture_waiting
    global captured_point

    capture_waiting = True

    captured_point = None

    while capture_waiting:

        root.update()

        time.sleep(0.05)

    return list(captured_point)

# =========================================================
# CAPTURE TRACK COORDINATES
# =========================================================

def capture_track_points():

    global track_points

    root.iconify()

    for track in sorted(track_points.keys()):

        messagebox.showinfo(
            "TRACK CAPTURE",
            f"Move mouse on track {track}\n\nPress SPACE"
        )

        track_points[track] = capture_point()

    root.deiconify()

    log("ALL TRACK POINTS CAPTURED")

def save_config():
    global config_file_path

    if config_file_path and os.path.exists(config_file_path):

        wb = load_workbook(config_file_path)

        if "ROUTE_POINTS" in wb.sheetnames:
            del wb["ROUTE_POINTS"]

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
        "Menu_X",
        "Menu_Y",
        "Normal_X",
        "Normal_Y",
        "Reverse_X",
        "Reverse_Y",
        "Free_X",
        "Free_Y"
    ])

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    ws_ch = wb.create_sheet("CRANK_HANDLES")

    ws_ch.append([
        "CRANK_HANDLE",
        "Menu_X",
        "Menu_Y",
        "IN_X",
        "IN_Y",
        "OUT_X",
        "OUT_Y",
        "ECH_X",
        "ECH_Y",
        "FREE_X",
        "FREE_Y"
    ])

    # =====================================================
    # LC GATES
    # =====================================================

    ws_lc = wb.create_sheet("LC_GATES")

    ws_lc.append([
        "LC_GATE",
        "Menu_X",
        "Menu_Y",
        "IN_X",
        "IN_Y",
        "OUT_X",
        "OUT_Y"
    ])

    # =====================================================
    # ROUTE POINTS
    # =====================================================

    ws_route = wb.create_sheet("ROUTE_POINTS")

    ws_route.append([

        "SIGNAL",

        "ROUTE",

        "X",

        "Y"

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
            indicator = info.get("state_indicator")
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

            menu = info.get("menu") or info.get("coordinate")
            normal = info.get("normal")
            reverse = info.get("reverse")
            free = info.get("free")

            ws_point.append([
                sig,
                menu[0] if menu else None,
                menu[1] if menu else None,
                normal[0] if normal else None,
                normal[1] if normal else None,
                reverse[0] if reverse else None,
                reverse[1] if reverse else None,
                free[0] if free else None,
                free[1] if free else None
            ])

        # =================================================
        # CRANK HANDLE
        # =================================================

        elif typ == "CH":

            menu = info.get("menu") or info.get("coordinate")
            in_pt = info.get("IN")
            out_pt = info.get("OUT")
            ech = info.get("ECH")
            free = info.get("FREE")

            ws_ch.append([
                sig,
                menu[0] if menu else None,
                menu[1] if menu else None,
                in_pt[0] if in_pt else None,
                in_pt[1] if in_pt else None,
                out_pt[0] if out_pt else None,
                out_pt[1] if out_pt else None,
                ech[0] if ech else None,
                ech[1] if ech else None,
                free[0] if free else None,
                free[1] if free else None
            ])

        # =================================================
        # LC GATE
        # =================================================

        elif typ == "LC":

            menu = info.get("menu") or info.get("coordinate")
            in_pt = info.get("in")
            out_pt = info.get("out")

            ws_lc.append([
                sig,
                menu[0] if menu else None,
                menu[1] if menu else None,
                in_pt[0] if in_pt else None,
                in_pt[1] if in_pt else None,
                out_pt[0] if out_pt else None,
                out_pt[1] if out_pt else None
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

        name = name.strip().upper()

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

        name = name.strip().upper()

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

        name = name.strip().upper()

        signals[name] = {
            "type": "SHUNT",
            "menu": None,
            "state_indicator": None,
            "initial_snapshot": None,
            "route_init": None
        }

        root.iconify()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture MENU\n\nPress SPACE"
        )
        signals[name]["menu"] = capture_point()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture ASPECT INDICATOR\n\nPress SPACE"
        )
        indicator = capture_point()
        signals[name]["state_indicator"] = indicator

        if indicator:
            x, y = indicator
            screenshot = pyautogui.screenshot()
            region = screenshot.crop((x - 50, y - 50, x + 50, y + 50))
            snapshot_file = os.path.abspath(
                os.path.join(SNAPSHOTS_FOLDER, f"{name}_initial.png")
            )
            region.save(snapshot_file)
            signals[name]["initial_snapshot"] = snapshot_file

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture ROUTE INITIATION INDICATOR\n\nPress SPACE"
        )
        signals[name]["route_init"] = capture_point()

        root.deiconify()

    log("ALL SIGNALS RECORDED")

# =========================================================
# PREPARE CAPTURE
# =========================================================

def prepare_capture():

    global capture_queue
    global capture_index
    global current_step

    capture_queue = list(signals.keys())

    capture_index = 0

    current_step = 0

    root.iconify()

    next_capture()

# =========================================================
# NEXT CAPTURE
# =========================================================

def next_capture():

    global current_signal

    if capture_index >= len(capture_queue):

        root.deiconify()

        save_config()

        log("ALL SIGNALS RECORDED")

        return

    current_signal = capture_queue[capture_index]

    info = signals[current_signal]

    sig_type = info["type"]

    # =====================================================
    # MAIN
    # =====================================================

    if sig_type == "MAIN":

        aspects = info["aspects"]

        steps = [

            "MENU",

            "RED",

            "YELLOW",

            "DOUBLE_YELLOW",

            "GREEN"
        ]

        if aspects == 2:

            steps = ["MENU", "RED", "GREEN"]

        elif aspects == 3:

            steps = ["MENU", "RED", "YELLOW", "GREEN"]

        current = steps[current_step]

        log(
            f"{current_signal} -> Capture {current}"
        )

    # =====================================================
    # CALLING ON
    # =====================================================

    elif sig_type == "CALLING_ON":

        steps = ["MENU", "YELLOW", "ROUTE_INIT"]

        current = steps[current_step]

        log(
            f"{current_signal} -> Capture {current}"
        )

    # =====================================================
    # SHUNT
    # =====================================================

    elif sig_type == "SHUNT":

        steps = ["MENU", "INDICATOR", "ROUTE_INIT"]

        current = steps[current_step]

        log(
            f"{current_signal} -> Capture {current}"
        )

# =========================================================
# SAVE POINT
# =========================================================

def save_point(x, y):

    global current_step
    global capture_index

    signal = current_signal
    info = signals[signal]
    sig_type = info["type"]

    if sig_type == "MAIN":
        aspects = info["aspects"]
        if aspects == 2:
            steps = ["MENU", "RED", "GREEN"]
        elif aspects == 3:
            steps = ["MENU", "RED", "YELLOW", "GREEN"]
        else:
            steps = ["MENU", "RED", "YELLOW", "DOUBLE_YELLOW", "GREEN"]

        current_name = steps[current_step]
        signals[signal][current_name] = [x, y]

    elif sig_type == "CALLING_ON":
        steps = ["MENU", "YELLOW", "ROUTE_INIT"]
        current_name = steps[current_step]

        if current_name == "MENU":
            signals[signal]["menu"] = [x, y]
        elif current_name == "YELLOW":
            signals[signal]["YELLOW"] = [x, y]
        elif current_name == "ROUTE_INIT":
            signals[signal]["route_init"] = [x, y]

    elif sig_type == "SHUNT":
        steps = ["MENU", "INDICATOR", "ROUTE_INIT"]
        current_name = steps[current_step]

        if current_name == "MENU":
            signals[signal]["menu"] = [x, y]

        elif current_name == "INDICATOR":
            signals[signal]["state_indicator"] = [x, y]
            screenshot = pyautogui.screenshot()
            region = screenshot.crop((x - 50, y - 50, x + 50, y + 50))
            snapshot_file = os.path.abspath(
                os.path.join(SNAPSHOTS_FOLDER, f"{signal}_initial.png")
            )
            region.save(snapshot_file)
            signals[signal]["initial_snapshot"] = snapshot_file

        elif current_name == "ROUTE_INIT":
            signals[signal]["route_init"] = [x, y]

    else:
        steps = ["MENU"]
        current_name = steps[current_step]
        signals[signal][current_name] = [x, y]

    log(f"{signal} {current_name} Saved")

    current_step += 1

    if current_step >= len(steps):
        current_step = 0
        capture_index += 1

    next_capture()

# =========================================================
# KEYBOARD
# =========================================================


# =========================================================
# SAVE CONFIG
# =========================================================

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

    config_file_path = file_path

    signals = {}

    wb = load_workbook(config_file_path)

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

            "menu": [row[1], row[2]],

            "YELLOW": [row[3], row[4]],

            "route_init": [row[5], row[6]]
            if len(row) > 6 and row[5] is not None else None

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

        snapshot_path = row[7] if len(row) > 7 else None
        if snapshot_path:
            snapshot_path = os.path.abspath(str(snapshot_path))

        signals[signal_name] = {
            "type": "SHUNT",
            "menu": [row[1], row[2]],
            "state_indicator": [row[3], row[4]],
            "route_init": [row[5], row[6]],
            "initial_snapshot": (
                snapshot_path
                if snapshot_path and os.path.exists(snapshot_path)
                else None
            )
        }

    # =====================================================
    # POINTS
    # =====================================================

    if "POINTS" in wb.sheetnames:

        ws = wb["POINTS"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            point_name = str(row[0]).replace(".0", "").strip().upper()

            menu = [row[1], row[2]] if len(row) >= 3 and row[1] is not None and row[2] is not None else None
            normal = [row[3], row[4]] if len(row) >= 5 and row[3] is not None and row[4] is not None else None
            reverse = [row[5], row[6]] if len(row) >= 7 and row[5] is not None and row[6] is not None else None
            free = [row[7], row[8]] if len(row) >= 9 and row[7] is not None and row[8] is not None else None

            signals[point_name] = {
                "type": "POINT",
                "menu": menu,
                "normal": normal,
                "reverse": reverse,
                "free": free,
                # Backward-compatible alias.
                "coordinate": menu
            }

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    if "CRANK_HANDLES" in wb.sheetnames:

        ws = wb["CRANK_HANDLES"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            ch_name = str(row[0]).replace(".0", "").strip().upper()

            menu = [row[1], row[2]] if len(row) >= 3 and row[1] is not None and row[2] is not None else None
            in_pt = [row[3], row[4]] if len(row) >= 5 and row[3] is not None and row[4] is not None else None
            out_pt = [row[5], row[6]] if len(row) >= 7 and row[5] is not None and row[6] is not None else None
            ech = [row[7], row[8]] if len(row) >= 9 and row[7] is not None and row[8] is not None else None
            free = [row[9], row[10]] if len(row) >= 11 and row[9] is not None and row[10] is not None else None

            signals[ch_name] = {
                "type": "CH",
                "menu": menu,
                "IN": in_pt,
                "OUT": out_pt,
                "ECH": ech,
                "FREE": free,
                "coordinate": menu
            }

    # =====================================================
    # LC GATES
    # =====================================================

    if "LC_GATES" in wb.sheetnames:

        ws = wb["LC_GATES"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            lc_name = str(row[0]).replace(".0", "").strip().upper()

            menu = [row[1], row[2]] if len(row) >= 3 and row[1] is not None and row[2] is not None else None
            in_pt = [row[3], row[4]] if len(row) >= 5 and row[3] is not None and row[4] is not None else None
            out_pt = [row[5], row[6]] if len(row) >= 7 and row[5] is not None and row[6] is not None else None

            signals[lc_name] = {
                "type": "LC",
                "menu": menu,
                "in": in_pt,
                "out": out_pt,
                "coordinate": menu
            }
    log("CONFIG LOADED")

def load_universal_coordinates():
    """Loads coordinates from the multi-sheet Universal Yard Coordinate
    file (the same file every other suite program reads from), instead
    of this program's own native config format (MAIN_SIGNALS /
    CALLING_ON / SHUNT / POINTS / CRANK_HANDLES / LC_GATES). Populates
    the same `signals` dict that load_config() populates, using the
    universal file's actual sheet/column names (MAIN / CAL / SHUNT /
    POINT / CH / LC)."""

    global signals
    global config_file_path

    if not lock_routes_data:
        messagebox.showwarning(
            "WARNING",
            "Please load TOC file first."
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

    loaded_main = loaded_cal = loaded_shunt = loaded_point = loaded_ch = loaded_lc = 0

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
                snap_val = cell(row, hi.get("SNAPSHOT_PATH"))
                snapshot_path = os.path.abspath(str(snap_val).strip()) if snap_val else None
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "state_indicator": parse_point(cell(row, hi.get("INDICATOR_X")), cell(row, hi.get("INDICATOR_Y"))),
                    "route_init": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y"))),
                    "initial_snapshot": snapshot_path if snapshot_path and os.path.exists(snapshot_path) else None
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
                menu = parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y")))
                signals[name] = {
                    "type": "POINT",
                    "menu": menu,
                    "normal": parse_point(cell(row, hi.get("NORMAL_X")), cell(row, hi.get("NORMAL_Y"))),
                    "reverse": parse_point(cell(row, hi.get("REVERSE_X")), cell(row, hi.get("REVERSE_Y"))),
                    "free": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y"))),
                    "coordinate": menu
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
                menu = parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y")))
                signals[name] = {
                    "type": "CH",
                    "menu": menu,
                    "IN": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "OUT": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y"))),
                    "ECH": parse_point(cell(row, hi.get("ECH_X")), cell(row, hi.get("ECH_Y"))),
                    "FREE": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y"))),
                    "coordinate": menu
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
                name = str(cell(row, name_col)).replace(".0", "").strip().upper()
                menu = parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y")))
                signals[name] = {
                    "type": "LC",
                    "menu": menu,
                    "in": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "out": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y"))),
                    "coordinate": menu
                }
                loaded_lc += 1

    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_cal} CALLING-ON, {loaded_shunt} SHUNT, "
        f"{loaded_point} POINT, {loaded_ch} CH, {loaded_lc} LC"
    )


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
# YELLOW DETECT
# =========================================================

def is_yellow(point):

    if point is None:
        return False

    screenshot = pyautogui.screenshot()

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

def is_route_initiated(point):
    """Return True when the configured route-initiation indicator is yellow."""
    if point is None:
        return False

    screenshot = pyautogui.screenshot()
    r, g, b = get_avg_color(point[0], point[1], screenshot)
    initiated = r > 150 and g > 150 and b < 140
    log(f"ROUTE INIT RGB = {r},{g},{b} -> {'YELLOW' if initiated else 'NOT YELLOW'}")
    return initiated


def shunt_compare_by_features(signal, initial_path, screenshot=None):
    """Compare the current SHUNT aspect-indicator region with its initial snapshot."""
    try:
        info = signals.get(signal)
        if not info:
            return 0.0, "Signal not configured"

        point = info.get("state_indicator")
        if point is None:
            return 0.0, "No aspect-indicator coordinate"

        if not initial_path or not os.path.exists(initial_path):
            return 0.0, "Initial snapshot not found"

        initial_img = cv2.imread(initial_path)
        if initial_img is None:
            return 0.0, "Cannot load initial snapshot"

        screenshot = screenshot or pyautogui.screenshot()
        x, y = point
        current_pil = screenshot.crop((x - 50, y - 50, x + 50, y + 50))
        current_img = cv2.cvtColor(np.array(current_pil), cv2.COLOR_RGB2BGR)

        initial_gray = cv2.cvtColor(initial_img, cv2.COLOR_BGR2GRAY)
        current_gray = cv2.cvtColor(current_img, cv2.COLOR_BGR2GRAY)

        kp1, des1 = orb.detectAndCompute(initial_gray, None)
        kp2, des2 = orb.detectAndCompute(current_gray, None)

        if des1 is None or des2 is None:
            return 0.0, "Cannot detect features"

        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        matches = sorted(bf.match(des1, des2), key=lambda m: m.distance)
        good = [m for m in matches if m.distance < 50]
        denom = max(len(kp1), len(kp2))
        score = (len(good) / denom * 100.0) if denom else 0.0

        log(f"{signal} SHUNT Features: Initial={len(kp1)}, Current={len(kp2)}, Matches={len(matches)}, Good={len(good)}")
        log(f"{signal} SHUNT Match Score: {score:.2f}%")

        return score, f"Features matched: {len(matches)}"

    except Exception as e:
        log(f"SHUNT Feature Comparison Error: {e}")
        return 0.0, str(e)


def shunt_restored_to_initial(signal, screenshot=None):
    """For AHEAD testing, a SHUNT is considered restored when its indicator
    region closely matches the captured initial reference after route release."""
    info = signals.get(signal)
    if not info or info.get("type") != "SHUNT":
        return None

    score, details = shunt_compare_by_features(
        signal,
        info.get("initial_snapshot"),
        screenshot
    )

    log(f"{signal} SHUNT restored score = {score:.2f}% | {details}")
    return score >= 70.0

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
# =========================================================
# DETECT SIGNAL TYPE
# =========================================================

def detect_signal_type(lock_route):

    lr = str(lock_route).strip().upper()
    left = lr.split("-")[0]

    # SHUNT has priority over CALLING-ON.
    if left.endswith("SH"):
        return left, "SHUNT"

    if left.endswith("C"):
        return left, "CALLING_ON"

    return left, "MAIN"

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
# =========================================================
# CREATE ADVANCED REPORT
# =========================================================


# =========================================================
# RUN ENGINE
# =========================================================


# =========================================================
# AHEAD SIGNAL TEST - EXCEL LOADER, TEST ENGINE AND REPORT
# =========================================================

def _header_key(value):
    return str(value or "").strip().upper().replace(" ", "_")


def load_lock_routes():
    """Load the TOC sheet used by the ahead-signal red-aspect test."""
    global lock_routes_data, track_points

    file_path = filedialog.askopenfilename(
        title="Select TOC Excel file",
        filetypes=[("Excel", "*.xlsx")]
    )
    if not file_path:
        return

    wb = load_workbook(file_path, data_only=True)
    ws = None
    columns = None
    required = {"SIGNAL", "ROUTE", "SIGNAL_AHEAD"}

    # Accept any sheet name.  The headings, not the tab name, identify the TOC sheet.
    for candidate in wb.worksheets:
        headings = {
            _header_key(cell.value): index
            for index, cell in enumerate(candidate[1])
            if cell.value is not None
        }
        if required.issubset(headings):
            ws = candidate
            columns = headings
            break

    if ws is None:
        messagebox.showerror(
            "TOC FORMAT ERROR",
            "Use headings: SIGNAL, ROUTE, SIGNAL_AHEAD"
        )
        return

    lock_routes_data = []
    track_points = {}

    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(value is not None and str(value).strip() for value in values):
            continue

        def value_for(name):
            index = columns.get(name)
            return values[index] if index is not None and index < len(values) else None

        signal = str(value_for("SIGNAL") or "").replace(".0", "").strip().upper()
        route = str(value_for("ROUTE") or "").replace(".0", "").strip().upper()
        ahead = str(value_for("SIGNAL_AHEAD") or "").replace(".0", "").strip().upper()

        if not signal or not route:
            log("Skipped an Excel row with no SIGNAL or ROUTE")
            continue

        lock_routes_data.append({
            "signal": signal,
            "route": route,
            "signal_ahead": ahead,
        })

    refresh_table()
    log(f"TOC LOADED: {len(lock_routes_data)} routes from sheet '{ws.title}'")


def get_route_signal_type(signal_name):
    """
    Determine whether a route belongs to MAIN, CALLING-ON,
    or SHUNT for display in the three dashboard tabs.

    Priority:
    1. Use the loaded signal configuration when available.
    2. Calling-on fallback: signal name ending in C.
    3. Otherwise treat it as MAIN.
    """

    name = str(signal_name or "").strip().upper()

    info = signals.get(name)

    if info:
        signal_type = info.get("type")

        if signal_type == "CALLING_ON":
            return "CALLING_ON"

        if signal_type == "SHUNT":
            return "SHUNT"

        if signal_type == "MAIN":
            return "MAIN"

    # Existing Ahead Signal convention:
    # signal names ending in C are Calling-On.
    if name.endswith("C"):
        return "CALLING_ON"

    return "MAIN"


def refresh_one_signal_table(tree_widget, signal_type):
    """Refresh one dashboard tab with only its signal type."""

    tree_widget.delete(
        *tree_widget.get_children()
    )

    display_number = 1

    for row in lock_routes_data:

        route_signal = str(
            row.get("signal", "")
        ).strip().upper()

        row_type = get_route_signal_type(
            route_signal
        )

        if row_type != signal_type:
            continue

        tree_widget.insert(
            "",
            "end",
            values=(
                display_number,
                row.get("signal", ""),
                row.get("route", ""),
                row.get("signal_ahead")
                or "NOT APPLICABLE",
            )
        )

        display_number += 1


def refresh_table():
    """
    Refresh all three dashboard tabs:
    MAIN SIGNALS, CALLING-ON SIGNALS and SHUNT SIGNALS.
    """

    refresh_one_signal_table(
        main_tree,
        "MAIN"
    )

    refresh_one_signal_table(
        calling_tree,
        "CALLING_ON"
    )

    refresh_one_signal_table(
        shunt_tree,
        "SHUNT"
    )



def is_signal_red(signal_name, screenshot=None):
    """Return True only when the configured red aspect is visibly ON."""
    signal_name = str(signal_name).replace(".0", "").strip().upper()
    info = signals.get(signal_name)
    if not info:
        log(f"{signal_name}: signal is not present in the configuration")
        return None
    if info.get("type") != "MAIN" or not info.get("RED"):
        log(f"{signal_name}: MAIN signal RED coordinate is not configured")
        return None

    screenshot = screenshot or pyautogui.screenshot()
    x, y = info["RED"]
    red, green, blue = get_avg_color(x, y, screenshot)
    is_red = red > 140 and green < 130 and blue < 130
    log(f"{signal_name} RED RGB = {red},{green},{blue} -> {'ON' if is_red else 'NOT ON'}")
    return is_red

def is_signal_in_expected_final_state(signal_name, screenshot=None):
    """Check the final state according to the configured signal type."""
    signal_name = str(signal_name).replace(".0", "").strip().upper()
    info = signals.get(signal_name)
    if not info:
        log(f"{signal_name}: signal is not present in the configuration")
        return None, "NOT CONFIGURED"

    sig_type = info.get("type")
    if sig_type == "MAIN":
        result = is_signal_red(signal_name, screenshot)
        return result, ("RED (ON)" if result else "NOT RED")

    if sig_type == "SHUNT":
        result = shunt_restored_to_initial(signal_name, screenshot)
        if result is None:
            return None, "SHUNT NOT CONFIGURED"
        return result, ("INITIAL STATE RESTORED" if result else "STATE CHANGED")

    if sig_type == "CALLING_ON":
        yellow = info.get("YELLOW")
        route_init = info.get("route_init")

        if yellow is None:
            return None, "YELLOW NOT CONFIGURED"
        if route_init is None:
            return None, "ROUTE INITIATION NOT CONFIGURED"

        screenshot = screenshot or pyautogui.screenshot()
        r, g, b = get_avg_color(yellow[0], yellow[1], screenshot)
        # After route cancel/release the calling-on signal should not remain yellow.
        not_yellow = not (r > 150 and g > 150)
        return not_yellow, ("NOT YELLOW" if not_yellow else "YELLOW STILL ON")

    return None, "UNSUPPORTED TYPE"


def are_signals_steady_red(route_signal, ahead_signal):
    """Backward-compatible name; now performs type-aware final-state checks."""
    route_states = []
    ahead_states = []
    route_labels = []
    ahead_labels = []

    for sample in range(12):
        if not running:
            return False, False

        pause_event.wait()
        screenshot = pyautogui.screenshot()

        route_state, route_label = is_signal_in_expected_final_state(route_signal, screenshot)
        ahead_state, ahead_label = is_signal_in_expected_final_state(ahead_signal, screenshot)

        route_states.append(route_state is True)
        ahead_states.append(ahead_state is True)
        route_labels.append(route_label)
        ahead_labels.append(ahead_label)

        if sample < 11:
            time.sleep(0.35)

    route_ok = all(route_states)
    ahead_ok = all(ahead_states)

    if not route_ok:
        log(f"{route_signal}: final-state check failed ({route_labels[-1]})")
    if not ahead_ok:
        log(f"{ahead_signal}: final-state check failed ({ahead_labels[-1]})")

    return route_ok, ahead_ok


def create_report():
    global REPORT_FILE

    wb = Workbook()
    ws = wb.active
    ws.title = "AHEAD SIGNAL REPORT"

    # =====================================================
    # REPORT HEADER
    # =====================================================

    ws.merge_cells("A1:E1")
    ws["A1"] = "AHEAD SIGNAL ASPECT TEST REPORT"
    ws["A1"].font = Font(
        name="Calibri",
        size=22,
        bold=True
    )
    ws["A1"].fill = PatternFill(fill_type=None)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:E2")
    ws["A2"] = (
        f"Generated: "
        f"{datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
    )
    ws["A2"].alignment = Alignment(horizontal="center")

    # =====================================================
    # TEST RESULT TABLE
    # =====================================================

    headers = [
        "SIGNAL",
        "ROUTE",
        "AHEAD SIGNAL",
        "TEST RESULT",
        "DATE & TIME"
    ]

    ws.append(headers)

    thin = Side(style="thin", color="000000")
    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in ws[3]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type=None)
        cell.alignment = Alignment(horizontal="left")
        cell.border = border

    for report in report_rows:

        ws.append([
            report.get("ROUTE_SIGNAL", ""),
            report.get("ROUTE", ""),
            report.get("AHEAD_SIGNAL", "NOT APPLICABLE"),
            report.get("RESULT", ""),
            report.get("DATE & TIME", "")
        ])

    # =====================================================
    # ROUTE SUMMARY
    # =====================================================
    #
    # Example:
    #
    # TOTAL ROUTES          12
    #
    # INITIATED ROUTES       5    1_A, 2_K, 3_K, 8_L, 25_J
    #
    # NON INITIATED ROUTES   7
    #
    # Only PASS routes are listed under INITIATED ROUTES.
    # Routes are counted uniquely, not once per report row.
    # =====================================================

    unique_routes = []
    seen_routes = set()

    for row in lock_routes_data:

        route_name = str(
            row.get("route", "")
        ).strip()

        if not route_name:
            continue

        route_key = route_name.upper()

        if route_key not in seen_routes:
            seen_routes.add(route_key)
            unique_routes.append(route_name)

    total_routes = len(unique_routes)

    # Collect results for each route.
    route_results = {}

    for report in report_rows:

        route_name = str(
            report.get("ROUTE", "")
        ).strip()

        if not route_name:
            continue

        route_key = route_name.upper()

        result = str(
            report.get("RESULT", "")
        ).strip().upper()

        route_results.setdefault(
            route_key,
            []
        ).append(result)

    initiated_route_names = []

    for route_name in unique_routes:

        route_key = route_name.upper()

        results = route_results.get(
            route_key,
            []
        )

        # A route is initiated only when it has a
        # recorded test and all recorded tests PASS.
        if results and all(
            result == "PASS"
            for result in results
        ):
            initiated_route_names.append(route_name)

    initiated_routes = len(
        initiated_route_names
    )

    initiated_route_text = ", ".join(
        initiated_route_names
    )

    non_initiated_routes = max(
        total_routes - initiated_routes,
        0
    )

    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL ROUTES"
    ws[f"B{start}"] = total_routes

    ws[f"A{start + 1}"] = "INITIATED ROUTES"
    ws[f"B{start + 1}"] = initiated_routes
    ws[f"C{start + 1}"] = initiated_route_text

    ws[f"A{start + 2}"] = "NON INITIATED ROUTES"
    ws[f"B{start + 2}"] = non_initiated_routes

    for r in range(start, start + 3):

        for column in ("A", "B", "C"):
            ws[f"{column}{r}"].font = Font(
                bold=True
            )
            ws[f"{column}{r}"].border = border

        ws[f"A{r}"].alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        ws[f"B{r}"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws[f"C{r}"].alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=False
        )

    if initiated_route_text:
        ws.column_dimensions["C"].width = min(
            max(len(initiated_route_text) + 5, 25),
            100
        )

    # =====================================================
    # REPORT ROW FORMATTING
    # =====================================================
    # PASS and NOT APPLICABLE rows stay white.
    # FAIL rows are highlighted RED.

    fail_fill = PatternFill(
        "solid",
        fgColor="FF6666"
    )

    white_fill = PatternFill(
        fill_type=None
    )

    for cells in ws.iter_rows(
        min_row=4,
        max_row=ws.max_row
    ):

        result = str(
            cells[3].value or ""
        ).strip().upper()

        row_fill = (
            fail_fill
            if result == "FAIL"
            else white_fill
        )

        for cell in cells:
            cell.fill = row_fill
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # Summary rows stay white.
    for r in range(start, start + 3):

        ws[f"A{r}"].fill = white_fill
        ws[f"B{r}"].fill = white_fill
        ws[f"C{r}"].fill = white_fill

        ws[f"A{r}"].alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        ws[f"B{r}"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws[f"C{r}"].alignment = Alignment(
            horizontal="left",
            vertical="center",
            wrap_text=False
        )

    # =====================================================
    # COLUMN WIDTHS
    # =====================================================

    widths = [16, 18, 18, 16, 22]

    for index, width in enumerate(
        widths,
        start=1
    ):
        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    # Keep the route-name column wide enough.
    if initiated_route_text:
        ws.column_dimensions["C"].width = min(
            max(
                ws.column_dimensions["C"].width,
                len(initiated_route_text) + 5
            ),
            100
        )

    # =====================================================
    # TIMESTAMPED REPORT FILE
    # =====================================================

    REPORT_FILE = (
        "AHEAD_SIGNAL_REPORT_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        ".xlsx"
    )

    wb.save(REPORT_FILE)

    log(
        f"REPORT SAVED: {REPORT_FILE}"
    )

def run_engine():
    """Set route -> cancel -> release -> prove route and ahead signals are red."""
    global running, report_rows
    report_rows = []
    root.iconify()
    time.sleep(2)

    for row in lock_routes_data:
        pause_event.wait()
        if not running:
            break

        route_signal = row["signal"]
        route = row["route"]
        ahead_signal = row["signal_ahead"]

        if not ahead_signal or ahead_signal in {"NA", "N/A", "NOT APPLICABLE"}:
            result = "NOT APPLICABLE"
            route_red = "NOT CHECKED"
            ahead_red = "STOP BOARD / NO AHEAD SIGNAL"
            log(f"{route_signal} {route}: skipped - no ahead signal")
        elif route_signal not in signals or ahead_signal not in signals:
            result = "FAIL"
            route_red = "NOT CONFIGURED"
            ahead_red = "NOT CONFIGURED"
            log(f"{route_signal} {route}: route or ahead signal is not configured")
        else:
            log(f"SETTING ROUTE: {route_signal} -> {route}; AHEAD SIGNAL: {ahead_signal}")
            click(signals[route_signal]["menu"])
            time.sleep(1)
            route_set = click_menu_item(route.replace("-", "_"))

            # Give the panel enough time to react before checking the
            # Route Initiation Indicator.  Calling-On follows the TL
            # workflow: route initiation is checked before cleanup.
            wait_after_route = 7 if signals[route_signal].get("type") == "CALLING_ON" else 5
            time.sleep(wait_after_route)

            if not route_set:
                result = "FAIL"
                route_red = "ROUTE NOT SET"
                ahead_red = "NOT CHECKED"
                log(f"{route_signal} {route}: route menu item was not found")
            else:
                route_type = signals[route_signal].get("type")
                route_initiated = True

                # -----------------------------------------------------
                # CALLING-ON: Route Initiation Indicator is mandatory.
                # Do not send Signal Cancel / Route Release unless the
                # route was actually initiated.
                # -----------------------------------------------------
                if route_type == "CALLING_ON":
                    route_init_point = signals[route_signal].get("route_init")

                    if route_init_point is None:
                        route_initiated = False
                        log(
                            f"{route_signal}: Route Initiation coordinate is not configured"
                        )
                    else:
                        route_initiated = is_route_initiated(route_init_point)
                        log(
                            f"{route_signal} Route Initiation Indicator: "
                            f"{'YELLOW' if route_initiated else 'NOT YELLOW'}"
                        )

                if route_type == "CALLING_ON" and not route_initiated:
                    result = "FAIL"
                    route_red = "ROUTE NOT INITIATED"
                    ahead_red = "NOT CHECKED"
                    log(
                        f"{route_signal} {route}: Route Initiation Indicator "
                        "did not become YELLOW. Signal Cancel / Route Release skipped."
                    )

                else:
                    # Signal cancel
                    click(signals[route_signal]["menu"])
                    time.sleep(1)
                    cancel_done = click_menu_item("Signal Cancel")
                    time.sleep(2)

                    # Route release
                    click(signals[route_signal]["menu"])
                    time.sleep(1)
                    release_done = click_menu_item("Route Release")
                    time.sleep(6)

                    route_is_red, ahead_is_red = are_signals_steady_red(
                        route_signal,
                        ahead_signal
                    )
                    route_red = "RED (ON)" if route_is_red else "NOT RED"
                    ahead_red = "RED (ON)" if ahead_is_red else "NOT RED"

                    result = (
                        "PASS"
                        if cancel_done and release_done and route_is_red and ahead_is_red
                        else "FAIL"
                    )

                    if route_type == "CALLING_ON":
                        log(
                            f"{route_signal} {route}: Route Initiated={route_initiated} | {result}"
                        )
                    else:
                        log(f"{route_signal} {route}: {result}")

            pyautogui.press("esc")
            time.sleep(1)

        report_rows.append({
            "ROUTE_SIGNAL": route_signal,
            "ROUTE": route,
            "AHEAD_SIGNAL": ahead_signal or "NOT APPLICABLE",
            "ROUTE_SIGNAL_RED": route_red,
            "AHEAD_SIGNAL_RED": ahead_red,
            "ROUTE_INITIATED": (
                "YES"
                if route_signal in signals
                and signals[route_signal].get("type") == "CALLING_ON"
                and 'route_initiated' in locals()
                and route_initiated
                else (
                    "NO"
                    if route_signal in signals
                    and signals[route_signal].get("type") == "CALLING_ON"
                    and 'route_initiated' in locals()
                    else "NOT CHECKED"
                )
            ),
            "RESULT": result,
            "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        })

    running = False
    status_label.config(text="COMPLETED", fg="#16a34a")
    create_report()
    root.deiconify()
    log("AHEAD SIGNAL TESTING COMPLETED")
    try:
        os.startfile(REPORT_FILE)
    except Exception:
        pass


# =========================================================
# START
# =========================================================

def start_automation():

    global running

    if not signals:

        log("LOAD CONFIG FIRST")

        return

    if not lock_routes_data:

        log("LOAD LOCK ROUTES FIRST")

        return

    # =========================================================
    # PRE-FLIGHT CHECK
    # If none of the TOC signals match a loaded coordinate, every
    # row will hit "NOT FOUND" and the run will finish in a couple
    # seconds with a report full of failures - looking exactly
    # like nothing was tested. Catch that here, loudly, before
    # starting.
    # =========================================================

    toc_signals = {row["signal"] for row in lock_routes_data}
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
    "AHEAD SIGNAL TESTING AUTOMATION"
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

logo_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Indian_Railway_Logo_2.png"
)

try:

    logo_image = PILImage.open(logo_path)
    logo_image.thumbnail((85, 85), PILImage.Resampling.LANCZOS)

    railway_logo = ImageTk.PhotoImage(logo_image)

    logo_label = tk.Label(
        left_logo_frame,
        image=railway_logo,
        bg="#0f172a",
        bd=0
    )

    logo_label.image = railway_logo
    logo_label.pack()

except Exception:

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

    text="AHEAD SIGNAL TESTING SYSTEM",

    font=("Segoe UI", 24, "bold"),

    bg="#0f172a",

    fg="white"

).pack(
    pady=(15, 0)
)

tk.Label(

    title_frame,

    text="COORDINATE CAPTURE  |  ROUTE TESTING  |  REPORTING",

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
        file="edrc_logo.png"
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
    pady=(15, 45)
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
    "LOAD TOC EXCEL",
    load_lock_routes,
    "#2563eb"
).pack(pady=8)

create_button(
    "SAVE NEW SIGNAL COORDINATES",
    save_new_signal_coordinates,
    "#16a34a"
).pack(pady=8)

create_button(
    "LOAD YARD COORDINATES",
    load_universal_coordinates,
    "#7c3aed"
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
    lambda e: refresh_table()
)

# =========================================================
# SIGNAL TABLE FACTORY
# =========================================================

def create_signal_tree(parent):

    frame = tk.Frame(
        parent,
        bg="white",
        bd=1,
        relief="solid"
    )

    frame.pack(
        fill="both",
        expand=True
    )

    scrollbar = ttk.Scrollbar(
        frame
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )

    tree_widget = ttk.Treeview(

        frame,

        columns=(
            "NO",
            "SIGNAL",
            "ROUTE",
            "AHEAD_SIGNAL",
        ),

        show="headings",

        yscrollcommand=scrollbar.set
    )

    scrollbar.config(
        command=tree_widget.yview
    )

    tree_widget.heading(
        "NO",
        text="NO"
    )

    tree_widget.heading(
        "SIGNAL",
        text="ROUTE SIGNAL"
    )

    tree_widget.heading(
        "ROUTE",
        text="ROUTE"
    )

    tree_widget.heading(
        "AHEAD_SIGNAL",
        text="AHEAD SIGNAL"
    )

    tree_widget.column(
        "NO",
        width=80,
        anchor="center"
    )

    tree_widget.column(
        "SIGNAL",
        width=140,
        anchor="center"
    )

    tree_widget.column(
        "ROUTE",
        width=150,
        anchor="center"
    )

    tree_widget.column(
        "AHEAD_SIGNAL",
        width=150,
        anchor="center"
    )

    tree_widget.pack(
        fill="both",
        expand=True
    )

    return tree_widget


# =========================================================
# MAIN / CALLING-ON / SHUNT TABLES
# =========================================================

main_tree = create_signal_tree(
    main_tab
)

calling_tree = create_signal_tree(
    calling_tab
)

shunt_tree = create_signal_tree(
    shunt_tab
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
# FIXED BOTTOM FOOTER / STATUS BAR
# =========================================================
# Use place() so the footer stays visually fixed at the very
# bottom of the application window. The main content leaves
# enough bottom space so the footer text cannot be clipped.

FOOTER_HEIGHT = 36

footer = tk.Frame(
    root,
    bg="#0f172a",
    height=FOOTER_HEIGHT
)

footer.place(
    relx=0,
    rely=1.0,
    anchor="sw",
    relwidth=1.0,
    height=FOOTER_HEIGHT
)

footer_label = tk.Label(
    footer,
    text=(
        "SPACE = CAPTURE COORDINATES   |   "
        "P = PAUSE / RESUME   |   "
        "AHEAD SIGNAL TESTING SYSTEM"
    ),
    bg="#0f172a",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    anchor="center"
)

footer_label.pack(
    fill="both",
    expand=True
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

root.mainloop()
