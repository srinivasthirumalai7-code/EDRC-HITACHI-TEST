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
import pyautogui

import win32api
import pythoncom
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

REPORT_FILE = "CASCADING_AUTOMATION_REPORT.xlsx"

# =========================================================
# PERSISTENT LOG FILE
# -----------------------------------------------------
# The GUI window can close before anyone gets a chance to
# read the on-screen log (e.g. when run unattended through
# the suite, or if it exits quickly). Every log() call also
# gets written here, flushed immediately, so there's always
# a readable trace afterward regardless of window timing.
# =========================================================

LOG_FILE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    f"CASCADING_LOG_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
)

try:
    _log_file_handle = open(LOG_FILE_PATH, "a", encoding="utf-8")
    _log_file_handle.write(
        f"===== CASCADING.py SESSION STARTED "
        f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====\n"
    )
    _log_file_handle.flush()
except Exception:
    _log_file_handle = None

# =========================================================
# GLOBALS
# =========================================================

signals = {}

lock_routes_data = []

track_points = {}

running = False

capture_module = None

# Route capture globals
route_capture_mode = False
route_capture_index = 0
route_points = {}
routes_data = []
cascading_routes_data = []

capture_index = 0

capture_master_list = []

record_stage = None

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
# =========================================================
# LOG
# =========================================================

def log(msg):

    line = f"[{time.strftime('%H:%M:%S')}] {msg}"

    if _log_file_handle is not None:
        try:
            _log_file_handle.write(line + "\n")
            _log_file_handle.flush()
        except Exception:
            pass

    def write():

        log_text.config(state="normal")

        log_text.insert(
            tk.END,
            line + "\n"
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

    capture_index += 1

    if capture_index < len(capture_master_list):

        set_initial_stage_for_current()

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

        if record_stage == "menu":

            push_undo(
                lambda s=signal:
                signals[s].__setitem__("menu", None)
            )

            signals[signal]["menu"] = [x, y]

            log(f"{signal} MENU Saved")

            record_stage = "indicator"

            master_next_capture()
            return

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

            log(f"{signal} INDICATOR Saved")

            log(
                f"{signal} Snapshot: {snapshot_file}"
            )

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

def dispatch_save_point(x, y):

    global capture_module
    global route_capture_mode
    global route_capture_index

    log(f"Dispatch -> {x}, {y}")

    if route_capture_mode:

        row = routes_data[route_capture_index]

        key = f'{row["signal"]}_{row["route"]}'

        route_points[key] = [x, y]

        log(f"Route Point Saved : {key}")

        route_capture_index += 1

        update_route_capture()

        return

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
# CLICK
# =========================================================

def click(point):

    if point is None:
        return False

    # -----------------------------------------
    # ALWAYS ACTIVATE TEST PANEL FIRST
    # -----------------------------------------

    if not bring_test_panel_to_front():

        log("WARNING: TEST PANEL COULD NOT BE ACTIVATED")

        return False

    x, y = point

    pyautogui.moveTo(
        x,
        y,
        duration=0.5
    )

    log(
        f"Mouse moved to ({x}, {y})"
    )

    time.sleep(1)

    pyautogui.click()

    log(
        f"Clicked coordinate ({x}, {y})"
    )

    return True

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


# =========================================================
# SAVE CONFIG
# =========================================================

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
        "Route_Init_X",
        "Route_Init_Y"
    ])

    # =====================================================
    # SHUNT
    # =====================================================

    ws_shunt = wb.create_sheet("SHUNT")

    ws_shunt.append([
        "Signal",
        "Menu_X",
        "Menu_Y",
        "Indicator_X",
        "Indicator_Y",
        "Route_Init_X",
        "Route_Init_Y"
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

            indicator = info.get("indicator")

            route_init = info.get("route_init")

            ws_shunt.append([

                sig,

                menu[0] if menu else None,

                menu[1] if menu else None,

                indicator[0] if indicator else None,

                indicator[1] if indicator else None,

                route_init[0] if route_init else None,

                route_init[1] if route_init else None

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
    # WRITE ROUTE POINTS
    # =====================================================

    for key, point in route_points.items():
        signal, route = key.split("_", 1)

        ws_route.append([

            signal,

            route,

            point[0],

            point[1]

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
# LOAD CONFIG
# =========================================================

def load_config():

    global signals
    global route_points
    global config_file_path
    global cascading_routes_data

    # Never reset signal coordinates out from under an in-progress
    # automation run - this can be triggered redundantly by the suite
    # even after auto_start_from_launcher() has already started things.
    if running:
        log("LOAD CONFIG BLOCKED : AUTOMATION RUNNING")
        return

    log("================================")
    log("LOAD CONFIG STARTED")

    # =====================================================
    # CHECK TOC
    # =====================================================

    log(
        f"TOC rows loaded = "
        f"{len(cascading_routes_data)}"
    )

    if not cascading_routes_data:

        log(
            "ERROR: cascading_routes_data IS EMPTY"
        )

        messagebox.showwarning(
            "WARNING",
            "Please load the Sequential Route Release workbook first."
        )

        return

    # =====================================================
    # GET CONFIG PATH
    # =====================================================

    file_path = os.environ.get("EDRC_COORDS")

    log(
        f"EDRC_COORDS RAW = {file_path}"
    )

    # -----------------------------------------------------
    # If launcher supplied only filename,
    # convert to absolute path
    # -----------------------------------------------------

    if file_path and not os.path.isabs(file_path):

        file_path = os.path.abspath(file_path)

    log(
        f"EDRC_COORDS FULL = {file_path}"
    )

    log(
        f"CONFIG EXISTS = "
        f"{os.path.exists(file_path) if file_path else False}"
    )

    # =====================================================
    # FALLBACK FILE DIALOG
    # =====================================================

    if not file_path or not os.path.exists(file_path):

        log(
            "CONFIG FILE NOT FOUND FROM EDRC_COORDS"
        )

        log(
            "OPENING CONFIG FILE DIALOG"
        )

        file_path = filedialog.askopenfilename(

            title="Select Signal Configuration",

            filetypes=[
                ("Excel Files", "*.xlsx")
            ]
        )

    if not file_path:

        log(
            "CONFIG FILE SELECTION CANCELLED"
        )

        return

    # =====================================================
    # SAVE PATH
    # =====================================================

    config_file_path = file_path

    log(
        f"OPENING CONFIG = {config_file_path}"
    )

    # =====================================================
    # LOAD EXCEL
    # =====================================================

    try:

        wb = load_workbook(
            config_file_path
        )

        log(
            "CONFIG OPENED SUCCESSFULLY"
        )

        log(
            f"SHEETS = {wb.sheetnames}"
        )

    except Exception as e:

        log(
            f"CONFIG LOAD ERROR = {type(e).__name__}: {e}"
        )

        messagebox.showerror(
            "CONFIG LOAD ERROR",
            f"Could not open configuration file.\n\n{e}"
        )

        return

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

        log("CONFIG LOADED")

        return

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

        aspects = int(row[1]) if row[1] is not None else 0

        green_coord = (
            [row[10], row[11]]
            if row[10] is not None and row[11] is not None
            else None
        )

        yellow_coord = (
            [row[6], row[7]]
            if row[6] is not None and row[7] is not None
            else None
        )

        # 2-ASPECT SIGNAL:
        # Yellow and Green use the same physical coordinate.
        if aspects == 2 and yellow_coord is None:
            yellow_coord = green_coord
            log(
                f"{signal_name} 2-ASPECT: "
                f"YELLOW coordinate missing, using GREEN coordinate {green_coord}"
            )

        signals[signal_name] = {
            "type": "MAIN",
            "aspects": aspects,
            "menu": [row[2], row[3]],
            "RED": [row[4], row[5]],
            "YELLOW": yellow_coord,
            "DOUBLE_YELLOW": [row[8], row[9]]
            if row[8] is not None and row[9] is not None else None,
            "GREEN": green_coord,
            "ROUTE_INDICATOR": [row[12], row[13]]
            if row[12] is not None and row[13] is not None else None
        }

    # =====================================================
    # CALLING ON
    # =====================================================

    if "CALLING_ON" in wb.sheetnames:

        ws = wb["CALLING_ON"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row or not row[0]:
                continue

            signal_name = str(
                row[0]
            ).replace(".0", "").strip().upper()

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

    if "SHUNT" in wb.sheetnames:

        ws = wb["SHUNT"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row or not row[0]:
                continue

            signal_name = str(
                row[0]
            ).replace(".0", "").strip().upper()

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

                "initial_snapshot": None

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
            # IN = OPEN
            signals[name]["in"] = (
                [row[3], row[4]]
                if row[3] is not None and row[4] is not None
                else None
            )
            # OUT = CLOSE
            signals[name]["out"] = (
                [row[5], row[6]]
                if row[5] is not None and row[6] is not None
                else None
            )

    # =====================================================
    # LOAD ROUTE POINTS
    # =====================================================

    if "ROUTE_POINTS" in wb.sheetnames:

        ws = wb["ROUTE_POINTS"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row[0]:
                continue

            signal = str(row[0]).strip().upper()
            route = str(row[1]).strip().upper()

            route_points[f"{signal}_{route}"] = [
                int(row[2]),
                int(row[3])
            ]

        log(f"ROUTE POINTS LOADED : {len(route_points)}")

    log("CONFIG LOADED")


def _load_universal_signals_from_workbook(wb):
    """Populates `signals` (and route_points, where applicable) from an
    already-open Universal Yard Coordinate workbook, reading MAIN / CAL
    / SHUNT / POINT / CH / LC sheets by HEADER NAME. Shared by
    load_config()'s auto-detection above and load_universal_coordinates()
    below, so both paths behave identically."""

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
    # route_points has no equivalent data in the universal file - left
    # empty when loading this way.


def load_universal_coordinates():
    """Explicit button/manual entry point: prompts for a Universal Yard
    Coordinate file and loads it via the same parser load_config() uses
    automatically when it detects the format."""

    global signals
    global config_file_path

    if running:
        log("LOAD COORDINATES BLOCKED : AUTOMATION RUNNING")
        return

    if not cascading_routes_data:
        messagebox.showwarning(
            "WARNING",
            "Please load the Sequential Route Release / Cascading TOC first."
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
# LOAD LOCK ROUTE EXCEL
# =========================================================

# =========================================================
# LOAD LOCK ROUTE EXCEL
# HEADER-BASED COLUMN DETECTION
# =========================================================

def load_cascading_routes():

    global cascading_routes_data

    # Never reset the TOC out from under an in-progress automation
    # run - this can be triggered redundantly by the suite even after
    # auto_start_from_launcher() has already started things.
    if running:
        log("LOAD TOC BLOCKED : AUTOMATION RUNNING")
        return

    file = (
        os.environ.get("EDRC_TOC")
        or os.environ.get("EDRC_LIST")
    )

    # =====================================================
    # SELECT FILE
    # =====================================================

    if not file or not os.path.exists(file):

        file = filedialog.askopenfilename(
            title="Select Cascading TOC",
            filetypes=[("Excel", "*.xlsx")]
        )

    if not file:
        log("TOC FILE SELECTION CANCELLED")
        return

    # =====================================================
    # OPEN WORKBOOK
    # =====================================================

    try:

        wb = load_workbook(file, data_only=True)

    except Exception as e:

        log(f"TOC LOAD ERROR : {type(e).__name__}: {e}")

        messagebox.showerror(
            "TOC LOAD ERROR",
            f"Could not open TOC workbook.\n\n{e}"
        )

        return

    log(f"TOC WORKBOOK OPENED : {file}")
    log(f"AVAILABLE SHEETS : {wb.sheetnames}")

    # =====================================================
    # FIND TOC SHEET
    # =====================================================

    ws = None

    for sheet_name in wb.sheetnames:

        if sheet_name.strip().upper() == "TOC":

            ws = wb[sheet_name]

            log(f"TOC SHEET FOUND : {sheet_name}")

            break

    if ws is None:

        messagebox.showerror(
            "ERROR",
            "TOC SHEET NOT FOUND"
        )

        log("ERROR : TOC SHEET NOT FOUND")

        return

    # =====================================================
    # FIND HEADER ROW
    # =====================================================
    #
    # We search the first several rows for the required
    # column headers.
    #
    # This means the headers do NOT have to be in row 1.
    #
    # =====================================================

    required_headers = {

        "signal": [
            "SIGNAL",
            "SIGNA",
            "MAIN SIGNAL"
        ],

        "route": [
            "ROUTE",
            "MAIN ROUTE"
        ],

        "cascading_signal": [
            "CASCADING SIGNAL",
            "CASCADE SIGNAL"
        ],

        "cascading_route": [
            "CASCADING ROUTE",
            "CASCADE ROUTE"
        ],

        "cascading": [
            "CASCADING",
            "CONTROL BIT",
            "CONTROL BITS"
        ],

        "extra_signal": [
            "EXTRA SIGNAL"
        ],

        "extra_route": [
            "EXTRA ROUTE"
        ]

    }

    header_row_number = None
    column_map = {}

    # Search the ENTIRE TOC sheet for the header row.
    # Column position does not matter.
    # Header row position does not matter.

    for row_number in range(1, ws.max_row + 1):

        current_headers = {}

        for column_number in range(1, ws.max_column + 1):

            value = ws.cell(
                row=row_number,
                column=column_number
            ).value

            if value is None:
                continue

            header = str(value).strip().upper()

            # Treat underscore and space as the same
            header = header.replace("_", " ")

            # Remove extra spaces
            header = " ".join(header.split())

            current_headers[header] = column_number

        # -------------------------------------------------
        # Check whether this row contains required headers
        # -------------------------------------------------

        found_map = {}

        for field_name, possible_names in required_headers.items():

            for possible_name in possible_names:

                possible_name = " ".join(
                    possible_name.strip().upper().split()
                )

                if possible_name in current_headers:

                    found_map[field_name] = current_headers[
                        possible_name
                    ]

                    break

        # -------------------------------------------------
        # Minimum required columns
        # -------------------------------------------------

        required_core = [

            "signal",
            "route",
            "cascading_signal",
            "cascading_route"

        ]

        if all(
            field in found_map
            for field in required_core
        ):

            header_row_number = row_number
            column_map = found_map

            break

    # =====================================================
    # HEADER VALIDATION
    # =====================================================

    if header_row_number is None:

        log("ERROR : REQUIRED TOC HEADERS NOT FOUND")

        messagebox.showerror(
            "TOC HEADER ERROR",
            "Could not find the required TOC columns.\n\n"
            "Required headers:\n"
            "SIGNAL\n"
            "ROUTE\n"
            "CASCADING SIGNAL\n"
            "CASCADING ROUTE"
        )

        return

    # =====================================================
    # LOG DETECTED COLUMNS
    # =====================================================

    log(
        f"TOC HEADER ROW FOUND : {header_row_number}"
    )

    for field_name, column_number in column_map.items():

        column_letter = get_column_letter(column_number)

        log(
            f"TOC COLUMN : {field_name.upper()} "
            f"= {column_letter}"
        )

    # =====================================================
    # CHECK OPTIONAL COLUMNS
    # =====================================================

    if "cascading" not in column_map:

        log(
            "WARNING : CASCADING / CONTROL BIT COLUMN "
            "NOT FOUND"
        )

    if "extra_signal" not in column_map:

        log(
            "INFO : EXTRA SIGNAL COLUMN NOT FOUND"
        )

    if "extra_route" not in column_map:

        log(
            "INFO : EXTRA ROUTE COLUMN NOT FOUND"
        )

    # =====================================================
    # READ TOC DATA
    # =====================================================

    cascading_routes_data = []

    for row_number in range(
        header_row_number + 1,
        ws.max_row + 1
    ):

        # -------------------------------------------------
        # GET VALUE HELPER
        # -------------------------------------------------

        def get_value(field_name):

            column_number = column_map.get(field_name)

            if column_number is None:
                return ""

            value = ws.cell(
                row=row_number,
                column=column_number
            ).value

            if value is None:
                return ""

            # Excel may return 1.0 for a numeric value that is really 1.
            # Convert only genuine integer floats.
            if isinstance(value, float) and value.is_integer():
                value = int(value)

            return str(value).strip().upper()

        # -------------------------------------------------
        # MAIN REQUIRED DATA
        # -------------------------------------------------

        signal = get_value("signal")
        route = get_value("route")
        cascading_signal = get_value("cascading_signal")
        cascading_route = get_value("cascading_route")

        # -------------------------------------------------
        # Ignore completely blank rows
        # -------------------------------------------------

        if not signal and not route and not cascading_signal and not cascading_route:
            continue

        # -------------------------------------------------
        # Required fields must exist
        # -------------------------------------------------

        if not signal:
            log(
                f"TOC ROW {row_number} SKIPPED : "
                "SIGNAL EMPTY"
            )
            continue

        if not route:
            log(
                f"TOC ROW {row_number} SKIPPED : "
                "ROUTE EMPTY"
            )
            continue

        if not cascading_signal:
            log(
                f"TOC ROW {row_number} SKIPPED : "
                "CASCADING SIGNAL EMPTY"
            )
            continue

        if not cascading_route:
            log(
                f"TOC ROW {row_number} SKIPPED : "
                "CASCADING ROUTE EMPTY"
            )
            continue

        # -------------------------------------------------
        # OPTIONAL CASCADING CONTROL BITS
        # -------------------------------------------------

        cascading = get_value("cascading")

        # -------------------------------------------------
        # OPTIONAL EXTRA SIGNAL
        # -------------------------------------------------

        extra_signal = get_value("extra_signal")

        # -------------------------------------------------
        # OPTIONAL EXTRA ROUTE
        # -------------------------------------------------

        extra_route = get_value("extra_route")

        # -------------------------------------------------
        # STORE ROW
        # -------------------------------------------------

        cascading_routes_data.append({

            "signal": signal,

            "route": route,

            "cascading_signal": cascading_signal,

            "cascading_route": cascading_route,

            "cascading": cascading,

            "extra_signal": extra_signal,

            "extra_route": extra_route

        })

        log(
            f"TOC ROW {row_number} LOADED : "
            f"{signal} | "
            f"{route} | "
            f"{cascading_signal} | "
            f"{cascading_route}"
        )

    # =====================================================
    # REFRESH TABLE
    # =====================================================

    refresh_table()

    # =====================================================
    # FINAL RESULT
    # =====================================================

    log(
        f"TOC LOADED SUCCESSFULLY : "
        f"{len(cascading_routes_data)} ROUTE ROW(S)"
    )

# =========================================================
# TABLE
# =========================================================

def refresh_table():

    # Clear existing rows
    tree.delete(*tree.get_children())

    selected_tab = notebook.tab(notebook.select(), "text")

    count = 1

    for row in cascading_routes_data:

        signal = str(row["signal"]).strip().upper()

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

                row["cascading_signal"],

                row["cascading_route"],

                row["cascading"],

                row["extra_signal"],

                row["extra_route"]

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

# =========================================================
# LC CLOSED DETECTION
# =========================================================

# =========================================================
# CHECK LC CLOSE INDICATION
# =========================================================

def is_lc_closed(lc_name):

    if lc_name not in signals:
        log(f"{lc_name} NOT FOUND IN CONFIG")
        return False

    info = signals[lc_name]

    close_point = info.get("out")

    if close_point is None:
        log(f"{lc_name} CLOSE coordinate NOT CONFIGURED")
        return False

    if not bring_test_panel_to_front():
        log("TEST PANEL NOT AVAILABLE")
        return False

    time.sleep(1)

    x, y = close_point

    screenshot = pyautogui.screenshot()

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    log(
        f"{lc_name} CLOSE POINT = ({x},{y})"
    )

    log(
        f"{lc_name} CLOSE RGB = R={r}, G={g}, B={b}"
    )

    # GREEN CLOSE indication
    closed = (
        g > 150 and
        g > r * 1.5 and
        g > b * 1.5
    )

    if closed:

        log(
            f"{lc_name} CLOSE INDICATION = GREEN"
        )

        log(
            f"{lc_name} IS ALREADY CLOSED"
        )

        return True

    log(
        f"{lc_name} CLOSE INDICATION = NOT GREEN"
    )

    log(
        f"{lc_name} IS NOT CLOSED"
    )

    return False

# =========================================================
# FIND LC GATE FROM SIGNAL CONFIG
# =========================================================

def get_lc_gate():

    for name, info in signals.items():

        if info.get("type") == "LC":

            log(f"LC GATE FOUND IN CONFIG : {name}")

            return name

    log("NO LC GATE FOUND IN SIGNAL CONFIG")

    return None

# =========================================================
# LC RECEIVE
# =========================================================

def receive_lc(lc_name):

    if lc_name not in signals:

        log(f"{lc_name} NOT FOUND IN CONFIG")

        return False

    info = signals[lc_name]

    menu_point = info.get("menu")

    if menu_point is None:

        log(
            f"{lc_name} MENU coordinate not configured"
        )

        return False

    # -----------------------------------------
    # OPEN LC MENU
    # -----------------------------------------

    if not click(menu_point):

        log(
            f"FAILED TO CLICK {lc_name} MENU"
        )

        return False

    time.sleep(1)

    # -----------------------------------------
    # CLICK RECEIVE
    # -----------------------------------------

    if not click_menu_item("Receive"):

        log(
            f"{lc_name} RECEIVE MENU ITEM NOT FOUND"
        )

        return False

    log(
        f"{lc_name} RECEIVE CLICKED"
    )

    time.sleep(3)

    # -----------------------------------------
    # VERIFY LC CLOSED
    # -----------------------------------------

    if is_lc_closed(lc_name):

        log(
            f"{lc_name} CLOSED SUCCESSFULLY"
        )

        return True

    log(
        f"{lc_name} DID NOT BECOME CLOSED"
    )

    return False

# =========================================================
# PREPARE LC BEFORE AUTOMATION
# CHECK ONLY ONCE
# =========================================================

# =========================================================
# PREPARE LC BEFORE AUTOMATION
# =========================================================

def prepare_lc_before_automation():

    lc_name = get_lc_gate()

    if not lc_name:

        log("CANNOT START AUTOMATION")
        log("LC GATE NOT FOUND IN SIGNAL CONFIG")

        return False

    log("--------------------------------")
    log("CHECKING LC BEFORE AUTOMATION")
    log(f"LC : {lc_name}")
    log("--------------------------------")

    # =====================================================
    # FIRST CHECK CLOSE INDICATION
    # =====================================================

    log(
        f"CHECKING {lc_name} CLOSE INDICATION..."
    )

    if is_lc_closed(lc_name):

        # -------------------------------------------------
        # ALREADY CLOSED
        # -------------------------------------------------

        log("--------------------------------")
        log(
            f"{lc_name} CLOSE = GREEN"
        )
        log(
            f"{lc_name} ALREADY CLOSED"
        )
        log(
            "SKIPPING LC MENU"
        )
        log(
            "SKIPPING RECEIVE"
        )
        log(
            "LC CONDITION OK"
        )
        log("--------------------------------")

        return True

    # =====================================================
    # NOT CLOSED
    # =====================================================

    log("--------------------------------")
    log(
        f"{lc_name} CLOSE = NOT GREEN"
    )
    log(
        f"{lc_name} IS NOT CLOSED"
    )
    log(
        "OPENING LC MENU"
    )
    log("--------------------------------")

    # =====================================================
    # RECEIVE LC
    # =====================================================

    if not receive_lc(lc_name):

        log(
            f"FAILED TO CLOSE {lc_name}"
        )

        return False

    # =====================================================
    # RECEIVE FUNCTION VERIFIED CLOSE
    # =====================================================

    log("--------------------------------")
    log(
        f"{lc_name} CLOSED SUCCESSFULLY"
    )
    log(
        "LC PREPARATION COMPLETE"
    )
    log("--------------------------------")

    return True

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

    # Close the Control Bit window
    if not find_and_click(
        "Cancel",
        control_type="Button"
    ):
        log("Cancel Button Not Found")
        return False

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

    if signal not in signals:
        return False

    info = signals[signal]

    indicator = info.get("indicator")

    if indicator is None:
        log(f"{signal} SHUNT INDICATOR NOT CONFIGURED")
        return False

    # Your SHUNT indicator checking logic goes here

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
# FIND AND CLICK UI
# =========================================================

def _cascade_rect_ok(rect):
    """True if the rectangle is a real on-screen area on ANY monitor
    (never click a hidden / off-screen element -> taskbar clicks)."""
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


def _find_in_container(container, name, control_type):
    """Find a control by exact text inside ONE window only."""
    for item in container.descendants():
        try:
            if item.window_text().strip().upper() != name.upper():
                continue
            if control_type:
                current_type = str(item.element_info.control_type)
                if control_type.upper() not in current_type.upper():
                    continue
            return item
        except Exception:
            pass
    return None


def find_and_click(name, control_type=None):
    """Click a control by its text.

    1st  : search inside the TEST PANEL window (single screen - the
           Station windows open inside the Test Panel).
    2nd  : if not found there, search the separate 'Station ...' windows
           (dual screen - after CTRL+B / 50051 the Station Bit window can
           open as its own window on the OTHER monitor). That window is
           brought to the front first, so the click selects the item
           instead of only activating the window.
    """

    if not bring_test_panel_to_front():

        log(
            f"TEST PANEL NOT AVAILABLE FOR : {name}"
        )

        return False

    try:

        desktop = Desktop(backend="uia")

        test_panel = None
        station_windows = []

        for win in desktop.windows():

            try:

                title = win.window_text().strip().upper()

                if title == "TEST PANEL":
                    test_panel = win

                elif title.startswith("STATION"):
                    station_windows.append(win)

            except Exception:
                pass

        if test_panel is None:

            log("TEST PANEL WINDOW NOT FOUND")

            return False

        # ---------------------------------------------
        # 1. INSIDE TEST PANEL (single screen)
        # ---------------------------------------------
        item = _find_in_container(test_panel, name, control_type)
        owner = test_panel

        # ---------------------------------------------
        # 2. SEPARATE STATION WINDOW (other screen)
        # ---------------------------------------------
        if item is None:

            for win in station_windows:

                item = _find_in_container(win, name, control_type)

                if item is not None:
                    owner = win
                    break

        if item is None:

            log(f"NOT FOUND IN TEST PANEL / STATION WINDOWS : {name}")

            return False

        # ---------------------------------------------
        # CLICK AT THE CENTRE OF THE CONTROL
        # ---------------------------------------------
        try:
            item.iface_scroll_item.ScrollIntoView()
            time.sleep(0.3)
        except Exception:
            pass

        rect = item.rectangle()

        if not _cascade_rect_ok(rect):

            log(
                f"{name} NOT VISIBLE ON ANY SCREEN "
                f"(L={rect.left} T={rect.top} R={rect.right} B={rect.bottom})"
                f" - CLICK SKIPPED"
            )

            return False

        x = (rect.left + rect.right) // 2
        y = (rect.top + rect.bottom) // 2

        if owner is test_panel:
            bring_test_panel_to_front()
        else:
            try:
                owner.set_focus()
            except Exception:
                pass

        time.sleep(0.3)

        pyautogui.click(x, y)

        where = "TEST PANEL" if owner is test_panel else owner.window_text()

        log(f"CLICKED : {name} at ({x},{y}) IN {where}")

        return True

    except Exception as e:

        log(
            f"find_and_click ERROR : {e}"
        )

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
# =========================================================
# CREATE ADVANCED REPORT
# =========================================================

def create_report():

    wb = Workbook()

    ws = wb.active

    ws.title = "CASCADING AUTOMATION REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:F1")

    title_cell = ws["A1"]

    title_cell.value = "CASCADING AUTOMATION REPORT"

    title_cell.fill = PatternFill(fill_type=None)

    title_cell.font = Font(
        bold=True,
        size=18,
        color="000000"
    )

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 35

    # =====================================================
    # DATE
    # =====================================================

    ws.merge_cells("A2:F2")

    date_cell = ws["A2"]

    date_cell.value = f"Generated : {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"

    date_cell.font = Font(
        bold=True,
        size=11
    )

    date_cell.alignment = Alignment(
        horizontal="center"
    )

    # =====================================================
    # HEADERS
    # =====================================================

    headers = [

        "MAIN SIGNAL",
        "MAIN ROUTE",
        "CASCADING SIGNAL",
        "CASCADING ROUTE",
        "RESULT",
        "DATE & TIME"

    ]

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

    for cell in ws[3]:
        cell.fill = PatternFill(fill_type=None)

        cell.font = Font(
            bold=True,
            color="000000"
        )

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
            row["CASCADING_SIGNAL"],
            row["CASCADING_ROUTE"],
            row["RESULT"],
            row["DATE & TIME"]

        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================



    fail_fill = PatternFill(
        "solid",
        fgColor="FEE2E2"
    )

    for row in ws.iter_rows(min_row=4):

        result_cell = row[4]

        for cell in row:
            cell.fill = PatternFill(fill_type=None)  # default white
            cell.font = Font(color="000000")  # black text
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # ONLY FAIL -> RED
        if str(result_cell.value).strip().upper() == "FAIL":

            for cell in row:
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="FF0000"
                )
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
    # SUMMARY (EXACT FORMAT LIKE SCREENSHOT)
    # =====================================================

    # All routes
    all_routes = set([x["MAIN_ROUTE"] for x in report_rows])

    # PASS routes (initiated)
    initiated_routes = sorted(set([
        x["MAIN_ROUTE"] for x in report_rows
        if x["RESULT"] == "PASS"
    ]))

    # FAIL routes (non-initiated)
    failed_routes = sorted(set([
        x["MAIN_ROUTE"] for x in report_rows
        if x["RESULT"] == "FAIL"
    ]))

    start = ws.max_row + 3

    # TOTAL ROUTES
    ws[f"A{start}"] = "TOTAL ROUTES"
    ws[f"B{start}"] = len(all_routes)

    # INITIATED ROUTES (show values)
    ws[f"A{start + 1}"] = "INITIATED ROUTES"
    ws[f"B{start + 1}"] = ", ".join(initiated_routes)

    # NON-INITIATED ROUTES
    ws[f"A{start + 2}"] = "NON-INITIATED ROUTES"

    if failed_routes:
        ws[f"B{start + 2}"] = ", ".join(failed_routes)
    else:
        ws[f"B{start + 2}"] = "0"

    # Bold labels
    for r in range(start, start + 3):
        ws[f"A{r}"].font = Font(bold=True)

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(REPORT_FILE)

    log(f"REPORT SAVED : {REPORT_FILE}")
# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():

    global running

    # uiautomation/pywinauto are COM-based. Calls made from a thread
    # that hasn't initialized COM (every background thread, by
    # default) can hang indefinitely instead of timing out - this is
    # almost certainly why automation was getting stuck silently after
    # "Clicked coordinate (...)" with no further log output. The main
    # thread gets this for free from Tkinter/pywin32 import machinery;
    # background threads need it explicitly.
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        log(f"COM INIT WARNING : {e}")

    try:

        _run_engine_impl()

    except Exception as e:

        import traceback

        log("=" * 60)
        log("FATAL ERROR IN run_engine() - AUTOMATION STOPPED")
        log(f"{type(e).__name__}: {e}")

        for line in traceback.format_exc().splitlines():
            log(line)

        log("=" * 60)

        running = False

    finally:

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _run_engine_impl():

    global running

    # -----------------------------------------
    # HIDE AUTOMATION DASHBOARD
    # -----------------------------------------

    root.iconify()

    time.sleep(2)

    # -----------------------------------------
    # MAKE TEST PANEL ACTIVE
    # -----------------------------------------

    if not bring_test_panel_to_front():

        log("ERROR: TEST PANEL NOT FOUND")

        running = False

        root.deiconify()

        status_label.config(
            text="STOPPED",
            fg="red"
        )

        return

    log(
        "TEST PANEL READY - AUTOMATION STARTING"
    )

    time.sleep(1)

    # LC check removed - cascading starts directly
    log("STARTING CASCADING AUTOMATION")

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

    create_report()

    running = False

    status_label.config(
        text="STOPPED",
        fg="red"
    )

    log("AUTOMATION COMPLETED")


# =========================================================
# START
# =========================================================

def start_automation():

    global running

    # =========================================================
    # CONCURRENCY GUARD
    # This program auto-starts itself via auto_start_from_launcher()
    # as soon as EDRC_TOC/EDRC_COORDS are detected, AND the suite's
    # own driver separately clicks this same function. Without this
    # guard, both can fire and run TWO automation threads at once,
    # both clicking the same UI at conflicting moments - which looks
    # exactly like getting stuck right after a click with nothing
    # further happening.
    # =========================================================

    if running:

        log("START AUTOMATION IGNORED : ALREADY RUNNING")

        return

    if not signals:

        log("LOAD CONFIG FIRST")

        return

    if not cascading_routes_data:

        log("LOAD LOCK ROUTES FIRST")

        return

    # =========================================================
    # PRE-FLIGHT CHECK
    # If none of the TOC's Signal values match a loaded coordinate,
    # every route will fail immediately and the run will look like
    # it finished instantly with nothing actually tested. Catch that
    # here, loudly, before starting.
    # =========================================================

    toc_signals = {
        str(row.get("signal", "")).strip().upper()
        for row in cascading_routes_data
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


def _log_tk_callback_exception(exc_type, exc_value, exc_tb):
    """Catches exceptions from ANY Tkinter callback (button clicks,
    etc.) that would otherwise just print to a console the user may
    never see, and writes them to the persistent log file/GUI log
    instead."""

    import traceback

    tb_text = "".join(
        traceback.format_exception(exc_type, exc_value, exc_tb)
    )

    if _log_file_handle is not None:
        try:
            _log_file_handle.write("=" * 60 + "\n")
            _log_file_handle.write("UNCAUGHT EXCEPTION IN TKINTER CALLBACK\n")
            _log_file_handle.write(tb_text)
            _log_file_handle.write("=" * 60 + "\n")
            _log_file_handle.flush()
        except Exception:
            pass

    try:
        log("=" * 60)
        log("UNCAUGHT EXCEPTION IN TKINTER CALLBACK")
        for line in tb_text.splitlines():
            log(line)
        log("=" * 60)
    except Exception:
        pass


root.report_callback_exception = _log_tk_callback_exception

root.title(
    "LOCK ROUTES TESTING AUTOMATION"
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

    text="CASCADING AUTOMATION SYSTEM",

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
    load_cascading_routes,
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

notebook.bind("<<NotebookTabChanged>>", lambda e: refresh_table())

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

        "CASCADING_SIGNAL",

        "CASCADING_ROUTE",

        "CASCADING",

        "EXTRA SIGNAL",

         "EXTRA ROUTE"

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

tree.heading("CASCADING_SIGNAL", text="CASCADING SIGNAL")

tree.heading("CASCADING_ROUTE", text="CASCADING ROUTE")

tree.heading("CASCADING", text="CASCADING")

tree.heading("EXTRA SIGNAL", text="EXTRA SIGNAL")

tree.heading("EXTRA ROUTE", text="EXTRA ROUTE")

tree.column("NO", width=70, anchor="center")

tree.column("SIGNAL", width=100, anchor="center")

tree.column("ROUTE", width=100, anchor="center")

tree.column("CASCADING_SIGNAL", width=140, anchor="center")

tree.column("CASCADING_ROUTE", width=150, anchor="center")

tree.column("CASCADING", width=220, anchor="center")

tree.column("EXTRA SIGNAL", width=120, anchor="center")

tree.column("EXTRA ROUTE", width=120, anchor="center")

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
# WALK-AWAY AUTO START
# =========================================================

def auto_start_from_launcher():

    log("================================")
    log("WALK-AWAY AUTO START")
    log("================================")

    toc_file = (
        os.environ.get("EDRC_TOC")
        or os.environ.get("EDRC_LIST")
    )

    coords_file = os.environ.get("EDRC_COORDS")

    log(f"EDRC_TOC    = {toc_file}")
    log(f"EDRC_COORDS = {coords_file}")

    # -----------------------------------------------------
    # LOAD TOC FIRST
    # -----------------------------------------------------

    if not toc_file:
        log("ERROR: EDRC_TOC / EDRC_LIST NOT SET")
        return

    if not os.path.exists(toc_file):
        log(f"ERROR: TOC FILE NOT FOUND: {toc_file}")
        return

    log(f"LOADING TOC: {toc_file}")

    load_cascading_routes()

    if not cascading_routes_data:
        log("ERROR: TOC LOADED BUT NO ROUTES FOUND")
        return

    log(
        f"TOC LOADED SUCCESSFULLY: "
        f"{len(cascading_routes_data)} ROUTE ROW(S)"
    )

    # -----------------------------------------------------
    # LOAD SIGNAL CONFIG
    # -----------------------------------------------------

    if not coords_file:
        log("ERROR: EDRC_COORDS NOT SET")
        return

    if not os.path.exists(coords_file):
        log(f"ERROR: CONFIG FILE NOT FOUND: {coords_file}")
        return

    log(f"LOADING CONFIG: {coords_file}")

    load_config()

    if not signals:
        log("ERROR: SIGNAL CONFIG LOADED BUT SIGNALS ARE EMPTY")
        return

    log(
        f"SIGNAL CONFIG LOADED: "
        f"{len(signals)} ELEMENT(S)"
    )

    # -----------------------------------------------------
    # START AUTOMATION
    # -----------------------------------------------------

    log("================================")
    log("STARTING CASCADING AUTOMATION")
    log("================================")

    start_automation()

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
# WALK-AWAY START
# =========================================================

root.after(
    1000,
    auto_start_from_launcher
)

# =========================================================
# RUN
# =========================================================

root.mainloop()