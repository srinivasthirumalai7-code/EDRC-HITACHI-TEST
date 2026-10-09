
from PIL import Image, ImageTk
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
from PIL import ImageTk
from PIL import Image as PILImage
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
import re

import pyautogui

import win32api
import win32con


import uiautomation as auto

from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = ""

# =========================================================
# GLOBALS
# =========================================================

signals = {}

lock_routes_data = []

track_points = {}

running = False

capture_queue = []

capture_master_list = []

capture_index = 0

record_stage = "coordinate"

capture_module = None

current_aspects = None

capture_overlay = None

capture_mode = None

current_signal = None

current_step = 0

last_capture_time = 0

capture_waiting = False

captured_point = None

undo_stack = []

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

report_saved = False
current_test = ""
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
    capture_overlay.bind("<BackSpace>", lambda e: "break")

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
    """Build the TL-style coordinate capture queue."""
    capture_list = []

    for signal, info in signals.items():
        if info.get("type") == "MAIN":
            capture_list.append(("MAIN", signal))

    for signal, info in signals.items():
        if info.get("type") == "CALLING_ON":
            capture_list.append(("CALLING_ON", signal))

    for signal, info in signals.items():
        if info.get("type") == "SHUNT":
            capture_list.append(("SHUNT", signal))

    for signal, info in signals.items():
        if info.get("type") == "POINT":
            capture_list.append(("POINT", signal))

    for signal, info in signals.items():
        if info.get("type") == "CH":
            capture_list.append(("CH", signal))

    for signal, info in signals.items():
        if info.get("type") == "LC":
            capture_list.append(("LC", signal))

    return capture_list

def set_initial_stage_for_current():
    global record_stage

    sig_type, signal = capture_master_list[capture_index]

    if sig_type == "MAIN":
        record_stage = "coordinate"
    else:
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
                "YELLOW": None
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
        log("ALL COORDINATES CAPTURED")
        capture_module = None
        destroy_capture_overlay()
        root.deiconify()

        messagebox.showinfo(
            "Completed",
            "Signal / element coordinate configuration saved successfully."
        )
        return

    sig_type, signal = capture_master_list[capture_index]

    if sig_type == "MAIN":

        if record_stage == "coordinate":
            update_capture_overlay(
                "MAIN", signal, "CLICK : SIGNAL MENU",
                "Move the mouse onto the signal menu.\n\nPress SPACE."
            )
        elif record_stage == "ask_aspects":
            show_aspect_buttons(signal)
        elif record_stage == "RED":
            update_capture_overlay(
                "MAIN", signal, "CLICK : RED ASPECT",
                "Move the mouse onto the RED lamp.\n\nPress SPACE."
            )
        elif record_stage == "YELLOW":
            update_capture_overlay(
                "MAIN", signal, "CLICK : YELLOW ASPECT",
                "Move the mouse onto the YELLOW lamp.\n\nPress SPACE."
            )
        elif record_stage == "DOUBLE_YELLOW":
            update_capture_overlay(
                "MAIN", signal, "CLICK : DOUBLE YELLOW",
                "Move the mouse onto the DOUBLE YELLOW lamp.\n\nPress SPACE."
            )
        elif record_stage == "GREEN":
            update_capture_overlay(
                "MAIN", signal, "CLICK : GREEN ASPECT",
                "Move the mouse onto the GREEN lamp.\n\nPress SPACE."
            )
        elif record_stage == "ROUTE_INDICATOR":
            update_capture_overlay(
                "MAIN", signal, "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator.\n\nPress SPACE."
            )

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":
            update_capture_overlay(
                "CALLING_ON", signal, "CLICK : MENU",
                "Move the mouse onto the signal menu.\n\nPress SPACE."
            )
        elif record_stage == "YELLOW":
            update_capture_overlay(
                "CALLING_ON", signal, "CLICK : YELLOW",
                "Move the mouse onto the YELLOW lamp.\n\nPress SPACE."
            )
        elif record_stage == "ROUTE_INIT":
            update_capture_overlay(
                "CALLING_ON", signal, "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator.\n\nPress SPACE."
            )

    elif sig_type == "SHUNT":

        if record_stage == "menu":
            update_capture_overlay(
                "SHUNT", signal, "CLICK : MENU",
                "Move the mouse onto the menu.\n\nPress SPACE."
            )
        elif record_stage == "state_indicator":
            update_capture_overlay(
                "SHUNT", signal, "CLICK : ASPECT INDICATOR",
                "Move the mouse onto the SHUNT aspect/state indicator.\n\nPress SPACE."
            )
        elif record_stage == "route_init":
            update_capture_overlay(
                "SHUNT", signal, "CLICK : ROUTE INITIATION",
                "Move the mouse onto the route initiation indicator.\n\nPress SPACE."
            )

    elif sig_type == "POINT":

        if record_stage == "menu":
            update_capture_overlay(
                "POINT", signal, "CLICK : MENU",
                "Move the mouse onto the point menu.\n\nPress SPACE."
            )
        elif record_stage == "normal":
            update_capture_overlay(
                "POINT", signal, "CLICK : NORMAL",
                "Move the mouse onto the NORMAL indication.\n\nPress SPACE."
            )
        elif record_stage == "reverse":
            update_capture_overlay(
                "POINT", signal, "CLICK : REVERSE",
                "Move the mouse onto the REVERSE indication.\n\nPress SPACE."
            )
        elif record_stage == "free":
            update_capture_overlay(
                "POINT", signal, "CLICK : FREE",
                "Move the mouse onto the FREE indication.\n\nPress SPACE."
            )

    elif sig_type == "CH":

        if record_stage == "menu":
            update_capture_overlay(
                "CH", signal, "CLICK : MENU",
                "Move the mouse onto the crank handle menu.\n\nPress SPACE."
            )
        elif record_stage == "IN":
            update_capture_overlay(
                "CH", signal, "CLICK : IN",
                "Move the mouse onto the IN indication.\n\nPress SPACE."
            )
        elif record_stage == "OUT":
            update_capture_overlay(
                "CH", signal, "CLICK : OUT",
                "Move the mouse onto the OUT indication.\n\nPress SPACE."
            )
        elif record_stage == "ECH":
            update_capture_overlay(
                "CH", signal, "CLICK : ECH",
                "Move the mouse onto the ECH indication.\n\nPress SPACE."
            )
        elif record_stage == "FREE":
            update_capture_overlay(
                "CH", signal, "CLICK : FREE",
                "Move the mouse onto the FREE indication.\n\nPress SPACE."
            )

    elif sig_type == "LC":

        if record_stage == "menu":
            update_capture_overlay(
                "LC", signal, "CLICK : MENU",
                "Move the mouse onto the LC Gate menu.\n\nPress SPACE."
            )
        elif record_stage == "IN":
            update_capture_overlay(
                "LC", signal, "CLICK : IN",
                "Move the mouse onto the IN indication.\n\nPress SPACE."
            )
        elif record_stage == "OUT":
            update_capture_overlay(
                "LC", signal, "CLICK : OUT",
                "Move the mouse onto the OUT indication.\n\nPress SPACE."
            )

def master_save_point(x, y):
    global record_stage

    if capture_index >= len(capture_master_list):
        return

    sig_type, signal = capture_master_list[capture_index]

    if sig_type == "MAIN":

        if record_stage == "coordinate":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "ask_aspects"
            master_next_capture()
            return

        if record_stage == "ask_aspects":
            return

        if record_stage == "RED":
            push_undo(lambda s=signal: signals[s].__setitem__("RED", None))
            signals[signal]["RED"] = [x, y]
            log(f"{signal} RED Saved")
            record_stage = "GREEN" if current_aspects == 2 else "YELLOW"
            master_next_capture()
            return

        if record_stage == "YELLOW":
            push_undo(lambda s=signal: signals[s].__setitem__("YELLOW", None))
            signals[signal]["YELLOW"] = [x, y]
            log(f"{signal} YELLOW Saved")
            record_stage = "DOUBLE_YELLOW" if current_aspects == 4 else "GREEN"
            master_next_capture()
            return

        if record_stage == "DOUBLE_YELLOW":
            push_undo(lambda s=signal: signals[s].__setitem__("DOUBLE_YELLOW", None))
            signals[signal]["DOUBLE_YELLOW"] = [x, y]
            log(f"{signal} DOUBLE YELLOW Saved")
            record_stage = "GREEN"
            master_next_capture()
            return

        if record_stage == "GREEN":
            push_undo(lambda s=signal: signals[s].__setitem__("GREEN", None))
            signals[signal]["GREEN"] = [x, y]
            log(f"{signal} GREEN Saved")
            record_stage = "ROUTE_INDICATOR"
            master_next_capture()
            return

        if record_stage == "ROUTE_INDICATOR":
            push_undo(lambda s=signal: signals[s].__setitem__("ROUTE_INDICATOR", None))
            signals[signal]["ROUTE_INDICATOR"] = [x, y]
            log(f"{signal} ROUTE INITIATION Saved")
            advance_to_next_signal()
            return

    elif sig_type == "CALLING_ON":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
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

    elif sig_type == "SHUNT":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "state_indicator"
            master_next_capture()
            return

        if record_stage == "state_indicator":
            push_undo(lambda s=signal: signals[s].__setitem__("state_indicator", None))
            signals[signal]["state_indicator"] = [x, y]
            # Keep old fields for existing Button Block logic/config compatibility.
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

    elif sig_type == "POINT":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "normal"
            master_next_capture()
            return

        if record_stage == "normal":
            push_undo(lambda s=signal: signals[s].__setitem__("normal", None))
            signals[signal]["normal"] = [x, y]
            log(f"{signal} NORMAL Saved")
            record_stage = "reverse"
            master_next_capture()
            return

        if record_stage == "reverse":
            push_undo(lambda s=signal: signals[s].__setitem__("reverse", None))
            signals[signal]["reverse"] = [x, y]
            log(f"{signal} REVERSE Saved")
            record_stage = "free"
            master_next_capture()
            return

        if record_stage == "free":
            push_undo(lambda s=signal: signals[s].__setitem__("free", None))
            signals[signal]["free"] = [x, y]
            log(f"{signal} FREE Saved")
            advance_to_next_signal()
            return

    elif sig_type == "CH":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "IN"
            master_next_capture()
            return

        if record_stage == "IN":
            push_undo(lambda s=signal: signals[s].__setitem__("IN", None))
            signals[signal]["IN"] = [x, y]
            log(f"{signal} IN Saved")
            record_stage = "OUT"
            master_next_capture()
            return

        if record_stage == "OUT":
            push_undo(lambda s=signal: signals[s].__setitem__("OUT", None))
            signals[signal]["OUT"] = [x, y]
            log(f"{signal} OUT Saved")
            record_stage = "ECH"
            master_next_capture()
            return

        if record_stage == "ECH":
            push_undo(lambda s=signal: signals[s].__setitem__("ECH", None))
            signals[signal]["ECH"] = [x, y]
            log(f"{signal} ECH Saved")
            record_stage = "FREE"
            master_next_capture()
            return

        if record_stage == "FREE":
            push_undo(lambda s=signal: signals[s].__setitem__("FREE", None))
            signals[signal]["FREE"] = [x, y]
            log(f"{signal} FREE Saved")
            advance_to_next_signal()
            return

    elif sig_type == "LC":

        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "IN"
            master_next_capture()
            return

        if record_stage == "IN":
            push_undo(lambda s=signal: signals[s].__setitem__("IN", None))
            signals[signal]["IN"] = [x, y]
            log(f"{signal} IN Saved")
            record_stage = "OUT"
            master_next_capture()
            return

        if record_stage == "OUT":
            push_undo(lambda s=signal: signals[s].__setitem__("OUT", None))
            signals[signal]["OUT"] = [x, y]
            log(f"{signal} OUT Saved")
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
    """Create the TL-style capture list while retaining existing data keys."""

    global signals

    signals = {}
    combined = []
    existing = set()

    count = prompt_count("MAIN Signals")
    for i in range(count):
        name = prompt_name("MAIN Signal", i, count, existing)
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

    count = prompt_count("CALLING-ON Signals")
    for i in range(count):
        name = prompt_name("CALLING-ON Signal", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "CALLING_ON",
            "menu": None,
            "YELLOW": None,
            "ROUTE_INIT": None
        }
        combined.append(("CALLING_ON", name))

    count = prompt_count("SHUNT Signals")
    for i in range(count):
        name = prompt_name("SHUNT Signal", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "SHUNT",
            "menu": None,
            "state_indicator": None,
            "route_init": None,
            "C1": None,
            "C2": None,
            "C3": None
        }
        combined.append(("SHUNT", name))

    count = prompt_count("POINTS")
    for i in range(count):
        name = prompt_name("POINT", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "POINT",
            "menu": None,
            "normal": None,
            "reverse": None,
            "free": None
        }
        combined.append(("POINT", name))

    count = prompt_count("CRANK HANDLES")
    for i in range(count):
        name = prompt_name("CRANK HANDLE", i, count, existing)
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

    count = prompt_count("LC GATES")
    for i in range(count):
        name = prompt_name("LC GATE", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "LC",
            "menu": None,
            "IN": None,
            "OUT": None
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
# CREATE SIGNAL SETUP
# =========================================================




# =========================================================
# SAVE POINT
# =========================================================

def save_point(x, y):
    """Compatibility wrapper for older capture callers."""
    if capture_module == "MASTER":
        master_save_point(x, y)

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

        "Indicator_X", "Indicator_Y", "Route_Init_X", "Route_Init_Y"

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

                (info.get("state_indicator") or info.get("C1") or [None, None])[0],
                (info.get("state_indicator") or info.get("C1") or [None, None])[1],
                (info.get("route_init") or info.get("C3") or info.get("C2") or [None, None])[0],
                (info.get("route_init") or info.get("C3") or info.get("C2") or [None, None])[1]

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

            "YELLOW": [row[3], row[4]]
            if row[3] is not None else None,

            "ROUTE_INIT": [row[5], row[6]]
            if len(row) > 6 and row[5] is not None else None

        }

    # =====================================================
    # SHUNT
    # =====================================================

    ws = wb["SHUNT"]

    # Read SHUNT columns by header name. Supports the new
    # Indicator/Route_Init format and old C1/C2/C3 files.
    headers = {}
    for col_index, cell in enumerate(ws[1], start=1):
        if cell.value is not None:
            headers[str(cell.value).strip().upper()] = col_index

    def shunt_cell(row_values, header_name):
        col = headers.get(header_name.upper())
        if col is None or col > len(row_values):
            return None
        return row_values[col - 1]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row or not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        menu = [
            shunt_cell(row, "MENU_X"),
            shunt_cell(row, "MENU_Y")
        ]

        indicator = [
            shunt_cell(row, "INDICATOR_X"),
            shunt_cell(row, "INDICATOR_Y")
        ]

        # =====================================================
        # SHUNT ROUTE INITIATOR
        # =====================================================

        route_init_x = shunt_cell(row, "ROUTE_INIT_X")

        if route_init_x is None:
            route_init_x = shunt_cell(row, "ROUTEINIT_X")

        route_init_y = shunt_cell(row, "ROUTE_INIT_Y")

        if route_init_y is None:
            route_init_y = shunt_cell(row, "ROUTEINIT_Y")

        route_init = [
            route_init_x,
            route_init_y
        ]

        signals[signal_name] = {
            "type": "SHUNT",
            "menu": menu,
            "state_indicator": indicator,
            "route_init": route_init,

            # Legacy aliases retained for compatibility.
            "C1": indicator,
            "C2": route_init,
            "C3": route_init
        }

    log("CONFIG LOADED")


def load_universal_coordinates():
    """Loads coordinates from the multi-sheet Universal Yard Coordinate
    file (the same file every other suite program reads from), instead
    of this program's own native config format (MAIN_SIGNALS /
    CALLING_ON / SHUNT). Populates the same `signals` dict that
    load_config() populates, using the universal file's actual
    sheet/column names (MAIN / CAL / SHUNT). This program doesn't use
    POINT/CH/LC, so those sheets are ignored even if present."""

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

    loaded_main = loaded_cal = loaded_shunt = 0

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
                    "GREEN": parse_point(cell(row, hi.get("GREEN_X")), cell(row, hi.get("GREEN_Y")))
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
                    "ROUTE_INIT": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
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
                indicator = parse_point(cell(row, hi.get("INDICATOR_X")), cell(row, hi.get("INDICATOR_Y")))
                route_init = parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "state_indicator": indicator,
                    "route_init": route_init,
                    "C1": indicator,
                    "C2": route_init,
                    "C3": route_init
                }
                loaded_shunt += 1

    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_cal} CALLING-ON, {loaded_shunt} SHUNT"
    )

    # -----------------------------------------------------------
    # DIAGNOSTIC: show whether the signals just loaded actually
    # overlap with the signals already loaded from the TOC. If
    # there's zero overlap, every route is about to fail with
    # "NOT FOUND IN SIGNAL CONFIG" and the run will look like it
    # finished instantly with a full report of failures.
    # -----------------------------------------------------------

    toc_signals = {row["signal"] for row in lock_routes_data}
    matched = toc_signals & set(signals.keys())

    log(f"TOC SIGNALS : {sorted(toc_signals)}")
    log(f"MATCHED IN COORDINATES : {sorted(matched)}")

    if toc_signals and not matched:
        log(
            "WARNING: NONE of the TOC signals match the loaded "
            "coordinates - check you selected the right universal "
            "coordinate file for this yard."
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
# LOAD LOCK ROUTE EXCEL
# =========================================================

def _header_key(value):
    """
    Normalize Excel headers so all of these are treated as the same:
        APPROACH_TRACK
        Approach_Track
        Approach Track
        APPROACH-TRACK
    """
    text = str(value or "").strip().upper()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return text.strip("_")


def load_lock_routes():

    global lock_routes_data

    file_path = filedialog.askopenfilename(
        title="Select TOC Excel file",
        filetypes=[("Excel", "*.xlsx")]
    )

    log(f"TOC FILE : {file_path}")

    if not file_path:
        return

    wb = load_workbook(file_path, data_only=True)

    # Prefer TOC, but if the workbook has only one sheet, use it.
    if "TOC" in wb.sheetnames:
        ws = wb["TOC"]
    elif len(wb.sheetnames) == 1:
        ws = wb[wb.sheetnames[0]]
        log(
            f"TOC SHEET NOT NAMED 'TOC' - USING SHEET : {ws.title}"
        )
    else:
        messagebox.showerror(
            "ERROR",
            "TOC sheet not found in Excel file."
        )
        return

    log(f"READING WORKSHEET : {ws.title}")

    columns = {
        _header_key(cell.value): index
        for index, cell in enumerate(ws[1])
        if cell.value is not None
    }

    log(
        "TOC HEADERS : "
        + ", ".join(columns.keys())
    )

    required = ["SIGNAL", "ROUTE"]
    missing = [
        name for name in required
        if name not in columns
    ]

    if missing:
        messagebox.showerror(
            "TOC FORMAT ERROR",
            "Missing required column(s): " + ", ".join(missing)
        )
        return

    # Track-drop input is taken from the TOC "Track" column.
    # (Headers are normalized to UPPERCASE by _header_key, so the
    # candidate must be "TRACK".) Approach_Track kept only as fallback.
    approach_column = None

    for candidate in (
        "TRACK",
        "APPROACH_TRACK",
        "APPROACHTRACK",
        "APPROACH",
    ):
        if candidate in columns:
            approach_column = columns[candidate]
            break

    if approach_column is not None:
        log(
            f"TRACK COLUMN FOUND : "
            f"{ws.cell(1, approach_column + 1).value}"
        )
    else:
        log(
            "APPROACH_TRACK COLUMN NOT FOUND"
        )
        log(
            "CALLING-ON routes will be reported as "
            "APPROACH TRACK NOT CONFIGURED."
        )

    grouped_routes = {}

    for excel_row_number, values in enumerate(
        ws.iter_rows(
            min_row=2,
            values_only=True
        ),
        start=2
    ):

        if not any(
            value is not None and str(value).strip()
            for value in values
        ):
            continue

        def value_for(header_name):
            column_index = columns.get(
                _header_key(header_name)
            )

            if (
                column_index is None
                or column_index >= len(values)
            ):
                return None

            return values[column_index]

        signal = str(
            value_for("SIGNAL") or ""
        ).replace(".0", "").strip().upper()

        route = str(
            value_for("ROUTE") or ""
        ).replace(".0", "").strip().upper()

        approach_track = ""

        if approach_column is not None:
            raw_approach = values[approach_column]

            if raw_approach is not None:
                approach_track = (
                    str(raw_approach)
                    .replace(".0", "")
                    .strip()
                    .upper()
                )

        if not signal or not route:
            continue

        if signal not in grouped_routes:
            grouped_routes[signal] = {}

        route_list = [
            r.strip().upper()
            for r in route.split(",")
            if r.strip()
        ]

        # Normally there is one route and one approach track per row.
        # If several comma-separated routes/tracks are supplied,
        # pair them by position.
        approach_list = [
            a.strip().upper()
            for a in approach_track.split(",")
            if a.strip()
        ]

        for index, route_name in enumerate(route_list):

            if not approach_list:
                route_track = ""

            elif len(approach_list) == 1:
                route_track = approach_list[0]

            elif index < len(approach_list):
                route_track = approach_list[index]

            else:
                route_track = approach_list[-1]

            grouped_routes[signal][route_name] = route_track

            # Explicit Calling-On diagnostic.
            if (
                (
                    signal.endswith("C")
                    and signal[:-1].isdigit()
                )
                or signal in ("1C", "2C", "3C", "30C", "31C", "32C")
            ):
                log(
                    f"TOC ROW {excel_row_number}: "
                    f"CALLING-ON {signal} -> {route_name} "
                    f"-> APPROACH_TRACK = "
                    f"{route_track or '[EMPTY]'}"
                )

    lock_routes_data = []

    for signal, route_map in grouped_routes.items():

        route_entries = [
            {
                "route": route_name,
                "approach_track": approach_track
            }
            for route_name, approach_track
            in route_map.items()
        ]

        routes = [
            entry["route"]
            for entry in route_entries
        ]

        lock_routes_data.append({
            "signal": signal,
            "routes": routes,
            "route_entries": route_entries
        })

        log(
            f"Loaded -> SIGNAL = {signal}"
        )
        log(
            f"Routes -> {routes}"
        )

        for entry in route_entries:
            log(
                f"{signal} {entry['route']} -> "
                f"APPROACH_TRACK = "
                f"{entry['approach_track'] or '[EMPTY]'}"
            )

    refresh_table()

    log(
        f"TOC LOADED: "
        f"{len(lock_routes_data)} SIGNALS"
    )

# =========================================================
# TABLE
# =========================================================

def get_button_block_signal_type(signal_name):
    """
    Use the existing signal configuration to classify the signal.
    No automation logic is changed.
    """

    name = str(signal_name or "").strip().upper()

    # ---------------------------------------------------------
    # SHUNT SIGNAL NAMING RULE - CHECK THIS FIRST
    # ---------------------------------------------------------
    # The TOC may already classify a signal as MAIN, but for
    # Button Block display, names such as:
    #
    #       9SH
    #       17SH
    #       SH9
    #       SH17
    #
    # must be displayed in the SHUNT SIGNALS tab.
    #
    # This check MUST happen BEFORE signals.get(name), otherwise
    # an existing MAIN entry in the configuration will override
    # the SHUNT naming rule.
    # ---------------------------------------------------------
    if (
        (name.endswith("SH") and name[:-2].isdigit())
        or
        (name.startswith("SH") and name[2:].isdigit())
    ):
        return "SHUNT"


    # ---------------------------------------------------------
    # Existing configured signal type
    # ---------------------------------------------------------

    info = signals.get(name)

    if info:
        signal_type = str(
            info.get("type", "MAIN")
        ).upper()

        if signal_type == "CALLING_ON":
            return "CALLING_ON"

        if signal_type == "SHUNT":
            return "SHUNT"

        return "MAIN"

    # Existing project convention/fallback.
    if name.endswith("C"):
        return "CALLING_ON"

    return "MAIN"


def refresh_signal_tree(tree_widget, signal_type):

    tree_widget.delete(
        *tree_widget.get_children()
    )

    display_no = 1

    for row in lock_routes_data:

        signal_name = str(
            row.get("signal", "")
        ).strip()

        row_type = get_button_block_signal_type(
            signal_name
        )

        if row_type != signal_type:
            continue

        tree_widget.insert(
            "",
            "end",
            values=(
                display_no,
                signal_name,
                ", ".join(row.get("routes", []))
            )
        )

        display_no += 1


def refresh_table():
    """
    Refresh MAIN, CALLING-ON and SHUNT tables.
    """

    refresh_signal_tree(
        main_tree,
        "MAIN"
    )

    refresh_signal_tree(
        calling_on_tree,
        "CALLING_ON"
    )

    refresh_signal_tree(
        shunt_tree,
        "SHUNT"
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
            item.Click()

            return True

    except:
        pass

    return False

def is_signal_already_blocked(signal_name):

    info = signals[signal_name]

    click(info["menu"])
    time.sleep(1)

    if not click_menu_item("Critical"):

        pyautogui.press("esc")

        return False

    time.sleep(1)

    try:

        unblock_item = auto.MenuItemControl(
            searchDepth=15,
            Name="Button Unblock"
        )

        # Button Unblock is enabled only when the signal is already blocked
        if unblock_item.Exists(2) and unblock_item.IsEnabled:

            log(f"{signal_name} IS ALREADY BLOCKED")

            pyautogui.press("esc")

            return True

    except:

        pass

    pyautogui.press("esc")

    return False


def get_route_approach_track(toc_row, route_name):
    """
    Return the Approach Track belonging to this exact TOC route.

    This is deliberately route-specific because the same Calling-On
    signal can have multiple routes.
    """

    route_name = str(
        route_name or ""
    ).strip().upper()

    for entry in (
        toc_row.get("route_entries") or []
    ):
        entry_route = str(
            entry.get("route", "")
        ).strip().upper()

        if entry_route == route_name:
            return str(
                entry.get("approach_track", "") or ""
            ).strip().upper()

    return ""



# =========================================================
# BUTTON BLOCK
# =========================================================

def button_block(signal_name):

    info = signals[signal_name]

    log(f"{signal_name} -> CHECKING BUTTON BLOCK")

    # Open the signal menu only once
    click(info["menu"])
    time.sleep(1)

    if not click_menu_item("Critical"):

        log("CRITICAL MENU NOT FOUND")

        return "ERROR"

    time.sleep(1)

    # If Button Unblock is enabled, this signal was already blocked
    try:

        unblock_item = auto.MenuItemControl(
            searchDepth=15,
            Name="Button Unblock"
        )

        if unblock_item.Exists(2) and unblock_item.IsEnabled:

            log(f"{signal_name} IS ALREADY BLOCKED")

            pyautogui.press("esc")

            return "ALREADY_BLOCKED"

    except:

        pass

    # Signal is not blocked: apply Button Block
    if not click_menu_item("Button Block"):

        log("BUTTON BLOCK OPTION NOT FOUND")

        return "ERROR"

    time.sleep(1)

    if not click_yes_dialog():

        log("BUTTON BLOCK YES CONFIRMATION NOT FOUND")

        return "ERROR"

    time.sleep(2)

    log(f"{signal_name} BUTTON BLOCK DONE")

    return "BLOCKED"

# =========================================================
# BUTTON UNBLOCK
# =========================================================

def button_unblock(signal_name):

    info = signals[signal_name]

    log(f"{signal_name} -> BUTTON UNBLOCK")

    click(info["menu"])
    time.sleep(1)

    if not click_menu_item("Critical"):
        log("CRITICAL MENU NOT FOUND")
        return False

    time.sleep(1)

    if not click_menu_item("Button Unblock"):
        log("BUTTON UNBLOCK OPTION NOT FOUND")
        return False

    time.sleep(1)

    if not click_yes_dialog():
        log("BUTTON UNBLOCK YES CONFIRMATION NOT FOUND")
        return False

    time.sleep(2)

    log(f"{signal_name} BUTTON UNBLOCK DONE")

    return True

# =========================================================
# CANCEL + RELEASE ROUTE
# =========================================================

def cancel_and_release_route(signal_name):

    if signal_name not in signals:
        log(f"{signal_name} NOT FOUND IN SIGNAL CONFIG")
        return False

    info = signals[signal_name]

    # =====================================================
    # SIGNAL CANCEL
    # =====================================================

    log(f"{signal_name} -> SIGNAL CANCEL")

    click(info["menu"])
    time.sleep(1)

    if not click_menu_item("Signal Cancel"):

        log(
            f"{signal_name} -> "
            f"SIGNAL CANCEL NOT FOUND"
        )

        pyautogui.press("esc")
        time.sleep(1)

        # Retry once
        click(info["menu"])
        time.sleep(1)

        if not click_menu_item("Signal Cancel"):

            log(
                f"{signal_name} -> "
                f"SIGNAL CANCEL FAILED"
            )

            pyautogui.press("esc")
            return False

    log(
        f"{signal_name} -> "
        f"SIGNAL CANCEL DONE"
    )

    time.sleep(3)

    # =====================================================
    # ROUTE RELEASE
    # =====================================================

    log(
        f"{signal_name} -> ROUTE RELEASE"
    )

    click(info["menu"])
    time.sleep(1)

    if not click_menu_item("Route Release"):

        log(
            f"{signal_name} -> "
            f"ROUTE RELEASE NOT FOUND"
        )

        pyautogui.press("esc")
        time.sleep(1)

        # Retry once
        click(info["menu"])
        time.sleep(1)

        if not click_menu_item("Route Release"):

            log(
                f"{signal_name} -> "
                f"ROUTE RELEASE FAILED"
            )

            pyautogui.press("esc")
            return False

    log(
        f"{signal_name} -> "
        f"ROUTE RELEASE COMMAND SENT"
    )

    # Give the interlocking time to completely release
    time.sleep(20)

    log(
        f"{signal_name} -> "
        f"ROUTE RELEASE COMPLETED"
    )

    return True


def calling_on_cleanup_and_restore(
    signal_name,
    approach_track,
    status,
    result
):
    """
    Execute the complete Calling-On post-test sequence:

        Signal Cancel
        Route Release
        50051 restore/transmit

    Returns updated (status, result).
    """

    cleanup_ok = calling_on_cancel_and_release(
        signal_name
    )

    if not cleanup_ok:

        status = (
            f"{status} - "
            "SIGNAL CANCEL/ROUTE RELEASE FAILED"
        )

        result = "FAIL"

    restore_ok = transmit_calling_on_approach_track(
        approach_track,
        phase="RESTORE"
    )

    if not restore_ok:

        status = (
            f"{status} - "
            "50051 RESTORE FAILED"
        )

        result = "FAIL"

    return status, result



def test_other_signals_can_set(blocked_signal):

    log(f"CHECKING OTHER SIGNALS WHILE {blocked_signal} IS BLOCKED")

    for toc_row in lock_routes_data:

        other_signal = str(
            toc_row["signal"]
        ).replace(".0", "").strip().upper()

        # Do not test routes of the signal that is currently blocked.
        if other_signal == blocked_signal:
            continue

        if other_signal not in signals:

            report_rows.append({
                "BLOCKED_SIGNAL": blocked_signal,
                "MAIN_SIGNAL": other_signal,
                "MAIN_ROUTE": "",
                "TYPE": "",
                "STATUS": (
                    f"SIGNAL NOT CONFIGURED WHILE "
                    f"{blocked_signal} BLOCKED"
                ),
                "RESULT": "FAIL",
                "DATE & TIME":
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            continue

        other_info = signals[other_signal]
        other_type = other_info["type"]

        route_entries = toc_row.get("route_entries") or [
            {"route": route_name, "approach_track": ""}
            for route_name in toc_row.get("routes", [])
        ]

        for other_entry in route_entries:

            other_route = str(
                other_entry.get("route", "")
            ).strip().upper()

            other_approach_track = get_route_approach_track(
                toc_row,
                other_route
            )

            log(
                f"{blocked_signal} BLOCKED -> "
                f"TESTING {other_signal} {other_route}"
            )

            # Calling-On routes also need the approach track prepared
            # through 50051 before route selection.
            if other_type == "CALLING_ON":

                if not other_approach_track:

                    report_rows.append({
                        "BLOCKED_SIGNAL": blocked_signal,
                        "MAIN_SIGNAL": other_signal,
                        "MAIN_ROUTE": other_route,
                        "TYPE": other_type,
                        "STATUS": "APPROACH TRACK NOT CONFIGURED",
                        "RESULT": "FAIL",
                        "DATE & TIME":
                            datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    })

                    continue

                if not transmit_calling_on_approach_track(
                    other_approach_track,
                    phase="ISOLATION PREPARE"
                ):

                    report_rows.append({
                        "BLOCKED_SIGNAL": blocked_signal,
                        "MAIN_SIGNAL": other_signal,
                        "MAIN_ROUTE": other_route,
                        "TYPE": other_type,
                        "STATUS": "50051 APPROACH TRACK PREPARATION FAILED",
                        "RESULT": "FAIL",
                        "DATE & TIME":
                            datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    })

                    continue

            click(other_info["menu"])
            time.sleep(1)

            if not click_menu_item(
                other_route.replace("-", "_")
            ):

                report_rows.append({
                    "BLOCKED_SIGNAL": blocked_signal,
                    "MAIN_SIGNAL": other_signal,
                    "MAIN_ROUTE": other_route,
                    "TYPE": other_type,
                    "STATUS": (
                        f"ROUTE OPTION NOT FOUND WHILE "
                        f"{blocked_signal} BLOCKED"
                    ),
                    "RESULT": "FAIL",
                    "DATE & TIME":
                        datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                })

                pyautogui.press("esc")

                if other_type == "CALLING_ON":
                    transmit_calling_on_approach_track(
                        other_approach_track,
                        phase="ISOLATION RESTORE"
                    )

                continue

            time.sleep(1)

            # Other signals must be able to receive confirmation and set route.
            click_yes_dialog()
            time.sleep(1)

            if other_type == "MAIN":
                route_set = is_signal_changed(other_signal)

            elif other_type == "CALLING_ON":

                log(
                    f"{other_signal} {other_route} -> "
                    "VERIFYING ROUTE INITIATOR YELLOW"
                )

                route_set = calling_on_route_set(
                    other_signal
                )

                log(
                    f"{other_signal} {other_route} -> "
                    f"ROUTE INITIATOR RESULT = "
                    f"{'YELLOW / SET' if route_set else 'NOT YELLOW / NOT SET'}"
                )

            elif other_type == "SHUNT":
                route_set = shunt_fail(other_signal)

            else:
                route_set = False

            if route_set:

                result = "PASS"
                status = (
                    f"ROUTE INITIATOR YELLOW - "
                    f"ROUTE SET WHILE {blocked_signal} "
                    f"BUTTON IS BLOCKED"
                )

            else:

                result = "FAIL"
                status = (
                    f"ROUTE INITIATOR NOT YELLOW - "
                    f"ROUTE NOT SET WHILE {blocked_signal} "
                    f"BUTTON IS BLOCKED"
                )

                pyautogui.press("esc")
                time.sleep(1)

            # Calling-On isolation cleanup:
            #
            # ONLY if the route actually SET (Route Initiator YELLOW),
            # perform Signal Cancel + Route Release.
            #
            # If the route did NOT set, skip both commands and only
            # restore the 50051 approach-track preparation.
            if other_type == "CALLING_ON":

                if route_set:

                    log(
                        f"{other_signal} {other_route} -> "
                        "ROUTE SET DETECTED: "
                        "PERFORMING SIGNAL CANCEL + ROUTE RELEASE"
                    )

                    status, result = calling_on_cleanup_and_restore(
                        other_signal,
                        other_approach_track,
                        status,
                        result
                    )

                else:

                    log(
                        f"{other_signal} {other_route} -> "
                        "ROUTE NOT SET: "
                        "SKIPPING SIGNAL CANCEL + ROUTE RELEASE"
                    )

                    restore_ok = transmit_calling_on_approach_track(
                        other_approach_track,
                        phase="ISOLATION RESTORE"
                    )

                    if not restore_ok:

                        result = "FAIL"
                        status = (
                            f"{status} - "
                            "50051 RESTORE FAILED"
                        )

            elif route_set:

                cleanup_ok = cancel_and_release_route(
                    other_signal
                )

                if not cleanup_ok:
                    result = "FAIL"
                    status = (
                        f"{status} - "
                        "CLEANUP FAILED"
                    )

            report_rows.append({
                "BLOCKED_SIGNAL": blocked_signal,
                "MAIN_SIGNAL": other_signal,
                "MAIN_ROUTE": other_route,
                "TYPE": other_type,
                "STATUS": status,
                "RESULT": result,
                "DATE & TIME":
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            log(
                f"{other_signal} {other_route} -> "
                f"{result}"
            )

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

    result = (
        r > 150 and
        g > 150 and
        b < 130
    )

    log(
        f"YELLOW CHECK "
        f"({x},{y}) RGB={r},{g},{b} -> {result}"
    )

    return result

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
    """
    SHUNT route detection.

    ONLY the ROUTE INITIATION indicator decides whether the route
    is set:

        YELLOW     -> ROUTE SET
        NOT YELLOW -> ROUTE NOT SET

    C1/C2/C3 are not used for the route-set decision.
    """

    if signal not in signals:
        log(f"{signal} NOT FOUND IN SIGNAL CONFIG")
        return False

    info = signals[signal]

    if str(info.get("type", "")).upper() != "SHUNT":
        log(f"{signal} IS NOT A SHUNT SIGNAL")
        return False

    route_init = info.get("route_init")

    # Compatibility with old in-memory configuration.
    if not route_init:
        route_init = info.get("C3") or info.get("C2")

    if not route_init:
        log(f"{signal} ROUTE INITIATION POINT NOT CONFIGURED")
        return False

    log(
        f"{signal} ROUTE INITIATION CHECK -> "
        f"({route_init[0]},{route_init[1]})"
    )

    route_set = is_yellow(route_init)

    if route_set:
        log(
            f"{signal} -> SHUNT ROUTE SET "
            f"(ROUTE INITIATION IS YELLOW)"
        )
    else:
        log(
            f"{signal} -> SHUNT ROUTE NOT DETECTED "
            f"(ROUTE INITIATION IS NOT YELLOW)"
        )

    return route_set

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



# =========================================================
# CALLING-ON / 50051 BIT CHART SUPPORT
# =========================================================

CALLING_ON_YELLOW_TIMEOUT = 20.0
CALLING_ON_YELLOW_POLL = 0.5


def find_and_click_50051(window_title, control_name, control_type=None):
    """
    EXACT UI pattern used by the working Emergency Crank Handle code.

    IMPORTANT:
    The Emergency Crank Handle implementation does NOT depend on a
    specific Station 50051 window title.  It searches all desktop windows,
    checks visible text, checks the control type when requested, and then
    uses click_input().

    Calling-On uses the same proven method here.
    """
    try:
        desktop = Desktop(backend="uia")

        for window in desktop.windows():
            try:
                for item in window.descendants():
                    try:
                        text = item.window_text().strip().upper()

                        if control_name.upper() not in text:
                            continue

                        if control_type:
                            current_type = str(
                                item.element_info.control_type
                            ).upper()

                            if control_type.upper() not in current_type:
                                continue

                        log(
                            f"Found: {item.window_text()} | "
                            f"Type: {item.element_info.control_type}"
                        )

                        item.click_input()
                        log(f"{control_name} CLICKED")
                        return True

                    except Exception as exc:
                        log(str(exc))

            except Exception as exc:
                log(str(exc))

    except Exception as exc:
        log(str(exc))

    log(f"{control_name} NOT FOUND")
    return False


def open_50051_bit():
    """
    OPEN 50051 exactly like Emergency Crank Handle:

        click panel
        -> Ctrl+B
        -> wait 3 sec
        -> click 50051 ONCE
        -> wait 1 sec
        -> ENTER
        -> wait 3 sec
        -> verify Stations - Select Station is CLOSED

    There is NO double-click of 50051.
    """
    log("CALLING-ON -> OPENING 50051")

    try:
        pyautogui.click(500, 500)
    except Exception:
        pass

    time.sleep(1)
    pause_event.wait()

    pyautogui.hotkey("ctrl", "b")
    log("CTRL+B PRESSED")
    time.sleep(3)

    # EXACT Emergency Crank Handle method.
    if not find_and_click_50051(
        "Stations - Select Station",
        "50051"
    ):
        log("50051 NOT FOUND")
        return False

    time.sleep(1)
    pause_event.wait()
    pyautogui.press("enter")
    log("ENTER PRESSED")
    time.sleep(3)

    # EXACT Emergency Crank Handle verification:
    # the station-selection dialog must be closed.
    try:
        desktop = Desktop(backend="uia")

        for window in desktop.windows():
            try:
                title = window.window_text().strip().upper()

                if "STATIONS - SELECT STATION" in title:
                    log("STATION WINDOW STILL OPEN")
                    return False

            except Exception as exc:
                log(str(exc))

    except Exception as exc:
        log(str(exc))

    log("STATION WINDOW CLOSED")
    log("50051 STATION SELECTED")
    return True


def select_50051_track(approach_track):
    """
    Select the Calling-On approach track in Station 50051.

    EXACT pattern from Emergency Crank Handle:

        find_and_click("Station 50051", ITEM_NAME, "ListItem")

    The helper intentionally searches the same way as the working code.
    """
    track_name = str(approach_track or "").strip().upper()

    if not track_name:
        log("APPROACH_TRACK NOT CONFIGURED")
        return False

    log(
        f"50051 -> SELECTING CALLING-ON APPROACH TRACK : {track_name}"
    )

    if not find_and_click_50051(
        "Station 50051",
        track_name,
        "ListItem"
    ):
        log(
            f"APPROACH TRACK {track_name} "
            "NOT FOUND IN 50051"
        )
        return False

    log(f"APPROACH TRACK {track_name} CLICKED")
    time.sleep(1)
    return True


def transmit_50051():
    """Click Transmit exactly like Emergency Crank Handle."""
    if not find_and_click_50051(
        "Station 50051",
        "Transmit",
        "Button"
    ):
        log("TRANSMIT BUTTON NOT FOUND")
        return False

    log("TRANSMIT CLICKED")
    time.sleep(3)
    return True


def close_50051_bit():
    """Click Cancel exactly like Emergency Crank Handle."""
    if not find_and_click_50051(
        "Station 50051",
        "Cancel",
        "Button"
    ):
        log("CANCEL BUTTON NOT FOUND")
        return False

    log("CANCEL CLICKED")
    time.sleep(1)
    return True


def transmit_calling_on_approach_track(approach_track, phase="PREPARE"):
    """
    CALLING-ON 50051 CONCEPT.

    PREPARE:
        1. Open Station 50051 using Ctrl+B.
        2. Select the Calling-On approach track from the TOC.
        3. Transmit.
        4. Cancel/close the 50051 station dialog.

    RESTORE:
        The exact same 50051 sequence is repeated after Signal Cancel /
        Route Release so the approach-track state is restored.
    """
    track = str(approach_track or "").strip().upper()

    if not track:
        log(
            f"CALLING-ON {phase} FAILED: "
            "APPROACH_TRACK IS EMPTY"
        )
        return False

    log(
        f"CALLING-ON {phase} -> "
        f"50051 APPROACH TRACK = {track}"
    )

    # 1. Open 50051 using the proven Emergency Crank Handle concept.
    if not open_50051_bit():
        log(f"CALLING-ON {phase} -> 50051 OPEN FAILED")
        # Do not leave a half-open Station dialog on screen -
        # it blocks the route menu of the next test.
        pyautogui.press("esc")
        time.sleep(1)
        return False

    # 2. Select the exact approach track from TOC.
    if not select_50051_track(track):
        log(
            f"CALLING-ON {phase} -> "
            f"APPROACH TRACK {track} SELECTION FAILED"
        )
        pyautogui.press("esc")
        return False

    # 3. Transmit the approach-track command.
    if not transmit_50051():
        log(f"CALLING-ON {phase} -> 50051 TRANSMIT FAILED")
        pyautogui.press("esc")
        return False

    # Allow the station command to propagate.
    time.sleep(1)

    # 4. Close/cancel exactly like Emergency Crank Handle.
    if not close_50051_bit():
        log(
            f"CALLING-ON {phase} -> "
            "50051 CANCEL FAILED - PRESSING ESC"
        )
        pyautogui.press("esc")
        time.sleep(1)
        return False

    log(
        f"CALLING-ON {phase} -> "
        "50051 APPROACH TRACK OPERATION COMPLETED"
    )
    return True

def wait_for_yellow(
    point,
    timeout=CALLING_ON_YELLOW_TIMEOUT,
    interval=CALLING_ON_YELLOW_POLL
):
    """
    Wait until the Calling-On Route Initiator becomes YELLOW.

    The lamp can take several seconds after route selection, so this
    deliberately polls the screen instead of doing one immediate check.
    """

    if point is None:
        log("ROUTE INITIATOR POINT NOT CONFIGURED")
        return False

    log(
        f"WAITING FOR ROUTE INITIATOR YELLOW "
        f"AT ({point[0]},{point[1]}) "
        f"TIMEOUT={timeout}s"
    )

    end_time = time.time() + timeout

    while time.time() < end_time:

        pause_event.wait()

        if not running:
            return False

        if is_yellow(point):
            log(
                "ROUTE INITIATOR -> "
                "YELLOW DETECTED -> ROUTE IS SET"
            )
            return True

        time.sleep(interval)

    log(
        "ROUTE INITIATOR -> "
        f"YELLOW NOT DETECTED WITHIN {timeout} SECONDS"
    )

    return False


def wait_for_route_initiator_off(
    point,
    timeout=15.0,
    interval=0.5
):
    """
    Wait until the Route Initiator is no longer yellow.

    This prevents a yellow indication left from a previous Calling-On
    route from being mistaken for the new route.
    """

    if point is None:
        log("ROUTE INITIATOR POINT NOT CONFIGURED")
        return False

    log(
        f"WAITING FOR ROUTE INITIATOR TO TURN OFF "
        f"AT ({point[0]},{point[1]})"
    )

    end_time = time.time() + timeout

    while time.time() < end_time:

        pause_event.wait()

        if not running:
            return False

        if not is_yellow(point):
            log(
                "ROUTE INITIATOR -> "
                "YELLOW OFF / READY"
            )
            return True

        time.sleep(interval)

    log(
        "ROUTE INITIATOR REMAINED YELLOW "
        f"FOR {timeout} SECONDS"
    )

    return False


def calling_on_route_set(signal_name):
    """
    Calling-On route verification MUST use the Route Initiator.

    Expected indication:
        Route Initiator lamp -> YELLOW

    A yellow lamp means the Calling-On route has actually been
    established.
    """

    if signal_name not in signals:
        log(
            f"{signal_name} NOT FOUND IN SIGNAL CONFIG"
        )
        return False

    info = signals[signal_name]

    route_init = info.get("ROUTE_INIT")

    if not route_init:
        log(
            f"{signal_name} ROUTE INITIATOR "
            "POINT NOT CONFIGURED"
        )
        return False

    return wait_for_yellow(route_init)


def calling_on_cancel_and_release(signal_name):
    """
    Calling-On cleanup sequence.

    EXACT REQUIRED ORDER:

        Signal Cancel
            ↓
        wait
            ↓
        Route Release
            ↓
        wait for release
            ↓
        return

    This is intentionally separate from the generic Main/Shunt cleanup
    so Calling-On cleanup is always executed after the Route Initiator
    verification.
    """

    if signal_name not in signals:
        log(
            f"{signal_name} NOT FOUND IN SIGNAL CONFIG"
        )
        return False

    info = signals[signal_name]
    route_init = info.get("ROUTE_INIT")

    cleanup_ok = True

    # =====================================================
    # 1. SIGNAL CANCEL
    # =====================================================

    log(
        f"{signal_name} -> "
        "CALLING-ON SIGNAL CANCEL"
    )

    signal_cancel_done = False

    for attempt in range(1, 4):

        try:
            pyautogui.press("esc")
            time.sleep(0.5)

            click(info["menu"])
            time.sleep(1)

            if click_menu_item("Signal Cancel"):

                log(
                    f"{signal_name} -> "
                    f"SIGNAL CANCEL CLICKED "
                    f"(ATTEMPT {attempt})"
                )

                signal_cancel_done = True
                break

        except Exception as exc:

            log(
                f"{signal_name} -> "
                f"SIGNAL CANCEL ATTEMPT {attempt} "
                f"ERROR : {exc}"
            )

        pyautogui.press("esc")
        time.sleep(1)

    if not signal_cancel_done:

        log(
            f"{signal_name} -> "
            "SIGNAL CANCEL FAILED"
        )

        cleanup_ok = False

    else:

        time.sleep(3)

        # After Signal Cancel, the Route Initiator should eventually
        # leave yellow.  This is diagnostic/state verification; even if
        # the indication cannot be detected, continue to Route Release.
        if route_init:

            wait_for_route_initiator_off(
                route_init,
                timeout=15.0,
                interval=0.5
            )

    # =====================================================
    # 2. ROUTE RELEASE
    # =====================================================

    log(
        f"{signal_name} -> "
        "CALLING-ON ROUTE RELEASE"
    )

    route_release_done = False

    for attempt in range(1, 4):

        try:
            pyautogui.press("esc")
            time.sleep(0.5)

            click(info["menu"])
            time.sleep(1)

            if click_menu_item("Route Release"):

                log(
                    f"{signal_name} -> "
                    f"ROUTE RELEASE CLICKED "
                    f"(ATTEMPT {attempt})"
                )

                route_release_done = True
                break

        except Exception as exc:

            log(
                f"{signal_name} -> "
                f"ROUTE RELEASE ATTEMPT {attempt} "
                f"ERROR : {exc}"
            )

        pyautogui.press("esc")
        time.sleep(1)

    if not route_release_done:

        log(
            f"{signal_name} -> "
            "ROUTE RELEASE FAILED"
        )

        cleanup_ok = False

    else:

        log(
            f"{signal_name} -> "
            "ROUTE RELEASE COMMAND SENT"
        )

        # Same interlocking release wait used elsewhere in the project.
        time.sleep(20)

        log(
            f"{signal_name} -> "
            "CALLING-ON ROUTE RELEASE COMPLETED"
        )

    return cleanup_ok

# =========================================================
# CREATE ADVANCED REPORT
# =========================================================

def create_report():

    global REPORT_FILE

    wb = Workbook()

    ws = wb.active

    ws.title = "BUTTON BLOCK REPORT"


    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:G1")

    title_cell = ws["A1"]

    title_cell.value = "BUTTON BLOCK TEST REPORT"

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

    ws.merge_cells("A2:G2")

    date_cell = ws["A2"]

    date_cell.value = (
        "Generated : "
        f"{datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
    )

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
        "BLOCKED SIGNAL",
        "TEST SIGNAL",
        "ROUTE",
        "TYPE",
        "STATUS",
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
            row.get("BLOCKED_SIGNAL", ""),
            row.get(
                "MAIN_SIGNAL",
                row.get("SIGNAL", "")
            ),
            row.get(
                "MAIN_ROUTE",
                "N/A"
            ),
            row.get("TYPE", ""),
            row.get("STATUS", ""),
            row.get("RESULT", ""),
            row.get("DATE & TIME", "")
        ])


    # =====================================================
    # ROW FORMATTING
    # =====================================================
    # ONLY FAIL ROWS ARE RED.
    # PASS rows remain white.

    fail_fill = PatternFill(
        "solid",
        fgColor="FF6666"
    )

    white_fill = PatternFill(
        fill_type=None
    )

    for row in ws.iter_rows(
        min_row=4
    ):

        result = str(
            row[5].value or ""
        ).strip().upper()

        fill = (
            fail_fill
            if result == "FAIL"
            else white_fill
        )

        for cell in row:

            cell.fill = fill

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )


    # =====================================================
    # SUMMARY
    # =====================================================

    route_results = {}

    route_display_names = {}

    for row in report_rows:

        route = str(
            row.get(
                "MAIN_ROUTE",
                "N/A"
            )
        ).strip()

        if route == "N/A" or not route:
            continue

        route_key = route.upper()

        route_display_names[
            route_key
        ] = route

        route_results.setdefault(
            route_key,
            []
        ).append(
            str(
                row.get("RESULT", "")
            ).strip().upper()
        )


    total_routes = len(
        route_display_names
    )

    initiated_route_names = []

    for route_key, results in route_results.items():

        if results and all(
            result == "PASS"
            for result in results
        ):

            initiated_route_names.append(
                route_display_names[route_key]
            )


    initiated_routes = len(
        initiated_route_names
    )

    non_initiated_routes = max(
        total_routes - initiated_routes,
        0
    )

    initiated_route_text = ", ".join(
        initiated_route_names
    )


    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL ROUTES"

    ws[f"B{start}"] = total_routes


    ws[f"A{start + 1}"] = "INITIATED ROUTES"

    ws[f"B{start + 1}"] = initiated_routes

    ws[f"C{start + 1}"] = initiated_route_text


    ws[f"A{start + 2}"] = "NON INITIATED ROUTES"

    ws[f"B{start + 2}"] = non_initiated_routes


    for r in range(
        start,
        start + 3
    ):

        for column in (
            "A",
            "B",
            "C"
        ):

            ws[f"{column}{r}"].font = Font(
                bold=True
            )

            ws[f"{column}{r}"].fill = white_fill

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
            vertical="center"
        )


    if initiated_route_text:

        ws.column_dimensions["C"].width = min(
            max(
                len(initiated_route_text) + 5,
                25
            ),
            100
        )


    # =====================================================
    # AUTO WIDTH
    # =====================================================

    for column in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column[0].column
        )

        for cell in column:

            try:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

            except Exception:
                pass

        adjusted = max_length + 5

        if column_letter != "C":

            ws.column_dimensions[
                column_letter
            ].width = min(
                max(adjusted, 12),
                45
            )


    # =====================================================
    # TIMESTAMPED SAVE
    # =====================================================

    REPORT_FILE = (
        "BUTTON_BLOCK_REPORT_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        ".xlsx"
    )

    wb.save(
        REPORT_FILE
    )

    log(
        f"REPORT SAVED : {REPORT_FILE}"
    )

# =========================================================
# CLICK YES ON SIGNAL BUTTON DIALOG
# =========================================================

def click_yes_dialog():

    log("Waiting for confirmation dialog...")

    end_time = time.time() + 3

    while time.time() < end_time:

        try:

            panel = auto.WindowControl(Name="Test Panel")

            if panel.Exists(1):

                dialog = panel.WindowControl(searchDepth=10)

                if dialog.Exists(1):

                    log(f"Dialog Found : {dialog.Name}")

                    yes_btn = dialog.ButtonControl(Name="Yes")

                    if yes_btn.Exists(1):

                        yes_btn.Click()

                        log("YES BUTTON CLICKED")

                        return True

        except Exception as e:
            log(f"Dialog Error : {e}")

        time.sleep(0.5)

    log("Confirmation dialog not found")

    return False
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
            break

        signal_name = str(
            row["signal"]
        ).replace(".0", "").strip().upper()

        route_entries = row.get("route_entries") or [
            {
                "route": route_name,
                "approach_track": ""
            }
            for route_name in row.get("routes", [])
        ]

        log("=" * 60)
        log(f"SIGNAL : {signal_name}")

        # ==========================================
        # CHECK CONFIG
        # ==========================================

        if signal_name not in signals:

            log(f"{signal_name} NOT FOUND IN SIGNAL CONFIG")

            report_rows.append({
                "BLOCKED_SIGNAL": signal_name,
                "MAIN_SIGNAL": signal_name,
                "MAIN_ROUTE": "N/A",
                "TYPE": "UNKNOWN",
                "STATUS": "SIGNAL NOT CONFIGURED",
                "RESULT": "FAIL",
                "DATE & TIME":
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            continue

        sig_info = signals[signal_name]
        sig_type = sig_info["type"]

        log(f"SIGNAL TYPE : {sig_type}")

        # ==========================================
        # BUTTON BLOCK TEST
        # ==========================================

        block_status = button_block(signal_name)

        if block_status != "BLOCKED":

            report_rows.append({
                "BLOCKED_SIGNAL": signal_name,
                "MAIN_SIGNAL": signal_name,
                "MAIN_ROUTE": "",
                "TYPE": sig_type,
                "STATUS": (
                    "SIGNAL ALREADY BLOCKED"
                    if block_status == "ALREADY_BLOCKED"
                    else "BUTTON BLOCK FAILED"
                ),
                "RESULT": "FAIL",
                "DATE & TIME":
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            continue

        # ==========================================
        # TEST EACH TOC ROUTE
        # ==========================================

        for route_entry in route_entries:

            main_route = str(
                route_entry.get("route", "")
            ).strip().upper()

            approach_track = get_route_approach_track(
                row,
                main_route
            )

            log(
                f"TESTING : {signal_name} -> {main_route}"
            )

            # --------------------------------------------------
            # CALLING-ON ONLY: PREPARE APPROACH TRACK FIRST
            # --------------------------------------------------
            if sig_type == "CALLING_ON":

                log(
                    f"{signal_name} {main_route} -> "
                    f"CALLING-ON APPROACH TRACK = "
                    f"{approach_track or '[EMPTY]'}"
                )

                if not approach_track:

                    log(
                        f"{signal_name} {main_route} -> "
                        "APPROACH_TRACK MISSING IN TOC"
                    )

                    report_rows.append({
                        "BLOCKED_SIGNAL": signal_name,
                        "MAIN_SIGNAL": signal_name,
                        "MAIN_ROUTE": main_route,
                        "TYPE": sig_type,
                        "STATUS": "APPROACH TRACK NOT CONFIGURED",
                        "RESULT": "FAIL",
                        "DATE & TIME":
                            datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    })

                    continue

                if not transmit_calling_on_approach_track(
                    approach_track,
                    phase="PREPARE"
                ):

                    report_rows.append({
                        "BLOCKED_SIGNAL": signal_name,
                        "MAIN_SIGNAL": signal_name,
                        "MAIN_ROUTE": main_route,
                        "TYPE": sig_type,
                        "STATUS": "50051 APPROACH TRACK PREPARATION FAILED",
                        "RESULT": "FAIL",
                        "DATE & TIME":
                            datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    })

                    continue

                # The previous Calling-On route must be completely clear
                # before this route is selected.  Otherwise an old yellow
                # Route Initiator can be mistaken for the new route.
                route_init = sig_info.get("ROUTE_INIT")

                if route_init:

                    wait_for_route_initiator_off(
                        route_init,
                        timeout=15.0,
                        interval=0.5
                    )

            # --------------------------------------------------
            # SELECT ROUTE
            # --------------------------------------------------

            click(sig_info["menu"])
            time.sleep(1)

            if not click_menu_item(
                main_route.replace("-", "_")
            ):

                log(
                    f"ROUTE OPTION NOT FOUND : {main_route}"
                )

                report_rows.append({
                    "BLOCKED_SIGNAL": signal_name,
                    "MAIN_SIGNAL": signal_name,
                    "MAIN_ROUTE": main_route,
                    "TYPE": sig_type,
                    "STATUS": "ROUTE OPTION NOT FOUND",
                    "RESULT": "FAIL",
                    "DATE & TIME":
                        datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                })

                pyautogui.press("esc")

                # Calling-On must restore the approach-track state before
                # moving to the next route.
                if sig_type == "CALLING_ON":
                    transmit_calling_on_approach_track(
                        approach_track,
                        phase="RESTORE"
                    )

                continue

            time.sleep(1)

            # Route confirmation may not appear when Button Block works.
            click_yes_dialog()
            time.sleep(1)

            # --------------------------------------------------
            # ROUTE-SET VERIFICATION
            # --------------------------------------------------

            if sig_type == "MAIN":

                route_set = is_signal_changed(signal_name)

            elif sig_type == "CALLING_ON":

                log(
                    f"{signal_name} {main_route} -> "
                    "VERIFYING ROUTE INITIATOR YELLOW"
                )

                # Calling-On Route Initiator changes to YELLOW only
                # after the route is actually established.  It may take
                # several seconds, therefore poll until timeout.
                route_set = calling_on_route_set(signal_name)

                log(
                    f"{signal_name} {main_route} -> "
                    f"ROUTE INITIATOR RESULT = "
                    f"{'YELLOW / SET' if route_set else 'NOT YELLOW / NOT SET'}"
                )

            elif sig_type == "SHUNT":

                log(
                    f"{signal_name} -> SHUNT ROUTE CHECK "
                    "USING ROUTE INITIATION ONLY"
                )

                route_set = shunt_fail(signal_name)

            else:
                route_set = False

            # --------------------------------------------------
            # RESULT + CLEANUP
            # --------------------------------------------------

            if route_set:

                # Route Initiator became YELLOW.
                # Therefore the Calling-On route actually set.
                result = "FAIL"
                status = (
                    "ROUTE INITIATOR YELLOW - "
                    "ROUTE SET AFTER BUTTON BLOCK"
                )

            else:

                # Route Initiator did not become YELLOW within the
                # configured timeout, so Button Block is working.
                result = "PASS"
                status = (
                    "ROUTE INITIATOR NOT YELLOW - "
                    "BUTTON BLOCK WORKING"
                )

            # --------------------------------------------------
            # CALLING-ON CLEANUP
            #
            # IMPORTANT:
            # Signal Cancel + Route Release are required ONLY when
            # the Calling-On route actually SET (Route Initiator
            # became YELLOW).
            #
            # If Route Initiator did NOT become YELLOW:
            #   - DO NOT Signal Cancel
            #   - DO NOT Route Release
            #   - Only restore the 50051 approach-track preparation
            #
            # If Route Initiator DID become YELLOW:
            #   - Signal Cancel
            #   - Route Release
            #   - Restore 50051
            # --------------------------------------------------

            if sig_type == "CALLING_ON":

                if route_set:

                    log(
                        f"{signal_name} {main_route} -> "
                        "ROUTE SET DETECTED: "
                        "PERFORMING SIGNAL CANCEL + ROUTE RELEASE"
                    )

                    status, result = calling_on_cleanup_and_restore(
                        signal_name,
                        approach_track,
                        status,
                        result
                    )

                else:

                    log(
                        f"{signal_name} {main_route} -> "
                        "ROUTE NOT SET: "
                        "SKIPPING SIGNAL CANCEL + ROUTE RELEASE"
                    )

                    # Route did not establish, so there is nothing to
                    # cancel or release. Only restore the 50051 state.
                    #
                    # Close the leftover route menu / popup first
                    # (same as the other-signals path), otherwise
                    # Ctrl+B does not open the Station window.
                    pyautogui.press("esc")
                    time.sleep(1)

                    restore_ok = transmit_calling_on_approach_track(
                        approach_track,
                        phase="RESTORE"
                    )

                    if not restore_ok:

                        log(
                            f"{signal_name} {main_route} -> "
                            "50051 RESTORE FAILED"
                        )

                        status = (
                            f"{status} - "
                            "50051 RESTORE FAILED"
                        )

                        result = "FAIL"

            else:

                # Main/Shunt retain the existing cleanup behavior.
                if route_set:

                    cleanup_ok = cancel_and_release_route(
                        signal_name
                    )

                    if not cleanup_ok:

                        log(
                            f"{signal_name} {main_route} -> "
                            "ROUTE CLEANUP FAILED"
                        )

                        status = (
                            f"{status} - "
                            "CLEANUP FAILED"
                        )

                        result = "FAIL"

                else:

                    pyautogui.press("esc")
                    time.sleep(1)

            report_rows.append({
                "BLOCKED_SIGNAL": signal_name,
                "MAIN_SIGNAL": signal_name,
                "MAIN_ROUTE": main_route,
                "TYPE": sig_type,
                "STATUS": status,
                "RESULT": result,
                "DATE & TIME":
                    datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            })

            log(
                f"{signal_name} {main_route} -> {result}"
            )

        # ==========================================
        # ISOLATION TEST
        # ==========================================

        test_other_signals_can_set(signal_name)

        # ==========================================
        # UNBLOCK SIGNAL
        # ==========================================

        if not button_unblock(signal_name):
            log(
                f"{signal_name} BUTTON UNBLOCK FAILED"
            )

        log(
            f"{signal_name} TEST COMPLETED"
        )

    # =========================================================
    # FINISH AUTOMATION
    # =========================================================

    running = False

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    create_report()

    root.deiconify()

    log("=" * 60)
    log("BUTTON BLOCK AUTOMATION COMPLETED")
    log("=" * 60)

    try:
        os.startfile(REPORT_FILE)
    except:
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
    # row will hit "NOT FOUND IN SIGNAL CONFIG" and the run will
    # finish in a couple seconds with a report full of failures -
    # looking exactly like nothing was tested. Catch that here
    # instead, loudly, before starting.
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
# =========================================================
# GUI
# =========================================================

root = tk.Tk()

root.title(
    "BUTTON BLOCK TESTING AUTOMATION"
)

root.geometry(
    "1450x900"
)

root.configure(
    bg="#e9edf2"
)

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
    rowheight=32,
    fieldbackground="white",
    font=("Segoe UI", 10)
)

style.configure(
    "Treeview.Heading",
    background="#1e293b",
    foreground="white",
    font=("Segoe UI", 10, "bold")
)

style.configure(
    "TNotebook",
    background="#e9edf2",
    borderwidth=0
)

style.configure(
    "TNotebook.Tab",
    font=("Segoe UI", 10, "bold"),
    padding=(18, 8)
)


# =========================================================
# HEADER
# =========================================================

header_frame = tk.Frame(
    root,
    bg="#0f172a",
    height=100
)

header_frame.pack(
    side="top",
    fill="x"
)

header_frame.pack_propagate(False)


# ---------------------------------------------------------
# LEFT LOGO - INDIAN RAILWAYS
# ---------------------------------------------------------

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
        bg="#0f172a"
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


# ---------------------------------------------------------
# CENTER TITLE
# ---------------------------------------------------------

title_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

title_frame.pack(
    side="left",
    fill="both",
    expand=True
)

tk.Label(
    title_frame,
    text="BUTTON BLOCK TESTING SYSTEM",
    font=("Segoe UI", 24, "bold"),
    bg="#0f172a",
    fg="white"
).pack(
    pady=(13, 0)
)

tk.Label(
    title_frame,
    text="COORDINATE CAPTURE  |  ROUTE TESTING  |  REPORTING",
    font=("Segoe UI", 10),
    bg="#0f172a",
    fg="#cbd5e1"
).pack(
    pady=(3, 0)
)


# ---------------------------------------------------------
# RIGHT LOGO
# ---------------------------------------------------------

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
        text="EDRC TEAM",
        font=("Segoe UI", 10, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack(pady=5)


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
    pady=(12, 8)
)


# =========================================================
# LEFT CONTROL PANEL
# =========================================================

left_panel = tk.Frame(
    main_frame,
    bg="white",
    bd=1,
    relief="solid",
    width=245
)

left_panel.pack(
    side="left",
    fill="y",
    padx=(0, 10)
)

left_panel.pack_propagate(False)


tk.Label(
    left_panel,
    text="CONTROL PANEL",
    font=("Segoe UI", 15, "bold"),
    bg="white",
    fg="#0f172a"
).pack(
    pady=20
)


def create_button(
    text,
    command,
    color,
    pady=8
):

    button = tk.Button(
        left_panel,
        text=text,
        command=command,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 10, "bold"),
        width=25,
        height=2
    )

    button.pack(
        pady=pady,
        padx=5
    )

    return button


# ---------------------------------------------------------
# SIGNAL CONFIGURATION
# ---------------------------------------------------------

tk.Button(
    left_panel,
    text="CAPTURE SIGNALLING GEARS",
    command=start_quantity_capture,
    bg="#2563eb",
    fg="white",
    activebackground="#2563eb",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 10, "bold"),
    width=25,
    height=2
).pack(
    pady=(0, 8),
    padx=5
)


small_button_frame = tk.Frame(
    left_panel,
    bg="white"
)

small_button_frame.pack(
    pady=3
)


tk.Button(
    small_button_frame,
    text="NEW SIGNAL",
    command=start_quantity_capture,
    bg="#ea580c",
    fg="white",
    activebackground="#ea580c",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=12
).pack(
    side="left",
    padx=3
)


tk.Button(
    small_button_frame,
    text="EXISTING SIGNAL",
    command=load_config,
    bg="#ea580c",
    fg="white",
    activebackground="#ea580c",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=12
).pack(
    side="left",
    padx=3
)


create_button(
    "LOAD TOC EXCEL",
    load_lock_routes,
    "#2563eb",
    10
)

create_button(
    "SAVE NEW SIGNAL COORDINATES",
    save_new_signal_coordinates,
    "#16a34a",
    10
)

create_button(
    "LOAD YARD COORDINATES",
    load_universal_coordinates,
    "#7c3aed",
    10
)

create_button(
    "START TESTING",
    start_automation,
    "#16a34a",
    18
)

create_button(
    "STOP TESTING",
    stop_automation,
    "#dc2626",
    8
)


# ---------------------------------------------------------
# SYSTEM STATUS
# ---------------------------------------------------------

status_frame = tk.Frame(
    left_panel,
    bg="#f8fafc",
    bd=1,
    relief="solid"
)

status_frame.pack(
    fill="x",
    padx=12,
    pady=22
)

tk.Label(
    status_frame,
    text="SYSTEM STATUS",
    font=("Segoe UI", 10, "bold"),
    bg="#f8fafc",
    fg="#0f172a"
).pack(
    pady=(12, 5)
)

status_label = tk.Label(
    status_frame,
    text="READY",
    font=("Segoe UI", 15, "bold"),
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
# SIGNAL NOTEBOOK
# =========================================================

signal_notebook = ttk.Notebook(
    right_panel
)

signal_notebook.pack(
    fill="both",
    expand=True
)


main_signal_tab = tk.Frame(
    signal_notebook,
    bg="white"
)

calling_on_tab = tk.Frame(
    signal_notebook,
    bg="white"
)

shunt_signal_tab = tk.Frame(
    signal_notebook,
    bg="white"
)


signal_notebook.add(
    main_signal_tab,
    text="MAIN SIGNALS"
)

signal_notebook.add(
    calling_on_tab,
    text="CALLING-ON SIGNALS"
)

signal_notebook.add(
    shunt_signal_tab,
    text="SHUNT SIGNALS"
)


# =========================================================
# TABLE FACTORY
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
        frame,
        orient="vertical"
    )

    scrollbar.pack(
        side="right",
        fill="y"
    )

    signal_tree = ttk.Treeview(
        frame,
        columns=(
            "NO",
            "SIGNAL",
            "ROUTE"
        ),
        show="headings",
        yscrollcommand=scrollbar.set
    )

    scrollbar.config(
        command=signal_tree.yview
    )

    signal_tree.heading(
        "NO",
        text="NO"
    )

    signal_tree.heading(
        "SIGNAL",
        text="SIGNAL"
    )

    signal_tree.heading(
        "ROUTE",
        text="ROUTE"
    )

    signal_tree.column(
        "NO",
        width=80,
        anchor="center"
    )

    signal_tree.column(
        "SIGNAL",
        width=180,
        anchor="center"
    )

    signal_tree.column(
        "ROUTE",
        width=650,
        anchor="w"
    )

    signal_tree.pack(
        fill="both",
        expand=True
    )

    return signal_tree


main_tree = create_signal_tree(
    main_signal_tab
)

calling_on_tree = create_signal_tree(
    calling_on_tab
)

shunt_tree = create_signal_tree(
    shunt_signal_tab
)


# =========================================================
# LIVE OPERATION LOG
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
    pady=(10, 8)
)


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


log_text = tk.Text(
    log_frame,
    bg="white",
    fg="black",
    font=("Consolas", 10),
    relief="flat",
    wrap="word"
)

log_text.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

log_text.config(
    state="disabled"
)


# =========================================================
# FOOTER
# =========================================================

footer_frame = tk.Frame(
    root,
    bg="#0f172a",
    height=32
)

footer_frame.pack(
    side="bottom",
    fill="x"
)

footer_frame.pack_propagate(False)

tk.Label(
    footer_frame,
    text=(
        "SPACE = CAPTURE COORDINATES"
        "   |   "
        "P = PAUSE / RESUME"
        "   |   "
        "BUTTON BLOCK TESTING SYSTEM"
    ),
    bg="#0f172a",
    fg="white",
    font=("Segoe UI", 9, "bold")
).pack(
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
# =========================================================
# RUN
# =========================================================
listener = pynput_keyboard.Listener(on_press=on_press)
listener.start()

root.mainloop()


