# =========================================================
# UNIVERSAL LOCK ROUTE AUTOMATION SYSTEM
# =========================================================
# FEATURES
# =========================================================
# 1. MAIN SIGNAL SUPPORT
# 2. CALLING ON SUPPORT
# 3. SHUNT SUPPORT
# 4. SIGNAL CONFI
# G SAVE/LOAD
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

import win32api
import win32con
import win32gui
from collections import Counter

import uiautomation as auto

from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "NEW_SIGNAL.xlsx"


REPORT_FILE = "VISUAL_INSPECTION_Report.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

crank_handle_points = {}

lc_points = {}

point_indications = {}

point_visual_items = {}

lock_routes_data = []

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

def capture_space_from_overlay(event=None):

    global last_space_time

    now = time.time()

    # Prevent double capture
    if now - last_space_time < 0.4:
        return "break"

    last_space_time = now

    # Only capture during coordinate recording
    if capture_module != "MASTER":
        return "break"

    # Get current mouse position
    x, y = win32api.GetCursorPos()

    log(
        f"SPACE -> CAPTURE ({x}, {y})"
    )

    # Send to Tkinter main thread
    root.after(
        0,
        lambda x=x, y=y:
        dispatch_save_point(x, y)
    )

    return "break"

def create_capture_overlay():

    global capture_overlay
    global overlay_progress_label
    global overlay_signal_label
    global overlay_step_label
    global overlay_hint_label
    global overlay_button_frame

    # ---------------------------------------------------------
    # Destroy previous overlay if it exists
    # ---------------------------------------------------------

    if capture_overlay is not None:

        try:
            capture_overlay.destroy()
        except:
            pass

        capture_overlay = None

    # ---------------------------------------------------------
    # CREATE OVERLAY
    # ---------------------------------------------------------

    capture_overlay = tk.Toplevel(root)

    capture_overlay.title("Coordinate Capture Guide")

    capture_overlay.configure(
        bg="#0f172a"
    )

    capture_overlay.resizable(
        False,
        False
    )

    capture_overlay.attributes(
        "-topmost",
        True
    )

    # ---------------------------------------------------------
    # POSITION TOP RIGHT
    # ---------------------------------------------------------

    screen_w = capture_overlay.winfo_screenwidth()

    width = 420
    height = 320

    x = screen_w - width - 20
    y = 40

    capture_overlay.geometry(
        f"{width}x{height}+{x}+{y}"
    )

    # ---------------------------------------------------------
    # CLOSE
    # ---------------------------------------------------------

    capture_overlay.protocol(
        "WM_DELETE_WINDOW",
        cancel_master_capture
    )

    capture_overlay.bind(
        "<space>",
        capture_space_from_overlay
    )

    capture_overlay.bind(
        "<BackSpace>",
        lambda event: undo_last_capture()
    )

    # Prevent SPACE from activating buttons


    # ---------------------------------------------------------
    # TITLE
    # ---------------------------------------------------------

    tk.Label(
        capture_overlay,
        text="COORDINATE CAPTURE",
        font=("Segoe UI", 11, "bold"),
        bg="#0f172a",
        fg="#64748b"
    ).pack(
        pady=(16, 0)
    )

    # ---------------------------------------------------------
    # PROGRESS
    # ---------------------------------------------------------

    overlay_progress_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 10),
        bg="#0f172a",
        fg="#94a3b8"
    )

    overlay_progress_label.pack(
        pady=(2, 10)
    )

    # ---------------------------------------------------------
    # SIGNAL
    # ---------------------------------------------------------

    overlay_signal_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 18, "bold"),
        bg="#0f172a",
        fg="white"
    )

    overlay_signal_label.pack()

    # ---------------------------------------------------------
    # STEP
    # ---------------------------------------------------------

    overlay_step_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 15, "bold"),
        bg="#0f172a",
        fg="#22c55e"
    )

    overlay_step_label.pack(
        pady=(6, 8)
    )

    # ---------------------------------------------------------
    # HINT
    # ---------------------------------------------------------

    overlay_hint_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 10),
        bg="#0f172a",
        fg="#cbd5e1",
        wraplength=380,
        justify="center"
    )

    overlay_hint_label.pack(
        pady=(0, 10)
    )

    # ---------------------------------------------------------
    # BUTTON FRAME
    # ---------------------------------------------------------

    overlay_button_frame = tk.Frame(
        capture_overlay,
        bg="#0f172a"
    )

    overlay_button_frame.pack(
        pady=(0, 6)
    )

    # ---------------------------------------------------------
    # UNDO
    # ---------------------------------------------------------

    tk.Button(
        capture_overlay,
        text="↶  UNDO LAST CLICK",
        command=undo_last_capture,
        bg="#f59e0b",
        fg="#1a1a1a",
        activebackground="#fbbf24",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 10, "bold"),
        width=22,
        takefocus=0
    ).pack(
        pady=(10, 2)
    )

    tk.Label(
        capture_overlay,
        text="(or press BACKSPACE)",
        font=("Segoe UI", 8),
        bg="#0f172a",
        fg="#475569"
    ).pack()

    # ---------------------------------------------------------
    # CANCEL
    # ---------------------------------------------------------

    tk.Button(
        capture_overlay,
        text="CANCEL CAPTURE",
        command=cancel_master_capture,
        bg="#0f172a",
        fg="#64748b",
        activebackground="#0f172a",
        activeforeground="#ef4444",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 9, "underline"),
        takefocus=0
    ).pack(
        side="bottom",
        pady=(0, 10)
    )

    # ---------------------------------------------------------
    # SHOW
    # ---------------------------------------------------------

    capture_overlay.deiconify()
    capture_overlay.lift()
    capture_overlay.attributes("-topmost", True)
    capture_overlay.update_idletasks()
    capture_overlay.lift()
    capture_overlay.focus_force()
    capture_overlay.update_idletasks()

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

    tk.Button(
        capture_overlay,
        text="↶  UNDO LAST CAPTURE",
        command=undo_last_capture,
        bg="#f59e0b",
        fg="#1a1a1a",
        activebackground="#fbbf24",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 9, "bold"),
        width=22,
        takefocus=0
    ).pack(pady=(12, 2))

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
    """Build the complete-yard capture list in the same order and stages as TL.

    Order:
        MAIN -> SHUNT -> CALLING-ON -> POINT -> CRANK HANDLE -> LC GATE

    The function only prepares capture data/UI state; it does not change the
    Visual Inspection automation engine.
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
            "ROUTE_INIT": None,
            # Legacy key retained for existing Visual Inspection logic.
            "ROUTE_INDICATOR": None,
        }
        combined.append(("MAIN", name))

    # =====================================================
    # SHUNT SIGNALS  (TL: MENU -> INDICATOR -> ROUTE INIT)
    # =====================================================
    count = prompt_count("SHUNT Signals")
    for i in range(count):
        name = prompt_name("SHUNT Signal", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "SHUNT",
            "menu": None,
            "state_indicator": None,
            "route_init": None,
            "initial_snapshot": None,
        }
        combined.append(("SHUNT", name))

    # =====================================================
    # CALLING-ON SIGNALS  (TL: MENU -> YELLOW -> ROUTE INIT)
    # =====================================================
    count = prompt_count("CALLING-ON Signals")
    for i in range(count):
        name = prompt_name("CALLING-ON Signal", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "CALLING_ON",
            "menu": None,
            "YELLOW": None,
            "ROUTE_INIT": None,
        }
        combined.append(("CALLING_ON", name))

    # =====================================================
    # POINTS  (TL: MENU -> NORMAL -> REVERSE -> FREE)
    # =====================================================
    count = prompt_count("POINTS")
    for i in range(count):
        name = prompt_name("POINT", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "POINT",
            "menu": None,
            "normal": None,
            "reverse": None,
            "free": None,
        }
        combined.append(("POINT", name))

    # =====================================================
    # CRANK HANDLES  (TL: MENU -> IN -> OUT -> ECH -> FREE)
    # =====================================================
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
            "FREE": None,
        }
        combined.append(("CH", name))

    # =====================================================
    # LC GATES  (TL: MENU -> IN -> OUT)
    # =====================================================
    count = prompt_count("LC GATES")
    for i in range(count):
        name = prompt_name("LC GATE", i, count, existing)
        existing.add(name)
        signals[name] = {
            "type": "LC",
            "menu": None,
            "IN": None,
            "OUT": None,
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

    # IMPORTANT:
    # Tell keyboard handler that MASTER capture is starting
    capture_module = "MASTER"

    capture_master_list = build_quantity_capture_list()

    # -----------------------------------------------------
    # No signals
    # -----------------------------------------------------

    if not capture_master_list:

        capture_module = None

        log("No signals entered")

        return

    # -----------------------------------------------------
    # Start capture
    # -----------------------------------------------------

    capture_index = 0

    set_initial_stage_for_current()

    create_capture_overlay()

    root.iconify()

    log("Recording Signal Coordinates")

    log(
        f"Total Signals : {len(capture_master_list)}"
    )

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

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if sig_type == "MAIN":

        record_stage = "coordinate"

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif sig_type == "SHUNT":

        record_stage = "menu"

    # =====================================================
    # CALLING-ON SIGNAL
    # =====================================================

    elif sig_type == "CALLING_ON":

        record_stage = "menu"

    # =====================================================
    # POINT
    # =====================================================

    elif sig_type == "POINT":

        record_stage = "menu"

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif sig_type == "CH":

        record_stage = "menu"

    # =====================================================
    # LC GATE
    # =====================================================

    elif sig_type == "LC":

        record_stage = "menu"

    log(
        f"INITIAL STAGE: "
        f"{sig_type} -> {signal} -> {record_stage}"
    )

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
            "All signal coordinates have been captured.\n\n"
            "Now click 'SAVE NEW SIGNAL COORDINATES'."
        )
        return

    sig_type, signal = capture_master_list[capture_index]

    # =====================================================
    # MAIN SIGNAL - TL: MENU -> ASPECTS -> ROUTE INIT
    # =====================================================
    if sig_type == "MAIN":
        if record_stage == "coordinate":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )
        elif record_stage == "ask_aspects":
            show_aspect_buttons(signal)
        elif record_stage == "RED":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : RED LAMP",
                "Move the mouse onto the RED lamp, then press SPACE."
            )
        elif record_stage == "YELLOW":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : YELLOW LAMP",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )
        elif record_stage == "DOUBLE_YELLOW":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : DOUBLE YELLOW LAMP",
                "Move the mouse onto the DOUBLE YELLOW lamp, then press SPACE."
            )
        elif record_stage == "GREEN":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : GREEN LAMP",
                "Move the mouse onto the GREEN lamp, then press SPACE."
            )
        elif record_stage == "ROUTE_INIT":
            update_capture_overlay(
                "MAIN", signal,
                "CLICK : ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # CALLING-ON - TL: MENU -> YELLOW -> ROUTE INIT
    # =====================================================
    elif sig_type == "CALLING_ON":
        if record_stage == "menu":
            update_capture_overlay(
                "CALLING_ON", signal,
                "CLICK : MENU",
                "Move the mouse onto this Calling-ON signal's menu button, then press SPACE."
            )
        elif record_stage == "YELLOW":
            update_capture_overlay(
                "CALLING_ON", signal,
                "CLICK : YELLOW LAMP",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )
        elif record_stage == "ROUTE_INIT":
            update_capture_overlay(
                "CALLING_ON", signal,
                "CLICK : ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # SHUNT - TL: MENU -> INDICATOR -> ROUTE INIT
    # =====================================================
    elif sig_type == "SHUNT":
        if record_stage == "menu":
            update_capture_overlay(
                "SHUNT", signal,
                "CLICK : MENU",
                "Move the mouse onto this shunt signal's menu button, then press SPACE."
            )
        elif record_stage == "indicator":
            update_capture_overlay(
                "SHUNT", signal,
                "CLICK : ASPECT INDICATOR",
                "Move the mouse onto the aspect indicator, then press SPACE.\n"
                "This step also saves the reference snapshot used by the existing inspection logic."
            )
        elif record_stage == "route_init":
            update_capture_overlay(
                "SHUNT", signal,
                "CLICK : ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    # =====================================================
    # POINT - TL: MENU -> NORMAL -> REVERSE -> FREE
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
    # CRANK HANDLE - TL: MENU -> IN -> OUT -> ECH -> FREE
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
    # LC GATE - TL: MENU -> IN -> OUT
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
    global capture_index

    log(
        f"MASTER SAVE ENTERED: x={x}, y={y}, "
        f"index={capture_index}, stage={record_stage}"
    )

    if capture_index >= len(capture_master_list):
        return

    sig_type, signal = capture_master_list[capture_index]

    # =====================================================
    # MAIN
    # =====================================================
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
            record_stage = "ROUTE_INIT"
            master_next_capture()
            return

        if record_stage == "ROUTE_INIT":
            push_undo(lambda s=signal: signals[s].__setitem__("ROUTE_INIT", None))
            signals[signal]["ROUTE_INIT"] = [x, y]
            # Legacy Visual Inspection key remains synchronized.
            signals[signal]["ROUTE_INDICATOR"] = [x, y]
            log(f"{signal} ROUTE INITIATION Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # CALLING-ON - MENU -> YELLOW -> ROUTE INIT
    # =====================================================
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

    # =====================================================
    # SHUNT - MENU -> INDICATOR -> ROUTE INIT
    # =====================================================
    elif sig_type == "SHUNT":
        if record_stage == "menu":
            push_undo(lambda s=signal: signals[s].__setitem__("menu", None))
            signals[signal]["menu"] = [x, y]
            log(f"{signal} MENU Saved")
            record_stage = "indicator"
            master_next_capture()
            return

        if record_stage == "indicator":
            push_undo(lambda s=signal: signals[s].__setitem__("state_indicator", None))
            signals[signal]["state_indicator"] = [x, y]
            # Preserve the existing snapshot mechanism if it exists.
            try:
                if "initial_snapshot" in signals[signal]:
                    # Do not force a screenshot here; existing inspection code
                    # can continue to manage snapshots independently.
                    pass
            except Exception:
                pass
            log(f"{signal} ASPECT INDICATOR Saved")
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":
            push_undo(lambda s=signal: signals[s].__setitem__("route_init", None))
            signals[signal]["route_init"] = [x, y]
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

    previous_aspects = signals[signal].get("aspects")

    push_undo(
        lambda s=signal, old=previous_aspects:
        signals[s].__setitem__("aspects", old)
    )

    current_aspects = aspects
    signals[signal]["aspects"] = aspects
    record_stage = "RED"

    log(f"{signal} ASPECTS SET -> {aspects}")
    master_next_capture()

def dispatch_save_point(x, y):

    global capture_module
    global capture_index
    global record_stage

    log(
        f"Dispatch -> {x}, {y}"
    )

    log(
        f"CAPTURE STATE: "
        f"module={capture_module}, "
        f"index={capture_index}, "
        f"stage={record_stage}, "
        f"total={len(capture_master_list)}"
    )

    if capture_module != "MASTER":
        log(
            "CAPTURE BLOCKED: capture_module is not MASTER"
        )
        return

    if capture_index >= len(capture_master_list):
        log(
            "CAPTURE BLOCKED: capture_index out of range"
        )
        return

    sig_type, signal = capture_master_list[capture_index]

    log(
        f"CAPTURE TARGET: "
        f"type={sig_type}, signal={signal}"
    )

    master_save_point(
        x,
        y
    )

def dispatch_undo():

    if capture_module == "MASTER":
        undo_last_capture()

    elif capture_module == "ROUTE":
        pass

# =========================================================
# GLOBAL KEYBOARD CAPTURE
# =========================================================

last_space_time = 0
last_backspace_time = 0
last_digit_time = 0


def on_press(key):

    global last_space_time
    global last_backspace_time
    global last_digit_time

    try:

        # =====================================================
        # SPACE = CAPTURE COORDINATE
        # =====================================================

        if key == pynput_keyboard.Key.space:

            now = time.time()

            if now - last_space_time < 0.4:
                return

            last_space_time = now

            # Only capture while MASTER recording is active
            if capture_module != "MASTER":
                return

            # Get current mouse position
            x, y = win32api.GetCursorPos()

            log(
                f"SPACE PRESSED -> CAPTURE ({x}, {y})"
            )

            # Send back to Tkinter main thread
            root.after(
                0,
                lambda x=x, y=y:
                dispatch_save_point(x, y)
            )

            return

        # =====================================================
        # BACKSPACE = UNDO
        # =====================================================

        if key == pynput_keyboard.Key.backspace:

            now = time.time()

            if now - last_backspace_time < 0.4:
                return

            last_backspace_time = now

            if capture_module != "MASTER":
                return

            root.after(
                0,
                dispatch_undo
            )

            return

        # =====================================================
        # 2 / 3 / 4 = ASPECT COUNT
        # =====================================================

        char = getattr(key, "char", None)

        if char in ("2", "3", "4"):

            if capture_module != "MASTER":
                return

            if capture_index >= len(capture_master_list):
                return

            sig_type, signal = (
                capture_master_list[capture_index]
            )

            if (
                sig_type == "MAIN"
                and record_stage == "ask_aspects"
            ):

                now = time.time()

                if now - last_digit_time < 0.4:
                    return

                last_digit_time = now

                root.after(
                    0,
                    lambda n=int(char):
                    set_aspects_and_continue(
                        signal,
                        n
                    )
                )

    except Exception as e:

        log(
            f"KEYBOARD CAPTURE ERROR: {e}"
        )

# =========================================================
# CLICK
# =========================================================

def click(point):

    if point is None:
        return False

    x, y = point

    pyautogui.moveTo(x, y, duration=1)

    log(f"Mouse moved to ({x}, {y})")

    time.sleep(5)   # Don't click yet

    pyautogui.click()

    return True

# ======================================================
# SCREEN COLOR HELPER
# ======================================================

def get_pixel_color(point):

    if point is None:
        return None

    x, y = point

    hdc = win32gui.GetDC(0)

    color = win32gui.GetPixel(hdc, x, y)

    win32gui.ReleaseDC(0, hdc)

    r = color & 255
    g = (color >> 8) & 255
    b = (color >> 16) & 255

    return (r, g, b)


def rebuild_point_indications():

    global point_indications

    point_indications = {}

    for point_no, info in signals.items():

        if info.get("type") != "POINT":
            continue

        normal = info.get("normal")
        reverse = info.get("reverse")

        if not normal or not reverse:
            log(
                f"POINT {point_no}: "
                "N/R COORDINATES NOT CONFIGURED"
            )
            continue

        point_indications[point_no] = {

            # Old automation naming
            "N": normal,
            "R": reverse,

            # Optional additional indication
            "FREE": info.get("free"),

            # Menu coordinate if needed later
            "MENU": info.get("menu")
        }

        log(
            f"POINT {point_no} CONFIGURED "
            f"N={normal} R={reverse}"
        )

def rebuild_crank_handle_points():

    global crank_handle_points

    crank_handle_points = {}

    for handle, info in signals.items():

        if info.get("type") != "CH":
            continue

        in_point = info.get("IN")
        out_point = info.get("OUT")
        ech_point = info.get("ECH")
        free_point = info.get("FREE")

        if not in_point:
            log(f"CH {handle}: IN COORDINATE NOT CONFIGURED")
            continue

        if not out_point:
            log(f"CH {handle}: OUT COORDINATE NOT CONFIGURED")
            continue

        if not ech_point:
            log(f"CH {handle}: ECH COORDINATE NOT CONFIGURED")
            continue

        if not free_point:
            log(f"CH {handle}: FREE COORDINATE NOT CONFIGURED")
            continue

        crank_handle_points[handle] = {

            "IN": in_point,
            "OUT": out_point,
            "ECH": ech_point,
            "FREE": free_point,

            # Keep MENU available for future CH operations
            "MENU": info.get("menu"),

            # FIX: click_crank_button() expects "BUTTON"; the universal
            # file stores the click point as Menu_X/Y.
            "BUTTON": info.get("BUTTON") or info.get("menu")
        }

        log(
            f"CH {handle} CONFIGURED "
            f"IN={in_point} "
            f"OUT={out_point} "
            f"ECH={ech_point} "
            f"FREE={free_point}"
        )

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
# CAPTURE CRANK HANDLE POINT
# =========================================================

def capture_crank_point(handle, point_name):

    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.update()

    messagebox.showinfo(
        "CAPTURE",
        f"{handle}\n\nCapture {point_name}\n\nPress SPACE"
    )

    root.attributes("-topmost", False)
    root.iconify()

    return capture_point()

# =========================================================
# CAPTURE TRACK COORDINATES
# =========================================================


# =========================================================
# CREATE SIGNAL SETUP
# =========================================================

def create_setup():

    global signals
    global crank_handle_points

    if not lock_routes_data:

        messagebox.showwarning(
            "WARNING",
            "Please load TOC file first."
        )
        return

    signals = {}

    crank_handle_points = {}


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

        # ------------------------------
        # 2 Aspect
        # ------------------------------

        if aspects == 2:

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture GREEN\n\nPress SPACE"
            )

            signals[name]["GREEN"] = capture_point()

        # ------------------------------
        # 3 Aspect
        # ------------------------------

        elif aspects == 3:

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture YELLOW\n\nPress SPACE"
            )

            signals[name]["YELLOW"] = capture_point()

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture GREEN\n\nPress SPACE"
            )

            signals[name]["GREEN"] = capture_point()

        # ------------------------------
        # 4 Aspect
        # ------------------------------

        elif aspects == 4:

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture YELLOW\n\nPress SPACE"
            )

            signals[name]["YELLOW"] = capture_point()

            messagebox.showinfo(
                "CAPTURE",
                f"{name}\n\nCapture DOUBLE YELLOW\n\nPress SPACE"
            )

            signals[name]["DOUBLE_YELLOW"] = capture_point()

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

            "C1": None,

            "C2": None,

            "C3": None
        }

        root.iconify()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture MENU\n\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture C1\n\nPress SPACE"
        )

        signals[name]["C1"] = capture_point()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture C2\n\nPress SPACE"
        )

        signals[name]["C2"] = capture_point()

        messagebox.showinfo(
            "CAPTURE",
            f"{name}\n\nCapture C3\n\nPress SPACE"
        )

        signals[name]["C3"] = capture_point()

        root.deiconify()

    log("ALL SIGNALS RECORDED")

    # =====================================================
    # EMERGENCY CRANK HANDLE
    # =====================================================

    # =====================================================
    # EMERGENCY CRANK HANDLE
    # =====================================================

    crank_handle_points = {}

    total_ch = simpledialog.askinteger(
        "EMERGENCY CRANK HANDLE",
        "How Many Crank Handles?"
    )

    if total_ch is None:
        return

    for i in range(1, total_ch + 1):
        handle = f"CH{i}"

        crank_handle_points[handle] = {}

        crank_handle_points[handle]["BUTTON"] = capture_crank_point(handle, "RED BUTTON")

        crank_handle_points[handle]["IN"] = capture_crank_point(handle, "IN")

        crank_handle_points[handle]["OUT"] = capture_crank_point(handle, "OUT")

        crank_handle_points[handle]["ECH"] = capture_crank_point(handle, "ECH")

        crank_handle_points[handle]["FREE"] = capture_crank_point(handle, "FREE")

    root.deiconify()

    log("CRANK HANDLE COORDINATES CAPTURED")

    # =====================================================
    # LC COORDINATES
    # =====================================================

    # =====================================================
    # LC COORDINATES
    # =====================================================

    lc_points.clear()

    lc_points["BUTTON"] = capture_crank_point(
        "LC",
        "PURPLE BUTTON"
    )

    lc_points["OPEN"] = capture_crank_point(
        "LC",
        "OPEN YELLOW INDICATION"
    )

    lc_points["CLOSE"] = capture_crank_point(
        "LC",
        "CLOSE GREEN INDICATION"
    )

    log("LC COORDINATES CAPTURED")

    # =====================================================
    # POINT INDICATION COORDINATES
    # =====================================================

    total_points = simpledialog.askinteger(
        "POINTS",
        "How many points need to be configured?"
    )

    if total_points is None:
        return

    for i in range(total_points):

        point_no = simpledialog.askstring(
            "POINT",
            f"Enter Point Number {i + 1}\nExample: 50"
        )

        if not point_no:
            return

        point_no = point_no.strip().upper()

        point_indications[point_no] = {}

        point_indications[point_no]["N"] = capture_crank_point(
            f"POINT {point_no}",
            "N GREEN INDICATION"
        )

        point_indications[point_no]["R"] = capture_crank_point(
            f"POINT {point_no}",
            "R YELLOW INDICATION"
        )

    log("LC AND POINT COORDINATES CAPTURED")



    # =====================================================
    # CAPTURE TRACK POINTS
    # =====================================================


    log("SIGNAL AND CRANK HANDLE COORDINATES CAPTURED")

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

            steps = ["MENU", "RED", "YELLOW"]

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

        steps = ["MENU", "YELLOW"]

        current = steps[current_step]

        log(
            f"{current_signal} -> Capture {current}"
        )

    # =====================================================
    # SHUNT
    # =====================================================

    elif sig_type == "SHUNT":

        steps = ["MENU", "C1", "C2", "C3"]

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

    # =====================================================
    # MAIN
    # =====================================================

    if sig_type == "MAIN":

        aspects = info["aspects"]

        if aspects == 2:

            steps = ["MENU", "RED", "YELLOW"]

        elif aspects == 3:

            steps = ["MENU", "RED", "YELLOW", "GREEN"]

        else:

            steps = [
                "MENU",
                "RED",
                "YELLOW",
                "DOUBLE_YELLOW",
                "GREEN"
            ]

    elif sig_type == "CALLING_ON":

        steps = ["MENU", "YELLOW"]

    else:

        steps = ["MENU", "C1", "C2", "C3"]

    current_name = steps[current_step]

    signals[signal][current_name] = [x, y]

    log(
        f"{signal} {current_name} Saved"
    )

    current_step += 1

    if current_step >= len(steps):

        current_step = 0

        capture_index += 1

    next_capture()


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

        "RouteInit_Y"

    ])

    # =====================================================
    # POINTS
    # =====================================================

    ws_point = wb.create_sheet("POINTS")

    ws_point.append([
        "POINT",

        "MENU_X",
        "MENU_Y",

        "NORMAL_X",
        "NORMAL_Y",

        "REVERSE_X",
        "REVERSE_Y",

        "FREE_X",
        "FREE_Y"
    ])

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    ws_ch = wb.create_sheet("CRANK_HANDLES")

    ws_ch.append([
        "CRANK_HANDLE",

        "MENU_X",
        "MENU_Y",

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

        "MENU_X",
        "MENU_Y",

        "IN_X",
        "IN_Y",

        "OUT_X",
        "OUT_Y"
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

            route_init = info.get("ROUTE_INIT") or [None, None]

            ws_call.append([
                sig,
                info["menu"][0],
                info["menu"][1],
                info["YELLOW"][0],
                info["YELLOW"][1],
                route_init[0],
                route_init[1]
            ])

        # =================================================
        # SHUNT
        # =================================================

        elif typ == "SHUNT":

            indicator = info.get("state_indicator") or info.get("C1") or [None, None]
            route_init = info.get("route_init") or info.get("C2") or [None, None]

            ws_shunt.append([
                sig,
                info["menu"][0],
                info["menu"][1],
                indicator[0],
                indicator[1],
                route_init[0],
                route_init[1]
            ])

        # =================================================
        # POINT
        # =================================================

        elif typ == "POINT":

            menu = info.get("menu") or [None, None]

            normal = info.get("normal") or [None, None]

            reverse = info.get("reverse") or [None, None]

            free = info.get("free") or [None, None]

            ws_point.append([

                sig,

                menu[0],

                menu[1],

                normal[0],

                normal[1],

                reverse[0],

                reverse[1],

                free[0],

                free[1]

            ])

        # =================================================
        # CRANK HANDLE
        # =================================================

        elif typ == "CH":

            menu = info.get("menu") or [None, None]

            in_point = info.get("IN") or [None, None]

            out_point = info.get("OUT") or [None, None]

            ech = info.get("ECH") or [None, None]

            free = info.get("FREE") or [None, None]

            ws_ch.append([

                sig,

                menu[0],

                menu[1],

                in_point[0],

                in_point[1],

                out_point[0],

                out_point[1],

                ech[0],

                ech[1],

                free[0],

                free[1]

            ])

        # =================================================
        # LC GATE
        # =================================================

        elif typ == "LC":

            menu = info.get("menu") or [None, None]

            in_point = info.get("IN") or [None, None]

            out_point = info.get("OUT") or [None, None]

            ws_lc.append([

                sig,

                menu[0],

                menu[1],

                in_point[0],

                in_point[1],

                out_point[0],

                out_point[1]

            ])



    # =====================================================
    # WRITE ROUTE POINTS
    # =====================================================


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
    global config_file_path
    global crank_handle_points


    if not lock_routes_data:

        messagebox.showwarning(
            "WARNING",
            "Please load TOC file first."
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

    crank_handle_points = {}

    lc_points.clear()

    point_indications.clear()

    wb = load_workbook(config_file_path)

    # =====================================================
    # VALIDATE VISUAL CONFIGURATION FILE
    # =====================================================

    required_sheets = [
        "MAIN_SIGNALS",
        "CALLING_ON",
        "SHUNT",
        "POINTS",
        "CRANK_HANDLES",
        "LC_GATES"
    ]

    missing_sheets = [
        sheet for sheet in required_sheets
        if sheet not in wb.sheetnames
    ]

    if missing_sheets:
        messagebox.showerror(
            "INVALID VISUAL CONFIGURATION",
            "The selected file is not a Visual Inspection configuration file.\n\n"
            f"Missing sheets:\n{', '.join(missing_sheets)}"
        )
        return
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

        route_init = None
        if len(row) > 6 and row[5] is not None:
            route_init = [row[5], row[6]]

        signals[signal_name] = {
            "type": "CALLING_ON",
            "menu": [row[1], row[2]],
            "YELLOW": [row[3], row[4]],
            "ROUTE_INIT": route_init
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

        # New TL format: Signal, Menu_X, Menu_Y, Indicator_X, Indicator_Y,
        # RouteInit_X, RouteInit_Y.  Older files with C1/C2/C3 are still
        # accepted so existing configurations are not lost.
        indicator = [row[3], row[4]] if len(row) > 4 and row[3] is not None else None
        route_init = [row[5], row[6]] if len(row) > 6 and row[5] is not None else None

        signals[signal_name] = {
            "type": "SHUNT",
            "menu": [row[1], row[2]],
            "state_indicator": indicator,
            "route_init": route_init,
            "initial_snapshot": None
        }

    # =====================================================
    # POINTS
    # =====================================================

    # =====================================================
    # POINTS
    # =====================================================

    if "POINTS" in wb.sheetnames:

        ws = wb["POINTS"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row[0]:
                continue

            point_name = (
                str(row[0])
                .replace(".0", "")
                .strip()
                .upper()
            )

            signals[point_name] = {

                "type": "POINT",

                "menu": [
                    row[1],
                    row[2]
                ],

                "normal": [
                    row[3],
                    row[4]
                ],

                "reverse": [
                    row[5],
                    row[6]
                ],

                "free": [
                    row[7],
                    row[8]
                ]
            }

    # =====================================================
    # LOAD CRANK HANDLE
    # =====================================================

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    if "CRANK_HANDLES" in wb.sheetnames:

        ws = wb["CRANK_HANDLES"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row[0]:
                continue

            ch_name = (
                str(row[0])
                .replace(".0", "")
                .strip()
                .upper()
            )

            signals[ch_name] = {

                "type": "CH",

                "menu": [
                    row[1],
                    row[2]
                ],

                "IN": [
                    row[3],
                    row[4]
                ],

                "OUT": [
                    row[5],
                    row[6]
                ],

                "ECH": [
                    row[7],
                    row[8]
                ],

                "FREE": [
                    row[9],
                    row[10]
                ]
            }

    # =====================================================
    # LOAD LC
    # =====================================================

    # =====================================================
    # LC GATES
    # =====================================================

    if "LC_GATES" in wb.sheetnames:

        ws = wb["LC_GATES"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row[0]:
                continue

            lc_name = (
                str(row[0])
                .replace(".0", "")
                .strip()
                .upper()
            )

            signals[lc_name] = {

                "type": "LC",

                "menu": [
                    row[1],
                    row[2]
                ],

                "IN": [
                    row[3],
                    row[4]
                ],

                "OUT": [
                    row[5],
                    row[6]
                ]
            }

    rebuild_point_indications()
    rebuild_crank_handle_points()

    # =====================================================
    # POINT CONFIGURATION DEBUG
    # =====================================================

    log("=" * 60)
    log("CONFIGURATION CHECK")

    log(
        f"TOTAL SIGNALS : {len(signals)}"
    )

    total_points = sum(
        1
        for info in signals.values()
        if info.get("type") == "POINT"
    )

    log(
        f"TOTAL POINTS IN SIGNAL CONFIG : {total_points}"
    )

    log(
        f"CONFIGURED POINT INDICATIONS : "
        f"{list(point_indications.keys())}"
    )

    for point_no, info in signals.items():

        if info.get("type") == "POINT":
            log(
                f"POINT {point_no} : "
                f"N={info.get('normal')} "
                f"R={info.get('reverse')} "
                f"FREE={info.get('free')}"
            )

    log("=" * 60)

    log("SIGNAL CONFIG LOADED")

def load_universal_coordinates():
    """Loads coordinates from the multi-sheet Universal Yard Coordinate
    file (the same file every other suite program reads from), instead
    of this program's own native config format (MAIN_SIGNALS /
    CALLING_ON / SHUNT / POINTS / CRANK_HANDLES / LC_GATES, which
    load_config() validates strictly and rejects if any is missing -
    this loader does not apply that same requirement, since the
    universal file uses different sheet names). Populates the same
    `signals` dict that load_config() populates (MAIN / CAL / SHUNT /
    POINT / CH / LC types), then rebuilds point_indications and
    crank_handle_points from it, same as load_config() does."""

    global signals
    global config_file_path
    global crank_handle_points

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
    crank_handle_points = {}
    lc_points.clear()
    point_indications.clear()

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
                snap_val = cell(row, hi.get("SNAPSHOT_PATH"))
                snapshot_path = str(snap_val).strip() if snap_val else None
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "state_indicator": indicator,
                    "route_init": route_init,
                    "initial_snapshot": snapshot_path,
                    "C1": indicator,
                    "C2": route_init,
                    "C3": route_init
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
                name = str(cell(row, name_col)).replace(".0", "").strip().upper()
                signals[name] = {
                    "type": "LC",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "IN": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "OUT": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y")))
                }
                loaded_lc += 1

    rebuild_point_indications()
    rebuild_crank_handle_points()

    log("=" * 60)
    log("CONFIGURATION CHECK (UNIVERSAL)")
    log(f"TOTAL SIGNALS : {len(signals)}")
    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_cal} CALLING-ON, {loaded_shunt} SHUNT, "
        f"{loaded_point} POINT, {loaded_ch} CH, {loaded_lc} LC"
    )
    log("=" * 60)
    log("SIGNAL CONFIG LOADED (UNIVERSAL)")


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

def load_lock_routes():

    global lock_routes_data

    file = filedialog.askopenfilename(
        title="Select Visual Inspection TOC",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)
    # FIX: read the sheet named TOC (master workbook has several sheets)
    ws = wb["TOC"] if "TOC" in wb.sheetnames else wb.active
    lock_routes_data = []

    # ==========================================
    # READ HEADER ROW
    # ==========================================

    headers = []

    for cell in ws[1]:
        if cell.value is None:
            headers.append("")
        else:
            headers.append(
                str(cell.value).strip().upper()
            )

    # ==========================================
    # FIND COLUMNS
    # ==========================================

    signal_col = None
    route_col = None
    crank_col = None
    locks_detects_col = None
    lcp_col = None
    controlled_tracks_col = None
    approach_track_col = None

    for i, h in enumerate(headers):

        if h == "SIGNAL":
            signal_col = i

        elif h == "ROUTE":
            route_col = i

        # FIX: ignore CRANK_HANDLE_BITS (CH1KLNR...) - it was overwriting
        # the real CRANK_HANDLE column.
        elif "CRANK" in h and "BITS" not in h and crank_col is None:
            crank_col = i

        elif h == "CONTROLLED_BY_TRACKS":
            controlled_tracks_col = i

        elif h == "LOCKS_DETECTS_POINTS":
            locks_detects_col = i

        elif h == "44_LCP":
            lcp_col = i

        elif h in (
            "APPROACH_TRACK",
            "APPROACH TRACK",
            "APPROACH_TRACKS"
        ):
            approach_track_col = i

    # ==========================================
    # REQUIRED COLUMNS
    # ==========================================

    if signal_col is None:
        messagebox.showerror(
            "ERROR", "SIGNAL column not found."
        )
        return

    if route_col is None:
        messagebox.showerror(
            "ERROR", "ROUTE column not found."
        )
        return

    if crank_col is None:
        messagebox.showerror(
            "ERROR", "CRANK HANDLE column not found."
        )
        return

    if locks_detects_col is None:
        messagebox.showerror(
            "ERROR",
            "LOCKS_DETECTS_POINTS column not found."
        )
        return

    if lcp_col is None:
        messagebox.showerror(
            "ERROR", "44_LCP column not found."
        )
        return

    # FIX: TOC keeps the Calling-On approach track (1C -> 1CXTPR) in
    # the "Track" column - use it when no Approach_Track column exists.
    if approach_track_col is None and "TRACK" in headers:
        approach_track_col = headers.index("TRACK")

    if approach_track_col is not None:
        log(
            "APPROACH_TRACK COLUMN FOUND : "
            f"{headers[approach_track_col]}"
        )
    else:
        log(
            "APPROACH_TRACK COLUMN NOT FOUND - "
            "Calling-ON routes cannot be prepared."
        )

    # ==========================================
    # READ DATA
    # ==========================================

    for row in ws.iter_rows(
            min_row=2,
            values_only=True):

        if row is None:
            continue

        if (
            signal_col >= len(row)
            or row[signal_col] is None
        ):
            continue

        if (
            route_col >= len(row)
            or row[route_col] is None
        ):
            continue

        signal = str(
            row[signal_col]
        ).replace(".0", "").strip().upper()

        route = str(
            row[route_col]
        ).replace(".0", "").strip().upper()

        if signal == "" or route == "":
            continue

        crank_handle = ""

        if (
            crank_col is not None
            and crank_col < len(row)
            and row[crank_col] is not None
        ):
            crank_handle = str(
                row[crank_col]
            ).replace(".0", "").strip().upper()

        locks_detects_points = ""

        if (
            locks_detects_col is not None
            and locks_detects_col < len(row)
            and row[locks_detects_col] is not None
        ):
            locks_detects_points = str(
                row[locks_detects_col]
            ).replace(".0", "").strip().upper()

        lcp_value = ""

        if (
            lcp_col is not None
            and lcp_col < len(row)
            and row[lcp_col] is not None
        ):
            lcp_value = str(
                row[lcp_col]
            ).replace(".0", "").strip()

        controlled_by_tracks = ""

        if (
            controlled_tracks_col is not None
            and controlled_tracks_col < len(row)
            and row[controlled_tracks_col] is not None
        ):
            controlled_by_tracks = str(
                row[controlled_tracks_col]
            ).replace(".0", "").strip().upper()

        approach_track = ""

        if (
            approach_track_col is not None
            and approach_track_col < len(row)
            and row[approach_track_col] is not None
        ):
            approach_track = str(
                row[approach_track_col]
            ).replace(".0", "").strip().upper()

        lock_routes_data.append({
            "signal": signal,
            "route": route,
            "controlled_by_tracks":
                controlled_by_tracks,
            "approach_track":
                approach_track,
            "crank_handle":
                crank_handle,
            "locks_detects_points":
                locks_detects_points,
            "44_lcp":
                lcp_value
        })

        log(
            f"{signal} -> {route} "
            f"-> APPROACH_TRACK="
            f"{approach_track or '[EMPTY]'} "
            f"-> {locks_detects_points}"
        )

    refresh_table()
    log("VISUAL INSPECTION TOC LOADED")


# =========================================================
# TABLE
# =========================================================

def refresh_table():

    tree.delete(*tree.get_children())

    selected_tab = notebook.tab(
        notebook.select(),
        "text"
    )

    count = 1

    for row in lock_routes_data:

        signal = str(
            row["signal"]
        ).strip().upper()

        # Dashboard classification:
        # 1C  -> CALLING-ON
        # 9SH -> SHUNT
        # 1   -> MAIN
        is_calling_on = signal.endswith("C")
        is_shunt = signal.endswith("SH") or signal.startswith("SH")

        if selected_tab == "MAIN SIGNALS":
            if is_calling_on or is_shunt:
                continue

        elif selected_tab == "CALLING-ON SIGNALS":
            if not is_calling_on:
                continue

        elif selected_tab == "SHUNT SIGNALS":
            if not is_shunt:
                continue

        tree.insert(
            "",
            "end",
            values=(
                count,
                row["signal"],
                row["route"],
                row.get("controlled_by_tracks", ""),
                row.get("approach_track", ""),
                row["crank_handle"],
                row["locks_detects_points"],
                row["44_lcp"]
            )
        )

        count += 1


def update_point_visual(signal, route, required_point, detected, status):

    key = (signal, route, required_point)

    def update_ui():

        item_id = point_visual_items.get(key)

        if not item_id:
            return

        tree.item(
            item_id,
            values=(
                tree.item(item_id, "values")[0],
                signal,
                route,
                required_point,
                detected,
                status
            ),
            tags=(status,)
        )

    root.after(0, update_ui)

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
# SET ROUTE WITH RETRY
# =========================================================

def set_route(signal, route, retries=3):

    for attempt in range(1, retries + 1):

        log(f"SETTING ROUTE (Attempt {attempt}/{retries})")

        if not click(signals[signal]["menu"]):
            continue

        time.sleep(1.5)

        if click_menu_item(route.replace("-", "_")):
            log("ROUTE COMMAND SENT")
            return True

        time.sleep(1)

    log("FAILED TO SET ROUTE")
    return False

# =========================================================
# AVG COLOR
# =========================================================

def get_avg_color(x, y, screenshot):

    pixels = []

    for dx in range(-2, 4):
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



def is_red_color(r, g, b):

    return (
        r >= 140 and
        r > g + 40 and
        r > b + 40
    )


def is_yellow_color(r, g, b):
    return r > 170 and g > 170 and b < 120


def is_blue_color(r, g, b):
    return b > 170 and r < 120 and g < 170


def is_black_color(r, g, b):
    return r < 60 and g < 60 and b < 60

# =========================================================
# YELLOW DETECT
# =========================================================

def is_yellow(point):

    if point is None:
        return False

    screenshot = pyautogui.screenshot()

    x, y = point

    r, g, b = get_avg_color(x, y, screenshot)

    log(f"RGB : {r},{g},{b}")

    return (
            r >= 140 and
            g >= 140 and
            abs(r - g) <= 70 and
            b < 160
    )

def is_green(point):

    if point is None:
        return False

    screenshot = pyautogui.screenshot()

    x, y = point

    r, g, b = get_avg_color(x, y, screenshot)

    log(f"GREEN RGB : {r},{g},{b}")

    return (
        g >= 90 and
        g > r + 20 and
        g > b + 20
    )


def is_steady_yellow(point, samples=6):

    yellow = 0

    for _ in range(samples):

        if is_yellow(point):

            yellow += 1

        time.sleep(0.30)

    log(
        f"Yellow Samples : {yellow}/{samples}"
    )

    return yellow >= 4


def is_steady_green(point, samples=6):

    green = 0

    for _ in range(samples):

        if is_green(point):
            green += 1

        time.sleep(0.30)

    log(f"Green Samples : {green}/{samples}")

    return green >= 4


# =========================================================
# STEADY RED
# =========================================================

def is_steady_red(point, samples=6):

    if point is None:
        log("RED COORDINATE IS NONE")
        return False

    red = 0

    x, y = point

    for _ in range(samples):

        screenshot = pyautogui.screenshot()

        r, g, b = get_avg_color(
            x,
            y,
            screenshot
        )

        log(f"RED RGB : {r},{g},{b}")

        if is_red_color(r, g, b):
            red += 1

        time.sleep(0.30)

    log(f"Red Samples : {red}/{samples}")

    return red >= 4

def is_blinking_yellow(point, duration=5):

    """
    Returns True only if the indication
    repeatedly changes between
    Yellow and OFF.
    """

    yellow_seen = 0
    off_seen = 0

    start = time.time()

    while time.time() - start < duration:

        if is_yellow(point):
            yellow_seen += 1
        else:
            off_seen += 1

        time.sleep(0.25)

    log(
        f"OPEN Blink Check : Yellow={yellow_seen}  OFF={off_seen}"
    )

    return yellow_seen > 0 and off_seen > 0

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
# WAIT FOR ROUTE TO SET
# =========================================================

def wait_for_route_set(signal, timeout=30):

    log("WAITING FOR ROUTE TO SET")

    start = time.time()

    while time.time() - start < timeout:

        if validate_route_signal(signal):

            log("ROUTE SET SUCCESSFULLY")

            return True

        time.sleep(1)

    log("ROUTE SET TIMEOUT")

    return False

def is_shunt_route_initiator_yellow(signal, samples=6):
    """
    SHUNT route verification uses ONLY the captured Route Initiator.
    C1/C2/C3 are intentionally not used.
    """
    if signal not in signals:
        log(f"{signal} NOT CONFIGURED")
        return False

    info = signals[signal]
    route_init_point = info.get("route_init")

    if not route_init_point:
        log(f"{signal} SHUNT ROUTE INITIATOR COORDINATE NOT CONFIGURED")
        return False

    log(f"SHUNT -> ROUTE INITIATOR = {route_init_point}")

    # The Route Initiator is the only route-set indication checked.
    # Retry because the yellow lamp may take time to appear.
    for attempt in range(6):
        log(f"SHUNT ROUTE INITIATOR CHECK {attempt + 1}/6")

        if is_steady_yellow(route_init_point, samples=samples):
            log("SHUNT ROUTE INITIATOR YELLOW VERIFIED")
            return True

        if attempt < 5:
            time.sleep(1)

    log("SHUNT ROUTE INITIATOR NOT YELLOW")
    return False


def validate_route_signal(signal):

    if signal not in signals:

        log(f"{signal} NOT CONFIGURED")

        return False

    info = signals[signal]

    signal_type = info["type"]

    log("=" * 60)
    log(f"CHECKING {signal} ({signal_type})")

    # ===========================================
    # MAIN SIGNAL
    # ===========================================

    # ===========================================
    # MAIN SIGNAL
    # ===========================================

    if signal_type == "MAIN":

        aspects = int(info["aspects"])

        # ---------------------------------------
        # 2 Aspect (RED -> GREEN)
        # ---------------------------------------
        if aspects == 2:

            log("2 ASPECT MAIN SIGNAL")

            # First check the Yellow coordinate (if available)
            if info["YELLOW"] is not None:
                if is_steady_yellow(info["YELLOW"]):
                    log("YELLOW VERIFIED")
                    return True

            # If no Yellow coordinate is configured,
            # use the Green coordinate as the proceed indication.
            if info["GREEN"] is not None:

                # Actual lamp may appear Green
                if is_steady_green(info["GREEN"]):
                    log("GREEN VERIFIED")
                    return True

                # Or, if your mentor wants to accept a yellow-looking proceed aspect
                if is_steady_yellow(info["GREEN"]):
                    log("GREEN COORDINATE SHOWING YELLOW")
                    return True

            log("PROCEED ASPECT NOT VERIFIED")
            return False
        # ---------------------------------------
        # 3 Aspect (RED -> YELLOW -> GREEN)
        # ---------------------------------------
        elif aspects == 3:

            log("3 ASPECT MAIN SIGNAL")

            if is_steady_yellow(info["YELLOW"]):
                log("YELLOW VERIFIED")

                return True

            if is_steady_green(info["GREEN"]):
                log("GREEN VERIFIED")

                return True

            log("SIGNAL NOT YELLOW/GREEN")

            return False

        # ---------------------------------------
        # 4 Aspect
        # ---------------------------------------
        elif aspects == 4:

            log("4 ASPECT MAIN SIGNAL")

            if info["DOUBLE_YELLOW"] is not None:

                if is_steady_yellow(info["DOUBLE_YELLOW"]):
                    log("DOUBLE YELLOW VERIFIED")

                    return True

            if is_steady_yellow(info["YELLOW"]):
                log("YELLOW VERIFIED")

                return True

            if is_steady_green(info["GREEN"]):
                log("GREEN VERIFIED")

                return True

            log("NO VALID ASPECT FOUND")

            return False

    # ===========================================
    # CALLING ON
    # ===========================================

    elif signal_type == "CALLING_ON":

        yellow_point = info.get("YELLOW")
        route_init_point = info.get("ROUTE_INIT")

        if not yellow_point:
            log(
                "CALLING ON YELLOW COORDINATE "
                "NOT CONFIGURED"
            )
            return False

        if not route_init_point:
            log(
                "CALLING ON ROUTE INITIATOR "
                "COORDINATE NOT CONFIGURED"
            )
            return False

        # Calling-ON route is considered set only when
        # BOTH the yellow lamp and the route initiator
        # become steady yellow.
        if not is_steady_yellow(yellow_point):
            log(
                "CALLING ON YELLOW LAMP "
                "NOT YELLOW YET"
            )
            return False

        if not is_steady_yellow(route_init_point):
            log(
                "CALLING ON ROUTE INITIATOR "
                "NOT YELLOW YET"
            )
            return False

        log(
            "CALLING ON YELLOW + "
            "ROUTE INITIATOR VERIFIED"
        )
        return True

    # ===========================================
    # SHUNT
    # ===========================================

    elif signal_type == "SHUNT":

        # =====================================================
        # SHUNT ROUTE SET VERIFICATION
        # =====================================================
        # DO NOT check C1, C2 or C3.
        # ONLY the captured Route Initiator is used.
        # =====================================================

        log("SHUNT -> VERIFYING ROUTE INITIATOR ONLY")

        if is_shunt_route_initiator_yellow(signal, samples=6):
            log("SHUNT ROUTE INITIATOR YELLOW VERIFIED")
            log("SHUNT ROUTE SET SUCCESSFULLY")
            return True

        log("SHUNT ROUTE INITIATOR NOT YELLOW")
        return False

    log("UNKNOWN SIGNAL TYPE")

    return False

# =========================================================
# VERIFY SIGNAL RESET
# =========================================================

def verify_signal_reset(signal):

    if signal not in signals:
        return False

    info = signals[signal]

    signal_type = info["type"]

    # ---------------------------------------
    # MAIN SIGNAL
    # ---------------------------------------

    if signal_type == "MAIN":

        red_point = info.get("RED")

        log(f"RED COORDINATE : {red_point}")

        if not red_point:
            log(f"{signal}: RED COORDINATE NOT CONFIGURED")
            return False

        # Read the actual pixel
        # Read the average color around RED coordinate
        screenshot = pyautogui.screenshot()

        x, y = red_point

        r, g, b = get_avg_color(
            x,
            y,
            screenshot
        )

        log(
            f"RED RGB : "
            f"{r},{g},{b}"
        )

        if is_steady_red(red_point):
            log("MAIN SIGNAL RETURNED TO RED")

            return True

        log("MAIN SIGNAL DID NOT RETURN TO RED")

        return False

    # ---------------------------------------
    # CALLING ON
    # ---------------------------------------

    elif signal_type == "CALLING_ON":

        yellow_point = info.get("YELLOW")
        route_init_point = info.get("ROUTE_INIT")

        yellow_active = (
            is_yellow(yellow_point)
            if yellow_point
            else False
        )

        route_init_active = (
            is_yellow(route_init_point)
            if route_init_point
            else False
        )

        if not yellow_active and not route_init_active:
            log("CALLING ON RESET VERIFIED")
            return True

        log(
            "CALLING ON STILL ACTIVE "
            f"(YELLOW={yellow_active}, "
            f"ROUTE_INIT={route_init_active})"
        )
        return False

    # ---------------------------------------
    # SHUNT
    # ---------------------------------------

    elif signal_type == "SHUNT":

        # After Signal Cancel + Route Release, the Route Initiator
        # must return to its non-yellow/reset state.
        route_init_point = info.get("route_init")

        if not route_init_point:
            log("SHUNT ROUTE INITIATOR COORDINATE NOT CONFIGURED")
            return False

        log(f"SHUNT RESET -> ROUTE INITIATOR : {route_init_point}")

        for _ in range(6):
            if not is_yellow(route_init_point):
                log("SHUNT ROUTE INITIATOR RESET VERIFIED")
                return True
            time.sleep(0.5)

        log("SHUNT ROUTE INITIATOR STILL YELLOW")
        return False

    return False

# =========================================================
# PARSE SIGNAL NAME
# =========================================================

# =========================================================
# DROP TRACK
# =========================================================

# =========================================================
# FIND AND CLICK UI
# =========================================================



def find_and_click(name, target=None, control_type=None):
    """
    Find and click a UI element.

    Supported forms:
        find_and_click(element_name)
        find_and_click(element_name, control_type)
        find_and_click(window_title, element_name, control_type)
    """
    try:
        desktop = Desktop(backend="uia")

        # Backward-compatible two-argument form.
        if control_type is None and target in (
                "Button", "ListItem", "MenuItem", "Window",
                "Pane", "Text", "Edit", "ComboBox"):
            control_type = target
            target = None

        for win in desktop.windows():
            try:
                win_title = win.window_text().strip().upper()

                if target is not None:
                    if name.upper() not in win_title:
                        continue
                    search_name = target
                    descendants = win.descendants()
                else:
                    search_name = name
                    descendants = win.descendants()

                for item in descendants:
                    try:
                        item_text = item.window_text().strip().upper()

                        if item_text != search_name.upper():
                            continue

                        if control_type:
                            current_type = str(
                                item.element_info.control_type
                            )
                            if control_type.upper() not in current_type.upper():
                                continue

                        rect = item.rectangle()
                        x = (rect.left + rect.right) // 2
                        y = (rect.top + rect.bottom) // 2

                        pyautogui.click(x, y)
                        log(f"Clicked : {search_name}")
                        return True

                    except Exception:
                        pass

            except Exception:
                pass

    except Exception as e:
        log(f"FIND_AND_CLICK ERROR: {e}")

    return False


# =========================================================
# OPEN BIT CHART
# =========================================================


def find_and_click_50051(window_title, control_name, control_type=None):
    """
    Find a control inside the requested Station 50051 window.

    Important for Calling-On:
    1CXTPR may be exposed by the application as Text/Pane/ListItem
    depending on the UI state, so the track lookup does NOT require
    ListItem. Transmit/Cancel still use Button when requested.
    """

    wanted_window = str(
        window_title or ""
    ).strip().upper()

    wanted_control = str(
        control_name or ""
    ).strip().upper()

    desktop = Desktop(backend="uia")

    # Prefer windows whose title identifies Station 50051.
    candidate_windows = []

    for window in desktop.windows():

        try:
            title = str(
                window.window_text() or ""
            ).strip().upper()

            if wanted_window in title:
                candidate_windows.append(window)

        except Exception:
            pass

    # If UIA does not expose the title, fall back to all windows.
    if not candidate_windows:
        candidate_windows = desktop.windows()

    for window in candidate_windows:

        try:

            for item in window.descendants():

                try:

                    text = str(
                        item.window_text() or ""
                    ).strip().upper()

                    # Exact match is important for 1CXTPR.
                    if text != wanted_control:
                        continue

                    if control_type:

                        current_type = str(
                            item.element_info.control_type
                        ).upper()

                        if (
                            str(control_type).upper()
                            not in current_type
                        ):
                            continue

                    log(
                        f"50051 UI FOUND : "
                        f"{item.window_text()} | "
                        f"TYPE : "
                        f"{item.element_info.control_type}"
                    )

                    # Use UIA click, not a guessed screen coordinate.
                    item.click_input()

                    log(
                        f"50051 {control_name} CLICKED"
                    )

                    return True

                except Exception:
                    pass

        except Exception:
            pass

    log(
        f"50051 {control_name} NOT FOUND "
        f"IN {window_title}"
    )

    return False


def open_bit_chart():

    """
    Calling-On Station 50051 opening:

        CTRL+B
        -> CLICK 50051 ONCE
        -> ENTER
        -> wait until Station 50051 is actually ready

    We detect readiness using the Station 50051 controls as well as
    the title because UIA does not always expose the window title
    consistently.
    """

    pyautogui.click(500, 500)
    time.sleep(1)

    pyautogui.hotkey("ctrl", "b")
    log("CTRL+B Pressed")
    time.sleep(3)

    desktop = Desktop(backend="uia")
    selection_windows = []

    for win in desktop.windows():

        try:

            title = str(
                win.window_text() or ""
            ).strip().upper()

            if (
                "STATIONS - SELECT STATION" in title
                or "SELECT STATION" in title
            ):
                selection_windows.append(win)

        except Exception:
            pass

    if not selection_windows:
        selection_windows = desktop.windows()

    clicked_50051 = False

    for win in selection_windows:

        try:

            for item in win.descendants():

                try:

                    text = str(
                        item.window_text() or ""
                    ).strip().upper()

                    if text != "50051":
                        continue

                    log(
                        "50051 FOUND IN STATION "
                        "SELECTION WINDOW"
                    )

                    try:
                        item.click_input()
                    except Exception:
                        rect = item.rectangle()

                        x = (
                            rect.left + rect.right
                        ) // 2

                        y = (
                            rect.top + rect.bottom
                        ) // 2

                        pyautogui.click(x, y)

                    clicked_50051 = True
                    log("50051 CLICKED")

                    time.sleep(0.5)

                    # Confirm the selected station with ENTER.
                    pyautogui.press("enter")
                    log("50051 ENTER PRESSED")

                    break

                except Exception:
                    pass

            if clicked_50051:
                break

        except Exception:
            pass

    if not clicked_50051:

        log(
            "50051 NOT FOUND IN STATION "
            "SELECTION WINDOW"
        )

        return False

    # ---------------------------------------------------------
    # Wait for the actual Station 50051 controls.
    # This is more reliable than checking only the window title.
    # ---------------------------------------------------------
    end_time = time.time() + 12

    while time.time() < end_time:

        desktop = Desktop(backend="uia")

        station_ready = False

        for win in desktop.windows():

            try:

                title = str(
                    win.window_text() or ""
                ).strip().upper()

                # Title-based detection.
                if "STATION 50051" in title:
                    station_ready = True
                    break

                # Control-based detection.
                has_transmit = False
                has_cancel = False

                for item in win.descendants():

                    try:

                        text = str(
                            item.window_text() or ""
                        ).strip().upper()

                        if text == "TRANSMIT":
                            has_transmit = True

                        elif text == "CANCEL":
                            has_cancel = True

                        if has_transmit and has_cancel:
                            station_ready = True
                            break

                    except Exception:
                        pass

                if station_ready:
                    break

            except Exception:
                pass

        if station_ready:

            log(
                "STATION 50051 WINDOW OPENED / "
                "CONTROLS DETECTED"
            )

            time.sleep(1)
            return True

        time.sleep(0.25)

    log(
        "50051 CLICK + ENTER COMPLETED, "
        "BUT STATION 50051 CONTROLS WERE NOT DETECTED"
    )

    return False


def select_50051_approach_track(track):

    """
    Select the exact Calling-On Approach_Track in Station 50051.

    Example:
        1CXTPR

    Do NOT require ListItem. The application can expose this entry
    as Text, Pane, ListItem, or another UIA type.
    """

    track = str(
        track or ""
    ).strip().upper()

    if not track:

        log(
            "CALLING-ON APPROACH TRACK "
            "IS EMPTY"
        )

        return False

    log(
        f"CALLING-ON -> SEARCHING EXACT "
        f"50051 APPROACH TRACK : {track}"
    )

    # First try exact text without any control-type restriction.
    if find_and_click_50051(
        "Station 50051",
        track
    ):

        log(
            f"CALLING-ON -> {track} "
            "APPROACH TRACK CLICKED"
        )

        time.sleep(1)

        return True

    # Diagnostic fallback: inspect Station 50051 and print controls
    # containing the requested track.
    desktop = Desktop(backend="uia")

    for win in desktop.windows():

        try:

            title = str(
                win.window_text() or ""
            ).strip().upper()

            if "STATION 50051" not in title:
                continue

            for item in win.descendants():

                try:

                    text = str(
                        item.window_text() or ""
                    ).strip()

                    if track in text.upper():

                        rect = item.rectangle()

                        log(
                            f"CALLING-ON -> TRACK TEXT FOUND "
                            f"AS '{text}' "
                            f"TYPE={item.element_info.control_type} "
                            f"RECT=({rect.left},{rect.top},"
                            f"{rect.right},{rect.bottom})"
                        )

                        item.click_input()

                        log(
                            f"CALLING-ON -> {track} "
                            "APPROACH TRACK CLICKED "
                            "(PARTIAL-TEXT FALLBACK)"
                        )

                        time.sleep(1)

                        return True

                except Exception:
                    pass

        except Exception:
            pass

    log(
        f"CALLING-ON -> {track} "
        "APPROACH TRACK NOT FOUND IN STATION 50051"
    )

    return False


def close_50051_station():
    """Close Station 50051 after a Calling-ON transmit."""
    if not click_station_cancel():
        log("FAILED TO CLOSE STATION 50051")
        return False

    log("STATION 50051 CLOSED")
    time.sleep(2)
    return True


def prepare_calling_on_approach_track(approach_track):
    """
    Calling-ON preparation:
        CTRL+B
        -> CLICK 50051 ONCE + ENTER
        -> SELECT APPROACH_TRACK
        -> TRANSMIT
        -> CLOSE 50051
    """
    approach_track = str(
        approach_track or ""
    ).strip().upper()

    if not approach_track:
        log(
            "CALLING-ON APPROACH TRACK "
            "NOT CONFIGURED IN TOC"
        )
        return False

    log("=" * 70)
    log(
        f"CALLING-ON 50051 PREPARATION : "
        f"{approach_track}"
    )

    if not open_station_bits():
        log("CALLING-ON FAILED TO OPEN 50051")
        return False

    if not select_50051_approach_track(
            approach_track):
        return False

    if not click_station_transmit():
        log(
            "CALLING-ON APPROACH TRACK "
            "TRANSMIT FAILED"
        )
        return False

    log(
        f"CALLING-ON APPROACH TRACK "
        f"TRANSMITTED : {approach_track}"
    )

    if not close_50051_station():
        return False

    log("CALLING-ON 50051 PREPARATION COMPLETED")
    log("=" * 70)
    return True


def complete_calling_on_approach_track(approach_track):
    """
    Final Calling-ON step:
        CTRL+B
        -> CLICK 50051 ONCE + ENTER
        -> SELECT SAME APPROACH_TRACK
        -> TRANSMIT
        -> CLOSE 50051
    """
    approach_track = str(
        approach_track or ""
    ).strip().upper()

    if not approach_track:
        return False

    log("=" * 70)
    log(
        f"CALLING-ON FINAL 50051 TRANSMIT : "
        f"{approach_track}"
    )

    if not open_station_bits():
        log("CALLING-ON FINAL 50051 OPEN FAILED")
        return False

    if not select_50051_approach_track(
            approach_track):
        return False

    if not click_station_transmit():
        log(
            "CALLING-ON FINAL APPROACH TRACK "
            "TRANSMIT FAILED"
        )
        return False

    log(
        f"CALLING-ON FINAL APPROACH TRACK "
        f"TRANSMIT SUCCESS : {approach_track}"
    )

    if not close_50051_station():
        return False

    log("CALLING-ON FINAL 50051 STEP COMPLETED")
    log("=" * 70)
    return True


def click_crank_button(handle):

    handle = handle.strip().upper()

    if handle not in crank_handle_points:

        log(f"{handle} NOT FOUND")

        return False

    click(crank_handle_points[handle]["BUTTON"])

    log(f"{handle} BUTTON CLICKED")

    time.sleep(1)

    return True

def click_crank_transmit():

    found = click_menu_item("Transmit")

    if not found:

        log("TRANSMIT MENU NOT FOUND")

        return False

    log("TRANSMIT CLICKED")

    time.sleep(3)

    return True

def click_crank_receive():

    found = click_menu_item("Receive")

    if not found:

        log("RECEIVE MENU NOT FOUND")

        return False

    log("RECEIVE CLICKED")

    time.sleep(3)

    return True

# =========================================================
# VERIFY IN YELLOW
# =========================================================

def is_in_yellow(handle):

    handle = handle.strip().upper()

    if handle not in crank_handle_points:
        return False

    point = crank_handle_points[handle]["IN"]

    return is_yellow(point)

def is_out_blinking(handle):

    handle = handle.strip().upper()

    if handle not in crank_handle_points:
        return False

    point = crank_handle_points[handle]["OUT"]

    red_count = 0
    off_count = 0

    start = time.time()

    while time.time() - start < 5:

        screenshot = pyautogui.screenshot()

        r, g, b = get_avg_color(
            point[0],
            point[1],
            screenshot
        )

        log(f"{handle} OUT RGB = ({r}, {g}, {b})")

        if is_red_color(r, g, b):
            red_count += 1
        else:
            off_count += 1

        time.sleep(0.25)

    log(f"{handle} RED = {red_count}  OFF = {off_count}")

    return red_count > 0 and off_count > 0

def is_ech_blinking(handle):

    handle = handle.strip().upper()

    if handle not in crank_handle_points:
        return False

    point = crank_handle_points[handle]["ECH"]

    yellow = 0
    black = 0

    start = time.time()

    while time.time() - start < 5:

        screenshot = pyautogui.screenshot()

        r, g, b = get_avg_color(
            point[0],
            point[1],
            screenshot
        )

        log(f"{handle} ECH RGB = {r},{g},{b}")

        if is_yellow_color(r, g, b):
            yellow += 1
        else:
            black += 1

        time.sleep(0.25)

    log(f"{handle} Yellow={yellow}  Dark={black}")

    return yellow > 0 and black > 0

def is_out_off(handle):

    handle = handle.strip().upper()

    point = crank_handle_points[handle]["OUT"]

    screenshot = pyautogui.screenshot()

    r, g, b = get_avg_color(
        point[0],
        point[1],
        screenshot
    )

    log(f"{handle} OUT RGB = {r},{g},{b}")

    return is_black_color(r, g, b)



def open_station_bits():

    return open_bit_chart()

def select_crank_handle(handle):

    handle = handle.strip().upper()

    item = handle + "KLNR"

    if not find_and_click(
            "Station 50051",
            item,
            "ListItem"):

        log(f"{item} NOT FOUND")

        return False

    log(f"{item} CLICKED")

    time.sleep(1)

    return True


def click_station_transmit():
    """
    Click Transmit using the same 50051 UIA method as Button Block.
    """
    if not find_and_click_50051(
        "Station 50051",
        "Transmit",
        "Button"
    ):
        log("50051 TRANSMIT BUTTON NOT FOUND")
        return False

    log("50051 TRANSMIT CLICKED")
    time.sleep(3)
    return True


def click_station_cancel():
    """
    Click Cancel using the same 50051 UIA method as Button Block.
    """
    if not find_and_click_50051(
        "Station 50051",
        "Cancel",
        "Button"
    ):
        log("50051 CANCEL BUTTON NOT FOUND")
        return False

    log("50051 CANCEL CLICKED")
    time.sleep(1)
    return True

def is_out_steady(handle):

    handle = handle.strip().upper()

    point = crank_handle_points[handle]["OUT"]

    screenshot = pyautogui.screenshot()

    r, g, b = get_avg_color(
        point[0],
        point[1],
        screenshot
    )

    log(f"{handle} OUT RGB = {r},{g},{b}")

    return is_red_color(r, g, b)

def remove_crank_handle(handle):

    if not click_crank_button(handle):

        return False

    if not click_crank_receive():

        return False

    log(f"{handle} REMOVED")

    time.sleep(3)

    return True



def recover_panel(signal, handle):

    log("=" * 60)
    log("STARTING SAFE RECOVERY")

    try:
        click(signals[signal]["menu"])
        time.sleep(1)
        if not click_menu_item("Signal Cancel"):
            log("SIGNAL CANCEL FAILED DURING RECOVERY")
        log("SIGNAL CANCEL DONE")
    except Exception as e:
        log(str(e))

    time.sleep(1)

    try:
        click(signals[signal]["menu"])
        time.sleep(1)
        if not click_menu_item("Route Release"):
            log("ROUTE RELEASE FAILED DURING RECOVERY")
        log("ROUTE RELEASE DONE")
    except Exception as e:
        log(str(e))

    time.sleep(1)

    try:
        remove_crank_handle(handle)
        log("CRANK HANDLE REMOVED")
    except Exception as e:
        log(str(e))

    log("SAFE RECOVERY COMPLETED")


def fail_step(signal, handle, message):

    log(message)

    return "FAIL"

# =========================================================
# VALIDATE ALL CRANK HANDLES
# =========================================================

def validate_all_crank_handles():

    log("=" * 80)
    log("STARTING CRANK HANDLE VALIDATION")
    log("=" * 80)

    handles = sorted(crank_handle_points.keys())

    if len(handles) == 0:

        log("NO CRANK HANDLES AVAILABLE")

        return False

    # -----------------------------------------------------
    # STEP 1
    # CLICK RED BUTTON + TRANSMIT FOR ALL HANDLES
    # -----------------------------------------------------

    for handle in handles:

        log(f"VALIDATING {handle}")

        if not click_crank_button(handle):

            log(f"{handle} BUTTON FAILED")

            return False

        if not click_crank_transmit():

            log(f"{handle} TRANSMIT FAILED")

            return False

        time.sleep(2)

        # Verify OUT blinking

        if not is_out_blinking(handle):

            log(f"{handle} OUT NOT BLINKING")

            return False

        log(f"{handle} OUT BLINK VERIFIED")

    # -----------------------------------------------------
    # STEP 2
    # CTRL+B
    # -----------------------------------------------------

    if not open_station_bits():

        log("FAILED TO OPEN STATION WINDOW")

        return False

    # -----------------------------------------------------
    # STEP 3
    # VERIFY EVERY HANDLE
    # -----------------------------------------------------

    for handle in handles:

        log(f"CHECKING {handle}")

        if not select_crank_handle(handle):

            return False

        if not click_station_transmit():

            return False

        time.sleep(2)

        if not is_out_steady(handle):

            log(f"{handle} OUT NOT STEADY")

            return False

        log(f"{handle} OUT STEADY VERIFIED")

        if not select_crank_handle(handle):

            return False

        if not click_station_transmit():

            return False

        time.sleep(2)

        if not is_out_blinking(handle):

            log(f"{handle} OUT NOT BLINKING")

            return False

        log(f"{handle} OUT BLINK VERIFIED")

    # -----------------------------------------------------
    # STEP 4
    # CANCEL
    # -----------------------------------------------------

    if not click_station_cancel():

        log("FAILED TO CLICK CANCEL")

        return False

    time.sleep(2)

    # -----------------------------------------------------
    # STEP 5
    # RECEIVE ALL HANDLES
    # -----------------------------------------------------

    for handle in handles:

        if not click_crank_button(handle):

            return False

        if not click_crank_receive():

            return False

        time.sleep(2)

    log("=" * 80)
    log("ALL CRANK HANDLES VERIFIED SUCCESSFULLY")
    log("=" * 80)

    return True

# =========================================================
# DROP TRACK
# =========================================================


# =========================================================
# DETECT SIGNAL TYPE
# =========================================================


# =========================================================
# GET TRACK FOR LOCK ROUTE
# =========================================================


# =========================================================
# CREATE ADVANCED REPORT
# =========================================================

def create_report():

    wb = Workbook()

    ws = wb.active

    ws.title = "VISUAL INSPECTION REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:F1")

    title_cell = ws["A1"]

    title_cell.value = "VISUAL INSPECTION AUTOMATION REPORT"

    title_cell.font = Font(
        bold=True,
        size=16,
        color="000000"
    )

    title_cell.fill = PatternFill(fill_type=None)

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
        "SIGNAL",
        "ROUTE",
        "CRANK HANDLE",
        "RESULT",
        "REASON",
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

        cell.fill = PatternFill(fill_type=None)

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

            row["SIGNAL"],
            row["ROUTE"],
            row["CRANK_HANDLE"],
            row["RESULT"],
            row["REASON"],
            row["DATE & TIME"]

        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================

    for row in ws.iter_rows(min_row=4):

        # RESULT is column D
        result_value = str(row[3].value).strip().upper()

        if result_value == "FAIL":
            # Highlight the ENTIRE failed test row in red
            row_fill = PatternFill(
                fill_type="solid",
                fgColor="FFC7CE"
            )
            row_font = Font(
                color="9C0006"
            )
        else:
            row_fill = PatternFill(fill_type=None)
            row_font = Font(
                color="000000"
            )

        for cell in row:
            cell.fill = row_fill
            cell.font = row_font

            # Borders
            cell.border = border

            # Alignment
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
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


            except Exception as e:

                log(str(e))

        adjusted = max_length + 5

        ws.column_dimensions[column_letter].width = adjusted

    # =====================================================
    # SUMMARY
    # =====================================================

    total_routes = len(lock_routes_data)

    # A route is initiated when its result has been recorded.
    initiated_routes = len(report_rows)

    non_initiated_routes = max(
        total_routes - initiated_routes,
        0
    )

    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL ROUTES"
    ws[f"B{start}"] = total_routes

    ws[f"A{start + 1}"] = "INITIATED ROUTES"
    ws[f"B{start + 1}"] = initiated_routes

    ws[f"A{start + 2}"] = "NON INITIATED ROUTES"
    ws[f"B{start + 2}"] = non_initiated_routes

    for r in range(start, start + 3):

        ws[f"A{r}"].font = Font(bold=True)
        ws[f"B{r}"].font = Font(bold=True)

        ws[f"A{r}"].border = border
        ws[f"B{r}"].border = border

        ws[f"A{r}"].alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        ws[f"B{r}"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # CONSOLIDATED SUMMARY SHEET
    # =====================================================

    summary_ws = wb.create_sheet(
        "CONSOLIDATED SUMMARY"
    )

    # =====================================================
    # TITLE
    # =====================================================

    summary_ws.merge_cells("A1:C1")

    title = summary_ws["A1"]

    title.value = "CONSOLIDATED TEST SUMMARY"

    title.font = Font(
        bold=True,
        size=18,
        color="000000"
    )

    title.fill = PatternFill(fill_type=None)

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    summary_ws.row_dimensions[1].height = 35

    # =====================================================
    # HEADERS
    # =====================================================

    headers = [

        "SIGNAL",
        "TOTAL TESTS",
        "FINAL RESULT"

    ]

    summary_ws.append([])
    summary_ws.append(headers)



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

    for cell in summary_ws[3]:
        cell.font = Font(
            bold=True,
            color="000000"
        )

        cell.fill = PatternFill(fill_type=None)

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # =====================================================
    # GROUP RESULTS
    # =====================================================

    signal_summary = {}

    for row in report_rows:

        sig = row["SIGNAL"]

        result = row["RESULT"]

        if sig not in signal_summary:
            signal_summary[sig] = {

                "total": 0,
                "failed": 0

            }

        signal_summary[sig]["total"] += 1

        if result == "FAIL":
            signal_summary[sig]["failed"] += 1

    # =====================================================
    # WRITE SUMMARY
    # =====================================================


    red_fill = PatternFill(
        "solid",
        fgColor="FEE2E2"
    )

    for sig, data in signal_summary.items():

        total = data["total"]

        failed = data["failed"]

        if failed == 0:

            final_result = "ALL TEST CASES PASSED"


        else:

            failed_routes = []

            for r in report_rows:

                if (
                        r["SIGNAL"] == sig
                        and r["RESULT"] == "FAIL"
                ):
                    failed_routes.append(
                        f'{r["ROUTE"]} ({r["CRANK_HANDLE"]}) NOT PASSED'
                    )

            final_result = ", ".join(failed_routes)

            fill = red_fill

        summary_ws.append([

            sig,
            total,
            final_result

        ])

        current_row = summary_ws.max_row

        for cell in summary_ws[current_row]:
            if failed > 0:
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="FFC7CE"
                )
                cell.font = Font(
                    color="9C0006"
                )
            else:
                cell.fill = PatternFill(fill_type=None)
                cell.font = Font(
                    color="000000"
                )

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # =====================================================
    # AUTO WIDTH
    # =====================================================

    for column in summary_ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column[0].column
        )

        for cell in column:

            try:

                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))


            except Exception as e:

                log(str(e))

        adjusted = max_length + 5

        summary_ws.column_dimensions[
            column_letter
        ].width = adjusted

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(REPORT_FILE)

    log(f"REPORT SAVED : {REPORT_FILE}")

def ensure_lc_closed():

    log("=" * 70)
    log("LC VALIDATION STARTED")

    # =====================================================
    # FIND CONFIGURED LC
    # =====================================================

    lc_signal = None

    for name, info in signals.items():

        if info.get("type") == "LC":

            lc_signal = name
            break

    if lc_signal is None:

        log("NO LC GATE CONFIGURED")

        return False

    lc = signals[lc_signal]

    log(f"LC GATE : {lc_signal}")

    # =====================================================
    # GET LC COORDINATES
    # =====================================================

    menu_point = lc.get("menu")
    in_point = lc.get("IN")
    out_point = lc.get("OUT")

    if not menu_point:

        log("LC MENU COORDINATE NOT CONFIGURED")

        return False

    if not in_point:

        log("LC IN COORDINATE NOT CONFIGURED")

        return False

    if not out_point:

        log("LC OUT COORDINATE NOT CONFIGURED")

        return False

    log(f"LC MENU : {menu_point}")
    log(f"LC IN   : {in_point}")
    log(f"LC OUT  : {out_point}")

    # =====================================================
    # STEP 1
    # CHECK IF LC IS ALREADY CLOSED
    #
    # IN = GREEN
    # =====================================================

    log("CHECKING LC IN INDICATION")

    if is_steady_green(in_point):

        log("LC IN ALREADY STEADY GREEN")

        log("LC ALREADY CLOSED")

        return True

    # =====================================================
    # STEP 2
    # LC IS NOT CLOSED
    #
    # OUT = YELLOW / BLINKING
    # =====================================================

    log("LC IN IS NOT GREEN")

    log("VERIFYING LC OUT INDICATION")

    if not is_blinking_yellow(out_point):

        log("LC OUT IS NOT BLINKING YELLOW")

        return False

    log("LC OUT BLINK VERIFIED")

    # =====================================================
    # STEP 3
    # CLICK LC MENU / BUTTON
    # =====================================================

    log("CLICKING LC MENU")

    if not click(menu_point):

        log("FAILED TO CLICK LC MENU")

        return False

    time.sleep(1)

    # =====================================================
    # STEP 4
    # CLICK RECEIVE
    # =====================================================

    log("CLICKING RECEIVE")

    if not click_menu_item("Receive"):

        log("RECEIVE MENU NOT FOUND")

        return False

    # =====================================================
    # STEP 5
    # WAIT FOR LC IN GREEN
    #
    # IN = CLOSED
    # =====================================================

    log("WAITING FOR LC IN GREEN")

    timeout = 15

    start = time.time()

    while time.time() - start < timeout:

        if is_steady_green(in_point):

            log("LC CLOSED SUCCESSFULLY")

            return True

        time.sleep(1)

    # =====================================================
    # FAILED
    # =====================================================

    log("LC FAILED TO CLOSE")

    return False

def validate_locked_points(signal, route, points_text):
    """
    Validate all point indications required by the TOC.

    TOC logical format:
        50N / 50R
        65N / 65R

    Visual configuration may contain physical indications:
        50A, 50B
        65A, 65B

    Therefore a logical TOC point such as 50R is resolved to all
    configured physical members 50A / 50B when an exact point 50
    does not exist.

    N = GREEN
    R = YELLOW
    """

    if not points_text:
        log(
            f"{signal} -> {route}: "
            "NO POINT LOCKS/DETECTS REQUIRED"
        )
        return True

    required_points = [
        item.strip().upper()
        for item in points_text.split(",")
        if item.strip()
    ]

    log("=" * 60)
    log("POINT VALIDATION STARTED")
    log(f"SIGNAL : {signal}")
    log(f"ROUTE  : {route}")
    log(f"TOC POINTS : {required_points}")
    log("=" * 60)

    # Refresh the current physical point configuration.
    rebuild_point_indications()

    log(
        f"AVAILABLE CONFIGURED POINTS : "
        f"{list(point_indications.keys())}"
    )

    for required in required_points:

        # -----------------------------------------------------
        # CHECK FORMAT
        # -----------------------------------------------------
        if len(required) < 2:
            log(
                f"INVALID TOC POINT FORMAT : {required}"
            )
            return False

        point_no = required[:-1].strip().upper()
        required_position = required[-1].upper()

        log(
            f"PROCESSING POINT : "
            f"{point_no}{required_position}"
        )

        # -----------------------------------------------------
        # CHECK POSITION
        # -----------------------------------------------------
        if required_position not in ("N", "R"):
            log(
                f"INVALID POINT POSITION : "
                f"{required_position}"
            )
            return False

        # -----------------------------------------------------
        # RESOLVE LOGICAL TOC POINT TO PHYSICAL CONFIGURATION
        # -----------------------------------------------------
        point_infos = []

        # First prefer an exact configured point number.
        exact_point = signals.get(point_no)
        if exact_point:
            point_infos.append((point_no, exact_point))
        else:
            # Example:
            #     TOC 50R -> physical 50A + 50B
            #     TOC 65N -> physical 65A + 65B
            for configured_name, configured_info in signals.items():
                configured_name = str(configured_name).strip().upper()

                if configured_info.get("type") != "POINT":
                    continue

                suffix = configured_name[len(point_no):]

                if (
                    configured_name.startswith(point_no)
                    and suffix in ("A", "B")
                ):
                    point_infos.append(
                        (configured_name, configured_info)
                    )

        if not point_infos:
            log(
                f"POINT {point_no} NOT FOUND "
                "IN SIGNAL CONFIG"
            )
            log(
                f"AVAILABLE SIGNALS : "
                f"{list(signals.keys())}"
            )
            return False

        if len(point_infos) > 1:
            log(
                f"LOGICAL POINT {point_no}{required_position} "
                f"RESOLVED TO : "
                f"{[name for name, _ in point_infos]}"
            )

        # -----------------------------------------------------
        # VALIDATE EVERY PHYSICAL MEMBER
        # -----------------------------------------------------
        for physical_name, point_info in point_infos:

            log(
                f"VALIDATING PHYSICAL POINT : {physical_name}"
            )

            # -------------------------------------------------
            # CHECK TYPE
            # -------------------------------------------------
            if point_info.get("type") != "POINT":
                log(
                    f"{physical_name} EXISTS BUT TYPE IS "
                    f"{point_info.get('type')}"
                )
                return False

            # -------------------------------------------------
            # GET NORMAL / REVERSE COORDINATES
            # -------------------------------------------------
            normal_lamp = point_info.get("normal")
            reverse_lamp = point_info.get("reverse")

            if not normal_lamp:
                log(
                    f"POINT {physical_name} NORMAL "
                    "COORDINATE NOT CONFIGURED"
                )
                return False

            if not reverse_lamp:
                log(
                    f"POINT {physical_name} REVERSE "
                    "COORDINATE NOT CONFIGURED"
                )
                return False

            # -------------------------------------------------
            # SELECT EXPECTED INDICATION
            # -------------------------------------------------
            if required_position == "N":
                required_lamp = normal_lamp
                opposite_lamp = reverse_lamp

                log(
                    f"POINT {physical_name} EXPECTED : NORMAL"
                )
                log(f"N COORDINATE : {normal_lamp}")
                log(f"R COORDINATE : {reverse_lamp}")

            else:
                required_lamp = reverse_lamp
                opposite_lamp = normal_lamp

                log(
                    f"POINT {physical_name} EXPECTED : REVERSE"
                )
                log(f"R COORDINATE : {reverse_lamp}")
                log(f"N COORDINATE : {normal_lamp}")

            # -------------------------------------------------
            # REQUIRED INDICATION
            # N = GREEN
            # R = YELLOW
            # -------------------------------------------------
            verified = False

            for attempt in range(10):

                if required_position == "N":
                    if is_steady_green(required_lamp):
                        verified = True
                        break
                else:
                    if is_steady_yellow(required_lamp):
                        verified = True
                        break

                log(
                    f"Waiting for {physical_name}{required_position} "
                    f"({attempt + 1}/10)"
                )
                time.sleep(1)

            if not verified:
                log(
                    f"{physical_name}{required_position} "
                    "REQUIRED INDICATION FAILED"
                )
                return False

            log(
                f"{physical_name}{required_position} "
                "REQUIRED INDICATION VERIFIED"
            )

            # -------------------------------------------------
            # OPPOSITE INDICATION MUST BE OFF
            # -------------------------------------------------
            for _ in range(4):

                if is_yellow(opposite_lamp):
                    log(
                        f"{physical_name} "
                        f"{'N' if required_position == 'R' else 'R'} "
                        "SHOULD BE OFF"
                    )
                    return False

                if is_green(opposite_lamp):
                    log(
                        f"{physical_name} "
                        f"{'N' if required_position == 'R' else 'R'} "
                        "SHOULD BE OFF"
                    )
                    return False

                time.sleep(0.4)

            log(
                f"{physical_name}{required_position} PASSED"
            )

    log("=" * 60)
    log("ALL POINTS VALIDATED SUCCESSFULLY")
    log("=" * 60)

    return True

def validate_crank_handle_in_only(handle):

    handle = handle.strip().upper()

    if handle not in crank_handle_points:
        log(f"{handle} NOT CONFIGURED")
        return False

    info = crank_handle_points[handle]

    # IN must remain steady yellow.
    if not is_steady_yellow(info["IN"]):
        log(f"{handle}: IN NOT STEADY YELLOW")
        return False

    # OUT, ECH, and FREE must remain inactive.
    for _ in range(4):

        if is_yellow(info["OUT"]):
            log(f"{handle}: OUT YELLOW - FAIL")
            return False

        if is_yellow(info["ECH"]):
            log(f"{handle}: ECH YELLOW - FAIL")
            return False

        if is_green(info["FREE"]):
            log(f"{handle}: FREE GREEN - FAIL")
            return False

        time.sleep(0.4)

    log(f"{handle}: IN STEADY YELLOW ONLY - PASSED")
    return True

def cancel_and_release_route(signal):

    log("SIGNAL CANCEL")

    if not click(signals[signal]["menu"]):
        return False, "FAILED TO OPEN SIGNAL MENU FOR SIGNAL CANCEL"

    time.sleep(1)

    if not click_menu_item("Signal Cancel"):
        return False, "SIGNAL CANCEL FAILED"

    log("✓ SIGNAL CANCEL SUCCESS")

    time.sleep(2)

    log("ROUTE RELEASE")

    if not click(signals[signal]["menu"]):
        return False, "FAILED TO OPEN SIGNAL MENU FOR ROUTE RELEASE"

    time.sleep(1)

    if not click_menu_item("Route Release"):
        return False, "ROUTE RELEASE FAILED"

    log("ROUTE RELEASE SUCCESS")

    timeout = 15
    start = time.time()

    while time.time() - start < timeout:

        if verify_signal_reset(signal):
            return True, "SIGNAL CANCELLED AND ROUTE RELEASED"

        time.sleep(1)

    return False, "SIGNAL DID NOT RETURN TO NORMAL"

def test_route(row):

    signal = str(
        row["signal"]
    ).strip().upper()

    route = str(
        row["route"]
    ).strip().upper()

    crank_handle = str(
        row.get("crank_handle", "")
    ).strip().upper()

    points_text = str(
        row.get("locks_detects_points", "")
    ).strip().upper()

    lcp_value = str(
        row.get("44_lcp", "")
    ).strip()

    signal_type = (
        signals.get(signal, {}).get("type", "")
        if signal in signals
        else ""
    )

    approach_track = str(
        row.get("approach_track", "")
    ).strip().upper()

    # =====================================================
    # ROUTE START LOG
    # =====================================================

    log("=" * 80)
    log("STARTING ROUTE TEST")
    log(f"SIGNAL : {signal}")
    log(f"TYPE   : {signal_type}")
    log(f"ROUTE  : {route}")

    if signal_type == "CALLING_ON":
        log(
            f"APPROACH TRACK : "
            f"{approach_track or '[EMPTY]'}"
        )

    log("=" * 80)

    # =====================================================
    # SKIP ROUTE IF NOTHING TO VALIDATE
    #
    # Calling-ON is never skipped because its 50051
    # preparation and Route Initiator verification are
    # the Visual Inspection test itself.
    # =====================================================

    if (
        signal_type != "CALLING_ON"
        and lcp_value != "44"
        and points_text == ""
        and crank_handle == ""
    ):
        log("NO LC")
        log("NO POINT INDICATION")
        log("NO CRANK HANDLE")
        log("ROUTE SKIPPED")

        return (
            "NOT TESTED",
            "NO LC / POINT / CRANK HANDLE IN TOC"
        )

    handles = [
        handle.strip().upper()
        for handle in crank_handle.split(",")
        if handle.strip()
    ]

    def fail_after_route_set(reason):

        cleanup_ok, cleanup_message = (
            cancel_and_release_route(signal)
        )

        if not cleanup_ok:
            return (
                "FAIL",
                f"{reason} | "
                f"CLEANUP FAILED: {cleanup_message}"
            )

        log("ROUTE CLEANUP COMPLETED")

        # =====================================================
        # CALLING-ON FINAL 50051 RELOAD AFTER CLEANUP
        # =====================================================
        # IMPORTANT:
        # Once a Calling-On route has been SET, the final 50051
        # approach-track operation must happen AFTER:
        #     Signal Cancel -> Route Release -> Reset Verified
        # This is required even when point/crank validation fails.
        #
        # Sequence:
        #     CTRL+B
        #     -> click 50051
        #     -> ENTER
        #     -> select the SAME approach track (1CXTPR)
        #     -> Transmit
        #     -> Cancel
        # =====================================================
        if signal_type == "CALLING_ON":
            log(
                "CALLING-ON AFTER CLEANUP -> "
                "RELOADING 50051"
            )

            if not complete_calling_on_approach_track(
                    approach_track):
                return (
                    "FAIL",
                    f"{reason} | "
                    "SIGNAL CANCELLED - ROUTE RELEASED | "
                    "CALLING-ON FINAL 50051 RELOAD FAILED"
                )

            log(
                "CALLING-ON AFTER CLEANUP -> "
                f"{approach_track} RELOADED SUCCESSFULLY"
            )

        message = (
            f"{reason} | SIGNAL CANCELLED - ROUTE RELEASED"
        )

        if signal_type == "CALLING_ON":
            message += " | CALLING-ON 50051 RELOADED"

        return "FAIL", message

    # =====================================================
    # CALLING-ON 50051 PREPARATION
    #
    # CTRL+B
    # -> CLICK 50051 ONCE + ENTER
    # -> SELECT APPROACH TRACK FROM TOC
    # -> TRANSMIT
    # -> CLOSE 50051
    # -> SET ROUTE
    # =====================================================

    if signal_type == "CALLING_ON":

        if not approach_track:
            log(
                "CALLING-ON APPROACH TRACK "
                "NOT CONFIGURED IN TOC"
            )

            # Route was not set, so do NOT Signal Cancel
            # and do NOT Route Release.
            return (
                "FAIL",
                "CALLING-ON APPROACH TRACK "
                "NOT CONFIGURED"
            )

        if not prepare_calling_on_approach_track(
                approach_track):

            # Preparation failed before route set.
            # Do NOT Signal Cancel / Route Release.
            return (
                "FAIL",
                "CALLING-ON 50051 APPROACH TRACK "
                "PREPARATION FAILED"
            )

    # =====================================================
    # LC VALIDATION
    # =====================================================

    if lcp_value == "44":

        log("LC VALIDATION REQUIRED")

        if not ensure_lc_closed():
            return "FAIL", "LC VALIDATION FAILED"

    else:
        log("LC NOT REQUIRED FOR THIS ROUTE")

    # =====================================================
    # SET ROUTE
    # =====================================================

    if signal not in signals:
        return "FAIL", "SIGNAL NOT FOUND"

    if not set_route(signal, route):
        # Route was not established.
        # Do NOT Signal Cancel / Route Release.
        return "FAIL", "ROUTE SET FAILED"

    # =====================================================
    # WAIT FOR ROUTE SET
    #
    # CALLING-ON:
    #   YELLOW LAMP + ROUTE INITIATOR YELLOW
    # =====================================================

    route_set = False

    for attempt in range(3):

        if wait_for_route_set(
                signal,
                timeout=30):

            route_set = True
            break

        log("ROUTE NOT SET AFTER 30 SECONDS")

        if attempt < 2:
            log(
                f"RETRY ROUTE SET "
                f"{attempt + 1}"
            )
            set_route(signal, route)

    if not route_set:
        # Route never reached verified-set state.
        # Do NOT Signal Cancel / Route Release.
        return (
            "FAIL",
            "ROUTE DID NOT SET"
        )

    log(
        "ROUTE SET VERIFIED "
        "(INCLUDING CALLING-ON ROUTE INITIATOR)"
        if signal_type == "CALLING_ON"
        else "ROUTE SET VERIFIED"
    )

    time.sleep(2)

    # =====================================================
    # POINT VALIDATION
    # =====================================================

    if points_text != "":

        if not validate_locked_points(
                signal,
                route,
                points_text):

            return fail_after_route_set(
                "POINT VALIDATION FAILED"
            )

    else:
        log("NO POINTS REQUIRED")

    # =====================================================
    # CRANK HANDLE VALIDATION
    # =====================================================

    if crank_handle != "":

        handles = [
            h.strip().upper()
            for h in crank_handle.split(",")
            if h.strip()
        ]

        for handle in handles:

            if not validate_crank_handle_in_only(
                    handle):

                return fail_after_route_set(
                    f"{handle} FAILED"
                )

    else:
        log("NO CRANK HANDLE REQUIRED")

    # =====================================================
    # ROUTE WAS SUCCESSFULLY SET
    #
    # NOW:
    #   SIGNAL CANCEL
    #   ROUTE RELEASE
    # =====================================================

    cleanup_ok, cleanup_message = (
        cancel_and_release_route(signal)
    )

    if not cleanup_ok:
        return (
            "FAIL",
            f"CLEANUP FAILED: "
            f"{cleanup_message}"
        )

    # =====================================================
    # CALLING-ON FINAL 50051 TRANSMIT (SUCCESS PATH)
    #
    # Signal Cancel -> Route Release has already completed above.
    # Then: CTRL+B -> CLICK 50051 ONCE + ENTER
    #       -> SELECT SAME APPROACH TRACK
    #       -> TRANSMIT -> CLOSE 50051
    # =====================================================

    if signal_type == "CALLING_ON":

        if not complete_calling_on_approach_track(
                approach_track):

            return (
                "FAIL",
                "CALLING-ON FINAL 50051 "
                "APPROACH TRACK TRANSMIT FAILED"
            )

    # =====================================================
    # PASS
    # =====================================================

    log("SIGNAL RESET VERIFIED")
    log("=" * 80)
    log(f"{signal} -> {route} : PASS")
    log("=" * 80)

    if signal_type == "CALLING_ON":
        reason = (
            "CALLING-ON 50051 APPROACH TRACK "
            "TRANSMITTED - ROUTE INITIATOR YELLOW "
            "- SIGNAL CANCELLED - ROUTE RELEASED "
            "- FINAL 50051 TRANSMIT COMPLETED"
        )
    else:
        reason = (
            "ALL TESTS PASSED - "
            "SIGNAL CANCELLED - ROUTE RELEASED"
        )

    return "PASS", reason


# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():

    global running

    root.iconify()

    time.sleep(3)

    report_rows.clear()

    for row in lock_routes_data:

        pause_event.wait()

        if not running:
            break

        signal = row["signal"]
        route = row["route"]
        crank_handle = row["crank_handle"]

        log("=" * 80)
        log(f"STARTING ROUTE TEST: {signal} -> {route}")

        result, reason = test_route(row)

        if result != "NOT TESTED":
            report_rows.append({

                "SIGNAL": signal,
                "ROUTE": route,
                "CRANK_HANDLE": crank_handle,
                "RESULT": result,
                "REASON": reason,
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            })

        log(f"RESULT: {result} - {reason}")

        if result == "FAIL":

            log("=" * 80)
            log("ROUTE FAILED")
            log(reason)
            log("=" * 80)

            if "CLEANUP FAILED" in reason:
                log("AUTOMATION STOPPED - CLEANUP FAILED")

                running = False

                status_label.config(
                    text="FAILED",
                    fg="#dc2626"
                )

                create_report()

                root.deiconify()

                return

            time.sleep(2)

            continue

        # If PASS, test_route() has already done:
        log("=" * 80)
        log("ROUTE PASSED")
        log("SIGNAL CANCEL COMPLETED")
        log("ROUTE RELEASE COMPLETED")
        log("STARTING NEXT ROUTE")
        log("=" * 80)

        time.sleep(2)

    running = False

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    create_report()

    root.deiconify()

    log("ALL TOC ROUTES COMPLETED")

# =========================================================
# TEST SINGLE CRANK HANDLE
# =========================================================

def test_crank_handle(signal, route, handle):

    log("=" * 80)
    log(f"STARTING TEST : {handle}")

    # -------------------------------------------------
    # 1. SET ROUTE
    # -------------------------------------------------

    if signal not in signals:
        return fail_step(signal, handle, "SIGNAL NOT FOUND")

    if not click(signals[signal]["menu"]):
        return fail_step(signal, handle, "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item(route.replace("-", "_")):
        return fail_step(signal, handle, "ROUTE SET FAILED")

    log("ROUTE COMMAND SENT")

    time.sleep(2)

    # -------------------------------------------------
    # 2. SIGNAL CANCEL
    # -------------------------------------------------

    if not click(signals[signal]["menu"]):
        return fail_step(signal, handle, "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item("Signal Cancel"):
        return fail_step(signal, handle, "SIGNAL CANCEL FAILED")

    log("SIGNAL CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 3. CH BUTTON
    # -------------------------------------------------

    if not click_crank_button(handle):
        return fail_step(signal, handle, "FAILED TO CLICK CRANK BUTTON")

    # -------------------------------------------------
    # 4. TRANSMIT
    # -------------------------------------------------

    if not click_crank_transmit():
        return fail_step(signal, handle, "TRANSMIT FAILED")

    time.sleep(2)

    # -------------------------------------------------
    # 5. VERIFY ECH BLINK
    # -------------------------------------------------

    log("VERIFYING ECH YELLOW BLINK")

    if not is_ech_blinking(handle):
        return fail_step(signal, handle, "ECH NOT BLINKING")

    log("ECH BLINK VERIFIED")

    # -------------------------------------------------
    # 6. VERIFY OUT BLINK
    # -------------------------------------------------

    log("VERIFYING OUT RED BLINK")

    if not is_out_blinking(handle):
        return fail_step(signal, handle, "OUT NOT BLINKING")

    log("OUT BLINK VERIFIED")

    # -------------------------------------------------
    # 7. CTRL+B
    # 8. OPEN STATION
    # -------------------------------------------------

    if not open_station_bits():
        return fail_step(signal, handle, "FAILED TO OPEN STATION")

    # -------------------------------------------------
    # 9. SELECT CH1KLNR
    # -------------------------------------------------

    if not select_crank_handle(handle):
        return fail_step(signal, handle, "FAILED TO SELECT KLNR")

    # -------------------------------------------------
    # 10. TRANSMIT
    # -------------------------------------------------

    if not click_station_transmit():
        return fail_step(signal, handle, "STATION TRANSMIT FAILED")

    time.sleep(2)

    # -------------------------------------------------
    # 11. VERIFY OUT STEADY
    # -------------------------------------------------

    if not is_out_steady(handle):
        return fail_step(signal, handle, "OUT NOT STEADY")

    log("OUT STEADY VERIFIED")

    # -------------------------------------------------
    # 12. SELECT AGAIN
    # -------------------------------------------------

    if not select_crank_handle(handle):
        return fail_step(signal, handle, "FAILED TO SELECT KLNR AGAIN")

    # -------------------------------------------------
    # 13. TRANSMIT
    # -------------------------------------------------

    if not click_station_transmit():
        return fail_step(signal, handle, "SECOND TRANSMIT FAILED")

    time.sleep(2)

    # -------------------------------------------------
    # 14. VERIFY OUT BLINK
    # -------------------------------------------------

    if not is_out_blinking(handle):
        return fail_step(signal, handle, "OUT NOT BLINKING")

    log("OUT BLINK VERIFIED")

    # -------------------------------------------------
    # 15. CANCEL
    # -------------------------------------------------

    if not click_station_cancel():
        return fail_step(signal, handle, "STATION CANCEL FAILED")

    log("STATION CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 16. SIGNAL CANCEL
    # -------------------------------------------------

    if not click(signals[signal]["menu"]):
        return fail_step(signal, handle, "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item("Signal Cancel"):
        return fail_step(signal, handle, "SIGNAL CANCEL FAILED")

    log("SIGNAL CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 17. ROUTE RELEASE
    # -------------------------------------------------

    if not click(signals[signal]["menu"]):
        return fail_step(signal, handle, "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item("Route Release"):
        return fail_step(signal, handle, "ROUTE RELEASE FAILED")

    log("ROUTE RELEASE SUCCESS")

    time.sleep(5)

    # -------------------------------------------------
    # 18. CLICK CH BUTTON
    # -------------------------------------------------

    if not click_crank_button(handle):
        return fail_step(signal, handle, "FAILED TO CLICK CRANK BUTTON")

    # -------------------------------------------------
    # 19. RECEIVE
    # -------------------------------------------------

    if not click_crank_receive():
        return fail_step(signal, handle, "RECEIVE FAILED")

    log("CRANK RECEIVE SUCCESS")

    time.sleep(3)

    # -------------------------------------------------
    # 20. VERIFY IN YELLOW
    # -------------------------------------------------

    if not is_in_yellow(handle):
        return fail_step(signal, handle, "IN NOT YELLOW")

    log("IN YELLOW VERIFIED")

    # -------------------------------------------------
    # 21. VERIFY OUT OFF
    # -------------------------------------------------

    if not is_out_off(handle):
        return fail_step(signal, handle, "OUT NOT OFF")

    log("OUT OFF VERIFIED")

    # -------------------------------------------------
    # 22. PASS
    # -------------------------------------------------

    log(f"{handle} PASSED")

    return "PASS"

# =========================================================
# TEST MULTIPLE CRANK HANDLES
# =========================================================

def test_multiple_crank_handles(signal, route, handles):

    log("=" * 80)
    log(f"STARTING MULTIPLE CRANK HANDLE TEST : {', '.join(handles)}")

    # -------------------------------------------------
    # 1. SET ROUTE
    # -------------------------------------------------

    if signal not in signals:
        return fail_step(signal, ",".join(handles), "SIGNAL NOT FOUND")

    if not click(signals[signal]["menu"]):
        return fail_step(signal, ",".join(handles), "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item(route.replace("-", "_")):
        return fail_step(signal, ",".join(handles), "ROUTE SET FAILED")

    log("ROUTE COMMAND SENT")

    time.sleep(2)

    # -------------------------------------------------
    # 2. SIGNAL CANCEL
    # -------------------------------------------------

    if not click(signals[signal]["menu"]):
        return fail_step(signal, ",".join(handles), "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item("Signal Cancel"):
        return fail_step(signal, ",".join(handles), "SIGNAL CANCEL FAILED")

    log("SIGNAL CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 3. APPLY + VERIFY EACH HANDLE
    # -------------------------------------------------

    for handle in handles:

        log("=" * 60)
        log(f"APPLYING {handle}")

        if not click_crank_button(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} BUTTON FAILED"
            )

        if not click_crank_transmit():
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} TRANSMIT FAILED"
            )

        time.sleep(2)

        # -----------------------------
        # VERIFY ECH IMMEDIATELY
        # -----------------------------

        log(f"VERIFYING {handle} ECH BLINK")

        if not is_ech_blinking(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} ECH NOT BLINKING"
            )

        log(f"{handle} ECH VERIFIED")

        # -----------------------------
        # VERIFY OUT IMMEDIATELY
        # -----------------------------

        log(f"VERIFYING {handle} OUT BLINK")

        retry = 0

        while retry < 10:

            if is_out_blinking(handle):
                break

            retry += 1

            time.sleep(1)

        if retry == 10:
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} OUT NOT BLINKING"
            )

        log(f"{handle} OUT VERIFIED")

    log("ALL CRANK HANDLES APPLIED SUCCESSFULLY")

    # -------------------------------------------------
    # 6. OPEN STATION WINDOW
    # -------------------------------------------------

    log("OPENING STATION 50051")

    if not open_station_bits():

        return fail_step(
            signal,
            ",".join(handles),
            "FAILED TO OPEN STATION"
        )

    log("STATION 50051 OPENED")

    # -------------------------------------------------
    # 7A. FIRST TRANSMIT FOR EACH HANDLE
    # -------------------------------------------------

    for handle in handles:

        log("=" * 60)
        log(f"PROCESSING {handle}")

        if not select_crank_handle(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} KLNR SELECTION FAILED"
            )

        if not click_station_transmit():
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} STATION TRANSMIT FAILED"
            )

        time.sleep(2)

        log(f"VERIFYING {handle} OUT STEADY")

        if not is_out_steady(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} OUT NOT STEADY"
            )

        log(f"{handle} OUT STEADY VERIFIED")

    # -------------------------------------------------
    # 7B. SECOND TRANSMIT FOR EACH HANDLE
    # -------------------------------------------------

    for handle in handles:

        log("=" * 60)
        log(f"SECOND TRANSMIT : {handle}")

        if not select_crank_handle(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} SECOND KLNR FAILED"
            )

        if not click_station_transmit():
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} SECOND TRANSMIT FAILED"
            )

        time.sleep(2)

    # -------------------------------------------------
    # 7C. VERIFY OUT BLINK FOR ALL HANDLES
    # -------------------------------------------------

    for handle in handles:

        log(f"VERIFYING {handle} OUT BLINK")

        if not is_out_blinking(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} OUT NOT BLINKING"
            )

        log(f"{handle} OUT BLINK VERIFIED")

    # -------------------------------------------------
    # 8. CANCEL STATION WINDOW
    # -------------------------------------------------

    log("=" * 60)
    log("CANCELLING STATION WINDOW")

    if not click_station_cancel():
        return fail_step(
            signal,
            ",".join(handles),
            "STATION CANCEL FAILED"
        )

    log("STATION CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 9. SIGNAL CANCEL
    # -------------------------------------------------

    log("=" * 60)
    log("SIGNAL CANCEL")

    if not click(signals[signal]["menu"]):
        return fail_step(
            signal,
            ",".join(handles),
            "FAILED TO OPEN SIGNAL MENU"
        )

    time.sleep(1)

    if not click_menu_item("Signal Cancel"):
        return fail_step(
            signal,
            ",".join(handles),
            "SIGNAL CANCEL FAILED"
        )

    log("SIGNAL CANCEL SUCCESS")

    time.sleep(2)

    # -------------------------------------------------
    # 10. ROUTE RELEASE
    # -------------------------------------------------

    log("=" * 60)
    log("ROUTE RELEASE")

    if not click(signals[signal]["menu"]):
        return fail_step(
            signal,
            ",".join(handles),
            "FAILED TO OPEN SIGNAL MENU"
        )

    time.sleep(1)

    if not click_menu_item("Route Release"):
        return fail_step(
            signal,
            ",".join(handles),
            "ROUTE RELEASE FAILED"
        )

    log("ROUTE RELEASE SUCCESS")

    time.sleep(5)

    # -------------------------------------------------
    # 11. RECEIVE EACH HANDLE
    # -------------------------------------------------

    for handle in handles:

        log("=" * 60)
        log(f"RECEIVING {handle}")

        if not click_crank_button(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} BUTTON FAILED"
            )

        if not click_crank_receive():
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} RECEIVE FAILED"
            )

        time.sleep(3)

        log(f"VERIFYING {handle} IN YELLOW")

        if not is_in_yellow(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} IN NOT YELLOW"
            )

        log(f"{handle} IN YELLOW VERIFIED")

        log(f"VERIFYING {handle} OUT OFF")

        if not is_out_off(handle):
            return fail_step(
                signal,
                ",".join(handles),
                f"{handle} OUT NOT OFF"
            )

        log(f"{handle} OUT OFF VERIFIED")

    # -------------------------------------------------
    # 12. PASS
    # -------------------------------------------------

    log("=" * 80)
    log(f"MULTIPLE CRANK HANDLE TEST PASSED : {', '.join(handles)}")
    log("=" * 80)

    return "PASS"

# =========================================================
# START
# =========================================================

def start_automation():

    global running
    global REPORT_FILE

    if not signals:

        log("LOAD CONFIG FIRST")

        return

    if not lock_routes_data:

        log("LOAD TOC FIRST")

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

    # =====================================================
    # CREATE A UNIQUE REPORT FILE FOR THIS TEST RUN
    # =====================================================

    REPORT_FILE = (
        "VISUAL_INSPECTION_Report_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".xlsx"
    )

    log(f"NEW TEST RUN REPORT : {REPORT_FILE}")

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


        except Exception as e:

            log(str(e))
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
    "VISUAL INSPECTION TESTING AUTOMATION"
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

    text="VISUAL INSPECTION AUTOMATION SYSTEM",

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

tree_scroll = ttk.Scrollbar(table_frame)

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
        "CONTROLLED_BY_TRACKS",
        "APPROACH_TRACK",
        "CRANK_HANDLE",
        "LOCKS_DETECTS_POINTS",
        "LCP"
    ),
    show="headings",
    yscrollcommand=tree_scroll.set
)

tree_scroll.config(
    command=tree.yview
)

# =========================================================
# HEADINGS
# =========================================================

tree.heading("NO", text="NO")
tree.heading("SIGNAL", text="SIGNAL")
tree.heading("ROUTE", text="ROUTE")
tree.heading("CONTROLLED_BY_TRACKS", text="CONTROLLED BY TRACKS")
tree.heading("APPROACH_TRACK", text="APPROACH TRACK")
tree.heading("CRANK_HANDLE", text="CRANK HANDLE")
tree.heading("LOCKS_DETECTS_POINTS", text="LOCKS/DETECTS POINTS")
tree.heading("LCP", text="44_LCP")

# =========================================================
# COLUMN WIDTH
# =========================================================

tree.column("NO", width=55, anchor="center")
tree.column("SIGNAL", width=80, anchor="center")
tree.column("ROUTE", width=90, anchor="center")
tree.column("CONTROLLED_BY_TRACKS", width=320, anchor="w")
tree.column("APPROACH_TRACK", width=130, anchor="center")
tree.column("CRANK_HANDLE", width=120, anchor="center")
tree.column("LOCKS_DETECTS_POINTS", width=180, anchor="center")
tree.column("LCP", width=75, anchor="center")

tree.tag_configure("PASS", background="#DCFCE7", foreground="#166534")

tree.tag_configure("FAIL", background="#FEE2E2", foreground="#991B1B")

tree.tag_configure("CHECKING", background="#FEF3C7", foreground="#92400E")

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

    text="SPACE = CAPTURE COORDINATES  |  P = PAUSE / RESUME  |  VISUAL INSPECTION AUTOMATION SYSTEM",

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

root.mainloop()
