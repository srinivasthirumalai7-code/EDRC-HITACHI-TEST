"""
===========================================================
HITACHI C-FAT AUTOMATION
TRAIN PASSING THROUGH STATION AUTOMATION
===========================================================

DEVELOPER : Suvetha
===========================================================
"""
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass
import tkinter as tk
import keyboard
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
from pathlib import Path

import pyautogui

import win32api
import win32con


import uiautomation as auto

from pywinauto import Desktop
from PIL import ImageGrab
from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# MASTER INTEGRATION
# =========================================================

MASTER_MODE = os.environ.get("EDRC_UNATTENDED") == "1"


# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "HAH COORDINATES(8).xlsx"



# =========================================================
# APPLICATION INFORMATION
# =========================================================

APP_NAME = "Hitachi C-FAT Automation"

VERSION = "1.0.0"

DEVELOPER = "SUVETHA"
# =========================================================
# DEVELOPMENT SETTINGS
# =========================================================

DEBUG_MODE = True

CURRENT_TEST = "ROUTE ENGINE"

TOC_FILE = None

SIMULATOR_RUNNING = False

PANEL_RUNNING = False
# =========================================================
# GLOBALS
# =========================================================

toc_data = []

signals = {}

points = {}

crank_handles = {}

running = False
automation_running = False

capture_waiting = False

captured_point = None

shunt_signals = {}
starter_signals = {}
calling_on_signals = {}
lc_gates = {}

# =========================================================
# THROUGH TESTING REPORT STATE
# =========================================================
# The EDRC Master Suite looks for a NEW .xlsx file in the same
# folder as THROUGH_TESTING.py.  Keep the report self-contained
# here so the child process never depends on an external "report"
# Python package.
through_report_wb = None
through_report_ws = None
through_report_path = None
# Keep route-result rows separate from the fixed summary rows.
through_report_next_row = 4
# =========================================================
# COORDINATE CAPTURE GLOBALS
# =========================================================

capture_module = None
capture_index = 0
capture_master_list = []

record_stage = None
current_aspects = 2

undo_stack = []

main_signal_aspects = {}

last_space_time = 0
last_backspace_time = 0
last_digit_time = 0

# =========================================================
# CAPTURE OVERLAY
# =========================================================

capture_overlay = None

overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None

TYPE_LABELS = {
    "MAIN": "MAIN SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "CAL": "CALLING ON SIGNAL",
    "POINT": "POINT",
    "CH": "CRANK HANDLE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "SHUNT": "#9333ea",
    "CAL": "#ea580c",
    "POINT": "#14b8a6",
    "CH": "#ca8a04",
    "LC": "#dc2626"
}

def pause_sleep(seconds):

    end_time = time.time() + seconds

    while True:

        pause_event.wait()

        remaining = end_time - time.time()

        if remaining <= 0:
            break

        time.sleep(min(0.1, remaining))
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

    pause_sleep(0.4)

    pause_event.wait()

    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTDOWN,
        0,
        0
    )

    pause_sleep(0.1)

    win32api.mouse_event(
        win32con.MOUSEEVENTF_LEFTUP,
        0,
        0
    )

    pause_sleep(0.4)

    return True
def show_capture_popup(message):

    root.deiconify()
    root.lift()
    root.attributes("-topmost", True)
    root.focus_force()
    root.update()

    messagebox.showinfo(
        "CAPTURE",
        message,
        parent=root
    )

    root.attributes("-topmost", False)
# =========================================================
# CAPTURE POINT
# =========================================================

def capture_point():

    global capture_waiting
    global captured_point

    capture_waiting = True
    captured_point = None

    while capture_waiting:

        root.update_idletasks()
        pause_sleep(0.05)

    return list(captured_point)

# =========================================================
# CAPTURE MOUSE POSITION
# =========================================================

def capture_mouse_position():

    x, y = capture_point()

    log(f"Captured Coordinate : ({x}, {y})")

    return x, y


def start_master_capture():

    global capture_module

    capture_module = "MASTER"

    create_capture_overlay()

    root.iconify()

    master_next_capture()

# =========================================================
# MASTER CAPTURE ENGINE
# =========================================================

def master_next_capture():

    global capture_index
    global capture_master_list
    global record_stage

    if capture_index >= len(capture_master_list):

        if capture_overlay:
            capture_overlay.destroy()

        root.deiconify()
        log("========================================")
        log("Coordinate Capture Completed")
        log("Press 'SAVE COORDINATES' to save the captured coordinates.")
        log("========================================")

        messagebox.showinfo(
            "Capture Completed",
            "All coordinates have been captured successfully.\n\n"
            "Click the 'SAVE COORDINATES' button to save them."
        )

        return

    item = capture_master_list[capture_index]

    item_type = item[0]

    item_name = item[1]

    stage = item[2]

    record_stage = stage

    update_capture_overlay(
        item_type,
        item_name,
        stage
    )
    log("")
    log("========================================")
    log(f"{TYPE_LABELS[item_type]} : {item_name}")
    log(f"STEP : {stage.replace('_', ' ')}")
    log("Press SPACE to Capture")
    log("Press BACKSPACE to Undo")
    log("========================================")

# =========================================================
# SAVE CURRENT CAPTURE
# =========================================================

def master_save_point(x, y):

    global capture_index
    global calling_on_signals
    global lc_gates

    item_type, item_name, stage = capture_master_list[capture_index]

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if item_type == "MAIN":

        if item_name not in signals:

            signals[item_name] = {

                "type": "MAIN",

                "aspects": main_signal_aspects[item_name],

                "menu": None,

                "RED": None,

                "YELLOW": None,

                "DOUBLE_YELLOW": None,

                "GREEN": None,

                "ROUTE_INIT": None

            }

        if stage == "MENU":
            signals[item_name]["menu"] = [x, y]

        elif stage == "RED":
            signals[item_name]["RED"] = [x, y]

        elif stage == "YELLOW":
            signals[item_name]["YELLOW"] = [x, y]

        elif stage == "DOUBLE_YELLOW":
            signals[item_name]["DOUBLE_YELLOW"] = [x, y]

        elif stage == "GREEN":
            signals[item_name]["GREEN"] = [x, y]

        elif stage == "ROUTE_INIT":
            signals[item_name]["ROUTE_INIT"] = [x, y]

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif item_type == "SHUNT":

        if item_name not in shunt_signals:

            shunt_signals[item_name] = {

                "menu": None,

                "RED": None,

                "WHITE": None,

                "ROUTE_INIT": None

            }

        shunt = shunt_signals[item_name]

        if stage == "MENU":
            shunt["menu"] = [x, y]

        elif stage == "RED":
            shunt["RED"] = [x, y]

        elif stage == "WHITE":
            shunt["WHITE"] = [x, y]

        elif stage == "ROUTE_INIT":
            shunt["ROUTE_INIT"] = [x, y]

    # =====================================================
    # CALLING ON SIGNAL
    # =====================================================

    elif item_type == "CAL":

        if item_name not in calling_on_signals:

            calling_on_signals[item_name] = {

                "menu": None,

                "RED": None,

                "WHITE": None,

                "ROUTE_INIT": None

            }

        calling = calling_on_signals[item_name]

        if stage == "MENU":
            calling["menu"] = [x, y]

        elif stage == "RED":
            calling["RED"] = [x, y]

        elif stage == "WHITE":
            calling["WHITE"] = [x, y]

        elif stage == "ROUTE_INIT":
            calling["ROUTE_INIT"] = [x, y]
    # =====================================================
    # LC GATE
    # =====================================================

    elif item_type == "LC":

        if item_name not in lc_gates:

            lc_gates[item_name] = {

                "menu": None,

                "IN": None,

                "OUT": None

            }

        gate = lc_gates[item_name]

        if stage == "MENU":
            gate["menu"] = [x, y]

        elif stage == "IN":
            gate["IN"] = [x, y]

        elif stage == "OUT":
            gate["OUT"] = [x, y]

    # =====================================================
    # LC GATE
    # =====================================================

    elif item_type == "LC":

        if item_name not in lc_gates:

            lc_gates[item_name] = {

                "menu": None,

                "IN": None,

                "OUT": None

            }

        gate = lc_gates[item_name]

        if stage == "MENU":
            gate["menu"] = [x, y]

        elif stage == "IN":
            gate["IN"] = [x, y]

        elif stage == "OUT":
            gate["OUT"] = [x, y]
    # =====================================================
    # POINT
    # =====================================================

    elif item_type == "POINT":

        if item_name not in points:

            points[item_name] = {

                "menu": None,

                "NORMAL": None,

                "REVERSE": None,

                "FREE": None

            }

        if stage == "MENU":
            points[item_name]["menu"] = [x, y]

        elif stage == "NORMAL":
            points[item_name]["NORMAL"] = [x, y]

        elif stage == "REVERSE":
            points[item_name]["REVERSE"] = [x, y]

        elif stage == "FREE":
            points[item_name]["FREE"] = [x, y]

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif item_type == "CH":

        if item_name not in crank_handles:

            crank_handles[item_name] = {

                "menu": None,

                "IN": None,

                "OUT": None,

                "ECH": None,

                "FREE": None

            }

        if stage == "MENU":
            crank_handles[item_name]["menu"] = [x, y]

        elif stage == "IN":
            crank_handles[item_name]["IN"] = [x, y]

        elif stage == "OUT":
            crank_handles[item_name]["OUT"] = [x, y]

        elif stage == "ECH":
            crank_handles[item_name]["ECH"] = [x, y]

        elif stage == "FREE":
            crank_handles[item_name]["FREE"] = [x, y]

    log("----------------------------------------")
    log(f"{item_name}")
    log(f"{stage}")
    log(f"Captured : ({x}, {y})")
    log("----------------------------------------")

    capture_index += 1

    overlay_progress_label.config(
        text=f"{capture_index}/{len(capture_master_list)}"
    )

    master_next_capture()
# =========================================================
# CLEAR STORED COORDINATE
# =========================================================

def clear_current_capture(item_type, item_name, stage):

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if item_type == "MAIN":

        if item_name not in signals:
            return

        if stage == "MENU":
            signals[item_name]["menu"] = None

        elif stage == "RED":
            signals[item_name]["RED"] = None

        elif stage == "YELLOW":
            signals[item_name]["YELLOW"] = None

        elif stage == "DOUBLE_YELLOW":
            signals[item_name]["DOUBLE_YELLOW"] = None

        elif stage == "GREEN":
            signals[item_name]["GREEN"] = None

        elif stage == "ROUTE_INIT":
            signals[item_name]["ROUTE_INIT"] = None

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif item_type == "SHUNT":

        if item_name not in shunt_signals:
            return

        if stage == "MENU":
            shunt_signals[item_name]["menu"] = None

        elif stage == "RED":
            shunt_signals[item_name]["RED"] = None

        elif stage == "WHITE":
            shunt_signals[item_name]["WHITE"] = None

        elif stage == "ROUTE_INIT":
            shunt_signals[item_name]["ROUTE_INIT"] = None

    # =====================================================
    # CALLING ON SIGNAL
    # =====================================================

    elif item_type == "CAL":

        if item_name not in calling_on_signals:
            return

        if stage == "MENU":
            calling_on_signals[item_name]["menu"] = None

        elif stage == "RED":
            calling_on_signals[item_name]["RED"] = None

        elif stage == "WHITE":
            calling_on_signals[item_name]["WHITE"] = None

        elif stage == "ROUTE_INIT":
            calling_on_signals[item_name]["ROUTE_INIT"] = None

    # =====================================================
    # LC GATE
    # =====================================================

    elif item_type == "LC":

        if item_name not in lc_gates:
            return

        if stage == "MENU":
            lc_gates[item_name]["menu"] = None

        elif stage == "IN":
            lc_gates[item_name]["IN"] = None

        elif stage == "OUT":
            lc_gates[item_name]["OUT"] = None
    # =====================================================
    # POINT
    # =====================================================

    elif item_type == "POINT":

        if item_name not in points:
            return

        if stage == "MENU":
            points[item_name]["menu"] = None

        elif stage == "NORMAL":
            points[item_name]["NORMAL"] = None

        elif stage == "REVERSE":
            points[item_name]["REVERSE"] = None

        elif stage == "FREE":
            points[item_name]["FREE"] = None

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif item_type == "CH":

        if item_name not in crank_handles:
            return

        if stage == "MENU":
            crank_handles[item_name]["menu"] = None

        elif stage == "IN":
            crank_handles[item_name]["IN"] = None

        elif stage == "OUT":
            crank_handles[item_name]["OUT"] = None

        elif stage == "ECH":
            crank_handles[item_name]["ECH"] = None

        elif stage == "FREE":
            crank_handles[item_name]["FREE"] = None

# =========================================================
# UNDO LAST CAPTURE
# =========================================================

def undo_last_capture():

    global capture_index

    if capture_index == 0:

        log("Already at first capture")

        return

    capture_index -= 1

    item_type, item_name, stage = capture_master_list[capture_index]

    clear_current_capture(
        item_type,
        item_name,
        stage
    )

    log("----------------------------------------")
    log(f"UNDO : {item_name}")
    log(f"STEP : {stage}")
    log("----------------------------------------")

    master_next_capture()
# =========================================================
# CANCEL MASTER CAPTURE
# =========================================================

def cancel_master_capture():

    global capture_module

    capture_module = None

    if capture_overlay:

        capture_overlay.destroy()

    root.deiconify()

    log("Coordinate Capture Cancelled")

# =========================================================
# UPDATE OVERLAY
# =========================================================

def update_capture_overlay(item_type, item_name, stage):

    overlay_progress_label.config(
        text=f"STEP {capture_index + 1} OF {len(capture_master_list)}"
    )
    remaining = len(capture_master_list) - capture_index - 1

    overlay_signal_label.config(
        text=f"{item_name}",
        fg=TYPE_COLORS[item_type]
    )
    overlay_step_label.config(
        text=f"CLICK : {stage.replace('_', ' ')}"
    )

    overlay_hint_label.config(
        text=(
            f"Move the mouse onto this\n"
            f"{stage.replace('_', ' ')} location,\n"
            f"then press SPACE.\n\n"
            f"(or press BACKSPACE)"
        )
    )
# =========================================================
# CREATE CAPTURE OVERLAY
# =========================================================

def create_capture_overlay():

    global capture_overlay

    global overlay_progress_label
    global overlay_signal_label
    global overlay_step_label
    global overlay_hint_label
    global overlay_button_frame

    if capture_overlay:

        try:
            capture_overlay.destroy()
        except:
            pass

    capture_overlay = tk.Toplevel()

    capture_overlay.title("Coordinate Capture")

    capture_overlay.geometry("420x390")

    capture_overlay.configure(bg="#0f172a")

    capture_overlay.attributes("-topmost", True)

    capture_overlay.resizable(False, False)

    capture_overlay.protocol(
        "WM_DELETE_WINDOW",
        lambda: None
    )
    # ==========================================
    # TITLE
    # ==========================================

    tk.Label(

        capture_overlay,

        text="COORDINATE CAPTURE",

        bg="#0f172a",

        fg="white",

        font=("Segoe UI",15,"bold")

    ).pack(pady=(15,8))

    overlay_progress_label = tk.Label(

        capture_overlay,

        text="0 / 0",

        bg="#0f172a",

        fg="#22c55e",

        font=("Segoe UI",13,"bold")

    )

    overlay_progress_label.pack()

    overlay_signal_label = tk.Label(

        capture_overlay,

        text="",

        bg="#0f172a",

        fg="#3b82f6",

        font=("Segoe UI",18,"bold")

    )

    overlay_signal_label.pack(pady=(8,2))
    overlay_step_label = tk.Label(

        capture_overlay,

        text="",

        bg="#0f172a",

        fg="white",

        font=("Segoe UI",13)

    )

    overlay_step_label.pack(pady=(5,5))
    overlay_hint_label = tk.Label(

        capture_overlay,

        text="Move mouse\nPress SPACE",

        bg="#0f172a",

        fg="#facc15",

        font=("Segoe UI",11)

    )

    overlay_hint_label.pack(pady=(5,5))
    overlay_button_frame = tk.Frame(

        capture_overlay,

        bg="#0f172a"

    )

    overlay_button_frame.pack(pady=(10, 15))

    tk.Button(
        overlay_button_frame,
        text="↩ Undo",
        bg="#f59e0b",
        fg="white",
        relief="flat",
        width=12,
        font=("Segoe UI", 10, "bold"),
        command=undo_last_capture
    ).pack(side="left", padx=8)

    tk.Button(
        overlay_button_frame,
        text="✖ Cancel",
        bg="#dc2626",
        fg="white",
        relief="flat",
        width=12,
        font=("Segoe UI", 10, "bold"),
        command=cancel_master_capture
    ).pack(side="left", padx=8)
# =========================================================
# CUSTOM NUMBER DIALOG
# =========================================================

def show_quantity_dialog(title, question):

    result = {"value": None}

    win = tk.Toplevel(root)

    win.title(title)

    win.geometry("500x340")

    win.configure(bg="#0f172a")

    win.resizable(False, False)

    win.transient(root)

    win.grab_set()

    win.attributes("-topmost", True)

    tk.Label(
        win,
        text=title.upper(),
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(25,10))

    tk.Label(
        win,
        text=question,
        bg="#0f172a",
        fg="white",
        font=("Segoe UI",18,"bold"),
        justify="center"
    ).pack()

    entry = tk.Entry(
        win,
        font=("Segoe UI",20),
        justify="center",
        width=5,
        bg="#1e293b",
        fg="white",
        insertbackground="white",
        relief="flat"
    )

    entry.pack(pady=20)

    entry.focus()

    def confirm():

        try:

            result["value"] = int(entry.get())

        except:

            messagebox.showerror(
                "Invalid",
                "Enter a valid number.",
                parent=win
            )

            return

        win.destroy()

    def skip():

        result["value"] = 0

        win.destroy()

    button_frame = tk.Frame(
        win,
        bg="#0f172a"
    )

    button_frame.pack(pady=10)

    tk.Button(
        button_frame,
        text="CONFIRM",
        command=confirm,
        bg="#2563eb",
        fg="white",
        relief="flat",
        width=12,
        font=("Segoe UI",10,"bold")
    ).pack(side="left", padx=5)

    tk.Button(
        button_frame,
        text="SKIP (0)",
        command=skip,
        bg="#0f172a",
        fg="#94a3b8",
        relief="flat",
        width=10
    ).pack(side="left")

    win.wait_window()

    return result["value"]

# =========================================================
# CUSTOM NAME DIALOG
# =========================================================

def show_name_dialog(title, question):

    result = {"value": None}

    win = tk.Toplevel(root)

    win.title(title)

    win.geometry("500x330")

    win.configure(bg="#0f172a")

    win.resizable(False, False)

    win.transient(root)

    win.grab_set()

    win.attributes("-topmost", True)

    tk.Label(
        win,
        text=title.upper(),
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI",14,"bold")
    ).pack(pady=(25,10))

    tk.Label(
        win,
        text=question,
        bg="#0f172a",
        fg="white",
        font=("Segoe UI",18,"bold"),
        justify="center"
    ).pack()

    entry = tk.Entry(
        win,
        font=("Segoe UI",18),
        justify="center",
        width=12,
        bg="#1e293b",
        fg="white",
        insertbackground="white",
        relief="flat"
    )

    entry.pack(pady=20)

    entry.focus()

    def next_step():

        value = entry.get().strip().upper()

        if value == "":
            messagebox.showerror(
                "Required",
                "Please enter a name.",
                parent=win
            )
            return

        result["value"] = value

        win.destroy()

    tk.Button(
        win,
        text="NEXT",
        command=next_step,
        bg="#2563eb",
        fg="white",
        relief="flat",
        width=18,
        font=("Segoe UI",10,"bold")
    ).pack(pady=10)

    win.wait_window()

    return result["value"]

# =========================================================
# CUSTOM ASPECT DIALOG
# =========================================================

def show_aspect_dialog(signal_name):

    result = {"value": None}

    win = tk.Toplevel(root)

    win.title("Signal Aspects")
    win.geometry("500x380")
    win.configure(bg="#0f172a")

    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    win.attributes("-topmost", True)

    tk.Label(
        win,
        text="SELECT ASPECTS",
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI",14,"bold")
    ).pack(pady=(20,10))

    tk.Label(
        win,
        text=f"Which aspect does\n{signal_name}\nhave?",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI",18,"bold"),
        justify="center"
    ).pack(pady=(0,20))

    def choose(value):
        result["value"] = value
        win.destroy()

    for value in (2, 3, 4):

        tk.Button(
            win,
            text=f"{value} ASPECTS",
            command=lambda v=value: choose(v),
            bg="#2563eb",
            fg="white",
            relief="flat",
            width=22,
            height=2,
            font=("Segoe UI",11,"bold")
        ).pack(pady=6)

    win.wait_window()

    return result["value"]

# =========================================================
# CREATE MASTER SETUP
# =========================================================

def create_setup():

    global signals
    global shunt_signals
    global starter_signals
    global calling_on_signals
    global lc_gates
    global points
    global crank_handles

    global capture_master_list
    global capture_index
    global main_signal_aspects

    signals = {}
    shunt_signals = {}
    starter_signals = {}
    calling_on_signals = {}
    lc_gates = {}
    points = {}
    crank_handles = {}

    capture_master_list = []
    capture_index = 0
    main_signal_aspects = {}

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    total_main = show_quantity_dialog(
        "How Many?",
        "How many MAIN signals are in\nthis yard?"
    )

    if total_main is None:
        return

    for i in range(total_main):

        signal = show_name_dialog(
            "Main Signal",
            f"Enter MAIN Signal {i + 1} Name"
        )

        if not signal:
            continue

        signal = signal.strip().upper()

        if signal in main_signal_aspects:
            messagebox.showerror(
                "Duplicate",
                f"{signal} already exists."
            )
            return

        aspects = show_aspect_dialog(signal)

        if aspects is None:
            continue

        main_signal_aspects[signal] = aspects

        capture_master_list.append(("MAIN", signal, "MENU"))
        capture_master_list.append(("MAIN", signal, "RED"))

        if aspects >= 3:
            capture_master_list.append(("MAIN", signal, "YELLOW"))

        if aspects == 4:
            capture_master_list.append(("MAIN", signal, "DOUBLE_YELLOW"))

        capture_master_list.append(("MAIN", signal, "GREEN"))
        capture_master_list.append(("MAIN", signal, "ROUTE_INIT"))

    # =====================================================
    # SHUNT SIGNALS
    # =====================================================

    total_shunt = show_quantity_dialog(
        "Shunt Signals",
        "How many SHUNT signals\nare available?"
    )

    if total_shunt is None:
        return

    for i in range(total_shunt):

        signal = show_name_dialog(
            "Shunt Signal",
            f"Enter SHUNT Signal {i + 1} Name"
        )

        if not signal:
            continue

        signal = signal.strip().upper()

        if signal in shunt_signals:
            messagebox.showerror(
                "Duplicate",
                f"{signal} already exists."
            )
            return

        capture_master_list.append(("SHUNT", signal, "MENU"))
        capture_master_list.append(("SHUNT", signal, "RED"))
        capture_master_list.append(("SHUNT", signal, "WHITE"))
        capture_master_list.append(("SHUNT", signal, "ROUTE_INIT"))

    # =====================================================
    # CALLING-ON SIGNALS
    # =====================================================

    total_calling = show_quantity_dialog(
        "Calling-On Signals",
        "How many CALLING-ON\nsignals are available?"
    )

    if total_calling is None:
        return

    for i in range(total_calling):

        signal = show_name_dialog(
            "Calling-On Signal",
            f"Enter CALLING-ON Signal {i + 1} Name"
        )

        if not signal:
            continue

        signal = signal.strip().upper()

        if signal in calling_on_signals:
            messagebox.showerror(
                "Duplicate",
                f"{signal} already exists."
            )
            return

        capture_master_list.append(("CAL", signal, "MENU"))
        capture_master_list.append(("CAL", signal, "RED"))
        capture_master_list.append(("CAL", signal, "WHITE"))
        capture_master_list.append(("CAL", signal, "ROUTE_INIT"))

    # =====================================================
    # LC GATES
    # =====================================================

    total_lc = show_quantity_dialog(
        "LC Gates",
        "How many LC GATES\nare available?"
    )

    if total_lc is None:
        return

    for i in range(total_lc):

        gate = show_name_dialog(
            "LC Gate",
            f"Enter LC Gate {i + 1} Name"
        )

        if not gate:
            continue

        gate = gate.strip().upper()

        if gate in lc_gates:
            messagebox.showerror(
                "Duplicate",
                f"{gate} already exists."
            )
            return

        capture_master_list.append(("LC", gate, "MENU"))
        capture_master_list.append(("LC", gate, "IN"))
        capture_master_list.append(("LC", gate, "OUT"))

    # =====================================================
    # POINTS
    # =====================================================

    total_points = show_quantity_dialog(
        "Points",
        "How many POINTS are\navailable?"
    )

    if total_points is None:
        return

    for i in range(total_points):

        point = show_name_dialog(
            "Point",
            f"Enter Point {i + 1} Name"
        )

        if not point:
            continue

        point = point.strip().upper()

        if point in points:
            messagebox.showerror(
                "Duplicate",
                f"{point} already exists."
            )
            return

        capture_master_list.append(("POINT", point, "MENU"))
        capture_master_list.append(("POINT", point, "NORMAL"))
        capture_master_list.append(("POINT", point, "REVERSE"))
        capture_master_list.append(("POINT", point, "FREE"))

    # =====================================================
    # CRANK HANDLES
    # =====================================================

    total_crank = show_quantity_dialog(
        "Crank Handle",
        "How many CRANK HANDLES\nare available?"
    )

    if total_crank is None:
        return

    for i in range(total_crank):

        crank = show_name_dialog(
            "Crank Handle",
            f"Enter Crank Handle {i + 1} Name"
        )

        if not crank:
            continue

        crank = crank.strip().upper()

        if crank in crank_handles:
            messagebox.showerror(
                "Duplicate",
                f"{crank} already exists."
            )
            return

        capture_master_list.append(("CH", crank, "MENU"))
        capture_master_list.append(("CH", crank, "IN"))
        capture_master_list.append(("CH", crank, "OUT"))
        capture_master_list.append(("CH", crank, "ECH"))
        capture_master_list.append(("CH", crank, "FREE"))

    # =====================================================
    # START CAPTURE
    # =====================================================

    log("========================================")
    log(f"TOTAL CAPTURES : {len(capture_master_list)}")
    log("Starting Master Capture")
    log("========================================")

    start_master_capture()
# ============================================
# RECORD MAIN SIGNAL
# ============================================

def record_main_signal(index):

    global signals

    log("")
    log("========================================")
    log(f"MAIN SIGNAL {index}")
    log("========================================")

    signal = simpledialog.askstring(
        "MAIN SIGNAL",
        f"Enter Main Signal {index} Name"
    )

    if not signal:
        return

    signal = signal.strip().upper()

    aspects = simpledialog.askinteger(
        "ASPECTS",
        f"How many Aspects for {signal} ?\n\n"
        "2 = Red + Green\n"
        "3 = Red + Yellow + Green\n"
        "4 = Red + Yellow + Double Yellow + Green"
    )

    if aspects is None:
        return

    # ---------------- MENU ----------------

    messagebox.showinfo(
        "MENU",
        f"Click MENU of Signal {signal}"
    )

    menu_x, menu_y = capture_mouse_position()

    # ---------------- RED ----------------

    messagebox.showinfo(
        "RED",
        f"Click RED Lamp of {signal}"
    )

    red_x, red_y = capture_mouse_position()

    yellow = None
    double_yellow = None

    # ---------------- YELLOW ----------------

    if aspects >= 3:

        messagebox.showinfo(
            "YELLOW",
            f"Click YELLOW Lamp of {signal}"
        )

        yx, yy = capture_mouse_position()

        yellow = [yx, yy]

    # ---------------- DOUBLE YELLOW ----------------

    if aspects == 4:

        messagebox.showinfo(
            "DOUBLE YELLOW",
            f"Click DOUBLE YELLOW Lamp of {signal}"
        )

        dx, dy = capture_mouse_position()

        double_yellow = [dx, dy]

    # ---------------- GREEN ----------------

    messagebox.showinfo(
        "GREEN",
        f"Click GREEN Lamp of {signal}"
    )

    green_x, green_y = capture_mouse_position()

    # ---------------- ROUTE INIT ----------------

    messagebox.showinfo(
        "ROUTE INITIALIZATION",
        f"Click Route Initialization point of {signal}"
    )

    route_x, route_y = capture_mouse_position()

    # ---------------- STORE ----------------

    signals[signal] = {

        "type": "MAIN",

        "aspects": aspects,

        "menu": [menu_x, menu_y],

        "RED": [red_x, red_y],

        "YELLOW": yellow,

        "DOUBLE_YELLOW": double_yellow,

        "GREEN": [green_x, green_y],

        "ROUTE_INIT": [route_x, route_y]

    }

    log(f"{signal} Recorded Successfully")


# ============================================
# RECORD POINT
# ============================================

def record_point(index):

    global points

    log("")
    log("========================================")
    log(f"POINT {index}")
    log("========================================")

    point = simpledialog.askstring(
        "POINT",
        f"Enter Point {index} Name"
    )

    if not point:
        return

    # -----------------------------------------
    # POINT MENU
    # -----------------------------------------

    messagebox.showinfo(
        "POINT",
        f"Click Point Menu of {point}"
    )

    menu_x, menu_y = capture_mouse_position()

    # -----------------------------------------
    # NORMAL
    # -----------------------------------------

    messagebox.showinfo(
        "NORMAL",
        f"Click NORMAL Lamp of {point}"
    )

    normal_x, normal_y = capture_mouse_position()

    # -----------------------------------------
    # REVERSE
    # -----------------------------------------

    messagebox.showinfo(
        "REVERSE",
        f"Click REVERSE Lamp of {point}"
    )

    reverse_x, reverse_y = capture_mouse_position()

    # -----------------------------------------
    # FREE
    # -----------------------------------------

    messagebox.showinfo(
        "FREE",
        f"Click FREE indication of {point}"
    )

    free_x, free_y = capture_mouse_position()

    # -----------------------------------------
    # STORE
    # -----------------------------------------

    points[point] = {

        "menu": [menu_x, menu_y],

        "NORMAL": [normal_x, normal_y],

        "REVERSE": [reverse_x, reverse_y],

        "FREE": [free_x, free_y]

    }

    log(f"{point} Recorded Successfully")

def capture_coordinate(title, instruction):

    messagebox.showinfo(
        title,
        instruction + "\n\nPress SPACE to record."
    )

    # Wait until previous SPACE is released
    while keyboard.is_pressed("space"):
        time.sleep(0.05)

    # Wait for new SPACE press
    while not keyboard.is_pressed("space"):
        time.sleep(0.05)

    x, y = pyautogui.position()

    log(f"{title} Recorded : ({x}, {y})")

    # Wait until SPACE is released again
    while keyboard.is_pressed("space"):
        time.sleep(0.05)

    messagebox.showinfo(
        "Captured",
        f"{title} Recorded Successfully.\n\nNext Step..."
    )

    return x, y

# ============================================
# RECORD CRANK HANDLE
# ============================================

def record_crank_handle(index):

    global crank_handles

    log("")
    log("========================================")
    log(f"CRANK HANDLE {index}")
    log("========================================")

    crank = simpledialog.askstring(
        "CRANK HANDLE",
        f"Enter Crank Handle {index} Name"
    )

    if not crank:
        return

    crank = crank.strip().upper()

    messagebox.showinfo(
        "MENU",
        f"Click MENU of Crank Handle {crank}"
    )

    menu_x, menu_y = capture_mouse_position()

    messagebox.showinfo(
        "IN",
        f"Click IN indication of {crank}"
    )

    in_x, in_y = capture_mouse_position()

    messagebox.showinfo(
        "OUT",
        f"Click OUT indication of {crank}"
    )

    out_x, out_y = capture_mouse_position()

    messagebox.showinfo(
        "ECH",
        f"Click ECH indication of {crank}"
    )

    ech_x, ech_y = capture_mouse_position()

    messagebox.showinfo(
        "FREE",
        f"Click FREE indication of {crank}"
    )

    free_x, free_y = capture_mouse_position()

    crank_handles[crank] = {

        "menu": [menu_x, menu_y],

        "IN": [in_x, in_y],

        "OUT": [out_x, out_y],

        "ECH": [ech_x, ech_y],

        "FREE": [free_x, free_y]

    }

    log(f"{crank} Recorded Successfully")

# =========================================================
# KEYBOARD
# =========================================================

def capture_space():

    global capture_module
    global last_space_time

    if capture_module != "MASTER":
        return

    now = time.time()

    if now - last_space_time < 0.25:
        return

    last_space_time = now

    if capture_index >= len(capture_master_list):
        return

    x, y = win32api.GetCursorPos()

    master_save_point(x, y)


def capture_backspace():

    global last_backspace_time

    if capture_module != "MASTER":
        return

    now = time.time()

    if now - last_backspace_time < 0.25:
        return

    last_backspace_time = now

    undo_last_capture()


# def aspect_two():
#
#     if capture_module != "MASTER":
#         return
#
#     if record_stage == "ask_aspects":
#
#         set_aspects_and_continue(2)
#
#
# def aspect_three():
#
#     if capture_module != "MASTER":
#         return
#
#     if record_stage == "ask_aspects":
#
#         set_aspects_and_continue(3)
#
#
# def aspect_four():
#
#     if capture_module != "MASTER":
#         return
#
#     if record_stage == "ask_aspects":
#
#         set_aspects_and_continue(4)


keyboard.add_hotkey("space", capture_space)

keyboard.add_hotkey("backspace", capture_backspace)

# keyboard.add_hotkey("2", aspect_two)

# keyboard.add_hotkey("3", aspect_three)

# keyboard.add_hotkey("4", aspect_four)


# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():

    global signals
    global points
    global crank_handles

    log(f"Signals : {len(signals)}")
    log(f"Points  : {len(points)}")
    log(f"Cranks  : {len(crank_handles)}")

    wb = Workbook()

    # =====================================================
    # MAIN
    # =====================================================

    ws_main = wb.active
    ws_main.title = "MAIN"

    ws_main.append([
        "Signal",
        "Aspects",
        "Menu_X",
        "Menu_Y",
        "Red_X",
        "Red_Y",
        "Yellow_X",
        "Yellow_Y",
        "DoubleYellow_X",
        "DoubleYellow_Y",
        "Green_X",
        "Green_Y",
        "RouteInit_X",
        "RouteInit_Y"
    ])

    for signal, data in signals.items():
        menu = data["menu"]

        red = data["RED"]

        yellow = data["YELLOW"]

        dy = data["DOUBLE_YELLOW"]

        green = data["GREEN"]

        route = data["ROUTE_INIT"]

        ws_main.append([

            signal,

            data["aspects"],

            menu[0], menu[1],

            red[0], red[1],

            yellow[0] if yellow else "",
            yellow[1] if yellow else "",

            dy[0] if dy else "",
            dy[1] if dy else "",

            green[0], green[1],

            route[0] if route else "",
            route[1] if route else ""

        ])
    ws_shunt = wb.create_sheet("SHUNT")

    ws_shunt.append([
        "Signal",
        "Menu_X", "Menu_Y",
        "Red_X", "Red_Y",
        "White_X", "White_Y",
        "RouteInit_X", "RouteInit_Y"
    ])

    for signal, data in shunt_signals.items():
        menu = data["menu"]
        red = data["RED"]
        white = data["WHITE"]
        route = data["ROUTE_INIT"]

        ws_shunt.append([

            signal,

            menu[0], menu[1],

            red[0], red[1],

            white[0], white[1],

            route[0], route[1]

        ])

    # =====================================================
    # CALLING ON SIGNALS
    # =====================================================

    ws_calling = wb.create_sheet("CALLING_ON")

    ws_calling.append([
        "Signal",
        "Menu_X",
        "Menu_Y",
        "Red_X",
        "Red_Y",
        "White_X",
        "White_Y",
        "RouteInit_X",
        "RouteInit_Y"
    ])

    # =====================================================
    # SAVE CALLING ON SIGNALS
    # =====================================================

    for signal, data in calling_on_signals.items():
        menu = data.get("menu") or [None, None]
        red = data.get("RED") or [None, None]
        white = data.get("WHITE") or [None, None]
        route = data.get("ROUTE_INIT") or [None, None]

        ws_calling.append([
            signal,

            menu[0], menu[1],

            red[0], red[1],

            white[0], white[1],

            route[0], route[1]
        ])

    # =====================================================
    # LC GATES
    # =====================================================

    ws_lc = wb.create_sheet("LC")

    ws_lc.append([
        "LCGate",
        "Menu_X",
        "Menu_Y",
        "IN_X",
        "IN_Y",
        "OUT_X",
        "OUT_Y"
    ])
    # =====================================================
    # SAVE LC GATES
    # =====================================================

    for gate, data in lc_gates.items():
        menu = data.get("menu") or [None, None]
        in_pos = data.get("IN") or [None, None]
        out_pos = data.get("OUT") or [None, None]

        ws_lc.append([
            gate,
            menu[0], menu[1],
            in_pos[0], in_pos[1],
            out_pos[0], out_pos[1]
        ])
    # =====================================================
    # POINT
    # =====================================================

    ws_point = wb.create_sheet("POINT")

    ws_point.append([
        "Point",
        "Menu_X",
        "Menu_Y",
        "Normal_X",
        "Normal_Y",
        "Reverse_X",
        "Reverse_Y",
        "Free_X",
        "Free_Y"
    ])

    for point, data in points.items():
        ws_point.append([

            point,

            data["menu"][0],
            data["menu"][1],

            data["NORMAL"][0],
            data["NORMAL"][1],

            data["REVERSE"][0],
            data["REVERSE"][1],

            data["FREE"][0],
            data["FREE"][1]

        ])

    # =====================================================
    # CH
    # =====================================================

    ws_ch = wb.create_sheet("CH")

    ws_ch.append([
        "CrankHandle",
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

    for crank, data in crank_handles.items():
        menu = data["menu"]

        IN = data["IN"]

        OUT = data["OUT"]

        ECH = data["ECH"]

        FREE = data["FREE"]

        ws_ch.append([

            crank,

            menu[0], menu[1],

            IN[0], IN[1],

            OUT[0], OUT[1],

            ECH[0], ECH[1],

            FREE[0], FREE[1]

        ])

    # =====================================================
    # SAVE
    # =====================================================

    config_file_path = os.path.join(
        os.getcwd(),
        "HAH COORDINATES.xlsx"
    )

    wb.save(config_file_path)
    wb.close()

    log("========================================")
    log("CONFIG SAVED SUCCESSFULLY")
    log(f"Location          : {config_file_path}")
    log(f"Main Signals       : {len(signals)}")
    log(f"Shunt Signals      : {len(shunt_signals)}")
    log(f"Calling-On Signals : {len(calling_on_signals)}")
    log(f"LC Gates           : {len(lc_gates)}")
    log(f"Points             : {len(points)}")
    log(f"Crank Handles      : {len(crank_handles)}")
    log("========================================")

    messagebox.showinfo(
        "CONFIGURATION",
        f"Configuration Saved Successfully.\n\n"
        f"Signals : {len(signals)}\n"
        f"Points : {len(points)}\n"
        f"Crank Handles : {len(crank_handles)}"
    )
# =========================================================
# SHOW CAPTURE SUMMARY
# =========================================================

def show_capture_summary():

    messagebox.showinfo(

        "CAPTURE SUMMARY",

        f"""
Main Signals   : {len(signals)}

Points         : {len(points)}

Crank Handles  : {len(crank_handles)}

Configuration Saved Successfully.
"""
    )

# =========================================================
# LOAD CONFIG
# =========================================================

def load_config():

    global signals
    global CONFIG_FILE

    if MASTER_MODE:
        master_coords = os.environ.get("EDRC_COORDS", "").strip()
        if master_coords and os.path.exists(master_coords):
            CONFIG_FILE = master_coords
            log(f"MASTER MODE CONFIG PATH : {CONFIG_FILE}")
    global shunt_signals
    global starter_signals
    global calling_on_signals
    global lc_gates
    global points
    global crank_handles

    signals = {}
    shunt_signals = {}
    calling_on_signals = {}
    lc_gates = {}
    points = {}
    crank_handles = {}
    starter_signals = {}

    if not os.path.exists(CONFIG_FILE):

        log("CONFIG FILE NOT FOUND")
        return

    try:

        wb = load_workbook(CONFIG_FILE, data_only=True)

        # =====================================================
        # MAIN SIGNALS
        # =====================================================

        if "MAIN" in wb.sheetnames:

            ws = wb["MAIN"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                signals[signal] = {

                    "type": "MAIN",

                    "aspects": row[1],

                    "menu": [int(row[2]), int(row[3])],

                    "RED": [int(row[4]), int(row[5])]
                    if row[4] is not None and row[5] is not None
                    else None,

                    "YELLOW": (
                        [int(row[6]), int(row[7])]
                        if row[6] not in (None, "") and row[7] not in (None, "")
                        else None
                    ),

                    "DOUBLE_YELLOW": (
                        [int(row[8]), int(row[9])]
                        if row[8] not in (None, "") and row[9] not in (None, "")
                        else None
                    ),

                    "GREEN": [int(row[10]), int(row[11])]
                    if row[10] is not None and row[11] is not None
                    else None,

                    "ROUTE_INIT": (
                        [int(row[12]), int(row[13])]
                        if row[12] not in (None, "") and row[13] not in (None, "")
                        else None
                    )

                }

        # =====================================================
        # STARTER SIGNALS
        # =====================================================
        # Starter signals are stored in MAIN in the common coordinate
        # workbook. Reuse the MAIN entries for starter signal lookups.
        starter_signals = dict(signals)

        # SHUNT SIGNALS
        # =====================================================

        if "SHUNT" in wb.sheetnames:

            ws = wb["SHUNT"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                shunt_signals[signal] = {

                    "menu": [int(row[1]), int(row[2])]
                    if row[1] is not None and row[2] is not None
                    else None,

                    "RED": [int(row[3]), int(row[4])]
                    if row[3] is not None and row[4] is not None
                    else None,

                    "WHITE": [int(row[5]), int(row[6])]
                    if row[5] is not None and row[6] is not None
                    else None,

                    "ROUTE_INIT": [int(row[7]), int(row[8])]
                    if len(row) > 8 and row[7] is not None and row[8] is not None
                    else None

                }

        # =====================================================
        # CALLING ON SIGNALS
        # =====================================================

        if "CALLING_ON" in wb.sheetnames:

            ws = wb["CALLING_ON"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                calling_on_signals[signal] = {

                    "menu": [int(row[1]), int(row[2])]
                    if row[1] is not None and row[2] is not None
                    else None,

                    "RED": [int(row[3]), int(row[4])]
                    if row[3] is not None and row[4] is not None
                    else None,

                    "WHITE": [int(row[5]), int(row[6])]
                    if row[5] is not None and row[6] is not None
                    else None,

                    "ROUTE_INIT": [int(row[7]), int(row[8])]
                    if row[7] is not None and row[8] is not None
                    else None

                }
        # =====================================================
        # LC GATES
        # =====================================================

        if "LC" in wb.sheetnames:

            ws = wb["LC"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                gate = str(row[0]).strip().upper()

                lc_gates[gate] = {

                    "menu": [int(row[1]), int(row[2])]
                    if row[1] is not None and row[2] is not None
                    else None,

                    "IN": [int(row[3]), int(row[4])]
                    if row[3] is not None and row[4] is not None
                    else None,

                    "OUT": [int(row[5]), int(row[6])]
                    if row[5] is not None and row[6] is not None
                    else None

                }
        # =====================================================
        # POINTS
        # =====================================================

        if "POINT" in wb.sheetnames:

            ws = wb["POINT"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                point = str(row[0]).strip().upper()

                points[point] = {

                    "menu": [int(row[1]), int(row[2])]
                    if row[1] is not None and row[2] is not None
                    else None,

                    "NORMAL": [int(row[3]), int(row[4])]
                    if row[3] is not None and row[4] is not None
                    else None,

                    "REVERSE": [int(row[5]), int(row[6])]
                    if row[5] is not None and row[6] is not None
                    else None,

                    "FREE": [int(row[7]), int(row[8])]
                    if row[7] is not None and row[8] is not None
                    else None

                }

        # =====================================================
        # CRANK HANDLES
        # =====================================================

        if "CH" in wb.sheetnames:

            ws = wb["CH"]

            for row in ws.iter_rows(min_row=2, values_only=True):

                if not row[0]:
                    continue

                crank = str(row[0]).strip().upper()

                crank_handles[crank] = {

                    "menu": [int(row[1]), int(row[2])]
                    if row[1] is not None and row[2] is not None
                    else None,

                    "IN": [int(row[3]), int(row[4])]
                    if row[3] is not None and row[4] is not None
                    else None,

                    "OUT": [int(row[5]), int(row[6])]
                    if row[5] is not None and row[6] is not None
                    else None,

                    "ECH": [int(row[7]), int(row[8])]
                    if row[7] is not None and row[8] is not None
                    else None,

                    "FREE": [int(row[9]), int(row[10])]
                    if row[9] is not None and row[10] is not None
                    else None

                }

        wb.close()

        log("========================================")
        log("CONFIG LOADED SUCCESSFULLY")
        log(f"Main Signals Loaded       : {len(signals)}")
        log(f"Starter Signals Loaded    : {len(starter_signals)}")
        log(f"Shunt Signals Loaded      : {len(shunt_signals)}")
        log(f"Calling-On Signals Loaded : {len(calling_on_signals)}")
        log(f"LC Gates Loaded           : {len(lc_gates)}")
        log(f"Points Loaded             : {len(points)}")
        log(f"Crank Handles Loaded      : {len(crank_handles)}")
        log("========================================")

    except Exception as e:

        log(f"LOAD CONFIG ERROR : {e}")

def open_existing_config():

    global CONFIG_FILE

    file_path = filedialog.askopenfilename(
        title="Select Coordinate File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    CONFIG_FILE = file_path

    load_config()

    log("========================================")
    log("CONFIGURATION FILE LOADED")
    log(f"File : {CONFIG_FILE}")
    log("========================================")

    messagebox.showinfo(
        "Success",
        "Configuration loaded successfully."
    )

def edit_signal():

    global capture_master_list
    global capture_index
    global undo_stack

    load_config()

    if (
        len(signals) == 0
        and len(shunt_signals) == 0
        and len(calling_on_signals) == 0
        and len(lc_gates) == 0
        and len(points) == 0
        and len(crank_handles) == 0
    ):

        messagebox.showerror(
            "No Configuration",
            "Load configuration first."
        )
        return

    equipment = show_name_dialog(
        "Edit Equipment",
        "Enter Equipment Name"
    )

    if not equipment:
        return

    equipment = equipment.strip().upper()

    capture_master_list = []
    capture_index = 0
    undo_stack.clear()

    # =====================================================
    # MAIN SIGNAL
    # =====================================================

    if equipment in signals:

        capture_master_list.append(("MAIN", equipment, "MENU"))
        capture_master_list.append(("MAIN", equipment, "RED"))

        if signals[equipment]["aspects"] >= 3:
            capture_master_list.append(
                ("MAIN", equipment, "YELLOW")
            )

        if signals[equipment]["aspects"] == 4:
            capture_master_list.append(
                ("MAIN", equipment, "DOUBLE_YELLOW")
            )

        capture_master_list.append(
            ("MAIN", equipment, "GREEN")
        )

        capture_master_list.append(
            ("MAIN", equipment, "ROUTE_INIT")
        )

        start_master_capture()
        return

    # =====================================================
    # SHUNT SIGNAL
    # =====================================================

    elif equipment in shunt_signals:

        capture_master_list.append(("SHUNT", equipment, "MENU"))
        capture_master_list.append(("SHUNT", equipment, "RED"))
        capture_master_list.append(("SHUNT", equipment, "WHITE"))
        capture_master_list.append(("SHUNT", equipment, "ROUTE_INIT"))

        start_master_capture()
        return

    # =====================================================
    # CALLING-ON SIGNAL
    # =====================================================

    elif equipment in calling_on_signals:

        capture_master_list.append(("CAL", equipment, "MENU"))
        capture_master_list.append(("CAL", equipment, "RED"))
        capture_master_list.append(("CAL", equipment, "WHITE"))
        capture_master_list.append(("CAL", equipment, "ROUTE_INIT"))

        start_master_capture()
        return

    # =====================================================
    # LC GATE
    # =====================================================

    elif equipment in lc_gates:

        capture_master_list.append(("LC", equipment, "MENU"))
        capture_master_list.append(("LC", equipment, "IN"))
        capture_master_list.append(("LC", equipment, "OUT"))

        start_master_capture()
        return

    # =====================================================
    # POINT
    # =====================================================

    elif equipment in points:

        capture_master_list.append(("POINT", equipment, "MENU"))
        capture_master_list.append(("POINT", equipment, "NORMAL"))
        capture_master_list.append(("POINT", equipment, "REVERSE"))
        capture_master_list.append(("POINT", equipment, "FREE"))

        start_master_capture()
        return

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    elif equipment in crank_handles:

        capture_master_list.append(("CH", equipment, "MENU"))
        capture_master_list.append(("CH", equipment, "IN"))
        capture_master_list.append(("CH", equipment, "OUT"))
        capture_master_list.append(("CH", equipment, "ECH"))
        capture_master_list.append(("CH", equipment, "FREE"))

        start_master_capture()
        return

    # =====================================================
    # NOT FOUND
    # =====================================================

    messagebox.showerror(
        "Not Found",
        f"{equipment} not found."
    )

def get_main_signal(signal):

    try:
        signal = str(signal).split("_")[0].strip().upper()

        if signal in signals:
            x, y = signals[signal]["menu"]
            return int(x), int(y)

        if signal in starter_signals:
            x, y = starter_signals[signal]["menu"]
            return int(x), int(y)

        log(f"Signal {signal} not found")
        return None, None

    except Exception as e:
        log(f"GET SIGNAL ERROR : {e}")
        return None, None

def get_signal_yellow_coordinates(signal):

    try:

        signal = str(signal).split("_")[0].strip().upper()

        if signal not in signals:

            return None, None

        if signals[signal]["YELLOW"] is None:

            return None, None

        x, y = signals[signal]["YELLOW"]

        return int(x), int(y)

    except Exception as e:

        log(f"GET SIGNAL YELLOW ERROR : {e}")

        return None, None

from pywinauto import Desktop

def get_signal_menu(signal):

    try:

        signal = str(signal).split("_")[0].strip().upper()

        if signal not in signals:

            log(f"{signal} not found in config")

            return None, None

        x, y = signals[signal]["menu"]

        return int(x), int(y)

    except Exception as e:

        log(f"GET SIGNAL MENU ERROR : {e}")

        return None, None

def find_and_click(name, control_type=None):

    log(f"Searching for : {name}")

    try:

        desktop = Desktop(backend="uia")

        log("Searching active popup...")

        # Search only visible top-level windows
        for window in desktop.windows():

            try:

                if not window.is_visible():
                    continue

                # Search descendants only in this visible popup/window
                for control in window.descendants():

                    try:

                        text = control.window_text().strip()

                        if text != str(name).strip():
                            continue

                        if not control.is_visible():
                            continue

                        rect = control.rectangle()

                        if rect.width() <= 0 or rect.height() <= 0:
                            continue

                        x = (rect.left + rect.right) // 2
                        y = (rect.top + rect.bottom) // 2

                        log(f"FOUND : {name}")
                        log(f"Window : {window.window_text()}")
                        log(f"Click : ({x},{y})")

                        pyautogui.moveTo(x, y, duration=0.2)
                        pyautogui.click()

                        time.sleep(0.5)

                        return True

                    except Exception:
                        continue

            except Exception:
                continue

        log(f"{name} Not Found")
        return False

    except Exception as e:

        log(f"find_and_click ERROR : {e}")
        return False

def verify_main_signal(signal):

    log("--------------------------------")
    log(f"VERIFYING MAIN SIGNAL : {signal}")

    if signal not in signals:

        log("Signal not found")

        return False

    info = signals[signal]

    screenshot = pyautogui.screenshot()

    # GREEN

    if info.get("GREEN"):

        x, y = info["GREEN"]

        rgb = screenshot.getpixel((x, y))

        log(f"GREEN RGB : {rgb}")

        if rgb[1] > 180:

            log("Signal Still GREEN")

            return True

    # YELLOW

    if info.get("YELLOW"):

        x, y = info["YELLOW"]

        rgb = screenshot.getpixel((int(x), int(y)))

        log(f"YELLOW RGB : {rgb}")

        if rgb[0] > 180 and rgb[1] > 180:
            log("Signal Still YELLOW")

            return True

    log("Main Signal Disturbed")

    return False

def get_main_signal_aspect(signal):

    signal = str(signal).split("_")[0].strip().upper()

    if signal in signals:
        info = signals[signal]
    elif signal in starter_signals:
        info = starter_signals[signal]
    else:
        log(f"{signal} NOT FOUND IN CONFIG")
        return "UNKNOWN"

    screenshot = pyautogui.screenshot()
    screenshot.save("debug.png")
    print(signal, info)

    # GREEN

    if info.get("GREEN"):

        x, y = info["GREEN"]

        r, g, b = screenshot.getpixel((int(x), int(y)))

        log(f"{signal} GREEN RGB : ({r}, {g}, {b})")

        # Starter Signals
        if info["type"] == "STARTER":

            if g >= 170:
                return "GREEN"

        # Main Signals
        else:
            # 2-aspect signals use the SAME MAIN/GREEN coordinate for
            # both possible indications.  Depending on the signal,
            # that coordinate can illuminate either GREEN or YELLOW.
            # Therefore classify the actual RGB colour at that one
            # coordinate instead of requiring a separate YELLOW point.
            if info.get("aspects") == 2 and not info.get("YELLOW"):
                # Yellow: red and green components are both high and
                # blue is very low.
                if r >= 150 and g >= 150 and b < 80:
                    return "YELLOW"

                # Green: green component is dominant.
                if g > r and g > b and g >= 80:
                    return "GREEN"
            else:
                if g > r and g > b and g >= 80:
                    return "GREEN"
    # YELLOW
    if info.get("YELLOW"):

        x, y = info["YELLOW"]

        r, g, b = screenshot.getpixel((int(x), int(y)))

        log(f"{signal} YELLOW RGB : ({r}, {g}, {b})")

        if r >= 160 and g >= 160:
            return "YELLOW"

    # RED
    if info.get("RED"):

        x, y = info["RED"]

        r, g, b = screenshot.getpixel((int(x), int(y)))

        log(f"{signal} RED RGB : ({r}, {g}, {b})")

        if r > g and r > b and r >= 120:
            return "RED"

    return "UNKNOWN"

def prepare_master_inputs():
    """
    Prepare TOC + coordinates when launched by EDRC_AUTOMATION_SUITE.

    Direct/manual execution is unchanged: the normal GUI buttons still
    call upload_toc() and load_config().  In Master mode, the launcher
    provides EDRC_LIST and EDRC_COORDS and no file picker is required.
    """
    global CONFIG_FILE

    if not MASTER_MODE:
        return

    toc_path = os.environ.get("EDRC_LIST", "").strip()
    coords_path = (
        os.environ.get("EDRC_COORDS", "").strip()
        or toc_path
    )

    log("========================================")
    log("MASTER MODE INPUT PREPARATION")
    log(f"EDRC_LIST  : {toc_path}")
    log(f"EDRC_COORDS: {coords_path}")
    log("========================================")

    if toc_path and os.path.exists(toc_path):
        if not toc_data:
            upload_toc()
        else:
            log(f"TOC ALREADY LOADED : {len(toc_data)} rows")
    else:
        log("MASTER TOC PATH NOT FOUND")

    if coords_path and os.path.exists(coords_path):
        CONFIG_FILE = coords_path
        if not signals and not starter_signals:
            load_config()
        else:
            log(
                f"CONFIG ALREADY LOADED : "
                f"MAIN={len(signals)} STARTER={len(starter_signals)}"
            )
    else:
        log("MASTER COORDINATE PATH NOT FOUND")

    # Give the GUI event loop a moment to complete any dialog callbacks
    # triggered by the launcher.
    for _ in range(20):
        if toc_data and (signals or starter_signals):
            break
        time.sleep(0.25)

    log(
        f"MASTER INPUT STATUS : TOC={len(toc_data)} "
        f"MAIN={len(signals)} STARTER={len(starter_signals)}"
    )



# =========================================================
# THROUGH TESTING REPORT
# =========================================================

def create_report():
    """Create the Through Testing report in the format used by EDRC."""
    global through_report_wb
    global through_report_ws
    global through_report_path
    global through_report_next_row

    try:
        through_report_next_row = 4
        # EDRC Master provides the exact per-program report folder.
        # This is supplied by the EDRC Automation Suite as
        # EDRC_PROGRAM_DIR (for example ...\\03_THROUGH_TESTING).
        # Standalone runs still save beside this file.
        program_dir = os.environ.get("EDRC_PROGRAM_DIR", "").strip()
        session_dir = os.environ.get("EDRC_SESSION_DIR", "").strip()

        if program_dir:
            report_dir = Path(program_dir)
        elif session_dir:
            # Compatibility fallback for older EDRC Suite versions.
            report_dir = Path(session_dir) / "THROUGH_TESTING"
        else:
            report_dir = Path(__file__).resolve().parent

        report_dir.mkdir(parents=True, exist_ok=True)
        log(f"REPORT DIRECTORY : {report_dir}")

        through_report_wb = Workbook()
        through_report_ws = through_report_wb.active
        through_report_ws.title = "Through Testing"

        ws = through_report_ws

        # -----------------------------------------------------
        # REPORT HEADER
        # -----------------------------------------------------
        ws.merge_cells("A1:E1")
        ws["A1"] = "THROUGH TESTING FOR RELATED SIGNALS"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        # Keep row 2 blank, matching the required report format.
        headers = ["Sl.No", "Signal Sequence", "Route Sequence", "Result", "Date & Time"]

        for col, value in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col, value=value)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Borders / alignment.
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows(min_row=3, max_row=3, min_col=1, max_col=5):
            for cell in row:
                cell.border = border

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 25
        ws.column_dimensions["D"].width = 16
        ws.column_dimensions["E"].width = 24

        through_report_path = str(
            report_dir /
            f"THROUGH_TESTING_REPORT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        # Create the file immediately so the Master Suite can discover it.
        through_report_wb.save(through_report_path)

        if not os.path.isfile(through_report_path):
            raise IOError(f"Report file was not created: {through_report_path}")

        log("========================================")
        log("THROUGH TESTING REPORT CREATED")
        log(f"REPORT : {through_report_path}")
        log("========================================")

        return through_report_path

    except Exception as e:
        through_report_wb = None
        through_report_ws = None
        through_report_path = None
        through_report_next_row = 4
        log(f"REPORT CREATION ERROR : {e}")
        return None


def _report_rows():
    """Return ONLY the route-result rows, never the summary area."""
    if through_report_ws is None:
        return []

    # Results are deliberately written to rows 4,5,6,7...
    # The summary is fixed at rows 9-11 and must never be counted.
    end_row = through_report_next_row - 1
    if end_row < 4:
        return []

    return list(range(4, min(end_row, 8) + 1))


def _update_report_summary():
    """Write the three summary lines used by the reference report."""
    ws = through_report_ws
    if ws is None:
        return

    data_rows = _report_rows()

    total = len(data_rows)
    passed_routes = []
    failed_routes = []

    for r in data_rows:
        route_sequence = str(ws.cell(r, 3).value or "")
        first_route = route_sequence.split(",")[0].strip() if route_sequence else ""
        result = str(ws.cell(r, 4).value or "").strip().upper()

        if result == "PASS":
            if first_route:
                passed_routes.append(first_route)
        else:
            if first_route:
                failed_routes.append(first_route)

    # Clear previous summary area.
    for r in range(9, 12):
        for c in range(1, 3):
            ws.cell(r, c).value = None
            ws.cell(r, c).border = Border()

    ws["A9"] = "Total Number of Routes Tested"
    ws["B9"] = total

    ws["A10"] = "Total Number of Routes Passed"
    ws["B10"] = ", ".join(passed_routes) if passed_routes else "NONE"

    ws["A11"] = "Total Number of Routes Failed"
    ws["B11"] = ", ".join(failed_routes) if failed_routes else "NONE"

    for r in range(9, 12):
        for c in range(1, 3):
            cell = ws.cell(r, c)
            cell.border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )
            cell.alignment = Alignment(vertical="center")
            if c == 1:
                cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 35


def save_result(
    slno,
    signal_sequence,
    route_sequence,
    final_result,
    first_signal="",
    signal_results="",
    route_selection="PASS"
):
    """Append one result row and save the report immediately."""
    global through_report_wb
    global through_report_ws
    global through_report_path
    global through_report_next_row

    try:
        if through_report_ws is None:
            if create_report() is None:
                return None

        ws = through_report_ws

        # IMPORTANT: never use ws.max_row because the summary occupies
        # rows 9-11. Route results must stay in rows 4-7.
        row = through_report_next_row

        if row >= 9:
            raise RuntimeError(
                "Too many route-result rows for the report layout."
            )

        result_text = str(final_result or "FAIL").upper()
        ws.cell(row, 1, slno)
        ws.cell(row, 2, signal_sequence)
        ws.cell(row, 3, route_sequence)
        ws.cell(row, 4, result_text)
        ws.cell(row, 5, datetime.now().strftime("%d-%m-%Y %I:%M:%S %p"))

        for col in range(1, 6):
            cell = ws.cell(row, col)
            cell.border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )
            cell.alignment = Alignment(horizontal="center", vertical="center")

        through_report_next_row = row + 1

        _update_report_summary()

        # Save after every row so even a partial run is available.
        through_report_wb.save(through_report_path)

        log(
            f"REPORT ROW SAVED : S.NO={slno} "
            f"RESULT={result_text}"
        )

        return through_report_path

    except Exception as e:
        log(f"REPORT SAVE ERROR : {e}")
        return None


def finalize_report():
    """Finalize and close the Through Testing report."""
    global through_report_wb
    global through_report_ws
    global through_report_path
    global through_report_next_row

    if through_report_wb is None or through_report_path is None:
        log("NO THROUGH TESTING REPORT TO FINALIZE")
        return None

    try:
        _update_report_summary()

        through_report_wb.save(through_report_path)
        through_report_wb.close()

        path = through_report_path

        through_report_wb = None
        through_report_ws = None
        through_report_next_row = 4

        log("========================================")
        log("THROUGH TESTING REPORT SAVED")
        log(f"REPORT : {path}")
        log("========================================")

        return path

    except Exception as e:
        log(f"REPORT FINALIZATION ERROR : {e}")
        return None


# =========================================================
# THROUGH TESTING TIMING
# =========================================================
# These waits are intentionally short but leave enough time for
# the HMI menu to update.  The old 5-second release wait was
# making the final route unnecessarily slow.
SIGNAL_MENU_WAIT = 1.5
ROUTE_ESTABLISH_WAIT = 2.0
BETWEEN_SIGNALS_WAIT = 0.5
CANCEL_MENU_WAIT = 1.0
RELEASE_MENU_WAIT = 1.0
AFTER_ROUTE_RELEASE_WAIT = 1.0
# Required settling time between completion of one TOC route row and the next route selection.
BETWEEN_ROUTES_WAIT = 4.0

def _find_existing_test_panel():
    """Find the already-running Hitachi Test Panel window robustly."""
    try:
        desktop = Desktop(backend="uia")
        candidates = []
        for win in desktop.windows():
            try:
                title = (win.window_text() or "").strip()
                if not title:
                    continue
                visible = True
                try:
                    visible = win.is_visible()
                except Exception:
                    pass
                if not visible:
                    continue
                title_upper = title.upper()
                # Hitachi panel titles can vary between installations.
                if ("TEST PANEL" in title_upper or
                    "HITACHI" in title_upper or
                    "C-FAT" in title_upper or
                    "CFAT" in title_upper):
                    candidates.append(win)
            except Exception:
                continue

        if candidates:
            # Prefer an explicit Test Panel title when several windows match.
            candidates.sort(key=lambda w: "TEST PANEL" not in (w.window_text() or "").upper())
            return candidates[0]

        log("HITACHI TEST PANEL NOT FOUND")
        return None
    except Exception as e:
        log(f"TEST PANEL SEARCH ERROR : {e}")
        return None

def open_test_panel():
    """Bring the existing Hitachi Test Panel to the foreground automatically."""
    panel = _find_existing_test_panel()
    if panel is None:
        log("HITACHI TEST PANEL NOT FOUND")
        if not MASTER_MODE:
            messagebox.showerror("TEST PANEL NOT FOUND", "Please start the Hitachi Test Panel and try again.")
        return False
    try:
        try: panel.restore()
        except Exception: pass
        try: panel.set_focus()
        except Exception: pass
        try: panel.maximize()
        except Exception: pass
        time.sleep(1)
        try: root.iconify()
        except Exception: pass
        log("HITACHI TEST PANEL OPENED / FOCUSED")
        return True
    except Exception as e:
        log(f"OPEN TEST PANEL ERROR : {e}")
        if not MASTER_MODE:
            messagebox.showerror("TEST PANEL ERROR", str(e))
        return False


def test_route_engine():

    global signals
    global starter_signals
    global shunt_signals
    global calling_on_signals
    global points
    global crank_handles
    global lc_gates
    global automation_running
    global running
    log(">>> ENTERED test_route_engine() FROM EDRC MASTER")
    log(">>> PREPARING THROUGH TESTING REPORT")
    # =====================================================
    # MASTER STARTUP GUARD
    # =====================================================
    # AUTOFEED may invoke the start function before the child
    # GUI has finished loading its TOC/configuration.  Never
    # allow the railway automation to start with empty data.
    if automation_running:
        log("TRAIN PASSING THROUGH AUTOMATION ALREADY RUNNING")
        return

    if MASTER_MODE:
        prepare_master_inputs()

    automation_running = True
    running = True

    log("")
    log("================================================")
    log("TRAIN PASSING THROUGH AUTOMATION STARTED")
    log("================================================")

    # -----------------------------------------------------
    # SETTLE TIME AFTER INPUTS ARE LOADED
    # Wait before the first click on the yard, so the Test
    # Panel is ready and the operator can see it start.
    # -----------------------------------------------------
    log("AUTOMATION STARTING IN 2 SECONDS...")

    for i in (2, 1):
        log(f"{i}...")
        time.sleep(1)

    log(f"DEBUG Main Signals     : {len(signals)}")
    log(f"DEBUG Starter Signals : {len(starter_signals)}")
    log(f"DEBUG Signals Dict    : {signals}")


    create_report()

    if not open_test_panel():
        automation_running = False
        running = False
        return

    if len(toc_data) == 0:

        messagebox.showerror(
            "ERROR",
            "TOC NOT LOADED"
        )

        automation_running = False
        return

    # =====================================================
    # LOOP THROUGH TOC
    # =====================================================

    for row_index, item in enumerate(toc_data):

        if not automation_running:
            break

        signal_sequence = item["signal_sequence"]
        route_sequence = item["route_sequence"]

        selected_signals = [

            x.strip()

            for x in signal_sequence.split(",")

            if x.strip()

        ]

        routes = [

            x.strip()

            for x in route_sequence.split(",")

            if x.strip()

        ]
        # =====================================
        # ROW RESULT
        # =====================================

        row_result = "PASS"
        log("")
        log("================================================")
        log(f"SIGNALS : {signals}")
        log(f"ROUTES  : {routes}")
        log("================================================")
        # =====================================================
        # SELECT ALL SIGNALS & ROUTES
        # =====================================================

        selection_failed = False

        for signal, route in zip(selected_signals, routes):

            log("")
            log("----------------------------------------")
            log(f"Signal : {signal}")
            log(f"Route  : {route}")
            log("----------------------------------------")

            # ADD THIS LINE
            log(f"Opening Signal : {signal}")

            x, y = get_main_signal(signal)


            if x is None:
                log(f"{signal} Coordinates Not Found")


                selection_failed = True
                row_result = "FAIL"
                break

            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()

            time.sleep(SIGNAL_MENU_WAIT)

            # -------------------------------------
            # SELECT ROUTE
            # -------------------------------------
            log(f"Selecting Route : {route}")
            if not find_and_click(route):

                log("===================================")
                log(f"CURRENT SIGNAL : {signal}")
                log(f"EXPECTED ROUTE : {route}")
                log("===================================")

                log("Searching route now...")
                log(f"{route} Selection Failed")

                selection_failed = True
                row_result = "FAIL"
                break

            else:

                log(f"{route} Selected Successfully")

            # Wait for the route to establish
            time.sleep(ROUTE_ESTABLISH_WAIT)

            # -------------------------------
            # IMMEDIATE SIGNAL VERIFICATION
            # -------------------------------

            aspect = get_main_signal_aspect(signal)

            log(f"{signal} Immediate Aspect : {aspect}")

            if aspect not in ("GREEN", "YELLOW"):

                log(f"{signal} Immediate Verification Failed")

            else:

                log(f"{signal} Immediate Verification Passed")

            # Small delay before next signal
            time.sleep(BETWEEN_SIGNALS_WAIT)

        if selection_failed:
            log("Selection Failed")
            row_result = "FAIL"
            save_result(
                item["slno"],
                item["signal_sequence"],
                item["route_sequence"],
                row_result,
                route_selection="FAIL"
            )
            continue
        # =====================================================
        # VERIFY ALL SIGNALS
        # =====================================================

        log("")
        log("================================================")
        log("VERIFYING ALL SIGNALS")
        log("================================================")

        verification_failed = False

        for signal in selected_signals:

            aspect = "UNKNOWN"

            # Wait up to 10 seconds for the signal
            for attempt in range(10):

                aspect = get_main_signal_aspect(signal)

                log(f"{signal} Attempt {attempt + 1} : {aspect}")

                if aspect in ("GREEN", "YELLOW"):
                    break

                time.sleep(1)

            log(f"{signal} Final Aspect : {aspect}")

            if aspect not in ("GREEN", "YELLOW"):
                verification_failed = True
                row_result = "FAIL"

        if verification_failed:

            log("One or more signals are neither GREEN nor YELLOW")

        else:

            log("ALL SELECTED SIGNALS ARE GREEN/YELLOW")

            row_result = "PASS"

        if verification_failed:
            log("Signal Verification Failed")

        log("Continuing with Route Release...")

        # =====================================================
        # RELEASE FROM LAST SIGNAL
        # =====================================================

        log("")
        log("================================================")
        log("RELEASING ROUTES")
        log("================================================")

        for signal, route in reversed(list(zip(selected_signals, routes))):

            log("")
            log("----------------------------------------")
            log(f"Signal : {signal}")
            log(f"Route  : {route}")
            log("----------------------------------------")

            x, y = get_main_signal(signal)

            if x is None:
                log(f"{signal} Coordinates Not Found")
                continue

            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()

            time.sleep(SIGNAL_MENU_WAIT)

            if find_and_click("Signal Cancel"):

                log("Signal Cancel Successful")



            else:

                log("Signal Cancel Failed")

            # Allow the HMI to close/update the menu.
            time.sleep(CANCEL_MENU_WAIT)

            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()

            time.sleep(SIGNAL_MENU_WAIT)

            if find_and_click("Route Release"):

                log("Route Released")



            else:

                log("Route Release Failed")

            # Do not wait an unnecessary 5 seconds after every release.
            time.sleep(AFTER_ROUTE_RELEASE_WAIT)



        # =====================================
        # UPDATE EXCEL
        # =====================================
        log("==============================")
        log(f"FINAL ROW RESULT = {row_result}")
        if save_result is not None:
            save_result(
                item["slno"],
                item["signal_sequence"],
                item["route_sequence"],
                row_result,
                route_selection="PASS" if not selection_failed else "FAIL"
            )
        else:
            log("REPORT SAVE SKIPPED: report module unavailable")
        # -----------------------------------------------------
        # 4-SECOND SETTLING DELAY BEFORE NEXT ROUTE ROW
        # -----------------------------------------------------
        if automation_running and row_index < len(toc_data) - 1:
            log("")
            log("========================================")
            log("ROUTE COMPLETED")
            log("WAITING 4 SECONDS BEFORE NEXT ROUTE")
            log("========================================")

            for remaining in range(4, 0, -1):
                log(f"Next route selection in {remaining}...")
                time.sleep(1)

    # =====================================================
    # AUTOMATION COMPLETED
    # =====================================================

    automation_running = False
    running = False

    if finalize_report is not None:
        finalize_report()
    else:
        log("REPORT FINALIZATION SKIPPED: report module unavailable")

    log("")
    log("================================================")
    log("THROUGH TESTING AUTOMATION COMPLETED")
    log("================================================")

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    log(">>> THROUGH TESTING test_route_engine() FINISHED")
# =========================================================
# STOP
# =========================================================

def stop_automation():
    global automation_running
    global running
    global paused

    automation_running = False
    running = False
    paused = False

    pause_event.set()

    status_label.config(
        text="STOPPED",
        fg="#dc2626"
    )

    log("STOP BUTTON PRESSED")


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
# UPLOAD TOC
# =========================================================

def upload_toc():

    global TOC_FILE
    global toc_data

    # In Master Suite walk-away mode, use the exact TOC selected
    # in the Master UI.  Direct execution still uses the normal picker.
    if MASTER_MODE:
        file_path = os.environ.get("EDRC_LIST", "").strip()
        if file_path:
            log(f"MASTER MODE TOC PATH : {file_path}")
        else:
            log("MASTER MODE: EDRC_LIST is empty")
            return
    else:
        file_path = filedialog.askopenfilename(
            title="Select TOC Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

    if not file_path:
        return

    TOC_FILE = file_path

    tree.delete(*tree.get_children())
    toc_data.clear()

    try:

        wb = load_workbook(
            TOC_FILE,
            data_only=True
        )

        if "TOC" not in wb.sheetnames:

            messagebox.showerror(
                "ERROR",
                "TOC Sheet Not Found"
            )

            wb.close()
            return

        sheet = wb["TOC"]

        log("")
        log("========================================")
        log("LOADING TOC")
        log("========================================")

        # ---------------------------------------
        # CREATE HEADER MAP
        # ---------------------------------------

        header_map = {}

        for col in range(1, sheet.max_column + 1):

            value = sheet.cell(
                row=1,
                column=col
            ).value

            if value is not None:

                header_map[
                    str(value).strip().upper()
                ] = col

        # ---------------------------------------
        # REQUIRED HEADERS
        # ---------------------------------------

        required = [

            "SIGNAL_SEQUENCE",

            "ROUTE_SEQUENCE"

        ]

        for head in required:

            if head not in header_map:

                messagebox.showerror(

                    "ERROR",

                    f"{head} column not found."

                )

                wb.close()

                return

        # ---------------------------------------
        # S.NO COLUMN
        # ---------------------------------------

        slno_col = None

        for name in [

            "S.NO",

            "SL.NO",

            "SL NO",

            "S NO"

        ]:

            if name in header_map:

                slno_col = header_map[name]

                break

        # ---------------------------------------
        # READ DATA
        # ---------------------------------------

        for row in range(2, sheet.max_row + 1):

            signal_sequence = sheet.cell(

                row=row,

                column=header_map["SIGNAL_SEQUENCE"]

            ).value

            if signal_sequence is None:

                continue

            route_sequence = sheet.cell(

                row=row,

                column=header_map["ROUTE_SEQUENCE"]

            ).value

            if route_sequence is None:

                route_sequence = ""

            slno = row - 1

            if slno_col:

                value = sheet.cell(

                    row=row,

                    column=slno_col

                ).value

                if value is not None:

                    slno = value

            signal_sequence = str(signal_sequence).strip()

            route_sequence = str(route_sequence).strip()

            toc_data.append({

                "slno": slno,

                "signal_sequence": signal_sequence,

                "route_sequence": route_sequence

            })

            tree.insert(

                "",

                "end",

                values=(

                    slno,

                    signal_sequence,

                    route_sequence

                )

            )

            log(

                f"{slno} | "

                f"{signal_sequence} | "

                f"{route_sequence}"

            )

        wb.close()

        log("")
        log(f"TOTAL ROWS LOADED : {len(toc_data)}")
        log("TOC LOADED SUCCESSFULLY")

    except Exception as e:

        log(f"UPLOAD TOC ERROR : {e}")

        messagebox.showerror(
            "ERROR",
            str(e)
        )
def get_toc_details(slno):

    for item in toc_data:

        if str(item["slno"]) == str(slno):

            return (

                item["signal_sequence"],

                item["route_sequence"]

            )

    return None, None

# =========================================================
# ROOT WINDOW
# =========================================================

root = tk.Tk()
root.title(APP_NAME)
root.geometry("1500x900")
root.configure(bg="#e9edf2")

# =========================================================
# HEADER
# =========================================================

header = tk.Frame(
    root,
    bg="#0b1220",
    height=120,
    bd=0
)
header.pack(fill="x")
header.pack_propagate(False)

# ---------------- LEFT LOGO ----------------

left_logo_frame = tk.Frame(
    header,
    bg="#0b1220",
    width=220
)
left_logo_frame.pack(side="left", fill="y", padx=20)
left_logo_frame.pack_propagate(False)

try:
    railway_logo = tk.PhotoImage(file="indian_railways.png")

    # Change 4 to 5 if still large
    railway_logo = railway_logo.subsample(6, 6)

    railway_label = tk.Label(
        left_logo_frame,
        image=railway_logo,
        bg="#0f172a",
        bd=0
    )
    railway_label.image = railway_logo
    railway_label.pack(expand=True)

except Exception:

    tk.Label(
        left_logo_frame,
        text="INDIAN RAILWAYS",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(expand=True)

# ---------------- TITLE ----------------

title = tk.Label(
    header,
    text="THROUGH TESTING OF RELATED SIGNALS",
    bg="#0b1220",
    fg="white",
    font=("Segoe UI",24,"bold")
)
title.place(
    relx=0.5,
    rely=0.5,
    anchor="center"
)

# ---------------- RIGHT LOGO ----------------

right_logo_frame = tk.Frame(
    header,
    bg="#0b1220",
    width=220
)
right_logo_frame.pack(side="right", fill="y", padx=20)
right_logo_frame.pack_propagate(False)

try:
    company_logo = tk.PhotoImage(file="company_logo.png")
    company_logo = company_logo.subsample(2, 2)

    company_label = tk.Label(
        right_logo_frame,
        image=company_logo,
        bg="#0f172a",
        bd=0
    )
    company_label.image = company_logo
    company_label.pack(expand=True)

except Exception:

    tk.Label(
        right_logo_frame,
        text="COMPANY LOGO",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 12, "bold")
    ).pack(expand=True)

# =========================================================
# MAIN FRAME
# =========================================================

main_frame = tk.Frame(root, bg="#e9edf2")
main_frame.pack(
    fill="both",
    expand=True,
    padx=20,
    pady=15
)

# =========================================================
# BUTTON HOVER EFFECT
# =========================================================

def add_hover_effect(button, normal_color, hover_color):

    button.bind(
        "<Enter>",
        lambda e: button.config(bg=hover_color)
    )

    button.bind(
        "<Leave>",
        lambda e: button.config(bg=normal_color)
    )
# =========================================================
# BUTTON CREATOR
# =========================================================

def create_side_button(text, command, color, hover):

    btn = tk.Button(
        left_panel,
        text=text,
        command=command,
        bg=color,
        fg="white",
        activebackground=hover,
        activeforeground="white",
        relief="flat",
        bd=0,
        highlightthickness=0,
        cursor="hand2",
        font=("Segoe UI", 11, "bold"),
        height=3
    )

    btn.pack(
        fill="x",
        padx=15,
        pady=(0,10),
        ipady=4
    )

    add_hover_effect(btn, color, hover)

    return btn
# =========================================================
# LEFT PANEL STYLING
# =========================================================

SIDE_MARGIN = 20
BUTTON_GAP = 10
TOP_MARGIN = 18
# =========================================================
# LEFT PANEL - CONTROL
# =========================================================

left_panel = tk.Frame(
    main_frame,
    width=320,
    bg="white",
    relief="solid",
    bd=1
)

left_panel.pack(
    side="left",
    fill="y",
    padx=(0,15)
)

left_panel.pack_propagate(False)

# =========================================================
# TITLE
# =========================================================

tk.Label(
    left_panel,
    text="CONTROL PANEL",
    bg="white",
    fg="#0f172a",
    font=("Segoe UI",18,"bold")
).pack(
    pady=(12, 22)
)
# =========================================================
# COMMON BUTTON
# =========================================================

def create_panel_button(text, command, color):

    btn = tk.Button(

        left_panel,

        text=text,

        command=command,

        bg=color,

        fg="white",

        activebackground=color,

        activeforeground="white",

        relief="flat",

        bd=0,

        cursor="hand2",

        font=("Segoe UI",11,"bold")

    )

    btn.pack(
        fill="x",
        padx=SIDE_MARGIN,
        pady=(0,BUTTON_GAP),
        ipady=12
    )

    return btn
# =========================================================
# CAPTURE BUTTON
# =========================================================

tk.Label(
    left_panel,
    text="CAPTURE SIGNALLING GEARS",
    bg="#2563eb",
    fg="white",
    font=("Segoe UI", 11, "bold"),
    padx=10,
    pady=10
).pack(
    fill="x",
    padx=SIDE_MARGIN,
    pady=(0, BUTTON_GAP)
)

# =========================================================
# NEW / EXISTING
# =========================================================
signal_frame = tk.Frame(
    left_panel,
    bg="white"
)

signal_frame.pack(
    fill="x",
    padx=SIDE_MARGIN,
    pady=(0,BUTTON_GAP)
)

new_btn = tk.Button(
    signal_frame,
    text="NEW SIGNAL",
    command=create_setup,
    bg="#f97316",
    fg="white",
    relief="flat",
    bd=0,
    font=("Segoe UI",10,"bold"),
    cursor="hand2"
)

new_btn.pack(
    side="left",
    expand=True,
    fill="x",
    padx=(0,5),
    ipady=8
)

edit_btn = tk.Button(
    signal_frame,
    text="EXISTING SIGNAL",
    command=open_existing_config,
    bg="#f97316",
    fg="white",
    relief="flat",
    bd=0,
    font=("Segoe UI",10,"bold"),
    cursor="hand2"
)

edit_btn.pack(
    side="left",
    expand=True,
    fill="x",
    padx=(5,0),
    ipady=8
)
# =========================================================
# IMPORT MASTER TOC
# =========================================================

create_panel_button(
    "IMPORT MASTER TOC",
    upload_toc,
    "#2563eb"
)

# =========================================================
# SAVE COORDINATES
# =========================================================
#
# There is NO separate "save coordinates" function in your
# current backend. create_setup() already captures and
# saves the coordinates by calling save_config().
#
create_panel_button(
    "SAVE COORDINATES",
    save_config,
    "#16a34a"
)

# =========================================================
# START TESTING
# =========================================================

def start_testing_from_ui():
    if automation_running:
        log("AUTOMATION ALREADY RUNNING - START IGNORED")
        return

    threading.Thread(
        target=test_route_engine,
        daemon=True
    ).start()


create_panel_button(
    "START TESTING",
    start_testing_from_ui,
    "#16a34a"
)

# =========================================================
# STOP TESTING
# =========================================================

create_panel_button(
    "STOP TESTING",
    stop_automation,
    "#dc2626"
)

# =========================================================
# SPACER
# =========================================================
tk.Frame(
    left_panel,
    bg="white",
    height=30
).pack(
    fill="x"
)
# =========================================================
# STATUS
# =========================================================

status_frame = tk.Frame(
    left_panel,
    bg="#f8fafc",
    highlightbackground="#d1d5db",
    highlightthickness=1
)

status_frame.pack(
    fill="x",
    padx=18,
    pady=(0,20)
)

tk.Label(
    status_frame,
    text="SYSTEM STATUS",
    bg="#f8fafc",
    fg="#111827",
    font=("Segoe UI", 11, "bold"),
    anchor="center",
    justify="center"
).pack(
    fill="x",
    pady=(15,6)
)
status_label = tk.Label(
    status_frame,
    text="READY",
    bg="#f8fafc",
    fg="#2563eb",
    font=("Segoe UI",14,"bold")
)

status_label.pack(
    pady=(0,15)
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
# ROUTE DETAILS TITLE
# =========================================================

table_title = tk.Label(
    right_panel,
    text="ROUTE DETAILS",
    font=("Segoe UI", 16, "bold"),
    bg="#e9edf2",
    fg="#0f172a"
)

table_title.pack(
    anchor="w",
    pady=(0, 8)
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
    expand=True,
    pady=(5, 15)
)

# =========================================================
# TREE SCROLLBAR
# =========================================================

tree_scroll = ttk.Scrollbar(
    table_frame,
    orient="vertical"
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
        "ROUTE"
    ),
    show="headings",
    yscrollcommand=tree_scroll.set
)

tree.heading(
    "NO",
    text="NO"
)

tree.heading(
    "SIGNAL",
    text="SIGNAL SEQUENCE"
)

tree.heading(
    "ROUTE",
    text="ROUTE SEQUENCE"
)

tree.column(
    "NO",
    width=120,
    anchor="center"
)

tree.column(
    "SIGNAL",
    width=420,
    anchor="center"
)

tree.column(
    "ROUTE",
    width=420,
    anchor="center"
)

tree.pack(
    fill="both",
    expand=True,
    padx=8,
    pady=8
)

tree_scroll.config(
    command=tree.yview
)

# =========================================================
# TREE STYLE
# =========================================================

style = ttk.Style()

style.configure(
    "Treeview",
    rowheight=30,
    font=("Segoe UI", 10)
)

style.configure(
    "Treeview.Heading",
    font=("Segoe UI", 11, "bold")
)

# =========================================================
# LIVE OPERATION LOG TITLE
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
    pady=(15, 10)
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
    expand=True,
    pady=(5, 0)
)

# =========================================================
# LOG SCROLLBAR
# =========================================================

log_scroll = tk.Scrollbar(log_frame)

log_scroll.pack(
    side="right",
    fill="y"
)

# =========================================================
# LOG TEXT
# =========================================================

log_text = tk.Text(
    log_frame,
    bg="white",
    fg="black",
    font=("Consolas", 10),
    relief="flat",
    wrap="word",
    yscrollcommand=log_scroll.set
)

log_text.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

log_scroll.config(
    command=log_text.yview
)

log_text.config(
    state="disabled"
)
# =========================================================
# FOOTER
# =========================================================

footer = tk.Label(

    root,

    text="THROUGH TESTING",

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

root.mainloop()

