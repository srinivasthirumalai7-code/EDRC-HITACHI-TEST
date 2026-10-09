"""
HITACHI C-FAT AUTOMATION
UI-ONLY VERSION + NEW SIGNAL COORDINATE CAPTURE

This file contains the user-interface layer plus the NEW SIGNAL workflow
required to capture and save signal coordinates.

All signal/point/crank operation/testing automation, TOC processing,
report generation and unrelated backend logic remain removed.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import os
import time
import threading
import pyautogui
import keyboard
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pywinauto import Desktop, Application
import ctypes

# =========================================================
# APPLICATION INFORMATION
# =========================================================

APP_NAME = "Hitachi C-FAT Automation"
VERSION = "1.0.0"
DEVELOPER = "SUVETHA"

BG = "#0f172a"
BTN = "#2563eb"
GREEN = "#22c55e"
RED = "#ef4444"

ROUTE_REPORT_FILE = None
report_rows = []
# =========================================================
# UI STATE
# =========================================================

TOC_FILE = None
CONFIG_FILE = None
paused = False
toc_data = []

# EDRC integration state. These are initialized at module scope so the
# EDRC launcher can load data and start the existing automation engine
# before any manual setup workflow has been used.
MASTER_MODE = os.environ.get("EDRC_UNATTENDED", "") == "1"
running = False
automation_running = False

signals = {}
starter_signals = {}
shunt_signals = {}
calling_on_signals = {}
lc_gates = {}
points = {}
crank_handles = {}

# =========================================================
# UI LOG
# =========================================================

LOG_FILE = os.path.join(
    os.getcwd(),
    f"NEGATIVE_TESTING_SIGNAL_LOG_{time.strftime('%Y%m%d_%H%M%S')}.txt"
)


def log(msg):
    """Write a message to the UI log AND to a log file, so the complete
    log is available even in walk-away mode."""

    line = f"[{time.strftime('%H:%M:%S')}] {msg}"

    try:
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass

    try:
        print(line, flush=True)
    except Exception:
        pass

    def write():
        try:
            log_text.config(state="normal")
            log_text.insert(
                tk.END,
                f"[{time.strftime('%H:%M:%S')}] {msg}\n"
            )
            log_text.see(tk.END)
            log_text.config(state="disabled")
        except tk.TclError:
            pass

    try:
        root.after(0, write)
    except tk.TclError:
        pass


# =========================================================
# UI STATUS
# =========================================================

def set_status(text, color=BTN):
    """Update the system status shown in the UI."""
    try:
        status_label.config(text=text, fg=color)
    except tk.TclError:
        pass


# =========================================================
# UI-ONLY BUTTON ACTIONS
# =========================================================

# =========================================================
# NEW SIGNAL - COORDINATE CAPTURE
# =========================================================

new_signal_data = {}
capture_queue = []
capture_index = 0
capture_active = False
capture_overlay = None
capture_progress_label = None
capture_signal_label = None
capture_step_label = None
capture_hint_label = None
last_space_time = 0


# =========================================================
# FULL COORDINATE CAPTURE ENGINE
# Based on the coordinate-capture mechanism in
# NEGATIVE_TESTING_bp(1).py
# =========================================================

def capture_point():
    """Wait for the global capture event and return [x, y]."""
    global capture_waiting, captured_point

    capture_waiting = True
    captured_point = None

    while capture_waiting:
        root.update_idletasks()
        time.sleep(0.05)

    if captured_point is None:
        return None

    return list(captured_point)


def capture_mouse_position():
    """Return the screen mouse position after the capture event."""
    point = capture_point()

    if point is None:
        return None, None

    x, y = point
    log(f"Captured Coordinate : ({x}, {y})")
    return x, y


def _capture_steps(aspects):
    steps = [
        "MENU",
        "RED",
    ]

    if aspects >= 3:
        steps.append("YELLOW")

    if aspects == 4:
        steps.append("DOUBLE_YELLOW")

    steps.extend([
        "GREEN",
        "ROUTE_INIT",
    ])

    return steps


def show_new_signal_name_dialog():
    result = {"value": None}

    win = tk.Toplevel(root)
    win.title("NEW SIGNAL")
    win.geometry("500x330")
    win.configure(bg="#0f172a")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    win.attributes("-topmost", True)

    tk.Label(
        win,
        text="NEW SIGNAL",
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(25, 10))

    tk.Label(
        win,
        text="Enter Signal Name",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 18, "bold")
    ).pack()

    entry = tk.Entry(
        win,
        font=("Segoe UI", 18),
        justify="center",
        width=12,
        bg="#1e293b",
        fg="white",
        insertbackground="white",
        relief="flat"
    )
    entry.pack(pady=20)
    entry.focus()

    def confirm():
        value = entry.get().strip().upper()
        if not value:
            messagebox.showerror(
                "Required",
                "Please enter a signal name.",
                parent=win
            )
            return

        result["value"] = value
        win.destroy()

    tk.Button(
        win,
        text="NEXT",
        command=confirm,
        bg="#2563eb",
        fg="white",
        relief="flat",
        width=18,
        font=("Segoe UI", 10, "bold")
    ).pack(pady=10)

    win.wait_window()
    return result["value"]


def show_new_signal_aspect_dialog(signal_name):
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
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(20, 10))

    tk.Label(
        win,
        text=f"Which aspect does\n{signal_name}\nhave?",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 18, "bold"),
        justify="center"
    ).pack(pady=(0, 20))

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
            font=("Segoe UI", 11, "bold")
        ).pack(pady=6)

    win.wait_window()
    return result["value"]


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
        except Exception:
            pass

    capture_overlay = tk.Toplevel(root)
    capture_overlay.title("Coordinate Capture")
    capture_overlay.geometry("420x390")
    capture_overlay.configure(bg="#0f172a")
    capture_overlay.attributes("-topmost", True)
    capture_overlay.resizable(False, False)
    capture_overlay.protocol("WM_DELETE_WINDOW", lambda: None)

    tk.Label(
        capture_overlay,
        text="COORDINATE CAPTURE",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 15, "bold")
    ).pack(pady=(15, 8))

    overlay_progress_label = tk.Label(
        capture_overlay,
        text="0 / 0",
        bg="#0f172a",
        fg="#22c55e",
        font=("Segoe UI", 13, "bold")
    )
    overlay_progress_label.pack()

    overlay_signal_label = tk.Label(
        capture_overlay,
        text="",
        bg="#0f172a",
        fg="#3b82f6",
        font=("Segoe UI", 18, "bold")
    )
    overlay_signal_label.pack(pady=(8, 2))

    overlay_step_label = tk.Label(
        capture_overlay,
        text="",
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 13)
    )
    overlay_step_label.pack(pady=(5, 5))

    overlay_hint_label = tk.Label(
        capture_overlay,
        text="Move mouse\nPress SPACE",
        bg="#0f172a",
        fg="#facc15",
        font=("Segoe UI", 11)
    )
    overlay_hint_label.pack(pady=(5, 5))

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


def update_capture_overlay(item_type, item_name, stage):
    overlay_progress_label.config(
        text=f"STEP {capture_index + 1} OF {len(capture_master_list)}"
    )

    overlay_signal_label.config(
        text=f"{item_name}",
        fg="#3b82f6"
    )

    overlay_step_label.config(
        text=f"CAPTURE : {stage.replace('_', ' ')}"
    )

    overlay_hint_label.config(
        text=(
            f"Move the mouse onto this\n"
            f"{stage.replace('_', ' ')} location,\n"
            f"then press SPACE.\n\n"
            f"(or press BACKSPACE)"
        )
    )


# Current coordinate-capture mode (prevents keyboard listener NameError)
capture_module = None

def clear_current_capture(item_type, item_name, stage):
    if item_type != "MAIN":
        return

    if item_name not in new_signal_data:
        return

    if stage in new_signal_data[item_name]:
        new_signal_data[item_name][stage] = None


def master_next_capture():
    global capture_module
    global record_stage

    if capture_index >= len(capture_master_list):
        capture_module = None

        if capture_overlay:
            try:
                capture_overlay.destroy()
            except Exception:
                pass

        root.deiconify()
        root.lift()
        root.focus_force()

        log("========================================")
        log("Coordinate Capture Completed")
        log("Press 'SAVE COORDINATES' to save the captured coordinates.")
        log("========================================")

        messagebox.showinfo(
            "Capture Completed",
            "All coordinates have been captured successfully.\n\n"
            "Click the 'SAVE COORDINATES' button to save them.",
            parent=root
        )
        return

    item_type, item_name, stage = capture_master_list[capture_index]
    record_stage = stage
    capture_module = "MASTER"

    update_capture_overlay(item_type, item_name, stage)

    log("")
    log("========================================")
    log(f"{item_type} : {item_name}")
    log(f"STEP : {stage.replace('_', ' ')}")
    log("Move mouse to target and press SPACE")
    log("Press BACKSPACE to Undo")
    log("========================================")


def master_save_point(x, y):
    global capture_index

    if capture_index >= len(capture_master_list):
        return

    item_type, item_name, stage = capture_master_list[capture_index]

    if item_type != "MAIN":
        return

    if item_name not in new_signal_data:
        return

    new_signal_data[item_name][stage] = [int(x), int(y)]

    log("----------------------------------------")
    log(f"{item_name}")
    log(f"{stage}")
    log(f"Captured : ({x}, {y})")
    log("----------------------------------------")

    capture_index += 1

    if overlay_progress_label:
        overlay_progress_label.config(
            text=f"STEP {capture_index} OF {len(capture_master_list)}"
        )

    master_next_capture()


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


def cancel_master_capture():
    global capture_module
    global capture_waiting
    global captured_point

    capture_module = None
    capture_waiting = False
    captured_point = None

    if capture_overlay:
        try:
            capture_overlay.destroy()
        except Exception:
            pass

    root.deiconify()
    root.lift()
    root.focus_force()

    log("Coordinate Capture Cancelled")


def capture_space():
    global capture_module
    global last_space_time
    global capture_waiting
    global captured_point

    if capture_module != "MASTER":
        return

    now = time.time()
    if now - last_space_time < 0.25:
        return

    last_space_time = now

    if capture_index >= len(capture_master_list):
        return

    x, y = pyautogui.position()

    captured_point = [int(x), int(y)]
    capture_waiting = False


def capture_backspace():
    global last_backspace_time

    if capture_module != "MASTER":
        return

    now = time.time()
    if now - last_backspace_time < 0.25:
        return

    last_backspace_time = now
    undo_last_capture()


def start_master_capture():
    global capture_module

    capture_module = "MASTER"

    create_capture_overlay()
    capture_overlay.update_idletasks()

    # Keep the capture overlay visible while the main UI is minimized.
    try:
        capture_overlay.deiconify()
        capture_overlay.lift()
        capture_overlay.attributes("-topmost", True)
    except Exception:
        pass

    root.iconify()
    master_next_capture()


def start_new_signal_capture():
    global capture_master_list
    global capture_index
    global new_signal_data
    global main_signal_aspects
    global capture_module

    signal = show_new_signal_name_dialog()

    if not signal:
        return

    if signal in new_signal_data:
        messagebox.showerror(
            "Duplicate",
            f"{signal} already exists.",
            parent=root
        )
        return

    aspects = show_new_signal_aspect_dialog(signal)

    if aspects is None:
        return

    main_signal_aspects[signal] = aspects

    new_signal_data[signal] = {
        "type": "MAIN",
        "aspects": aspects,
        "MENU": None,
        "RED": None,
        "YELLOW": None,
        "DOUBLE_YELLOW": None,
        "GREEN": None,
        "ROUTE_INIT": None,
    }

    capture_master_list = [
        ("MAIN", signal, stage)
        for stage in _capture_steps(aspects)
    ]
    capture_index = 0
    capture_module = "MASTER"

    log("========================================")
    log(f"NEW SIGNAL : {signal}")
    log(f"ASPECTS    : {aspects}")
    log(f"TOTAL CAPTURES : {len(capture_master_list)}")
    log("Starting Coordinate Capture")
    log("========================================")

    start_master_capture()


# Global keyboard hooks are required because the main Tk window is minimized
# during capture.
keyboard.add_hotkey("space", capture_space)
keyboard.add_hotkey("backspace", capture_backspace)

# =========================================================

# UI DIALOGS
# =========================================================

def show_quantity_dialog(title, question):
    result = {"value": None}

    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("500x380")
    win.configure(bg="#0f172a")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    win.attributes("-topmost", True)

    # TITLE
    tk.Label(
        win,
        text=title.upper(),
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(28, 10))

    # QUESTION
    tk.Label(
        win,
        text=question,
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 18, "bold"),
        justify="center"
    ).pack(pady=(0, 22))

    # NUMBER ENTRY
    entry = tk.Entry(
        win,
        font=("Segoe UI", 18),
        justify="center",
        width=6,
        bg="#1e293b",
        fg="white",
        insertbackground="white",
        relief="flat"
    )
    entry.pack(pady=(0, 30))
    entry.focus_set()

    def confirm():
        value = entry.get().strip()

        if not value:
            messagebox.showerror(
                "Required",
                "Please enter a number.",
                parent=win
            )
            return

        try:
            value = int(value)
        except ValueError:
            messagebox.showerror(
                "Invalid",
                "Please enter a valid number.",
                parent=win
            )
            return

        if value < 0:
            messagebox.showerror(
                "Invalid",
                "Number cannot be negative.",
                parent=win
            )
            return

        result["value"] = value
        win.destroy()

    def skip():
        result["value"] = 0
        win.destroy()

    # BUTTON FRAME
    button_frame = tk.Frame(
        win,
        bg="#0f172a"
    )
    button_frame.pack()

    # CONFIRM
    tk.Button(
        button_frame,
        text="CONFIRM",
        command=confirm,
        bg="#2563eb",
        fg="white",
        activebackground="#1d4ed8",
        activeforeground="white",
        relief="flat",
        bd=0,
        width=14,
        height=2,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2"
    ).pack(side="left", padx=10)

    # SKIP
    tk.Button(
        button_frame,
        text="SKIP (0)",
        command=skip,
        bg="#0f172a",
        fg="#94a3b8",
        activebackground="#0f172a",
        activeforeground="white",
        relief="flat",
        bd=0,
        font=("Segoe UI", 10),
        cursor="hand2"
    ).pack(side="left", padx=5)

    win.bind("<Return>", lambda event: confirm())
    win.bind("<Escape>", lambda event: skip())

    win.wait_window()

    return result["value"]

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

    # TITLE
    tk.Label(
        win,
        text=title.upper(),
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(28, 10))

    # QUESTION
    tk.Label(
        win,
        text=question,
        bg="#0f172a",
        fg="white",
        font=("Segoe UI", 18, "bold"),
        justify="center"
    ).pack(pady=(0, 20))

    # SIGNAL NAME ENTRY
    entry = tk.Entry(
        win,
        font=("Segoe UI", 18),
        justify="center",
        width=14,
        bg="#1e293b",
        fg="white",
        insertbackground="white",
        relief="flat"
    )
    entry.pack(pady=(0, 25))
    entry.focus_set()

    def confirm():
        value = entry.get().strip().upper()

        if not value:
            messagebox.showerror(
                "Required",
                "Please enter a signal name.",
                parent=win
            )
            return

        result["value"] = value
        win.destroy()

    def cancel():
        result["value"] = None
        win.destroy()

    # BUTTON FRAME
    button_frame = tk.Frame(
        win,
        bg="#0f172a"
    )
    button_frame.pack()

    # CONFIRM
    tk.Button(
        button_frame,
        text="CONFIRM",
        command=confirm,
        bg="#2563eb",
        fg="white",
        activebackground="#1d4ed8",
        activeforeground="white",
        relief="flat",
        bd=0,
        width=14,
        height=2,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2"
    ).pack(side="left", padx=10)

    # CANCEL
    tk.Button(
        button_frame,
        text="CANCEL",
        command=cancel,
        bg="#0f172a",
        fg="#94a3b8",
        activebackground="#0f172a",
        activeforeground="white",
        relief="flat",
        bd=0,
        font=("Segoe UI", 10),
        cursor="hand2"
    ).pack(side="left", padx=5)

    win.bind("<Return>", lambda event: confirm())
    win.bind("<Escape>", lambda event: cancel())

    win.wait_window()

    return result["value"]


def show_aspect_dialog(signal_name):
    """UI-only aspect selection dialog."""
    result = {"value": None}

    win = tk.Toplevel(root)
    win.title("Signal Aspects")
    win.geometry("500x380")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()
    win.attributes("-topmost", True)

    tk.Label(
        win,
        text="SELECT ASPECTS",
        bg=BG,
        fg="#64748b",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(20, 10))

    tk.Label(
        win,
        text=f"Which aspect does\n{signal_name}\nhave?",
        bg=BG,
        fg="white",
        font=("Segoe UI", 18, "bold"),
        justify="center"
    ).pack(pady=(0, 20))

    def choose(value):
        result["value"] = value
        win.destroy()

    for value in (2, 3, 4):
        tk.Button(
            win,
            text=f"{value} ASPECTS",
            command=lambda v=value: choose(v),
            bg=BTN,
            fg="white",
            relief="flat",
            width=22,
            height=2,
            font=("Segoe UI", 11, "bold")
        ).pack(pady=6)

    win.wait_window()
    return result["value"]


def show_capture_popup(message):
    """UI-only information popup."""
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


def show_capture_summary():
    """UI-only capture summary popup."""
    messagebox.showinfo(
        "CAPTURE SUMMARY",
        "UI-only version.\n\n"
        "No capture/automation data is processed."
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
        pady=(0, 10),
        ipady=4
    )

    add_hover_effect(btn, color, hover)
    return btn


def create_panel_button(text, command, color):

    def button_clicked():
        log("")
        log("========================================")
        log(f"BUTTON CLICKED : {text}")
        log("========================================")

        try:
            command()

        except Exception as e:

            log(
                f"BUTTON ERROR : {type(e).__name__} : {e}"
            )

            import traceback

            log(
                traceback.format_exc()
            )

            messagebox.showerror(
                "AUTOMATION ERROR",
                str(e)
            )

    btn = tk.Button(
        left_panel,
        text=text,
        command=button_clicked,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        relief="flat",
        bd=0,
        cursor="hand2",
        font=("Segoe UI", 11, "bold")
    )

    btn.pack(
        fill="x",
        padx=SIDE_MARGIN,
        pady=(0, BUTTON_GAP),
        ipady=12
    )

    return btn
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
def load_config():

    global CONFIG_FILE
    global signals
    global starter_signals
    global shunt_signals
    global calling_on_signals
    global lc_gates
    global points
    global crank_handles

    signals = {}
    starter_signals = {}
    shunt_signals = {}
    calling_on_signals = {}
    lc_gates = {}
    points = {}
    crank_handles = {}

    if MASTER_MODE:
        edrc_coords = os.environ.get("EDRC_COORDS", "").strip()
        if edrc_coords:
            CONFIG_FILE = edrc_coords
            log(f"EDRC COORDINATE PATH : {CONFIG_FILE}")

    if not CONFIG_FILE:
        log("NO CONFIGURATION FILE SELECTED")
        return False

    if not os.path.exists(CONFIG_FILE):
        log("CONFIG FILE NOT FOUND")
        return False

    try:

        wb = load_workbook(
            CONFIG_FILE,
            data_only=True
        )

        # =====================================================
        # MAIN SIGNALS
        # =====================================================

        if "MAIN" in wb.sheetnames:

            ws = wb["MAIN"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                signals[signal] = {

                    "type": "MAIN",

                    "aspects": row[1],

                    "menu": (
                        [int(row[2]), int(row[3])]
                        if row[2] is not None
                        and row[3] is not None
                        else None
                    ),

                    "RED": (
                        [int(row[4]), int(row[5])]
                        if row[4] is not None
                        and row[5] is not None
                        else None
                    ),

                    "YELLOW": (
                        [int(row[6]), int(row[7])]
                        if row[6] is not None
                        and row[7] is not None
                        else None
                    ),

                    "DOUBLE_YELLOW": (
                        [int(row[8]), int(row[9])]
                        if row[8] is not None
                        and row[9] is not None
                        else None
                    ),

                    "GREEN": (
                        [int(row[10]), int(row[11])]
                        if row[10] is not None
                        and row[11] is not None
                        else None
                    ),

                    "ROUTE_INIT": (
                        [int(row[12]), int(row[13])]
                        if row[12] is not None
                        and row[13] is not None
                        else None
                    )
                }

        # =====================================================
        # STARTER SIGNALS
        # =====================================================

        if "STARTER" in wb.sheetnames:

            ws = wb["STARTER"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                starter_signals[signal] = {

                    "type": "STARTER",

                    "aspects": row[1],

                    "menu": (
                        [int(row[2]), int(row[3])]
                        if row[2] is not None
                        and row[3] is not None
                        else None
                    ),

                    "RED": (
                        [int(row[4]), int(row[5])]
                        if row[4] is not None
                        and row[5] is not None
                        else None
                    ),

                    "YELLOW": (
                        [int(row[6]), int(row[7])]
                        if row[6] is not None
                        and row[7] is not None
                        else None
                    ),

                    "DOUBLE_YELLOW": (
                        [int(row[8]), int(row[9])]
                        if row[8] is not None
                        and row[9] is not None
                        else None
                    ),

                    "GREEN": (
                        [int(row[10]), int(row[11])]
                        if row[10] is not None
                        and row[11] is not None
                        else None
                    ),

                    "ROUTE_INIT": None
                }

        # =====================================================
        # SHUNT SIGNALS
        # =====================================================
        # Coordinate file format:
        # SHUNT = Signal, Menu_X, Menu_Y, Indicator_X, Indicator_Y,
        #         RouteInit_X, RouteInit_Y, Snapshot_Path
        # The Indicator coordinate is the shunt indication used by
        # the automation as the WHITE indication point.

        if "SHUNT" in wb.sheetnames:

            ws = wb["SHUNT"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                shunt_signals[signal] = {

                    "menu": (
                        [int(row[1]), int(row[2])]
                        if len(row) > 2
                        and row[1] is not None
                        and row[2] is not None
                        else None
                    ),

                    # The supplied SHUNT sheet has one indication
                    # coordinate named Indicator_X / Indicator_Y.
                    "WHITE": (
                        [int(row[3]), int(row[4])]
                        if len(row) > 4
                        and row[3] is not None
                        and row[4] is not None
                        else None
                    ),

                    "ROUTE_INIT": (
                        [int(row[5]), int(row[6])]
                        if len(row) > 6
                        and row[5] is not None
                        and row[6] is not None
                        else None
                    )
                }

        # =====================================================
        # CALLING-ON (CAL)
        # =====================================================
        # Coordinate file format:
        # CAL = Signal, Menu_X, Menu_Y, Yellow_X, Yellow_Y,
        #       RouteInit_X, RouteInit_Y
        # Calling-On indication is YELLOW in this coordinate file.

        if "CAL" in wb.sheetnames:

            ws = wb["CAL"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                calling_on_signals[signal] = {

                    "menu": (
                        [int(row[1]), int(row[2])]
                        if len(row) > 2
                        and row[1] is not None
                        and row[2] is not None
                        else None
                    ),

                    "YELLOW": (
                        [int(row[3]), int(row[4])]
                        if len(row) > 4
                        and row[3] is not None
                        and row[4] is not None
                        else None
                    ),

                    "ROUTE_INIT": (
                        [int(row[5]), int(row[6])]
                        if len(row) > 6
                        and row[5] is not None
                        and row[6] is not None
                        else None
                    )
                }

        # Backward compatibility with older coordinate files.
        if "CALLING_ON" in wb.sheetnames:

            ws = wb["CALLING_ON"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                signal = str(row[0]).strip().upper()

                calling_on_signals[signal] = {
                    "menu": (
                        [int(row[1]), int(row[2])]
                        if len(row) > 2 and row[1] is not None and row[2] is not None
                        else None
                    ),
                    "YELLOW": (
                        [int(row[5]), int(row[6])]
                        if len(row) > 6 and row[5] is not None and row[6] is not None
                        else None
                    ),
                    "ROUTE_INIT": (
                        [int(row[7]), int(row[8])]
                        if len(row) > 8 and row[7] is not None and row[8] is not None
                        else None
                    )
                }

        # =====================================================
        # LC GATES
        # =====================================================

        if "LC" in wb.sheetnames:

            ws = wb["LC"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                gate = str(row[0]).strip().upper()

                lc_gates[gate] = {
                    "menu": (
                        [int(row[1]), int(row[2])]
                        if len(row) > 2 and row[1] is not None and row[2] is not None
                        else None
                    ),
                    "IN": (
                        [int(row[3]), int(row[4])]
                        if len(row) > 4 and row[3] is not None and row[4] is not None
                        else None
                    ),
                    "OUT": (
                        [int(row[5]), int(row[6])]
                        if len(row) > 6 and row[5] is not None and row[6] is not None
                        else None
                    )
                }

        # =====================================================
        # POINTS
        # =====================================================

        if "POINT" in wb.sheetnames:

            ws = wb["POINT"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                point = str(row[0]).strip().upper()

                points[point] = {

                    "menu": (
                        [int(row[1]), int(row[2])]
                        if row[1] is not None
                        and row[2] is not None
                        else None
                    ),

                    "NORMAL": (
                        [int(row[3]), int(row[4])]
                        if row[3] is not None
                        and row[4] is not None
                        else None
                    ),

                    "REVERSE": (
                        [int(row[5]), int(row[6])]
                        if row[5] is not None
                        and row[6] is not None
                        else None
                    ),

                    "FREE": (
                        [int(row[7]), int(row[8])]
                        if row[7] is not None
                        and row[8] is not None
                        else None
                    )
                }

        # =====================================================
        # CRANK HANDLES
        # =====================================================

        if "CH" in wb.sheetnames:

            ws = wb["CH"]

            for row in ws.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row or not row[0]:
                    continue

                crank = str(row[0]).strip().upper()

                crank_handles[crank] = {

                    "menu": (
                        [int(row[1]), int(row[2])]
                        if row[1] is not None
                        and row[2] is not None
                        else None
                    )
                }

        wb.close()

        log("========================================")
        log("CONFIG LOADED SUCCESSFULLY")
        log(
            f"Main Signals Loaded : "
            f"{len(signals)}"
        )
        log(
            f"Starter Signals Loaded : "
            f"{len(starter_signals)}"
        )
        log(
            f"Points Loaded : "
            f"{len(points)}"
        )
        log(
            f"Crank Handles Loaded : "
            f"{len(crank_handles)}"
        )
        log("========================================")

        return True

    except Exception as e:

        log(
            f"LOAD CONFIG ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return False


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

def _normalise_track_bits(value):
    """
    Normalize TRACK values from the TOC Excel sheet.

    Returns False for empty/invalid values and True when
    a usable track value is present.
    """

    if value is None:
        return False

    # Convert Excel numeric values safely
    if isinstance(value, (int, float)):
        return value != 0

    value = str(value).strip()

    if not value:
        return False

    # Treat common empty values as invalid
    if value.upper() in (
        "NONE",
        "NULL",
        "N/A",
        "NA",
        "-",
        "0",
        "FALSE",
    ):
        return False

    return True
# =========================================================
# REFRESH RIGHT PANEL
# =========================================================

def refresh_right_panel():
    """
    Refresh MAIN, CALLING-ON and SHUNT signal tables.
    Data comes directly from TOC.
    """

    try:

        # =====================================================
        # CLEAR TABLES
        # =====================================================

        for tree in (main_tree, cal_tree, shunt_tree):
            for item in tree.get_children():
                tree.delete(item)


        # =====================================================
        # MAIN SIGNAL TABLE
        # =====================================================

        main_signals = {}

        for row in toc_data:

            signal = str(
                row.get("main_signal", "")
            ).strip().upper()

            if not signal:
                continue

            main_signals[signal] = \
                main_signals.get(signal, 0) + 1

        no = 1

        for signal, route_count in main_signals.items():

            main_tree.insert(
                "",
                "end",
                values=(
                    no,
                    signal,
                    route_count
                )
            )

            no += 1


        # =====================================================
        # CALLING-ON TABLE
        # =====================================================

        calling_on_signals_from_toc = {}

        for row in toc_data:

            signal = str(
                row.get("calling_on", "")
            ).strip().upper()

            if not signal:
                continue

            main_signal = str(
                row.get("main_signal", "")
            ).strip().upper()

            track = str(
                row.get("track", "")
            ).strip().upper()

            key = signal

            if key not in calling_on_signals_from_toc:
                calling_on_signals_from_toc[key] = {
                    "route": main_signal,
                    "track": track
                }

        no = 1

        for signal, data in calling_on_signals_from_toc.items():

            cal_tree.insert(
                "",
                "end",
                values=(
                    no,
                    signal,
                    data["route"],
                    data["track"]
                )
            )

            no += 1


        # =====================================================
        # SHUNT TABLE
        # =====================================================

        shunt_signals_from_toc = {}

        for row in toc_data:

            signal = str(
                row.get("shunt", "")
            ).strip().upper()

            if not signal:
                continue

            route = str(
                row.get("main_signal", "")
            ).strip().upper()

            shunt_signals_from_toc.setdefault(
                signal,
                set()
            ).add(route)

        no = 1

        for signal, routes in shunt_signals_from_toc.items():

            shunt_tree.insert(
                "",
                "end",
                values=(
                    no,
                    signal,
                    len(routes)
                )
            )

            no += 1


        # =====================================================
        # LOG
        # =====================================================

        log(
            f"RIGHT PANEL REFRESHED : "
            f"MAIN={len(main_tree.get_children())}, "
            f"CALLING-ON={len(cal_tree.get_children())}, "
            f"SHUNT={len(shunt_tree.get_children())}"
        )

    except Exception as e:

        log(
            f"RIGHT PANEL REFRESH ERROR : "
            f"{type(e).__name__} : {e}"
        )

def upload_toc():
    global TOC_FILE
    global toc_data

    # In EDRC unattended mode the launcher supplies the TOC path through
    # EDRC_LIST. Normal/manual operation still uses the existing file picker.
    if MASTER_MODE:
        file_path = os.environ.get("EDRC_LIST", "").strip()
        if not file_path or not os.path.exists(file_path):
            log("EDRC TOC PATH NOT FOUND")
            log(f"EDRC_LIST : {file_path}")
            return False
        log(f"EDRC TOC PATH : {file_path}")
    else:
        file_path = filedialog.askopenfilename(
            title="Select TOC Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_path:
            return False

    TOC_FILE = file_path
    toc_data.clear()

    try:
        refresh_right_panel()
    except Exception:
        pass

    try:
        wb = load_workbook(TOC_FILE, data_only=True)

        log(f"Selected TOC : {TOC_FILE}")
        log(f"Available Sheets : {wb.sheetnames}")

        if "TOC" not in wb.sheetnames:
            messagebox.showerror("ERROR", "TOC sheet not found.")
            wb.close()
            return

        sheet = wb["TOC"]

        # -----------------------------------------------------
        # ONLY THESE TOC COLUMNS ARE USED.
        # Main / Shunt : Signal, Route, Points, CRANK_HANDLE,
        #                SIGNAL_BITS, POINT_BITS, CRANK_HANDLE_BITS
        # Calling-On : same columns + Track
        # -----------------------------------------------------
        allowed_headers = {
            "SIGNAL", "ROUTE", "POINTS", "CRANK_HANDLE",
            "SIGNAL_BITS", "POINT_BITS", "CRANK_HANDLE_BITS",
            "TRACK"
        }

        header_map = {}
        for col in range(1, sheet.max_column + 1):
            header = sheet.cell(1, col).value
            if header is not None:
                key = str(header).strip().upper()
                if key in allowed_headers:
                    header_map[key] = col

        log(f"TOC USED HEADERS : {list(header_map.keys())}")

        required = [
            "SIGNAL", "ROUTE", "POINTS", "CRANK_HANDLE",
            "SIGNAL_BITS", "POINT_BITS", "CRANK_HANDLE_BITS"
        ]
        missing = [x for x in required if x not in header_map]
        if missing:
            messagebox.showerror(
                "ERROR",
                "Required TOC columns missing:\n\n" + ", ".join(missing)
            )
            wb.close()
            return

        def split_values(value):
            if value is None:
                return []
            return [
                x.strip().upper()
                for x in str(value).split(",")
                if x.strip()
            ]

        for r in range(2, sheet.max_row + 1):
            signal_value = sheet.cell(r, header_map["SIGNAL"]).value
            route_value = sheet.cell(r, header_map["ROUTE"]).value

            if signal_value is None or route_value is None:
                continue

            signal = str(signal_value).strip().upper()
            route = str(route_value).strip().upper()
            if not signal or not route:
                continue

            points_value = sheet.cell(r, header_map["POINTS"]).value
            crank_handle_value = sheet.cell(r, header_map["CRANK_HANDLE"]).value
            signal_bits_value = sheet.cell(r, header_map["SIGNAL_BITS"]).value
            point_bits_value = sheet.cell(r, header_map["POINT_BITS"]).value
            crank_bits_value = sheet.cell(r, header_map["CRANK_HANDLE_BITS"]).value

            point_names = split_values(points_value)
            crank_handles_used = split_values(crank_handle_value)
            signal_bits = split_values(signal_bits_value)
            point_bits = split_values(point_bits_value)
            crank_bits = split_values(crank_bits_value)

            # Signal type from the SIGNAL NAME first, so the TOC can be
            # loaded BEFORE the coordinate file (the launcher feeds the
            # TOC first - with the old order 1C came out as UNKNOWN and
            # its calling-on Track was never read, so the track was not
            # dropped before the route was set).
            if signal.endswith("SH") or signal.startswith("SH"):
                signal_type = "SHUNT"
            elif signal.endswith("C"):
                signal_type = "CALLING_ON"
            elif signal in calling_on_signals:
                signal_type = "CALLING_ON"
            elif signal in shunt_signals:
                signal_type = "SHUNT"
            else:
                signal_type = "MAIN"

            # Track is allowed ONLY for Calling-On rows.
            track = ""
            if signal_type == "CALLING_ON" and "TRACK" in header_map:
                value = sheet.cell(r, header_map["TRACK"]).value
                if value is not None:
                    track = str(value).strip().upper()

            toc_data.append({
                "slno": len(toc_data) + 1,
                # Existing automation uses this field as the route key.
                # Therefore the new TOC Route column is mapped here.
                "main_signal": route,
                "signal": signal,
                "route": route,
                "signal_type": signal_type,
                "point_names": point_names,
                "crank_handles": crank_handles_used,
                "signal_bits": signal_bits,
                "point_bits": point_bits,
                "crank_bits": crank_bits,
                "track": track,
                "main_signals": [signal] if signal_type == "MAIN" else [],
                "calling_on_signals": [signal] if signal_type == "CALLING_ON" else [],
                "shunt_signals": [signal] if signal_type == "SHUNT" else []
            })

            log(
                f"TOC ROW {r} : SIGNAL={signal}, ROUTE={route}, "
                f"TYPE={signal_type}, POINTS={point_names}, "
                f"CRANK_HANDLE={crank_handles_used}, "
                f"SIGNAL_BITS={signal_bits}, POINT_BITS={point_bits}, "
                f"CRANK_HANDLE_BITS={crank_bits}, TRACK={track or 'NONE'}"
            )

        wb.close()

        log(f"TOC Loaded Successfully ({len(toc_data)} Rows)")
        log(
            f"MAIN={sum(1 for x in toc_data if x.get('signal_type') == 'MAIN')}, "
            f"CALLING-ON={sum(1 for x in toc_data if x.get('signal_type') == 'CALLING_ON')}, "
            f"SHUNT={sum(1 for x in toc_data if x.get('signal_type') == 'SHUNT')}"
        )

        try:
            refresh_right_panel()
        except Exception as e:
            log(f"RIGHT PANEL REFRESH WARNING : {e}")

        return True

    except Exception as e:
        log(f"TOC LOAD ERROR : {type(e).__name__} : {e}")
        messagebox.showerror(
            "TOC LOAD ERROR",
            f"{type(e).__name__} : {e}"
        )
        return False

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
def find_ui_element(name, control_type=None):
    """Find an exact UI element by visible text and click it."""
    try:
        desktop = Desktop(backend="uia")
        search_name = str(name).strip().upper()

        for window in desktop.windows():
            try:
                for control in window.descendants():
                    try:
                        text = control.window_text().strip()

                        if not text:
                            continue

                        if text.upper() != search_name:
                            continue

                        if control_type:
                            current_type = str(
                                control.element_info.control_type
                            )

                            if control_type.lower() not in current_type.lower():
                                continue

                        rect = control.rectangle()

                        if not _on_any_screen(rect):
                            continue

                        x = (rect.left + rect.right) // 2
                        y = (rect.top + rect.bottom) // 2

                        try:
                            window.set_focus()
                            time.sleep(0.2)
                        except Exception:
                            pass

                        log(f"FOUND : {name}")
                        log(f"CLICK : ({x},{y})")

                        pyautogui.moveTo(x, y, duration=0.2)
                        pyautogui.click()

                        time.sleep(0.5)

                        return True

                    except Exception:
                        continue

            except Exception:
                continue

    except Exception as e:
        log(f"UI SEARCH ERROR : {e}")

    log(f"NOT FOUND : {name}")
    return False


def open_test_panel_station():
    """Open Test Panel and select Station 50051."""

    log("========================================")
    log("OPENING TEST PANEL")
    log("========================================")

    # -------------------------------------------------
    # FIND AND FOCUS TEST PANEL
    # -------------------------------------------------
    try:
        desktop = Desktop(backend="uia")

        test_panel = desktop.window(
            title_re=".*Test Panel.*"
        )

        if not test_panel.exists(timeout=5):
            log("TEST PANEL WINDOW NOT FOUND")
            return False

        test_panel.set_focus()
        time.sleep(1)

        log("TEST PANEL FOUND AND FOCUSED")

    except Exception as e:
        log(f"TEST PANEL FOCUS ERROR : {e}")
        return False

    # Open station selection
    pyautogui.hotkey("ctrl", "b")

    log("CTRL+B PRESSED")

    time.sleep(3)

    # Select station 50051
    if not find_ui_element("50051"):
        log("50051 NOT FOUND")
        return False

    time.sleep(1)

    # Double-click selected station
    pyautogui.doubleClick()

    time.sleep(4)

    log("STATION 50051 OPENED")

    return True


def _on_any_screen(rect):
    """True if the rectangle is really visible on ANY monitor."""
    try:
        metrics = ctypes.windll.user32.GetSystemMetrics
        vx = metrics(76)   # SM_XVIRTUALSCREEN
        vy = metrics(77)   # SM_YVIRTUALSCREEN
        vw = metrics(78)   # SM_CXVIRTUALSCREEN
        vh = metrics(79)   # SM_CYVIRTUALSCREEN
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
        return (rect.right > rect.left and rect.bottom > rect.top
                and vx <= cx < vx + vw and vy <= cy < vy + vh)
    except Exception:
        return False


def find_station_relay(relay_name):
    """Find and click a relay inside the Station 50051 window.

    The Station 50051 window can be INSIDE the Test Panel (single
    screen) or a SEPARATE window on the other monitor (dual screen),
    so both places are searched. The window is brought to the front
    before clicking, so the click selects the relay instead of only
    activating the window.
    """

    relay_name = str(relay_name).strip().upper()

    log(f"Searching Station Relay : {relay_name}")

    try:
        desktop = Desktop(backend="uia")

        test_panel = desktop.window(
            title_re=".*Test Panel.*"
        )

        test_panel.wait(
            "exists",
            timeout=5
        )

        # Relay controls may take some time to populate
        for attempt in range(1, 11):

            # Test Panel + any separate "Station ..." window,
            # on either monitor
            containers = [test_panel]

            try:
                for win in desktop.windows():
                    try:
                        title = win.window_text().strip().upper()
                        if title.startswith("STATION"):
                            containers.append(win)
                    except Exception:
                        pass
            except Exception:
                pass

            for container in containers:

                try:
                    controls = container.descendants()
                except Exception:
                    controls = []

                for control in controls:

                    try:
                        text = control.window_text().strip()

                        if not text:
                            continue

                        if text.upper() != relay_name:
                            continue

                        try:
                            control.iface_scroll_item.ScrollIntoView()
                            time.sleep(0.2)
                        except Exception:
                            pass

                        rect = control.rectangle()

                        if not _on_any_screen(rect):
                            continue

                        x = (rect.left + rect.right) // 2
                        y = (rect.top + rect.bottom) // 2

                        try:
                            container.set_focus()
                            time.sleep(0.3)
                        except Exception:
                            pass

                        log(
                            f"FOUND RELAY : {relay_name} "
                            f"(attempt {attempt}) IN "
                            f"{container.window_text()}"
                        )

                        log(
                            f"RELAY CLICK : ({x},{y})"
                        )

                        pyautogui.moveTo(
                            x,
                            y,
                            duration=0.2
                        )

                        pyautogui.click()

                        time.sleep(0.5)

                        log(
                            f"{relay_name} SELECTED"
                        )

                        return True

                    except Exception:
                        continue

            if attempt < 10:
                time.sleep(0.75)

    except Exception as e:

        log(
            f"RELAY SEARCH ERROR "
            f"{relay_name} : {e}"
        )

    log(
        f"{relay_name} NOT FOUND"
    )

    return False
def set_all_points_normal():
    """Put every point from the POINT sheet back to NORMAL."""

    initial_points = list(points.keys())

    log("========================================")
    log("SETTING ALL POINTS TO NORMAL")
    log("========================================")

    if not initial_points:
        log("NO POINTS LOADED - NOTHING TO SET NORMAL")
        return False

    try:
        Desktop(backend="uia").window(
            title_re=".*Test Panel.*"
        ).set_focus()
        time.sleep(1)
    except Exception as e:
        log(f"TEST PANEL FOCUS ERROR : {e}")

    for point in initial_points:

        point = str(point).strip().upper()

        menu = points[point].get("menu")
        normal = points[point].get("NORMAL")

        if not menu:
            log(f"POINT MENU COORDINATE NOT FOUND : {point}")
            continue

        if not normal:
            log(f"POINT NORMAL COORDINATE NOT FOUND : {point}")
            continue

        log(f"SELECTING POINT MENU : {point}")

        pyautogui.click(
            menu[0],
            menu[1]
        )

        time.sleep(0.5)

        log(f"SELECTING NORMAL : {point}")

        pyautogui.click(
            normal[0],
            normal[1]
        )

        time.sleep(1)

    log("ALL POINTS SET TO NORMAL")

    return True

def read_initial_control_bits():
    """Read relay names from INITIAL_CONTROL_BITS sheet."""

    if not TOC_FILE:
        log("TOC FILE NOT LOADED")
        return []

    try:
        wb = load_workbook(
            TOC_FILE,
            data_only=True,
            read_only=True
        )

        if "INITIAL_CONTROL_BITS" not in wb.sheetnames:
            log("INITIAL_CONTROL_BITS SHEET NOT FOUND")
            wb.close()
            return []

        ws = wb["INITIAL_CONTROL_BITS"]

        # -------------------------------------------------
        # FIND RELAY COLUMN FROM HEADER
        # -------------------------------------------------
        relay_col = None

        for col in range(1, ws.max_column + 1):
            header = ws.cell(
                row=1,
                column=col
            ).value

            if (
                header is not None
                and str(header).strip().upper() == "RELAY"
            ):
                relay_col = col
                break

        if relay_col is None:
            log("RELAY COLUMN NOT FOUND")
            wb.close()
            return []

        log(
            f"INITIAL_CONTROL_BITS RELAY COLUMN : "
            f"{relay_col}"
        )

        # -------------------------------------------------
        # READ ALL RELAYS
        # -------------------------------------------------
        relays = []

        for row in range(2, ws.max_row + 1):

            relay = ws.cell(
                row=row,
                column=relay_col
            ).value

            if relay is None:
                continue

            relay = str(relay).strip()

            if not relay:
                continue

            relays.append(relay)

        wb.close()

        log(
            f"INITIAL CONTROL BITS LOADED : "
            f"{len(relays)}"
        )

        for relay in relays:
            log(f"  INITIAL BIT : {relay}")

        return relays

    except Exception as e:

        log(
            f"INITIAL CONTROL BITS READ ERROR : {e}"
        )

        return []
def restore_initial_control_bits():
    """
    Restore all relays listed in INITIAL_CONTROL_BITS
    at the END of the complete automation.
    """

    log("")
    log("========================================")
    log("RESTORING INITIAL CONTROL BITS")
    log("========================================")

    # -------------------------------------------------
    # READ INITIAL CONTROL BITS
    # -------------------------------------------------

    initial_bits = read_initial_control_bits()

    if not initial_bits:

        log(
            "INITIAL CONTROL BITS RESTORE FAILED : "
            "NO RELAYS FOUND"
        )

        return False

    log(
        f"INITIAL CONTROL BITS TO RESTORE : "
        f"{len(initial_bits)}"
    )

    # -------------------------------------------------
    # OPEN STATION 50051
    # -------------------------------------------------

    if not open_test_panel_station():

        log(
            "FAILED TO OPEN STATION 50051 "
            "FOR INITIAL CONTROL BIT RESTORE"
        )

        return False

    time.sleep(2)

    # -------------------------------------------------
    # SELECT ALL INITIAL RELAYS
    # -------------------------------------------------

    success = True

    for relay in initial_bits:

        relay = str(relay).strip()

        if not relay:
            continue

        log(
            f"RESTORING INITIAL RELAY : "
            f"{relay}"
        )

        if not find_station_relay(relay):

            log(
                f"FAILED TO RESTORE INITIAL RELAY : "
                f"{relay}"
            )

            success = False

            continue

        log(
            f"INITIAL RELAY SELECTED : "
            f"{relay}"
        )

        time.sleep(0.3)

    # -------------------------------------------------
    # TRANSMIT
    # -------------------------------------------------

    if not find_ui_element(
        "Transmit",
        control_type="Button"
    ):

        log(
            "INITIAL RESTORE TRANSMIT BUTTON "
            "NOT FOUND"
        )

        success = False

    else:

        time.sleep(1)

        log(
            "TRANSMITTING INITIAL CONTROL BITS RESTORE"
        )

        pyautogui.click()

        time.sleep(2)

        log(
            "INITIAL CONTROL BITS RESTORE TRANSMITTED"
        )

    # -------------------------------------------------
    # CLOSE WINDOW
    # -------------------------------------------------

    close_station_control_bits_window()

    time.sleep(1)

    # -------------------------------------------------
    # FINAL RESULT
    # -------------------------------------------------

    if success:

        log("")
        log("========================================")
        log("INITIAL CONTROL BITS RESTORED SUCCESSFULLY")
        log("========================================")

    else:

        log("")
        log("========================================")
        log("INITIAL CONTROL BITS RESTORE FAILED")
        log("========================================")

    return success

def close_station_control_bits_window():

    log(
        "SEARCHING STATION 50051 CONTROL BITS WINDOW"
    )

    # =====================================================
    # METHOD 1 - FIND WINDOW USING UI AUTOMATION
    # =====================================================

    try:

        desktop = Desktop(
            backend="uia"
        )

        windows = desktop.windows()

        for window in windows:

            try:

                title = (
                    window.window_text()
                    .strip()
                    .upper()
                )

                log(
                    f"CHECKING WINDOW FOR CLOSE : "
                    f"{title}"
                )

                # -------------------------------------------------
                # DO NOT REQUIRE ALL THREE WORDS.
                # STATION 50051 IS THE MAIN IDENTIFIER.
                # -------------------------------------------------

                if "50051" not in title:
                    continue

                log(
                    f"CONTROL BITS WINDOW FOUND : "
                    f"{title}"
                )

                # -------------------------------------------------
                # TRY NORMAL WINDOW CLOSE FIRST
                # -------------------------------------------------

                try:

                    window.set_focus()

                    time.sleep(0.3)

                    window.close()

                    time.sleep(1)

                    log(
                        "STATION 50051 CONTROL BITS "
                        "WINDOW CLOSED USING WINDOW.CLOSE()"
                    )

                    return True

                except Exception as e:

                    log(
                        f"WINDOW.CLOSE FAILED : "
                        f"{type(e).__name__} : {e}"
                    )

                # -------------------------------------------------
                # FALLBACK - CLICK TOP RIGHT X
                # -------------------------------------------------

                try:

                    rect = window.rectangle()

                    x = rect.right - 15
                    y = rect.top + 15

                    log(
                        f"CLICKING CONTROL BITS X : "
                        f"({x},{y})"
                    )

                    window.set_focus()

                    time.sleep(0.3)

                    pyautogui.moveTo(
                        x,
                        y,
                        duration=0.2
                    )

                    pyautogui.click()

                    time.sleep(1)

                    log(
                        "STATION 50051 CONTROL BITS "
                        "WINDOW CLOSED USING X"
                    )

                    return True

                except Exception as e:

                    log(
                        f"X CLICK CLOSE FAILED : "
                        f"{type(e).__name__} : {e}"
                    )

            except Exception:
                continue

    except Exception as e:

        log(
            f"CONTROL BITS WINDOW SEARCH ERROR : "
            f"{type(e).__name__} : {e}"
        )

    # =====================================================
    # METHOD 2 - ALT + F4 FALLBACK
    # =====================================================

    try:

        log(
            "USING ALT+F4 FALLBACK "
            "FOR CONTROL BITS WINDOW"
        )

        pyautogui.hotkey(
            "alt",
            "f4"
        )

        time.sleep(1)

        log(
            "ALT+F4 PRESSED - "
            "CONTROL BITS WINDOW CLOSED"
        )

        return True

    except Exception as e:

        log(
            f"ALT+F4 CLOSE ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return False

def open_signal_menu(signal):

    log(
        f"OPEN SIGNAL MENU : {signal}"
    )

    coordinates = get_signal_menu(signal)

    if not coordinates:
        log(
            f"SIGNAL MENU NOT FOUND : {signal}"
        )
        return False

    try:

        x, y = coordinates

        log(
            f"CLICKING SIGNAL MENU : "
            f"{signal} -> ({x},{y})"
        )

        pyautogui.moveTo(
            int(x),
            int(y),
            duration=0.4
        )

        time.sleep(0.3)

        pyautogui.click()

        time.sleep(1.5)

        return True

    except Exception as e:

        log(
            f"OPEN SIGNAL MENU ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return False

def establish_actual_route_sequence(main_signal):

    global route_selection_started

    route_selection_started = False

    main_signal = str(
        main_signal
    ).strip().upper()

    sequence = get_actual_route_sequence(
        main_signal
    )

    if not sequence:

        log(
            f"ROUTE ESTABLISHMENT FAILED : "
            f"{main_signal}"
        )

        return False

    log("")
    log("========================================")
    log("ESTABLISHING ACTUAL ROUTE SEQUENCE")
    log("========================================")

    log(
        f"MAIN SIGNAL : {main_signal}"
    )

    log(
        f"ACTUAL SEQUENCE : {sequence}"
    )

    total_steps = len(sequence)

    # =====================================================
    # PROCESS EACH SIGNAL + ROUTE
    # =====================================================

    for step_no, (
        signal,
        route
    ) in enumerate(
        sequence,
        start=1
    ):

        signal = str(
            signal
        ).strip().upper()

        route = str(
            route
        ).strip().upper()

        log("")
        log("----------------------------------------")
        log(
            f"ROUTE STEP {step_no}/{total_steps}"
        )
        log(
            f"SIGNAL : {signal}"
        )
        log(
            f"ROUTE  : {route}"
        )
        log("----------------------------------------")

        # =================================================
        # OPEN SIGNAL MENU
        # =================================================

        log(
            f"OPENING MENU FOR SIGNAL : {signal}"
        )

        if not open_signal_menu(signal):

            log(
                f"❌ SIGNAL MENU FAILED : "
                f"{signal}"
            )

            return False

        log(
            f"✅ SIGNAL MENU OPENED : "
            f"{signal}"
        )

        # =================================================
        # SELECT ACTUAL ROUTE
        # =================================================

        log(
            f"SEARCHING ACTUAL ROUTE : "
            f"{route}"
        )

        if not find_and_click(route):

            log(
                f"❌ ACTUAL ROUTE NOT FOUND : "
                f"{route}"
            )

            return False

        log(
            f"✅ ACTUAL ROUTE SELECTED : "
            f"{route}"
        )

        # -------------------------------------------------
        # FIRST SUCCESSFUL SELECTION MEANS CLEANUP
        # IS NOW MANDATORY
        # -------------------------------------------------

        route_selection_started = True

        # =================================================
        # WAIT BEFORE NEXT ROUTE
        # =================================================

        if step_no < total_steps:

            log(
                "WAITING BEFORE NEXT ROUTE..."
            )

            time.sleep(3)

    # =====================================================
    # ALL ROUTES COMPLETED
    # =====================================================

    log("")
    log("========================================")
    log("✅ FULL ACTUAL ROUTE ESTABLISHED")
    log("========================================")

    return True

def process_auxiliary_signal(signal, route, signal_type):
    """
    Process one CALLING-ON or SHUNT signal.

    Sequence:
        1. Open signal menu
        2. Select route
        3. Wait for indication
        4. Verify WHITE aspect
        5. Return PASS / FAIL

    signal_type:
        CALLING_ON
        SHUNT
    """

    signal = str(signal or "").strip().upper()
    route = str(route or "").strip().upper()
    signal_type = str(signal_type or "").strip().upper()

    if not signal:
        log(
            f"{signal_type} SIGNAL NOT PROVIDED"
        )
        return False

    if not route:
        log(
            f"{signal_type} ROUTE NOT PROVIDED : {signal}"
        )
        return False

    log("")
    log("========================================")
    log(f"PROCESSING {signal_type} SIGNAL")
    log("========================================")
    log(f"SIGNAL : {signal}")
    log(f"ROUTE  : {route}")

    # =====================================================
    # OPEN SIGNAL MENU
    # =====================================================

    log(
        f"OPENING MENU FOR {signal_type} : {signal}"
    )

    if not open_signal_menu(signal):

        log(
            f"{signal_type} MENU OPEN FAILED : {signal}"
        )

        return False

    log(
        f"{signal_type} MENU OPENED : {signal}"
    )

    # =====================================================
    # SELECT ROUTE
    # =====================================================

    log(
        f"SELECTING {signal_type} ROUTE : {route}"
    )

    if not find_and_click(route):

        log(
            f"{signal_type} ROUTE NOT FOUND : "
            f"{signal} -> {route}"
        )

        return False

    log(
        f"{signal_type} ROUTE SELECTED : "
        f"{signal} -> {route}"
    )

    # =====================================================
    # WAIT FOR INDICATION
    # =====================================================

    log(
        f"WAITING FOR {signal_type} INDICATION..."
    )

    time.sleep(3)

    # =====================================================
    # GET INDICATION COORDINATE
    # =====================================================

    if signal_type == "CALLING_ON":

        signal_data = calling_on_signals.get(signal)

    elif signal_type == "SHUNT":

        signal_data = shunt_signals.get(signal)

    else:

        log(
            f"INVALID AUXILIARY SIGNAL TYPE : "
            f"{signal_type}"
        )

        return False

    if not signal_data:

        log(
            f"{signal_type} CONFIG NOT FOUND : "
            f"{signal}"
        )

        return False

    # Calling-On uses YELLOW from the CAL sheet.
    # Shunt uses Indicator_X / Indicator_Y, loaded internally as WHITE.
    if signal_type == "CALLING_ON":
        indication_coord = signal_data.get("YELLOW")
        indication_name = "YELLOW"
    else:
        indication_coord = signal_data.get("WHITE")
        indication_name = "WHITE"

    if not indication_coord:

        log(
            f"{signal_type} {indication_name} COORDINATE NOT FOUND : "
            f"{signal}"
        )

        return False

    # =====================================================
    # READ WHITE LAMP
    # =====================================================

    try:

        screen_width, screen_height = pyautogui.size()

        BASE_WIDTH = 1920
        BASE_HEIGHT = 1080

        scale_x = screen_width / BASE_WIDTH
        scale_y = screen_height / BASE_HEIGHT

        screen_x = round(
            float(indication_coord[0]) * scale_x
        )

        screen_y = round(
            float(indication_coord[1]) * scale_y
        )

        screen_x = max(
            0,
            min(screen_x, screen_width - 1)
        )

        screen_y = max(
            0,
            min(screen_y, screen_height - 1)
        )

        rgb = pyautogui.pixel(
            screen_x,
            screen_y
        )

        rgb = (
            int(rgb[0]),
            int(rgb[1]),
            int(rgb[2])
        )

        log(
            f"{signal_type} {indication_name} RGB : "
            f"{signal} : {rgb}"
        )

    except Exception as e:

        log(
            f"{signal_type} WHITE LAMP READ ERROR : "
            f"{signal} : "
            f"{type(e).__name__} : {e}"
        )

        return False

    # =====================================================
    # VERIFY INDICATION
    # =====================================================

    if signal_type == "CALLING_ON":

        detected = is_yellow_lamp(rgb)

    else:

        r, g, b = rgb
        detected = (
            r >= 150
            and g >= 150
            and b >= 150
            and abs(r - g) <= 70
            and abs(g - b) <= 70
            and abs(r - b) <= 70
        )

    if detected:

        log(
            f"PASS : {signal_type} {indication_name} ASPECT "
            f"DETECTED : {signal}"
        )

        return True

    log(
        f"FAIL : {signal_type} {indication_name} ASPECT "
        f"NOT DETECTED : {signal} : RGB={rgb}"
    )

    return False

def analyze_actual_route_lamps(main_signal, shunt_baseline=None):

    main_signal = str(main_signal).strip().upper()

    sequence = get_actual_route_sequence(main_signal)

    if not sequence:
        log(f"LAMP ANALYSIS FAILED : NO ROUTE SEQUENCE FOR {main_signal}")
        return [], "FAIL"

    log("")
    log("========================================")
    log("ANALYZING ROUTE SIGNAL LAMPS")
    log("========================================")
    log(f"MAIN SIGNAL : {main_signal}")

    results = []

    for step_no, (signal, route) in enumerate(sequence, start=1):

        signal = str(signal).strip().upper()
        route = str(route).strip().upper()

        log("")
        log("----------------------------------------")
        log(f"LAMP CHECK {step_no}/{len(sequence)}")
        log(f"SIGNAL : {signal}")
        log(f"ROUTE  : {route}")
        log("----------------------------------------")

        # Calling-On can take up to 6 seconds to illuminate.
        # Keep checking only this signal type; normal MAIN/STARTER
        # lamp timing remains unchanged.
        if signal in calling_on_signals:
            colour = "UNKNOWN"
            deadline = time.time() + 6.0

            while time.time() < deadline:
                colour = read_signal_lamp_colour(signal)
                if colour == "YELLOW":
                    break
                time.sleep(0.25)

            if colour == "YELLOW":
                result = "PASS"
            else:
                result = "FAIL"
                log("CALLING-ON YELLOW NOT DETECTED WITHIN 6 SECONDS")

        elif signal in shunt_signals:
            # SHUNT verification uses the route-indicator coordinate
            # recorded in the SHUNT sheet (RouteInit_X / RouteInit_Y).
            # For this panel the established route indicator glows YELLOW.
            route_indicator_rgb = get_signal_lamp_rgb(signal, "ROUTE_INIT")

            yellow_detected = is_yellow_lamp(route_indicator_rgb)

            log(
                f"SHUNT ROUTE INDICATOR : {signal} : "
                f"RGB={route_indicator_rgb} : "
                f"YELLOW={yellow_detected}"
            )

            result = "PASS" if yellow_detected else "FAIL"
            colour = "YELLOW" if yellow_detected else "NOT YELLOW"

            if result == "PASS":
                log(f"PASS : SHUNT ROUTE INDICATOR YELLOW DETECTED : {signal}")
            else:
                log(f"FAIL : SHUNT ROUTE INDICATOR YELLOW NOT DETECTED : {signal}")

        else:
            colour = read_signal_lamp_colour(signal)

            if colour == "RED":
                result = "FAIL"
            elif colour in ("YELLOW", "GREEN"):
                result = "PASS"
            else:
                result = "FAIL"

        log(f"SIGNAL : {signal}")
        log(f"ROUTE  : {route}")
        log(f"LAMP COLOUR : {colour}")
        log(f"RESULT : {result}")

        results.append({
            "signal": signal,
            "route": route,
            "lamp": colour,
            "result": result
        })

        time.sleep(0.5)

    if results and all(
        item["result"] == "PASS"
        for item in results
    ):
        overall_result = "PASS"
    else:
        overall_result = "FAIL"

    log("")
    log("========================================")
    log(f"OVERALL LAMP RESULT : {overall_result}")
    log("========================================")

    return results, overall_result

def is_red_lamp(rgb):

    if rgb is None:
        return False

    r, g, b = rgb

    return (
        r > 150
        and r > g * 1.4
        and r > b * 1.4
    )

def is_yellow_lamp(rgb):

    if not rgb:
        return False

    r, g, b = rgb

    # =====================================================
    # YELLOW LAMP
    # =====================================================

    # Normal yellow
    if r >= 150 and g >= 150 and b <= 100:
        return True

    # Railway panel yellow variations
    if r >= 180 and g >= 180 and b <= 80:
        return True

    # Specific observed lamp:
    # RGB = (210, 210, 7)
    if (
        190 <= r <= 230
        and
        190 <= g <= 230
        and
        0 <= b <= 30
    ):
        return True

    return False

def is_green_lamp(rgb):

    if rgb is None:
        return False

    r, g, b = rgb

    return (
        g > 120
        and g > r * 1.2
        and g > b * 1.2
    )
def get_signal_lamp_rgb(signal, lamp_type):
    try:
        signal = str(signal).strip().upper()
        lamp_type = str(lamp_type).strip().upper()

        # =====================================================
        # FIND SIGNAL CONFIGURATION
        # =====================================================
        signal_data = signals.get(signal)

        if signal_data is None:
            signal_data = starter_signals.get(signal)

        # Calling-On coordinates are stored separately in the CAL sheet.
        if signal_data is None:
            signal_data = calling_on_signals.get(signal)

        # Shunt coordinates are stored separately in the SHUNT sheet.
        if signal_data is None:
            signal_data = shunt_signals.get(signal)

        if signal_data is None:
            log(f"LAMP CONFIG NOT FOUND : {signal}")
            return (0, 0, 0)

        # =====================================================
        # GET COORDINATE FROM HAH COORDINATES
        # =====================================================
        coord = signal_data.get(lamp_type)

        if coord is None:
            log(
                f"{signal} {lamp_type} "
                f"COORDINATE NOT AVAILABLE"
            )
            return (0, 0, 0)

        if not isinstance(coord, (list, tuple)) or len(coord) < 2:
            log(
                f"INVALID COORDINATE : "
                f"{signal} {lamp_type} : {coord}"
            )
            return (0, 0, 0)

        excel_x = float(coord[0])
        excel_y = float(coord[1])

        # =====================================================
        # CURRENT SCREEN SIZE
        # =====================================================
        screen_width, screen_height = pyautogui.size()

        # =====================================================
        # EXCEL COORDINATES ARE BASED ON 1920x1080
        # =====================================================
        BASE_WIDTH = 1920
        BASE_HEIGHT = 1080

        scale_x = screen_width / BASE_WIDTH
        scale_y = screen_height / BASE_HEIGHT

        screen_x = round(excel_x * scale_x)
        screen_y = round(excel_y * scale_y)

        # =====================================================
        # KEEP INSIDE SCREEN
        # =====================================================
        screen_x = max(
            0,
            min(screen_x, screen_width - 1)
        )

        screen_y = max(
            0,
            min(screen_y, screen_height - 1)
        )

        log(
            f"{signal} {lamp_type} : "
            f"EXCEL=({excel_x},{excel_y}) "
            f"SCREEN=({screen_x},{screen_y}) "
            f"RESOLUTION={screen_width}x{screen_height}"
        )

        # =====================================================
        # READ PIXEL
        # =====================================================
        rgb = pyautogui.pixel(
            screen_x,
            screen_y
        )

        rgb = tuple(int(v) for v in rgb)

        log(
            f"{signal} {lamp_type} RGB : {rgb}"
        )

        return rgb

    except Exception as e:

        log(
            f"GET LAMP RGB ERROR : "
            f"{signal} : {lamp_type} : "
            f"{type(e).__name__} : {e}"
        )

        return (0, 0, 0)

def read_signal_lamp_colour(signal):

    try:
        signal = str(signal).strip().upper()

        # =====================================================
        # CALLING-ON: EXPECT YELLOW ONLY
        # =====================================================
        if signal in calling_on_signals:

            yellow_rgb = get_signal_lamp_rgb(signal, "YELLOW")

            log(f"{signal} CALLING-ON YELLOW RGB : {yellow_rgb}")

            if is_yellow_lamp(yellow_rgb):
                log(f"{signal} : CALLING-ON YELLOW DETECTED")
                return "YELLOW"

            log(f"{signal} : CALLING-ON YELLOW NOT DETECTED")
            return "UNKNOWN"

        # =====================================================
        # SHUNT: EXPECT WHITE ONLY
        # =====================================================
        if signal in shunt_signals:

            white_rgb = get_signal_lamp_rgb(signal, "WHITE")

            log(f"{signal} SHUNT WHITE RGB : {white_rgb}")

            if white_rgb is not None:
                r, g, b = white_rgb
                detected = (
                    r >= 150
                    and g >= 150
                    and b >= 150
                    and abs(r - g) <= 70
                    and abs(g - b) <= 70
                    and abs(r - b) <= 70
                )
                if detected:
                    log(f"{signal} : SHUNT WHITE DETECTED")
                    return "WHITE"

            log(f"{signal} : SHUNT WHITE NOT DETECTED")
            return "UNKNOWN"

        # =====================================================
        # MAIN / STARTER: EXISTING RED/YELLOW/GREEN LOGIC
        # =====================================================

        red_rgb = get_signal_lamp_rgb(signal, "RED")
        log(f"{signal} RED RGB : {red_rgb}")

        if is_red_lamp(red_rgb):
            log(f"{signal} : RED DETECTED")
            return "RED"

        yellow_rgb = get_signal_lamp_rgb(signal, "YELLOW")
        log(f"{signal} YELLOW RGB : {yellow_rgb}")

        if is_yellow_lamp(yellow_rgb):
            log(f"{signal} : YELLOW DETECTED")
            return "YELLOW"

        green_rgb = get_signal_lamp_rgb(signal, "GREEN")
        log(f"{signal} SECOND ASPECT RGB : {green_rgb}")

        if is_yellow_lamp(green_rgb):
            log(f"{signal} : YELLOW DETECTED FROM SECOND ASPECT LOCATION")
            return "YELLOW"

        if is_green_lamp(green_rgb):
            log(f"{signal} : GREEN DETECTED FROM SECOND ASPECT LOCATION")
            return "GREEN"

        log(f"{signal} : LAMP COLOUR UNKNOWN")
        return "UNKNOWN"

    except Exception as e:
        log(
            f"LAMP READ ERROR : {signal} : "
            f"{type(e).__name__} : {e}"
        )
        return "UNKNOWN"

def analyze_toc_route_lamps(toc_item):

    route_sequence = get_toc_route_sequence(
        toc_item
    )

    results = []

    log("")
    log("========================================")
    log("ANALYZING ROUTE LAMP COLOURS")
    log("========================================")

    for route in route_sequence:

        signal = route.split("_")[0]

        colour = read_signal_lamp_colour(
            signal
        )

        if colour == "RED":

            result = "FAIL"

        elif colour in (
            "YELLOW",
            "GREEN"
        ):

            result = "PASS"

        else:

            result = "FAIL"

        result_data = {
            "signal": signal,
            "route": route,
            "lamp": colour,
            "result": result
        }

        results.append(
            result_data
        )

        log(
            f"SIGNAL : {signal} | "
            f"ROUTE : {route} | "
            f"LAMP : {colour} | "
            f"RESULT : {result}"
        )

    overall = (
        "PASS"
        if results
        and all(
            x["result"] == "PASS"
            for x in results
        )
        else "FAIL"
    )

    log(
        f"OVERALL ROUTE RESULT : "
        f"{overall}"
    )

    return results, overall

def release_toc_route_sequence(toc_item):

    route_sequence = get_toc_route_sequence(
        toc_item
    )

    if not route_sequence:

        log(
            "NO ROUTES AVAILABLE FOR RELEASE"
        )

        return False

    log("")
    log("========================================")
    log("RELEASING ROUTES")
    log("========================================")

    success = True

    # -------------------------------------------------
    # REVERSE ORDER
    # -------------------------------------------------

    for route in reversed(
        route_sequence
    ):

        signal = route.split("_")[0]

        log("")
        log("----------------------------------------")
        log(
            f"RELEASING SIGNAL : {signal}"
        )
        log(
            f"RELEASING ROUTE  : {route}"
        )
        log("----------------------------------------")

        # =================================================
        # 1. SIGNAL CANCEL
        # =================================================

        if not open_signal_menu(signal):

            log(
                f"FAILED TO OPEN SIGNAL MENU : "
                f"{signal}"
            )

            success = False
            continue

        if not find_and_click(
            "Signal Cancel"
        ):

            log(
                f"SIGNAL CANCEL FAILED : "
                f"{signal}"
            )

            success = False
            continue

        log(
            f"SIGNAL CANCEL SUCCESS : "
            f"{signal}"
        )

        time.sleep(1)

        # =================================================
        # 2. ROUTE RELEASE
        # =================================================

        if not open_signal_menu(signal):

            log(
                f"FAILED TO REOPEN SIGNAL MENU : "
                f"{signal}"
            )

            success = False
            continue

        if not find_and_click(
            "Route Release"
        ):

            log(
                f"ROUTE RELEASE FAILED : "
                f"{route}"
            )

            success = False
            continue

        log(
            f"ROUTE RELEASE SUCCESS : "
            f"{route}"
        )

        time.sleep(2)

    log("")
    log("========================================")
    log(
        "ROUTE RELEASE COMPLETED : "
        + (
            "PASS"
            if success
            else "FAIL"
        )
    )
    log("========================================")

    return success
def initialize_route_report():
    """Create ONE report for the complete automation run
    (same structure as the CASCADING report)."""

    global ROUTE_REPORT_FILE
    global report_rows

    try:
        report_folder = os.getcwd()

        os.makedirs(report_folder, exist_ok=True)

        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")

        ROUTE_REPORT_FILE = os.path.join(
            report_folder,
            f"NEGATIVE_TESTING_SIGNAL_REPORT_{timestamp}.xlsx"
        )

        report_rows = []

        write_route_report()

        log("")
        log("========================================")
        log("ROUTE REPORT CREATED")
        log(f"REPORT : {ROUTE_REPORT_FILE}")
        log("========================================")

        return True

    except Exception as e:

        log(
            f"REPORT INITIALIZATION ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return False


def write_route_report():
    """Write the whole report file from report_rows.
    Layout is the same as the CASCADING AUTOMATION REPORT."""

    if not ROUTE_REPORT_FILE:
        return False

    wb = Workbook()
    ws = wb.active
    ws.title = "NEGATIVE TESTING SIGNAL REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title_cell = ws["A1"]
    title_cell.value = "NEGATIVE TEST FOR SIGNAL, POINT AND CRANK HANDLE"
    title_cell.fill = PatternFill(fill_type=None)
    title_cell.font = Font(bold=True, size=18, color="000000")
    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 35

    # =====================================================
    # DATE
    # =====================================================

    ws.merge_cells("A2:D2")

    date_cell = ws["A2"]
    date_cell.value = (
        f"Generated : {time.strftime('%d-%m-%Y %H:%M:%S')}"
    )
    date_cell.font = Font(bold=True, size=11)
    date_cell.alignment = Alignment(horizontal="center")

    # =====================================================
    # HEADERS
    # =====================================================

    headers = [
        "SIGNAL",
        "ROUTE",
        "RESULT",
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
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font(bold=True, color="000000")
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
            row.get("SIGNAL", ""),
            row.get("ROUTE", ""),
            row.get("RESULT", ""),
            row.get("DATE & TIME", "")
        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================

    for row in ws.iter_rows(min_row=4):

        result_cell = row[2]

        for cell in row:
            cell.fill = PatternFill(fill_type=None)
            cell.font = Font(color="000000")
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
            except Exception:
                pass

        ws.column_dimensions[column_letter].width = max_length + 5

    # =====================================================
    # SUMMARY
    # =====================================================

    all_routes = set(
        x.get("ROUTE", "") for x in report_rows
    )

    initiated_routes = sorted(set(
        x.get("ROUTE", "")
        for x in report_rows
        if str(x.get("RESULT", "")).strip().upper() == "PASS"
    ))

    failed_routes = sorted(set(
        x.get("ROUTE", "")
        for x in report_rows
        if str(x.get("RESULT", "")).strip().upper() != "PASS"
    ))

    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL ROUTES"
    ws[f"B{start}"] = len(all_routes)

    ws[f"A{start + 1}"] = "INITIATED ROUTES"
    ws[f"B{start + 1}"] = ", ".join(initiated_routes)

    ws[f"A{start + 2}"] = "NON-INITIATED ROUTES"

    if failed_routes:
        ws[f"B{start + 2}"] = ", ".join(failed_routes)
    else:
        ws[f"B{start + 2}"] = "0"

    for r in range(start, start + 3):
        ws[f"A{r}"].font = Font(bold=True)

    wb.save(ROUTE_REPORT_FILE)

    return True


def add_route_result(signal, route, result):
    """Add one route result and rewrite the report file."""

    global report_rows

    report_rows.append({
        "SIGNAL": str(signal).strip(),
        "ROUTE": str(route).strip(),
        "RESULT": str(result).strip().upper(),
        "DATE & TIME": time.strftime("%d-%m-%Y %H:%M:%S")
    })

    try:
        write_route_report()
        log(f"REPORT UPDATED : {route} -> {result}")
    except Exception as e:
        log(
            f"REPORT UPDATE ERROR : "
            f"{type(e).__name__} : {e}"
        )

    return True


def save_toc_route_report(
    toc_item,
    lamp_results,
    overall_result
):
    """Add the result of one TOC route to the report."""

    return add_route_result(
        toc_item.get("signal", ""),
        toc_item.get("route", "")
        or toc_item.get("main_signal", ""),
        overall_result
    )

ROUTE_SEQUENCE = {

    "1_A": [
        ("1", "1_A"),
        ("3", "3_K")
    ],

    "1_A1": [
        ("1", "1_A1")
    ],

    "1_B": [
        ("1", "1_B"),
        ("2", "2_K")
    ],

    "1C_A": [
        ("1C", "1C_A")
    ],

    "1C_B": [
        ("1C", "1C_B")
    ],

    "2_K": [
        ("2", "2_K")
    ],

    "2C_K": [
        ("2C", "2C_K")
    ],

    "3_K": [
        ("3", "3_K")
    ],

    "3C_K": [
        ("3C", "3C_K")
    ],

    "9_A": [
        ("9", "9_A")
    ],

    "9_B": [
        ("9", "9_B")
    ],

    "17_A": [
        ("17", "17_A")
    ],

    "17_B": [
        ("17", "17_B")
    ],

    "30_H": [
        ("30", "30_H")
    ],

    "30C_H": [
        ("30C", "30C_H")
    ],

    "32_A": [
        ("32", "32_A"),
        ("30", "30_H")
    ],

    "32_A1": [
        ("32", "32_A1")
    ],

    "32_B": [
        ("32", "32_B"),
        ("31", "31_H")
    ],

    "31_H": [
        ("31", "31_H")
    ],

    "31C_H": [
        ("31C", "31C_H")
    ],
}

def get_actual_route_sequence(main_signal):

    main_signal = str(main_signal).strip().upper()

    # For Calling-On and Shunt TOC rows, the route belongs directly to
    # the signal named in that TOC row (e.g. 1C -> 1C_A, 9SH -> 9_A).
    # This avoids using a static MAIN signal alias such as 9 -> 9_A.
    # Existing MAIN/STARTER route sequences remain unchanged.
    for toc_item in toc_data:
        try:
            route_key = str(toc_item.get("main_signal", "")).strip().upper()
            signal_type = str(toc_item.get("signal_type", "")).strip().upper()
            signal = str(toc_item.get("signal", "")).strip().upper()
            route = str(toc_item.get("route", "")).strip().upper()

            if (
                route_key == main_signal
                and signal_type in ("CALLING_ON", "SHUNT")
                and signal
                and route
            ):
                sequence = [(signal, route)]
                log(
                    f"AUXILIARY ROUTE SEQUENCE FROM TOC : "
                    f"{main_signal} : {sequence}"
                )
                return sequence
        except Exception:
            continue

    sequence = ROUTE_SEQUENCE.get(main_signal)

    if sequence is None:
        log(
            f"NO ACTUAL ROUTE SEQUENCE FOUND "
            f"FOR MAIN SIGNAL : {main_signal}"
        )

        log(
            f"AVAILABLE ROUTE SEQUENCES : "
            f"{list(ROUTE_SEQUENCE.keys())}"
        )

        return []

    log(
        f"ACTUAL ROUTE SEQUENCE FOR "
        f"{main_signal} : {sequence}"
    )

    return sequence

def find_and_click(name, control_type=None):
    """
    Search for an item in the currently opened GSIM signal menu
    and click the matching menu item.

    Example:
        Signal 1 menu opened
        -> search 1_A
        -> click 1_A
    """

    target = str(name).strip().upper()

    log(f"SEARCHING MENU ITEM : {target}")

    try:
        desktop = Desktop(backend="uia")

        # ---------------------------------------------------------
        # SEARCH ONLY VISIBLE WINDOWS / CONTROLS
        # ---------------------------------------------------------
        for window in desktop.windows():

            try:
                if not window.is_visible():
                    continue
            except Exception:
                continue

            # -----------------------------------------------------
            # SEARCH CHILD CONTROLS
            # -----------------------------------------------------
            try:
                controls = window.descendants()
            except Exception:
                continue

            for control in controls:

                try:
                    if not control.is_visible():
                        continue
                except Exception:
                    pass

                try:
                    text = control.window_text().strip()
                except Exception:
                    continue

                if not text:
                    continue

                # -------------------------------------------------
                # EXACT MATCH
                # -------------------------------------------------
                if text.upper() != target:
                    continue

                # -------------------------------------------------
                # CONTROL TYPE CHECK
                # -------------------------------------------------
                if control_type:
                    try:
                        if control.element_info.control_type != control_type:
                            continue
                    except Exception:
                        continue

                # -------------------------------------------------
                # GET POSITION
                # -------------------------------------------------
                try:
                    rect = control.rectangle()

                    x = (rect.left + rect.right) // 2
                    y = (rect.top + rect.bottom) // 2

                except Exception:
                    continue

                log(
                    f"FOUND MENU ITEM : {target} "
                    f"AT ({x},{y})"
                )

                # -------------------------------------------------
                # CLICK THE ACTUAL MENU ITEM
                # -------------------------------------------------
                pyautogui.moveTo(
                    x,
                    y,
                    duration=0.2
                )

                time.sleep(0.3)

                pyautogui.click()

                time.sleep(1.5)

                log(
                    f"MENU ITEM CLICKED SUCCESSFULLY : {target}"
                )

                return True

        log(
            f"MENU ITEM NOT FOUND IN OPEN MENU : {target}"
        )

        return False

    except Exception as e:

        log(
            f"FIND AND CLICK ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return False
def get_signal_menu(signal):

    signal = str(signal).strip().upper()

    # Remove route suffix.
    # Example:
    # 1_A -> 1
    # 3_K -> 3
    lookup = signal.split("_")[0]

    log(
        f"GET SIGNAL MENU : "
        f"requested=[{signal}] "
        f"lookup=[{lookup}]"
    )

    # MAIN
    if lookup in signals:

        menu = signals[lookup].get("menu")

        if menu:
            return menu

    # STARTER
    if lookup in starter_signals:

        menu = starter_signals[lookup].get("menu")

        if menu:
            return menu

    # CALLING ON
    if lookup in calling_on_signals:

        menu = calling_on_signals[lookup].get("menu")

        if menu:
            return menu

    # SHUNT
    if lookup in shunt_signals:

        menu = shunt_signals[lookup].get("menu")

        if menu:
            return menu

    log(
        f"SIGNAL MENU NOT FOUND : {signal}"
    )

    return None

def close_station_control_bits_window():

    log(
        "SEARCHING STATION 50051 CONTROL BITS WINDOW"
    )

    try:

        desktop = Desktop(
            backend="uia"
        )

        for window in desktop.windows():

            try:

                title = (
                    window.window_text()
                    .strip()
                    .upper()
                )

                if (
                    "STATION 50051" not in title
                    or "INDICATIONS" not in title
                    or "CONTROLS" not in title
                ):
                    continue

                rect = window.rectangle()

                # Top-right X
                x = rect.right - 15
                y = rect.top + 15

                log(
                    f"CONTROL BITS WINDOW FOUND : "
                    f"{title}"
                )

                log(
                    f"CLICKING X : ({x},{y})"
                )

                pyautogui.moveTo(
                    x,
                    y,
                    duration=0.2
                )

                pyautogui.click()

                time.sleep(1)

                log(
                    "STATION 50051 CONTROL BITS CLOSED"
                )

                return True

            except Exception:
                continue

    except Exception as e:

        log(
            f"CLOSE CONTROL BITS ERROR : "
            f"{type(e).__name__} : {e}"
        )

    log(
        "CONTROL BITS WINDOW NOT FOUND"
    )

    return False

def transmit_initial_control_bits():
    """Select all INITIAL_CONTROL_BITS relays and transmit."""

    log("========================================")
    log("INITIAL CONTROL BIT DEPLOYMENT")
    log("========================================")

    relays = read_initial_control_bits()

    if not relays:
        # No INITIAL_CONTROL_BITS sheet / no relays listed.
        # Do not block the run (a popup would hang walk-away mode).
        log(
            "NO INITIAL CONTROL BITS FOUND - STEP SKIPPED "
            "(add an INITIAL_CONTROL_BITS sheet with a RELAY column "
            "to the TOC file if this step is needed)"
        )
        return True

    # Open station 50051
    if not open_test_panel_station():
        return False

    time.sleep(2)

    # Select each relay from Excel
    for relay in relays:

        log(
            f"SELECTING INITIAL RELAY : {relay}"
        )

        if not find_station_relay(relay):

            log(f"INITIAL RELAY NOT FOUND : {relay}")

            return False

        time.sleep(1)

    # Click TRANSMIT
    log("SEARCHING TRANSMIT BUTTON")

    if not find_ui_element(
        "Transmit",
        control_type="Button"
    ):
        log("TRANSMIT BUTTON NOT FOUND")

        return False

    # Wait for Transmit operation to complete
    time.sleep(2)

    # Click the Control Bits window's Cancel button after Transmit.
    log("SEARCHING CANCEL BUTTON")
    if not find_ui_element("Cancel", control_type="Button"):
        log("CANCEL BUTTON NOT FOUND AFTER INITIAL TRANSMIT")
        return False
    log("CANCEL CLICKED AFTER INITIAL TRANSMIT")
    time.sleep(1)

    log(
        "INITIAL CONTROL BIT DEPLOYMENT COMPLETED"
    )

    # -------------------------------------------------
    # CLOSE STATION 50051 CONTROL BITS WINDOW
    # -------------------------------------------------

    if not close_station_control_bits_window():

        log(
            "WARNING : "
            "CONTROL BITS WINDOW COULD NOT BE CLOSED"
        )

    else:

        log(
            "CONTROL BITS WINDOW CLOSED"
        )

    time.sleep(1)

    return True
def release_actual_route_sequence(main_signal):

    main_signal = str(
        main_signal
    ).strip().upper()

    sequence = get_actual_route_sequence(
        main_signal
    )

    if not sequence:

        log(
            f"NO ROUTE SEQUENCE FOR RELEASE : "
            f"{main_signal}"
        )

        return False

    log("")
    log("========================================")
    log("RELEASING ACTUAL ROUTE")
    log("========================================")

    success = True

    # =====================================================
    # REVERSE ACTUAL ROUTE SEQUENCE
    # =====================================================

    for step_no, (
        signal,
        route
    ) in enumerate(
        reversed(sequence),
        start=1
    ):

        signal = str(
            signal
        ).strip().upper()

        route = str(
            route
        ).strip().upper()

        log("")
        log("----------------------------------------")
        log(
            f"RELEASE STEP {step_no}"
        )
        log(
            f"SIGNAL : {signal}"
        )
        log(
            f"ROUTE  : {route}"
        )
        log("----------------------------------------")

        # =================================================
        # SIGNAL CANCEL
        # =================================================

        if not open_signal_menu(signal):

            log(
                f"❌ SIGNAL MENU OPEN FAILED : "
                f"{signal}"
            )

            success = False
            continue

        if not find_and_click(
            "Signal Cancel"
        ):

            log(
                f"❌ SIGNAL CANCEL FAILED : "
                f"{signal}"
            )

            success = False
            continue

        log(
            f"✅ SIGNAL CANCEL : "
            f"{signal}"
        )

        time.sleep(1)

        # =================================================
        # ROUTE RELEASE
        # =================================================

        if not open_signal_menu(signal):

            log(
                f"❌ SIGNAL MENU REOPEN FAILED : "
                f"{signal}"
            )

            success = False
            continue

        if not find_and_click(
            "Route Release"
        ):

            log(
                f"❌ ROUTE RELEASE FAILED : "
                f"{route}"
            )

            success = False
            continue

        log(
            f"✅ ROUTE RELEASED : "
            f"{route}"
        )

        time.sleep(2)

    log("")
    log("========================================")
    log(
        "ACTUAL ROUTE RELEASE RESULT : "
        + (
            "PASS"
            if success
            else "FAIL"
        )
    )
    log("========================================")

    return success
def activate_toc_control_bits(
    signal_bits,
    point_bits,
    crank_bits,
    track
):

    log("")
    log("========================================")
    log("ACTIVATING TOC CONTROL BITS")
    log("========================================")

    all_bits = []

    # TRACK
    if track:

        if isinstance(track, str):

            track_bits = [
                x.strip()
                for x in track.split(",")
                if x.strip()
            ]

        else:
            track_bits = [str(track)]

        all_bits.extend(
            track_bits
        )

    # SIGNAL
    all_bits.extend(
        signal_bits or []
    )

    # POINT
    all_bits.extend(
        point_bits or []
    )

    # CRANK
    all_bits.extend(
        crank_bits or []
    )

    if not all_bits:
        log(
            "NO TOC CONTROL BITS FOUND"
        )

        log(
            "BLANK CONTROL DATA - "
            "SKIPPING CURRENT TOC ITEM"
        )

        return True
    # -------------------------------------------------
    # OPEN CONTROL BITS
    # -------------------------------------------------

    if not open_test_panel_station():

        log(
            "FAILED TO OPEN STATION 50051"
        )

        return False

    time.sleep(2)

    # -------------------------------------------------
    # SELECT RELAYS
    # -------------------------------------------------

    for relay in all_bits:

        relay = str(
            relay
        ).strip()

        if not relay:
            continue

        log(
            f"SELECTING TOC RELAY : "
            f"{relay}"
        )

        if not find_station_relay(
            relay
        ):

            log(
                f"TOC RELAY NOT FOUND : "
                f"{relay}"
            )

            return False

        time.sleep(0.3)

    # -------------------------------------------------
    # TRANSMIT
    # -------------------------------------------------

    log(
        "SEARCHING TRANSMIT BUTTON"
    )

    if not find_ui_element(
        "Transmit",
        control_type="Button"
    ):

        log(
            "TRANSMIT BUTTON NOT FOUND"
        )

        return False

    time.sleep(2)

    log(
        "TOC CONTROL BITS TRANSMITTED"
    )

    # -------------------------------------------------
    # CANCEL CONTROL BITS
    # -------------------------------------------------
    log("SEARCHING CANCEL BUTTON")
    if not find_ui_element(
        "Cancel",
        control_type="Button"
    ):
        log("CANCEL BUTTON NOT FOUND AFTER TOC TRANSMIT")
        return False

    log("TOC CONTROL BITS CANCEL CLICKED")
    time.sleep(1)

    # -------------------------------------------------
    # CLOSE CONTROL BITS
    # -------------------------------------------------

    close_station_control_bits_window()

    time.sleep(1)

    log(
        "TOC CONTROL BITS WINDOW CLOSED"
    )

    return True
def restore_toc_control_bits(
    signal_bits,
    point_bits,
    crank_bits,
    track
):

    log("")
    log("========================================")
    log("RESTORING TOC CONTROL BITS")
    log("========================================")

    all_bits = []

    all_bits.extend(
        signal_bits or []
    )

    all_bits.extend(
        point_bits or []
    )

    all_bits.extend(
        crank_bits or []
    )

    if track:

        if isinstance(track, str):

            all_bits.extend(
                [
                    x.strip()
                    for x in track.split(",")
                    if x.strip()
                ]
            )

        else:
            all_bits.append(
                str(track)
            )

    if not all_bits:

        log(
            "NO CONTROL BITS TO RESTORE"
        )

        return True

    if not open_test_panel_station():

        log(
            "FAILED TO OPEN CONTROL BITS"
        )

        return False

    time.sleep(2)

    success = True

    for relay in all_bits:

        relay = str(
            relay
        ).strip()

        if not relay:
            continue

        log(
            f"RESTORING RELAY : "
            f"{relay}"
        )

        if not find_station_relay(
            relay
        ):

            log(
                f"FAILED TO RESTORE : "
                f"{relay}"
            )

            success = False

        time.sleep(0.3)

    if find_ui_element(
        "Transmit",
        control_type="Button"
    ):

        time.sleep(2)

        log(
            "CONTROL BIT RESTORE TRANSMITTED"
        )

        # Click Cancel after the restore Transmit as well.
        log("SEARCHING CANCEL BUTTON AFTER RESTORE")
        if find_ui_element(
            "Cancel",
            control_type="Button"
        ):
            log("CONTROL BIT RESTORE CANCEL CLICKED")
            time.sleep(1)
        else:
            log("RESTORE CANCEL BUTTON NOT FOUND")
            success = False

    else:

        log(
            "RESTORE TRANSMIT BUTTON NOT FOUND"
        )

        success = False

    # =================================================
    # CLOSE CONTROL BITS WINDOW
    # =================================================

    log(
        "CLOSING CONTROL BITS WINDOW "
        "AFTER RESTORE"
    )

    window_closed = (
        close_station_control_bits_window()
    )

    if window_closed:

        log(
            "CONTROL BITS WINDOW CLOSED SUCCESSFULLY"
        )

    else:

        log(
            "WARNING : CONTROL BITS WINDOW "
            "COULD NOT BE CLOSED"
        )

    time.sleep(1)

    return success

# =========================================================
# EDRC AUTO CONTROL BRIDGE
# =========================================================

def load_universal_coordinates():
    """Load the coordinate workbook supplied by the EDRC launcher."""
    global CONFIG_FILE

    if MASTER_MODE:
        coords_path = os.environ.get("EDRC_COORDS", "").strip()
        if not coords_path or not os.path.exists(coords_path):
            log("EDRC COORDINATE PATH NOT FOUND")
            log(f"EDRC_COORDS : {coords_path}")
            return False

        CONFIG_FILE = coords_path
        log(f"EDRC COORDINATE PATH : {CONFIG_FILE}")
        loaded = load_config()
        if loaded is False:
            log("EDRC COORDINATES LOAD FAILED")
            return False

        log("EDRC COORDINATES LOADED SUCCESSFULLY")
        return True

    # Manual mode keeps the original coordinate-file workflow.
    try:
        open_existing_config()
        return bool(CONFIG_FILE)
    except Exception as exc:
        log(f"COORDINATE LOAD ERROR : {type(exc).__name__} : {exc}")
        return False


def start_testing_from_ui():
    """EDRC start bridge; the existing start_testing engine is unchanged."""
    global running
    global automation_running

    if running or automation_running:
        log("EDRC AUTOMATION ALREADY RUNNING - START IGNORED")
        return False

    # Set this BEFORE scheduling the existing engine. The EDRC launcher
    # checks this module-level flag immediately after pressing START.
    running = True
    automation_running = True

    log("[EDRC] START ACCEPTED")
    log(f"[EDRC] TOC={TOC_FILE}")
    log(f"[EDRC] CONFIG={CONFIG_FILE}")
    log("[EDRC] QUEUING EXISTING TEST ENGINE ON TK MAIN THREAD")

    try:
        # Run the engine in a BACKGROUND THREAD.
        # Previously it ran on the Tk main thread, which blocked the UI
        # for the whole run and made the launcher believe the program
        # had never started.
        threading.Thread(
            target=start_testing,
            daemon=True
        ).start()
        return True
    except Exception as exc:
        running = False
        automation_running = False
        log(f"[EDRC] START QUEUE ERROR : {type(exc).__name__} : {exc}")
        return False


def start_testing():

    global running
    global automation_running
    global route_selection_started

    running = True
    automation_running = True

    set_status(
        "STARTING AUTOMATION",
        GREEN
    )

    log("")
    log("========================================")
    log("START AUTOMATION")
    log("========================================")

    # =====================================================
    # VALIDATION
    # =====================================================

    if not TOC_FILE:

        log("START FAILED : TOC FILE NOT LOADED")

        running = False
        automation_running = False
        return

    if not CONFIG_FILE:

        log("START FAILED : CONFIGURATION FILE NOT LOADED")

        running = False
        automation_running = False
        return

    if not toc_data:

        log("START FAILED : TOC DATA IS EMPTY")

        running = False
        automation_running = False
        return

    # =====================================================
    # 5 SECOND DELAY
    # =====================================================

    log(
        "AUTOMATION STARTING IN 5 SECONDS..."
    )

    for i in range(
        5,
        0,
        -1
    ):

        log(f"{i}...")

        time.sleep(1)

    # =====================================================
    # INITIALIZE ROUTE REPORT
    # =====================================================

    if not initialize_route_report():

        log("❌ ROUTE REPORT INITIALIZATION FAILED")

        running = False
        automation_running = False

        set_status(
            "REPORT INITIALIZATION FAILED",
            RED
        )

        return

    log("ROUTE REPORT INITIALIZED SUCCESSFULLY")
    # =====================================================
    # INITIAL POINTS - SET TO NORMAL
    # =====================================================

    # if not set_initial_points_normal():
    #     log(
    #         "INITIAL POINT NORMAL SETTING FAILED"
    #     )
    #
    #     running = False
    #     automation_running = False
    #
    #     set_status(
    #         "START FAILED",
    #         RED
    #     )
    #
    #     return
    #
    # log(
    #     "INITIAL POINTS 50A, 50B, 65A, 65B "
    #     "SET TO NORMAL"
    # )
    # =====================================================
    # INITIAL CONTROL BITS
    # =====================================================
    # =====================================================
    # INITIAL CONTROL BITS
    # =====================================================

    if not transmit_initial_control_bits():

        log(
            "INITIAL CONTROL BIT DEPLOYMENT FAILED"
        )

        running = False
        automation_running = False

        set_status(
            "START FAILED",
            RED
        )

        return

    log(
        "INITIAL CONTROL BITS COMPLETED"
    )

    # =====================================================
    # PROCESS EVERY TOC ROW
    # =====================================================

    for toc_item in toc_data:

        if not automation_running:
            break

        main_signal = str(
            toc_item.get(
                "main_signal",
                ""
            )
        ).strip().upper()

        log("")
        log("========================================")
        log(
            f"PROCESSING TOC : "
            f"{main_signal}"
        )
        log("========================================")

        signal_bits = toc_item.get(
            "signal_bits",
            []
        )

        point_bits = toc_item.get(
            "point_bits",
            []
        )

        crank_bits = toc_item.get(
            "crank_bits",
            []
        )

        track = toc_item.get(
            "track",
            ""
        )

        # =================================================
        # CALLING-ON / SHUNT
        # =================================================

        calling_on_list = [
            str(x).strip().upper()
            for x in toc_item.get(
                "calling_on_signals",
                []
            )
            if str(x).strip()
        ]

        shunt_list = [
            str(x).strip().upper()
            for x in toc_item.get(
                "shunt_signals",
                []
            )
            if str(x).strip()
        ]

        log(
            f"CALLING-ON SIGNALS : "
            f"{calling_on_list if calling_on_list else 'NONE'}"
        )

        log(
            f"SHUNT SIGNALS : "
            f"{shunt_list if shunt_list else 'NONE'}"
        )

        # Compatibility variables
        calling_on = (
            calling_on_list[0]
            if calling_on_list
            else ""
        )

        shunt = (
            shunt_list[0]
            if shunt_list
            else ""
        )
        # =================================================
        # REMOVE BLANK VALUES
        # =================================================

        signal_bits = [
            str(x).strip().upper()
            for x in (signal_bits or [])
            if str(x).strip()
        ]

        point_bits = [
            str(x).strip().upper()
            for x in (point_bits or [])
            if str(x).strip()
        ]

        crank_bits = [
            str(x).strip().upper()
            for x in (crank_bits or [])
            if str(x).strip()
        ]

        track = str(
            track or ""
        ).strip().upper()

        log(
            f"VALID SIGNAL BITS : {signal_bits}"
        )

        log(
            f"VALID POINT BITS : {point_bits}"
        )

        log(
            f"VALID CRANK BITS : {crank_bits}"
        )

        log(
            f"VALID TRACK : {track if track else 'NONE'}"
        )

        # =================================================
        # CALLING-ON / SHUNT SIGNALS
        # =================================================
        #
        # IMPORTANT:
        # Calling-On and Shunt come from the existing
        # SIGNAL column of the TOC.
        #
        # upload_toc() stores them as:
        #     calling_on_signals
        #     shunt_signals
        #
        # These variables are kept here as local automation
        # variables so the existing automation does not
        # generate "name 'calling_on' is not defined".
        # =================================================

        calling_on_list = toc_item.get(
            "calling_on_signals",
            []
        )

        shunt_list = toc_item.get(
            "shunt_signals",
            []
        )

        # -------------------------------------------------
        # COMPATIBILITY VARIABLES
        # -------------------------------------------------

        calling_on = (
            calling_on_list[0]
            if calling_on_list
            else ""
        )

        shunt = (
            shunt_list[0]
            if shunt_list
            else ""
        )

        # -------------------------------------------------
        # LOG
        # -------------------------------------------------

        log(
            f"CALLING-ON SIGNALS : "
            f"{calling_on_list if calling_on_list else 'NONE'}"
        )

        log(
            f"SHUNT SIGNALS : "
            f"{shunt_list if shunt_list else 'NONE'}"
        )
        # =================================================
        # CONTROL BITS FOR CURRENT TOC ROW
        # =================================================

        if not activate_toc_control_bits(
            signal_bits,
            point_bits,
            crank_bits,
            track
        ):

            log(
                f"CONTROL BIT ACTIVATION FAILED : "
                f"{main_signal}"
            )

            add_route_result(
                toc_item.get("signal", ""),
                main_signal,
                "FAIL"
            )

            continue

        # =================================================
        # ESTABLISH ALL ROUTES
        # =================================================

        if not establish_actual_route_sequence(
                main_signal
        ):

            log(
                f"❌ ACTUAL ROUTE ESTABLISHMENT FAILED : "
                f"{main_signal}"
            )

            add_route_result(
                toc_item.get("signal", ""),
                main_signal,
                "FAIL"
            )

            # Cleanup if at least one route was selected
            if route_selection_started:
                release_actual_route_sequence(
                    main_signal
                )

            continue
        # =================================================
        # WAIT FOR MAP TO SETTLE
        # =================================================

        log(
            "FULL ROUTE ESTABLISHED"
        )
        log(
            "WAITING 4 SECONDS BEFORE LAMP ANALYSIS..."
        )

        time.sleep(4)

        log(
            "4 SECOND WAIT COMPLETED - STARTING LAMP ANALYSIS"
        )

        # =================================================
        # ANALYZE ACTUAL ROUTE LAMPS
        # =================================================

        lamp_results, overall_result = (
            analyze_actual_route_lamps(
                main_signal
            )
        )

        # =================================================
        # UPDATE REPORT BEFORE RELEASE
        # =================================================

        save_toc_route_report(
            toc_item,
            lamp_results,
            overall_result
        )

        # =================================================
        # RELEASE ACTUAL ROUTES IN REVERSE
        # =================================================

        release_actual_route_sequence(
            main_signal
        )
        # =================================================
        # RESTORE CONTROL BITS
        # =================================================

        restore_toc_control_bits(
            signal_bits,
            point_bits,
            crank_bits,
            track
        )

        route_selection_started = False

        log("")
        log(
            f"TOC COMPLETED : "
            f"{main_signal}"
        )
    # =====================================================
    # RESTORE INITIAL CONTROL BITS
    # =====================================================

    log("")
    log("========================================")
    log("ALL TOC ROUTES COMPLETED")
    log("STARTING FINAL INITIAL CONTROL BIT RESTORE")
    log("========================================")

    if not restore_initial_control_bits():

        log(
            "WARNING : INITIAL CONTROL BITS "
            "COULD NOT BE RESTORED"
        )

    else:

        log(
            "FINAL INITIAL CONTROL BIT RESTORE COMPLETED"
        )

    # =====================================================
    # ALL POINTS BACK TO NORMAL
    # =====================================================

    set_all_points_normal()
    # =====================================================
    # COMPLETE
    # =====================================================

    automation_running = False
    running = False

    set_status(
        "AUTOMATION COMPLETED",
        GREEN
    )

    log("")
    log("========================================")
    log("TOC AUTOMATION COMPLETED")
    log("========================================")

    log("TOC ROUTE AUTOMATION COMPLETED")

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
    text="NEGATIVE TESTING FOR SIGNAL, POINT AND CRANK HANDLE",
    bg="#0b1220",
    fg="white",
    font=("Segoe UI", 24, "bold")
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
# LEFT PANEL
# =========================================================

SIDE_MARGIN = 20
BUTTON_GAP = 10
TOP_MARGIN = 18

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
    padx=(0, 15)
)
left_panel.pack_propagate(False)


tk.Label(
    left_panel,
    text="CONTROL PANEL",
    bg="white",
    fg="#0f172a",
    font=("Segoe UI", 18, "bold")
).pack(
    pady=(12, 22)
)


# =========================================================
# CAPTURE SIGNALLING GEARS
# =========================================================

tk.Label(
    left_panel,
    text="CAPTURE SIGNALLING GEARS",
    bg=BTN,
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
    pady=(0, BUTTON_GAP)
)

new_btn = tk.Button(
    signal_frame,
    text="NEW SIGNAL",
    command=create_setup,
    bg="#f97316",
    fg="white",
    relief="flat",
    bd=0,
    font=("Segoe UI", 10, "bold"),
    cursor="hand2"
)
new_btn.pack(
    side="left",
    expand=True,
    fill="x",
    padx=(0, 5),
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
    font=("Segoe UI", 10, "bold"),
    cursor="hand2"
)
edit_btn.pack(
    side="left",
    expand=True,
    fill="x",
    padx=(5, 0),
    ipady=8
)


# =========================================================
# MAIN CONTROL BUTTONS
# =========================================================

create_panel_button(
    "IMPORT MASTER TOC",
    upload_toc,
    BTN
)

create_panel_button(
    "SAVE COORDINATES",
    save_config,
    "#16a34a"
)

create_panel_button(
    "START AUTOMATION",
    start_testing,
    "#16a34a"
)

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
).pack(fill="x")


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
    pady=(0, 20)
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
    pady=(15, 6)
)

status_label = tk.Label(
    status_frame,
    text="READY",
    bg="#f8fafc",
    fg=BTN,
    font=("Segoe UI", 14, "bold")
)
status_label.pack(pady=(0, 15))


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
# NOTEBOOK
# =========================================================

notebook = ttk.Notebook(right_panel)
notebook.pack(
    fill="both",
    expand=True
)


# =========================================================
# MAIN SIGNALS TAB
# =========================================================

main_tab = tk.Frame(notebook, bg="white")
notebook.add(main_tab, text="MAIN SIGNALS")

main_tree_scroll = ttk.Scrollbar(
    main_tab,
    orient="vertical"
)
main_tree_scroll.pack(side="right", fill="y")

main_tree = ttk.Treeview(
    main_tab,
    columns=("NO", "SIGNAL", "ROUTES"),
    show="headings",
    yscrollcommand=main_tree_scroll.set
)

main_tree.heading("NO", text="NO")
main_tree.heading("SIGNAL", text="SIGNAL")
main_tree.heading("ROUTES", text="TOTAL ROUTES")

main_tree.column("NO", width=80, anchor="center")
main_tree.column("SIGNAL", width=250, anchor="center")
main_tree.column("ROUTES", width=200, anchor="center")

main_tree.pack(
    fill="both",
    expand=True,
    padx=8,
    pady=8
)

main_tree_scroll.config(command=main_tree.yview)


# =========================================================
# CALLING-ON SIGNALS TAB
# =========================================================

cal_tab = tk.Frame(notebook, bg="white")
notebook.add(cal_tab, text="CALLING-ON SIGNALS")

cal_tree_scroll = ttk.Scrollbar(
    cal_tab,
    orient="vertical"
)
cal_tree_scroll.pack(side="right", fill="y")

cal_tree = ttk.Treeview(
    cal_tab,
    columns=("NO", "SIGNAL", "ROUTE", "TRACK"),
    show="headings",
    yscrollcommand=cal_tree_scroll.set
)

cal_tree.heading("NO", text="NO")
cal_tree.heading("SIGNAL", text="SIGNAL")
cal_tree.heading("ROUTE", text="ROUTE")
cal_tree.heading("TRACK", text="TRACK")

cal_tree.column("NO", width=80, anchor="center")
cal_tree.column("SIGNAL", width=220, anchor="center")
cal_tree.column("ROUTE", width=300, anchor="center")
cal_tree.column("TRACK", width=220, anchor="center")

cal_tree.pack(
    fill="both",
    expand=True,
    padx=8,
    pady=8
)

cal_tree_scroll.config(command=cal_tree.yview)


# =========================================================
# SHUNT SIGNALS TAB
# =========================================================

shunt_tab = tk.Frame(notebook, bg="white")
notebook.add(shunt_tab, text="SHUNT SIGNALS")

shunt_tree_scroll = ttk.Scrollbar(
    shunt_tab,
    orient="vertical"
)
shunt_tree_scroll.pack(side="right", fill="y")

shunt_tree = ttk.Treeview(
    shunt_tab,
    columns=("NO", "SIGNAL", "ROUTES"),
    show="headings",
    yscrollcommand=shunt_tree_scroll.set
)

shunt_tree.heading("NO", text="NO")
shunt_tree.heading("SIGNAL", text="SIGNAL")
shunt_tree.heading("ROUTES", text="TOTAL ROUTES")

shunt_tree.column("NO", width=80, anchor="center")
shunt_tree.column("SIGNAL", width=250, anchor="center")
shunt_tree.column("ROUTES", width=200, anchor="center")

shunt_tree.pack(
    fill="both",
    expand=True,
    padx=8,
    pady=8
)

shunt_tree_scroll.config(command=shunt_tree.yview)


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
    pady=(15, 10)
)

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

log_scroll = ttk.Scrollbar(
    log_frame,
    orient="vertical"
)
log_scroll.pack(side="right", fill="y")

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

log_scroll.config(command=log_text.yview)
log_text.config(state="disabled")


# =========================================================
# FOOTER
# =========================================================

footer = tk.Label(
    root,
    text="SPACE KEY = NEGATIVE TEST FOR SIGNAL | POINT | CRANK HANDLE",
    bg=BG,
    fg="white",
    font=("Segoe UI", 10)
)
footer.pack(fill="x")


# =========================================================
# KEYBOARD SHORTCUT
# =========================================================

root.bind("<KeyPress-p>", lambda event: toggle_pause())


# =========================================================
# INITIAL UI LOG
# =========================================================
log("UI initialized successfully.")
log("BACKEND AUTOMATION ENABLED.")
set_status("READY", BTN)

# =========================================================
# RUN
# =========================================================

root.mainloop()
