
# =========================================================
# NEGATIVE TESTING FOR VITAL INPUTS
# FINAL - TL STYLE SIGNAL + TRACK CAPTURE
# =========================================================
# pip install pyautogui openpyxl pywinauto pynput pillow uiautomation pywin32
# =========================================================

import os
import time
import threading
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

import openpyxl
import pyautogui
import uiautomation as auto
import win32api
import win32con

from PIL import ImageGrab
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pynput import keyboard as pynput_keyboard

def debug_track_pixel(track_name):
    if track_name not in track_points:
        print(f"{track_name} coordinate not found")
        return

    x, y = track_points[track_name]

    screenshot = pyautogui.screenshot()

    print(f"{track_name} coordinate = ({x},{y})")
    print(f"{track_name} RGB = {screenshot.getpixel((x, y))}")

    screenshot.save(f"DEBUG_{track_name}.png")

# =========================================================
# FILES
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
COORD_FILE = os.environ.get(
    "EDRC_TRACK_COORDS",
    os.path.join(
        os.path.expanduser("~"),
        "Desktop",
        "track_coordinates.xlsx"
    )
)
# Report is written to the CURRENT WORKING FOLDER (same as the other
# EDRC programs), so the EDRC launcher finds it in its session folder.
# (Before, it was written next to the .py file - e.g. Downloads - where
# the launcher does not look, giving "NO NEW REPORT FOUND".)
REPORT_FILE = os.path.join(os.getcwd(), "NEGATIVE_TESTING_TRACKS_REPORT.xlsx")


# =========================================================
# GLOBALS
# =========================================================

signals = {}
lock_routes_data = []
track_points = {}
point_points = {}

track_capture_list = []
track_capture_index = 0
track_undo_stack = []

capture_module = None
capture_master_list = []
capture_index = 0
record_stage = None
current_aspects = 2
undo_stack = []

capture_overlay = None
overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None

running = False
paused = False
pause_event = threading.Event()
pause_event.set()

report_rows = []

last_space_time = 0
last_backspace_time = 0
last_digit_time = 0


# =========================================================
# COLORS
# =========================================================

BG = "#0f172a"
BLUE = "#2563eb"
GREEN = "#16a34a"
RED = "#dc2626"
ORANGE = "#ea580c"
PURPLE = "#9333ea"

TYPE_LABEL = {
    "MAIN": "MAIN SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "CALLING_ON": "CALLING-ON SIGNAL",
    "POINT": "POINT",
    "CH": "CRANK HANDLE",
    "LC": "LC GATE",
}

TYPE_COLOR = {
    "MAIN": "#3b82f6",
    "SHUNT": "#a855f7",
    "CALLING_ON": "#f97316",
    "POINT": "#14b8a6",
    "CH": "#eab308",
    "LC": "#f43f5e",
}


# =========================================================
# LOG
# =========================================================

def log(msg):
    def write():
        try:
            log_text.config(state="normal")
            log_text.insert(
                tk.END,
                f"[{time.strftime('%H:%M:%S')}] {msg}\n"
            )
            log_text.see(tk.END)
            log_text.config(state="disabled")
        except Exception:
            pass

    try:
        root.after(0, write)
    except Exception:
        pass


# =========================================================
# BASIC CLICK
# =========================================================

def click(point):
    if not point:
        return False

    try:
        x, y = int(point[0]), int(point[1])
        win32api.SetCursorPos((x, y))
        time.sleep(0.35)
        pause_event.wait()
        win32api.mouse_event(
            win32con.MOUSEEVENTF_LEFTDOWN, 0, 0
        )
        time.sleep(0.08)
        win32api.mouse_event(
            win32con.MOUSEEVENTF_LEFTUP, 0, 0
        )
        time.sleep(0.4)
        return True
    except Exception as e:
        log(f"CLICK ERROR : {e}")
        return False


# =========================================================
# TL STYLE COUNT / NAME DIALOGS
# =========================================================

def center_window(win, width, height):
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(
        f"{width}x{height}+{(sw-width)//2}+{(sh-height)//2}"
    )


def ask_count(label):
    result = {"value": None}

    win = tk.Toplevel(root)
    win.title("How Many?")
    win.configure(bg=BG)
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.transient(root)
    win.grab_set()

    tk.Label(
        win, text="HOW MANY?",
        font=("Segoe UI", 11, "bold"),
        bg=BG, fg="#64748b"
    ).pack(pady=(20, 4))

    tk.Label(
        win,
        text=f"How many {label} are in this yard?",
        font=("Segoe UI", 15, "bold"),
        bg=BG, fg="white",
        wraplength=340,
        justify="center"
    ).pack(pady=(0, 15))

    var = tk.StringVar(value="0")
    entry = tk.Entry(
        win, textvariable=var,
        font=("Segoe UI", 20, "bold"),
        justify="center", width=6,
        bg="#1e293b", fg="white",
        insertbackground="white", relief="flat"
    )
    entry.pack(ipady=6)
    entry.focus_set()
    entry.select_range(0, tk.END)

    tk.Label(
        win, text="Enter 0 to skip this category",
        font=("Segoe UI", 9),
        bg=BG, fg="#475569"
    ).pack(pady=(5, 15))

    def confirm():
        try:
            result["value"] = max(0, int(var.get().strip() or "0"))
            win.grab_release()
            win.destroy()
        except ValueError:
            messagebox.showerror(
                "INVALID NUMBER",
                "Enter a valid whole number.",
                parent=win
            )

    def cancel():
        result["value"] = None
        win.grab_release()
        win.destroy()

    f = tk.Frame(win, bg=BG)
    f.pack()

    tk.Button(
        f, text="CONFIRM", command=confirm,
        bg=BLUE, fg="white", relief="flat",
        font=("Segoe UI", 10, "bold"), width=12
    ).pack(side="left", padx=6)

    tk.Button(
        f, text="CANCEL", command=cancel,
        bg=BG, fg="#64748b", relief="flat",
        font=("Segoe UI", 10, "underline"), width=10
    ).pack(side="left", padx=6)

    win.bind("<Return>", lambda e: confirm())
    win.bind("<Escape>", lambda e: cancel())
    center_window(win, 400, 260)
    win.wait_window()
    return result["value"]


def ask_name(label, index, total, existing):
    while True:
        result = {"value": None}

        win = tk.Toplevel(root)
        win.title("Signal / Element ID")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.attributes("-topmost", True)
        win.transient(root)
        win.grab_set()

        tk.Label(
            win,
            text=f"{label.upper()}   |   {index + 1} OF {total}",
            font=("Segoe UI", 11, "bold"),
            bg=BG, fg="#64748b"
        ).pack(pady=(20, 4))

        tk.Label(
            win, text="ENTER ID",
            font=("Segoe UI", 16, "bold"),
            bg=BG, fg="white"
        ).pack(pady=(0, 12))

        var = tk.StringVar()
        entry = tk.Entry(
            win, textvariable=var,
            font=("Segoe UI", 16, "bold"),
            justify="center", width=18,
            bg="#1e293b", fg="white",
            insertbackground="white", relief="flat"
        )
        entry.pack(ipady=6)
        entry.focus_set()

        tk.Label(
            win,
            text="Enter the exact signalling-gear ID",
            font=("Segoe UI", 9),
            bg=BG, fg="#475569"
        ).pack(pady=(5, 15))

        def confirm():
            value = var.get().strip().upper()
            if not value:
                return
            result["value"] = value
            win.grab_release()
            win.destroy()

        def cancel():
            result["value"] = None
            win.grab_release()
            win.destroy()

        tk.Button(
            win, text="CONFIRM", command=confirm,
            bg=BLUE, fg="white", relief="flat",
            font=("Segoe UI", 10, "bold"), width=14
        ).pack(pady=(0, 20))

        win.bind("<Return>", lambda e: confirm())
        win.bind("<Escape>", lambda e: cancel())
        center_window(win, 400, 240)
        win.wait_window()

        value = result["value"]
        if value is None:
            return None

        if value in existing:
            messagebox.showwarning(
                "DUPLICATE ID",
                f"'{value}' is already used."
            )
            continue

        return value


# =========================================================
# NEW SIGNAL - EXACT QUANTITY FLOW
# =========================================================

def build_quantity_capture_list():
    global signals

    signals = {}
    result = []
    existing = set()

    categories = [
        ("MAIN SIGNALS", "MAIN SIGNAL", "MAIN"),
        ("SHUNT SIGNALS", "SHUNT SIGNAL", "SHUNT"),
        ("CALLING-ON SIGNALS", "CALLING-ON SIGNAL", "CALLING_ON"),
        ("POINTS", "POINT", "POINT"),
        ("CRANK HANDLES", "CRANK HANDLE", "CH"),
        ("LC GATES", "LC GATE", "LC"),
    ]

    for count_label, name_label, typ in categories:
        count = ask_count(count_label)
        if count is None:
            return []

        for i in range(count):
            name = ask_name(
                name_label, i, count, existing
            )
            if name is None:
                return []

            existing.add(name)

            if typ == "MAIN":
                info = {
                    "type": "MAIN",
                    "aspects": None,
                    "menu": None,
                    "RED": None,
                    "YELLOW": None,
                    "DOUBLE_YELLOW": None,
                    "GREEN": None,
                    "ROUTE_INDICATOR": None,
                }

            elif typ == "SHUNT":
                # SHUNT capture flow:
                # 1. MENU
                # 2. ASPECT INDICATOR
                # 3. ROUTE INITIATION INDICATOR
                info = {
                    "type": "SHUNT",
                    "menu": None,
                    "indicator": None,
                    "route_init": None,
                }

            elif typ == "CALLING_ON":
                info = {
                    "type": "CALLING_ON",
                    "menu": None,
                    "YELLOW": None,
                    "ROUTE_INIT": None,
                }

            elif typ == "POINT":
                info = {
                    "type": "POINT",
                    "menu": None,
                    "normal": None,
                    "reverse": None,
                    "free": None,
                }

            elif typ == "CH":
                info = {
                    "type": "CH",
                    "menu": None,
                    "IN": None,
                    "OUT": None,
                    "ECH": None,
                    "FREE": None,
                }

            else:
                info = {
                    "type": "LC",
                    "menu": None,
                    "IN": None,
                    "OUT": None,
                }

            signals[name] = info
            result.append((typ, name))

    return result


# =========================================================
# CAPTURE OVERLAY
# =========================================================

def destroy_capture_overlay():
    global capture_overlay

    if capture_overlay:
        try:
            capture_overlay.destroy()
        except Exception:
            pass

    capture_overlay = None


def clear_overlay_buttons():
    if overlay_button_frame:
        for w in overlay_button_frame.winfo_children():
            w.destroy()


def create_capture_overlay(close_command):
    global capture_overlay
    global overlay_progress_label
    global overlay_signal_label
    global overlay_step_label
    global overlay_hint_label
    global overlay_button_frame

    destroy_capture_overlay()

    capture_overlay = tk.Toplevel(root)
    capture_overlay.title("Coordinate Capture Guide")
    capture_overlay.configure(bg=BG)
    capture_overlay.resizable(False, False)
    capture_overlay.attributes("-topmost", True)

    sw = capture_overlay.winfo_screenwidth()
    capture_overlay.geometry(
        f"420x330+{sw-440}+40"
    )
    capture_overlay.protocol(
        "WM_DELETE_WINDOW", close_command
    )

    capture_overlay.bind("<space>", lambda e: "break")
    capture_overlay.bind("<Return>", lambda e: "break")

    tk.Label(
        capture_overlay,
        text="COORDINATE CAPTURE",
        font=("Segoe UI", 11, "bold"),
        bg=BG, fg="#64748b"
    ).pack(pady=(16, 0))

    overlay_progress_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 10),
        bg=BG, fg="#94a3b8"
    )
    overlay_progress_label.pack(pady=(2, 10))

    overlay_signal_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 18, "bold"),
        bg=BG, fg="white"
    )
    overlay_signal_label.pack()

    overlay_step_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 15, "bold"),
        bg=BG, fg=GREEN
    )
    overlay_step_label.pack(pady=(6, 8))

    overlay_hint_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 10),
        bg=BG, fg="#cbd5e1",
        wraplength=380,
        justify="center"
    )
    overlay_hint_label.pack(pady=(0, 10))

    overlay_button_frame = tk.Frame(
        capture_overlay, bg=BG
    )
    overlay_button_frame.pack()

    tk.Button(
        capture_overlay,
        text="↶  UNDO LAST CLICK",
        command=undo_last_capture,
        bg="#f59e0b", fg="#111827",
        relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"),
        width=22
    ).pack(pady=(12, 2))

    tk.Label(
        capture_overlay,
        text="(or press BACKSPACE)",
        font=("Segoe UI", 8),
        bg=BG, fg="#475569"
    ).pack()

    tk.Button(
        capture_overlay,
        text="CANCEL CAPTURE",
        command=close_command,
        bg=BG, fg="#64748b",
        activeforeground="#ef4444",
        relief="flat",
        font=("Segoe UI", 9, "underline")
    ).pack(side="bottom", pady=10)


def update_overlay(typ, name, step, hint):
    if not capture_overlay:
        return

    clear_overlay_buttons()

    total = len(capture_master_list)

    overlay_progress_label.config(
        text=f"{TYPE_LABEL.get(typ, typ)}   |   "
             f"{capture_index + 1} OF {total}"
    )
    overlay_signal_label.config(
        text=name,
        fg=TYPE_COLOR.get(typ, "white")
    )
    overlay_step_label.config(text=step)
    overlay_hint_label.config(text=hint)

    capture_overlay.lift()


def show_aspect_buttons(signal):
    clear_overlay_buttons()

    overlay_progress_label.config(
        text=f"MAIN SIGNAL   |   {capture_index + 1} OF {len(capture_master_list)}"
    )
    overlay_signal_label.config(
        text=signal, fg=TYPE_COLOR["MAIN"]
    )
    overlay_step_label.config(
        text="HOW MANY ASPECTS?"
    )
    overlay_hint_label.config(
        text="Choose 2, 3 or 4 aspects."
    )

    for n in (2, 3, 4):
        tk.Button(
            overlay_button_frame,
            text=str(n),
            width=6,
            font=("Segoe UI", 12, "bold"),
            bg=BLUE, fg="white",
            relief="flat",
            command=lambda x=n:
                set_aspects_and_continue(signal, x)
        ).pack(side="left", padx=6)

    capture_overlay.lift()


# =========================================================
# SIGNAL CAPTURE STAGES
# =========================================================

def set_initial_stage():
    global record_stage

    typ, _ = capture_master_list[capture_index]

    # MAIN starts with its signal-menu capture.
    # SHUNT also starts with MENU, then ASPECT INDICATOR,
    # then ROUTE INITIATION INDICATOR.
    if typ == "MAIN":
        record_stage = "coordinate"
    elif typ == "SHUNT":
        record_stage = "menu"
    else:
        record_stage = "menu"


def push_undo(clear_fn):
    undo_stack.append({
        "index": capture_index,
        "stage": record_stage,
        "clear": clear_fn,
    })


def undo_last_capture():
    global capture_index, record_stage

    if capture_module != "MASTER":
        return

    if not undo_stack:
        log("Nothing to undo")
        return

    action = undo_stack.pop()

    try:
        action["clear"]()
    except Exception as e:
        log(f"UNDO ERROR : {e}")

    capture_index = action["index"]
    record_stage = action["stage"]
    master_next_capture()
    log("LAST SIGNAL CAPTURE UNDONE")


def cancel_signal_capture():
    global capture_module

    capture_module = None
    undo_stack.clear()
    destroy_capture_overlay()
    root.deiconify()
    root.lift()
    log("SIGNAL CAPTURE CANCELLED")


def advance_signal():
    global capture_index

    capture_index += 1

    if capture_index < len(capture_master_list):
        set_initial_stage()

    master_next_capture()


def master_next_capture():
    if capture_module != "MASTER":
        return

    if capture_index >= len(capture_master_list):
        finish_signal_capture()
        return

    typ, name = capture_master_list[capture_index]

    if typ == "MAIN":
        steps = {
            "coordinate": (
                "CLICK : SIGNAL MENU",
                "Move onto the signal menu button, then press SPACE."
            ),
            "RED": (
                "CLICK : RED ASPECT",
                "Move onto RED, then press SPACE."
            ),
            "YELLOW": (
                "CLICK : YELLOW ASPECT",
                "Move onto YELLOW, then press SPACE."
            ),
            "DOUBLE_YELLOW": (
                "CLICK : DOUBLE YELLOW",
                "Move onto DOUBLE YELLOW, then press SPACE."
            ),
            "GREEN": (
                "CLICK : GREEN ASPECT",
                "Move onto GREEN, then press SPACE."
            ),
            "ROUTE_INDICATOR": (
                "CLICK : ROUTE INITIATION INDICATOR",
                "Move onto route initiation indicator, then press SPACE."
            ),
        }

        if record_stage == "ask_aspects":
            show_aspect_buttons(name)
        else:
            step, hint = steps[record_stage]
            update_overlay("MAIN", name, step, hint)

    elif typ == "SHUNT":
        steps = {
            "menu": (
                "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            ),
            "indicator": (
                "CLICK: ASPECT INDICATOR",
                "Move the mouse onto the aspect indicator, then press SPACE.\n"
                "This also saves a reference snapshot for comparison."
            ),
            "route_init": (
                "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            ),
        }

        step, hint = steps[record_stage]
        update_overlay("SHUNT", name, step, hint)

    elif typ == "CALLING_ON":
        steps = {
            "menu": ("CLICK : SIGNAL MENU",
                     "Move onto the Calling-On signal menu, then press SPACE."),
            "YELLOW": ("CLICK : YELLOW",
                       "Move onto the yellow indicator, then press SPACE."),
            "ROUTE_INIT": ("CLICK : ROUTE INITIATION INDICATOR",
                           "Move onto the route initiation indicator, then press SPACE."),
        }
        step, hint = steps[record_stage]
        update_overlay("CALLING_ON", name, step, hint)

    elif typ == "POINT":
        steps = {
            "menu": ("CLICK : POINT MENU",
                     "Move onto the point menu, then press SPACE."),
            "normal": ("CLICK : NORMAL",
                       "Move onto NORMAL indication, then press SPACE."),
            "reverse": ("CLICK : REVERSE",
                        "Move onto REVERSE indication, then press SPACE."),
            "free": ("CLICK : FREE",
                     "Move onto FREE indication, then press SPACE."),
        }
        step, hint = steps[record_stage]
        update_overlay("POINT", name, step, hint)

    elif typ == "CH":
        steps = {
            "menu": ("CLICK : CH MENU",
                     "Move onto the crank-handle menu, then press SPACE."),
            "IN": ("CLICK : IN",
                   "Move onto IN, then press SPACE."),
            "OUT": ("CLICK : OUT",
                    "Move onto OUT, then press SPACE."),
            "ECH": ("CLICK : ECH",
                    "Move onto ECH, then press SPACE."),
            "FREE": ("CLICK : FREE",
                     "Move onto FREE, then press SPACE."),
        }
        step, hint = steps[record_stage]
        update_overlay("CH", name, step, hint)

    elif typ == "LC":
        steps = {
            "menu": ("CLICK : LC MENU",
                     "Move onto the LC gate menu, then press SPACE."),
            "IN": ("CLICK : IN",
                   "Move onto IN, then press SPACE."),
            "OUT": ("CLICK : OUT",
                    "Move onto OUT, then press SPACE."),
        }
        step, hint = steps[record_stage]
        update_overlay("LC", name, step, hint)


def master_save_point(x, y):
    global record_stage

    if capture_module != "MASTER":
        return

    typ, name = capture_master_list[capture_index]

    if typ == "MAIN":
        if record_stage == "coordinate":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("menu", None)
            )
            signals[name]["menu"] = [x, y]
            record_stage = "ask_aspects"
            master_next_capture()
            return

        if record_stage == "RED":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("RED", None)
            )
            signals[name]["RED"] = [x, y]
            record_stage = (
                "GREEN" if current_aspects == 2
                else "YELLOW"
            )
            master_next_capture()
            return

        if record_stage == "YELLOW":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("YELLOW", None)
            )
            signals[name]["YELLOW"] = [x, y]
            record_stage = (
                "DOUBLE_YELLOW"
                if current_aspects == 4
                else "GREEN"
            )
            master_next_capture()
            return

        if record_stage == "DOUBLE_YELLOW":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("DOUBLE_YELLOW", None)
            )
            signals[name]["DOUBLE_YELLOW"] = [x, y]
            record_stage = "GREEN"
            master_next_capture()
            return

        if record_stage == "GREEN":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("GREEN", None)
            )
            signals[name]["GREEN"] = [x, y]
            record_stage = "ROUTE_INDICATOR"
            master_next_capture()
            return

        if record_stage == "ROUTE_INDICATOR":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("ROUTE_INDICATOR", None)
            )
            signals[name]["ROUTE_INDICATOR"] = [x, y]
            advance_signal()
            return

    elif typ == "SHUNT":
        if record_stage == "menu":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("menu", None)
            )
            signals[name]["menu"] = [x, y]
            record_stage = "indicator"
            master_next_capture()
            return

        if record_stage == "indicator":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("indicator", None)
            )
            signals[name]["indicator"] = [x, y]
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("route_init", None)
            )
            signals[name]["route_init"] = [x, y]
            advance_signal()
            return

    elif typ == "CALLING_ON":
        if record_stage == "menu":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("menu", None)
            )
            signals[name]["menu"] = [x, y]
            record_stage = "YELLOW"
            master_next_capture()
            return

        if record_stage == "YELLOW":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("YELLOW", None)
            )
            signals[name]["YELLOW"] = [x, y]
            record_stage = "ROUTE_INIT"
            master_next_capture()
            return

        if record_stage == "ROUTE_INIT":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("ROUTE_INIT", None)
            )
            signals[name]["ROUTE_INIT"] = [x, y]
            advance_signal()
            return

    elif typ == "POINT":
        order = {
            "menu": "normal",
            "normal": "reverse",
            "reverse": "free",
        }

        if record_stage in order:
            key = record_stage
            push_undo(
                lambda n=name, k=key:
                signals[n].__setitem__(k, None)
            )
            signals[name][key] = [x, y]
            record_stage = order[key]
            master_next_capture()
            return

        if record_stage == "free":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("free", None)
            )
            signals[name]["free"] = [x, y]
            advance_signal()
            return

    elif typ == "CH":
        order = {
            "menu": "IN",
            "IN": "OUT",
            "OUT": "ECH",
            "ECH": "FREE",
        }

        if record_stage in order:
            key = record_stage
            push_undo(
                lambda n=name, k=key:
                signals[n].__setitem__(k, None)
            )
            signals[name][key] = [x, y]
            record_stage = order[key]
            master_next_capture()
            return

        if record_stage == "FREE":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("FREE", None)
            )
            signals[name]["FREE"] = [x, y]
            advance_signal()
            return

    elif typ == "LC":
        if record_stage == "menu":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("menu", None)
            )
            signals[name]["menu"] = [x, y]
            record_stage = "IN"
            master_next_capture()
            return

        if record_stage == "IN":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("IN", None)
            )
            signals[name]["IN"] = [x, y]
            record_stage = "OUT"
            master_next_capture()
            return

        if record_stage == "OUT":
            push_undo(
                lambda n=name:
                signals[n].__setitem__("OUT", None)
            )
            signals[name]["OUT"] = [x, y]
            advance_signal()
            return


def set_aspects_and_continue(signal, aspects):
    global current_aspects, record_stage

    current_aspects = aspects
    signals[signal]["aspects"] = aspects
    record_stage = "RED"
    log(f"{signal} -> {aspects} ASPECTS")
    master_next_capture()


def start_quantity_capture():
    global capture_module, capture_master_list, capture_index

    if signals:
        ok = messagebox.askyesno(
            "Start New Recording",
            "This will clear the existing signalling-gear capture.\n\n"
            "Continue?"
        )
        if not ok:
            return

    capture_master_list = build_quantity_capture_list()

    if not capture_master_list:
        log("NEW SIGNAL CANCELLED")
        return

    capture_index = 0
    capture_module = "MASTER"
    undo_stack.clear()
    set_initial_stage()

    create_capture_overlay(cancel_signal_capture)
    root.iconify()

    log("============================================================")
    log("NEW SIGNAL CAPTURE STARTED")
    log(f"TOTAL ELEMENTS : {len(capture_master_list)}")
    log("ORDER : MAIN -> SHUNT(MENU -> ASPECT INDICATOR -> ROUTE INITIATION) -> CALLING-ON -> POINTS -> CH -> LC")
    log("============================================================")

    master_next_capture()


def finish_signal_capture():
    global capture_module

    capture_module = None
    destroy_capture_overlay()

    root.deiconify()
    root.lift()

    save_signal_config()

    log("ALL SIGNALLING-GEAR COORDINATES CAPTURED")

    messagebox.showinfo(
        "Capture Complete",
        "New signal / complete signalling-gear capture finished.\n\n"
        "Configuration saved."
    )


# =========================================================
# SIGNAL CONFIG SAVE / LOAD
# =========================================================

def pair(value):
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return value[0], value[1]
    return None, None


def save_signal_config():
    if not signals:
        return

    path = filedialog.asksaveasfilename(
        title="Save Signal Configuration",
        initialfile="SIGNAL_CONFIG.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not path:
        return

    wb = Workbook()
    wb.remove(wb.active)

    sheets = {
        "MAIN_SIGNALS": [
            "Signal", "Aspects",
            "Menu_X", "Menu_Y",
            "Red_X", "Red_Y",
            "Yellow_X", "Yellow_Y",
            "DY_X", "DY_Y",
            "Green_X", "Green_Y",
            "RouteInit_X", "RouteInit_Y"
        ],
        "CALLING_ON": [
            "Signal",
            "Menu_X", "Menu_Y",
            "Yellow_X", "Yellow_Y",
            "RouteInit_X", "RouteInit_Y"
        ],
        "SHUNT": [
            "Signal",
            "Coordinate_X", "Coordinate_Y"
        ],
        "POINT": [
            "Signal",
            "Menu_X", "Menu_Y",
            "Normal_X", "Normal_Y",
            "Reverse_X", "Reverse_Y",
            "Free_X", "Free_Y"
        ],
        "CH": [
            "Signal",
            "Menu_X", "Menu_Y",
            "IN_X", "IN_Y",
            "OUT_X", "OUT_Y",
            "ECH_X", "ECH_Y",
            "FREE_X", "FREE_Y"
        ],
        "LC": [
            "Signal",
            "Menu_X", "Menu_Y",
            "IN_X", "IN_Y",
            "OUT_X", "OUT_Y"
        ],
    }

    ws = {}
    for name, headers in sheets.items():
        ws[name] = wb.create_sheet(name)
        ws[name].append(headers)

    for name, info in signals.items():
        typ = info["type"]

        if typ == "MAIN":
            row = [name, info.get("aspects")]
            for key in (
                "menu", "RED", "YELLOW",
                "DOUBLE_YELLOW", "GREEN",
                "ROUTE_INDICATOR"
            ):
                row.extend(pair(info.get(key)))
            ws["MAIN_SIGNALS"].append(row)

        elif typ == "CALLING_ON":
            row = [name]
            for key in ("menu", "YELLOW", "ROUTE_INIT"):
                row.extend(pair(info.get(key)))
            ws["CALLING_ON"].append(row)

        elif typ == "SHUNT":
            row = [name]
            for key in ("menu", "indicator", "route_init"):
                row.extend(pair(info.get(key)))
            ws["SHUNT"].append(row)

        elif typ == "POINT":
            row = [name]
            for key in ("menu", "normal", "reverse", "free"):
                row.extend(pair(info.get(key)))
            ws["POINT"].append(row)

        elif typ == "CH":
            row = [name]
            for key in ("menu", "IN", "OUT", "ECH", "FREE"):
                row.extend(pair(info.get(key)))
            ws["CH"].append(row)

        elif typ == "LC":
            row = [name]
            for key in ("menu", "IN", "OUT"):
                row.extend(pair(info.get(key)))
            ws["LC"].append(row)

    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.font = Font(
                bold=True, color="FFFFFF"
            )
            cell.fill = PatternFill(
                "solid", fgColor="1E293B"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

        for col in sheet.columns:
            letter = get_column_letter(
                col[0].column
            )
            width = max(
                len(str(c.value))
                if c.value is not None else 0
                for c in col
            )
            sheet.column_dimensions[
                letter
            ].width = min(width + 4, 30)

    wb.save(path)
    log(f"SIGNAL CONFIG SAVED : {path}")


def save_new_signal_coordinates():
    save_signal_config()


def sheet_name(sheetnames, *names):
    upper = {
        str(s).strip().upper(): s
        for s in sheetnames
    }
    for n in names:
        if n.upper() in upper:
            return upper[n.upper()]
    return None


def get_pair(row, x, y):
    try:
        if row[x] is None or row[y] is None:
            return None
        return [int(row[x]), int(row[y])]
    except Exception:
        return None


def load_config():
    global signals

    path = filedialog.askopenfilename(
        title="Select Signal Configuration",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not path:
        return

    try:
        wb = load_workbook(path, data_only=True)
        signals = {}

        s = sheet_name(
            wb.sheetnames,
            "MAIN_SIGNALS", "MAIN"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "MAIN",
                    "aspects": int(row[1] or 2),
                    "menu": get_pair(row, 2, 3),
                    "RED": get_pair(row, 4, 5),
                    "YELLOW": get_pair(row, 6, 7),
                    "DOUBLE_YELLOW": get_pair(row, 8, 9),
                    "GREEN": get_pair(row, 10, 11),
                    "ROUTE_INDICATOR": get_pair(row, 12, 13),
                }

        s = sheet_name(
            wb.sheetnames,
            "CALLING_ON", "CAL"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "CALLING_ON",
                    "menu": get_pair(row, 1, 2),
                    "YELLOW": get_pair(row, 3, 4),
                    "ROUTE_INIT": get_pair(row, 5, 6),
                }

        s = sheet_name(
            wb.sheetnames,
            "SHUNT", "SHUNT_SIGNALS"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "SHUNT",
                    "menu": get_pair(row, 1, 2),
                    "indicator": get_pair(row, 3, 4),
                    "route_init": get_pair(row, 5, 6),
                }

        s = sheet_name(
            wb.sheetnames, "POINT", "POINTS"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "POINT",
                    "menu": get_pair(row, 1, 2),
                    "normal": get_pair(row, 3, 4),
                    "reverse": get_pair(row, 5, 6),
                    "free": get_pair(row, 7, 8),
                }

        s = sheet_name(
            wb.sheetnames, "CH", "CRANK_HANDLES"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "CH",
                    "menu": get_pair(row, 1, 2),
                    "IN": get_pair(row, 3, 4),
                    "OUT": get_pair(row, 5, 6),
                    "ECH": get_pair(row, 7, 8),
                    "FREE": get_pair(row, 9, 10),
                }

        s = sheet_name(
            wb.sheetnames, "LC", "LC_GATES"
        )
        if s:
            for row in wb[s].iter_rows(
                min_row=2, values_only=True
            ):
                if not row or row[0] is None:
                    continue
                n = str(row[0]).strip().upper()
                signals[n] = {
                    "type": "LC",
                    "menu": get_pair(row, 1, 2),
                    "IN": get_pair(row, 3, 4),
                    "OUT": get_pair(row, 5, 6),
                }

        wb.close()

        log(f"EXISTING SIGNAL CONFIG LOADED : {path}")
        log(f"TOTAL ELEMENTS : {len(signals)}")

    except Exception as e:
        messagebox.showerror(
            "CONFIG LOAD ERROR", str(e)
        )
        log(f"CONFIG LOAD ERROR : {e}")


# =========================================================
# TRACK COORDINATES
# =========================================================

def create_track_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "TRACK COORDINATES"

    ws.merge_cells("A1:C1")
    ws["A1"] = "TRACK COORDINATES"
    ws["A1"].font = Font(
        bold=True, size=16, color="FFFFFF"
    )
    ws["A1"].fill = PatternFill(
        "solid", fgColor="1E293B"
    )
    ws["A1"].alignment = Alignment(
        horizontal="center"
    )

    ws.append([])
    ws.append([
        "TRACK NAME",
        "X COORDINATE",
        "Y COORDINATE"
    ])

    thin = Side(style="thin", color="000000")
    border = Border(
        left=thin, right=thin,
        top=thin, bottom=thin
    )

    for c in ws[3]:
        c.font = Font(
            bold=True, color="FFFFFF"
        )
        c.fill = PatternFill(
            "solid", fgColor="2563EB"
        )
        c.border = border
        c.alignment = Alignment(
            horizontal="center"
        )

    wb.save(COORD_FILE)
    log(f"TRACK EXCEL CREATED : {COORD_FILE}")


def save_track_coordinate(name, x, y):
    if not os.path.exists(COORD_FILE):
        create_track_excel()

    wb = load_workbook(COORD_FILE)
    ws = wb.active

    target = None
    for r in range(4, ws.max_row + 1):
        value = ws.cell(r, 1).value
        if value and str(value).strip().upper() == name.upper():
            target = r
            break

    if target is None:
        target = ws.max_row + 1

    ws.cell(target, 1).value = name
    ws.cell(target, 2).value = int(x)
    ws.cell(target, 3).value = int(y)

    wb.save(COORD_FILE)


def load_track_coordinates():
    global track_points, point_points

    if not os.path.exists(COORD_FILE):
        messagebox.showerror(
            "TRACK COORDINATES",
            f"File not found:\n{COORD_FILE}"
        )
        return False

    try:
        wb = load_workbook(
            COORD_FILE, data_only=True
        )
        ws = wb.active

        track_points = {}
        point_points = {}

        for row in ws.iter_rows(
            min_row=2, values_only=True
        ):
            if not row or len(row) < 3:
                continue
            if row[0] is None:
                continue
            if row[1] is None or row[2] is None:
                continue

            name = str(
                row[0]
            ).strip().upper()

            try:
                p = [int(row[1]), int(row[2])]
            except Exception:
                continue

            if name.endswith("XT"):
                track_points[name] = p
            else:
                point_points[name] = p

        wb.close()

        log(
            f"TRACK COORDINATES LOADED : "
            f"{len(track_points)} tracks"
        )

        return True

    except Exception as e:
        log(f"TRACK LOAD ERROR : {e}")
        return False


def import_track_coordinates():
    """Import track names from an Excel file without assuming a fixed column.

    The track header may be in any column and the header row may be within
    the first 20 rows. If no recognised header is found, the first column is
    used as a compatibility fallback.
    """
    global track_capture_list

    path = filedialog.askopenfilename(
        title="Select Track Coordinate Excel",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not path:
        return

    try:
        wb = load_workbook(path, data_only=True)

        header_names = {
            "TRACK", "TRACK NAME", "TRACK_NAME", "TRACKS",
            "TRACK COORDINATES", "TRACK_COORDINATES",
            "CONTROLLED BY TRACKS", "CONTROLLED TRACKS",
            "REQUIRED TRACKS", "TRACK DATA"
        }

        result = []
        seen = set()
        found_header = False

        # Search all sheets and first 20 rows for a recognised track column.
        for ws in wb.worksheets:
            header_row = None
            track_col = None

            for r in range(1, min(20, ws.max_row) + 1):
                for c in range(1, ws.max_column + 1):
                    value = ws.cell(r, c).value
                    if value is None:
                        continue
                    header = str(value).strip().upper().replace("-", " ")
                    if header in header_names:
                        header_row = r
                        track_col = c
                        break
                if track_col is not None:
                    break

            if track_col is None:
                continue

            found_header = True

            for r in range(header_row + 1, ws.max_row + 1):
                value = ws.cell(r, track_col).value
                if value is None:
                    continue

                # A TOC cell can contain comma-separated controlled tracks.
                values = str(value).replace("\n", ",").split(",")
                for item in values:
                    name = str(item).replace(".0", "").strip().upper()
                    if not name or name in header_names:
                        continue
                    if name not in seen:
                        seen.add(name)
                        result.append(name)

        # Compatibility fallback: first column of the active sheet.
        if not found_header:
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if not row or row[0] is None:
                    continue
                for item in str(row[0]).replace("\n", ",").split(","):
                    name = str(item).replace(".0", "").strip().upper()
                    if not name or name in header_names:
                        continue
                    if name not in seen:
                        seen.add(name)
                        result.append(name)

        wb.close()

        # Remove obvious non-track labels accidentally picked from generic sheets.
        result = [
            x for x in result
            if x not in {"SIGNAL", "ROUTE", "POINTS", "POINT", "DATE", "TIME"}
        ]

        if not result:
            messagebox.showerror(
                "TRACK IMPORT",
                "No track names found in the selected Excel."
            )
            log("TRACK IMPORT ERROR : NO TRACK NAMES FOUND")
            return

        track_capture_list = result

        log(
            f"TRACK LIST IMPORTED : {len(result)}"
        )
        for track in result:
            log(f"TRACK READY : {track}")

        messagebox.showinfo(
            "TRACK LIST LOADED",
            f"{len(result)} tracks loaded.\n\n"
            "Click CAPTURE TRACK COORDINATES."
        )

    except Exception as e:
        messagebox.showerror(
            "TRACK IMPORT ERROR", str(e)
        )
        log(f"TRACK IMPORT ERROR : {e}")


# =========================================================
# TRACK CAPTURE
# =========================================================

def update_track_overlay():
    if not capture_overlay:
        return

    if track_capture_index >= len(track_capture_list):
        return

    clear_overlay_buttons()

    track = track_capture_list[
        track_capture_index
    ]

    overlay_progress_label.config(
        text=(
            f"TRACK {track_capture_index + 1} "
            f"OF {len(track_capture_list)}"
        )
    )
    overlay_signal_label.config(
        text=track, fg=PURPLE
    )
    overlay_step_label.config(
        text="CAPTURE COORDINATE"
    )
    overlay_hint_label.config(
        text=(
            "Move the mouse onto the track.\n\n"
            "Press SPACE to save.\n"
            "Press BACKSPACE to undo."
        )
    )

    capture_overlay.lift()


def capture_track_point(x, y):
    global track_capture_index

    if capture_module != "TRACK":
        return

    if track_capture_index >= len(
        track_capture_list
    ):
        return

    name = track_capture_list[
        track_capture_index
    ]

    old = track_points.get(name)

    track_undo_stack.append({
        "name": name,
        "old": old,
        "index": track_capture_index
    })

    track_points[name] = [
        int(x), int(y)
    ]

    save_track_coordinate(
        name, x, y
    )

    log(
        f"{name} SAVED : ({x},{y})"
    )

    track_capture_index += 1

    if track_capture_index >= len(
        track_capture_list
    ):
        finish_track_capture()
    else:
        update_track_overlay()


def undo_track():
    global track_capture_index

    if not track_undo_stack:
        log("Nothing to undo")
        return

    action = track_undo_stack.pop()
    name = action["name"]

    if action["old"] is None:
        track_points.pop(name, None)
    else:
        track_points[name] = list(
            action["old"]
        )

    track_capture_index = action["index"]
    update_track_overlay()

    log(
        f"TRACK UNDO : {name}"
    )


def cancel_track_capture():
    global capture_module

    capture_module = None
    track_undo_stack.clear()
    destroy_capture_overlay()
    root.deiconify()
    root.lift()

    log("TRACK CAPTURE CANCELLED")


def finish_track_capture():
    global capture_module

    capture_module = None
    destroy_capture_overlay()
    root.deiconify()
    root.lift()

    log(
        "ALL TRACK COORDINATES CAPTURED"
    )

    messagebox.showinfo(
        "Track Capture Complete",
        f"{len(track_capture_list)} "
        "track coordinates captured."
    )


def start_track_capture():
    global capture_module
    global track_capture_index
    global track_undo_stack

    if not track_capture_list:
        messagebox.showwarning(
            "TRACK CAPTURE",
            "Click IMPORT TRACK COORDINATES first."
        )
        return

    if not os.path.exists(COORD_FILE):
        create_track_excel()

    track_capture_index = 0
    track_undo_stack.clear()
    capture_module = "TRACK"

    create_capture_overlay(
        cancel_track_capture
    )

    root.iconify()

    log(
        f"TRACK CAPTURE STARTED : "
        f"{len(track_capture_list)} tracks"
    )

    update_track_overlay()


# =========================================================
# KEYBOARD
# =========================================================

def dispatch_save(x, y):
    if capture_module == "MASTER":
        master_save_point(x, y)
    elif capture_module == "TRACK":
        capture_track_point(x, y)


def on_press(key):
    global last_space_time
    global last_backspace_time
    global last_digit_time

    if key == pynput_keyboard.Key.space:
        now = time.time()
        if now - last_space_time < 0.35:
            return
        last_space_time = now

        if capture_module not in (
            "MASTER", "TRACK"
        ):
            return

        x, y = win32api.GetCursorPos()
        root.after(
            0,
            lambda: dispatch_save(x, y)
        )

    elif key == pynput_keyboard.Key.backspace:
        now = time.time()
        if now - last_backspace_time < 0.35:
            return
        last_backspace_time = now

        if capture_module == "MASTER":
            root.after(0, undo_last_capture)
        elif capture_module == "TRACK":
            root.after(0, undo_track)

    else:
        char = getattr(key, "char", None)

        if (
            char in ("2", "3", "4")
            and capture_module == "MASTER"
            and capture_index < len(
                capture_master_list
            )
        ):
            typ, name = capture_master_list[
                capture_index
            ]

            if (
                typ == "MAIN"
                and record_stage == "ask_aspects"
            ):
                now = time.time()
                if now - last_digit_time < 0.35:
                    return
                last_digit_time = now

                root.after(
                    0,
                    lambda n=int(char):
                    set_aspects_and_continue(
                        name, n
                    )
                )


# =========================================================
# TOC IMPORT
# =========================================================

def _norm_toc_header(value):
    if value is None:
        return ""
    return re.sub(
        r"[^A-Z0-9]+",
        " ",
        str(value).strip().upper()
    ).strip()


def _find_toc_columns(ws):
    signal_aliases = {
        "SIGNAL", "SIGNAL NAME", "SIGNAL ID", "SIGNAL NUMBER",
        "SIGNAL NO", "SIGNAL NO.", "FROM SIGNAL", "START SIGNAL"
    }
    route_aliases = {
        "ROUTE", "ROUTE NAME", "ROUTE ID", "ROUTE NUMBER",
        "ROUTE NO", "ROUTE NO.", "ROUTE CODE", "ROUTE DESCRIPTION"
    }
    track_aliases = {
        "CONTROLLED BY TRACKS", "CONTROLLED TRACKS", "CONTROL TRACKS",
        "REQUIRED TRACKS", "REQUIRED TRACK", "TRACKS REQUIRED",
        "TRACK REQUIRED", "TRACK DATA", "TRACKS", "TRACK",
        "TRACK CIRCUITS", "TRACK CIRCUIT", "ROUTE TRACKS",
        "ROUTE CONTROLLED TRACKS", "LOCK ROUTES", "LOCK ROUTE TRACKS"
    }
    point_aliases = {
        "POINTS", "POINT", "POINT DATA", "POINT STATE", "POINT STATES",
        "POINT SETTING", "POINT SETTINGS", "REQUIRED POINTS",
        "CONTROLLED POINTS", "ROUTE POINTS", "LOCKED POINTS"
    }

    best = None
    for r in range(1, min(20, ws.max_row) + 1):
        sig = route = track = point = None
        score = 0
        for c in range(1, ws.max_column + 1):
            h = _norm_toc_header(ws.cell(r, c).value)
            if not h:
                continue
            if h in signal_aliases:
                sig = c
                score += 100
            elif h in route_aliases:
                route = c
                score += 100
            elif h in track_aliases:
                track = c
                score += 50
            elif h in point_aliases:
                point = c
                score += 25
        if sig is not None and route is not None:
            candidate = (score, r, sig, route, track, point)
            if best is None or candidate[0] > best[0]:
                best = candidate

    return best


def _find_toc_extra_column(ws, header_row, aliases):
    """Find an optional TOC column such as APPROACH_TRACK."""
    aliases = {_norm_toc_header(x) for x in aliases}
    for c in range(1, ws.max_column + 1):
        if _norm_toc_header(ws.cell(header_row, c).value) in aliases:
            return c
    return None


def _split_track_names(value):
    """Return coordinate-level track names from TOC/bit-chart names.

    In this station the universal TOC/TrackIt uses the base track name
    (for example ``1CXT``), while the Test Panel bit chart displays the
    corresponding indication as ``1CXTPR``.  Coordinates are captured
    against the base TrackIt name, and select_tracks() already maps a base
    name to the Test Panel ``PR`` item.

    Therefore a trailing ``PR`` is removed here so the TOC approach-track
    value ``1CXTPR`` correctly resolves to the captured coordinate ``1CXT``.
    """
    if value is None:
        return set()

    result = set()

    for item in str(value).replace("\\n", ",").replace(";", ",").split(","):
        name = str(item).replace(".0", "").strip().upper()
        if not name:
            continue

        # Test Panel indication suffix: 1CXT -> 1CXTPR.
        # Store/use the TrackIt coordinate name: 1CXT.
        if name.endswith("PR") and name[:-2].endswith("XT"):
            name = name[:-2]

        result.add(name)

    return result


def load_lock_routes():
    global lock_routes_data, track_points

    path = filedialog.askopenfilename(
        title="Select TOC / RCC Excel",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not path:
        return

    try:
        wb = load_workbook(path, data_only=True)

        toc_name = None
        for s in wb.sheetnames:
            if str(s).strip().upper() == "TOC":
                toc_name = s
                break

        if toc_name is None:
            # If a TOC sheet is not named exactly TOC, use the best sheet.
            toc_name = wb.sheetnames[0]

        ws = wb[toc_name]
        columns = _find_toc_columns(ws)

        if columns is None:
            # Universal fallback for the common A/B/C/D TOC format.
            header_row = 1
            signal_col, route_col, track_col, point_col = 1, 2, 3, 4
        else:
            _, header_row, signal_col, route_col, track_col, point_col = columns

        # Universal TOC optional columns used by subsidiary signals.
        approach_col = _find_toc_extra_column(
            ws,
            header_row,
            {
                "APPROACH_TRACK", "APPROACH TRACK",
                "APPROACH", "APPROACH TRACK CIRCUIT"
            }
        )
        sequential_col = _find_toc_extra_column(
            ws,
            header_row,
            {
                "SEQUENTIAL_TRACKS", "SEQUENTIAL TRACKS",
                "SEQUENTIAL TRACK", "SEQUENTIAL"
            }
        )

        lock_routes_data = []
        track_points = {}

        for r in range(header_row + 1, ws.max_row + 1):
            signal_value = ws.cell(r, signal_col).value
            route_value = ws.cell(r, route_col).value

            if signal_value is None or route_value is None:
                continue

            signal = str(signal_value).replace(".0", "").strip().upper()
            route = str(route_value).replace(".0", "").strip().upper()

            if not signal or not route:
                continue

            tracks = ""
            if track_col is not None:
                value = ws.cell(r, track_col).value
                if value is not None:
                    tracks = str(value).replace("\n", "").strip().upper()

            points = {}
            if point_col is not None:
                value = ws.cell(r, point_col).value
                if value:
                    for item in str(value).replace("\n", ",").upper().split(","):
                        item = item.strip()
                        if len(item) >= 2:
                            points[item[:-1].strip()] = item[-1].strip()

            approach_track = ""
            if approach_col is not None:
                value = ws.cell(r, approach_col).value
                if value is not None:
                    approach_track = str(value).replace("\n", "").strip().upper()

            sequential_tracks = ""
            if sequential_col is not None:
                value = ws.cell(r, sequential_col).value
                if value is not None:
                    sequential_tracks = str(value).replace("\n", "").strip().upper()

            lock_routes_data.append({
                "signal": signal,
                "route": route,
                "track_data": tracks,
                "point_data": points,
                "approach_track": approach_track,
                "sequential_tracks": sequential_tracks
            })

            for track in tracks.split(","):
                track = track.strip().upper()
                if track and not track.endswith("_VPR"):
                    track_points.setdefault(track, ())

            # Calling-on approach tracks must also be available to the
            # negative-test engine.  Their screenshot coordinates are loaded
            # from the same universal track-coordinate file later.
            if approach_track:
                for track in _split_track_names(approach_track):
                    if track and not track.endswith("_VPR"):
                        track_points.setdefault(track, ())

        wb.close()

        if not lock_routes_data:
            raise ValueError(
                "No valid SIGNAL/ROUTE rows were found in the TOC."
            )

        refresh_table()

        log(
            f"TOC LOADED : {len(lock_routes_data)} routes"
        )
        log(
            f"TOC COLUMNS : SIGNAL={signal_col}, ROUTE={route_col}, "
            f"TRACKS={track_col}, POINTS={point_col}, "
            f"APPROACH={approach_col}, SEQUENTIAL={sequential_col}"
        )

    except Exception as e:
        messagebox.showerror(
            "TOC LOAD ERROR", str(e)
        )
        log(f"TOC LOAD ERROR : {e}")


# =========================================================
# TABLE
# =========================================================

def refresh_table():
    tree.delete(*tree.get_children())

    for i, row in enumerate(
        lock_routes_data, 1
    ):
        tree.insert(
            "",
            "end",
            values=(
                i,
                row["signal"],
                row["route"],
                row["track_data"]
            )
        )


# =========================================================
# TEST PANEL HELPERS
# =========================================================

def click_menu_item(name):
    try:
        item = auto.MenuItemControl(
            searchDepth=15,
            Name=name
        )

        if item.Exists(3):
            pause_event.wait()
            item.Click(
                simulateMove=False
            )
            time.sleep(0.5)
            log(f"MENU CLICKED : {name}")
            return True

    except Exception as e:
        log(f"MENU CLICK ERROR : {e}")

    return False


def open_indications():
    panel = auto.WindowControl(
        Name="Test Panel"
    )

    if not panel.Exists(5):
        raise Exception(
            "Test Panel not found"
        )

    panel.SetActive()
    pyautogui.hotkey(
        "ctrl", "b"
    )
    time.sleep(1)

    station_window = panel.WindowControl(
        searchDepth=10,
        Name="Stations - Select Station"
    )

    # Window may open as a separate window on the other screen
    if not station_window.Exists(5):
        station_window = auto.WindowControl(
            searchDepth=1,
            Name="Stations - Select Station"
        )

    if not station_window.Exists(3):
        raise Exception(
            "Station selection window not found"
        )

    list_ctrl = station_window.ListControl()

    station = list_ctrl.ListItemControl(
        Name="50051"
    )

    if not station.Exists(5):
        raise Exception(
            "Station 50051 not found"
        )

    station.Click()
    time.sleep(0.4)

    station_window.ButtonControl(
        Name="OK"
    ).Click()

    time.sleep(2)

    panel = auto.WindowControl(
        Name="Test Panel"
    )

    ind = panel.WindowControl(
        searchDepth=10,
        Name="Station 50051: (Indications | Controls)"
    )

    # Station Bit window may open on the SECOND screen as a
    # separate top-level window - look there too.
    if not ind.Exists(5):
        ind = auto.WindowControl(
            searchDepth=1,
            Name="Station 50051: (Indications | Controls)"
        )

    if not ind.Exists(3):
        raise Exception(
            "Indications window not found"
        )

    ind.SetActive()
    time.sleep(0.5)

    r = ind.BoundingRectangle
    log(
        f"STATION BIT WINDOW AT L={r.left} T={r.top} "
        f"R={r.right} B={r.bottom}"
    )

    return ind


def get_track_color(item):
    try:
        # Bring the bit into view (long lists scroll)
        try:
            item.GetScrollItemPattern().ScrollIntoView()
            time.sleep(0.2)
        except Exception:
            pass

        r = item.BoundingRectangle

        # all_screens=True -> works when the Station Bit window
        # opens on the SECOND screen as well as the first.
        # (Without it only the primary screen is grabbed, colour
        # comes as UNKNOWN and no track gets selected.)
        img = ImageGrab.grab(
            bbox=(
                int(r.left),
                int(r.top),
                int(r.right),
                int(r.bottom)
            ),
            all_screens=True
        ).convert("RGB")

        red_pixels = 0
        blue_pixels = 0

        for y in range(img.height):
            for x in range(img.width):
                rr, gg, bb = img.getpixel(
                    (x, y)
                )

                if (
                    rr > 170
                    and gg < 120
                    and bb < 120
                ):
                    red_pixels += 1

                elif (
                    bb > 170
                    and rr < 120
                    and gg < 170
                ):
                    blue_pixels += 1

        if (
            red_pixels > blue_pixels
            and red_pixels > 10
        ):
            return "RED"

        if (
            blue_pixels > red_pixels
            and blue_pixels > 10
        ):
            return "BLUE"

    except Exception:
        pass

    return "UNKNOWN"


def select_tracks(track_names, mode, include_vpr=False):
    """
    Select/restore normal track indications and, when requested,
    all *_VPR tracks in the bit chart.

    mode:
        break   -> BLUE -> select -> Transmit
        restore -> RED  -> select -> Transmit

    include_vpr=True is intentionally used for the "break ALL"
    operation so VPR tracks are also broken.
    """
    try:
        names = {
            str(x).strip().upper()
            for x in track_names
            if str(x).strip()
        }

        ind = open_indications()
        track_list = ind.ListControl()

        if not track_list.Exists(5):
            raise Exception("Track list not found")

        selected_names = set()
        selected = 0

        # -------------------------------------------------
        # NORMAL TRACKS
        # -------------------------------------------------
        for name in sorted(names):
            item = track_list.ListItemControl(
                Name=f"{name}PR"
            )

            if not item.Exists(2):
                item = track_list.ListItemControl(
                    Name=name
                )

            if not item.Exists(2):
                log(
                    f"{name} NOT FOUND IN BIT CHART"
                )
                continue

            color = get_track_color(item)
            should_select = (
                mode == "break" and color == "BLUE"
            ) or (
                mode == "restore" and color == "RED"
            )

            log(
                f"{name} -> {color} "
                f"(TARGET={mode.upper()})"
            )

            if should_select:
                item.Click()
                selected += 1
                selected_names.add(name)
                log(
                    f"{name} -> {mode.upper()}"
                )

        # -------------------------------------------------
        # VPR TRACKS
        # -------------------------------------------------
        vpr_found = 0
        vpr_selected = 0

        if include_vpr:
            try:
                children = track_list.GetChildren()
            except Exception:
                children = []

            for child in children:
                try:
                    child_name = (
                        str(child.Name).strip().upper()
                    )

                    if not child_name.endswith("_VPR"):
                        continue

                    vpr_found += 1

                    color = get_track_color(child)

                    should_select = (
                        mode == "break"
                        and color == "BLUE"
                    ) or (
                        mode == "restore"
                        and color == "RED"
                    )

                    log(
                        f"{child_name} -> {color} "
                        f"(VPR TARGET={mode.upper()})"
                    )

                    if should_select:
                        child.Click()
                        selected += 1
                        vpr_selected += 1
                        log(
                            f"{child_name} -> "
                            f"{mode.upper()}"
                        )

                except Exception:
                    pass

        log(
            f"VPR TRACKS FOUND : {vpr_found} | "
            f"VPR SELECTED : {vpr_selected}"
        )

        # -------------------------------------------------
        # TRANSMIT ONCE
        # -------------------------------------------------
        if selected:
            transmit = ind.ButtonControl(
                Name="Transmit"
            )

            if not transmit.Exists(3):
                raise Exception(
                    "Transmit button not found"
                )

            transmit.Click()
            time.sleep(1)

        log(
            f"{selected} TRACKS TRANSMITTED "
            f"(MODE={mode.upper()})"
        )

        # -------------------------------------------------
        # CLOSE BIT CHART
        # -------------------------------------------------
        try:
            close = (
                ind.TitleBarControl()
                .ButtonControl(Name="Close")
            )

            if close.Exists(2):
                close.Click()

        except Exception:
            pass

        time.sleep(0.5)
        return True

    except Exception as e:
        log(
            f"SELECT TRACKS FAILED : {e}"
        )

        try:
            close = (
                ind.TitleBarControl()
                .ButtonControl(Name="Close")
            )
            if close.Exists(1):
                close.Click()
        except Exception:
            pass

        return False


def break_all_tracks(all_tracks):
    """
    Reset the station to the negative-test starting state:
    break every imported/required track AND every *_VPR track.
    """
    log("============================================================")
    log("BREAKING ALL TRACKS + ALL VPR TRACKS")
    log(f"TRACKS TO BREAK : {len(all_tracks)}")
    log("============================================================")

    ok = select_tracks(
        sorted(all_tracks),
        "break",
        include_vpr=True
    )

    time.sleep(2)

    if ok:
        log(
            "ALL TRACKS + VPR BREAK OPERATION COMPLETE"
        )
    else:
        log(
            "ALL TRACKS + VPR BREAK OPERATION FAILED"
        )

    return ok


# =========================================================
# TRACK INDICATION DETECTION
# =========================================================

def avg_color(x, y, img):
    pixels = []

    for dx in range(-2, 3):
        for dy in range(-2, 3):
            try:
                pixels.append(
                    img.getpixel(
                        (x + dx, y + dy)
                    )
                )
            except Exception:
                pass

    if not pixels:
        return 0, 0, 0

    return (
        sum(p[0] for p in pixels) // len(pixels),
        sum(p[1] for p in pixels) // len(pixels),
        sum(p[2] for p in pixels) // len(pixels),
    )


def is_track_active(point):
    if not point:
        return False

    x, y = map(int, point)

    for _ in range(10):

        img = pyautogui.screenshot()

        # Check a small area around the captured coordinate
        yellow_found = False
        best_rgb = (0, 0, 0)
        best_score = 0

        for dy in range(-5, 6):
            for dx in range(-5, 6):

                px = x + dx
                py = y + dy

                r, g, b = img.getpixel((px, py))

                # Yellow / active track detection
                if r > 180 and g > 180 and b < 120:

                    score = r + g

                    if score > best_score:
                        best_score = score
                        best_rgb = (r, g, b)

                    yellow_found = True

        log(
            f"TRACK RGB = {best_rgb[0]},{best_rgb[1]},{best_rgb[2]}"
        )

        if yellow_found:
            return True

        time.sleep(0.2)

    return False

# =========================================================
# CREATE NEGATIVE TESTING REPORT
# =========================================================

def create_report():

    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from datetime import datetime

    # =====================================================
    # CREATE WORKBOOK
    # =====================================================

    wb = Workbook()

    ws = wb.active

    ws.title = "NEGATIVE TESTING FOR TRACKS REPORT"

    # =====================================================
    # BORDER
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

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    ws["A1"] = "NEGATIVE TESTING FOR TRACKS REPORT"

    ws["A1"].font = Font(
        bold=True,
        size=18,
        color="000000"
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 30

    # =====================================================
    # GENERATED TIME
    # =====================================================

    ws.merge_cells("A2:D2")

    ws["A2"] = (
        "Generated : "
        + datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    )

    ws["A2"].font = Font(
        size=11,
        color="000000"
    )

    ws["A2"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # =====================================================
    # HEADERS
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
            bold=True,
            size=11,
            color="000000"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border

    # =====================================================
    # RESULTS
    # =====================================================

    row = 4

    initiated_routes = []
    non_initiated_routes = []

    # =====================================================
    # USE ACTUAL VARIABLE FROM YOUR PROGRAM
    # =====================================================

    for result in report_rows:

        signal = str(
            result.get("SIGNAL", "")
        )

        route = str(
            result.get("ROUTE", "")
        )

        overall_result = str(
            result.get("RESULT", "FAIL")
        ).upper()

        test_time = result.get(
            "DATE & TIME",
            datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        )

        # =================================================
        # INITIATED / NON-INITIATED
        # =================================================

        if overall_result == "PASS":

            initiated_routes.append(signal)

        else:

            non_initiated_routes.append(signal)

        # =================================================
        # WRITE DATA
        # =================================================

        ws.cell(
            row=row,
            column=1,
            value=signal
        )

        ws.cell(
            row=row,
            column=2,
            value=route
        )

        ws.cell(
            row=row,
            column=3,
            value=overall_result
        )

        ws.cell(
            row=row,
            column=4,
            value=test_time
        )

        # =================================================
        # NORMAL FORMATTING
        # =================================================

        for col in range(1, 5):

            cell = ws.cell(
                row=row,
                column=col
            )

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            cell.font = Font(
                size=11,
                color="000000"
            )

        # =================================================
        # ONLY FAILED ROW = RED
        # =================================================

        if overall_result == "FAIL":

            for col in range(1, 5):

                cell = ws.cell(
                    row=row,
                    column=col
                )

                cell.font = Font(
                    bold=True,
                    size=11,
                    color="FF0000"
                )

        row += 1

    # =====================================================
    # SUMMARY
    # =====================================================

    summary_row = row + 2

    total_routes = len(report_rows)

    # =====================================================
    # TOTAL ROUTES
    # =====================================================

    ws.cell(
        row=summary_row,
        column=1,
        value="TOTAL ROUTES"
    )

    ws.cell(
        row=summary_row,
        column=2,
        value=total_routes
    )

    # =====================================================
    # INITIATED ROUTES
    # =====================================================

    ws.cell(
        row=summary_row + 1,
        column=1,
        value="INITIATED ROUTES"
    )

    ws.cell(
        row=summary_row + 1,
        column=2,
        value=(
            ",".join(initiated_routes)
            if initiated_routes
            else "0"
        )
    )

    # =====================================================
    # NON-INITIATED ROUTES
    # =====================================================

    ws.cell(
        row=summary_row + 2,
        column=1,
        value="NON-INITIATED ROUTES"
    )

    ws.cell(
        row=summary_row + 2,
        column=2,
        value=(
            ",".join(non_initiated_routes)
            if non_initiated_routes
            else "0"
        )
    )

    # =====================================================
    # SUMMARY FORMATTING
    # =====================================================

    for r in range(
        summary_row,
        summary_row + 3
    ):

        for c in range(1, 3):

            cell = ws.cell(
                row=r,
                column=c
            )

            cell.border = border

            cell.font = Font(
                bold=True,
                size=11,
                color="000000"
            )

            cell.alignment = Alignment(
                horizontal=(
                    "left"
                    if c == 1
                    else "center"
                ),
                vertical="center"
            )

    # =====================================================
    # COLUMN WIDTHS
    # =====================================================

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 25

    # =====================================================
    # ROW HEIGHTS
    # =====================================================

    for r in range(3, row):

        ws.row_dimensions[r].height = 22

    # =====================================================
    # SAVE REPORT
    # =====================================================

    report_path = REPORT_FILE

    # Make sure the destination folder exists
    os.makedirs(os.path.dirname(report_path), exist_ok=True)

    wb.save(report_path)

    log(
        f"REPORT SAVED : {report_path}"
    )
# =========================================================
# NEGATIVE TEST ENGINE
# =========================================================

def cancel_signal_and_release(signal, menu_point):
    """Cancel the current route and release it before the next test."""
    cancel_ok = False
    release_ok = False

    log("STEP 5 : SIGNAL CANCEL")

    if click(menu_point):
        time.sleep(1)
        cancel_ok = click_menu_item("Signal Cancel")
        if cancel_ok:
            log(f"{signal} SIGNAL CANCEL DONE")
        else:
            log(f"SIGNAL CANCEL MENU ITEM NOT FOUND : {signal}")
    else:
        log(f"SIGNAL CANCEL MENU CLICK FAILED : {signal}")

    time.sleep(2)

    log("STEP 6 : ROUTE RELEASE")

    if click(menu_point):
        time.sleep(1)
        release_ok = click_menu_item("Route Release")
        if release_ok:
            log(f"{signal} ROUTE RELEASE DONE")
        else:
            log(f"ROUTE RELEASE MENU ITEM NOT FOUND : {signal}")
    else:
        log(f"ROUTE RELEASE MENU CLICK FAILED : {signal}")

    time.sleep(5)
    return cancel_ok and release_ok


def _find_temporary_point_tracks(tracks, point_data):
    """Return the opposite point track(s) for the route."""
    temporary_tracks = set()

    for point in point_data:
        point_name = str(point).strip().upper()

        if point_name == "50":
            if "50AXT" in tracks and "50BXT" not in tracks:
                temporary_tracks.add("50BXT")
            elif "50BXT" in tracks and "50AXT" not in tracks:
                temporary_tracks.add("50AXT")

        elif point_name == "65":
            if "65AXT" in tracks and "65BXT" not in tracks:
                temporary_tracks.add("65BXT")
            elif "65BXT" in tracks and "65AXT" not in tracks:
                temporary_tracks.add("65AXT")

    return temporary_tracks


def _set_route_from_menu(signal, route):
    """Use the existing TL-style signal menu route selection."""
    info = signals.get(signal, {})
    menu_point = info.get("menu")

    if not click(menu_point):
        log(f"MENU CLICK FAILED : {signal}")
        return False

    time.sleep(1)

    route_item = route.replace("-", "_")

    selected = click_menu_item(route_item)
    if not selected:
        selected = click_menu_item(route)

    if selected:
        log(f"ROUTE SELECTED : {signal} -> {route}")
        time.sleep(2)
        return True

    log(f"ROUTE NOT SELECTED : {signal} -> {route}")
    return False


def _initiate_subsidiary_route(signal):
    """
    Calling-on and shunt signals have a dedicated route-initiation
    coordinate captured in SIGNAL_CONFIG.

    This is intentionally performed only for subsidiary signal types;
    main-signal routes keep the already-working main flow.
    """
    info = signals.get(signal, {})
    typ = info.get("type")

    if typ == "CALLING_ON":
        point = info.get("ROUTE_INIT")
    elif typ == "SHUNT":
        point = info.get("route_init")
    else:
        return True

    if not point:
        log(f"{signal} ROUTE INITIATION COORDINATE NOT CAPTURED")
        return False

    log(f"{signal} ROUTE INITIATION : CLICK")
    ok = click(point)
    time.sleep(2)

    if ok:
        log(f"{signal} ROUTE INITIATION DONE")
    else:
        log(f"{signal} ROUTE INITIATION FAILED")

    return ok


def _verify_tracks(tracks):
    """Verify every required controlled track using the existing RGB logic."""
    overall_result = "PASS"

    for track in sorted(tracks):
        pause_event.wait()

        if not running:
            return "STOPPED"

        if track not in track_points or not track_points[track]:
            result = "FAIL"
            log(f"{track} -> FAIL (COORDINATE NOT FOUND)")
        else:
            active = is_track_active(track_points[track])
            result = "PASS" if active else "FAIL"
            log(f"{track} -> {result}")

        if result == "FAIL":
            overall_result = "FAIL"

    return overall_result


def _run_engine_with_com():
    """Run the automation engine on a worker thread with COM initialized.

    uiautomation uses Windows COM/UI Automation APIs.  Each worker thread
    must initialize COM before touching those APIs.
    """
    import ctypes
    ole32 = ctypes.windll.ole32
    hr = ole32.CoInitialize(None)
    try:
        run_engine()
    finally:
        # S_OK (0) and S_FALSE (1) both mean this thread entered COM.
        if hr >= 0:
            ole32.CoUninitialize()


def run_engine():
    """
    NEGATIVE TEST FLOW - MAIN + CALLING-ON + SHUNT

    MAIN:
      Break all -> restore route tracks + opposite point track ->
      set route -> break opposite point track -> verify route tracks ->
      Signal Cancel -> Route Release -> reset all.

    CALLING-ON:
      Break all -> restore the calling-on APPROACH TRACK
      (for example 1C_XTPR) + controlled route tracks +
      opposite point track -> set route -> initiate Calling-On using
      the captured ROUTE_INIT coordinate -> break the opposite point
      track -> verify controlled route tracks -> Signal Cancel ->
      Route Release -> reset all.

      The approach track is kept restored/occupied during the route test,
      which is the calling-on condition from the universal TOC.

    SHUNT:
      Break all -> restore controlled route tracks + opposite point
      track -> set route -> initiate the shunt route using the captured
      ROUTE_INIT coordinate -> break the opposite point track ->
      verify controlled route tracks -> Signal Cancel -> Route Release ->
      reset all.

    The existing main-signal negative-test behaviour is preserved.
    """

    global running, report_rows

    report_rows = []

    try:
        root.iconify()
        time.sleep(2)

        all_tracks = {
            str(t).strip().upper()
            for t in track_points.keys()
            if str(t).strip()
        }

        if not all_tracks:
            raise Exception("No track coordinates loaded.")

        # =====================================================
        # STEP 0 - INITIAL NEGATIVE STATE
        # =====================================================
        log("============================================================")
        log("NEGATIVE TESTING STARTED")
        log("MAIN + CALLING-ON + SHUNT")
        log("STEP 0 : BREAK ALL TRACKS INCLUDING VPR")
        log("============================================================")

        if not break_all_tracks(all_tracks):
            raise Exception(
                "Unable to perform initial ALL TRACK + VPR break."
            )

        # =====================================================
        # ROUTE-BY-ROUTE TESTING
        # =====================================================
        for route_index, row in enumerate(lock_routes_data, start=1):
            pause_event.wait()

            if not running:
                return

            signal = (
                str(row["signal"])
                .replace(".0", "")
                .strip()
                .upper()
            )
            route = (
                str(row["route"])
                .replace(".0", "")
                .strip()
                .upper()
            )

            tracks = {
                x.strip().upper()
                for x in str(row["track_data"]).split(",")
                if x.strip()
            }

            point_data = row.get("point_data", {}) or {}
            approach_tracks = _split_track_names(
                row.get("approach_track", "")
            )

            if signal not in signals:
                log(f"{signal} NOT IN CONFIG -> SKIPPING")
                continue

            signal_type = signals[signal].get("type", "MAIN")

            if signal_type not in ("MAIN", "CALLING_ON", "SHUNT"):
                log(
                    f"SKIPPING {signal} -> {route} "
                    f"(TYPE={signal_type})"
                )
                continue

            log("")
            log("============================================================")
            log(
                f"NEGATIVE TEST {route_index}/{len(lock_routes_data)} : "
                f"{signal} -> {route} [{signal_type}]"
            )
            log("============================================================")

            temporary_tracks = _find_temporary_point_tracks(
                tracks, point_data
            )

            log(f"SIGNAL TYPE : {signal_type}")
            log(f"REQUIRED TRACKS : {sorted(tracks)}")
            log(f"REQUIRED POINTS : {point_data}")
            log(
                f"NON-REQUIRED POINT TRACKS : "
                f"{sorted(temporary_tracks)}"
            )

            if signal_type == "CALLING_ON":
                log(
                    f"CALLING-ON APPROACH TRACK : "
                    f"{sorted(approach_tracks)}"
                )

                if not approach_tracks:
                    log(
                        f"WARNING : NO APPROACH TRACK IN TOC FOR "
                        f"CALLING-ON {signal} -> {route}"
                    )

            # -------------------------------------------------
            # STEP 1 - RESTORE REQUIRED CONDITIONS
            # -------------------------------------------------
            # IMPORTANT CALLING-ON NEGATIVE TEST CONDITION:
            #
            # The Calling-On approach track MUST remain BROKEN.
            #
            # Example:
            #   TOC APPROACH_TRACK = 1CXTPR
            #   TrackIt / coordinate name = 1CXT
            #
            # 1CXT must NOT be restored.
            #
            # When 1CXT is healthy:
            #   1C_A / 1C_B are disabled.
            #
            # When 1CXT is broken:
            #   1C_A / 1C_B become enabled.
            #
            # Therefore only the controlled route tracks and
            # temporary point-side tracks are restored.
            restore_tracks = set(tracks) | set(temporary_tracks)

            if signal_type == "CALLING_ON":
                log(
                    "CALLING-ON CONDITION : APPROACH TRACK "
                    "WILL REMAIN BROKEN"
                )
                log(
                    f"CALLING-ON APPROACH TRACK(S) KEPT BROKEN : "
                    f"{sorted(approach_tracks)}"
                )

            log("STEP 1 : RESTORE REQUIRED TRACKS / CONDITIONS")
            log(f"RESTORING : {sorted(restore_tracks)}")

            missing_restore_coords = {
                t for t in restore_tracks
                if t not in track_points or not track_points[t]
            }

            if missing_restore_coords:
                log(
                    "RESTORE FAILED - MISSING COORDINATES : "
                    f"{sorted(missing_restore_coords)}"
                )

                report_rows.append({
                    "SIGNAL": signal,
                    "ROUTE": route,
                    "TRACKS": ",".join(sorted(tracks)),
                    "RESULT": "FAIL",
                    "DATE & TIME": datetime.now().strftime(
                        "%d-%m-%Y %H:%M:%S"
                    ),
                })

                break_all_tracks(all_tracks)
                time.sleep(2)
                continue

            if restore_tracks:
                if not select_tracks(
                    sorted(restore_tracks),
                    "restore"
                ):
                    log(f"RESTORE FAILED : {signal} -> {route}")
                    report_rows.append({
                        "SIGNAL": signal,
                        "ROUTE": route,
                        "TRACKS": ",".join(sorted(tracks)),
                        "RESULT": "FAIL",
                        "DATE & TIME": datetime.now().strftime(
                            "%d-%m-%Y %H:%M:%S"
                        ),
                    })
                    break_all_tracks(all_tracks)
                    time.sleep(2)
                    continue

            time.sleep(2)

            # -------------------------------------------------
            # STEP 2 - SET ROUTE
            # -------------------------------------------------
            log("STEP 2 : SETTING ROUTE")

            if not _set_route_from_menu(signal, route):
                report_rows.append({
                    "SIGNAL": signal,
                    "ROUTE": route,
                    "TRACKS": ",".join(sorted(tracks)),
                    "RESULT": "FAIL",
                    "DATE & TIME": datetime.now().strftime(
                        "%d-%m-%Y %H:%M:%S"
                    ),
                })
                break_all_tracks(all_tracks)
                time.sleep(2)
                continue

            # -------------------------------------------------
            # STEP 3 - CALLING-ON / SHUNT ROUTE INITIATION
            # -------------------------------------------------
            route_init_ok = True

            if signal_type in ("CALLING_ON", "SHUNT"):
                log(
                    f"STEP 3 : {TYPE_LABEL.get(signal_type, signal_type)} "
                    "ROUTE INITIATION"
                )

                route_init_ok = _initiate_subsidiary_route(signal)

                if not route_init_ok:
                    log(
                        f"{signal} SUBSIDIARY ROUTE INITIATION FAILED"
                    )

            else:
                log("STEP 3 : MAIN SIGNAL ROUTE ALREADY INITIATED")

            # -------------------------------------------------
            # STEP 4 - BREAK NON-REQUIRED POINT TRACK
            # -------------------------------------------------
            # ALL SIGNAL TYPES (MAIN / CALLING-ON / SHUNT):
            # tracks stay as restored in STEP 1. No point track is
            # broken after setting the route; all tracks go down
            # only after Signal Cancel + Route Release (STEP 7).
            log(
                "STEP 4 : NO POINT TRACK BREAK "
                "(TRACKS KEPT AS RESTORED)"
            )

            # -------------------------------------------------
            # STEP 5 - NEGATIVE RESULT
            # -------------------------------------------------
            log(
                "STEP 5 : VERIFYING / REPORTING ROUTE CONDITION"
            )

            overall_result = "PASS"

            if not route_init_ok:
                overall_result = "FAIL"

            # For calling-on, the approach track is a condition track,
            # not a controlled route track, so it is not included in
            # the final controlled-track PASS/FAIL row.
            track_result = _verify_tracks(tracks)

            if track_result != "PASS":
                overall_result = "FAIL"

            if track_result == "STOPPED":
                return

            test_time = datetime.now().strftime(
                "%d-%m-%Y %H:%M:%S"
            )

            report_rows.append({
                "SIGNAL": signal,
                "ROUTE": route,
                "TRACKS": ",".join(sorted(tracks)),
                "RESULT": overall_result,
                "DATE & TIME": test_time,
            })

            log(
                f"NEGATIVE TEST RESULT : "
                f"{signal} -> {route} -> {overall_result}"
            )

            # -------------------------------------------------
            # STEP 6 - SIGNAL CANCEL + ROUTE RELEASE
            # -------------------------------------------------
            menu_point = signals[signal].get("menu")

            cleanup_ok = cancel_signal_and_release(
                signal,
                menu_point
            )

            if not cleanup_ok:
                log(
                    f"WARNING : SIGNAL CANCEL / ROUTE RELEASE "
                    f"NOT FULLY CONFIRMED : {signal}"
                )

            # -------------------------------------------------
            # STEP 7 - PREPARE FOR NEXT ROUTE / FINAL RESTORE
            # -------------------------------------------------
            # IMPORTANT:
            # For negative testing, all tracks must be broken again
            # only when another route still has to be tested.
            #
            # After the LAST route, DO NOT leave the station with
            # all tracks broken. Restore ALL tracks, including ALL
            # *_VPR tracks, so the panel is clean for the next task.
            if route_index < len(lock_routes_data):
                log(
                    "STEP 7 : BREAK ALL TRACKS INCLUDING VPR "
                    "FOR NEXT ROUTE"
                )

                if not break_all_tracks(all_tracks):
                    raise Exception(
                        f"Failed to reset all tracks after route {route}."
                    )

                time.sleep(2)
            else:
                log(
                    "STEP 7 : LAST ROUTE COMPLETED - "
                    "SKIPPING FINAL ALL-TRACK BREAK"
                )

        # =====================================================
        # FINAL RESTORE - ALL TRACKS + ALL VPR TRACKS
        # =====================================================
        log("============================================================")
        log("FINAL RESTORE : ALL TRACKS + ALL VPR TRACKS")
        log("The station will be left in the restored/healthy state.")
        log("============================================================")

        restore_ok = select_tracks(
            sorted(all_tracks),
            "restore",
            include_vpr=True
        )

        time.sleep(3)

        if restore_ok:
            log(
                "FINAL RESTORE COMPLETE : "
                "ALL TRACKS + ALL VPR TRACKS RETRIEVED"
            )
        else:
            log(
                "FINAL RESTORE FAILED : "
                "CHECK TEST PANEL BEFORE NEXT AUTOMATION"
            )

        # =====================================================
        # COMPLETE
        # =====================================================
        running = False

        status_label.config(
            text="COMPLETED",
            fg=GREEN
        )

        create_report()

        root.deiconify()

        log("============================================================")
        log("NEGATIVE TESTING COMPLETED")
        log("MAIN + CALLING-ON + SHUNT")
        log("============================================================")

        try:
            os.startfile(REPORT_FILE)
        except Exception:
            pass

    except Exception as e:
        running = False
        root.deiconify()

        status_label.config(
            text="ERROR",
            fg=RED
        )

        log(f"NEGATIVE TEST ERROR : {e}")

        try:
            create_report()
        except Exception:
            pass

        messagebox.showerror(
            "AUTOMATION ERROR",
            str(e)
        )


# =========================================================
# START / STOP / PAUSE
# =========================================================

def start_automation():

    global running
    global capture_waiting
    global capture_module
    global track_capture_mode

    # ============================================
    # RESET CAPTURE FLAGS
    # ============================================

    capture_waiting = False
    capture_module = None
    track_capture_mode = False

    log("============================================================")
    log("START AUTOMATION CHECK")
    log("============================================================")

    # ============================================
    # LOAD TRACK COORDINATES
    # ============================================

    log(f"COORD_FILE = {COORD_FILE}")
    log(f"COORD_FILE EXISTS = {os.path.exists(COORD_FILE)}")

    try:

        load_track_coordinates()

        log(
            f"TRACK COORDINATES COUNT = "
            f"{len(track_points)}"
        )

        if track_points:
            log(
                f"TRACK NAMES = "
                f"{list(track_points.keys())}"
            )

    except Exception as e:

        log(
            f"TRACK COORDINATE LOAD ERROR : {e}"
        )

        messagebox.showerror(
            "ERROR",
            f"Unable to load Track Coordinates.\n\n{e}"
        )

        return

    # ============================================
    # CHECK SIGNAL CONFIG
    # ============================================

    log(
        f"SIGNAL CONFIG COUNT = "
        f"{len(signals)}"
    )

    if signals:
        log(
            f"SIGNALS = "
            f"{list(signals.keys())}"
        )
    else:
        log(
            "SIGNAL CONFIG IS EMPTY"
        )

    if not signals:

        log(
            "START FAILED : LOAD CONFIG FIRST"
        )

        return

    # ============================================
    # CHECK LOCK ROUTES
    # ============================================

    log(
        f"LOCK ROUTE COUNT = "
        f"{len(lock_routes_data)}"
    )

    if lock_routes_data:

        for row in lock_routes_data[:10]:

            log(
                f"ROUTE : "
                f"{row.get('signal')} -> "
                f"{row.get('route')}"
            )

    else:

        log(
            "LOCK ROUTES ARE EMPTY"
        )

    if not lock_routes_data:

        log(
            "START FAILED : LOAD LOCK ROUTES FIRST"
        )

        return

    # ============================================
    # FINAL VALIDATION
    # ============================================

    if not track_points:

        log(
            "START FAILED : NO TRACK COORDINATES"
        )

        return

    # ============================================
    # EVERYTHING IS READY
    # ============================================

    log("============================================================")
    log("ALL DATA VALIDATED")
    log(
        f"TRACKS  = {len(track_points)}"
    )
    log(
        f"SIGNALS = {len(signals)}"
    )
    log(
        f"ROUTES  = {len(lock_routes_data)}"
    )
    log("============================================================")

    running = True

    pause_event.set()

    status_label.config(
        text="RUNNING",
        fg="#16a34a"
    )

    log("RUNNING = TRUE")
    log("STARTING RUN ENGINE")

    threading.Thread(
        target=_run_engine_with_com,
        daemon=True
    ).start()


def stop_automation():
    global running, paused

    running = False
    paused = False
    pause_event.set()

    status_label.config(
        text="STOPPED",
        fg=RED
    )

    root.deiconify()
    log("AUTOMATION STOPPED")

    if report_rows:
        create_report()


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
            fg=GREEN
        )
        log("AUTOMATION RESUMED")


# =========================================================
# GUI
# =========================================================

root = tk.Tk()
root.title(
    "NEGATIVE TESTING FOR VITAL INPUTS"
)
root.geometry("1450x900")
root.configure(bg="#e9edf2")
root.state("zoomed")


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


# =========================================================
# HEADER
# =========================================================

header = tk.Frame(
    root, bg=BG, height=100
)
header.pack(fill="x")
header.pack_propagate(False)

tk.Label(
    header,
    text="INDIAN RAILWAYS",
    font=("Segoe UI", 12, "bold"),
    bg=BG, fg="white"
).pack(
    side="left", padx=20
)

title_frame = tk.Frame(
    header, bg=BG
)
title_frame.pack(
    side="left", expand=True
)

tk.Label(
    title_frame,
    text="NEGATIVE TESTING FOR VITAL INPUTS",
    font=("Segoe UI", 24, "bold"),
    bg=BG, fg="white"
).pack(pady=(18, 0))

tk.Label(
    title_frame,
    text="NEGATIVE TESTING FOR TRACKS AUTOMATION",
    font=("Segoe UI", 10),
    bg=BG, fg="#cbd5e1"
).pack()

tk.Label(
    header,
    text="COMPANY LOGO",
    font=("Segoe UI", 12, "bold"),
    bg=BG, fg="white"
).pack(
    side="right", padx=20
)


# =========================================================
# MAIN AREA
# =========================================================

main_frame = tk.Frame(
    root, bg="#e9edf2"
)
main_frame.pack(
    fill="both", expand=True,
    padx=15, pady=15
)


# =========================================================
# LEFT CONTROL PANEL
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

tk.Label(
    left_panel,
    text="CONTROL PANEL",
    font=("Segoe UI", 15, "bold"),
    bg="white", fg="#0f172a"
).pack(pady=20)


def button(text, command, color):
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


signal_frame = tk.Frame(
    left_panel, bg="white"
)
signal_frame.pack(pady=10)

tk.Button(
    signal_frame,
    text="CAPTURE SIGNALLING GEARS",
    bg=BLUE, fg="white",
    activebackground=BLUE,
    relief="flat",
    font=("Segoe UI", 11, "bold"),
    width=26, height=2
).pack(pady=(0, 8))

small = tk.Frame(
    signal_frame, bg="white"
)
small.pack()

tk.Button(
    small,
    text="NEW SIGNAL",
    command=start_quantity_capture,
    bg=ORANGE, fg="white",
    activebackground=ORANGE,
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=15, height=1
).pack(
    side="left", padx=5
)

tk.Button(
    small,
    text="EXISTING SIGNAL",
    command=load_config,
    bg=ORANGE, fg="white",
    activebackground=ORANGE,
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=15, height=1
).pack(
    side="left", padx=5
)


button(
    "LOAD TOC",
    load_lock_routes,
    BLUE
).pack(pady=8)

button(
    "SAVE NEW SIGNAL COORDINATES",
    save_new_signal_coordinates,
    GREEN
).pack(pady=8)

button(
    "IMPORT TRACK COORDINATES",
    import_track_coordinates,
    BLUE
).pack(pady=8)

button(
    "CAPTURE TRACK COORDINATES",
    start_track_capture,
    PURPLE
).pack(pady=8)

button(
    "START TESTING",
    start_automation,
    GREEN
).pack(pady=20)

button(
    "STOP TESTING",
    stop_automation,
    RED
).pack(pady=8)


# =========================================================
# STATUS
# =========================================================

status_frame = tk.Frame(
    left_panel,
    bg="#f8fafc",
    bd=1,
    relief="solid"
)
status_frame.pack(
    fill="x",
    padx=15, pady=25
)

tk.Label(
    status_frame,
    text="SYSTEM STATUS",
    font=("Segoe UI", 11, "bold"),
    bg="#f8fafc", fg="#0f172a"
).pack(pady=10)

status_label = tk.Label(
    status_frame,
    text="READY",
    font=("Segoe UI", 16, "bold"),
    bg="#f8fafc",
    fg=GREEN
)
status_label.pack(pady=(0, 15))


# =========================================================
# RIGHT PANEL
# =========================================================

right_panel = tk.Frame(
    main_frame, bg="#e9edf2"
)
right_panel.pack(
    side="left",
    fill="both", expand=True
)

tk.Label(
    right_panel,
    text="TRACK DETAILS",
    font=("Segoe UI", 16, "bold"),
    bg="#e9edf2", fg="#0f172a"
).pack(
    anchor="w", pady=(0, 10)
)

table_frame = tk.Frame(
    right_panel,
    bg="white",
    bd=1,
    relief="solid"
)
table_frame.pack(
    fill="both", expand=True
)

scroll = ttk.Scrollbar(
    table_frame
)
scroll.pack(
    side="right", fill="y"
)

tree = ttk.Treeview(
    table_frame,
    columns=(
        "NO",
        "SIGNAL",
        "ROUTE",
        "CONTROLLED"
    ),
    show="headings",
    yscrollcommand=scroll.set
)

scroll.config(
    command=tree.yview
)

tree.heading(
    "NO", text="NO"
)
tree.heading(
    "SIGNAL", text="SIGNAL"
)
tree.heading(
    "ROUTE", text="ROUTE"
)
tree.heading(
    "CONTROLLED",
    text="CONTROLLED BY TRACKS"
)

tree.column(
    "NO", width=80,
    anchor="center"
)
tree.column(
    "SIGNAL", width=180,
    anchor="center"
)
tree.column(
    "ROUTE", width=220,
    anchor="center"
)
tree.column(
    "CONTROLLED", width=600,
    anchor="w"
)

tree.pack(
    fill="both", expand=True
)


# =========================================================
# LOG
# =========================================================

tk.Label(
    right_panel,
    text="LIVE OPERATION LOG",
    font=("Segoe UI", 16, "bold"),
    bg="#e9edf2", fg="#0f172a"
).pack(
    anchor="w", pady=(20, 10)
)

log_frame = tk.Frame(
    right_panel,
    bg="white",
    bd=1,
    relief="solid"
)
log_frame.pack(
    fill="both", expand=True
)

log_text = tk.Text(
    log_frame,
    bg="white",
    fg="black",
    font=("Consolas", 10),
    relief="flat"
)
log_text.pack(
    fill="both", expand=True,
    padx=10, pady=10
)
log_text.config(
    state="disabled"
)


# =========================================================
# FOOTER
# =========================================================

tk.Label(
    root,
    text=(
        "SPACE = CAPTURE | BACKSPACE = UNDO | "
        "NEW SIGNAL ORDER: MAIN → SHUNT(MENU → ASPECT → ROUTE INIT) → CALLING-ON → "
        "POINTS → CH → LC"
    ),
    bg=BG, fg="white",
    font=("Segoe UI", 10)
).pack(fill="x")


# =========================================================
# GLOBAL KEYBOARD LISTENER
# =========================================================

listener = pynput_keyboard.Listener(
    on_press=on_press
)
listener.daemon = True
listener.start()


def pause_key_listener(key):
    try:
        if key.char and key.char.lower() == "p":
            root.after(0, toggle_pause)
    except Exception:
        pass


pause_listener = pynput_keyboard.Listener(
    on_press=pause_key_listener
)
pause_listener.daemon = True
pause_listener.start()


# =========================================================
# RUN
# =========================================================

root.mainloop()
