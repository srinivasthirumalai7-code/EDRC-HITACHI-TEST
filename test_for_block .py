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
import os
import tkinter as tk
import keyboard
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
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "TRACK_FAILURE_REPORT.xlsx"

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

track_points = {}

running = False

capture_queue = []

capture_index = 0

capture_mode = None

capture_master_list = []

undo_stack = []

record_stage = None

current_aspects = 0

current_signal = None

current_step = 0

last_capture_time = 0

capture_waiting = False

captured_point = None

# Global keyboard listener for SPACE-based coordinate capture.
capture_module = None
track_capture_mode = False
capture_keyboard_listener = None

config_file_path = ""

# ==========================================
# NEW UI GLOBALS (TL STYLE)
# ==========================================

main_data = {}
cal_signal_data = {}
shunt_data = {}
point_data = {}
ch_data = {}
lc_data = {}

main_signal_order = []
cal_signal_order = []
shunt_signal_order = []
point_order = []
ch_order = []
lc_order = []

# =========================================================
# TRACK CAPTURE GLOBALS
# =========================================================

track_capture_list = []
track_capture_data = {}
current_track_index = 0
track_undo_stack = []

main_signal_aspects = {}

SNAPSHOTS_FOLDER = "snapshots"

os.makedirs(SNAPSHOTS_FOLDER, exist_ok=True)
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
# REPORT DATA STORAGE
# =========================================================

report_rows = []

# =========================================================
# TRACK STATE MANAGEMENT
# =========================================================

currently_broken_tracks = set()


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
        except Exception:
            pass

        capture_overlay = None


def clear_overlay_buttons():

    if overlay_button_frame is None:
        return

    for widget in overlay_button_frame.winfo_children():
        widget.destroy()


def update_capture_overlay(sig_type, signal, step_text, hint_text):

    # IMPORTANT:
    # The 2/3/4 aspect buttons belong ONLY to the current MAIN
    # signal while we are asking for its aspect count.
    #
    # After the aspect count is selected, the same overlay is
    # reused for RED/YELLOW/GREEN/ROUTE INIT and then for
    # CALLING-ON, SHUNT, POINT, CH, LC, etc.
    #
    # Therefore remove any old aspect buttons whenever a normal
    # capture step is displayed. Without this, the old 2/3/4
    # buttons remain visible and look as if every element is
    # asking for aspect count.
    clear_overlay_buttons()

    if sig_type == "TRACK":

        overlay_progress_label.config(
            text=f"Track {current_track_index + 1} of {len(track_capture_list)}   |   TRACK"
        )

        overlay_signal_label.config(
            text=signal,
            fg="#38bdf8"
        )

    else:

        overlay_progress_label.config(
            text=f"Signal {capture_index + 1} of {len(capture_master_list)}   |   {TYPE_LABELS[sig_type]}"
        )

        overlay_signal_label.config(
            text=signal,
            fg=TYPE_COLORS[sig_type]
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

    if capture_module == "TRACK":
        undo_last_track()
        return

    global capture_index, record_stage

    if not undo_stack:
        log("Nothing To Undo")
        return

    entry = undo_stack.pop()

    entry["clear"]()

    capture_index = entry["capture_index"]
    record_stage = entry["record_stage"]

    log("Last Capture Undone")

    master_next_capture()

def undo_last_track():

    global current_track_index

    if not track_undo_stack:
        log("Nothing To Undo")
        return

    last_track = track_undo_stack.pop()

    if last_track in track_capture_data:
        del track_capture_data[last_track]

    current_track_index -= 1

    if current_track_index < 0:
        current_track_index = 0

    log(f"{last_track} Undo")

    show_next_track()

def cancel_master_capture():

    global capture_module
    global capture_waiting
    global track_capture_mode

    capture_module = None
    capture_waiting = False
    track_capture_mode = False

    destroy_capture_overlay()

    root.deiconify()

    log("CAPTURE CANCELLED BY USER")

    undo_stack.clear()


def set_aspects_and_continue(signal, aspects):

    global record_stage, current_aspects

    push_undo(lambda s=signal: main_signal_aspects.pop(s, None))

    main_signal_aspects[signal] = aspects
    current_aspects = aspects
    record_stage = "RED"

    log(f"{signal} Aspects Set To {aspects}")

    master_next_capture()

    # SPACE is disabled while the 2/3/4 aspect buttons are shown.
    # Re-enable it after the aspect count is selected.
    global capture_waiting
    capture_waiting = True

def set_initial_stage_for_current():

    global record_stage

    sig_type, signal = capture_master_list[capture_index]

    # A new signal starts with its own capture flow.
    # Only MAIN starts with coordinate -> aspect-count.
    # CALLING-ON / SHUNT / POINT / CH / LC start directly
    # with their respective menu/element capture stage.
    record_stage = "coordinate" if sig_type == "MAIN" else "menu"

    # Never carry MAIN's 2/3/4 aspect buttons into another
    # signal/element.
    clear_overlay_buttons()

def advance_to_next_signal():

    global capture_index

    capture_index += 1

    if capture_index < len(capture_master_list):
        set_initial_stage_for_current()

    master_next_capture()

# =========================================================================
#  CAPTURE GUIDE POPUP
#  A small always-on-top window that tells the operator exactly what to
#  click next. Stays visible the whole time (even while the main window
#  is minimized), never blocks input, and needs no keyboard focus - the
#  operator just reads it, clicks the target on the real panel, and
#  presses SPACE (SPACE capture works globally, regardless of focus).
# =========================================================================

TYPE_LABELS = {
    "MAIN": "MAIN SIGNAL",
    "CAL": "CALLING-ON SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "POINT": "POINT MACHINE",
    "CH": "CRANK HANDLE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CAL": "#f97316",
    "SHUNT": "#a855f7",
    "POINT": "#14b8a6",
    "CH": "#eab308",
    "LC": "#f43f5e"
}

capture_overlay = None
overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None


def master_next_capture():

    global capture_module

    if capture_index >= len(capture_master_list):

        capture_module = None

        destroy_capture_overlay()

        root.deiconify()

        refresh_table()

        log("ALL COORDINATES CAPTURED")

        undo_stack.clear()

        messagebox.showinfo(
            "Capture Complete",
            f"All {len(capture_master_list)} elements captured successfully."
        )

        return

    sig_type, signal = capture_master_list[capture_index]

    log(f"TYPE = {sig_type}")
    log(f"SIGNAL = {signal}")
    log(f"STAGE = {record_stage}")

    print(sig_type, signal, record_stage)

    if sig_type == "MAIN":

        if record_stage == "coordinate":

            update_capture_overlay(
                sig_type, signal, "CLICK: SIGNAL MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "ask_aspects":

            show_aspect_buttons(signal)

        elif record_stage == "RED":

            update_capture_overlay(
                sig_type, signal, "CLICK: RED ASPECT",
                "Move the mouse onto the RED lamp, then press SPACE."
            )

        elif record_stage == "YELLOW":

            update_capture_overlay(
                sig_type, signal, "CLICK: YELLOW ASPECT",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )

        elif record_stage == "DOUBLE_YELLOW":

            update_capture_overlay(
                sig_type, signal, "CLICK: DOUBLE YELLOW ASPECT",
                "Move the mouse onto the DOUBLE YELLOW lamp, then press SPACE."
            )

        elif record_stage == "GREEN":

            update_capture_overlay(
                sig_type, signal, "CLICK: GREEN ASPECT",
                "Move the mouse onto the GREEN lamp, then press SPACE."
            )

        elif record_stage == "ROUTE_INIT":

            update_capture_overlay(
                sig_type, signal, "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    elif sig_type == "CAL":

        if record_stage == "menu":

            update_capture_overlay(
                sig_type, signal, "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "yellow":

            update_capture_overlay(
                sig_type, signal, "CLICK: YELLOW LAMP",
                "Move the mouse onto the YELLOW lamp, then press SPACE."
            )

        elif record_stage == "route_init":

            update_capture_overlay(
                sig_type, signal, "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    elif sig_type == "SHUNT":

        if record_stage == "menu":

            update_capture_overlay(
                sig_type, signal, "CLICK: MENU",
                "Move the mouse onto this signal's menu button, then press SPACE."
            )

        elif record_stage == "indicator":

            update_capture_overlay(
                sig_type, signal, "CLICK: ASPECT INDICATOR",
                "Move the mouse onto the aspect indicator, then press SPACE.\n"
                "This also saves a reference snapshot for comparison."
            )

        elif record_stage == "route_init":

            update_capture_overlay(
                sig_type, signal, "CLICK: ROUTE INITIATION INDICATOR",
                "Move the mouse onto the route initiation indicator, then press SPACE."
            )

    elif sig_type == "POINT":

        if record_stage == "menu":

            update_capture_overlay(
                sig_type, signal, "CLICK: MENU",
                "Move the mouse onto this point's menu button, then press SPACE."
            )

        elif record_stage == "normal":

            update_capture_overlay(
                sig_type, signal, "CLICK: NORMAL ASPECT",
                "Move the mouse onto the NORMAL aspect indicator, then press SPACE."
            )

        elif record_stage == "reverse":

            update_capture_overlay(
                sig_type, signal, "CLICK: REVERSE ASPECT",
                "Move the mouse onto the REVERSE aspect indicator, then press SPACE."
            )

        elif record_stage == "free":

            update_capture_overlay(
                sig_type, signal, "CLICK: FREE ASPECT",
                "Move the mouse onto the FREE (out of correspondence) indicator, then press SPACE."
            )

    elif sig_type == "CH":

        if record_stage == "menu":

            update_capture_overlay(
                sig_type, signal, "CLICK: MENU",
                "Move the mouse onto this crank handle's menu button, then press SPACE."
            )

        elif record_stage == "IN":

            update_capture_overlay(
                sig_type, signal, "CLICK: IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )

        elif record_stage == "OUT":

            update_capture_overlay(
                sig_type, signal, "CLICK: OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )

        elif record_stage == "ECH":

            update_capture_overlay(
                sig_type, signal, "CLICK: ECH ASPECT",
                "Move the mouse onto the ECH indicator, then press SPACE."
            )

        elif record_stage == "FREE":

            update_capture_overlay(
                sig_type, signal, "CLICK: FREE ASPECT",
                "Move the mouse onto the FREE indicator, then press SPACE."
            )

    elif sig_type == "LC":

        if record_stage == "menu":

            update_capture_overlay(
                sig_type, signal, "CLICK: MENU",
                "Move the mouse onto this LC gate's menu button, then press SPACE."
            )

        elif record_stage == "in":

            update_capture_overlay(
                sig_type, signal, "CLICK: IN ASPECT",
                "Move the mouse onto the IN indicator, then press SPACE."
            )

        elif record_stage == "out":

            update_capture_overlay(
                sig_type, signal, "CLICK: OUT ASPECT",
                "Move the mouse onto the OUT indicator, then press SPACE."
            )


def get_calling_on_signal_for(main_signal):
    """Finds the CALLING-ON signal linked to a MAIN signal (e.g. '1' -> '1C')."""

    candidate = (main_signal + "C").upper()

    for s in cal_signal_order:

        if s.upper() == candidate:
            return s

    return None


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
    """TOC-INDEPENDENT capture list. Asks HOW MANY of each element type,
    then a name for each one, in this fixed order:

        MAIN -> SHUNT -> CALLING-ON -> POINT -> CRANK HANDLE -> LC GATE

    Entering 0 (or Cancel) for any category skips it and moves straight
    to the next one. This is what "RECORD COMPLETE YARD" uses instead of
    reading names off an imported TOC - the coordinate file it produces
    is independent of any particular TOC and can be reused by any tool
    that needs coordinates for these signal/element names."""

    global main_signal_order, shunt_signal_order, cal_signal_order
    global point_order, ch_order, lc_order

    combined = []

    all_names = (
        set(main_signal_order) | set(shunt_signal_order) | set(cal_signal_order)
        | set(point_order) | set(ch_order) | set(lc_order)
    )
    log(f"ALL_NAMES BEFORE INPUT = {all_names}")

    # ---- MAIN ----
    n = prompt_count("MAIN signals")
    for i in range(n):
        name = prompt_name("MAIN Signal", i, n, all_names)
        all_names.add(name)
        if name not in main_data:
            main_data[name] = {
                "open_menu": None, "RED": None, "YELLOW": None,
                "DOUBLE_YELLOW": None, "GREEN": None, "ROUTE_INIT": None
            }
        if name not in main_signal_order:
            main_signal_order.append(name)
        combined.append(("MAIN", name))

    # ---- SHUNT ----
    n = prompt_count("SHUNT signals")
    for i in range(n):
        name = prompt_name("SHUNT Signal", i, n, all_names)
        all_names.add(name)
        if name not in shunt_data:
            shunt_data[name] = {
                "menu_coordinate": None, "state_indicator": None,
                "initial_snapshot": None, "route_init": None
            }
        if name not in shunt_signal_order:
            shunt_signal_order.append(name)
        combined.append(("SHUNT", name))

    # ---- CALLING-ON ----
    n = prompt_count("CALLING-ON signals")
    for i in range(n):
        name = prompt_name("CALLING-ON Signal", i, n, all_names)
        all_names.add(name)
        if name not in cal_signal_data:
            cal_signal_data[name] = {"menu": None, "yellow": None, "route_init": None}
        if name not in cal_signal_order:
            cal_signal_order.append(name)
        combined.append(("CAL", name))

    # ---- POINTS ----
    n = prompt_count("Points (point machines)")
    for i in range(n):
        name = prompt_name("Point", i, n, all_names)
        all_names.add(name)
        if name not in point_data:
            point_data[name] = {"menu": None, "normal": None, "reverse": None, "free": None}
        if name not in point_order:
            point_order.append(name)
        combined.append(("POINT", name))

    # ---- CRANK HANDLE ----
    n = prompt_count("Crank Handles")
    for i in range(n):
        name = prompt_name("Crank Handle", i, n, all_names)
        all_names.add(name)
        if name not in ch_data:
            ch_data[name] = {"menu": None, "IN": None, "OUT": None, "ECH": None, "FREE": None}
        if name not in ch_order:
            ch_order.append(name)
        combined.append(("CH", name))

    # ---- LC GATE ----
    n = prompt_count("LC Gates")
    for i in range(n):
        name = prompt_name("LC Gate", i, n, all_names)
        all_names.add(name)
        if name not in lc_data:
            lc_data[name] = {"menu": None, "in": None, "out": None}
        if name not in lc_order:
            lc_order.append(name)
        combined.append(("LC", name))

    return combined

def build_capture_master_list():
    """Builds one ordered list covering every signal: each MAIN signal is
    immediately followed by its CALLING-ON signal (if any), then all
    SHUNT signals are appended at the end."""

    combined = []
    used_cal = set()

    for signal in main_signal_order:

        combined.append(("MAIN", signal))

        cal_signal = get_calling_on_signal_for(signal)

        if cal_signal:
            combined.append(("CAL", cal_signal))
            used_cal.add(cal_signal)

    for s in cal_signal_order:

        if s not in used_cal:
            combined.append(("CAL", s))

    for s in shunt_signal_order:

        combined.append(("SHUNT", s))

    return combined


# =========================================================
# CREATE SIGNAL SETUP
# =========================================================

def create_setup():

    log("NEW SIGNAL BUTTON CLICKED")

    global signals
    global capture_master_list
    global capture_index
    global current_step

    main_data.clear()
    cal_signal_data.clear()
    shunt_data.clear()
    point_data.clear()
    ch_data.clear()
    lc_data.clear()

    main_signal_order.clear()
    cal_signal_order.clear()
    shunt_signal_order.clear()
    point_order.clear()
    ch_order.clear()
    lc_order.clear()

    signals = {}

    log(f"MAIN = {main_signal_order}")
    log(f"CALLING = {cal_signal_order}")
    log(f"SHUNT = {shunt_signal_order}")

    capture_master_list = build_quantity_capture_list()

    for sig in main_signal_order:

        if sig not in main_data:
            main_data[sig] = {
                "open_menu": None,
                "RED": None,
                "YELLOW": None,
                "DOUBLE_YELLOW": None,
                "GREEN": None,
                "ROUTE_INIT": None
            }

    for sig in cal_signal_order:

        if sig not in cal_signal_data:
            cal_signal_data[sig] = {
                "menu": None,
                "yellow": None,
                "route_init": None
            }

    for sig in shunt_signal_order:

        if sig not in shunt_data:
            shunt_data[sig] = {
                "menu_coordinate": None,
                "state_indicator": None,
                "initial_snapshot": None,
                "route_init": None
            }

    log(f"TOTAL ITEMS TO CAPTURE = {len(capture_master_list)}")

    if not capture_master_list:
        return

    capture_index = 0
    current_step = 0

    set_initial_stage_for_current()

    create_capture_overlay()

    global capture_module
    global capture_waiting

    capture_module = "MASTER"
    capture_waiting = True

    log(f"capture_module = {capture_module}")
    log(f"capture_waiting = {capture_waiting}")
    log(str(main_data))

    master_next_capture()


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
        destroy_capture_overlay()

        root.deiconify()

        save_config()

        undo_stack.clear()

        log("ALL SIGNALS RECORDED")

        messagebox.showinfo(
            "Capture Complete",
            "All coordinates captured successfully."
        )

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

        update_capture_overlay(

            "MAIN",

            current_signal,

            f"CLICK : {current}",

            "Move the mouse over the required location and press SPACE."

        )

    # =====================================================
    # CALLING ON
    # =====================================================

    elif sig_type == "CALLING_ON":

        steps = ["MENU", "YELLOW"]

        current = steps[current_step]

        update_capture_overlay(

            "CAL",

            current_signal,

            f"CLICK : {current}",

            "Move the mouse over the required location and press SPACE."

        )

    # =====================================================
    # SHUNT
    # =====================================================

    elif sig_type == "SHUNT":

        steps = ["MENU", "C1", "C2", "C3"]

        current = steps[current_step]

        update_capture_overlay(

            "SHUNT",

            current_signal,

            f"CLICK : {current}",

            "Move the mouse over the required location and press SPACE."

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

            steps = ["MENU", "RED", "GREEN"]

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

    push_undo(
        signal,
        current_name,
        signals[signal].get(current_name)
    )

    signals[signal][current_name] = [x, y]

    log(
        f"{signal} {current_name} Saved"
    )

    current_step += 1

    if current_step >= len(steps):

        current_step = 0

        capture_index += 1

    master_next_capture()


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

            push_undo(lambda s=signal: main_data[s].__setitem__("open_menu", None))

            main_data[signal]["open_menu"] = [x, y]
            log(f"{signal} Menu Saved ({x},{y})")
            record_stage = "ask_aspects"
            master_next_capture()
            return

        if record_stage == "ask_aspects":

            log(f"{signal} - SPACE does nothing here. Click 2 / 3 / 4 on the guide window, or press the 2, 3, or 4 key.")
            return

        if record_stage == "RED":

            push_undo(lambda s=signal: main_data[s].__setitem__("RED", None))

            main_data[signal]["RED"] = [x, y]
            log(f"{signal} RED Saved")
            record_stage = "GREEN" if current_aspects == 2 else "YELLOW"
            master_next_capture()
            return

        if record_stage == "YELLOW":

            push_undo(lambda s=signal: main_data[s].__setitem__("YELLOW", None))

            main_data[signal]["YELLOW"] = [x, y]
            log(f"{signal} YELLOW Saved")
            record_stage = "DOUBLE_YELLOW" if current_aspects == 4 else "GREEN"
            master_next_capture()
            return

        if record_stage == "DOUBLE_YELLOW":

            push_undo(lambda s=signal: main_data[s].__setitem__("DOUBLE_YELLOW", None))

            main_data[signal]["DOUBLE_YELLOW"] = [x, y]
            log(f"{signal} DOUBLE YELLOW Saved")
            record_stage = "GREEN"
            master_next_capture()
            return

        if record_stage == "GREEN":

            push_undo(lambda s=signal: main_data[s].__setitem__("GREEN", None))

            main_data[signal]["GREEN"] = [x, y]
            log(f"{signal} GREEN Saved")
            record_stage = "ROUTE_INIT"
            master_next_capture()
            return

        if record_stage == "ROUTE_INIT":

            push_undo(lambda s=signal: main_data[s].__setitem__("ROUTE_INIT", None))

            main_data[signal]["ROUTE_INIT"] = [x, y]
            log(f"{signal} Route Initiation Indicator Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # CALLING-ON
    # =====================================================
    elif sig_type == "CAL":

        if record_stage == "menu":

            push_undo(lambda s=signal: cal_signal_data[s].__setitem__("menu", None))

            cal_signal_data[signal]["menu"] = [x, y]
            log(f"{signal} Menu Coordinate Saved")
            record_stage = "yellow"
            master_next_capture()
            return

        if record_stage == "yellow":

            push_undo(lambda s=signal: cal_signal_data[s].__setitem__("yellow", None))

            cal_signal_data[signal]["yellow"] = [x, y]
            log(f"{signal} Yellow Lamp Saved")
            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":

            push_undo(lambda s=signal: cal_signal_data[s].__setitem__("route_init", None))

            cal_signal_data[signal]["route_init"] = [x, y]
            log(f"{signal} Route Initiation Indicator Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # SHUNT
    # =====================================================
    elif sig_type == "SHUNT":

        if record_stage == "menu":

            push_undo(lambda s=signal: shunt_data[s].__setitem__("menu_coordinate", None))

            shunt_data[signal]["menu_coordinate"] = [x, y]
            log(f"{signal} Menu Saved ({x},{y})")
            record_stage = "indicator"
            master_next_capture()
            return

        if record_stage == "indicator":

            screenshot = pyautogui.screenshot()
            region = screenshot.crop((x - 50, y - 50, x + 50, y + 50))

            snapshot_file = os.path.join(SNAPSHOTS_FOLDER, f"{signal}_initial.png")

            region.save(snapshot_file)

            def _undo_shunt_indicator(s=signal, f=snapshot_file):

                shunt_data[s]["state_indicator"] = None
                shunt_data[s]["initial_snapshot"] = None

                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass

            push_undo(_undo_shunt_indicator)

            shunt_data[signal]["state_indicator"] = [x, y]
            shunt_data[signal]["initial_snapshot"] = snapshot_file

            log(f"{signal} Indicator Saved ({x},{y})")
            log(f"{signal} Snapshot: {snapshot_file}")

            record_stage = "route_init"
            master_next_capture()
            return

        if record_stage == "route_init":

            push_undo(lambda s=signal: shunt_data[s].__setitem__("route_init", None))

            shunt_data[signal]["route_init"] = [x, y]
            log(f"{signal} Route Initiation Indicator Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # POINT
    # =====================================================
    elif sig_type == "POINT":

        if record_stage == "menu":

            push_undo(lambda s=signal: point_data[s].__setitem__("menu", None))

            point_data[signal]["menu"] = [x, y]
            log(f"{signal} Menu Saved ({x},{y})")
            record_stage = "normal"
            master_next_capture()
            return

        if record_stage == "normal":

            push_undo(lambda s=signal: point_data[s].__setitem__("normal", None))

            point_data[signal]["normal"] = [x, y]
            log(f"{signal} Normal Aspect Saved")
            record_stage = "reverse"
            master_next_capture()
            return

        if record_stage == "reverse":

            push_undo(lambda s=signal: point_data[s].__setitem__("reverse", None))

            point_data[signal]["reverse"] = [x, y]
            log(f"{signal} Reverse Aspect Saved")
            record_stage = "free"
            master_next_capture()
            return

        if record_stage == "free":

            push_undo(lambda s=signal: point_data[s].__setitem__("free", None))

            point_data[signal]["free"] = [x, y]
            log(f"{signal} Free Aspect Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # CRANK HANDLE
    # =====================================================
    elif sig_type == "CH":

        if record_stage == "menu":

            push_undo(lambda s=signal: ch_data[s].__setitem__("menu", None))

            ch_data[signal]["menu"] = [x, y]
            log(f"{signal} Menu Saved ({x},{y})")
            record_stage = "IN"
            master_next_capture()
            return

        if record_stage == "IN":

            push_undo(lambda s=signal: ch_data[s].__setitem__("IN", None))

            ch_data[signal]["IN"] = [x, y]
            log(f"{signal} IN Aspect Saved")
            record_stage = "OUT"
            master_next_capture()
            return

        if record_stage == "OUT":

            push_undo(lambda s=signal: ch_data[s].__setitem__("OUT", None))

            ch_data[signal]["OUT"] = [x, y]
            log(f"{signal} OUT Aspect Saved")
            record_stage = "ECH"
            master_next_capture()
            return

        if record_stage == "ECH":

            push_undo(lambda s=signal: ch_data[s].__setitem__("ECH", None))

            ch_data[signal]["ECH"] = [x, y]
            log(f"{signal} ECH Aspect Saved")
            record_stage = "FREE"
            master_next_capture()
            return

        if record_stage == "FREE":

            push_undo(lambda s=signal: ch_data[s].__setitem__("FREE", None))

            ch_data[signal]["FREE"] = [x, y]
            log(f"{signal} FREE Aspect Saved")
            advance_to_next_signal()
            return

    # =====================================================
    # LC GATE
    # =====================================================
    elif sig_type == "LC":

        if record_stage == "menu":

            push_undo(lambda s=signal: lc_data[s].__setitem__("menu", None))

            lc_data[signal]["menu"] = [x, y]
            log(f"{signal} Menu Saved ({x},{y})")
            record_stage = "in"
            master_next_capture()
            return

        if record_stage == "in":

            push_undo(lambda s=signal: lc_data[s].__setitem__("in", None))

            lc_data[signal]["in"] = [x, y]
            log(f"{signal} IN Aspect Saved")
            record_stage = "out"
            master_next_capture()
            return

        if record_stage == "out":

            push_undo(lambda s=signal: lc_data[s].__setitem__("out", None))

            lc_data[signal]["out"] = [x, y]
            log(f"{signal} OUT Aspect Saved")
            advance_to_next_signal()
            return





# =========================================================
# KEYBOARD
# =========================================================

def on_press(key):

    global capture_waiting
    global captured_point
    global capture_module

    # SPACE is the universal capture key. pynput receives it
    # even when the Test Panel or capture overlay has focus.
    if key != pynput_keyboard.Key.space:
        return

    if not capture_waiting:
        return

    # Consume this press immediately to prevent double capture.
    capture_waiting = False

    x, y = win32api.GetCursorPos()
    captured_point = (x, y)

    # pynput runs on a background thread. Dispatch the actual
    # capture to Tk's main thread.
    def dispatch_capture():

        global capture_waiting

        if capture_module == "MASTER":

            master_save_point(x, y)

            # Do not re-arm while the aspect-count buttons are open.
            if (
                capture_module == "MASTER"
                and capture_index < len(capture_master_list)
                and record_stage != "ask_aspects"
            ):
                capture_waiting = True

        elif capture_module == "TRACK":

            save_current_track(x, y)

            if track_capture_mode:
                capture_waiting = True

    try:
        root.after(0, dispatch_capture)
    except Exception:
        capture_waiting = False

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

        "Yellow_Y"

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

        "C1_X",

        "C1_Y",

        "C2_X",

        "C2_Y",

        "C3_X",

        "C3_Y"

    ])

    # =====================================================
    # WRITE MAIN
    # =====================================================

    for sig in main_signal_order:
        info = main_data[sig]

        ws_main.append([

            sig,
            main_signal_aspects.get(sig),

            *(info["open_menu"] or [None, None]),

            *(info["RED"] or [None, None]),

            *(info["YELLOW"] or [None, None]),

            *(info["DOUBLE_YELLOW"] or [None, None]),

            *(info["GREEN"] or [None, None]),

        ])

    # =====================================================
    # CALLING ON
    # =====================================================

    for sig in cal_signal_order:
        info = cal_signal_data[sig]

        ws_call.append([

            sig,

            *(info["menu"] or [None, None]),

            *(info["yellow"] or [None, None])

        ])

    # =====================================================
    # SHUNT
    # =====================================================

    for sig in shunt_signal_order:
        info = shunt_data[sig]

        ws_shunt.append([

            sig,

            *(info["menu_coordinate"] or [None, None]),

            *(info["state_indicator"] or [None, None]),

            *(info["route_init"] or [None, None])

        ])

    ws_point = wb.create_sheet("POINTS")

    ws_point.append([
        "Point",
        "Menu_X", "Menu_Y",
        "Normal_X", "Normal_Y",
        "Reverse_X", "Reverse_Y",
        "Free_X", "Free_Y"
    ])

    for sig in point_order:
        info = point_data[sig]

        ws_point.append([

            sig,

            *(info["menu"] or [None, None]),

            *(info["normal"] or [None, None]),

            *(info["reverse"] or [None, None]),

            *(info["free"] or [None, None])

        ])

    ws_ch = wb.create_sheet("CRANK_HANDLE")

    ws_ch.append([
        "CH",
        "Menu_X", "Menu_Y",
        "IN_X", "IN_Y",
        "OUT_X", "OUT_Y",
        "ECH_X", "ECH_Y",
        "FREE_X", "FREE_Y"
    ])

    for sig in ch_order:
        info = ch_data[sig]

        ws_ch.append([

            sig,

            *(info["menu"] or [None, None]),

            *(info["IN"] or [None, None]),

            *(info["OUT"] or [None, None]),

            *(info["ECH"] or [None, None]),

            *(info["FREE"] or [None, None])

        ])

    ws_lc = wb.create_sheet("LC_GATES")

    ws_lc.append([
        "LC",
        "Menu_X", "Menu_Y",
        "IN_X", "IN_Y",
        "OUT_X", "OUT_Y"
    ])

    for sig in lc_order:
        info = lc_data[sig]

        ws_lc.append([

            sig,

            *(info["menu"] or [None, None]),

            *(info["in"] or [None, None]),

            *(info["out"] or [None, None])

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

def save_track_coordinates():

    wb = Workbook()

    ws = wb.active

    ws.title = "TRACK_COORDINATES"

    ws.append(["Track", "X", "Y"])

    for track, coord in track_capture_data.items():

        ws.append([track, coord[0], coord[1]])

    save_path = filedialog.asksaveasfilename(
        title="Save Track Coordinates",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not save_path:
        return

    wb.save(save_path)

    log(f"TRACK COORDINATES SAVED : {save_path}")

    messagebox.showinfo(
        "SUCCESS",
        f"Track coordinates saved.\n\n{save_path}"
    )

# =========================================================
# IMPORT TRACK COORDINATES
# =========================================================

def import_track_coordinates():

    global track_coordinate_file
    global track_capture_list
    global current_track_index
    global track_capture_data

    file_path = filedialog.askopenfilename(
        title="Select Track Coordinate Excel",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file_path:
        return

    track_coordinate_file = file_path

    wb = load_workbook(file_path)

    ws = wb.active

    log(f"Globals contains track_capture_list = {'track_capture_list' in globals()}")
    log(f"Globals contains track_capture_data = {'track_capture_data' in globals()}")
    log(f"Globals contains current_track_index = {'current_track_index' in globals()}")

    track_capture_list.clear()
    track_capture_data.clear()

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row[0]:
            continue

        track = str(row[0]).strip().upper()

        track_capture_list.append(track)

        track_capture_data[track] = None

    current_track_index = 0

    log(f"TRACK EXCEL SELECTED : {track_coordinate_file}")
    log(f"TOTAL TRACKS TO CAPTURE : {len(track_capture_list)}")

# =========================================================
# LOAD CONFIG
# =========================================================

def load_config():

    global signals
    global config_file_path

    global main_data
    global cal_signal_data
    global shunt_data
    global point_data
    global ch_data
    global lc_data

    main_data = {}
    cal_signal_data = {}
    shunt_data = {}
    point_data = {}
    ch_data = {}
    lc_data = {}

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

    wb = load_workbook(config_file_path, data_only=True)

    # -----------------------------------------------------
    # Supports BOTH coordinate file formats:
    #   universal E2E file : MAIN / CAL / SHUNT (/POINT/CH/LC)
    #   old format         : MAIN_SIGNALS / CALLING_ON / SHUNT
    # -----------------------------------------------------
    sheet_lookup = {str(n).strip().upper(): n for n in wb.sheetnames}

    main_sheet = sheet_lookup.get("MAIN") or sheet_lookup.get("MAIN_SIGNALS")
    calling_sheet = sheet_lookup.get("CAL") or sheet_lookup.get("CALLING_ON")
    shunt_sheet = sheet_lookup.get("SHUNT")

    log(f"CONFIG SHEETS FOUND : {wb.sheetnames}")

    if not (main_sheet and calling_sheet and shunt_sheet):
        messagebox.showerror(
            "ERROR",
            "Signal configuration sheets not found.\n\n"
            "Required: MAIN + CAL + SHUNT\n"
            "or MAIN_SIGNALS + CALLING_ON + SHUNT"
        )
        log("SIGNAL CONFIGURATION SHEETS NOT FOUND")
        return

    log(f"MAIN : {main_sheet} | CALLING-ON : {calling_sheet} | SHUNT : {shunt_sheet}")

    # =====================================================
    # MAIN
    # =====================================================

    ws = wb[main_sheet]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        signals[signal_name] = {

            "type": "MAIN",

            "aspects": row[1] if len(row) > 1 else 1,

            "menu": [row[2], row[3]] if len(row) > 3 else [None, None],

            "RED": [row[4], row[5]] if len(row) > 5 else [None, None],

            "YELLOW": [row[6], row[7]]
            if len(row) > 7 and row[6] else None,

            "DOUBLE_YELLOW": [row[8], row[9]]
            if len(row) > 9 and row[8] else None,

            "GREEN": [row[10], row[11]] if len(row) > 11 else [None, None]

        }

        main_data[signal_name] = signals[signal_name]

    # =====================================================
    # CALLING ON
    # =====================================================

    ws = wb[calling_sheet]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        signals[signal_name] = {

            "type": "CALLING_ON",

            "menu": [row[1], row[2]] if len(row) > 2 else [None, None],

            "YELLOW": [row[3], row[4]] if len(row) > 4 else [None, None],

            "ROUTE_INIT": [row[5], row[6]]
            if len(row) > 6 and row[5] is not None else None

        }

        cal_signal_data[signal_name] = signals[signal_name]

    # =====================================================
    # SHUNT
    # =====================================================

    ws = wb[shunt_sheet]

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row[0]:
            continue

        signal_name = str(row[0]).replace(".0", "").strip().upper()

        signals[signal_name] = {

            "type": "SHUNT",

            "menu": [row[1], row[2]] if len(row) > 2 else [None, None],

            "route_init": [row[3], row[4]]
            if len(row) > 4 and row[3] is not None else None,

            "C1": [row[3], row[4]] if len(row) > 4 else None,

            "C2": [row[5], row[6]] if len(row) > 6 else None,

            "C3": [row[7], row[8]] if len(row) > 8 else None

        }

        shunt_data[signal_name] = signals[signal_name]

    print(signals.keys())
    log(f"AVAILABLE SIGNALS : {list(signals.keys())}")

    log("CONFIG LOADED")


def save_new_signal_coordinates():

    global config_file_path

    if (
            not main_data and
            not cal_signal_data and
            not shunt_data and
            not point_data and
            not ch_data and
            not lc_data
    ):
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

    tree.delete(*tree.get_children())

    count = 1

    for row in lock_routes_data:

        # Show only rows used for Block Test
        if not row["sequential_tracks"].strip():
            continue

        tree.insert(
            "",
            "end",
            values=(
                count,
                row["signal"],
                row["route"],
                row["track_data"],
                row["sequential_tracks"]
            )
        )

        count += 1

# =========================================================
# LOAD LOCK ROUTE EXCEL
# =========================================================

def load_lock_routes():

    global lock_routes_data
    global track_points

    file = filedialog.askopenfilename(
        title="Select TOC / RCC Excel",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    try:
        wb = load_workbook(file, data_only=True)
    except Exception as e:
        messagebox.showerror(
            "ERROR",
            f"Unable to open TOC file.\n\n{e}"
        )
        return

    # -----------------------------------------------------
    # Find TOC sheet
    # -----------------------------------------------------
    ws = None

    for sheet_name in wb.sheetnames:
        if sheet_name.strip().upper() == "TOC":
            ws = wb[sheet_name]
            break

    if ws is None:
        messagebox.showerror(
            "ERROR",
            "TOC SHEET NOT FOUND"
        )
        return

    # -----------------------------------------------------
    # READ COLUMN POSITIONS FROM HEADERS
    #
    # Your common TOC is not limited to four columns.
    # Example:
    # A = SIGNAL
    # B = ROUTE
    # C = CONTROLLED_BY_TRACKS
    # D = Points
    # E = Approach_Track
    # F = SEQUENTIAL_TRACKS
    #
    # Earlier code assumed SEQUENTIAL_TRACKS was column D,
    # which caused 50R, 65R, 65N, etc. to be treated as
    # sequential tracks.
    # -----------------------------------------------------
    header_map = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value

        if value is None:
            continue

        header = (
            str(value)
            .strip()
            .upper()
            .replace(" ", "_")
            .replace("-", "_")
        )

        header_map[header] = col

    def find_col(*names):
        for name in names:
            key = (
                name.strip()
                .upper()
                .replace(" ", "_")
                .replace("-", "_")
            )
            if key in header_map:
                return header_map[key]
        return None

    signal_col = find_col("SIGNAL")
    route_col = find_col("ROUTE")

    controlled_col = find_col(
        "CONTROLLED_BY_TRACKS",
        "CONTROLLED BY TRACKS",
        "CONTROLLED_TRACKS"
    )

    sequential_col = find_col(
        "SEQUENTIAL_TRACKS",
        "SEQUENTIAL TRACKS",
        "SEQUENTIAL"
    )

    points_col = find_col(
        "POINTS",
        "POINT"
    )

    approach_col = find_col(
        "APPROACH_TRACK",
        "APPROACH TRACK",
        "TRACK"                   # common TOC header
    )

    # -----------------------------------------------------
    # Fallbacks only for the mandatory fields.
    # This keeps compatibility with older 4-column TOCs.
    # -----------------------------------------------------
    if signal_col is None:
        signal_col = 1

    if route_col is None:
        route_col = 2

    if controlled_col is None:
        controlled_col = 3

    # IMPORTANT:
    # If SEQUENTIAL_TRACKS exists in the header, ALWAYS use
    # that exact column. Never fall back to column D when the
    # common 6-column TOC has a separate sequential column.
    if sequential_col is None:
        if ws.max_column >= 6:
            sequential_col = 6
        else:
            sequential_col = 4

    log(
        "TOC COLUMNS : "
        f"SIGNAL={signal_col}, "
        f"ROUTE={route_col}, "
        f"CONTROLLED_BY_TRACKS={controlled_col}, "
        f"POINTS={points_col or '-'}, "
        f"APPROACH_TRACK={approach_col or '-'}, "
        f"SEQUENTIAL_TRACKS={sequential_col}"
    )

    # -----------------------------------------------------
    # RESET
    # -----------------------------------------------------
    lock_routes_data = []
    track_points = {}

    main_signal_order.clear()
    cal_signal_order.clear()
    shunt_signal_order.clear()

    # -----------------------------------------------------
    # READ TOC
    # -----------------------------------------------------
    for row_num in range(2, ws.max_row + 1):

        signal_value = ws.cell(
            row=row_num,
            column=signal_col
        ).value

        route_value = ws.cell(
            row=row_num,
            column=route_col
        ).value

        # Automatic signals (e.g. 8, 25) have NO route in TOC -
        # keep them; only a row without a signal is skipped.
        if signal_value is None:
            continue

        if route_value is None:
            route_value = ""

        signal = (
            str(signal_value)
            .replace(".0", "")
            .strip()
            .upper()
        )

        route = (
            str(route_value)
            .replace(".0", "")
            .strip()
            .upper()
        )

        # -------------------------------------------------
        # CONTROLLED BY TRACKS
        # -------------------------------------------------
        controlled_value = ws.cell(
            row=row_num,
            column=controlled_col
        ).value

        track_data = ""

        if controlled_value is not None:
            track_data = (
                str(controlled_value)
                .replace(".0", "")
                .strip()
                .upper()
            )

        # -------------------------------------------------
        # SEQUENTIAL TRACKS
        # -------------------------------------------------
        sequential_value = ws.cell(
            row=row_num,
            column=sequential_col
        ).value

        sequential_tracks = ""

        if sequential_value is not None:
            sequential_tracks = (
                str(sequential_value)
                .replace(".0", "")
                .strip()
                .upper()
            )

        # -------------------------------------------------
        # Optional fields retained for future/common-TOC use
        # -------------------------------------------------
        point_value = (
            ws.cell(row=row_num, column=points_col).value
            if points_col
            else None
        )

        approach_value = (
            ws.cell(row=row_num, column=approach_col).value
            if approach_col
            else None
        )

        point_data = ""
        approach_track = ""

        if point_value is not None:
            point_data = str(point_value).strip().upper()

        if approach_value is not None:
            approach_track = str(approach_value).strip().upper()

        # -------------------------------------------------
        # Signal classification is kept for coordinate
        # capture compatibility, but BLOCK TEST itself
        # does NOT use this classification to decide what
        # to test. The decisive condition is:
        #
        #     sequential_tracks is not empty
        # -------------------------------------------------
        if signal.endswith("SH"):
            if signal not in shunt_signal_order:
                shunt_signal_order.append(signal)

        elif signal.endswith("C"):
            if signal not in cal_signal_order:
                cal_signal_order.append(signal)

        else:
            if signal not in main_signal_order:
                main_signal_order.append(signal)

        if route.endswith("_C"):
            if signal not in cal_signal_order:
                cal_signal_order.append(signal)

        elif route.endswith("_SH"):
            if signal not in shunt_signal_order:
                shunt_signal_order.append(signal)

        else:
            if signal not in main_signal_order:
                main_signal_order.append(signal)

        # -------------------------------------------------
        # Add controlled tracks to coordinate dictionary.
        # Sequential tracks are also added because Block Test
        # directly operates on them.
        # -------------------------------------------------
        all_track_values = []

        if track_data:
            all_track_values.extend(
                t.strip().upper()
                for t in track_data.split(",")
                if t.strip()
            )

        if sequential_tracks:
            all_track_values.extend(
                t.strip().upper()
                for t in sequential_tracks.split(",")
                if t.strip()
            )

        for track in all_track_values:
            if track not in track_points:
                track_points[track] = ()

        lock_routes_data.append({
            "signal": signal,
            "route": route,
            "track_data": track_data,
            "sequential_tracks": sequential_tracks,
            "point_data": point_data,
            "approach_track": approach_track
        })

    # -----------------------------------------------------
    # BLOCK TEST FILTER
    #
    # ONLY rows with a real SEQUENTIAL_TRACKS value are
    # displayed/testable.
    # -----------------------------------------------------
    block_rows = [
        row for row in lock_routes_data
        if str(row.get("sequential_tracks", "")).strip()
    ]

    refresh_table()

    log(f"TOTAL TOC ROUTES LOADED : {len(lock_routes_data)}")
    log(f"BLOCK TEST ROUTES       : {len(block_rows)}")

    for row in lock_routes_data:

        seq = str(row.get("sequential_tracks", "")).strip()

        if seq:
            log(
                f"BLOCK TEST CANDIDATE : "
                f"{row['signal']} | {row['route']} | "
                f"SEQUENTIAL = {seq}"
            )
        else:
            log(
                f"SKIPPED (NO SEQUENTIAL TRACKS) : "
                f"{row['signal']} | {row['route']}"
            )

    log("TOC LOADED")

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

    ws.title = "TRACK FAILURE REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = "TEST FOR BLOCK INSTRUMENTS REPORT"

    title_cell.font = Font(
        size=16,
        bold=True
    )

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # Remove any fill
    title_cell.fill = PatternFill(fill_type=None)

    # =====================================================
    # DATE
    # =====================================================

    ws.merge_cells("A2:E2")

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
        "TEST",
        "PASS/FAIL",
        "DATE & TIME"

    ]

    ws.append(headers)

    header_fill = PatternFill(fill_type=None)

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
            bold=True
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

            row["SIGNAL"],
            row["ROUTE"],
            row["TRACK"],
            row["RESULT"],
            row["DATE & TIME"]

        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================

    fail_fill = PatternFill(
        fill_type="solid",
        fgColor="FF9999"
    )

    for row in ws.iter_rows(min_row=4):

        result = row[3].value

        for cell in row:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        if result == "FAIL":
            for cell in row:
                cell.fill = fail_fill

    # =====================================================
    # AUTO WIDTH
    # =====================================================

    for col in range(1, 6):

        letter = get_column_letter(col)

        max_len = 0

        for cell in ws[letter]:

            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[letter].width = max_len + 4

    # =====================================================
    # SUMMARY
    # =====================================================

    initiated = []
    not_initiated = []

    for row in report_rows:

        signal = str(row["SIGNAL"])

        if row["RESULT"] == "PASS":
            if signal not in initiated:
                initiated.append(signal)

        else:
            if signal not in not_initiated:
                not_initiated.append(signal)

    start = ws.max_row + 3

    labels = [
        ("TOTAL ROUTES", len(report_rows)),
        ("INITIATED ROUTES", ", ".join(initiated) if initiated else "0"),
        ("NON-INITIATED ROUTES", ", ".join(not_initiated) if not_initiated else "0")
    ]

    for i, (label, value) in enumerate(labels):
        ws.cell(row=start + i, column=1).value = label
        ws.cell(row=start + i, column=3).value = value

        ws.cell(row=start + i, column=1).font = Font(bold=True)

        ws.cell(row=start + i, column=1).alignment = Alignment(horizontal="left")
        ws.cell(row=start + i, column=3).alignment = Alignment(horizontal="left")

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(REPORT_FILE)

    log(f"REPORT SAVED : {REPORT_FILE}")


def check_track_status(track):

    if track not in track_points:
        log(f"{track} NOT FOUND")
        return "FAIL"

    x, y = map(int, track_points[track])

    log(f"CHECKING {track} AT X={x} Y={y}")

    SEARCH = 8     # Search 20 pixels around the stored point

    img = pyautogui.screenshot(
        region=(x - SEARCH, y - SEARCH,
                SEARCH * 2 + 1,
                SEARCH * 2 + 1)
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

    img = pyautogui.screenshot(region=(x, y, 1, 1))
    r, g, b = img.getpixel((0, 0))

    log(f"{main_signal} SIGNAL RGB = ({r},{g},{b})")

    # Signal RED
    if r > 180 and g < 120 and b < 120:
        return "FAIL"

    return "PASS"

# =========================================================
# CHECK GREEN
# =========================================================

def _signal_rgb(point):
    """Average 5x5 colour at a signal lamp coordinate, on ANY screen."""
    if not point or point[0] is None or point[1] is None:
        return None
    x, y = int(float(point[0])), int(float(point[1]))
    img = ImageGrab.grab(
        bbox=(x - 2, y - 2, x + 3, y + 3),
        all_screens=True
    ).convert("RGB")
    w, h = img.size
    px = [img.getpixel((i, j)) for i in range(w) for j in range(h)]
    return (
        sum(p[0] for p in px) // len(px),
        sum(p[1] for p in px) // len(px),
        sum(p[2] for p in px) // len(px)
    )


def is_signal_green(main_signal):
    """Signal is OFF (cleared): GREEN lit, or YELLOW / DOUBLE YELLOW lit.
    Automatic block signals and 2-aspect signals may clear to YELLOW
    instead of GREEN, so any proceed aspect is accepted."""

    if main_signal not in signals:
        log(f"{main_signal} NOT IN SIGNAL CONFIG")
        return False

    sig = signals[main_signal]

    for aspect in ("GREEN", "YELLOW", "DOUBLE_YELLOW"):

        point = sig.get(aspect)
        rgb = _signal_rgb(point)

        if rgb is None:
            continue

        r, g, b = rgb
        log(f"{main_signal} {aspect} AT {point} RGB = ({r},{g},{b})")

        if aspect == "GREEN" and g > r + 40 and g > b + 40:
            log(f"{main_signal} CLEARED : GREEN")
            return True

        if aspect != "GREEN" and r > 150 and g > 150 and b < 120:
            log(f"{main_signal} CLEARED : {aspect}")
            return True

    log(f"{main_signal} NOT CLEARED (NO GREEN / YELLOW LIT)")

    # Diagnosis: is the signal at RED, or is no lamp lit at all?
    red_rgb = _signal_rgb(sig.get("RED"))
    if red_rgb is not None:
        r, g, b = red_rgb
        log(f"{main_signal} RED AT {sig.get('RED')} RGB = ({r},{g},{b})")
        if r > 180 and g < 120:
            log(
                f"{main_signal} IS AT RED - a track in its block section "
                f"is DOWN (check sequential tracks / VPR are UP)"
            )
        else:
            log(
                f"{main_signal} NO LAMP LIT AT RED/GREEN COORDINATES - "
                f"check the signal coordinates in the MAIN sheet"
            )

    return False

# =========================================================
def is_signal_red(main_signal):

    if main_signal not in signals:
        return False

    rgb = _signal_rgb(signals[main_signal].get("RED"))

    if rgb is None:
        log(f"{main_signal} RED COORDINATE MISSING")
        return False

    r, g, b = rgb

    log(f"{main_signal} RED RGB = ({r},{g},{b})")

    if r > 180 and g < 120:
        return True

    return False


def toggle_track(track_name):

    try:

        log(f"TOGGLING {track_name}")

        # =====================================
        # Open Bit Chart
        # =====================================
        pyautogui.hotkey("ctrl", "b")
        time.sleep(1)

        # =====================================
        # Station Selection
        # =====================================

        panel = auto.WindowControl(Name="Test Panel")

        if not panel.Exists(5):
            raise Exception("Test Panel not found")

        station_window = panel.WindowControl(
            searchDepth=10,
            Name="Stations - Select Station"
        )

        # May open as a separate window on the OTHER screen
        if not station_window.Exists(5):
            station_window = auto.WindowControl(
                searchDepth=1,
                Name="Stations - Select Station"
            )

        if not station_window.Exists(3):
            raise Exception("Station selection window not found")

        list_ctrl = station_window.ListControl()

        if not list_ctrl.Exists(5):
            raise Exception("List control not found")

        station = list_ctrl.ListItemControl(Name="50051")

        if not station.Exists(5):
            raise Exception("Station 50051 not found")

        station.Click()
        time.sleep(0.5)

        station_window.ButtonControl(Name="OK").Click()
        time.sleep(2)

        # =====================================
        # Indications Window
        # =====================================
        panel = auto.WindowControl(Name="Test Panel")

        if not panel.Exists(5):
            raise Exception("Test Panel not found")

        ind_window = panel.WindowControl(
            searchDepth=10,
            Name="Station 50051: (Indications | Controls)"
        )

        # May open as a separate window on the OTHER screen
        if not ind_window.Exists(5):
            ind_window = auto.WindowControl(
                searchDepth=1,
                Name="Station 50051: (Indications | Controls)"
            )

        if not ind_window.Exists(3):
            raise Exception("Indications window not found")

        ind_window.SetActive()
        time.sleep(0.5)

        # =====================================
        # Clear Previous Selection (Optional)
        # =====================================
        #deselect_btn = ind_window.ButtonControl(Name="Deselect All")

       #if deselect_btn.Exists(3):
        #    deselect_btn.Click()
        #    time.sleep(0.5)

        # =====================================
        # Select Track
        # =====================================
        track_name = track_name.strip().upper()

        track_list = ind_window.ListControl()

        if not track_list.Exists(5):
            raise Exception("Track list not found")

        # =====================================
        # Select Track / VPR
        # =====================================

        if track_name.endswith("_VPR"):

            search_name = track_name

        else:

            search_name = f"{track_name}PR"

        track_item = track_list.ListItemControl(Name=search_name)

        if not track_item.Exists(5):
            raise Exception(f"{search_name} not found")

        track_item.Click()
        track_item.SetFocus()
        time.sleep(0.5)

        # =====================================
        # Toggle Track
        # =====================================
        transmit_btn = ind_window.ButtonControl(Name="Transmit")

        if not transmit_btn.Exists(3):
            raise Exception("Transmit button not found")

        transmit_btn.Click()

        time.sleep(1)

        title = ind_window.TitleBarControl()
        close_btn = title.ButtonControl(Name="Close")

        if close_btn.Exists(3):
            close_btn.Click()
            time.sleep(1)

        log(f"{track_name} TOGGLED")

        return True

    except Exception as e:

        log(f"TOGGLE FAILED : {track_name} : {e}")

        return False


# =========================================================
# SEQUENTIAL ROUTE RELEASE
# =========================================================

# =========================================================
# SEQUENTIAL ROUTE RELEASE
# =========================================================

def sequential_route_release(track_list):

    global currently_broken_tracks

    try:
        log("STARTING SEQUENTIAL ROUTE RELEASE")

        # =====================================
        # Open Bit Chart
        # =====================================
        pyautogui.hotkey("ctrl", "b")
        time.sleep(1)

        # =====================================
        # Station Selection
        # =====================================
        panel = auto.WindowControl(Name="Test Panel")

        if not panel.Exists(5):
            raise Exception("Test Panel not found")

        station_window = panel.WindowControl(
            searchDepth=10,
            Name="Stations - Select Station"
        )

        # May open as a separate window on the OTHER screen
        if not station_window.Exists(5):
            station_window = auto.WindowControl(
                searchDepth=1,
                Name="Stations - Select Station"
            )

        if not station_window.Exists(3):
            raise Exception("Station selection window not found")

        list_ctrl = station_window.ListControl()

        if not list_ctrl.Exists(5):
            raise Exception("Station list not found")

        station = list_ctrl.ListItemControl(Name="50051")

        if not station.Exists(5):
            raise Exception("Station 50051 not found")

        station.Click()
        time.sleep(0.5)

        station_window.ButtonControl(Name="OK").Click()
        time.sleep(2)

        # =====================================
        # Indications Window
        # =====================================
        panel = auto.WindowControl(Name="Test Panel")

        if not panel.Exists(5):
            raise Exception("Test Panel not found")

        ind_window = panel.WindowControl(
            searchDepth=10,
            Name="Station 50051: (Indications | Controls)"
        )

        # May open as a separate window on the OTHER screen
        if not ind_window.Exists(5):
            ind_window = auto.WindowControl(
                searchDepth=1,
                Name="Station 50051: (Indications | Controls)"
            )

        if not ind_window.Exists(3):
            raise Exception("Indications window not found")

        ind_window.SetActive()
        time.sleep(1)

        # =====================================
        # NORMALIZE TRACK LIST
        # =====================================
        track_list = [
            str(t).strip().upper()
            for t in track_list
            if str(t).strip()
        ]

        if not track_list:
            raise Exception("No tracks available for sequential release")

        log(f"SEQUENTIAL TRACK LIST : {track_list}")

        track_ctrl = ind_window.ListControl()

        if not track_ctrl.Exists(5):
            raise Exception("Track list not found")

        transmit_btn = ind_window.ButtonControl(Name="Transmit")

        if not transmit_btn.Exists(5):
            raise Exception("Transmit button not found")

        # =====================================
        # BREAK ALL TRACKS
        # =====================================
        for track in track_list:

            log(f"BREAKING {track}")

            if track.endswith("_VPR"):
                item_name = track
            else:
                item_name = f"{track}PR"

            track_item = track_ctrl.ListItemControl(
                Name=item_name
            )

            if not track_item.Exists(5):
                raise Exception(
                    f"{item_name} NOT FOUND DURING BREAK"
                )

            track_item.Click()
            track_item.SetFocus()

            time.sleep(0.5)

            transmit_btn.Click()

            time.sleep(2)

            # IMPORTANT:
            # Track is now physically broken, so remember it.
            currently_broken_tracks.add(track)

            log(f"{track} BROKEN")

        log("ALL TRACKS BROKEN")

        # =====================================
        # RESTORE ALL TRACKS
        # =====================================
        for track in track_list:

            log(f"RESTORING {track}")

            if track.endswith("_VPR"):
                item_name = track
            else:
                item_name = f"{track}PR"

            track_item = track_ctrl.ListItemControl(
                Name=item_name
            )

            if not track_item.Exists(5):
                raise Exception(
                    f"{item_name} NOT FOUND DURING RESTORE"
                )

            track_item.Click()
            track_item.SetFocus()

            time.sleep(0.5)

            transmit_btn.Click()

            time.sleep(2)

            # IMPORTANT:
            # Track has now been restored.
            currently_broken_tracks.discard(track)

            log(f"{track} RESTORED")

        log("ALL TRACKS RESTORED")

        # =====================================
        # CLOSE WINDOW
        # =====================================
        title = ind_window.TitleBarControl()

        close_btn = title.ButtonControl(Name="Close")

        if close_btn.Exists(3):
            close_btn.Click()
            time.sleep(1)

        log("SEQUENTIAL ROUTE RELEASE COMPLETED")

        return True

    except Exception as e:

        log(

            f"SEQUENTIAL ROUTE RELEASE FAILED : {e}"

        )

        log(

            f"TRACKS STILL BROKEN : "

            f"{sorted(currently_broken_tracks)}"

        )

        return False

# =========================================================
# CLEANUP ROUTE
# =========================================================

def cleanup_route(signal):

    try:

        log(f"CLEANING ROUTE : {signal}")

        click(signals[signal]["menu"])
        time.sleep(1)

        click_menu_item("Signal Cancel")

        log(f"{signal} SIGNAL CANCEL DONE")

        time.sleep(2)

        click(signals[signal]["menu"])
        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{signal} ROUTE RELEASE DONE")

        time.sleep(6)

    except Exception as e:

        log(f"CLEANUP FAILED : {e}")

# =========================================================
# SET ROUTE
# =========================================================

def set_route(signal, route):

    try:

        log("SETTING ROUTE")

        click(signals[signal]["menu"])

        time.sleep(1)

        click_menu_item(route.replace("-", "_"))

        time.sleep(2)

        log("ROUTE SET COMPLETED")

        time.sleep(2)

        return True

    except Exception as e:

        log(f"SET ROUTE FAILED : {e}")

        return False

# =========================================================
# FAIL CURRENT ROUTE
# =========================================================

# =========================================================
# FAIL CURRENT ROUTE
# =========================================================

def fail_current_route(signal, route, reason):

    global currently_broken_tracks

    log("=" * 60)
    log(f"BLOCK TEST FAILED : {reason}")
    log(
        f"BROKEN TRACKS BEFORE CLEANUP : "
        f"{sorted(currently_broken_tracks)}"
    )
    log("=" * 60)

    # =====================================================
    # RESTORE ANY TRACK THAT IS STILL BROKEN
    # =====================================================

    broken_tracks = list(currently_broken_tracks)

    if broken_tracks:

        log(
            f"EMERGENCY RESTORE STARTED : "
            f"{broken_tracks}"
        )

        for track in broken_tracks:

            log(
                f"EMERGENCY RESTORE : {track}"
            )

            restored = False

            # Try up to 2 times
            for attempt in range(1, 3):

                log(
                    f"{track} RESTORE ATTEMPT "
                    f"{attempt}/2"
                )

                if toggle_track(track):

                    restored = True

                    log(
                        f"{track} EMERGENCY RESTORE SUCCESS"
                    )

                    currently_broken_tracks.discard(
                        track
                    )

                    break

                else:

                    log(
                        f"{track} RESTORE ATTEMPT "
                        f"{attempt} FAILED"
                    )

                    time.sleep(1)

            if not restored:

                log(
                    f"CRITICAL : {track} "
                    f"COULD NOT BE RESTORED"
                )

    else:

        log(
            "NO TRACKS REQUIRE EMERGENCY RESTORE"
        )

    # =====================================================
    # NORMAL ROUTE CLEANUP
    # =====================================================

    if str(route).strip():
        cleanup_route(signal)
    else:
        log(f"{signal} : AUTOMATIC SIGNAL - NO ROUTE CLEANUP")

    # =====================================================
    # REPORT
    # =====================================================

    report_rows.append({

        "SIGNAL": signal,

        "ROUTE": route,

        "TRACK": f"BLOCK TEST - {reason}",

        "RESULT": "FAIL",

        "DATE & TIME":
            datetime.now().strftime(
                "%d-%m-%Y %H:%M:%S"
            )

    })

# =========================================================
# MAIN SIGNAL TEST
# =========================================================

def main_signal_test(row):

    global currently_broken_tracks

    signal = str(row["signal"]).replace(".0", "").strip().upper()

    route = row["route"]

    seq = row["sequential_tracks"]

    log("=" * 60)
    log(f"MAIN SIGNAL TEST : {signal}")
    log(f"ROUTE : {route}")
    log(f"SEQUENTIAL : {seq}")
    log("=" * 60)

    tracks = [
        x.strip().upper()
        for x in seq.split(",")
        if x.strip()
    ]

    # Reset broken-track state for this route
    currently_broken_tracks.clear()

    vpr = tracks[-1]

    log(f"VPR TRACK : {vpr}")
    log(f"TRACKS : {tracks}")

    # ============================================
    # SET ROUTE
    # ============================================

    log(f"SIGNAL FROM TOC = '{signal}'")
    log(f"SIGNAL TYPE = {type(signal)}")
    log(f"CONFIG KEYS = {list(signals.keys())}")

    if signal not in signals:
        fail_current_route(
            signal,
            route,
            "SIGNAL NOT FOUND IN CONFIG"
        )

        return

    if not str(route).strip():
        # AUTOMATIC SIGNAL - no route in TOC, nothing to set
        log(f"{signal} : AUTOMATIC SIGNAL - NO ROUTE TO SET")
        time.sleep(3)

    elif not set_route(signal, route):
        fail_current_route(
            signal,
            route,
            "SET ROUTE FAILED"
        )

        return

    if is_signal_green(signal):

        log("GREEN CHECK : PASS")



    else:

        fail_current_route(

            signal,

            route,

            "GREEN CHECK FAILED"

        )

        return

    # ============================================
    # BREAK VPR
    # ============================================

    log(f"BREAKING {vpr}")

    if not toggle_track(vpr):
        fail_current_route(
            signal,
            route,
            f"{vpr} BREAK FAILED"
        )

        return

    # IMPORTANT: VPR is now known to be broken
    currently_broken_tracks.add(vpr)

    time.sleep(3)

    log(f"{vpr} BROKEN")

    # ============================================
    # VERIFY RED
    # ============================================

    time.sleep(2)

    if is_signal_red(signal):

        log("RED CHECK : PASS")



    else:

        fail_current_route(

            signal,

            route,

            "RED CHECK FAILED"

        )

        return

    # ============================================
    # RESTORE VPR
    # ============================================

    log(f"RESTORING {vpr}")

    if not toggle_track(vpr):
        fail_current_route(
            signal,
            route,
            f"{vpr} RESTORE FAILED"
        )

        return

    # VPR is now restored
    currently_broken_tracks.discard(vpr)

    time.sleep(3)

    log(f"{vpr} RESTORED")

    if is_signal_green(signal):

        log("GREEN RESTORED : PASS")

        xt_track = tracks[1]

        log(f"XT TRACK : {xt_track}")

        log(f"BREAKING {xt_track}")

        if not toggle_track(xt_track):
            fail_current_route(
                signal,
                route,
                f"{xt_track} BREAK FAILED"
            )

            return

        # IMPORTANT: XT is now broken
        currently_broken_tracks.add(xt_track)

        time.sleep(3)

        log(f"{xt_track} BROKEN")

        # ============================================
        # VERIFY RED AFTER XT
        # ============================================

        time.sleep(2)

        if is_signal_red(signal):

            log("RED AFTER XT : PASS")



        else:

            fail_current_route(

                signal,

                route,

                "RED AFTER XT FAILED"

            )

            return

        # ============================================
        # RESTORE XT
        # ============================================

        log(f"RESTORING {xt_track}")

        if not toggle_track(xt_track):
            fail_current_route(
                signal,
                route,
                f"{xt_track} RESTORE FAILED"
            )

            return

        # XT is now restored
        currently_broken_tracks.discard(xt_track)

        time.sleep(3)

        log(f"{xt_track} RESTORED")

        # ============================================
        # SIGNAL SHOULD STILL BE RED
        # ============================================

        time.sleep(2)

        if is_signal_red(signal):

            log("SIGNAL STILL RED : PASS")



        else:

            fail_current_route(

                signal,

                route,

                "SIGNAL NOT RED AFTER XT RESTORE"

            )

            return

        if not sequential_route_release(tracks):
            fail_current_route(
                signal,
                route,
                "SEQUENTIAL ROUTE RELEASE FAILED"
            )

            return

        report_rows.append({

            "SIGNAL": signal,

            "ROUTE": route,

            "TRACK": "BLOCK TEST",

            "RESULT": "PASS",

            "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

        })



    else:

        fail_current_route(

            signal,

            route,

            "GREEN NOT RESTORED AFTER VPR"

        )

        return

# =========================================================
# RUN ENGINE
# =========================================================

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

    global running
    global report_rows

    report_rows = []

    root.iconify()

    time.sleep(2)

    # =====================================================
    # BLOCK TEST SCOPE
    #
    # ONLY TOC rows having SEQUENTIAL_TRACKS are testable.
    # This automatically excludes:
    #   - ordinary routes such as 1_A / 2_K
    #   - SHUNT routes such as 9_A
    #   - CALLING-ON routes such as 1C_A
    #   - any future route without SEQUENTIAL_TRACKS
    #
    # The filter is based on the actual TOC field, not on
    # signal naming.
    # =====================================================
    block_test_rows = []

    for row in lock_routes_data:

        sequential = str(
            row.get("sequential_tracks", "")
        ).strip()

        if not sequential:
            log(
                f"BLOCK TEST SKIP : "
                f"{row['signal']} -> {row['route']} "
                f"(NO SEQUENTIAL_TRACKS)"
            )
            continue

        block_test_rows.append(row)

    log(
        f"BLOCK TEST SCOPE : "
        f"{len(block_test_rows)} route(s) selected"
    )

    if not block_test_rows:
        running = False

        status_label.config(
            text="NO ROUTES",
            fg="#dc2626"
        )

        root.deiconify()

        messagebox.showwarning(
            "BLOCK TEST",
            "No routes with SEQUENTIAL_TRACKS were found in the TOC."
        )

        return

    # =====================================================
    # TEST ONLY SELECTED ROUTES
    # =====================================================
    for index, row in enumerate(block_test_rows, start=1):

        pause_event.wait()

        if not running:
            break

        signal = str(
            row["signal"]
        ).replace(".0", "").strip().upper()

        route = str(
            row["route"]
        ).replace(".0", "").strip().upper()

        sequential = str(
            row["sequential_tracks"]
        ).strip().upper()

        log("")
        log("=" * 60)
        log(
            f"BLOCK TEST {index}/{len(block_test_rows)} : "
            f"{signal} -> {route}"
        )
        log(f"SEQUENTIAL TRACKS : {sequential}")
        log("=" * 60)

        main_signal_test(row)

    running = False

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    create_report()

    # Reset capture mode
    global capture_waiting
    global capture_module
    global track_capture_mode

    capture_waiting = False
    capture_module = None
    track_capture_mode = False

    root.deiconify()

    log(
        f"BLOCK TEST COMPLETED : "
        f"{len(block_test_rows)} route(s) tested"
    )

    try:
        os.startfile(REPORT_FILE)
    except:
        pass

def create_excel():

    wb = Workbook()
    ws = wb.active
    ws.title = "TRACK COORDINATES"

    # Title
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

    # Headers
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

    wb = openpyxl.load_workbook(COORD_FILE)
    ws = wb.active

    ws.append([track, x, y])

    row = ws.max_row

    thin = Side(style="thin", color="000000")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in ws[row]:
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    # Auto-width
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

    if not os.path.exists(COORD_FILE):

        messagebox.showerror(
            "ERROR",
            "Track Coordinate Excel not found."
        )
        return

    wb = openpyxl.load_workbook(COORD_FILE)
    ws = wb.active

    track_points = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        if not row:
            continue

        if row[0] is None:
            continue

        track = str(row[0]).strip().upper()

        try:
            x = int(float(row[1]))
            y = int(float(row[2]))
        except (TypeError, ValueError, IndexError):
            log(f"SKIPPED ROW (NOT A COORDINATE) : {row}")
            continue

        track_points[track] = [x, y]

        log(f"{track} -> ({x},{y})")

    log("TRACK COORDINATES LOADED")


def show_next_track():

    global current_track_index

    if current_track_index >= len(track_capture_list):

        log("TRACK CAPTURE COMPLETED")

        messagebox.showinfo(
            "DONE",
            "All track coordinates captured."
        )
        return

    track = track_capture_list[current_track_index]

    update_capture_overlay(
        "TRACK",
        track,
        "CLICK : TRACK",
        "Move the mouse over this track, then press SPACE."
    )

def start_track_capture():

    global current_track_index

    if not track_capture_list:

        messagebox.showwarning(
            "WARNING",
            "Please import Track Coordinate Excel first."
        )
        return

    current_track_index = 0

    global capture_waiting
    global capture_module
    global track_capture_mode

    capture_module = "TRACK"
    track_capture_mode = True
    capture_waiting = True

    # Create popup first
    create_capture_overlay()

    # Then update it
    show_next_track()

def save_current_track(x, y):

    global current_track_index
    global track_capture_mode
    global capture_waiting

    track = track_capture_list[current_track_index]

    track_capture_data[track] = [x, y]

    track_undo_stack.append(track)

    log(f"{track} -> ({x},{y})")

    current_track_index += 1

    if current_track_index >= len(track_capture_list):

        track_capture_mode = False
        capture_waiting = False

        global capture_module
        capture_module = None

        destroy_capture_overlay()

        log("ALL TRACKS CAPTURED")

        save_track_coordinates()

        return

    show_next_track()


def start_capture():

    if not track_capture_list:

        messagebox.showwarning(
            "WARNING",
            "Import Track Coordinates first."
        )

        return

    start_track_capture()

# =========================================================
# START
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

    # ============================================
    # LOAD TRACK COORDINATES
    # ============================================

    try:

        load_track_coordinates()

        log("TRACKS AFTER LOADING:")
        log(str(track_points))

    except Exception as e:

        messagebox.showerror(
            "ERROR",
            f"Unable to load Track Coordinates.\n\n{e}"
        )

        return

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
        target=_run_engine_with_com,
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
    "TEST FOR BLOCK INSTRUMENTS"
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

    text="TEST FOR BLOCK INSTRUMENTS",

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

    command=create_setup,

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
    "IMPORT TRACK COORDINATES",
    import_track_coordinates,
    "#2563eb"
).pack(pady=8)

# ==========================
# NEW BUTTON
# ==========================

create_button(
    "CAPTURE TRACK COORDINATES",
    start_track_capture,
    "#9333ea"
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

        "LOCK_ROUTE",

        "SEQUENTIAL"

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
tree.heading(
    "SEQUENTIAL",
    text="SEQUENTIAL TRACKS"
)

tree.column("NO", width=80, anchor="center")

tree.column("SIGNAL", width=200, anchor="center")

tree.column("ROUTE", width=220, anchor="center")

tree.column(
    "LOCK_ROUTE",
    width=350,
    anchor="w"
)

tree.column(
    "SEQUENTIAL",
    width=350,
    anchor="w"
)

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
# GLOBAL SPACE CAPTURE LISTENER
# =========================================================
# The original Block Test code defined on_press(), but never
# started a pynput Listener. Therefore SPACE never reached
# on_press(). This listener remains active for the lifetime of
# the UI and only acts while capture_waiting is True.
# =========================================================

def start_global_capture_listener():

    global capture_keyboard_listener

    if capture_keyboard_listener is not None:
        return

    capture_keyboard_listener = pynput_keyboard.Listener(
        on_press=on_press
    )

    capture_keyboard_listener.daemon = True
    capture_keyboard_listener.start()

    log("GLOBAL SPACE CAPTURE LISTENER STARTED")


# =========================================================
# RUN
# =========================================================

start_global_capture_listener()

root.mainloop()
