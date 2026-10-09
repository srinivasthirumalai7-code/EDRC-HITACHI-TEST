# =========================================================
# UNIVERSAL YARD COORDINATE RECORDER
#
# Records the complete yard ONCE - MAIN signals, SHUNT signals,
# CALLING-ON signals, Points, Crank Handles, and LC Gates - and saves it
# as one universal multi-sheet coordinate file (MAIN / SHUNT / CAL /
# POINT / CH / LC).
#
# This file is what the Signal Clearance and Route Initiation programs
# both load via their "LOAD YARD COORDINATES" button - each one just
# reads whatever columns it actually needs and ignores the rest. This
# is the ONLY place coordinates get recorded; the testing programs no
# longer have their own capture flow.
# =========================================================

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog

import win32api

import time
import os
import threading

from openpyxl import Workbook, load_workbook

from pynput import keyboard

import pyautogui

from PIL import Image, ImageTk

# =========================================================
# FOLDERS
# =========================================================
SNAPSHOTS_FOLDER = "snapshots"

if not os.path.exists(SNAPSHOTS_FOLDER):
    os.makedirs(SNAPSHOTS_FOLDER)

# =========================================================
# GLOBALS - MAIN SIGNAL
# =========================================================
main_data = {}
main_signal_order = []
main_signal_aspects = {}

# =========================================================
# GLOBALS - CALLING ON
# =========================================================
cal_signal_data = {}
cal_signal_order = []

# =========================================================
# GLOBALS - SHUNT
# =========================================================
shunt_data = {}
shunt_signal_order = []

# =========================================================
# GLOBALS - POINTS
# =========================================================
point_data = {}
point_order = []

# =========================================================
# GLOBALS - CRANK HANDLE
# =========================================================
ch_data = {}
ch_order = []

# =========================================================
# GLOBALS - LC GATE
# =========================================================
lc_data = {}
lc_order = []

# =========================================================
# GLOBALS - TRACK COORDINATES
# =========================================================
track_data = {}
track_order = []

# =========================================================
# GLOBALS - SYSTEM CONTROLS
# =========================================================
system_control_data = {}
system_control_order = []

# =========================================================
# GLOBALS - SHARED CAPTURE STATE
# =========================================================
capture_module = None          # "MASTER" while a recording pass is active, else None
capture_index = 0
capture_list = []
record_stage = None
current_aspects = 2


# =========================================================================
# =========================================================================
#  SHARED UTILITIES
# =========================================================================
# =========================================================================

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


def autosize_columns(ws):

    for column in ws.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[column_letter].width = max_length + 5




# =========================================================================
# =========================================================================
#  SHARED CAPTURE DISPATCHER
# =========================================================================
# =========================================================================

last_space_time = 0
last_backspace_time = 0
last_digit_time = 0


def on_press(key):

    global last_space_time, last_backspace_time, last_digit_time

    if key == keyboard.Key.space:

        now = time.time()

        if now - last_space_time < 0.4:
            log("SPACE ignored (pressed too soon after previous capture)")
            return

        last_space_time = now

        if capture_module == "MASTER":
            x, y = win32api.GetCursorPos()
            root.after(0, lambda: dispatch_save_point(x, y))

        elif capture_module == "SINGLE":
            x, y = win32api.GetCursorPos()
            root.after(0, lambda: finish_single_capture(x, y))

        else:
            return

    elif key == keyboard.Key.backspace:

        now = time.time()

        if now - last_backspace_time < 0.4:
            return

        last_backspace_time = now

        if capture_module != "MASTER":
            return

        root.after(0, dispatch_undo)

    else:

        # Number keys 2 / 3 / 4 as a keyboard shortcut for the aspect-count
        # step, so the operator never has to reach for the mouse mid-flow.
        char = getattr(key, "char", None)

        if char in ("2", "3", "4") and capture_module == "MASTER":

            if capture_index < len(capture_master_list):

                sig_type, signal = capture_master_list[capture_index]

                if sig_type == "MAIN" and record_stage == "ask_aspects":

                    now = time.time()

                    if now - last_digit_time < 0.4:
                        return

                    last_digit_time = now

                    n = int(char)

                    root.after(0, lambda: set_aspects_and_continue(signal, n))


def dispatch_save_point(x, y):

    if capture_module == "MASTER":
        master_save_point(x, y)


def dispatch_undo():

    if capture_module == "MASTER":
        undo_last_capture()


keyboard.Listener(on_press=on_press).start()


# =========================================================================
#  SINGLE COORDINATE CAPTURE
# =========================================================================

single_capture_event = None
single_capture_result = None
single_capture_overlay = None
single_overlay_progress_label = None
single_overlay_signal_label = None
single_overlay_step_label = None
single_overlay_hint_label = None


def create_single_capture_overlay(title, instruction):
    """Create a single-coordinate guide using the same overlay style as the
    main universal recorder. It is non-blocking visually, always-on-top,
    uses the same dark panel/typography, and SPACE is captured globally by
    pynput just like the normal recorder."""

    global single_capture_overlay
    global single_overlay_progress_label
    global single_overlay_signal_label
    global single_overlay_step_label
    global single_overlay_hint_label

    single_capture_overlay = tk.Toplevel(root)
    single_capture_overlay.title("Coordinate Capture Guide")
    single_capture_overlay.configure(bg="#0f172a")
    single_capture_overlay.resizable(False, False)
    single_capture_overlay.attributes("-topmost", True)

    screen_w = single_capture_overlay.winfo_screenwidth()
    single_capture_overlay.geometry(f"420x320+{screen_w - 440}+40")

    single_capture_overlay.protocol(
        "WM_DELETE_WINDOW",
        cancel_single_capture
    )

    # Same keyboard protection as the main recorder overlay.
    single_capture_overlay.bind("<space>", lambda e: "break")
    single_capture_overlay.bind("<Return>", lambda e: "break")

    tk.Label(
        single_capture_overlay,
        text="COORDINATE CAPTURE",
        font=("Segoe UI", 11, "bold"),
        bg="#0f172a",
        fg="#64748b"
    ).pack(pady=(16, 0))

    single_overlay_progress_label = tk.Label(
        single_capture_overlay,
        text="UNIVERSAL RECORDER",
        font=("Segoe UI", 10),
        bg="#0f172a",
        fg="#94a3b8"
    )
    single_overlay_progress_label.pack(pady=(2, 10))

    single_overlay_signal_label = tk.Label(
        single_capture_overlay,
        text=title,
        font=("Segoe UI", 18, "bold"),
        bg="#0f172a",
        fg="#22c55e"
    )
    single_overlay_signal_label.pack()

    single_overlay_step_label = tk.Label(
        single_capture_overlay,
        text="CLICK THE REQUIRED ITEM",
        font=("Segoe UI", 15, "bold"),
        bg="#0f172a",
        fg="#22c55e"
    )
    single_overlay_step_label.pack(pady=(6, 8))

    single_overlay_hint_label = tk.Label(
        single_capture_overlay,
        text=instruction,
        font=("Segoe UI", 10),
        bg="#0f172a",
        fg="#cbd5e1",
        wraplength=380,
        justify="center"
    )
    single_overlay_hint_label.pack(pady=(0, 18))

    tk.Label(
        single_capture_overlay,
        text="Press SPACE to save coordinate",
        font=("Segoe UI", 10, "bold"),
        bg="#0f172a",
        fg="#22c55e"
    ).pack(pady=(0, 8))

    tk.Button(
        single_capture_overlay,
        text="CANCEL CAPTURE",
        command=cancel_single_capture,
        bg="#0f172a",
        fg="#64748b",
        activebackground="#0f172a",
        activeforeground="#ef4444",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 9, "underline"),
        takefocus=0
    ).pack(side="bottom", pady=(0, 10))

    single_capture_overlay.lift()
    single_capture_overlay.focus_force()


def start_single_capture(title, instruction):
    """Capture exactly one coordinate using the same global SPACE mechanism
    as the complete-yard recorder."""

    global capture_module, single_capture_event
    global single_capture_result, single_capture_overlay

    if capture_module is not None:
        log("Another coordinate capture is already active")
        return None

    capture_module = "SINGLE"
    single_capture_event = threading.Event()
    single_capture_result = None

    create_single_capture_overlay(title, instruction)

    # Same operator workflow as the main recorder: the guide stays visible
    # while the real panel remains accessible.
    root.iconify()

    while not single_capture_event.is_set():
        try:
            root.update()
            time.sleep(0.05)
        except Exception:
            break

    result = single_capture_result

    try:
        if single_capture_overlay is not None:
            single_capture_overlay.destroy()
    except Exception:
        pass

    single_capture_overlay = None
    capture_module = None

    try:
        root.deiconify()
        root.lift()
    except Exception:
        pass

    return result


def finish_single_capture(x, y):
    global single_capture_result

    if capture_module != "SINGLE":
        return

    single_capture_result = [int(x), int(y)]

    log(f"Coordinate Captured : ({int(x)}, {int(y)})")

    if single_capture_event is not None:
        single_capture_event.set()


def cancel_single_capture():
    global single_capture_result

    single_capture_result = None

    if single_capture_event is not None:
        single_capture_event.set()


# =========================================================================
# =========================================================================
#  UNIFIED COORDINATE CAPTURE
# =========================================================================
# =========================================================================

capture_master_list = []       # list of (signal_type, signal_name) tuples
undo_stack = []                 # each entry: {"capture_index", "record_stage", "clear"}

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
    "LC": "LC GATE",
    "SYSCTL": "SYSTEM CONTROL"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CAL": "#f97316",
    "SHUNT": "#a855f7",
    "POINT": "#14b8a6",
    "CH": "#eab308",
    "LC": "#f43f5e",
    "SYSCTL": "#0ea5e9"
}

capture_overlay = None
overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None


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

    if capture_overlay is None:
        return

    clear_overlay_buttons()

    overlay_progress_label.config(
        text=f"Signal {capture_index + 1} of {len(capture_master_list)}   |   {TYPE_LABELS[sig_type]}"
    )
    overlay_signal_label.config(text=signal, fg=TYPE_COLORS[sig_type])
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


def set_aspects_and_continue(signal, aspects):

    global record_stage, current_aspects

    push_undo(lambda s=signal: main_signal_aspects.pop(s, None))

    main_signal_aspects[signal] = aspects
    current_aspects = aspects
    record_stage = "RED"

    log(f"{signal} Aspects Set To {aspects}")

    master_next_capture()


def cancel_master_capture():

    global capture_module

    capture_module = None

    destroy_capture_overlay()

    root.deiconify()

    log("CAPTURE CANCELLED BY USER")

    undo_stack.clear()


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


# The Emergency Point Key / SM Key coordinates. They are stored in the
# SYSTEM_CONTROLS sheet under these exact names, which is what
# EMERGENCY_POINT_KEY_TEST.py reads (get_system_control).
KEY_CONTROLS = [
    ("EMERGENCY POINT KEY", "the Emergency Point Key control itself"),
    ("EMERGENCY POINT KEY IN", "the Emergency Point Key IN position indicator"),
    ("EMERGENCY POINT KEY OUT", "the Emergency Point Key OUT position indicator"),
    ("SM KEY IN", "the SM Key IN position indicator"),
    ("SM KEY OUT", "the SM Key OUT position indicator"),
]


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

    # ---- EMERGENCY POINT KEY / SM KEY ----
    #
    # Asked right after the crank handles, as 5 fixed coordinates.
    # No count/name prompts - they always have the same names, and they
    # are written into the SYSTEM_CONTROLS sheet so
    # EMERGENCY_POINT_KEY_TEST.py can read them directly.
    want_keys = messagebox.askyesno(
        "Emergency Point Key / SM Key",
        "Record the Emergency Point Key and SM Key coordinates?\n\n"
        "5 coordinates :\n"
        "  1. EMERGENCY POINT KEY (control)\n"
        "  2. EMERGENCY POINT KEY IN\n"
        "  3. EMERGENCY POINT KEY OUT\n"
        "  4. SM KEY IN\n"
        "  5. SM KEY OUT",
        parent=root
    )

    if want_keys:
        for control_name, _hint in KEY_CONTROLS:
            combined.append(("SYSCTL", control_name))

    # ---- LC GATE ----
    #
    # LC names are entered manually, exactly like MAIN / SHUNT / CAL /
    # POINT / CH. This is important because different TOCs may use
    # different LC identifiers, for example 44_LCP.
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


def start_quantity_capture():
    """Entry point for the new 'RECORD COMPLETE YARD' button. Does NOT
    require a TOC import - builds its own list purely from the quantities
    and names the operator enters, then reuses the exact same overlay /
    save-point / undo engine as the TOC-based flow.

    Each run is a FRESH, standalone recording of the whole yard - any
    previously recorded MAIN/SHUNT/CAL/POINT/CH/LC data is cleared first
    so new names never collide with leftover data from an earlier session.
    The main window stays normal/visible while you answer the count and
    ID prompts, and only minimizes once actual on-panel coordinate
    capture begins."""

    global capture_module, capture_master_list, capture_index
    global main_data, main_signal_order, main_signal_aspects
    global shunt_data, shunt_signal_order
    global cal_signal_data, cal_signal_order
    global point_data, point_order
    global ch_data, ch_order
    global lc_data, lc_order

    has_existing = bool(
        main_signal_order or shunt_signal_order or cal_signal_order
        or point_order or ch_order or lc_order
    )

    if has_existing:
        proceed = messagebox.askyesno(
            "Start Fresh Recording",
            "This clears all previously recorded yard coordinates and "
            "starts a brand new complete-yard recording.\n\nContinue?"
        )
        if not proceed:
            log("Record Complete Yard cancelled - existing data kept")
            return

    main_data = {}
    main_signal_order = []
    main_signal_aspects = {}

    shunt_data = {}
    shunt_signal_order = []

    cal_signal_data = {}
    cal_signal_order = []

    point_data = {}
    point_order = []

    ch_data = {}
    ch_order = []

    lc_data = {}
    lc_order = []

    refresh_summary()

    capture_master_list = build_quantity_capture_list()

    if not capture_master_list:
        log("No Elements Entered - Nothing To Capture")
        return

    capture_index = 0
    capture_module = "MASTER"

    undo_stack.clear()

    set_initial_stage_for_current()

    create_capture_overlay()

    root.iconify()

    log("Recording COMPLETE YARD (MAIN + SHUNT + CALLING-ON + POINTS + CRANK HANDLE + LC GATES)")
    log(f"Total Elements To Capture: {len(capture_master_list)}")

    master_next_capture()


def set_initial_stage_for_current():

    global record_stage

    sig_type, signal = capture_master_list[capture_index]

    if sig_type in ("MAIN", "SYSCTL"):
        record_stage = "coordinate"
    else:
        record_stage = "menu"


def master_next_capture():

    global capture_module

    if capture_index >= len(capture_master_list):

        capture_module = None

        destroy_capture_overlay()

        root.deiconify()

        refresh_summary()

        log("ALL COORDINATES CAPTURED")

        undo_stack.clear()

        messagebox.showinfo(
            "Capture Complete",
            f"All {len(capture_master_list)} elements captured successfully."
        )

        return

    sig_type, signal = capture_master_list[capture_index]

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

    elif sig_type == "SYSCTL":

        hint = dict(KEY_CONTROLS).get(signal, signal)

        update_capture_overlay(
            sig_type, signal, f"CLICK: {signal}",
            f"Move the mouse onto {hint}, then press SPACE."
        )


def advance_to_next_signal():

    global capture_index

    capture_index += 1

    if capture_index < len(capture_master_list):
        set_initial_stage_for_current()

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

    # =====================================================
    # EMERGENCY POINT KEY / SM KEY  (SYSTEM_CONTROLS sheet)
    # =====================================================
    elif sig_type == "SYSCTL":

        def _undo_sysctl(name=signal):
            system_control_data.pop(name, None)
            if name in system_control_order:
                system_control_order.remove(name)

        push_undo(_undo_sysctl)

        system_control_data[signal] = [x, y]

        if signal not in system_control_order:
            system_control_order.append(signal)

        log(f"{signal} Saved ({x},{y})")

        advance_to_next_signal()
        return


def record_track_coordinate():
    """Record one TRACK_CONFIG entry."""

    track = simpledialog.askstring(
        "Track Recorder",
        "ENTER TRACK NAME\n\nExample: 1CXTPR",
        parent=root
    )

    if not track or not track.strip():
        return

    track = track.strip().upper()

    result = start_single_capture(
        "TRACK COORDINATE RECORDER",
        f"Move the mouse to the centre of:\n\n{track}\n\n"
        "Press SPACE to save the coordinate."
    )

    if result is None:
        log("Track coordinate recording cancelled")
        return

    track_data[track] = result

    if track not in track_order:
        track_order.append(track)

    refresh_summary()

    log(f"Track Saved : {track} -> ({result[0]}, {result[1]})")

    messagebox.showinfo(
        "Track Recorder",
        f"Track coordinate saved.\n\n{track}\n"
        f"X = {result[0]}\nY = {result[1]}",
        parent=root
    )


def record_system_control_coordinate():
    """Record one SYSTEM_CONTROLS entry. The control name is entered by
    the operator, so the same recorder can store EMERGENCY POINT KEY,
    PC SMKEY, or any other panel/system control required by the test."""

    control = simpledialog.askstring(
        "System Control Recorder",
        "ENTER SYSTEM CONTROL NAME\n\n"
        "Examples:\n"
        "EMERGENCY POINT KEY\n"
        "PC SMKEY",
        parent=root
    )

    if not control or not control.strip():
        return

    control = control.strip().upper()

    result = start_single_capture(
        "SYSTEM CONTROL RECORDER",
        f"Move the mouse to the control:\n\n{control}\n\n"
        "Press SPACE to save the coordinate."
    )

    if result is None:
        log("System control recording cancelled")
        return

    system_control_data[control] = result

    if control not in system_control_order:
        system_control_order.append(control)

    refresh_summary()

    log(
        f"System Control Saved : {control} -> "
        f"({result[0]}, {result[1]})"
    )

    messagebox.showinfo(
        "System Controls",
        f"System control saved.\n\n{control}\n"
        f"X = {result[0]}\nY = {result[1]}",
        parent=root
    )


def master_export_coordinates():
    """Saves the COMPLETE YARD to one workbook, one sheet per element type:
    MAIN / SHUNT / CAL / POINT / CH / LC. Any other program only needs to
    open the sheet(s) it cares about - e.g. a shunt-only tool just reads
    the SHUNT sheet and ignores the rest."""

    if not (main_signal_order or cal_signal_order or shunt_signal_order
            or point_order or ch_order or lc_order):
        log("No Coordinates Available")
        return

    file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel File", "*.xlsx")]
    )

    if not file:
        return

    wb = Workbook()
    wb.remove(wb.active)   # drop the default blank sheet - we add named ones below

    # ---- MAIN ----
    ws = wb.create_sheet("MAIN")
    ws.append(["Signal", "Aspects", "Menu_X", "Menu_Y", "RED_X", "RED_Y",
               "YELLOW_X", "YELLOW_Y", "DOUBLE_YELLOW_X", "DOUBLE_YELLOW_Y",
               "GREEN_X", "GREEN_Y", "RouteInit_X", "RouteInit_Y"])

    for signal in main_signal_order:
        d = main_data.get(signal, {})
        menu = d.get("open_menu")
        red = d.get("RED")
        yellow = d.get("YELLOW")
        dy = d.get("DOUBLE_YELLOW")
        green = d.get("GREEN")
        ri = d.get("ROUTE_INIT")
        ws.append([
            signal, main_signal_aspects.get(signal, ""),
            menu[0] if menu else None, menu[1] if menu else None,
            red[0] if red else None, red[1] if red else None,
            yellow[0] if yellow else None, yellow[1] if yellow else None,
            dy[0] if dy else None, dy[1] if dy else None,
            green[0] if green else None, green[1] if green else None,
            ri[0] if ri else None, ri[1] if ri else None,
        ])
    autosize_columns(ws)

    # ---- SHUNT ----
    ws = wb.create_sheet("SHUNT")
    ws.append(["Signal", "Menu_X", "Menu_Y", "Indicator_X", "Indicator_Y",
               "RouteInit_X", "RouteInit_Y", "Snapshot_Path"])

    for signal in shunt_signal_order:
        d = shunt_data.get(signal, {})
        menu = d.get("menu_coordinate")
        ind = d.get("state_indicator")
        ri = d.get("route_init")
        snap = d.get("initial_snapshot")
        ws.append([
            signal,
            menu[0] if menu else None, menu[1] if menu else None,
            ind[0] if ind else None, ind[1] if ind else None,
            ri[0] if ri else None, ri[1] if ri else None,
            snap if snap else "",
        ])
    autosize_columns(ws)

    # ---- CALLING-ON ----
    ws = wb.create_sheet("CAL")
    ws.append(["Signal", "Menu_X", "Menu_Y", "Yellow_X", "Yellow_Y",
               "RouteInit_X", "RouteInit_Y"])

    for signal in cal_signal_order:
        d = cal_signal_data.get(signal, {})
        menu = d.get("menu")
        yellow = d.get("yellow")
        ri = d.get("route_init")
        ws.append([
            signal,
            menu[0] if menu else None, menu[1] if menu else None,
            yellow[0] if yellow else None, yellow[1] if yellow else None,
            ri[0] if ri else None, ri[1] if ri else None,
        ])
    autosize_columns(ws)

    # ---- POINT ----
    ws = wb.create_sheet("POINT")
    ws.append(["Point", "Menu_X", "Menu_Y", "Normal_X", "Normal_Y",
               "Reverse_X", "Reverse_Y", "Free_X", "Free_Y"])

    for name in point_order:
        d = point_data.get(name, {})
        menu = d.get("menu")
        nrm = d.get("normal")
        rev = d.get("reverse")
        free = d.get("free")
        ws.append([
            name,
            menu[0] if menu else None, menu[1] if menu else None,
            nrm[0] if nrm else None, nrm[1] if nrm else None,
            rev[0] if rev else None, rev[1] if rev else None,
            free[0] if free else None, free[1] if free else None,
        ])
    autosize_columns(ws)

    # ---- CRANK HANDLE ----
    ws = wb.create_sheet("CH")
    ws.append(["CrankHandle", "Menu_X", "Menu_Y", "IN_X", "IN_Y",
               "OUT_X", "OUT_Y", "ECH_X", "ECH_Y", "FREE_X", "FREE_Y"])

    for name in ch_order:
        d = ch_data.get(name, {})
        menu = d.get("menu")
        i_ = d.get("IN")
        o_ = d.get("OUT")
        e_ = d.get("ECH")
        f_ = d.get("FREE")
        ws.append([
            name,
            menu[0] if menu else None, menu[1] if menu else None,
            i_[0] if i_ else None, i_[1] if i_ else None,
            o_[0] if o_ else None, o_[1] if o_ else None,
            e_[0] if e_ else None, e_[1] if e_ else None,
            f_[0] if f_ else None, f_[1] if f_ else None,
        ])
    autosize_columns(ws)

    # ---- LC GATE ----
    #
    # The first-column value is the actual LC identifier used by the
    # current TOC: 44_LCP.
    ws = wb.create_sheet("LC")
    ws.append(["LCGate", "Menu_X", "Menu_Y", "IN_X", "IN_Y", "OUT_X", "OUT_Y"])

    for name in lc_order:
        d = lc_data.get(name, {})
        menu = d.get("menu")
        i_ = d.get("in")
        o_ = d.get("out")
        ws.append([
            name,
            menu[0] if menu else None, menu[1] if menu else None,
            i_[0] if i_ else None, i_[1] if i_ else None,
            o_[0] if o_ else None, o_[1] if o_ else None,
        ])
    autosize_columns(ws)

    # ---- TRACK CONFIG ----
    ws = wb.create_sheet("TRACK_CONFIG")
    ws.append(["Track", "X", "Y"])

    for track in track_order:
        point = track_data.get(track)
        if point:
            ws.append([track, point[0], point[1]])

    autosize_columns(ws)

    # ---- SYSTEM CONTROLS ----
    ws = wb.create_sheet("SYSTEM_CONTROLS")
    ws.append(["Control", "X", "Y"])

    for control in system_control_order:
        point = system_control_data.get(control)
        if point:
            ws.append([control, point[0], point[1]])

    autosize_columns(ws)

    wb.save(file)

    log(f"Complete Yard Coordinates Saved : {file}")


def master_import_coordinates():
    """Loads a COMPLETE YARD workbook (MAIN / SHUNT / CAL / POINT / CH / LC
    sheets).

    If MAIN/SHUNT/CAL signals are already listed (typically because a TOC
    was imported first), this load only fills in coordinates for those
    already-known signals - it will NOT add new rows for signals in the
    file that aren't part of the current TOC, so the table only ever shows
    what you're actually testing this session. Their coordinate data is
    still stored in memory in case a later TOC import includes them.

    If nothing is listed yet (no TOC imported this session), every signal
    found in the file is added, so the file can still be reviewed/loaded
    standalone before any TOC exists."""

    global main_data, main_signal_aspects, main_signal_order
    global cal_signal_data, cal_signal_order
    global shunt_data, shunt_signal_order
    global point_data, point_order
    global ch_data, ch_order
    global lc_data, lc_order

    main_order_locked = bool(main_signal_order)
    shunt_order_locked = bool(shunt_signal_order)
    cal_order_locked = bool(cal_signal_order)

    skipped_main = 0
    skipped_shunt = 0
    skipped_cal = 0

    file = filedialog.askopenfilename(filetypes=[("Excel File", "*.xlsx")])

    if not file:
        return

    try:

        wb = load_workbook(file)

        def cell(row, idx):
            return row[idx] if len(row) > idx else None

        def parse_int(value):
            """Safely converts a cell value to int. Returns None for blank
            cells, empty strings, or anything not a valid number - instead
            of raising, which used to abort the whole import on one stray
            empty cell (e.g. after hand-editing/renaming rows)."""

            if value is None:
                return None

            if isinstance(value, str) and value.strip() == "":
                return None

            try:
                return int(value)
            except (TypeError, ValueError):
                return None

        def parse_point(x_val, y_val):
            """Converts an (x, y) cell pair into [x, y] only if BOTH sides
            are valid numbers - a half-filled row (one coordinate typed,
            the other still blank) is treated as not-yet-recorded rather
            than a broken single-value point."""

            x = parse_int(x_val)
            y = parse_int(y_val)

            if x is None or y is None:
                return None

            return [x, y]

        # ---- MAIN ----
        if "MAIN" in wb.sheetnames:
            for row in wb["MAIN"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                signal = str(row[0]).strip().upper()
                aspects = parse_int(cell(row, 1))

                if signal not in main_data:
                    main_data[signal] = {
                        "open_menu": None, "RED": None, "YELLOW": None,
                        "DOUBLE_YELLOW": None, "GREEN": None, "ROUTE_INIT": None
                    }
                if signal not in main_signal_order:
                    if main_order_locked:
                        skipped_main += 1
                    else:
                        main_signal_order.append(signal)

                main_signal_aspects[signal] = aspects if aspects else 2

                pt = parse_point(cell(row, 2), cell(row, 3))
                if pt: main_data[signal]["open_menu"] = pt

                pt = parse_point(cell(row, 4), cell(row, 5))
                if pt: main_data[signal]["RED"] = pt

                pt = parse_point(cell(row, 6), cell(row, 7))
                if pt: main_data[signal]["YELLOW"] = pt

                pt = parse_point(cell(row, 8), cell(row, 9))
                if pt: main_data[signal]["DOUBLE_YELLOW"] = pt

                pt = parse_point(cell(row, 10), cell(row, 11))
                if pt: main_data[signal]["GREEN"] = pt

                pt = parse_point(cell(row, 12), cell(row, 13))
                if pt: main_data[signal]["ROUTE_INIT"] = pt

        # ---- SHUNT ----
        if "SHUNT" in wb.sheetnames:
            for row in wb["SHUNT"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                signal = str(row[0]).strip().upper()
                snap = cell(row, 7)

                if signal not in shunt_data:
                    shunt_data[signal] = {
                        "menu_coordinate": None, "state_indicator": None,
                        "initial_snapshot": None, "route_init": None
                    }
                if signal not in shunt_signal_order:
                    if shunt_order_locked:
                        skipped_shunt += 1
                    else:
                        shunt_signal_order.append(signal)

                pt = parse_point(cell(row, 1), cell(row, 2))
                if pt: shunt_data[signal]["menu_coordinate"] = pt

                pt = parse_point(cell(row, 3), cell(row, 4))
                if pt: shunt_data[signal]["state_indicator"] = pt

                pt = parse_point(cell(row, 5), cell(row, 6))
                if pt: shunt_data[signal]["route_init"] = pt

                if snap:
                    snap = str(snap).strip()
                    if snap and os.path.exists(snap):
                        shunt_data[signal]["initial_snapshot"] = snap
                    elif snap:
                        log(f"{signal} Snapshot path not found: {snap}")

        # ---- CALLING-ON ----
        if "CAL" in wb.sheetnames:
            for row in wb["CAL"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                signal = str(row[0]).strip().upper()

                if signal not in cal_signal_data:
                    cal_signal_data[signal] = {"menu": None, "yellow": None, "route_init": None}
                if signal not in cal_signal_order:
                    if cal_order_locked:
                        skipped_cal += 1
                    else:
                        cal_signal_order.append(signal)

                pt = parse_point(cell(row, 1), cell(row, 2))
                if pt: cal_signal_data[signal]["menu"] = pt

                pt = parse_point(cell(row, 3), cell(row, 4))
                if pt: cal_signal_data[signal]["yellow"] = pt

                pt = parse_point(cell(row, 5), cell(row, 6))
                if pt: cal_signal_data[signal]["route_init"] = pt

        # ---- POINT ----
        if "POINT" in wb.sheetnames:
            for row in wb["POINT"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                name = str(row[0]).strip().upper()

                if name not in point_data:
                    point_data[name] = {"menu": None, "normal": None, "reverse": None, "free": None}
                if name not in point_order:
                    point_order.append(name)

                pt = parse_point(cell(row, 1), cell(row, 2))
                if pt: point_data[name]["menu"] = pt

                pt = parse_point(cell(row, 3), cell(row, 4))
                if pt: point_data[name]["normal"] = pt

                pt = parse_point(cell(row, 5), cell(row, 6))
                if pt: point_data[name]["reverse"] = pt

                pt = parse_point(cell(row, 7), cell(row, 8))
                if pt: point_data[name]["free"] = pt

        # ---- CRANK HANDLE ----
        if "CH" in wb.sheetnames:
            for row in wb["CH"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                name = str(row[0]).strip().upper()

                if name not in ch_data:
                    ch_data[name] = {"menu": None, "IN": None, "OUT": None, "ECH": None, "FREE": None}
                if name not in ch_order:
                    ch_order.append(name)

                pt = parse_point(cell(row, 1), cell(row, 2))
                if pt: ch_data[name]["menu"] = pt

                pt = parse_point(cell(row, 3), cell(row, 4))
                if pt: ch_data[name]["IN"] = pt

                pt = parse_point(cell(row, 5), cell(row, 6))
                if pt: ch_data[name]["OUT"] = pt

                pt = parse_point(cell(row, 7), cell(row, 8))
                if pt: ch_data[name]["ECH"] = pt

                pt = parse_point(cell(row, 9), cell(row, 10))
                if pt: ch_data[name]["FREE"] = pt

        # ---- LC GATE ----
        if "LC" in wb.sheetnames:
            for row in wb["LC"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                name = str(row[0]).strip().upper()

                if name not in lc_data:
                    lc_data[name] = {"menu": None, "in": None, "out": None}
                if name not in lc_order:
                    lc_order.append(name)

                pt = parse_point(cell(row, 1), cell(row, 2))
                if pt: lc_data[name]["menu"] = pt

                pt = parse_point(cell(row, 3), cell(row, 4))
                if pt: lc_data[name]["in"] = pt

                pt = parse_point(cell(row, 5), cell(row, 6))
                if pt: lc_data[name]["out"] = pt

        # ---- TRACK CONFIG ----
        if "TRACK_CONFIG" in wb.sheetnames:
            for row in wb["TRACK_CONFIG"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue

                name = str(row[0]).strip().upper()
                pt = parse_point(cell(row, 1), cell(row, 2))

                if pt:
                    track_data[name] = pt

                    if name not in track_order:
                        track_order.append(name)

        # ---- SYSTEM CONTROLS ----
        if "SYSTEM_CONTROLS" in wb.sheetnames:
            for row in wb["SYSTEM_CONTROLS"].iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue

                name = str(row[0]).strip().upper()
                pt = parse_point(cell(row, 1), cell(row, 2))

                if pt:
                    system_control_data[name] = pt

                    if name not in system_control_order:
                        system_control_order.append(name)

        refresh_summary()

        log("Complete Yard Coordinates Loaded Successfully")

        total_skipped = skipped_main + skipped_shunt + skipped_cal
        if total_skipped:
            log(
                f"{total_skipped} signal(s) in the file are not part of the current "
                f"TOC (MAIN: {skipped_main}, SHUNT: {skipped_shunt}, CAL: {skipped_cal}) - "
                f"not added to the table, but their coordinates are kept in memory "
                f"in case a later TOC import includes them."
            )

    except Exception as e:

        log(f"Load Error : {e}")


def master_load_config():
    """No longer requires a TOC import first - the file itself carries
    every signal/element name, so it can be loaded standalone."""

    master_import_coordinates()




# =========================================================================
# =========================================================================
#  SUMMARY DISPLAY
# =========================================================================
# =========================================================================

def refresh_summary():

    summary_tree.delete(*summary_tree.get_children())

    rows = [
        ("MAIN Signals", len(main_signal_order)),
        ("SHUNT Signals", len(shunt_signal_order)),
        ("CALLING-ON Signals", len(cal_signal_order)),
        ("Points", len(point_order)),
        ("Crank Handles", len(ch_order)),
        ("LC Gates", len(lc_order)),
        ("Tracks", len(track_order)),
        ("System Controls", len(system_control_order)),
    ]

    for name, count in rows:
        summary_tree.insert("", "end", values=(name, count))

    total = sum(c for _, c in rows)

    summary_tree.insert("", "end", values=("TOTAL", total))


# =========================================================================
# =========================================================================
#  GUI
# =========================================================================
# =========================================================================

root = tk.Tk()

root.title("UNIVERSAL YARD COORDINATE RECORDER")

root.geometry("1300x850")
root.configure(bg="#e9edf2")
root.state("zoomed")

# =========================================================
# STYLE
# =========================================================

style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Treeview",
    background="white", foreground="black",
    rowheight=32, fieldbackground="white",
    font=("Segoe UI", 10)
)

style.configure(
    "Treeview.Heading",
    background="#1e293b", foreground="white",
    font=("Segoe UI", 11, "bold")
)

style.map("Treeview", background=[("selected", "#2563eb")])

# =========================================================
# HEADER
# =========================================================

header_frame = tk.Frame(root, bg="#0f172a", height=100)
header_frame.pack(fill="x")
header_frame.pack_propagate(False)

left_logo_frame = tk.Frame(header_frame, bg="#0f172a")
left_logo_frame.pack(side="left", padx=20)

try:
    ir_logo_path = r"C:\Users\Balamurali\PyCharmMiscProject\indian_railways.png"
    ir_logo_img = Image.open(ir_logo_path)
    ir_logo_img.thumbnail((80, 80), Image.Resampling.LANCZOS)
    ir_logo_photo = ImageTk.PhotoImage(ir_logo_img)
    ir_logo_label = tk.Label(left_logo_frame, image=ir_logo_photo, bg="#0f172a")
    ir_logo_label.image = ir_logo_photo
    ir_logo_label.pack(pady=5)
except Exception:
    tk.Label(
        left_logo_frame, text="INDIAN RAILWAYS",
        font=("Segoe UI", 10, "bold"), bg="#0f172a", fg="white"
    ).pack()

title_frame = tk.Frame(header_frame, bg="#0f172a")
title_frame.pack(side="left", expand=True)

tk.Label(
    title_frame, text="UNIVERSAL YARD COORDINATE RECORDER",
    font=("Segoe UI", 22, "bold"), bg="#0f172a", fg="white"
).pack(pady=(15, 0))

tk.Label(
    title_frame, text="Record once here - Signal Clearance and Route Initiation both load the same file",
    font=("Segoe UI", 10), bg="#0f172a", fg="#cbd5e1"
).pack()

right_logo_frame = tk.Frame(header_frame, bg="#0f172a")
right_logo_frame.pack(side="right", padx=20)

try:
    company_logo_path = r"C:\Users\Balamurali\PyCharmMiscProject\company_logo.png"
    company_logo_img = Image.open(company_logo_path)
    company_logo_img.thumbnail((130, 130), Image.Resampling.LANCZOS)
    company_logo_photo = ImageTk.PhotoImage(company_logo_img)
    company_logo_label = tk.Label(right_logo_frame, image=company_logo_photo, bg="#0f172a")
    company_logo_label.image = company_logo_photo
    company_logo_label.pack(pady=5)
except Exception:
    tk.Label(
        right_logo_frame, text="EDRC TEAM",
        font=("Segoe UI", 10, "bold"), bg="#0f172a", fg="white"
    ).pack()

# =========================================================
# MAIN AREA
# =========================================================

main_frame = tk.Frame(root, bg="#e9edf2")
main_frame.pack(fill="both", expand=True, padx=15, pady=15)

# =========================================================
# LEFT PANEL - CONTROL
# =========================================================

left_panel = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
left_panel.pack(side="left", fill="y", padx=(0, 10))

tk.Label(
    left_panel, text="CONTROL PANEL",
    font=("Segoe UI", 15, "bold"), bg="white", fg="#0f172a"
).pack(pady=20)


def create_button(text, command, color, width=26, height=2, font_size=11):

    return tk.Button(
        left_panel, text=text, command=command,
        bg=color, fg="white",
        activebackground=color, activeforeground="white",
        relief="flat", cursor="hand2",
        font=("Segoe UI", font_size, "bold"),
        width=width, height=height
    )


create_button(
    "RECORD COMPLETE YARD", start_quantity_capture, "#ea580c"
).pack(pady=(10, 8))

create_button(
    "LOAD EXISTING RECORDING", master_load_config, "#7c3aed"
).pack(pady=8)

create_button(
    "SAVE COORDINATES", master_export_coordinates, "#16a34a"
).pack(pady=8)

create_button(
    "RECORD TRACK", record_track_coordinate, "#ea580c"
).pack(pady=8)

create_button(
    "SYSTEM CONTROLS", record_system_control_coordinate, "#059669"
).pack(pady=8)

status_frame = tk.Frame(left_panel, bg="#f8fafc", bd=1, relief="solid")
status_frame.pack(fill="x", padx=15, pady=25)

tk.Label(
    status_frame, text="LAST ACTION",
    font=("Segoe UI", 11, "bold"), bg="#f8fafc", fg="#0f172a"
).pack(pady=10)

status_label = tk.Label(
    status_frame, text="READY",
    font=("Segoe UI", 16, "bold"), bg="#f8fafc", fg="#16a34a"
)
status_label.pack(pady=(0, 15))

# =========================================================
# RIGHT PANEL - SUMMARY + LOG
# =========================================================

right_panel = tk.Frame(main_frame, bg="#e9edf2")
right_panel.pack(side="left", fill="both", expand=True)

summary_title = tk.Label(
    right_panel, text="RECORDED SO FAR",
    font=("Segoe UI", 16, "bold"), bg="#e9edf2", fg="#0f172a"
)
summary_title.pack(anchor="w", pady=(0, 10))

summary_frame = tk.Frame(right_panel, bg="white", bd=1, relief="solid")
summary_frame.pack(fill="x")

summary_tree = ttk.Treeview(
    summary_frame, columns=("TYPE", "COUNT"),
    show="headings", height=7
)
summary_tree.heading("TYPE", text="ELEMENT TYPE")
summary_tree.heading("COUNT", text="RECORDED COUNT")
summary_tree.column("TYPE", width=300, anchor="w")
summary_tree.column("COUNT", width=200, anchor="center")
summary_tree.pack(fill="x", padx=10, pady=10)

refresh_summary()

# =========================================================
# LOG
# =========================================================

log_title = tk.Label(
    right_panel, text="LIVE OPERATION LOG",
    font=("Segoe UI", 16, "bold"), bg="#e9edf2", fg="#0f172a"
)
log_title.pack(anchor="w", pady=(20, 10))

log_frame = tk.Frame(right_panel, bg="white", bd=1, relief="solid")
log_frame.pack(fill="both", expand=True)

log_text = tk.Text(log_frame, bg="white", fg="black", font=("Consolas", 10), relief="flat")
log_text.pack(fill="both", expand=True, padx=10, pady=10)
log_text.config(state="disabled")

# =========================================================
# FOOTER
# =========================================================

footer = tk.Label(
    root,
    text="SPACE = CAPTURE  |  BACKSPACE = UNDO LAST  |  UNIVERSAL YARD COORDINATE RECORDER",
    bg="#0f172a", fg="white", font=("Segoe UI", 10)
)
footer.pack(fill="x")

# =========================================================
# RUN
# =========================================================

root.mainloop()
