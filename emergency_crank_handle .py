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
from PIL import ImageTk
from PIL import Image as PILImage
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

import pyautogui

import win32api
import win32con
import win32gui

import uiautomation as auto

from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "Emergency_Crank_Handle.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

crank_handle_points = {}
point_data = {}
lc_gate_points = {}
shunt_snapshots = {}

lock_routes_data = []

running = False

capture_queue = []

capture_index = 0

capture_mode = None

current_signal = None

current_step = 0

last_capture_time = 0

capture_waiting = False
captured_point = None

keyboard_listener = None

# =========================================================
# UNDO / BACKSPACE CONTROL
# =========================================================
undo_requested = False
undo_last_point = None

CONFIG_FILE = ""

config_file_path = ""

capture_current = 0
capture_total = 0

capture_window = None

lbl_type = None
lbl_signal = None
lbl_step = None
lbl_hint = None
lbl_progress = None
lbl_mouse = None

capture_overlay = None

overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None

TYPE_LABELS = {
    "MAIN": "MAIN SIGNAL",
    "CALLING_ON": "CALLING ON SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "CH": "CRANK HANDLE",
    "POINT": "POINT MACHINE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CALLING_ON": "#f97316",
    "SHUNT": "#a855f7",
    "CH": "#eab308",
    "POINT": "#14b8a6",
    "LC": "#f43f5e"
}
# =========================================================
# PAUSE CONTROL
# ====================================================
# =====

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
# CAPTURE OVERLAY GLOBALS
# =========================================================

TYPE_LABELS = {
    "MAIN": "MAIN SIGNAL",
    "CALLING_ON": "CALLING ON SIGNAL",
    "SHUNT": "SHUNT SIGNAL",
    "CH": "CRANK HANDLE",
    "POINT": "POINT MACHINE",
    "LC": "LC GATE"
}

TYPE_COLORS = {
    "MAIN": "#3b82f6",
    "CALLING_ON": "#f97316",
    "SHUNT": "#a855f7",
    "CH": "#eab308",
    "POINT": "#14b8a6",
    "LC": "#f43f5e"
}

capture_overlay = None
overlay_progress_label = None
overlay_signal_label = None
overlay_step_label = None
overlay_hint_label = None
overlay_button_frame = None

# =========================================================
# REPORT DATA STORAGE
# =========================================================

report_rows = []

# Stores the exact reason for the latest failed test.
last_fail_reason = ""

# =========================================================
# CALLING-ON 50051 RECOVERY STATE
# =========================================================
# Keeps track of the Calling-On approach track (for example
# 1CXTPR) so it can ALWAYS be restored after a failed test.
active_calling_on_track = ""
active_calling_on_test = False


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


def undo_last_capture():
    """
    Undo the current/last coordinate capture.

    BACKSPACE and the UNDO LAST CLICK button both call this function.
    """
    global capture_waiting
    global captured_point
    global undo_requested
    global undo_last_point

    # Only allow undo while coordinate capture is active
    if not capture_waiting:
        return

    undo_requested = True
    undo_last_point = captured_point

    # Stop the current capture_point() waiting loop
    captured_point = None
    capture_waiting = False

    log("UNDO LAST CLICK REQUESTED")


def create_capture_window():
    global capture_window
    global lbl_type
    global lbl_signal
    global lbl_step
    global lbl_progress
    global lbl_hint
    global lbl_mouse
    global overlay_button_frame

    capture_window = tk.Toplevel(root)

    capture_window.title("Coordinate Capture Guide")

    screen_w = capture_window.winfo_screenwidth()

    capture_window.geometry(
        f"420x320+{screen_w - 440}+40"
    )

    capture_window.configure(bg="#0f172a")

    capture_window.attributes("-topmost", True)

    capture_window.resizable(False, False)

    tk.Label(
        capture_window,
        text="COORDINATE CAPTURE",
        bg="#0f172a",
        fg="#64748b",
        font=("Segoe UI", 11, "bold")
    ).pack(pady=(16, 0))

    lbl_progress = tk.Label(
        capture_window,
        text="",
        bg="#0f172a",
        fg="#94a3b8",
        font=("Segoe UI", 10)
    )

    lbl_progress.pack(pady=(2, 10))

    # Retained for the existing capture logic. Its value is presented in the
    # progress line below, matching NEW_UI.py.
    lbl_type = tk.Label(capture_window, bg="#0f172a")

    lbl_signal = tk.Label(
        capture_window,
        text="S1",
        bg="#0f172a",
        fg="#3b82f6",
        font=("Segoe UI", 18, "bold")
    )

    lbl_signal.pack(pady=(0, 6))

    lbl_step = tk.Label(
        capture_window,
        text="",
        bg="#0f172a",
        fg="#22c55e",
        font=("Segoe UI", 15, "bold")
    )

    lbl_step.pack(pady=(0, 8))

    lbl_hint = tk.Label(
        capture_window,
        text="",
        bg="#0f172a",
        fg="#cbd5e1",
        font=("Segoe UI", 10),
        wraplength=380,
        justify="center"
    )

    lbl_hint.pack(pady=(0, 10))

    # The controls below mirror the guide layout in the team application.
    overlay_button_frame = tk.Frame(capture_window, bg="#0f172a")
    overlay_button_frame.pack(pady=(0, 6))

    tk.Button(
        capture_window, text="↶  UNDO LAST CLICK", command=undo_last_capture,
        bg="#f59e0b", fg="#1a1a1a", activebackground="#fbbf24",
        activeforeground="#1a1a1a", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=22, height=1, takefocus=0
    ).pack(pady=(10, 2))
    tk.Label(
        capture_window, text="(or press BACKSPACE)", font=("Segoe UI", 8),
        bg="#0f172a", fg="#475569"
    ).pack()
    tk.Button(
        capture_window, text="CANCEL CAPTURE", command=lambda: None,
        bg="#0f172a", fg="#64748b", activebackground="#0f172a",
        activeforeground="#ef4444", relief="flat", cursor="hand2",
        font=("Segoe UI", 9, "underline"), takefocus=0
    ).pack(side="bottom", pady=(0, 10))

    # Kept only because the existing capture code declares this label.
    lbl_mouse = tk.Label(capture_window, bg="#0f172a")

    # Match the team application: the guide is not visible at startup. It is
    # shown by update_capture_window() only after the setup prompts finish.
    capture_window.withdraw()

    # =========================================================
    # LOCAL KEYBOARD CAPTURE
    # =========================================================
    capture_window.bind("<space>", capture_space)
    capture_window.bind("<Key-space>", capture_space)
    capture_window.bind("<BackSpace>", undo_last_capture)


def update_capture_window(sig_type, signal, step, hint):
    type_colors = TYPE_COLORS
    type_names = TYPE_LABELS

    lbl_type.config(
        text=type_names.get(sig_type, sig_type),
        fg=type_colors.get(sig_type, "white")
    )

    lbl_progress.config(text=type_names.get(sig_type, sig_type))

    lbl_signal.config(fg=type_colors.get(sig_type, "white"))

    lbl_signal.config(text=signal)

    lbl_step.config(text=step.replace("CAPTURE ", "CLICK: "))

    lbl_hint.config(
        text=hint.replace("Move mouse to ", "Move the mouse onto ")
        .replace("\nPress SPACE", ", then press SPACE.")
    )

    capture_window.deiconify()
    capture_window.lift()
    capture_window.attributes("-topmost", True)
    capture_window.focus_force()
    capture_window.update()


def update_mouse_position():
    """Kept for compatibility; the refreshed guide uses the NEW_UI prompt text."""
    return


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


def capture_space(event=None):
    """Capture the current mouse position when SPACE is pressed."""
    global capture_waiting
    global captured_point

    if not capture_waiting:
        return "break"

    captured_point = win32api.GetCursorPos()
    capture_waiting = False

    log(
        f"SPACE PRESSED - COORDINATE CAPTURED: "
        f"{captured_point}"
    )

    return "break"


# =========================================================
# CAPTURE POINT
# =========================================================

def capture_point():
    global capture_waiting
    global captured_point
    global undo_requested

    capture_waiting = True
    captured_point = None
    undo_requested = False

    # Make sure the capture window receives keyboard focus
    capture_window.deiconify()
    capture_window.lift()
    capture_window.attributes("-topmost", True)
    capture_window.focus_force()

    while capture_waiting:
        root.update()

        # Allow Backspace / Undo to be processed
        capture_window.update()

        time.sleep(0.05)

    capture_window.withdraw()

    # ---------------------------------------------------------
    # BACKSPACE / UNDO
    # ---------------------------------------------------------
    if undo_requested:
        undo_requested = False

        log("CAPTURE CANCELLED BY BACKSPACE / UNDO")

        return None

    # ---------------------------------------------------------
    # NORMAL SPACE CAPTURE
    # ---------------------------------------------------------
    if captured_point is None:
        return None

    return list(captured_point)


# =========================================================
# CAPTURE CRANK HANDLE POINT
# =========================================================

def capture_crank_point(handle, point_name):
    root.attributes("-topmost", False)
    root.iconify()

    # Determine step color/type
    update_capture_window(
        "CH",
        handle,
        f"CAPTURE {point_name}",
        f"Move mouse to {point_name}\nPress SPACE"
    )

    return capture_point()


# =========================================================
# CAPTURE TRACK COORDINATES
# =========================================================


# =========================================================
# CREATE SIGNAL SETUP
# =========================================================

def themed_input_dialog(title, prompt, integer_only=False):
    """Display a modal input dialog that matches the application theme."""
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.configure(bg="#f8fafc")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    result = {"value": None}

    header = tk.Frame(dialog, bg="#0f172a", height=52)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text=title, bg="#0f172a", fg="white",
             font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=18, pady=14)

    content = tk.Frame(dialog, bg="#f8fafc")
    content.pack(fill="both", expand=True, padx=18, pady=16)
    tk.Label(content, text=prompt, bg="#f8fafc", fg="#0f172a",
             font=("Segoe UI", 10), justify="left", wraplength=330).pack(anchor="w")

    value_var = tk.StringVar()
    entry = tk.Entry(content, textvariable=value_var, font=("Segoe UI", 11),
                     relief="solid", bd=1, highlightthickness=1,
                     highlightbackground="#cbd5e1", highlightcolor="#2563eb")
    entry.pack(fill="x", pady=(12, 4), ipady=6)
    error_label = tk.Label(content, text="", bg="#f8fafc", fg="#dc2626",
                           font=("Segoe UI", 9))
    error_label.pack(anchor="w")

    buttons = tk.Frame(content, bg="#f8fafc")
    buttons.pack(fill="x", pady=(14, 0))

    def cancel():
        dialog.destroy()

    def submit(event=None):
        value = value_var.get().strip()
        if integer_only:
            try:
                value = int(value)
            except ValueError:
                error_label.config(text="Please enter a whole number.")
                entry.focus_set()
                return
        result["value"] = value
        dialog.destroy()

    tk.Button(buttons, text="CANCEL", command=cancel, bg="#e2e8f0", fg="#0f172a",
              activebackground="#cbd5e1", activeforeground="#0f172a", relief="flat",
              font=("Segoe UI", 9, "bold"), width=11, height=1).pack(side="right")
    tk.Button(buttons, text="CONTINUE", command=submit, bg="#2563eb", fg="white",
              activebackground="#1d4ed8", activeforeground="white", relief="flat",
              font=("Segoe UI", 9, "bold"), width=11, height=1).pack(side="right", padx=(0, 8))

    dialog.bind("<Return>", submit)
    dialog.bind("<Escape>", lambda event: cancel())
    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.update_idletasks()
    dialog.geometry(f"380x{dialog.winfo_reqheight()}+{root.winfo_rootx() + 80}+{root.winfo_rooty() + 80}")
    entry.focus_set()
    root.wait_window(dialog)
    return result["value"]


def center_dialog(dlg, width=380, height=280):
    """Use the same centred dialog placement as NEW_UI.py."""
    screen_w = dlg.winfo_screenwidth()
    screen_h = dlg.winfo_screenheight()
    dlg.geometry(f"{width}x{height}+{(screen_w - width) // 2}+{(screen_h - height) // 2}")


def themed_askinteger(title, prompt):
    """Exact NEW_UI-style count prompt."""
    result = {"value": 0}
    dlg = tk.Toplevel(root)
    dlg.title("How Many?")
    dlg.configure(bg="#0f172a")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.transient(root)
    dlg.grab_set()

    tk.Label(dlg, text="HOW MANY?", font=("Segoe UI", 11, "bold"),
             bg="#0f172a", fg="#64748b").pack(pady=(20, 4), padx=40)
    tk.Label(dlg, text=prompt, font=("Segoe UI", 16, "bold"),
             bg="#0f172a", fg="white", wraplength=340,
             justify="center").pack(pady=(0, 16), padx=20)

    entry_var = tk.StringVar(value="0")
    entry = tk.Entry(dlg, textvariable=entry_var, font=("Segoe UI", 20, "bold"),
                     justify="center", width=6, bg="#1e293b", fg="white",
                     insertbackground="white", relief="flat")
    entry.pack(pady=(0, 6), ipady=6)
    entry.focus_set()
    entry.select_range(0, tk.END)
    tk.Label(dlg, text="Enter 0 to skip this category", font=("Segoe UI", 9),
             bg="#0f172a", fg="#475569").pack(pady=(0, 14))

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
    tk.Button(btn_frame, text="CONFIRM", command=confirm, bg="#2563eb", fg="white",
              activebackground="#2563eb", activeforeground="white", relief="flat",
              cursor="hand2", font=("Segoe UI", 10, "bold"), width=12,
              height=1, takefocus=0).pack(side="left", padx=6)
    tk.Button(btn_frame, text="SKIP (0)", command=skip, bg="#0f172a", fg="#64748b",
              activebackground="#0f172a", activeforeground="#ef4444", relief="flat",
              cursor="hand2", font=("Segoe UI", 10, "underline"),
              takefocus=0).pack(side="left", padx=6)
    dlg.bind("<Return>", lambda event: confirm())
    center_dialog(dlg, 380, 260)
    dlg.wait_window()
    return result["value"]


def themed_askstring(title, prompt):
    """NEW_UI-style identifier prompt for signal names."""
    result = {"value": None}
    dlg = tk.Toplevel(root)
    dlg.title("Signal / Element ID")
    dlg.configure(bg="#0f172a")
    dlg.resizable(False, False)
    dlg.attributes("-topmost", True)
    dlg.transient(root)
    dlg.grab_set()

    tk.Label(dlg, text=title.upper(), font=("Segoe UI", 11, "bold"),
             bg="#0f172a", fg="#64748b").pack(pady=(20, 4), padx=40)
    tk.Label(dlg, text=prompt, font=("Segoe UI", 16, "bold"),
             bg="#0f172a", fg="white", wraplength=360,
             justify="center").pack(pady=(0, 12), padx=20)

    entry_var = tk.StringVar()
    entry = tk.Entry(dlg, textvariable=entry_var, font=("Segoe UI", 16, "bold"),
                     justify="center", width=18, bg="#1e293b", fg="white",
                     insertbackground="white", relief="flat")
    entry.pack(pady=(0, 6), ipady=6, padx=20)
    entry.focus_set()

    def confirm():
        result["value"] = entry_var.get()
        dlg.destroy()

    tk.Button(dlg, text="CONFIRM", command=confirm, bg="#2563eb", fg="white",
              activebackground="#2563eb", activeforeground="white", relief="flat",
              cursor="hand2", font=("Segoe UI", 10, "bold"), width=14,
              height=1, takefocus=0).pack(pady=(14, 20))
    dlg.bind("<Return>", lambda event: confirm())
    center_dialog(dlg, 400, 220)
    dlg.wait_window()
    return result["value"]


def collect_element_names(title, count_prompt, name_prompt, prefix):
    """Collect all IDs before beginning the coordinate capture session."""
    names = []
    for index in range(themed_askinteger(title, count_prompt)):
        name = themed_askstring(title, f"{name_prompt} {index + 1}")
        names.append((name or f"{prefix}{index + 1}").strip().upper())
    return names


def select_main_aspects_during_capture(signal):
    """The team UI selects 2/3/4 aspects after the MENU coordinate."""
    result = {"value": 0}
    # Use the guide's dedicated button row. It sits above Undo/Cancel, just
    # like the team application, so the aspect buttons are never clipped.
    choices = overlay_button_frame
    for widget in choices.winfo_children():
        widget.destroy()

    lbl_progress.config(text="MAIN SIGNAL")
    lbl_signal.config(text=signal, fg="#3b82f6")
    lbl_step.config(text="HOW MANY ASPECTS?")
    lbl_hint.config(text="Click 2 / 3 / 4 below, or press the matching number key.")

    def choose(value):
        result["value"] = value
        for widget in choices.winfo_children():
            widget.destroy()

    for value in (2, 3, 4):
        tk.Button(
            choices, text=str(value), command=lambda n=value: choose(n), width=6,
            height=1, font=("Segoe UI", 12, "bold"), bg="#2563eb", fg="white",
            activebackground="#2563eb", activeforeground="white", relief="flat",
            cursor="hand2", takefocus=0
        ).pack(side="left", padx=6)

    capture_window.deiconify()
    capture_window.lift()
    capture_window.bind("<Key-2>", lambda event: choose(2))
    capture_window.bind("<Key-3>", lambda event: choose(3))
    capture_window.bind("<Key-4>", lambda event: choose(4))
    capture_window.focus_force()
    while not result["value"]:
        root.update()
        time.sleep(0.05)
    capture_window.unbind("<Key-2>")
    capture_window.unbind("<Key-3>")
    capture_window.unbind("<Key-4>")
    capture_window.withdraw()
    return result["value"]


def capture_team_step(sig_type, name, step, hint):
    update_capture_window(
        sig_type,
        name,
        step,
        hint
    )

    point = capture_point()

    return point


def _save_shunt_snapshot(name):
    """Capture a reference image for SHUNT without changing test logic."""
    global shunt_snapshots
    try:
        folder = os.path.join(os.getcwd(), "shunt_snapshots")
        os.makedirs(folder, exist_ok=True)
        safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in name)
        path = os.path.join(folder, f"{safe}_initial.png")
        pyautogui.screenshot().save(path)
        shunt_snapshots[name] = path
        return path
    except Exception as e:
        log(f"SHUNT SNAPSHOT FAILED: {e}")
        shunt_snapshots[name] = None
        return None


def create_setup_team_style():
    """TL-style coordinate capture UI; existing automation logic is preserved."""
    global signals, crank_handle_points, point_data, lc_gate_points, shunt_snapshots

    if not lock_routes_data:
        messagebox.showwarning("WARNING", "Please load TOC file first.")
        return

    main_names = collect_element_names("MAIN SIGNALS", "How Many Main Signals?", "Enter Main Signal Name", "MAIN")
    calling_names = collect_element_names("CALLING ON", "How Many Calling ON Signals?", "Enter Calling ON Signal Name",
                                          "CALLING")
    shunt_names = collect_element_names("SHUNT", "How Many Shunt Signals?", "Enter Shunt Signal Name", "SHUNT")
    handle_names = collect_element_names("EMERGENCY CRANK HANDLE", "How Many Crank Handles?", "Enter Crank Handle Name",
                                         "CH")
    point_names = collect_element_names("POINT MACHINE", "How Many Points?", "Enter Point Name", "POINT")
    lc_names = collect_element_names("LC GATE", "How Many LC Gates?", "Enter LC Gate Name", "LC")

    signals = {}
    crank_handle_points = {}
    point_data = {}
    lc_gate_points = {}
    shunt_snapshots = {}

    for name in main_names:
        signals[name] = {"type": "MAIN", "aspects": 0, "menu": None, "RED": None,
                         "YELLOW": None, "DOUBLE_YELLOW": None, "GREEN": None,
                         "ROUTE_INIT": None}
    for name in calling_names:
        signals[name] = {"type": "CALLING_ON", "menu": None, "YELLOW": None, "ROUTE_INIT": None}
    for name in shunt_names:
        # Keep C1/C2/C3 for compatibility with existing automation/config.
        # New UI meanings: C1=aspect indicator, C2/C3=route-init compatibility.
        signals[name] = {"type": "SHUNT", "menu": None,
                         "C1": None, "C2": None, "C3": None,
                         "state_indicator": None, "route_init": None,
                         "initial_snapshot": None}
    for name in handle_names:
        crank_handle_points[name] = {"BUTTON": None, "IN": None, "OUT": None, "ECH": None, "FREE": None}
    for name in point_names:
        point_data[name] = {"MENU": None, "NORMAL": None, "REVERSE": None, "FREE": None}
    for name in lc_names:
        lc_gate_points[name] = {"MENU": None, "IN": None, "OUT": None}

    if not signals and not crank_handle_points and not point_data and not lc_gate_points:
        log("No elements entered - nothing to capture")
        return

    root.iconify()
    try:
        # ---------------- MAIN ----------------
        for name in main_names:
            info = signals[name]
            info["menu"] = capture_team_step("MAIN", name, "CAPTURE MENU", "Move mouse to MENU\nPress SPACE")
            info["aspects"] = select_main_aspects_during_capture(name)
            info["RED"] = capture_team_step("MAIN", name, "CAPTURE RED", "Move mouse to RED\nPress SPACE")
            if info["aspects"] >= 3:
                info["YELLOW"] = capture_team_step("MAIN", name, "CAPTURE YELLOW", "Move mouse to YELLOW\nPress SPACE")
            if info["aspects"] == 4:
                info["DOUBLE_YELLOW"] = capture_team_step("MAIN", name, "CAPTURE DOUBLE YELLOW",
                                                          "Move mouse to DOUBLE YELLOW\nPress SPACE")
            info["GREEN"] = capture_team_step("MAIN", name, "CAPTURE GREEN", "Move mouse to GREEN\nPress SPACE")
            info["ROUTE_INIT"] = capture_team_step("MAIN", name, "CAPTURE ROUTE INITIATION",
                                                   "Move mouse to ROUTE INITIATION\nPress SPACE")

        # ---------------- CALLING ON ----------------
        for name in calling_names:
            info = signals[name]
            info["menu"] = capture_team_step("CALLING_ON", name, "CAPTURE MENU", "Move mouse to MENU\nPress SPACE")
            info["YELLOW"] = capture_team_step("CALLING_ON", name, "CAPTURE YELLOW",
                                               "Move mouse to YELLOW\nPress SPACE")
            info["ROUTE_INIT"] = capture_team_step("CALLING_ON", name, "CAPTURE ROUTE INITIATION",
                                                   "Move mouse to ROUTE INITIATION INDICATOR\nPress SPACE")

        # ---------------- SHUNT ----------------
        for name in shunt_names:
            info = signals[name]
            info["menu"] = capture_team_step("SHUNT", name, "CAPTURE MENU", "Move mouse to MENU\nPress SPACE")
            info["state_indicator"] = capture_team_step("SHUNT", name, "CAPTURE ASPECT INDICATOR",
                                                        "Move mouse to the aspect indicator\nPress SPACE")
            info["C1"] = info["state_indicator"]
            info["initial_snapshot"] = _save_shunt_snapshot(name)
            info["route_init"] = capture_team_step("SHUNT", name, "CAPTURE ROUTE INITIATION",
                                                   "Move mouse to ROUTE INITIATION INDICATOR\nPress SPACE")
            # Keep legacy C2/C3 populated so existing automation/config code does not crash.
            info["C2"] = info["route_init"]
            info["C3"] = info["route_init"]

        # ---------------- EMERGENCY CRANK HANDLE ----------------
        for name in handle_names:
            info = crank_handle_points[name]
            for key, label in (("BUTTON", "RED BUTTON"), ("IN", "IN"), ("OUT", "OUT"), ("ECH", "ECH"),
                               ("FREE", "FREE")):
                info[key] = capture_team_step("CH", name, f"CAPTURE {label}", f"Move mouse to {label}\nPress SPACE")

        # ---------------- POINT MACHINE ----------------
        for name in point_names:
            info = point_data[name]
            info["MENU"] = capture_team_step("POINT", name, "CAPTURE MENU", "Move mouse to MENU\nPress SPACE")
            info["NORMAL"] = capture_team_step("POINT", name, "CAPTURE NORMAL",
                                               "Move mouse to NORMAL indication\nPress SPACE")
            info["REVERSE"] = capture_team_step("POINT", name, "CAPTURE REVERSE",
                                                "Move mouse to REVERSE indication\nPress SPACE")
            info["FREE"] = capture_team_step("POINT", name, "CAPTURE FREE",
                                             "Move mouse to FREE / OUT OF CORRESPONDENCE indication\nPress SPACE")

        # ---------------- LC GATE ----------------
        for name in lc_names:
            info = lc_gate_points[name]
            info["MENU"] = capture_team_step("LC", name, "CAPTURE MENU", "Move mouse to MENU\nPress SPACE")
            info["IN"] = capture_team_step("LC", name, "CAPTURE IN", "Move mouse to IN indication\nPress SPACE")
            info["OUT"] = capture_team_step("LC", name, "CAPTURE OUT", "Move mouse to OUT indication\nPress SPACE")

        log("ALL COORDINATES CAPTURED")
    finally:
        capture_window.withdraw()
        root.deiconify()


def create_setup():
    return create_setup_team_style()

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
    point_data = {}
    lc_gate_points = {}
    shunt_snapshots = {}

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    total_main = themed_askinteger(
        "MAIN SIGNALS",
        "How Many Main Signals?"
    )

    if total_main is None:
        return

    for i in range(total_main):

        # =================================================
        # NAME
        # =================================================

        name = themed_askstring(
            "MAIN SIGNAL",
            f"Enter Main Signal Name {i + 1}"
        )

        if not name:
            return

        name = name.strip().upper()

        # =================================================
        # ASPECTS
        # =================================================

        aspects = themed_askinteger(
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

            "GREEN": None,

            "ROUTE_INIT": None
        }

        # =================================================
        # MENU
        # =================================================

        root.iconify()

        update_capture_window(
            "MAIN",
            name,
            "CAPTURE MENU",
            "Move mouse to MENU\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        # =================================================
        # RED
        # =================================================

        update_capture_window(
            "MAIN",
            name,
            "CAPTURE RED",
            "Move mouse to RED\nPress SPACE"
        )

        signals[name]["RED"] = capture_point()

        # =================================================
        # YELLOW
        # =================================================

        if aspects >= 3:
            update_capture_window(
                "MAIN",
                name,
                "CAPTURE YELLOW",
                "Move mouse to YELLOW\nPress SPACE"
            )

            signals[name]["YELLOW"] = capture_point()

        # =================================================
        # DOUBLE YELLOW
        # =================================================

        if aspects == 4:
            update_capture_window(
                "MAIN",
                name,
                "CAPTURE DOUBLE YELLOW",
                "Move mouse to DOUBLE YELLOW\nPress SPACE"
            )

            signals[name]["DOUBLE_YELLOW"] = capture_point()

        # =================================================
        # GREEN
        # =================================================

        update_capture_window(
            "MAIN",
            name,
            "CAPTURE GREEN",
            "Move mouse to GREEN\nPress SPACE"
        )

        signals[name]["GREEN"] = capture_point()

        update_capture_window(
            "MAIN",
            name,
            "CAPTURE ROUTE INITIATION",
            "Move mouse to ROUTE INITIATION\nPress SPACE"
        )

        signals[name]["ROUTE_INIT"] = capture_point()

        root.deiconify()

    # =====================================================
    # CALLING ON
    # =====================================================

    total_calling = themed_askinteger(
        "CALLING ON",
        "How Many Calling ON Signals?"
    )

    if total_calling is None:
        return

    for i in range(total_calling):

        name = themed_askstring(
            "CALLING ON",
            f"Enter Calling ON Signal Name {i + 1}"
        )

        if not name:
            return

        signals[name] = {

            "type": "CALLING_ON",

            "menu": None,

            "YELLOW": None
        }

        root.iconify()

        update_capture_window(
            "CALLING_ON",
            name,
            "CAPTURE MENU",
            "Move mouse to MENU\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        update_capture_window(
            "CALLING_ON",
            name,
            "CAPTURE YELLOW",
            "Move mouse to YELLOW\nPress SPACE"
        )

        signals[name]["YELLOW"] = capture_point()

        root.deiconify()

    # =====================================================
    # SHUNT
    # =====================================================

    total_shunt = themed_askinteger(
        "SHUNT",
        "How Many Shunt Signals?"
    )

    if total_shunt is None:
        return

    for i in range(total_shunt):

        name = themed_askstring(
            "SHUNT",
            f"Enter Shunt Signal Name {i + 1}"
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

        update_capture_window(
            "SHUNT",
            name,
            "CAPTURE MENU",
            "Move mouse to MENU\nPress SPACE"
        )

        signals[name]["menu"] = capture_point()

        update_capture_window(
            "SHUNT",
            name,
            "CAPTURE C1",
            "Move mouse to C1\nPress SPACE"
        )

        signals[name]["C1"] = capture_point()

        update_capture_window(
            "SHUNT",
            name,
            "CAPTURE C2",
            "Move mouse to C2\nPress SPACE"
        )

        signals[name]["C2"] = capture_point()

        update_capture_window(
            "SHUNT",
            name,
            "CAPTURE C3",
            "Move mouse to C3\nPress SPACE"
        )

        signals[name]["C3"] = capture_point()
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

    total_ch = themed_askinteger(
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

            "GREEN",

            "ROUTE_INIT"
        ]

        if aspects == 2:

            steps = ["MENU", "RED", "GREEN", "ROUTE_INIT"]

        elif aspects == 3:

            steps = ["MENU", "RED", "YELLOW", "GREEN", "ROUTE_INIT"]

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

            steps = ["MENU", "RED", "GREEN", "ROUTE_INIT"]

        elif aspects == 3:

            steps = ["MENU", "RED", "YELLOW", "GREEN", "ROUTE_INIT"]

        else:

            steps = [
                "MENU",
                "RED",
                "YELLOW",
                "DOUBLE_YELLOW",
                "GREEN",
                "ROUTE_INIT"
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
# KEYBOARD
# =========================================================

def on_press(key):
    global capture_waiting
    global captured_point
    global undo_requested

    # =========================================================
    # SPACE = CAPTURE COORDINATE
    # =========================================================

    if key == pynput_keyboard.Key.space:

        if capture_waiting:
            captured_point = win32api.GetCursorPos()

            capture_waiting = False

            log(
                f"COORDINATE CAPTURED: "
                f"{captured_point}"
            )

        return

    # =========================================================
    # BACKSPACE = UNDO / CANCEL CURRENT CAPTURE
    # =========================================================

    if key == pynput_keyboard.Key.backspace:

        if capture_waiting:
            undo_requested = True

            captured_point = None

            capture_waiting = False

            log("BACKSPACE PRESSED - UNDO LAST CLICK")

        return


# =========================================================
# START GLOBAL KEYBOARD LISTENER
# =========================================================

keyboard_listener = pynput_keyboard.Listener(
    on_press=on_press
)

keyboard_listener.daemon = True
keyboard_listener.start()


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

        "Green_Y",

        "Route_Init_X",

        "Route_Init_Y"

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

        "C1_X",

        "C1_Y",

        "C2_X",

        "C2_Y",

        "C3_X",

        "C3_Y",

        "Indicator_X",

        "Indicator_Y",

        "Route_Init_X",

        "Route_Init_Y",

        "Snapshot_Path"

    ])

    # =====================================================
    # TRACKS
    # =====================================================

    # =====================================================
    # WRITE TRACK DATA
    # =====================================================

    # =====================================================
    # CRANK HANDLE
    # =====================================================

    ws_crank = wb.create_sheet("CRANK_HANDLES")

    ws_crank.append([

        "HANDLE",

        "BUTTON_X",
        "BUTTON_Y",

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
    # POINT MACHINE
    # =====================================================
    ws_point = wb.create_sheet("POINTS")
    ws_point.append([
        "POINT", "Menu_X", "Menu_Y",
        "Normal_X", "Normal_Y",
        "Reverse_X", "Reverse_Y",
        "Free_X", "Free_Y"
    ])

    # =====================================================
    # LC GATE
    # =====================================================
    ws_lc = wb.create_sheet("LC_GATES")
    ws_lc.append([
        "LC_GATE", "Menu_X", "Menu_Y",
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

                info["GREEN"][0],
                info["GREEN"][1],

                (info.get("ROUTE_INIT") or [None, None])[0],
                (info.get("ROUTE_INIT") or [None, None])[1]

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

                info["C1"][0],
                info["C1"][1],

                info["C2"][0],
                info["C2"][1],

                info["C3"][0],
                info["C3"][1],

                (info.get("state_indicator") or info.get("C1") or [None, None])[0],
                (info.get("state_indicator") or info.get("C1") or [None, None])[1],
                (info.get("route_init") or info.get("C2") or [None, None])[0],
                (info.get("route_init") or info.get("C2") or [None, None])[1],
                info.get("initial_snapshot")

            ])

    # =====================================================
    # WRITE CRANK HANDLE DATA
    # =====================================================

    for handle, info in crank_handle_points.items():
        ws_crank.append([

            handle,

            info["BUTTON"][0],
            info["BUTTON"][1],

            info["IN"][0],
            info["IN"][1],

            info["OUT"][0],
            info["OUT"][1],

            info["ECH"][0],
            info["ECH"][1],

            info["FREE"][0],
            info["FREE"][1]

        ])

    for name, info in point_data.items():
        ws_point.append([
            name,
            (info.get("MENU") or [None, None])[0], (info.get("MENU") or [None, None])[1],
            (info.get("NORMAL") or [None, None])[0], (info.get("NORMAL") or [None, None])[1],
            (info.get("REVERSE") or [None, None])[0], (info.get("REVERSE") or [None, None])[1],
            (info.get("FREE") or [None, None])[0], (info.get("FREE") or [None, None])[1]
        ])

    for name, info in lc_gate_points.items():
        ws_lc.append([
            name,
            (info.get("MENU") or [None, None])[0], (info.get("MENU") or [None, None])[1],
            (info.get("IN") or [None, None])[0], (info.get("IN") or [None, None])[1],
            (info.get("OUT") or [None, None])[0], (info.get("OUT") or [None, None])[1]
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
    time.sleep(1)

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
    global point_data, lc_gate_points, shunt_snapshots

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

            "GREEN": [row[10], row[11]],

            "ROUTE_INIT": (
                [row[12], row[13]]
                if len(row) > 13 and row[12] is not None
                else None
            )

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
            "ROUTE_INIT": ([row[5], row[6]] if len(row) > 6 and row[5] is not None else None)

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

        signals[signal_name] = {

            "type": "SHUNT",

            # MENU
            "menu": [row[1], row[2]],

            # NEW SIGNAL CONFIG
            # Indicator_X / Indicator_Y
            "state_indicator": (
                [row[3], row[4]]
                if len(row) > 4 and row[3] is not None
                else None
            ),

            # Route_Init_X / Route_Init_Y
            "route_init": (
                [row[5], row[6]]
                if len(row) > 6 and row[5] is not None
                else None
            ),

            "initial_snapshot": None,

            # -------------------------------------------------
            # LEGACY C1/C2/C3
            # Keep these because the existing automation
            # still expects them.
            # -------------------------------------------------
            "C1": (
                [row[3], row[4]]
                if len(row) > 4 and row[3] is not None
                else None
            ),

            "C2": (
                [row[5], row[6]]
                if len(row) > 6 and row[5] is not None
                else None
            ),

            "C3": (
                [row[5], row[6]]
                if len(row) > 6 and row[5] is not None
                else None
            )
        }

    # =====================================================
    # LOAD TRACKS
    # =====================================================

    # =====================================================
    # LOAD CRANK HANDLE
    # =====================================================

    if "CRANK_HANDLES" in wb.sheetnames:

        ws = wb["CRANK_HANDLES"]

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            if not row[0]:
                continue

            handle = str(row[0]).strip().upper()

            crank_handle_points[handle] = {

                "BUTTON": [row[1], row[2]],
                "IN": [row[3], row[4]],
                "OUT": [row[5], row[6]],
                "ECH": [row[7], row[8]],
                "FREE": [row[9], row[10]]

            }

    if "POINTS" in wb.sheetnames:
        ws = wb["POINTS"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            name = str(row[0]).strip().upper()
            point_data[name] = {
                "MENU": [row[1], row[2]],
                "NORMAL": [row[3], row[4]],
                "REVERSE": [row[5], row[6]],
                "FREE": [row[7], row[8]]
            }

    if "LC_GATES" in wb.sheetnames:
        ws = wb["LC_GATES"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            name = str(row[0]).strip().upper()
            lc_gate_points[name] = {
                "MENU": [row[1], row[2]],
                "IN": [row[3], row[4]],
                "OUT": [row[5], row[6]]
            }

    log("SIGNAL CONFIG LOADED")
    log("CRANK HANDLE CONFIG LOADED")
    log(f"Signals loaded: {list(signals.keys())}")


def load_universal_coordinates():
    """Loads coordinates from the multi-sheet Universal Yard Coordinate
    file (the same file every other suite program reads from), instead
    of this program's own native config format (MAIN_SIGNALS /
    CALLING_ON / SHUNT / CRANK_HANDLES / POINTS / LC_GATES). Populates
    `signals` (MAIN/CALLING_ON/SHUNT) plus the separate
    crank_handle_points / point_data / lc_gate_points dicts, using the
    universal file's actual sheet/column names (MAIN / CAL / SHUNT /
    CH / POINT / LC).

    NOTE: this program's crank handle points use the key "BUTTON" for
    the click coordinate (same role as the universal file's Menu_X/Y),
    and also want C1/C2/C3 for SHUNT - the universal file only stores
    one indicator point (Indicator_X/Y) plus Route Init, so C1 is
    filled from Indicator, C2/C3 from Route Init (matching the newer
    state_indicator/route_init fields this program already prefers)."""

    global signals
    global config_file_path
    global crank_handle_points
    global point_data, lc_gate_points

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
    point_data = {}
    lc_gate_points = {}

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
    loaded_ch = loaded_point = loaded_lc = 0

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
                    "ROUTE_INIT": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
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
                    "C1": indicator,
                    "C2": route_init,
                    "C3": route_init,
                    "state_indicator": indicator,
                    "route_init": route_init,
                    "initial_snapshot": snapshot_path
                }
                loaded_shunt += 1

    # ---- CH ----
    if "CH" in wb.sheetnames:
        ws_u = wb["CH"]
        hi = header_index_of(ws_u)
        name_col = hi.get("CRANKHANDLE")
        if name_col is not None:
            for row in ws_u.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, name_col) is None:
                    continue
                handle = str(cell(row, name_col)).strip().upper()
                crank_handle_points[handle] = {
                    "BUTTON": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "IN": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "OUT": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y"))),
                    "ECH": parse_point(cell(row, hi.get("ECH_X")), cell(row, hi.get("ECH_Y"))),
                    "FREE": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y")))
                }
                loaded_ch += 1

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
                point_data[name] = {
                    "MENU": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "NORMAL": parse_point(cell(row, hi.get("NORMAL_X")), cell(row, hi.get("NORMAL_Y"))),
                    "REVERSE": parse_point(cell(row, hi.get("REVERSE_X")), cell(row, hi.get("REVERSE_Y"))),
                    "FREE": parse_point(cell(row, hi.get("FREE_X")), cell(row, hi.get("FREE_Y")))
                }
                loaded_point += 1

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
                lc_gate_points[name] = {
                    "MENU": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "IN": parse_point(cell(row, hi.get("IN_X")), cell(row, hi.get("IN_Y"))),
                    "OUT": parse_point(cell(row, hi.get("OUT_X")), cell(row, hi.get("OUT_Y")))
                }
                loaded_lc += 1

    log("SIGNAL CONFIG LOADED (UNIVERSAL)")
    log("CRANK HANDLE CONFIG LOADED (UNIVERSAL)")
    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_cal} CALLING-ON, {loaded_shunt} SHUNT, "
        f"{loaded_ch} CRANK HANDLE, {loaded_point} POINT, {loaded_lc} LC"
    )
    log(f"Signals loaded: {list(signals.keys())}")


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
        title="Select Emergency Crank Handle TOC",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)

    # FIX: always read the sheet named TOC (master workbook has several
    # sheets); fall back to the active sheet for single-sheet files.
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
            headers.append(str(cell.value).strip().upper())

    # ==========================================
    # FIND REQUIRED COLUMNS
    # ==========================================

    signal_col = None
    route_col = None
    crank_col = None
    approach_track_col = None

    for i, h in enumerate(headers):

        if h == "SIGNAL":
            signal_col = i

        elif h == "ROUTE":
            route_col = i

        # FIX: take only the CRANK_HANDLE column. The TOC also has a
        # CRANK_HANDLE_BITS column (CH1KLNR,CH2KLNR) which was overwriting
        # this and made every handle "CH1KLNR NOT FOUND".
        elif "CRANK" in h and "BITS" not in h and crank_col is None:
            crank_col = i

        # -------------------------------------------------
        # CALLING-ON APPROACH TRACK
        # -------------------------------------------------
        # Your TOC contains:
        #     Approach_Track
        #
        # Example:
        #     1C -> 1C_A -> 1CXTPR
        #
        # This value is required for:
        #     CTRL+B -> 50051 -> Approach Track -> Transmit
        # -------------------------------------------------
        elif (
                h.replace("_", "")
                        .replace("-", "")
                        .replace(" ", "")
                        .strip()
                in ("APPROACHTRACK", "APPROACH")
        ):
            approach_track_col = i

    # ==========================================
    # CHECK REQUIRED COLUMNS
    # ==========================================

    if signal_col is None:
        messagebox.showerror(
            "ERROR",
            "SIGNAL column not found."
        )
        return

    if route_col is None:
        messagebox.showerror(
            "ERROR",
            "ROUTE column not found."
        )
        return

    if crank_col is None:
        messagebox.showerror(
            "ERROR",
            "CRANK HANDLE column not found."
        )
        return

    # FIX: TOC has no Approach_Track column - the Calling-On approach
    # track (1C -> 1CXTPR) is in the "Track" column. Use it as fallback.
    if approach_track_col is None and "TRACK" in headers:
        approach_track_col = headers.index("TRACK")

    if approach_track_col is not None:
        log(
            f"APPROACH_TRACK COLUMN FOUND : "
            f"{headers[approach_track_col]}"
        )
    else:
        log(
            "APPROACH_TRACK COLUMN NOT FOUND - "
            "Calling-On 50051 preparation needs a TOC value."
        )

    # ==========================================
    # READ DATA
    # ==========================================

    for row in ws.iter_rows(min_row=2, values_only=True):

        if row is None:
            continue

        if row[signal_col] is None:
            continue

        signal = str(row[signal_col]).replace(".0", "").strip().upper()

        route = str(row[route_col]).replace(".0", "").strip().upper()

        crank_handle = ""

        if row[crank_col]:
            crank_handle = str(
                row[crank_col]
            ).strip().upper()

        approach_track = ""

        if (
                approach_track_col is not None
                and approach_track_col < len(row)
                and row[approach_track_col] is not None
        ):

            approach_track = (
                str(row[approach_track_col])
                .strip()
                .upper()
            )

            # Remove a numeric Excel suffix only when it is actually
            # present. Do not alter names such as 1CXTPR.
            if approach_track.endswith(".0"):
                approach_track = approach_track[:-2].strip()

        lock_routes_data.append({
            "signal": signal,
            "route": route,
            "crank_handle": crank_handle,
            "approach_track": approach_track
        })

        if signal == "1C" or get_signal_type(signal) == "CALLING_ON":
            log(
                f"TOC -> {signal} {route} -> "
                f"APPROACH_TRACK = "
                f"{approach_track or '[EMPTY]'}"
            )

    refresh_table()

    log("EMERGENCY CRANK HANDLE TOC LOADED")


# =========================================================
# TABLE
# =========================================================

# One Treeview is maintained for each signal category.
# This affects display only; Emergency Crank Handle automation
# continues to use the existing lock_routes_data and test logic.
signal_trees = {}

# Kept for compatibility with any existing code that refers to
# the original "tree" variable.
tree = None


def split_dashboard_signals(signal):
    """
    Split a TOC SIGNAL cell into individual signal names for dashboard
    display only.

    Example:
        "1C,9SH,1" -> ["1C", "9SH", "1"]

    The original lock_routes_data value is NOT modified, so the
    Emergency Crank Handle automation data remains unchanged.
    """
    if signal is None:
        return []

    text = str(signal).replace(".0", "").strip().upper()

    if not text:
        return []

    return [
        part.strip()
        for part in text.split(",")
        if part.strip()
    ]


def get_signal_type(signal):
    """
    Classify one individual signal for the Emergency Crank Handle
    dashboard.

    Priority:
        1. Numeric SHUNT naming:
           9SH, 17SH, SH9, SH17 -> SHUNT
        2. Explicit signal configuration type
        3. Signal ending with C -> CALLING ON
        4. Otherwise -> MAIN
    """
    signal = str(signal or "").strip().upper()

    # ---------------------------------------------------------
    # SHUNT NAMING RULE - MUST BE CHECKED FIRST
    # ---------------------------------------------------------
    # This prevents a TOC/configuration entry from incorrectly
    # placing 9SH / SH9 in MAIN SIGNALS.
    if (
            (signal.endswith("SH") and signal[:-2].isdigit())
            or
            (signal.startswith("SH") and signal[2:].isdigit())
    ):
        return "SHUNT"

    # ---------------------------------------------------------
    # EXISTING CONFIGURED TYPE
    # ---------------------------------------------------------
    if signal in signals:
        configured_type = str(
            signals[signal].get("type", "MAIN")
        ).upper()

        if configured_type in ("MAIN", "CALLING_ON", "SHUNT"):
            return configured_type

    # ---------------------------------------------------------
    # CALLING ON NAMING RULE
    # ---------------------------------------------------------
    if signal.endswith("C"):
        return "CALLING_ON"

    return "MAIN"


def refresh_table():
    """
    Refresh the MAIN, CALLING ON and SHUNT dashboard tabs.

    A single TOC SIGNAL cell may contain multiple comma-separated
    signals, for example:

        1C,9SH,1

    Each signal is displayed separately in its correct dashboard tab:

        1C  -> CALLING ON SIGNALS
        9SH -> SHUNT SIGNALS
        1   -> MAIN SIGNALS

    Only the dashboard display is split. The original TOC row and
    automation data remain unchanged.
    """

    for table in signal_trees.values():
        table.delete(*table.get_children())

    counters = {
        "MAIN": 0,
        "CALLING_ON": 0,
        "SHUNT": 0
    }

    for row in lock_routes_data:

        raw_signal = str(
            row.get("signal", "")
        ).strip().upper()

        dashboard_signals = split_dashboard_signals(
            raw_signal
        )

        # If the signal cell is empty, do not create a dashboard row.
        if not dashboard_signals:
            continue

        for signal in dashboard_signals:

            signal_type = get_signal_type(signal)

            if signal_type not in signal_trees:
                signal_type = "MAIN"

            counters[signal_type] += 1

            signal_trees[signal_type].insert(
                "",
                "end",
                values=(
                    counters[signal_type],
                    signal,
                    row.get("route", ""),
                    row.get("crank_handle", "")
                )
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


def is_red_color(r, g, b):
    return r > 170 and g < 90 and b < 90


def is_yellow_color(r, g, b):
    return r > 140 and g > 140 and b < 120


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


# =========================================================
# VERIFY ALL CONTROLLED TRACKS
# =========================================================


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

# =========================================================
# DROP TRACK
# =========================================================

# =========================================================
# FIND AND CLICK UI
# =========================================================


def find_and_click(window_title, control_name, control_type=None):
    try:
        desktop = Desktop(backend="uia")

        for window in desktop.windows():

            try:
                for item in window.descendants():

                    try:
                        text = item.window_text().strip().upper()

                        if control_name.upper() in text:

                            if control_type:

                                current_type = str(item.element_info.control_type).upper()

                                if control_type.upper() not in current_type:
                                    continue

                            log(
                                f"Found: {item.window_text()} | "
                                f"Type: {item.element_info.control_type}"
                            )

                            item.click_input()

                            log(f"{control_name} CLICKED")

                            return True

                    except Exception as e:

                        log(str(e))

            except Exception as e:

                log(str(e))

    except Exception as e:
        log(str(e))

    log(f"{control_name} NOT FOUND")

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
        "Stations - Select Station",
        "50051"
    )

    if not found:
        log("50051 Not Found")

        return False

    time.sleep(1)

    pyautogui.press("enter")

    log("ENTER PRESSED")

    time.sleep(3)

    desktop = Desktop(backend="uia")

    for w in desktop.windows():
        if "STATIONS - SELECT STATION" in w.window_text().upper():
            log("STATION WINDOW STILL OPEN")
            return False

    log("STATION WINDOW CLOSED")

    return True


# =========================================================
# OPEN EMERGENCY CRANK HANDLE WINDOW
# =========================================================


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
    if not find_and_click(
            "Station 50051",
            "Transmit",
            "Button"):
        log("TRANSMIT BUTTON NOT FOUND")

        return False

    log("TRANSMIT CLICKED")

    time.sleep(3)

    return True


def click_station_cancel():
    if not find_and_click(
            "Station 50051",
            "Cancel",
            "Button"):
        log("CANCEL BUTTON NOT FOUND")

        return False

    log("CANCEL CLICKED")

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
    global last_fail_reason
    global active_calling_on_track
    global active_calling_on_test

    last_fail_reason = str(message)

    log(message)

    # -------------------------------------------------
    # IMPORTANT: CALLING-ON FAILURE RECOVERY
    # -------------------------------------------------
    # If a Calling-On test fails at ANY step, restore the
    # approach track in Station 50051 (for example 1CXTPR).
    # This prevents the next Calling-On test from starting
    # with 50051 left in the previous test state.
    if active_calling_on_test:

        track = active_calling_on_track

        if track:
            log(
                f"{signal} -> CALLING-ON FAILURE RECOVERY: "
                f"RELOADING {track}"
            )

            try:
                if restore_calling_on_50051(track):
                    log(
                        f"{signal} -> CALLING-ON "
                        f"{track} RELOADED AFTER FAILURE"
                    )
                else:
                    log(
                        f"{signal} -> CALLING-ON "
                        f"{track} RELOAD FAILED AFTER FAILURE"
                    )
            except Exception as e:
                log(
                    f"{signal} -> CALLING-ON "
                    f"{track} RELOAD ERROR: {e}"
                )
        else:
            log(
                f"{signal} -> CALLING-ON FAILURE RECOVERY "
                "SKIPPED: APPROACH_TRACK EMPTY"
            )

        # Clear the state so another test does not accidentally
        # restore the previous signal's approach track.
        active_calling_on_track = ""
        active_calling_on_test = False

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
    ws.title = "EMERGENCY CRANK HANDLE REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:F1")

    title_cell = ws["A1"]
    title_cell.value = "EMERGENCY CRANK HANDLE AUTOMATION REPORT"
    title_cell.font = Font(bold=True, size=18)
    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 35

    # =====================================================
    # GENERATED DATE / TIME
    # =====================================================

    ws.merge_cells("A2:F2")

    date_cell = ws["A2"]
    date_cell.value = (
        f"Generated : "
        f"{datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
    )
    date_cell.font = Font(bold=True, size=11)
    date_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # =====================================================
    # HEADERS
    # =====================================================

    ws.append([
        "SIGNAL",
        "ROUTE",
        "CRANK HANDLE",
        "RESULT",
        "REASON",
        "DATE & TIME"
    ])

    thin = Side(style="thin", color="000000")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for cell in ws[3]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        cell.border = border

    # =====================================================
    # DATA
    # =====================================================

    for row_data in report_rows:
        ws.append([
            row_data.get("SIGNAL", ""),
            row_data.get("ROUTE", ""),
            row_data.get("CRANK_HANDLES", ""),
            row_data.get("RESULT", ""),
            row_data.get("REASON", ""),
            row_data.get("DATE & TIME", "")
        ])

    # =====================================================
    # FAIL = ENTIRE ROW RED
    # =====================================================

    fail_fill = PatternFill(
        fill_type="solid",
        fgColor="FFC7CE"
    )

    fail_font = Font(
        color="9C0006"
    )

    for row in ws.iter_rows(min_row=4):

        result_value = str(
            row[3].value or ""
        ).strip().upper()

        if result_value == "FAIL":
            row_fill = fail_fill
            row_font = fail_font
        else:
            row_fill = PatternFill(fill_type=None)
            row_font = Font(color="000000")

        for cell in row:
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # =====================================================
    # COLUMN WIDTHS
    # =====================================================

    widths = {
        "A": 14,
        "B": 18,
        "C": 18,
        "D": 12,
        "E": 55,
        "F": 22
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    # =====================================================
    # SUMMARY
    # =====================================================

    # Count UNIQUE routes from the TOC.
    unique_routes = []
    seen_routes = set()

    for route_row in lock_routes_data:

        route_name = str(
            route_row.get("route", "")
        ).strip().upper()

        if route_name and route_name not in seen_routes:
            seen_routes.add(route_name)
            unique_routes.append(route_name)

    total_routes = len(unique_routes)

    # =====================================================
    # INITIATED / PASSED ROUTES
    # =====================================================
    # A route is considered INITIATED only when all report
    # entries belonging to that route have RESULT = PASS.
    #
    # Example:
    # TOTAL ROUTES          12
    # INITIATED ROUTES       5    1_A, 2_K, 3_K, 8_L, 25_J
    # NON INITIATED ROUTES   7
    # =====================================================

    route_results = {}

    for row_data in report_rows:

        route_name = str(
            row_data.get("ROUTE", "")
        ).strip().upper()

        if not route_name:
            continue

        result = str(
            row_data.get("RESULT", "")
        ).strip().upper()

        route_results.setdefault(
            route_name,
            []
        ).append(result)

    initiated_route_names = []

    for route_name in unique_routes:

        results = route_results.get(
            route_name,
            []
        )

        if results and all(
                result == "PASS"
                for result in results
        ):
            initiated_route_names.append(
                route_name
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

    start_row = ws.max_row + 3

    ws[f"A{start_row}"] = "TOTAL ROUTES"
    ws[f"B{start_row}"] = total_routes

    ws[f"A{start_row + 1}"] = "INITIATED ROUTES"
    ws[f"B{start_row + 1}"] = initiated_routes
    ws[f"C{start_row + 1}"] = initiated_route_text

    ws[f"A{start_row + 2}"] = "NON INITIATED ROUTES"
    ws[f"B{start_row + 2}"] = non_initiated_routes

    for r in range(start_row, start_row + 3):
        ws[f"A{r}"].font = Font(bold=True)
        ws[f"B{r}"].font = Font(bold=True)
        ws[f"C{r}"].font = Font(bold=True)

        ws[f"A{r}"].border = border
        ws[f"B{r}"].border = border
        ws[f"C{r}"].border = border

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

    # Make room for the comma-separated initiated route names.
    if initiated_route_text:
        ws.column_dimensions["C"].width = max(
            ws.column_dimensions["C"].width or 0,
            min(len(initiated_route_text) + 5, 100)
        )

    # =====================================================
    # CONSOLIDATED SUMMARY
    # =====================================================

    summary_ws = wb.create_sheet(
        "CONSOLIDATED SUMMARY"
    )

    summary_ws.merge_cells("A1:C1")

    summary_title = summary_ws["A1"]
    summary_title.value = "CONSOLIDATED TEST SUMMARY"
    summary_title.font = Font(bold=True, size=18)
    summary_title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    summary_ws.row_dimensions[1].height = 35

    summary_ws.append([])

    summary_ws.append([
        "SIGNAL",
        "TOTAL ROUTES",
        "FINAL RESULT"
    ])

    for cell in summary_ws[3]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        cell.border = border

    signal_summary = {}

    for row_data in report_rows:

        sig = row_data.get("SIGNAL", "")
        result = str(
            row_data.get("RESULT", "")
        ).upper()

        if sig not in signal_summary:
            signal_summary[sig] = {
                "total": 0,
                "failed": 0,
                "not_tested": 0
            }

        signal_summary[sig]["total"] += 1

        if result == "FAIL":
            signal_summary[sig]["failed"] += 1
        elif result == "NOT TESTED":
            signal_summary[sig]["not_tested"] += 1

    for sig, data in signal_summary.items():

        if data["failed"] > 0:

            failed_routes = []

            for r in report_rows:
                if (
                        r.get("SIGNAL") == sig
                        and str(r.get("RESULT", "")).upper() == "FAIL"
                ):
                    failed_routes.append(
                        f'{r.get("ROUTE", "")} '
                        f'({r.get("CRANK_HANDLE", "")}) NOT PASSED'
                    )

            final_result = ", ".join(failed_routes)

            row_fill = PatternFill(
                fill_type="solid",
                fgColor="FFC7CE"
            )
            row_font = Font(color="9C0006")

        elif data["not_tested"] > 0:

            final_result = "PARTIALLY TESTED"

            row_fill = PatternFill(
                fill_type="solid",
                fgColor="FFF2CC"
            )
            row_font = Font(color="7F6000")

        else:

            final_result = "ALL TESTS PASSED"

            row_fill = PatternFill(fill_type=None)
            row_font = Font(color="000000")

        summary_ws.append([
            sig,
            data["total"],
            final_result
        ])

        current_row = summary_ws.max_row

        for cell in summary_ws[current_row]:
            cell.fill = row_fill
            cell.font = row_font
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # =====================================================
    # CONSOLIDATED TOTALS
    # =====================================================

    summary_start = summary_ws.max_row + 3

    total = len(report_rows)

    passed = sum(
        1 for x in report_rows
        if str(x.get("RESULT", "")).upper() == "PASS"
    )

    failed = sum(
        1 for x in report_rows
        if str(x.get("RESULT", "")).upper() == "FAIL"
    )

    not_tested = sum(
        1 for x in report_rows
        if str(x.get("RESULT", "")).upper() == "NOT TESTED"
    )

    summary_ws[f"A{summary_start}"] = "TOTAL TESTS"
    summary_ws[f"B{summary_start}"] = total

    summary_ws[f"A{summary_start + 1}"] = "PASSED"
    summary_ws[f"B{summary_start + 1}"] = passed

    summary_ws[f"A{summary_start + 2}"] = "FAILED"
    summary_ws[f"B{summary_start + 2}"] = failed

    summary_ws[f"A{summary_start + 3}"] = "NOT TESTED"
    summary_ws[f"B{summary_start + 3}"] = not_tested

    for r in range(summary_start, summary_start + 4):
        summary_ws[f"A{r}"].font = Font(bold=True)
        summary_ws[f"B{r}"].font = Font(bold=True)
        summary_ws[f"A{r}"].border = border
        summary_ws[f"B{r}"].border = border

    summary_ws.column_dimensions["A"].width = 18
    summary_ws.column_dimensions["B"].width = 18
    summary_ws.column_dimensions["C"].width = 60

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(REPORT_FILE)

    log(
        f"REPORT SAVED : {REPORT_FILE}"
    )


# =========================================================
# RUN ENGINE
# =========================================================

# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():
    global running

    root.iconify()

    time.sleep(3)

    report_rows.clear()

    # =====================================================
    # VALIDATE ALL CRANK HANDLES (ONLY ONCE)
    # =====================================================

    log("=" * 80)
    log("STARTING CRANK HANDLE VALIDATION")
    log("=" * 80)

    if not validate_all_crank_handles():
        log("CRANK HANDLE VALIDATION FAILED")

        running = False

        status_label.config(
            text="FAILED",
            fg="#dc2626"
        )

        root.deiconify()

        return

    log("CRANK HANDLE VALIDATION PASSED")

    # =====================================================
    # START ROUTE TESTING
    # =====================================================

    for row in lock_routes_data:

        pause_event.wait()

        if not running:
            break

        signal = row["signal"]
        route = row["route"]
        crank_handle = row["crank_handle"]

        approach_track = row.get(
            "approach_track",
            ""
        )

        # ---------------------------------------------
        # Skip if no crank handle
        # ---------------------------------------------

        # ---------------------------------------------
        # GET CRANK HANDLE
        # ---------------------------------------------

        if crank_handle is None or str(crank_handle).strip() == "":

            log(
                f"TOC HAS NO CRANK HANDLE FOR "
                f"{signal} -> {route}"
            )

            # -------------------------------------------------
            # USE CRANK HANDLES LOADED FROM NEW SIGNAL CONFIG
            # -------------------------------------------------

            if crank_handle_points:

                crank_handle = ",".join(
                    sorted(crank_handle_points.keys())
                )

                log(
                    f"USING CONFIGURED CRANK HANDLES : "
                    f"{crank_handle}"
                )

            else:

                log(
                    f"SKIPPING {signal} -> {route} "
                    f"(NO CRANK HANDLE IN TOC OR CONFIG)"
                )

                continue

        log("=" * 80)
        log(f"SIGNAL             : {signal}")
        log(f"ROUTE              : {route}")
        log(f"CRANK HANDLE       : {crank_handle}")

        if get_signal_type(signal) == "CALLING_ON":
            log(
                f"APPROACH TRACK     : "
                f"{approach_track or '[EMPTY]'}"
            )

        # ---------------------------------------------
        # CH1 / CH2 / CH1,CH2
        # ---------------------------------------------

        # ---------------------------------------------
        # CH1 / CH2 / CH1,CH2
        # ---------------------------------------------

        # ---------------------------------------------
        # TEST EACH CRANK HANDLE ONE BY ONE
        # ---------------------------------------------

        handles = [
            h.strip().upper()
            for h in crank_handle.split(",")
            if h.strip()
        ]

        # -----------------------------------------
        # SINGLE CRANK HANDLE
        # -----------------------------------------

        if len(handles) == 1:

            handle = handles[0]

            log("=" * 70)
            log(f"STARTING TEST FOR {handle}")

            result = test_crank_handle(
                signal,
                route,
                handle,
                approach_track
            )

            if result == "PASS":
                reason = (
                    "ALL TESTS PASSED - "
                    "SIGNAL CANCELLED - "
                    "ROUTE RELEASED"
                )
            else:
                reason = (
                    last_fail_reason
                    if last_fail_reason
                    else "TEST FAILED"
                )

            report_rows.append({

                "SIGNAL": signal,
                "ROUTE": route,
                "CRANK_HANDLES": handle,
                "RESULT": result,
                "REASON": reason,
                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            })

        # -----------------------------------------
        # MULTIPLE CRANK HANDLES
        # -----------------------------------------

        else:

            log("=" * 70)
            log(f"STARTING MULTIPLE HANDLE TEST : {handles}")

            result = test_multiple_crank_handles(
                signal,
                route,
                handles,
                approach_track
            )

            first = True

            for handle in handles:

                if first:

                    report_result = result

                    first = False

                else:

                    if result == "FAIL":

                        report_result = "NOT TESTED"

                    else:

                        report_result = "PASS"

                if report_result == "PASS":
                    reason = (
                        "ALL TESTS PASSED - "
                        "SIGNAL CANCELLED - "
                        "ROUTE RELEASED"
                    )
                elif report_result == "FAIL":
                    reason = (
                        last_fail_reason
                        if last_fail_reason
                        else "TEST FAILED"
                    )
                else:
                    reason = (
                        "NOT TESTED - "
                        "PREVIOUS CRANK HANDLE TEST FAILED"
                    )

                report_rows.append({

                    "SIGNAL": signal,
                    "ROUTE": route,
                    "CRANK_HANDLES": handle,
                    "RESULT": report_result,
                    "REASON": reason,
                    "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

                })

    running = False

    status_label.config(

        text="COMPLETED",

        fg="#16a34a"

    )

    create_report()

    root.deiconify()

    log("=" * 80)
    log("ALL EMERGENCY CRANK HANDLE TESTS COMPLETED")
    log("=" * 80)

    try:

        os.startfile(REPORT_FILE)


    except Exception as e:

        log(str(e))


# =========================================================
# TEST SINGLE CRANK HANDLE
# =========================================================


# =========================================================
# CALLING-ON 50051 ROUTE PREPARATION
# =========================================================

def select_calling_on_50051_track(approach_track):
    track = str(approach_track or "").strip().upper()

    if not track:
        log("CALLING-ON -> APPROACH_TRACK NOT CONFIGURED")
        return False

    if not find_and_click(
            "Station 50051",
            track,
            "ListItem"
    ):
        log(
            f"CALLING-ON -> APPROACH TRACK "
            f"{track} NOT FOUND IN 50051"
        )
        return False

    log(
        f"CALLING-ON -> APPROACH TRACK "
        f"{track} SELECTED"
    )

    time.sleep(1)
    return True


def prepare_calling_on_50051(approach_track):
    """
    Calling-On route preparation:

        Ctrl+B -> 50051 -> ENTER
        -> Approach_Track -> Transmit -> Cancel
    """

    track = str(approach_track or "").strip().upper()

    if not track:
        log(
            "CALLING-ON -> 50051 PREPARATION FAILED: "
            "APPROACH_TRACK EMPTY"
        )
        return False

    log(
        f"CALLING-ON -> PREPARE 50051 "
        f"APPROACH_TRACK = {track}"
    )

    if not open_station_bits():
        log(
            "CALLING-ON -> FAILED TO OPEN STATION 50051"
        )
        return False

    if not select_calling_on_50051_track(track):

        try:
            click_station_cancel()
        except Exception:
            pyautogui.press("esc")

        return False

    if not click_station_transmit():

        log(
            "CALLING-ON -> 50051 TRANSMIT FAILED"
        )

        try:
            click_station_cancel()
        except Exception:
            pyautogui.press("esc")

        return False

    log(
        f"CALLING-ON -> 50051 {track} TRANSMITTED"
    )

    time.sleep(1)

    if not click_station_cancel():
        log(
            "CALLING-ON -> 50051 CANCEL FAILED"
        )

        pyautogui.press("esc")
        return False

    log(
        "CALLING-ON -> 50051 PREPARATION COMPLETED"
    )

    return True


def wait_for_calling_on_route_set(
        signal,
        timeout=30.0,
        interval=0.5
):
    if signal not in signals:
        log(
            f"{signal} -> SIGNAL NOT FOUND "
            "FOR CALLING-ON ROUTE CHECK"
        )
        return False

    route_init = signals[signal].get("ROUTE_INIT")

    if not route_init:
        log(
            f"{signal} -> CALLING-ON ROUTE_INIT "
            "POINT NOT CONFIGURED"
        )
        return False

    log(
        f"{signal} -> WAITING FOR CALLING-ON "
        f"ROUTE_INIT YELLOW AT "
        f"({route_init[0]}, {route_init[1]})"
    )

    end_time = time.time() + timeout

    while time.time() < end_time:

        pause_event.wait()

        if not running:
            return False

        if is_yellow(route_init):
            log(
                f"{signal} -> CALLING-ON "
                "ROUTE_INIT YELLOW DETECTED - ROUTE SET"
            )

            return True

        time.sleep(interval)

    log(
        f"{signal} -> CALLING-ON ROUTE_INIT "
        f"YELLOW NOT DETECTED WITHIN {timeout} SECONDS"
    )

    return False


def restore_calling_on_50051(approach_track):
    """Reload the Calling-On approach track in Station 50051."""

    if not approach_track:
        log(
            "CALLING-ON -> 50051 RESTORE SKIPPED: "
            "APPROACH_TRACK EMPTY"
        )
        return True

    return prepare_calling_on_50051(
        approach_track
    )


def test_crank_handle(signal, route, handle, approach_track=""):
    log("=" * 80)
    log(f"STARTING TEST : {handle}")

    # -------------------------------------------------
    # 1. SET ROUTE
    # -------------------------------------------------

    if signal not in signals:
        return fail_step(signal, handle, "SIGNAL NOT FOUND")

    signal_type = get_signal_type(signal)

    global active_calling_on_track
    global active_calling_on_test

    # Reset recovery state for this test.
    active_calling_on_track = ""
    active_calling_on_test = False

    if signal_type == "CALLING_ON":
        active_calling_on_track = str(approach_track or "").strip().upper()
        active_calling_on_test = True

    # Calling-On ONLY: prepare 50051 before setting the route.
    if signal_type == "CALLING_ON":

        if not prepare_calling_on_50051(
                approach_track
        ):
            return fail_step(
                signal,
                handle,
                "CALLING-ON 50051 APPROACH TRACK PREPARATION FAILED"
            )

    if not click(signals[signal]["menu"]):
        return fail_step(signal, handle, "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item(route.replace("-", "_")):
        return fail_step(signal, handle, "ROUTE SET FAILED")

    log("ROUTE COMMAND SENT")

    time.sleep(2)

    if signal_type == "CALLING_ON":

        if not wait_for_calling_on_route_set(
                signal,
                timeout=30
        ):
            log(
                f"{signal} -> CALLING-ON ROUTE DID NOT SET"
            )

            # Do not issue Signal Cancel / Route Release
            # when the Calling-On route never became SET.
            # fail_step() will restore 50051 / 1CXTPR.

            return fail_step(
                signal,
                handle,
                "CALLING-ON ROUTE DID NOT SET"
            )

        log(
            f"{signal} -> CALLING-ON ROUTE SET VERIFIED"
        )

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

    if signal_type == "CALLING_ON":

        if not restore_calling_on_50051(
                approach_track
        ):
            return fail_step(
                signal,
                handle,
                "50051 RESTORE FAILED"
            )

        log(
            f"{signal} -> CALLING-ON "
            "50051 RESTORE COMPLETED"
        )

        # Successful restore completed. Clear recovery state.
        active_calling_on_track = ""
        active_calling_on_test = False

    # -------------------------------------------------
    # 22. PASS
    # -------------------------------------------------

    log(f"{handle} PASSED")

    return "PASS"


# =========================================================
# TEST MULTIPLE CRANK HANDLES
# =========================================================

def test_multiple_crank_handles(signal, route, handles, approach_track=""):
    log("=" * 80)
    log(f"STARTING MULTIPLE CRANK HANDLE TEST : {', '.join(handles)}")

    # -------------------------------------------------
    # 1. SET ROUTE
    # -------------------------------------------------

    if signal not in signals:
        return fail_step(signal, ",".join(handles), "SIGNAL NOT FOUND")

    signal_type = get_signal_type(signal)

    global active_calling_on_track
    global active_calling_on_test

    # Reset recovery state for this test.
    active_calling_on_track = ""
    active_calling_on_test = False

    if signal_type == "CALLING_ON":
        active_calling_on_track = str(approach_track or "").strip().upper()
        active_calling_on_test = True

    # Calling-On ONLY: prepare 50051 before setting the route.
    if signal_type == "CALLING_ON":

        if not prepare_calling_on_50051(
                approach_track
        ):
            return fail_step(
                signal,
                ",".join(handles),
                "CALLING-ON 50051 APPROACH TRACK PREPARATION FAILED"
            )

    if not click(signals[signal]["menu"]):
        return fail_step(signal, ",".join(handles), "FAILED TO OPEN SIGNAL MENU")

    time.sleep(1)

    if not click_menu_item(route.replace("-", "_")):
        return fail_step(signal, ",".join(handles), "ROUTE SET FAILED")

    log("ROUTE COMMAND SENT")

    time.sleep(2)

    if signal_type == "CALLING_ON":

        if not wait_for_calling_on_route_set(
                signal,
                timeout=30
        ):
            log(
                f"{signal} -> CALLING-ON ROUTE DID NOT SET"
            )

            # fail_step() will restore 50051 / 1CXTPR.

            return fail_step(
                signal,
                ",".join(handles),
                "CALLING-ON ROUTE DID NOT SET"
            )

        log(
            f"{signal} -> CALLING-ON ROUTE SET VERIFIED"
        )

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

    if signal_type == "CALLING_ON":

        if not restore_calling_on_50051(
                approach_track
        ):
            return fail_step(
                signal,
                ",".join(handles),
                "50051 RESTORE FAILED"
            )

        log(
            f"{signal} -> CALLING-ON "
            "50051 RESTORE COMPLETED"
        )

        # Successful restore completed. Clear recovery state.
        active_calling_on_track = ""
        active_calling_on_test = False

    return "PASS"


# =========================================================
# START
# =========================================================

def start_automation():
    global running
    global REPORT_FILE
    global last_fail_reason
    global active_calling_on_track
    global active_calling_on_test

    log("=" * 70)
    log("START AUTOMATION REQUESTED")
    log(f"SIGNALS LOADED      : {len(signals)}")
    log(f"TOC ROWS LOADED      : {len(lock_routes_data)}")
    log(f"CRANK HANDLES LOADED : {len(crank_handle_points)}")
    log("=" * 70)

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
    # UNIQUE REPORT FOR THIS TEST RUN
    # =====================================================

    REPORT_FILE = (
            "ECH_REPORT_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
            + ".xlsx"
    )

    report_rows.clear()
    last_fail_reason = ""
    active_calling_on_track = ""
    active_calling_on_test = False

    log(
        f"NEW TEST RUN REPORT : {REPORT_FILE}"
    )

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
    "EMERGENCY CRANK HANDLING AUTOMATION SYSTEM"
)

root.geometry("1500x950")

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
    rowheight=32,
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

style.configure(
    "TNotebook.Tab",
    font=("Segoe UI", 10, "bold"),
    padding=[14, 6]
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

    text="EMERGENCY CRANK HANDLING AUTOMATION SYSTEM",

    font=("Segoe UI", 22, "bold"),

    bg="#0f172a",

    fg="white"

).pack(
    pady=(15, 0)
)

tk.Label(

    title_frame,

    text="COORDINATE CAPTURE  |  ROUTE TESTING  |  REPORTING",

    font=("Segoe UI", 10),

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

table_title = tk.Label(

    right_panel,

    text="EMERGENCY CRANK HANDLE DETAILS",

    font=("Segoe UI", 16, "bold"),

    bg="#e9edf2",

    fg="#0f172a"

)

table_title.pack(
    anchor="w",
    pady=(0, 10)
)

# The tab caption below replaces this legacy heading, as in NEW_UI.py.
table_title.pack_forget()

# The data view follows the tabbed layout used in NEW_UI.py.
notebook = ttk.Notebook(right_panel)
notebook.pack(
    fill="both",
    expand=True
)

# =========================================================
# SIGNAL TYPE TABS
# =========================================================

# Dashboard shows ONLY:
#   1. MAIN SIGNALS
#   2. CALLING ON SIGNALS
#   3. SHUNT SIGNALS
#
# Emergency Crank Handle, Point and LC Gate data are still captured,
# saved and used by the automation. They are NOT shown as dashboard tabs.

signal_notebook = ttk.Notebook(right_panel)

signal_notebook.pack(
    fill="both",
    expand=True
)

# IMPORTANT:
# Only these three signal tabs are created.
signal_tabs = {
    "MAIN": tk.Frame(
        signal_notebook,
        bg="white",
        bd=1,
        relief="solid"
    ),
    "CALLING_ON": tk.Frame(
        signal_notebook,
        bg="white",
        bd=1,
        relief="solid"
    ),
    "SHUNT": tk.Frame(
        signal_notebook,
        bg="white",
        bd=1,
        relief="solid"
    )
}

signal_notebook.add(
    signal_tabs["MAIN"],
    text="MAIN SIGNALS"
)

signal_notebook.add(
    signal_tabs["CALLING_ON"],
    text="CALLING ON SIGNALS"
)

signal_notebook.add(
    signal_tabs["SHUNT"],
    text="SHUNT SIGNALS"
)

# Keep the original variable name for compatibility.
notebook = signal_notebook

# =========================================================
# CREATE TABLES FOR THE THREE SIGNAL TABS ONLY
# =========================================================

signal_trees = {}

for signal_type, tab_frame in signal_tabs.items():
    tab_scroll = ttk.Scrollbar(tab_frame)

    tab_scroll.pack(
        side="right",
        fill="y"
    )

    tab_tree = ttk.Treeview(
        tab_frame,
        columns=(
            "NO",
            "SIGNAL",
            "ROUTE",
            "CRANK_HANDLE",
        ),
        show="headings",
        yscrollcommand=tab_scroll.set
    )

    tab_scroll.config(
        command=tab_tree.yview
    )

    tab_tree.heading("NO", text="NO")
    tab_tree.heading("SIGNAL", text="SIGNAL")
    tab_tree.heading("ROUTE", text="ROUTE")
    tab_tree.heading("CRANK_HANDLE", text="CRANK HANDLE")

    tab_tree.column(
        "NO",
        width=80,
        anchor="center"
    )

    tab_tree.column(
        "SIGNAL",
        width=220,
        anchor="center"
    )

    tab_tree.column(
        "ROUTE",
        width=300,
        anchor="center"
    )

    tab_tree.column(
        "CRANK_HANDLE",
        width=220,
        anchor="center"
    )

    tab_tree.pack(
        fill="both",
        expand=True
    )

    signal_trees[signal_type] = tab_tree

# Original "tree" reference points to the MAIN table for
# compatibility with existing code.
tree = signal_trees["MAIN"]

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

    text="SPACE = CAPTURE COORDINATES  |  P = PAUSE / RESUME  |  EMERGENCY CRANK HANDLING AUTOMATION SYSTEM",

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

create_capture_window()
# =========================================================
# RUN
# =========================================================

root.mainloop()

