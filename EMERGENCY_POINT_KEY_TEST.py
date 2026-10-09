# =====================================================================
# HITACHI C-FAT AUTOMATION SYSTEM
# =====================================================================
#
# Project :
#     Hitachi Computer Based Interlocking (CBI)
#     C-FAT Automation Tool
#
# Developed For :
#     Railway Signalling Automation
#
# Purpose :
#     Automate C-FAT testing using
#     • TOC.xlsx
#     • Hitachi Test Panel
#     • Hitachi Simulator
#
# Current Development Stage :
#     Phase 1 - Route Engine
#
# Future Modules :
#     ✓ Route Engine
#     ✓ Track Lock Test
#     ✓ Point Operation Test
#     ✓ Crank Handle Test
#     ✓ Emergency Point Operation
#     ✓ Route Lock Test
#     ✓ Auto Throw Test
#     ✓ Report Generation
#
# =====================================================================
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
import re
import subprocess
import pyautogui

import uiautomation as auto

from pywinauto import Desktop

from PIL import Image, ImageTk, ImageGrab
import ctypes
from openpyxl import Workbook
from openpyxl import load_workbook

SEARCH_TIMEOUT = 5
SEARCH_RETRY = 0.25
BIT_CHART_TIMEOUT = 5
from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "HAH COORDINATES1.xlsx"

REPORT_FILE = "SRINI.xlsx"

# =========================================================
# APPLICATION INFORMATION
# =========================================================

APP_NAME = "Hitachi C-FAT Automation"

VERSION = "1.0.0"

DEVELOPER = "Thirumalai Srinivas"
# =========================================================
# DEVELOPMENT SETTINGS
# =========================================================

DEBUG_MODE = True

CURRENT_TEST = "POINT/CRANK HANDLE"

TOC_FILE = None

SIMULATOR_RUNNING = False

PANEL_RUNNING = False
# =========================================================
# GLOBALS
# =========================================================

signals = {}

auto_throw_data = []
loaded_signal_list = []
loaded_point_list = []
loaded_ch_list = []
loaded_shunt_list = []
loaded_calling_on_list = []
loaded_lc_list = []
loaded_track_list = []
points = {}
running = False
EDRC_UNATTENDED = os.environ.get("EDRC_UNATTENDED") == "1"



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


from tkinter import filedialog
from openpyxl import load_workbook


# =========================================================
# NEW TOC HEADER HELPERS
# =========================================================
# The current TOC is header-driven.  Never depend on fixed Excel positions
# for Track / Points / CRANK_HANDLE because the new TOC contains additional
# columns between them.

def _norm_toc_header(value):
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "_", str(value).strip().upper()).strip("_")


def _get_toc_headers(ws):
    header_map = {}
    first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())

    for index, value in enumerate(first_row):
        key = _norm_toc_header(value)
        if key and key not in header_map:
            header_map[key] = index

    return header_map


def _toc_col(headers, *names):
    for name in names:
        key = _norm_toc_header(name)
        if key in headers:
            return headers[key]
    return None


def _toc_value(row, headers, *names):
    col = _toc_col(headers, *names)
    if col is None or col >= len(row):
        return None
    return row[col]


def _split_toc_list(value):
    """Split a TOC cell into clean comma/space separated items."""
    if value is None:
        return []

    text = str(value).strip()
    if not text:
        return []

    # The new TOC uses comma-separated values.  Keep the existing handling
    # for workbooks where multiple spaces were used as separators.
    result = [item.strip() for item in re.split(r"[,;]", text) if item.strip()]

    if len(result) == 1 and "  " in result[0]:
        result = [
            item.strip()
            for item in re.split(r"\s{2,}", result[0])
            if item.strip()
        ]

    return result


def _first_route_rows(ws):
    """Return only the first TOC route for each signal."""
    headers = _get_toc_headers(ws)

    signal_col = _toc_col(headers, "Signal")
    route_col = _toc_col(headers, "Route")

    if signal_col is None or route_col is None:
        raise ValueError("TOC must contain Signal and Route columns.")

    rows = []
    seen_signals = set()

    for excel_row_no, row in enumerate(
        ws.iter_rows(min_row=2, values_only=True),
        start=2
    ):
        signal = _toc_value(row, headers, "Signal")
        route = _toc_value(row, headers, "Route")

        if signal is None and route is None:
            continue

        signal_text = str(signal).strip() if signal is not None else ""
        route_text = str(route).strip() if route is not None else ""

        if not signal_text or not route_text:
            continue

        signal_key = signal_text.upper()
        if signal_key in seen_signals:
            continue

        seen_signals.add(signal_key)
        rows.append((excel_row_no, row))

    return headers, rows


def _find_selected_toc_row(ws, signal, route):
    """Find the exact GUI-selected route in the new header-driven TOC."""
    headers = _get_toc_headers(ws)
    signal_col = _toc_col(headers, "Signal")
    route_col = _toc_col(headers, "Route")

    if signal_col is None or route_col is None:
        return headers, None, None

    signal_key = str(signal or "").strip().upper()
    route_key = str(route or "").strip().upper()
    seen_signals = set()

    for excel_row_no, row in enumerate(
        ws.iter_rows(min_row=2, values_only=True),
        start=2
    ):
        row_signal = _toc_value(row, headers, "Signal")
        row_route = _toc_value(row, headers, "Route")

        if row_signal is None or row_route is None:
            continue

        row_signal_key = str(row_signal).strip().upper()
        row_route_key = str(row_route).strip().upper()

        # Keep the same first-route-per-signal rule used by the GUI.
        if row_signal_key in seen_signals:
            continue
        seen_signals.add(row_signal_key)

        if row_signal_key == signal_key and row_route_key == route_key:
            return headers, excel_row_no, row

    return headers, None, None


def load_toc():
    """
    Load the new header-driven TOC.

    IMPORTANT:
        Signal / Route / Lock_Route / Track / Points / CRANK_HANDLE are read
        by column name.  Only the FIRST route for each signal is displayed.
    """
    global TOC_FILE

    file_path = None
    wb = None

    try:
        log("--------------------------------")
        log("OPENING TOC FILE SELECTOR")
        log("--------------------------------")

        file_path = filedialog.askopenfilename(
            parent=root,
            title="Select TOC Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xlsm"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            log("TOC selection cancelled")
            return

        file_path = os.path.abspath(file_path).strip()
        log(f"Selected TOC : {file_path}")

        if not file_path.lower().endswith((".xlsx", ".xlsm")):
            messagebox.showerror("TOC Error", "Please select an Excel .xlsx or .xlsm file.")
            return

        if not os.path.isfile(file_path):
            messagebox.showerror("TOC Error", "The selected TOC file does not exist.")
            return

        wb = load_workbook(file_path, data_only=True, read_only=True)

        toc_sheet_name = next(
            (name for name in wb.sheetnames if str(name).strip().upper() == "TOC"),
            None
        )

        if toc_sheet_name is None:
            available = ", ".join(str(x) for x in wb.sheetnames)
            messagebox.showerror(
                "TOC Error",
                "The selected Excel file does not contain a TOC sheet.\n\n"
                f"Available sheets:\n{available}"
            )
            return

        ws = wb[toc_sheet_name]
        headers, first_rows = _first_route_rows(ws)

        required_columns = {
            "Signal": _toc_col(headers, "Signal"),
            "Route": _toc_col(headers, "Route"),
            "Track": _toc_col(headers, "Track"),
            "Points": _toc_col(headers, "Points", "Point"),
            "CRANK_HANDLE": _toc_col(headers, "CRANK_HANDLE", "Crank Handle", "CH"),
        }

        missing = [name for name, col in required_columns.items() if col is None]
        if missing:
            messagebox.showerror(
                "TOC Error",
                "The new TOC is missing required columns:\n\n" +
                "\n".join(missing)
            )
            log("TOC ERROR : Missing columns : " + ", ".join(missing))
            return

        if not first_rows:
            messagebox.showwarning("TOC Empty", "The selected TOC contains no valid first-route records.")
            return

        tree.delete(*tree.get_children())

        for index, (_excel_row_no, row) in enumerate(first_rows, start=1):
            signal = _toc_value(row, headers, "Signal") or ""
            route = _toc_value(row, headers, "Route") or ""
            lock_route = _toc_value(row, headers, "Lock_Route", "Lock Route", "Lock_Routes") or ""
            points = _toc_value(row, headers, "Points", "Point") or ""
            crank_handle = _toc_value(row, headers, "CRANK_HANDLE", "Crank Handle", "CH") or ""

            tree.insert(
                "",
                "end",
                values=(index, signal, route, lock_route, points, crank_handle)
            )

        TOC_FILE = file_path

        log("--------------------------------")
        log("TOC LOADED SUCCESSFULLY")
        log(f"TOC File       : {TOC_FILE}")
        log(f"TOC Sheet      : {toc_sheet_name}")
        log(f"FIRST ROUTE PER SIGNAL : {len(first_rows)}")
        log("TOC ACTIVE COLUMNS : Signal, Route, Lock_Route, Track, Points, CRANK_HANDLE")
        log("--------------------------------")

        try:
            status_label.config(text="TOC LOADED", fg="#16a34a")
        except Exception:
            pass

        messagebox.showinfo(
            "TOC Loaded",
            f"TOC loaded successfully.\n\nFirst route per signal : {len(first_rows)}"
        )

    except PermissionError as e:
        log(f"TOC ERROR : PermissionError : {e}")
        messagebox.showerror("TOC Error", "The TOC file is being used by another program or access is denied.")

    except Exception as e:
        log(f"TOC ERROR : {type(e).__name__} : {e}")
        messagebox.showerror("TOC Loading Error", f"Unable to load the selected TOC.\n\n{type(e).__name__}: {e}")

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def log(msg):
    """Write to the child GUI and also to stdout for the EDRC helper."""
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"

    # EDRC captures stdout from the helper process.  This makes startup
    # failures visible instead of appearing as a generic 6-second timeout.
    try:
        print(line, flush=True)
    except Exception:
        pass

    def write():
        try:
            log_text.config(state="normal")
            log_text.insert(tk.END, line + "\n")
            log_text.see(tk.END)
            log_text.config(state="disabled")
        except Exception:
            pass

    try:
        root.after(0, write)
    except Exception:
        pass

# =========================================================
# CLICK
# =========================================================

def click(point):
    """
    Mouse click helper using pyautogui only.
    No Windows API mouse calls are used.
    """
    if point is None:
        return False

    x, y = point

    try:
        pyautogui.moveTo(x, y, duration=0.25)
        pause_sleep(0.4)
        pause_event.wait()

        pyautogui.mouseDown(button="left")
        pause_sleep(0.1)
        pyautogui.mouseUp(button="left")
        pause_sleep(0.4)

        return True

    except Exception as e:
        log(f"CLICK ERROR : {e}")
        return False


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

    global loaded_point_list

    loaded_point_list.clear()

    all_names = set()

    n = prompt_count("Points")

    for i in range(n):

        name = prompt_name(
            "Point",
            i,
            n,
            all_names
        )

        all_names.add(name)

        loaded_point_list.append(name)

    log(f"Total Points : {len(loaded_point_list)}")

    for point in loaded_point_list:

        log(f"POINT : {point}")

    return loaded_point_list


def build_ch_capture_list():
    global loaded_ch_list

    loaded_ch_list.clear()

    all_names = set()

    n = prompt_count("Crank Handles")

    for i in range(n):
        name = prompt_name(
            "Crank Handle",
            i,
            n,
            all_names
        )

        all_names.add(name)

        loaded_ch_list.append(name)

    log(f"Total Crank Handles : {len(loaded_ch_list)}")

    for ch in loaded_ch_list:
        log(f"CH : {ch}")

    return loaded_ch_list
# =========================================================
# CAPTURE POINT
# =========================================================


TYPE_LABELS = {
    "POINT": "POINT COORDINATE"
}

TYPE_COLORS = {
    "POINT": "#3b82f6"
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
    capture_overlay.geometry(f"450x360+{screen_w - 470}+40")

    capture_overlay.protocol("WM_DELETE_WINDOW", capture_cancel)

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
        capture_overlay,
        text="",
        font=("Segoe UI", 22, "bold"),
        bg="#0f172a",
        fg="white",
        justify="center"
    )
    overlay_signal_label.pack()

    overlay_step_label = tk.Label(
        capture_overlay,
        text="",
        font=("Segoe UI", 18, "bold"),
        bg="#0f172a",
        fg="#22c55e",
        justify="center"
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
        capture_overlay, text="\u21b6  UNDO LAST CLICK", command=undo_capture,
        bg="#f59e0b", fg="#1a1a1a", activebackground="#fbbf24",
        activeforeground="#1a1a1a", relief="flat", cursor="hand2",
        font=("Segoe UI", 10, "bold"), width=22, height=1, takefocus=0
    ).pack(pady=(10, 2))

    tk.Label(
        capture_overlay, text="(or press BACKSPACE)",
        font=("Segoe UI", 8), bg="#0f172a", fg="#475569"
    ).pack()

    tk.Button(
        capture_overlay, text="CANCEL CAPTURE", command=capture_cancel,
        bg="#0f172a", fg="#64748b", activebackground="#0f172a",
        activeforeground="#ef4444", relief="flat", cursor="hand2",
        font=("Segoe UI", 9, "underline"), takefocus=0
    ).pack(side="bottom", pady=(0, 10))


capture_waiting = False
captured_point = None

capture_result = None
capture_action = None

def capture_point(title, step):

    global capture_result
    global capture_action

    capture_result = None
    capture_action = None

    overlay_signal_label.config(
        text=title.upper()
    )

    capture_overlay.title(title.upper())
    overlay_step_label.config(text=step)
    overlay_progress_label.config(
        text="SPACE = Capture    BACKSPACE = Undo    ESC = Cancel"
    )
    overlay_hint_label.config(
        text="Move the mouse to the required coordinate\nPress SPACE to capture"
    )

    capture_overlay.withdraw()

    capture_overlay.deiconify()
    capture_overlay.lift()
    capture_overlay.focus_force()

    while capture_action is None:
        root.update()
        time.sleep(0.02)

    capture_overlay.withdraw()

    return capture_action, capture_result


def capture_space():

    global capture_action
    global capture_result

    if not capture_overlay.winfo_viewable():
        return
    print("SPACE pressed")
    capture_result = pyautogui.position()
    capture_action = "SPACE"


def undo_capture():

    global capture_action

    if not capture_overlay.winfo_viewable():
        return

    capture_action = "UNDO"

def capture_cancel():

    global capture_action

    if not capture_overlay.winfo_viewable():
        return

    capture_action = "CANCEL"


# =========================================================
# REPORT STYLES
# =========================================================

FAIL_FILL = PatternFill(
    fill_type="solid",
    fgColor="FFC7CE"
)

FAIL_FONT = Font(
    bold=True,
    color="9C0006"
)

CENTER = Alignment(
    horizontal="center",
    vertical="center"
)

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)


from tkinter import simpledialog
#===========================================================
# CREATE REPORT
#===========================================================
from openpyxl import Workbook
from datetime import datetime

report_rows = []


def create_report():
    """Create the report file for this run (cascading report layout)."""

    global REPORT_FILE
    global report_rows

    # Report goes to the EDRC session folder when the launcher gives one,
    # otherwise to the working folder - never a "Reports" sub-folder,
    # where the launcher does not look.
    folder = (
        os.environ.get("EDRC_PROGRAM_DIR")
        or os.environ.get("EDRC_REPORT_DIR")
        or os.getcwd()
    )

    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)

    log(f"REPORT DESTINATION : {folder}")

    REPORT_FILE = os.path.join(
        folder,
        f"EMERGENCY_POINT_KEY_REPORT_"
        f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    )

    report_rows = []

    write_report()

    log(f"REPORT FILE CREATED : {REPORT_FILE}")


def write_report():
    """Write the whole report from report_rows."""

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "EMERGENCY POINT KEY REPORT"

    headers = [
        "POINT",
        "CRANK HANDLE",
        "POINT LOCK",
        "EMERGENCY OPERATION",
        "RESULT",
        "DATE & TIME"
    ]

    last_col = get_column_letter(len(headers))

    # TITLE
    ws.merge_cells(f"A1:{last_col}1")
    title = ws["A1"]
    title.value = "EMERGENCY POINT KEY TEST REPORT"
    title.font = Font(bold=True, size=18, color="000000")
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # DATE
    ws.merge_cells(f"A2:{last_col}2")
    date_cell = ws["A2"]
    date_cell.value = (
        f"Generated : {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
    )
    date_cell.font = Font(bold=True, size=11)
    date_cell.alignment = Alignment(horizontal="center")

    # HEADERS
    ws.append(headers)

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[3]:
        cell.font = Font(bold=True, color="000000")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # DATA
    for row in report_rows:
        ws.append([
            row.get("POINT", ""),
            row.get("CRANK HANDLE", ""),
            row.get("POINT LOCK", ""),
            row.get("EMERGENCY OPERATION", ""),
            row.get("RESULT", ""),
            row.get("DATE & TIME", "")
        ])

    for row in ws.iter_rows(min_row=4):

        result_cell = row[4]

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
                cell.font = Font(color="FFFFFF", bold=True)

    # AUTO WIDTH
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

    # SUMMARY
    passed = sorted(set(
        str(x.get("POINT", ""))
        for x in report_rows
        if str(x.get("RESULT", "")).strip().upper() == "PASS"
    ))

    failed = sorted(set(
        str(x.get("POINT", ""))
        for x in report_rows
        if str(x.get("RESULT", "")).strip().upper() != "PASS"
    ))

    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL POINTS TESTED"
    ws[f"B{start}"] = len(report_rows)

    ws[f"A{start + 1}"] = "PASSED POINTS"
    ws[f"B{start + 1}"] = ", ".join(passed) if passed else "0"

    ws[f"A{start + 2}"] = "FAILED POINTS"
    ws[f"B{start + 2}"] = ", ".join(failed) if failed else "0"

    for r in range(start, start + 3):
        ws[f"A{r}"].font = Font(bold=True)

    wb.save(REPORT_FILE)
    wb.close()


#============================================================
#SAVE RESULT
#============================================================

from openpyxl import load_workbook


def save_result(
        point,
        crank_handle,
        point_lock,
        emergency_operation
):
    """Add one point's result and rewrite the report."""

    global report_rows

    # A skipped check (no track in the TOC) must not fail the point.
    checks = [crank_handle, point_lock, emergency_operation]

    if any(str(c).strip().upper() == "FAIL" for c in checks):
        result = "FAIL"
    else:
        result = "PASS"

    report_rows.append({
        "POINT": point,
        "CRANK HANDLE": crank_handle,
        "POINT LOCK": point_lock,
        "EMERGENCY OPERATION": emergency_operation,
        "RESULT": result,
        "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    })

    try:
        write_report()
        log(f"REPORT UPDATED : POINT {point} -> {result}")
    except Exception as e:
        log(f"REPORT UPDATE ERROR : {type(e).__name__} : {e}")


#===========================================================
#FINALIZE REPORT
#===========================================================
def finalize_report():

    try:
        write_report()
    except Exception as e:
        log(f"REPORT FINALIZE ERROR : {e}")
        return

    log(f"REPORT SAVED : {REPORT_FILE}")

    if EDRC_UNATTENDED:
        log("REPORT SAVED - EDRC WILL HANDLE SESSION REPORT DISPLAY")
    else:
        try:
            os.startfile(REPORT_FILE)
        except Exception as e:
            log(f"REPORT OPEN ERROR : {e}")


# =========================================================
# KEYBOARD
# =========================================================


keyboard.add_hotkey("space", capture_space)
keyboard.add_hotkey("backspace", undo_capture)
keyboard.add_hotkey("esc", capture_cancel)

# =========================================================
# CONFIGURATION MIGRATION
# =========================================================

def _ensure_point_sheet_schema(wb):
    """Upgrade older POINT sheets without destroying indication data."""
    if "POINT" not in wb.sheetnames:
        ws = wb.create_sheet("POINT")
    else:
        ws = wb["POINT"]

    headers = [
        "Point", "Menu_X", "Menu_Y",
        "Normal_X", "Normal_Y",
        "Reverse_X", "Reverse_Y",
        "Free_X", "Free_Y",
        "Control_B_X", "Control_B_Y"
    ]

    existing = [
        str(ws.cell(1, c).value or "").strip().upper()
        for c in range(1, max(ws.max_column, len(headers)) + 1)
    ]

    # If this is a brand-new/empty sheet, create the full schema.
    if ws.max_row == 1 and all(
        ws.cell(1, c).value in (None, "")
        for c in range(1, len(headers) + 1)
    ):
        for c, header in enumerate(headers, 1):
            ws.cell(1, c).value = header
        return ws

    # Normalize the first nine standard headers.
    for c, header in enumerate(headers, 1):
        if ws.cell(1, c).value in (None, ""):
            ws.cell(1, c).value = header

    # Older workbooks had only A:I. J:K are new Control-B columns.
    if str(ws.cell(1, 10).value or "").strip() == "":
        ws.cell(1, 10).value = "Control_B_X"
    if str(ws.cell(1, 11).value or "").strip() == "":
        ws.cell(1, 11).value = "Control_B_Y"

    return ws


# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():

    global CONFIG_FILE

    file = filedialog.asksaveasfilename(
        title="Create / Open Configuration",
        defaultextension=".xlsx",
        initialfile="HAH COORDINATES1.xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    CONFIG_FILE = file

    # -------------------------------------------------
    # OPEN EXISTING OR CREATE NEW
    # -------------------------------------------------

    if os.path.exists(CONFIG_FILE):

        wb = load_workbook(CONFIG_FILE)

        log("Existing Configuration Opened")

    else:

        wb = Workbook()

        log("New Configuration Created")

    # -------------------------------------------------
    # POINT SHEET
    # -------------------------------------------------

    _ensure_point_sheet_schema(wb)

    if False:  # retained structure marker; schema was created above


        if wb.active.title == "Sheet":

            ws = wb.active
            ws.title = "POINT"

        else:

            ws = wb.create_sheet("POINT")

        ws.append([
            "Point",
            "Menu_X",
            "Menu_Y",
            "Normal_X",
            "Normal_Y",
            "Reverse_X",
            "Reverse_Y",
            "Free_X",
            "Free_Y",
            "Control_B_X",
            "Control_B_Y"
        ])

    # -------------------------------------------------
    # CH SHEET
    # -------------------------------------------------

    if "CH" not in wb.sheetnames:

        ws = wb.create_sheet("CH")

        ws.append([
            "POINT",
            "CONTROL_X",
            "CONTROL_Y",
            "IN_X",
            "IN_Y",
            "OUT_X",
            "OUT_Y",
            "FREE_X",
            "FREE_Y"
        ])

    # -------------------------------------------------
    # TRACK CONFIG
    # -------------------------------------------------

    if "TRACK_CONFIG" not in wb.sheetnames:
        ws = wb.create_sheet("TRACK_CONFIG")
        # Track | X | Y.  This sheet intentionally has no header because
        # the existing workflow reads it as a simple 3-column map.

    # -------------------------------------------------
    # SYSTEM CONTROLS
    # -------------------------------------------------

    if "SYSTEM_CONTROLS" not in wb.sheetnames:

        ws = wb.create_sheet("SYSTEM_CONTROLS")

        ws.append([
            "CONTROL",
            "X",
            "Y"
        ])

    ws = wb["SYSTEM_CONTROLS"]

    found = False

    for row in ws.iter_rows(min_row=2):

        if row[0].value == "EMERGENCY POINT KEY":

            found = True
            break

    if not found:

        ws.append([
            "EMERGENCY POINT KEY",
            None,
            None
        ])

    # -------------------------------------------------
    # AUTO WIDTH
    # -------------------------------------------------

    for sheet in wb.worksheets:

        for cell in sheet[1]:

            cell.font = Font(bold=True)

            cell.alignment = CENTER

            cell.border = THIN_BORDER

        for column in sheet.columns:

            length = max(
                len(str(c.value or ""))
                for c in column
            )

            sheet.column_dimensions[
                get_column_letter(column[0].column)
            ].width = length + 5

    wb.save(CONFIG_FILE)

    wb.close()

    log(f"Configuration Ready : {CONFIG_FILE}")

    messagebox.showinfo(
        "Configuration",
        "Configuration is ready.\n\n"
        "Existing coordinates were preserved."
    )
# =========================================================
# LOAD CONFIG
# =========================================================

def load_config():
    """
    Load the HAH coordinate workbook used by the point/crank-handle/
    emergency-point-key workflow.

    Coordinate mapping comes from:
        MAIN             -> signal/menu/aspect coordinates
        POINT            -> point menu + normal/reverse/free indication
        CH               -> crank-handle menu + IN/OUT/FREE indication
        TRACK_CONFIG     -> track coordinate mapping
        SYSTEM_CONTROLS  -> Emergency Point Key coordinate

    The workbook is selected by the user; no coordinates are hard-coded.
    """
    global CONFIG_FILE

    try:
        log("--------------------------------")
        log("LOAD HAH COORDINATE CONFIGURATION")
        log("--------------------------------")

        file_path = filedialog.askopenfilename(
            parent=root,
            title="Select HAH COORDINATES Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xlsm"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            log("Configuration selection cancelled")
            return

        wb = load_workbook(file_path, data_only=True, read_only=True)
        sheets = {str(s).strip().upper(): s for s in wb.sheetnames}

        required = ["POINT", "CH", "TRACK_CONFIG", "SYSTEM_CONTROLS"]
        missing = [s for s in required if s not in sheets]

        if missing:
            wb.close()
            messagebox.showerror(
                "Invalid HAH Configuration",
                "The selected workbook is missing:\n\n" +
                "\n".join(missing) +
                "\n\nPlease select HAH COORDINATES1.xlsx."
            )
            log(f"CONFIG ERROR : Missing sheets {missing}")
            return

        # Validate the coordinate headers so a different workbook cannot
        # silently be used with this workflow.
        checks = {
            "POINT": ["Point", "Menu_X", "Menu_Y", "Normal_X", "Normal_Y",
                      "Reverse_X", "Reverse_Y", "Free_X", "Free_Y",
                      "Control_B_X", "Control_B_Y"],
            "CH": ["Point", "Menu_X", "Menu_Y", "IN_X", "IN_Y",
                   "OUT_X", "OUT_Y", "FREE_X", "FREE_Y"],
            "SYSTEM_CONTROLS": ["CONTROL", "X", "Y"]
        }

        for sheet_key, expected in checks.items():
            ws = wb[sheets[sheet_key]]
            headers = [
                str(c).strip().upper() if c is not None else ""
                for c in next(ws.iter_rows(min_row=1, max_row=1,
                                           values_only=True))
            ]
            missing_headers = [
                h for h in expected if h.upper() not in headers
            ]
            if missing_headers:
                wb.close()
                messagebox.showerror(
                    "Invalid HAH Configuration",
                    f"{sheet_key} sheet is missing:\n\n" +
                    "\n".join(missing_headers)
                )
                log(
                    f"CONFIG ERROR : {sheet_key} missing headers "
                    f"{missing_headers}"
                )
                return

        wb.close()

        CONFIG_FILE = os.path.abspath(file_path)

        log("--------------------------------")
        log("HAH CONFIGURATION LOADED")
        log(f"CONFIG FILE : {CONFIG_FILE}")
        log("Coordinate source : Excel workbook")
        log("POINT -> POINT sheet")
        log("CH -> CH sheet")
        log("TRACK -> TRACK_CONFIG sheet")
        log("EMERGENCY KEY -> SYSTEM_CONTROLS sheet")
        log("--------------------------------")

        try:
            status_label.config(text="CONFIG LOADED", fg="#16a34a")
        except Exception:
            pass

        messagebox.showinfo(
            "Configuration Loaded",
            "HAH COORDINATES configuration loaded successfully."
        )

    except Exception as e:
        log(f"CONFIG ERROR : {type(e).__name__} : {e}")
        messagebox.showerror(
            "Configuration Error",
            f"Unable to load HAH coordinate configuration.\n\n{e}"
        )


def record_point_indication_coordinate():

    log("===================================")
    log("POINT INDICATION COORDINATE RECORDER")
    log("===================================")

    build_quantity_capture_list()

    if not loaded_point_list:
        log("No Points To Record")
        return

    wb = load_workbook(CONFIG_FILE)

    for point in loaded_point_list:

        coords = [None, None]

        step = 0

        while step < 2:

            if step == 0:

                action, result = capture_point(
                    f"POINT {point}",
                    "RECORD NORMAL INDICATION"
                )

            else:

                action, result = capture_point(
                    f"POINT {point}",
                    "RECORD REVERSE INDICATION"
                )

            if action == "CANCEL":
                wb.close()
                return

            if action == "UNDO":

                if step > 0:
                    coords[step - 1] = None
                    step -= 1

                continue

            coords[step] = result
            step += 1

        save_point_indication_coordinate(
            point,
            coords[0][0],
            coords[0][1],
            coords[1][0],
            coords[1][1]
        )

        log(f"{point} NORMAL : {coords[0]}")
        log(f"{point} REVERSE : {coords[1]}")

    wb.close()

    log("Point Indication Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Point Indication Coordinates Saved Successfully."
    )

def save_point_indication_coordinate(
        point,
        n_x,
        n_y,
        r_x,
        r_y,
):
    wb = load_workbook(CONFIG_FILE)

    sheet = wb["POINT"]

    updated = False

    for row in sheet.iter_rows(min_row=2):

        if str(row[0].value).strip() == str(point).strip():

            # D/E = NORMAL indication
            row[3].value = n_x
            row[4].value = n_y

            # F/G = REVERSE indication
            row[5].value = r_x
            row[6].value = r_y

            updated = True
            break

    if not updated:
        sheet.append([
            point,
            None, None,          # CONTROL X/Y
            n_x, n_y,            # NORMAL X/Y
            r_x, r_y,            # REVERSE X/Y
            None, None           # FREE X/Y
        ])

    wb.save(CONFIG_FILE)
    wb.close()


def record_point_control_coordinate():

    log("================================")
    log("POINT CONTROL COORDINATE RECORDER")
    log("================================")

    build_quantity_capture_list()

    if not loaded_point_list:
        log("No Points To Record")
        return

    wb = load_workbook(CONFIG_FILE)


    for index, point in enumerate(loaded_point_list, start=1):
        coords = [None, None]

        step = 0
        while step < 2:

            if step == 0:

                action, result = capture_point(
                    f"POINT {index} OF {len(loaded_point_list)}\n\nPOINT ID : {point}",
                    "RECORD POINT A"
                )

            else:

                action, result = capture_point(
                    f"POINT {index} OF {len(loaded_point_list)}\n\nPOINT ID : {point}",
                    "RECORD POINT B"
                )

            if action == "CANCEL":
                wb.close()
                return

            if action == "UNDO":

                if step > 0:
                    coords[step - 1] = None
                    step -= 1

                continue

            coords[step] = result
            step += 1

        save_point_control_coordinate(
            point,
            coords[0][0],
            coords[0][1],
            coords[1][0],
            coords[1][1]
        )

        log(f"{point} A : {coords[0]}")
        log(f"{point} B : {coords[1]}")

    wb.close()

    log("Point Control Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Point Control Coordinates Saved Successfully."
    )

def save_point_control_coordinate(
        point,
        a_x,
        a_y,
        b_x,
        b_y
):
    """
    Save point control coordinates without overwriting indication data.

    A/B/C/D/E/F/G/H/I/J/K mapping:
        A Point
        B/C Menu (Control A)
        D/E Normal
        F/G Reverse
        H/I Free
        J/K Control B
    """
    wb = load_workbook(CONFIG_FILE)
    sheet = wb["POINT"]

    # Make sure the extended columns exist even when an older workbook
    # was selected.
    headers = [
        "Point", "Menu_X", "Menu_Y",
        "Normal_X", "Normal_Y",
        "Reverse_X", "Reverse_Y",
        "Free_X", "Free_Y",
        "Control_B_X", "Control_B_Y"
    ]
    for col, header in enumerate(headers, 1):
        sheet.cell(row=1, column=col).value = header

    updated = False

    for row in sheet.iter_rows(min_row=2):
        if _norm_id(row[0].value) == _norm_id(point):
            # Control A / menu remains in B/C.
            row[1].value = a_x
            row[2].value = a_y

            # Control B is stored separately in J/K.
            row[9].value = b_x
            row[10].value = b_y

            updated = True
            break

    if not updated:
        sheet.append([
            point,
            a_x, a_y,
            None, None,
            None, None,
            None, None,
            b_x, b_y
        ])

    wb.save(CONFIG_FILE)
    wb.close()


def get_point_indication_coordinates(point):
    """
    Read POINT indication coordinates from the active HAH workbook.

    POINT layout:
        A Point
        B/C Menu (Control A)
        D/E Normal indication
        F/G Reverse indication
        H/I Free indication
        J/K Control B
    """
    mapping = get_point_control_coordinates(point)
    if not mapping:
        return None

    normal = mapping.get("NORMAL")
    reverse = mapping.get("REVERSE")

    if normal is None or reverse is None:
        log(
            f"POINT {point} indication mapping incomplete : "
            f"NORMAL={normal}, REVERSE={reverse}"
        )
        return None

    coordinates = (
        normal[0], normal[1],
        reverse[0], reverse[1]
    )

    log(f"POINT {point} INDICATION COORDINATES : {coordinates}")
    return coordinates


# =========================================================
# POINT LAMP DETECTION
# =========================================================

def is_normal_lamp_on(rgb):

    r, g, b = rgb

    return (
        g >= 150 and
        r <= 80 and
        b <= 120
    )


def is_reverse_lamp_on(rgb):

    r, g, b = rgb

    log(f"Reverse Check -> R={r} G={g} B={b}")
    log(f"R>=180 : {r >= 180}")
    log(f"G>=180 : {g >= 180}")
    log(f"B<=100 : {b <= 100}")

    return (
        r >= 180 and
        g >= 180 and
        b <= 100
    )
def detect_point_state(point):

    log("--------------------------------")
    log(f"CHECKING POINT {point} INDICATION")

    coordinates = get_point_indication_coordinates(point)

    if coordinates is None:
        log("Point indication coordinates not found")
        return None

    n_x, n_y, r_x, r_y = coordinates

    # Grab ALL monitors, so the reading also works when the Test Panel
    # sits on the second screen.
    try:
        screen = ImageGrab.grab(all_screens=True).convert("RGB")
        offset_x = ctypes.windll.user32.GetSystemMetrics(76)
        offset_y = ctypes.windll.user32.GetSystemMetrics(77)
    except Exception:
        screen = pyautogui.screenshot()
        offset_x = 0
        offset_y = 0

    def search_lamp(x, y):

        normal_found = False
        reverse_found = False

        best_green = (0, 0, 0)
        best_yellow = (0, 0, 0)

        # Search 11x11 area around recorded coordinate
        for dx in range(-5, 6):
            for dy in range(-5, 6):

                try:
                    rgb = screen.getpixel(
                        (x + dx - offset_x, y + dy - offset_y)
                    )[:3]
                except Exception:
                    continue

                r, g, b = rgb

                # Green lamp
                if (
                        g > r + 40 and
                        g > b + 40 and
                        g > 120
                ):
                    normal_found = True

                    if g > best_green[1]:
                        best_green = rgb

                # Yellow lamp
                if (
                        r > 140 and
                        g > 140 and
                        abs(r - g) < 70 and
                        b < 130
                ):
                    reverse_found = True

                    if (r + g) > (best_yellow[0] + best_yellow[1]):
                        best_yellow = rgb

        return (
            normal_found,
            reverse_found,
            best_green,
            best_yellow
        )

    n_green, n_yellow, green_rgb, yellow_rgb = search_lamp(n_x, n_y)

    r_green, r_yellow, green2_rgb, yellow2_rgb = search_lamp(r_x, r_y)

    log(f"Normal Area Best Green  : {green_rgb}")
    log(f"Normal Area Best Yellow : {yellow_rgb}")

    log(f"Reverse Area Best Green : {green2_rgb}")
    log(f"Reverse Area Best Yellow: {yellow2_rgb}")

    # Normal indication lamp
    if n_green and not r_yellow:

        log(f"Point {point} Current State : NORMAL")
        return "N"

    # Reverse indication lamp
    if r_yellow and not n_green:

        log(f"Point {point} Current State : REVERSE")
        return "R"

    # Cross-check
    if r_green:
        log(f"Point {point} Current State : NORMAL")
        return "N"

    if n_yellow:
        log(f"Point {point} Current State : REVERSE")
        return "R"

    # Both indication lamps dark : the point is showing FREE / out of
    # correspondence (for example the crank handle is still keyed OUT),
    # or nothing is lit at the recorded coordinates.
    free_coord = (get_point_control_coordinates(point) or {}).get("FREE")

    if free_coord:
        f_green, f_yellow, f_green_rgb, f_yellow_rgb = search_lamp(
            free_coord[0], free_coord[1]
        )
        log(
            f"Free Area Best Green : {f_green_rgb} | "
            f"Best Yellow : {f_yellow_rgb}"
        )
        if f_green or f_yellow:
            log(
                f"Point {point} Current State : FREE "
                f"(not NORMAL/REVERSE) - check the crank handle is IN"
            )
            return None

    log(
        f"Point {point} Current State : UNKNOWN - no lamp lit at "
        f"NORMAL {(n_x, n_y)} / REVERSE {(r_x, r_y)}"
    )
    return None


def is_point_free(point):
    """True when the point's FREE lamp is lit (out of correspondence,
    typically because its crank handle is still keyed OUT)."""

    free_coord = (get_point_control_coordinates(point) or {}).get("FREE")

    if not free_coord:
        return False

    try:
        screen = ImageGrab.grab(all_screens=True).convert("RGB")
        offset_x = ctypes.windll.user32.GetSystemMetrics(76)
        offset_y = ctypes.windll.user32.GetSystemMetrics(77)
    except Exception:
        screen = pyautogui.screenshot()
        offset_x = 0
        offset_y = 0

    x, y = free_coord

    for dx in range(-5, 6):
        for dy in range(-5, 6):
            try:
                r, g, b = screen.getpixel(
                    (x + dx - offset_x, y + dy - offset_y)
                )[:3]
            except Exception:
                continue

            if g > r + 40 and g > b + 40 and g > 120:
                return True

            if r > 140 and g > 140 and b < 130:
                return True

    return False


def recover_point_from_free(point, ch):
    """A point left FREE from an earlier run cannot be operated.
    Put its crank handle back IN (Receive) and re-check."""

    if not is_point_free(point):
        return False

    log(
        f"POINT {point} IS FREE - crank handle {ch} is still OUT. "
        f"Putting it back IN before the test."
    )

    if not operate_crank_handle(ch, "Receive"):
        log(f"CRANK HANDLE {ch} RECEIVE FAILED - point {point} still FREE")
        return False

    time.sleep(2)

    state = detect_point_state(point)

    if state is None:
        log(f"POINT {point} STILL NOT IN NORMAL/REVERSE AFTER CRANK HANDLE IN")
        return False

    log(f"POINT {point} RECOVERED : {state}")
    return True


def _norm_id(value):
    """Normalize Excel/UI identifiers for reliable workbook matching."""
    if value is None:
        return ""

    text = str(value).strip()

    # Excel may store numeric point IDs as 50.0.
    if text.endswith(".0"):
        text = text[:-2]

    # Treat spaces, underscores and hyphens as the same separator.
    # Example:
    #   Emergency Point Key
    #   EMERGENCY_POINT_KEY
    #   Emergency-Point-Key
    # all become the same key.
    text = re.sub(r"[^A-Z0-9]+", " ", text.upper())
    return " ".join(text.split())


def _open_config():
    if not CONFIG_FILE or not os.path.isfile(CONFIG_FILE):
        log(f"CONFIG FILE NOT FOUND : {CONFIG_FILE}")
        return None
    try:
        return load_workbook(CONFIG_FILE, data_only=True, read_only=True)
    except Exception as e:
        log(f"CONFIG OPEN ERROR : {e}")
        return None


def _find_row_by_first_column(ws, value):
    wanted = _norm_id(value)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        if _norm_id(row[0]) == wanted:
            return row
    return None


def get_point_control_coordinates(point):
    """
    POINT Excel mapping:
        Point, Menu_X, Menu_Y, Normal_X, Normal_Y,
        Reverse_X, Reverse_Y, Free_X, Free_Y,
        Control_B_X, Control_B_Y
    """
    wb = _open_config()
    if wb is None:
        return None

    try:
        if "POINT" not in wb.sheetnames:
            log("POINT sheet not found")
            return None

        row = _find_row_by_first_column(wb["POINT"], point)
        if row is None:
            log(f"POINT {point} not found in Excel mapping")
            return None

        values = list(row) + [None] * 11

        def pair(ix, iy):
            if values[ix] is None or values[iy] is None:
                return None
            try:
                return int(values[ix]), int(values[iy])
            except (TypeError, ValueError):
                return None

        result = {
            "MENU": pair(1, 2),
            "CONTROL_A": pair(1, 2),
            "NORMAL": pair(3, 4),
            "REVERSE": pair(5, 6),
            "FREE": pair(7, 8),
            "CONTROL_B": pair(9, 10),
        }

        log(f"POINT {point} EXCEL MAPPING : {result}")
        return result

    except Exception as e:
        log(f"POINT COORDINATE ERROR : {e}")
        return None
    finally:
        wb.close()



def get_ch_coordinates(ch):
    """
    CH Excel mapping:
        Point, Menu_X, Menu_Y, IN_X, IN_Y, OUT_X, OUT_Y, FREE_X, FREE_Y
    """
    wb = _open_config()
    if wb is None:
        return None

    try:
        if "CH" not in wb.sheetnames:
            log("CH sheet not found")
            return None

        row = _find_row_by_first_column(wb["CH"], ch)
        if row is None:
            log(f"CH {ch} not found in Excel mapping")
            return None

        values = list(row) + [None] * 11

        def pair(ix, iy):
            if values[ix] is None or values[iy] is None:
                return None
            return int(values[ix]), int(values[iy])

        # Columns are located by HEADER NAME, so a CH sheet written by the
        # universal yard recorder (Menu / IN / OUT / ECH / FREE) is read
        # correctly instead of taking ECH as FREE.
        headers = [
            str(h).strip().upper() if h is not None else ""
            for h in next(
                wb["CH"].iter_rows(min_row=1, max_row=1, values_only=True),
                ()
            )
        ]

        def pair_by_header(header_name, default_ix, default_iy):
            for idx, head in enumerate(headers):
                if head == header_name:
                    return pair(idx, idx + 1)
            return pair(default_ix, default_iy)

        result = {
            "MENU": pair_by_header("MENU_X", 1, 2),
            "IN": pair_by_header("IN_X", 3, 4),
            "OUT": pair_by_header("OUT_X", 5, 6),
            "FREE": pair_by_header("FREE_X", 7, 8),
        }

        log(f"CH {ch} EXCEL MAPPING : {result}")
        return result

    except Exception as e:
        log(f"CH COORDINATE ERROR : {e}")
        return None
    finally:
        wb.close()


def _track_name_variants(track):
    """1CXTPR <-> 1CXT : match a track with or without the PR suffix."""
    name = _norm_id(track)
    variants = [name]
    if name.endswith("PR"):
        variants.append(name[:-2])
    else:
        variants.append(name + "PR")
    return variants


def _track_from_edrc_file(track):
    """Track coordinates recorded by the EDRC AUTOMATION SUITE
    (TRACK_COORDINATES_CAPTURED.xlsx, passed as TRACK COORDS)."""

    candidates = [
        os.environ.get("EDRC_TRACK_COORDS"),
        os.environ.get("EDRC_TRACKS"),
        os.environ.get("EDRC_TRACK_COORDINATES"),
        os.path.join(os.getcwd(), "TRACK_COORDINATES_CAPTURED.xlsx"),
    ]

    wanted = _track_name_variants(track)

    for item in candidates:

        if not item:
            continue

        item = os.path.abspath(item)

        if not os.path.exists(item):
            continue

        try:
            wb = load_workbook(item, data_only=True)

            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    if not row or row[0] is None or len(row) < 3:
                        continue
                    if _norm_id(row[0]) not in wanted:
                        continue
                    try:
                        result = (int(float(row[1])), int(float(row[2])))
                    except (TypeError, ValueError):
                        continue
                    wb.close()
                    log(
                        f"TRACK {track} FROM "
                        f"{os.path.basename(item)} : {result}"
                    )
                    return result

            wb.close()

        except Exception as e:
            log(f"TRACK COORDINATE FILE ERROR : {item} : {e}")

    return None


def auto_track_for_point(point, side=None):
    """Work out the point's own track when the TOC Track column is empty.

    Point 50 -> 50AXTPR / 50BXTPR, point 65 -> 65AXTPR / 65BXTPR.
    The first name that exists in the track coordinate source is used,
    so TRACK DOWN still runs without a TOC Track entry."""

    point_id = _norm_id(point).replace(" ", "")

    if not point_id:
        return None

    candidates = []

    if side:
        candidates.append(f"{point_id}{str(side).strip().upper()}XTPR")

    candidates += [
        f"{point_id}AXTPR",
        f"{point_id}BXTPR",
        f"{point_id}XTPR"
    ]

    seen = []

    for name in candidates:

        if name in seen:
            continue

        seen.append(name)

        if get_track_coordinate(name) is not None:
            log(f"AUTO TRACK FOR POINT {point} : {name}")
            return name

    log(
        f"AUTO TRACK FOR POINT {point} : none of {seen} found in the "
        f"track coordinate file"
    )

    return None


def get_track_coordinate(track):
    """Track coordinate from TRACK_CONFIG in the coordinate workbook, or
    from the EDRC track-coordinate file when that sheet is not there."""
    wb = _open_config()
    if wb is None:
        return _track_from_edrc_file(track)

    try:
        if "TRACK_CONFIG" not in wb.sheetnames:
            return _track_from_edrc_file(track)

        ws = wb["TRACK_CONFIG"]
        wanted = _norm_id(track)

        for row in ws.iter_rows(values_only=True):
            if not row:
                continue

            # HAH COORDINATES1.xlsx currently stores TRACK_CONFIG as:
            # Track, X, Y without a header row.
            if len(row) >= 3 and _norm_id(row[0]) == wanted:
                if row[1] is None or row[2] is None:
                    return None
                result = (int(row[1]), int(row[2]))
                log(f"TRACK {track} EXCEL MAPPING : {result}")
                return result

        log(f"TRACK {track} not in TRACK_CONFIG - checking EDRC track file")
        return _track_from_edrc_file(track)

    except Exception as e:
        log(f"TRACK COORDINATE ERROR : {e}")
        return None
    finally:
        wb.close()


def get_system_control(control_name):
    """Read system-control coordinates, especially EMERGENCY POINT KEY."""
    wb = _open_config()
    if wb is None:
        return None

    try:
        if "SYSTEM_CONTROLS" not in wb.sheetnames:
            log("SYSTEM_CONTROLS sheet not found")
            return None

        ws = wb["SYSTEM_CONTROLS"]
        wanted = _norm_id(control_name)

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue

            if _norm_id(row[0]) == wanted:
                if row[1] is None or row[2] is None:
                    log(f"SYSTEM CONTROL {control_name} has no coordinate")
                    return None

                result = (int(row[1]), int(row[2]))
                log(
                    f"SYSTEM CONTROL {control_name} "
                    f"EXCEL MAPPING : {result}"
                )
                return result

        log(f"SYSTEM CONTROL {control_name} not found")
        return None

    except Exception as e:
        log(f"System Control Error : {e}")
        return None
    finally:
        wb.close()


# =========================================================
# GET POINT CONTROL
# =========================================================

def get_point_control(point):
    """Backward-compatible helper: return the POINT menu/control-A coordinate."""
    mapping = get_point_control_coordinates(point)
    if not mapping:
        return None
    return mapping.get("CONTROL_A")

def operate_crank_handle(point, operation):
    """
    Operate a crank handle using CH sheet coordinate mapping.

    The CH sheet supplies:
        MENU_X/Y       -> click the crank handle control
        IN_X/Y         -> IN indication
        OUT_X/Y        -> OUT indication
        FREE_X/Y       -> FREE indication

    The Hitachi menu command remains UIA/pywinauto based:
        Transmit -> crank handle OUT
        Receive  -> crank handle IN
    """
    log("================================")
    log(f"CRANK HANDLE : {operation}")
    log(f"CRANK HANDLE ID : {point}")
    log("================================")

    mapping = get_ch_coordinates(point)

    if not mapping or mapping.get("MENU") is None:
        log(f"Crank Handle {point} MENU coordinate not found in Excel")
        return False

    menu_x, menu_y = mapping["MENU"]

    log(f"CH MENU FROM EXCEL : ({menu_x}, {menu_y})")
    safe_click(menu_x, menu_y)
    time.sleep(1)

    if not find_and_click(operation):
        log(f"{operation} not found in Hitachi menu")
        return False

    log(f"{operation} clicked")
    time.sleep(3)

    # Verify the expected indication using the same Excel mapping.
    expected = "OUT" if str(operation).strip().lower() == "transmit" else "IN"
    for attempt in range(5):
        state = get_crank_handle_state(point)
        if state == expected:
            log(f"CRANK HANDLE {point} STATE : {state}")
            log(f"CRANK HANDLE {operation} : PASS")
            return True
        time.sleep(1)

    log(
        f"CRANK HANDLE {operation} verification failed. "
        f"Expected={expected}, Actual={get_crank_handle_state(point)}"
    )
    return False




from pywinauto import Desktop



def find_and_click(name, control_type=None):
    log(f"Searching for : {name}")

    try:
        desktop = Desktop(backend="uia")

        elements = desktop.windows()

        log("UI search started")

        for window in elements:

            try:
                controls = window.descendants()

                for control in controls:

                    try:
                        text = control.window_text().strip()

                        if text != str(name).strip():
                            continue

                        rect = control.rectangle()

                        left = rect.left
                        top = rect.top
                        right = rect.right
                        bottom = rect.bottom

                        # Ignore invalid/off-screen controls
                        if (
                                left < 0 or
                                top < 0 or
                                right > 1920 or
                                bottom > 1080 or
                                right <= left or
                                bottom <= top
                        ):
                            continue

                        x = (left + right) // 2
                        y = (top + bottom) // 2

                        log(
                            f"{name} -> "
                            f"L={left}, T={top}, "
                            f"R={right}, B={bottom}"
                        )

                        log(f"Click Position : ({x}, {y})")

                        safe_click(x, y, duration=0.3)

                        log(f"Clicked : {name}")

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



def open_test_panel():

    """
    EDRC / walk-away panel handling.

    The Hitachi Test Panel must already be open before this test starts.
    This function DOES NOT launch Panel.exe and therefore does not use a
    hard-coded application/configuration path.

    It only finds the existing Test Panel, brings it to the foreground, and
    then the automation GUI is minimized so the panel remains visible.
    """

    log("========================================")
    log("USING EXISTING TEST PANEL")
    log("========================================")

    try:

        # get_test_panel() is defined later in this file, but is available
        # by the time this function is called.
        window = get_test_panel()

        if window is None:
            log("EXISTING TEST PANEL NOT FOUND")
            log("PANEL MUST BE OPEN BEFORE STARTING THE EDRC TEST")

            if not EDRC_UNATTENDED:
                messagebox.showerror(
                    "Panel Not Found",
                    "The Hitachi Test Panel is not open.\\n\\n"
                    "Open the panel first, then start the EDRC test."
                )

            return False

        try:
            window.restore()
        except Exception:
            pass

        try:
            window.set_focus()
        except Exception:
            pass

        try:
            window.maximize()
        except Exception:
            pass

        time.sleep(1)

        log("EXISTING TEST PANEL FOUND")
        log("TEST PANEL BROUGHT TO FOREGROUND")

        # Minimize this automation window. The Hitachi panel stays visible.
        try:
            root.iconify()
            root.update_idletasks()
            root.update()
            log("AUTOMATION WINDOW MINIMIZED")
        except Exception as e:
            log(f"AUTOMATION WINDOW MINIMIZE ERROR : {e}")

        time.sleep(1)

        return True

    except Exception as e:

        log(f"EXISTING TEST PANEL ERROR : {e}")

        if not EDRC_UNATTENDED:
            messagebox.showerror(
                "Panel Error",
                f"Unable to use the existing Test Panel.\\n\\n{e}"
            )

        return False

# =========================================================
# RECORD CRANK HANDLE
# =========================================================

def record_crank_handle_coordinate():

    log("===================================")
    log("CRANK HANDLE COORDINATE RECORDER")
    log("===================================")
    build_ch_capture_list()

    if not loaded_ch_list:
        log("No Crank Handles To Record")
        return

    wb = load_workbook(CONFIG_FILE)
    ws = wb["CH"]

    for point in loaded_ch_list:

        coords = [None, None, None, None]

        steps = [
            "RECORD CONTROL",
            "RECORD IN",
            "RECORD OUT",
            "RECORD FREE"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"CH {point}",
                steps[step]
            )

            if action == "CANCEL":
                wb.close()
                return

            if action == "UNDO":

                if step > 0:
                    coords[step - 1] = None
                    step -= 1

                continue

            coords[step] = result
            step += 1

        updated = False

        for row in range(2, ws.max_row + 1):

            value = ws[f"A{row}"].value

            if value and str(value).strip() == str(point).strip():
                ws[f"B{row}"] = coords[0][0]
                ws[f"C{row}"] = coords[0][1]

                ws[f"D{row}"] = coords[1][0]
                ws[f"E{row}"] = coords[1][1]

                ws[f"F{row}"] = coords[2][0]
                ws[f"G{row}"] = coords[2][1]

                ws[f"H{row}"] = coords[3][0]
                ws[f"I{row}"] = coords[3][1]

                updated = True
                break

        if not updated:
            ws.append([
                point,
                coords[0][0], coords[0][1],
                coords[1][0], coords[1][1],
                coords[2][0], coords[2][1],
                coords[3][0], coords[3][1]
            ])

        log(f"{point} CONTROL : {coords[0]}")
        log(f"{point} IN      : {coords[1]}")
        log(f"{point} OUT     : {coords[2]}")
        log(f"{point} FREE    : {coords[3]}")

    wb.save(CONFIG_FILE)
    wb.close()

    log("Crank Handle Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Crank Handle Coordinates Saved Successfully."
    )

# =========================================================
# GET CRANK HANDLE STATE
# =========================================================

def get_crank_handle_state(point):

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "CH" not in wb.sheetnames:
        wb.close()

        log("CH sheet not found")

        return None

    sheet = wb["CH"]

    control = in_lamp = out_lamp = free_lamp = None

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip() != str(point).strip():
            continue

        control = (
            int(row[1]),
            int(row[2])
        )

        in_lamp = (
            int(row[3]),
            int(row[4])
        )

        out_lamp = (
            int(row[5]),
            int(row[6])
        )

        free_lamp = (
            int(row[7]),
            int(row[8])
        )

        break

    wb.close()

    if in_lamp is None:

        log(f"Crank Handle {point} not found")

        return None

    image = pyautogui.screenshot()

    in_rgb = image.getpixel(in_lamp)

    out_rgb = image.getpixel(out_lamp)

    free_rgb = image.getpixel(free_lamp)

    log("--------------------------------")
    log(f"CRANK HANDLE : {point}")
    log(f"IN RGB   : {in_rgb}")
    log(f"OUT RGB  : {out_rgb}")
    log(f"FREE RGB : {free_rgb}")

    # Yellow Lamp
    in_on = (
        in_rgb[0] > 180 and
        in_rgb[1] > 180 and
        in_rgb[2] < 120
    )

    # Red Lamp
    out_on = (
        out_rgb[0] > 180 and
        out_rgb[1] < 100 and
        out_rgb[2] < 100
    )

    # Green Lamp
    free_on = (
        free_rgb[1] > 180 and
        free_rgb[0] < 100 and
        free_rgb[2] < 100
    )

    log(f"IN    : {in_on}")
    log(f"OUT   : {out_on}")
    log(f"FREE  : {free_on}")

    if in_on:

        log("Crank Handle State : IN")

        return "IN"

    if out_on:

        log("Crank Handle State : OUT")

        return "OUT"

    if free_on:

        log("Crank Handle State : FREE")

        return "FREE"

    log("Crank Handle State : UNKNOWN")

    return "UNKNOWN"


def record_emergency_point_key():
    global CONFIG_FILE

    if not CONFIG_FILE:
        messagebox.showwarning(
            "Configuration Required",
            "Please load HAH COORDINATES1.xlsx before recording "
            "the Emergency Point Key."
        )
        return


    log("===================================")
    log("EMERGENCY POINT KEY RECORDER")
    log("===================================")

    wb = load_workbook(CONFIG_FILE)

    if "SYSTEM_CONTROLS" not in wb.sheetnames:

        wb.close()

        messagebox.showerror(
            "Error",
            "SYSTEM_CONTROLS sheet not found."
        )

        return

    ws = wb["SYSTEM_CONTROLS"]

    row = None

    for r in range(2, ws.max_row + 1):

        value = ws[f"A{r}"].value

        if value and str(value).strip().upper() == "EMERGENCY POINT KEY":

            row = r
            break

    if row is None:

        ws.append([
            "EMERGENCY POINT KEY",
            None,
            None
        ])

        row = ws.max_row

    while True:

        action, result = capture_point(
            "SYSTEM CONTROL",
            "RECORD EMERGENCY POINT KEY"
        )

        if action == "CANCEL":

            wb.close()
            return

        if action == "UNDO":
            continue

        ws[f"B{row}"] = result.x
        ws[f"C{row}"] = result.y

        log(f"Emergency Point Key : ({result.x}, {result.y})")

        break

    wb.save(CONFIG_FILE)
    wb.close()

    log("Emergency Point Key Saved")

    messagebox.showinfo(
        "Completed",
        "Emergency Point Key Saved Successfully."
    )

import pyautogui
import time

def safe_click(x, y, duration=0.5, tolerance=3):

    while True:

        log(f"Moving to ({x}, {y})")

        pyautogui.moveTo(x, y, duration=duration)

        time.sleep(0.1)

        mx, my = pyautogui.position()

        # Mouse reached the destination
        if abs(mx - x) <= tolerance and abs(my - y) <= tolerance:

            pyautogui.click()

            log("Click Successful")

            return

        # User moved the mouse
        log("Mouse movement detected.")
        log("Waiting for mouse to become idle...")

        while True:

            p1 = pyautogui.position()

            time.sleep(0.5)

            p2 = pyautogui.position()

            if p1 == p2:

                log("Mouse is idle. Retrying click...")

                break




#=========================================================
# OPERATE POINT
#=========================================================

def get_point_operation_coordinate(point, side):
    """
    Resolve the actual HAH point operation coordinate.

    From the uploaded HAH COORDINATES workbook:
        POINT row 50:
            Menu_X/Menu_Y     -> Point 50A
            Control_B_X/Y     -> Point 50B

        POINT row 65:
            Menu_X/Menu_Y     -> Point 65A
            Control_B_X/Y     -> Point 65B
    """
    mapping = get_point_control_coordinates(point)

    if not mapping:
        log(f"POINT {point} mapping not found")
        return None

    side = str(side or "A").strip().upper()

    if side == "B":
        coordinate = mapping.get("CONTROL_B")
        label = "B"
    else:
        coordinate = mapping.get("MENU")
        label = "A"

    if coordinate is None:
        log(
            f"POINT {point}{label} operation coordinate "
            f"not found in Excel"
        )
        return None

    log(
        f"POINT {point}{label} OPERATION COORDINATE "
        f"FROM EXCEL : {coordinate}"
    )
    return coordinate


def operate_point(point, track, side=None):
    """
    Operate the base POINT row using the A/B side encoded in TRACK.

    Examples:
        track 50AXTPR -> use POINT row 50, A coordinate
        track 50BXTPR -> use POINT row 50, B coordinate
        track 65AXTPR -> use POINT row 65, A coordinate
        track 65BXTPR -> use POINT row 65, B coordinate

    Indication detection always uses the base POINT row (50/65), because
    that is exactly how the uploaded HAH configuration is structured.
    """
    if side is None:
        _point, side = resolve_point_from_track(track, fallback_point=point)

    log("================================")
    log(f"OPERATING POINT : {point}{side or ''}")
    log(f"POINT CONFIG ROW: {point}")
    log(f"TRACK CONTEXT   : {track}")
    log(f"OPERATING SIDE  : {side}")
    log("================================")

    current_state = None

    for attempt in range(5):
        current_state = detect_point_state(point)
        if current_state is not None:
            break
        log(f"Point indication retry : {attempt + 1}/5")
        time.sleep(1)

    if current_state is None:
        log(f"Unable to detect Point {point}")
        return False

    operation_coordinate = get_point_operation_coordinate(point, side)

    if operation_coordinate is None:
        return False

    op_x, op_y = operation_coordinate

    # This coordinate is the actual A/B point operation target.
    log(
        f"CLICKING POINT {point}{side or ''} "
        f"AT ({op_x}, {op_y})"
    )
    safe_click(op_x, op_y)
    time.sleep(1)

    operation = "Reverse" if current_state == "N" else "Normal"

    log(f"CURRENT STATE : {current_state}")
    log(f"SELECTING OPERATION : {operation}")

    if not find_and_click(operation):
        log(f"{operation} menu item not found")
        return False

    for attempt in range(8):
        time.sleep(1)
        new_state = detect_point_state(point)

        if new_state is not None and new_state != current_state:
            log(f"NEW STATE : {new_state}")
            log(f"POINT {point}{side or ''} OPERATION : PASS")
            return True

    log(f"NEW STATE : {detect_point_state(point)}")
    log(f"POINT {point}{side or ''} OPERATION : FAIL")
    return False


# ==========================================================
# GET TEST PANEL WINDOW
# ==========================================================

def get_test_panel():

    desktop = Desktop(backend="uia")

    end = time.time() + 5

    while time.time() < end:

        try:

            for window in desktop.windows():

                title = window.window_text().strip()

                if "Test Panel" in title:
                    return window

        except Exception:
            pass

        time.sleep(0.2)

    return None
# ==========================================================
# FIND CONTROL
# ==========================================================

def find_control(name, timeout=SEARCH_TIMEOUT):

    log("--------------------------------")
    log(f"Searching : {name}")

    end = time.time() + timeout

    while time.time() < end:

        window = get_test_panel()

        if window is None:
            time.sleep(SEARCH_RETRY)
            continue

        try:

            for control in window.descendants():

                try:

                    text = control.window_text().strip()

                    if text != str(name).strip():
                        continue

                    rect = control.rectangle()

                    if (
                        rect.left < 0 or
                        rect.top < 0 or
                        rect.right <= rect.left or
                        rect.bottom <= rect.top
                    ):
                        continue

                    log(f"Found : {name}")

                    return control

                except Exception:
                    pass

        except Exception:
            pass

        time.sleep(SEARCH_RETRY)

    log(f"{name} NOT FOUND")

    return None


# ==========================================================
# CLICK CONTROL
# ==========================================================

def click_control(name):

    control = find_control(name)

    if control is None:
        log(f"{name} not found")
        return False

    rect = control.rectangle()

    x = (rect.left + rect.right) // 2
    y = (rect.top + rect.bottom) // 2

    pyautogui.moveTo(x, y, duration=0.25)
    pyautogui.click()

    log(f"Clicked : {name}")

    return True
# ==========================================================
# OPEN BIT CHART
# ==========================================================

def open_bit_chart():

    log("================================")
    log("OPEN BIT CHART")
    log("================================")

    for attempt in range(1, 4):

        log(f"Attempt : {attempt}")
        try:
            windows = Desktop(backend="uia").windows()

            for w in windows:
                if "Test Panel" in w.window_text():
                    w.set_focus()
                    break

            time.sleep(1)

        except Exception as e:
            log(f"Focus Error : {e}")

        pyautogui.hotkey("ctrl", "b")

        time.sleep(2)
        window = get_test_panel()

        if window is None:
            log("Test Panel Not Found")
            continue

        try:
            window.set_focus()
        except:
            pass

        log("============================")

        if not find_and_click("50051"):

            log("Station Not Found")

            continue

        time.sleep(1)

        if not find_and_click("OK"):

            log("OK Button Not Found")

            continue

        time.sleep(2)

        if verify_bit_chart():

            log("Bit Chart Open Success")

            return True

        log("Retry Opening Bit Chart")

    log("Unable To Open Bit Chart")

    return False
# ==========================================================
# VERIFY BIT CHART
# ==========================================================

def verify_bit_chart():

    end = time.time() + BIT_CHART_TIMEOUT

    while time.time() < end:

        window = get_test_panel()

        if window is None:

            time.sleep(0.25)
            continue

        try:

            for control in window.descendants():

                try:

                    text = control.window_text().strip()

                    if text == "Transmit":

                        log("Bit Chart Verified")

                        return True

                except Exception:
                    pass

        except Exception:
            pass

        time.sleep(0.25)

    log("Bit Chart Window Not Detected")

    return False


# ==========================================================
# CLOSE BIT CHART
# ==========================================================

def close_bit_chart():

    log("--------------------------------")
    log("Closing Bit Chart")

    if not find_and_click("Cancel"):
        log("Failed to close Bit Chart")
        return False

    time.sleep(1)

    log("Bit Chart Closed")

    return True
# ==========================================================
# TOGGLE TRACK
# ==========================================================

def toggle_track(track):

    log("================================")
    log(f"TOGGLING TRACK : {track}")
    log("================================")

    # --------------------------------
    # Track
    # --------------------------------

    if not click_control(track):

        log(f"Failed to click track : {track}")

        return False

    time.sleep(1)

    # --------------------------------
    # Transmit
    # --------------------------------

    if not click_control("Transmit"):

        log("Transmit button not found")

        return False

    log("Transmit Clicked")

    time.sleep(2)

    # --------------------------------
    # Cancel
    # --------------------------------

    if not click_control("Cancel"):

        log("Cancel button not found")

        return False

    log("Cancel Clicked")

    time.sleep(1)

    log("Track Toggle Completed")

    return True

# ==========================================================
# TRACK DOWN
# ==========================================================

def track_down(track):

    log("")
    log("================================")
    log(f"TRACK DOWN : {track}")
    log("================================")

    if not open_bit_chart():
        return False

    if not toggle_track(track):

        log("--------------------------------")
        log("TRACK DOWN FAILED")
        log("--------------------------------")

        close_bit_chart()

        return False

    close_bit_chart()

    log("--------------------------------")
    log("TRACK DOWN COMPLETED")
    log("--------------------------------")

    return True

# ==========================================================
# TRACK UP
# ==========================================================

def track_up(track):

    log("")
    log("================================")
    log(f"TRACK UP : {track}")
    log("================================")

    if not open_bit_chart():
        return False

    if not toggle_track(track):

        log("--------------------------------")
        log("TRACK UP FAILED")
        log("--------------------------------")

        close_bit_chart()

        return False

    close_bit_chart()

    log("--------------------------------")
    log("TRACK UP COMPLETED")
    log("--------------------------------")

    return True

# =========================================================
# UIA LOGIN / PASSWORD HELPERS
# =========================================================
#
# The Hitachi password dialogs are handled through pywinauto's UIA
# backend.  No Windows native backend or keyboard text injection is used here.
#
# The helper looks for Edit controls in the password dialog and writes
# the values directly through UI Automation. This is more reliable than
# sending keyboard text blindly to the active window.
# =========================================================

def _uia_window_text(window):
    try:
        return window.window_text().strip()
    except Exception:
        return ""


def _uia_visible_edits(window):
    edits = []

    try:
        controls = window.descendants()
    except Exception:
        return edits

    for control in controls:
        try:
            if control.element_info.control_type != "Edit":
                continue

            rect = control.rectangle()

            if (
                rect.right <= rect.left
                or rect.bottom <= rect.top
            ):
                continue

            edits.append(control)

        except Exception:
            continue

    # Top-to-bottom order normally corresponds to username/password order.
    edits.sort(
        key=lambda c: (
            c.rectangle().top,
            c.rectangle().left
        )
    )

    return edits


def _uia_set_text(control, value):
    """Set text directly through UIA without keyboard send-keys."""
    try:
        control.set_edit_text(str(value))
        return True
    except Exception as e:
        log(f"UIA set_edit_text failed : {e}")

    # Some legacy controls expose ValuePattern rather than the pywinauto
    # Edit wrapper method. Try the UIA value pattern directly.
    try:
        pattern = control.iface_value
        pattern.SetValue(str(value))
        return True
    except Exception as e:
        log(f"UIA ValuePattern failed : {e}")

    return False


def _uia_click_button(window, names):
    """Invoke a dialog button through UIA."""
    wanted = {
        str(name).strip().lower()
        for name in names
    }

    try:
        controls = window.descendants()
    except Exception:
        controls = []

    for control in controls:
        try:
            if control.element_info.control_type != "Button":
                continue

            caption = control.window_text().strip().lower()

            if caption not in wanted:
                continue

            try:
                control.invoke()
                log(f"UIA invoked button : {control.window_text()}")
                return True
            except Exception as e:
                log(f"UIA invoke failed for {caption} : {e}")

                try:
                    control.click_input()
                    log(f"UIA clicked button : {control.window_text()}")
                    return True
                except Exception as e2:
                    log(f"UIA click_input failed : {e2}")

        except Exception:
            continue

    return False


def enter_panel_credentials(
    username="ETOE",
    password="ETOE",
    preferred_titles=None,
    timeout=8
):
    """
    Find the currently displayed Hitachi credential dialog and enter
    username/password through UIA Edit controls.

    preferred_titles:
        Optional list of exact/partial dialog titles.  If supplied,
        matching windows are checked first.

    Returns True only after both fields are populated and the dialog's
    confirmation button is invoked.
    """
    if preferred_titles is None:
        preferred_titles = [
            "Password for Key in",
            "Password",
            "Emergency",
            "Key Out"
        ]

    log("--------------------------------")
    log("UIA CREDENTIAL ENTRY STARTED")
    log("--------------------------------")

    # Only the Test Panel's own windows may receive the credentials.
    # Without this the search walked EVERY desktop window and typed the
    # username/password into whatever had edit boxes - for example a
    # File Explorer window.
    panel_pid = None

    try:
        for window in Desktop(backend="uia").windows():
            try:
                if "TEST PANEL" in _uia_window_text(window).upper():
                    panel_pid = window.element_info.process_id
                    break
            except Exception:
                continue
    except Exception:
        panel_pid = None

    log(f"CREDENTIAL DIALOG : Test Panel process = {panel_pid}")

    BLOCKED_WINDOWS = (
        "file explorer",
        "program manager",
        "task manager",
        "windows explorer",
        "chrome",
        "edge",
        "firefox",
        "notepad",
        "excel",
        "visual studio",
        "pycharm",
        "command prompt",
        "powershell",
        "edrc",
        "automation"
    )

    def _credential_window_allowed(window, title):

        text = (title or "").lower()

        # Never type credentials into another application, whatever its
        # window title happens to contain (a folder named
        # 01_EMERGENCY_POINT_KEY_TEST matched "Emergency" before).
        for blocked in BLOCKED_WINDOWS:
            if blocked in text:
                return False

        # The Test Panel's own windows are the only valid target.
        if panel_pid is not None:
            try:
                return window.element_info.process_id == panel_pid
            except Exception:
                return False

        # Test Panel process unknown : fall back to the dialog titles.
        for wanted_title in preferred_titles:
            wanted = wanted_title.strip().lower()
            if wanted and wanted in text:
                return True

        return False

    deadline = time.time() + timeout

    while time.time() < deadline:
        try:
            desktop = Desktop(backend="uia")
            windows = desktop.windows()

            # First check preferred dialog titles.
            ordered_windows = []

            for wanted_title in preferred_titles:
                wanted = wanted_title.strip().lower()

                for window in windows:
                    title = _uia_window_text(window).lower()

                    if wanted and wanted in title:
                        if window not in ordered_windows:
                            ordered_windows.append(window)

            # Then check remaining windows. This handles dialogs whose
            # caption differs slightly between panel versions.
            for window in windows:
                if window not in ordered_windows:
                    ordered_windows.append(window)

            for window in ordered_windows:
                title = _uia_window_text(window)

                if not title:
                    continue

                if not _credential_window_allowed(window, title):
                    continue

                edits = _uia_visible_edits(window)

                if not edits:
                    continue

                log(f"Credential dialog candidate : {title}")
                log(f"UIA Edit controls found : {len(edits)}")

                try:
                    window.set_focus()
                except Exception:
                    pass

                # Normal Hitachi credential dialog:
                # Edit 1 = username
                # Edit 2 = password
                if len(edits) >= 2:
                    if not _uia_set_text(edits[0], username):
                        log("Username field could not be populated")
                        continue

                    if not _uia_set_text(edits[1], password):
                        log("Password field could not be populated")
                        continue

                    log("Username field populated through UIA")
                    log("Password field populated through UIA")

                elif len(edits) == 1:
                    # Some panel builds may show only the password field.
                    # In that case, populate the single field with the
                    # supplied password.
                    if not _uia_set_text(edits[0], password):
                        log("Password field could not be populated")
                        continue

                    log("Single password field populated through UIA")

                else:
                    continue

                # Confirm without keyboard text injection.
                if _uia_click_button(
                    window,
                    ["OK", "Login", "Enter", "Apply", "Yes"]
                ):
                    log("Credential dialog confirmed through UIA")
                    return True

                # If the dialog has no named button, try the UIA default
                # button. This confirmation is still performed through UI Automation.
                try:
                    buttons = [
                        c for c in window.descendants()
                        if c.element_info.control_type == "Button"
                    ]

                    for button in buttons:
                        try:
                            if button.is_enabled():
                                button.invoke()
                                log(
                                    "Credential dialog confirmed "
                                    "using first enabled UIA button"
                                )
                                return True
                        except Exception:
                            continue
                except Exception:
                    pass

        except Exception as e:
            log(f"UIA credential search error : {e}")

        time.sleep(0.25)

    log("Credential dialog was not completed through UIA")
    return False


def emergency_point_key_in():

    log("================================")
    log("EMERGENCY POINT KEY IN")
    log("================================")

    coord = get_system_control("EMERGENCY POINT KEY")

    if coord is None:
        log("Emergency Point Key Coordinate Not Found")
        return False

    x, y = coord

    safe_click(x, y)

    time.sleep(1)

    if not find_and_click("Key in"):
        log("Key in Menu Not Found")
        return False

    log("Key in Clicked")
    time.sleep(2)

    if not enter_panel_credentials(
        username="ETOE",
        password="ETOE",
        preferred_titles=["Password for Key in", "Password"],
        timeout=8
    ):
        log("Username/Password entry failed")
        return False

    log("Username and Password Entered Through UIA")
    log("Login Successful")

    time.sleep(1)

    return True


def emergency_point_operation(point, track, default_point_mode=False):

    log("================================")
    log("EMERGENCY POINT OPERATION")
    log("================================")

    state = detect_point_state(point)

    if state is None:
        log("Unable to detect point state")
        return False

    coordinates = get_point_control_coordinates(point)

    if coordinates is None:
        log("Point Control Coordinates Missing")
        return False

    # Emergency Point Key test uses the configured default point only.
    # Do not require a separate CONTROL_B coordinate for a BXTPR track.
    # The default point's recorded control coordinate is the one used.
    control_side = "CONTROL_A"

    if not default_point_mode and "BXTPR" in str(track).upper():
        control_side = "CONTROL_B"

    control_coord = coordinates.get(control_side)

    if control_coord is None:
        log(
            f"Point {point} {control_side} coordinate not found "
            f"in POINT sheet"
        )
        return False

    x, y = control_coord

    if default_point_mode:
        log(
            "DEFAULT POINT MODE : using configured default point "
            "control coordinate (CONTROL_A)"
        )

    log(
        f"EMERGENCY CONTROL {control_side} FROM EXCEL : "
        f"({x}, {y})"
    )

    safe_click(x, y)

    time.sleep(0.5)
    time.sleep(1)

    if state == "N":

        log("Current State : NORMAL")
        log("Trying Emergency Reverse")

        if not find_and_click("Emergency Reverse"):
            return False

    else:

        log("Current State : REVERSE")
        log("Trying Emergency Normal")

        if not find_and_click("Emergency Normal"):
            return False

    log("Emergency Command Clicked")

    # ---------------------------------------------------------
    # PASSWORD
    # ---------------------------------------------------------
    # Replaced keyboard text injection with direct UIA
    # entry into the username/password controls.
    # ---------------------------------------------------------

    if not enter_panel_credentials(
        username="ETOE",
        password="ETOE",
        preferred_titles=["Password", "Emergency"],
        timeout=8
    ):
        log("Emergency operation password entry failed")
        return False

    log("Password Entered Through UIA")

    time.sleep(3)
    time.sleep(1)

    # --------------------------------
    # VERIFY
    # --------------------------------

    new_state = detect_point_state(point)

    if new_state is None:

        log("Unable to Verify Point")
        return False

    if new_state != state:

        log("EMERGENCY POINT OPERATION : PASS")
        return True

    log("EMERGENCY POINT OPERATION : FAIL")
    return False


def emergency_point_key_out():

    log("================================")
    log("EMERGENCY POINT KEY OUT")
    log("================================")

    coord = get_system_control("EMERGENCY POINT KEY")

    if coord is None:

        log("Emergency Point Key Not Found")
        return False

    x, y = coord

    safe_click(x, y)

    time.sleep(1)

    if not find_and_click("Critical"):

        log("Critical Not Found")
        return False

    time.sleep(1)

    if not find_and_click("Key Out"):

        log("Key Out Not Found")
        return False

    time.sleep(1)

    if not enter_panel_credentials(
        username="ETOE",
        password="ETOE",
        preferred_titles=["Password for Key Out", "Password", "Key Out"],
        timeout=8
    ):
        log("Key Out password entry failed")
        return False

    log("Emergency Point Key Returned")

    time.sleep(2)

    return True





# =========================================================
# TOC TRACK -> POINT / SIDE RESOLUTION
# =========================================================
def resolve_point_from_track(track, fallback_point=None):
    """
    Resolve the point row and the A/B operating side from the TOC TRACK.

    The uploaded HAH configuration stores:
        POINT 50  -> Menu_X/Menu_Y = Point 50A operation coordinate
                     Control_B_X/Y = Point 50B operation coordinate
        POINT 65  -> Menu_X/Menu_Y = Point 65A operation coordinate
                     Control_B_X/Y = Point 65B operation coordinate

    Therefore:
        50AXTPR -> POINT row 50, SIDE A
        50BXTPR -> POINT row 50, SIDE B
        65AXTPR -> POINT row 65, SIDE A
        65BXTPR -> POINT row 65, SIDE B

    IMPORTANT:
        The POINT column in the TOC (50R, 65R) is the expected indication
        state. It is not the Excel POINT row name.
    """
    if track is not None:
        track_text = str(track).strip().upper()

        # A/B are part of the track identity, but the POINT sheet uses
        # the base point number as its row key.
        match = re.search(r"^(.+?)([AB])XTPR$", track_text)

        if match:
            point = match.group(1)
            side = match.group(2)

            log(
                f"TOC TRACK RESOLUTION : {track_text} -> "
                f"POINT {point}, SIDE {side}"
            )
            return point, side

    # Fallback for a non-XTPR track, using the TOC POINT value.
    if fallback_point:
        fallback = str(fallback_point).strip().upper()
        log(
            f"TOC TRACK RESOLUTION : {track} -> "
            f"FALLBACK POINT {fallback}, SIDE A"
        )
        return fallback, "A"

    log(f"TOC TRACK RESOLUTION FAILED : {track}")
    return None, None


# Backward-compatible helper for any other code that only needs the point row.
def get_point_from_track(track, fallback_point=None):
    point, _side = resolve_point_from_track(track, fallback_point)
    return point


def get_default_point_from_config():
    """
    Return the first usable point recorded in the POINT sheet.

    For the Emergency Point Key test, this is the ONLY point that is
    verified.  The test does not require every point listed in the TOC
    route.
    """
    wb = _open_config()
    if wb is None:
        return None

    try:
        if "POINT" not in wb.sheetnames:
            log("POINT sheet not found")
            return None

        ws = wb["POINT"]

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue

            value = row[0]
            if value is None or str(value).strip() == "":
                continue

            point = str(value).strip().upper()

            # The first non-empty POINT row is the configured default point.
            log(f"DEFAULT POINT FROM CONFIG : {point}")
            return point

        log("No default point is recorded in POINT sheet")
        return None

    except Exception as e:
        log(f"DEFAULT POINT READ ERROR : {type(e).__name__} : {e}")
        return None

    finally:
        wb.close()


# =========================================================
# TOC ITEM PARSER
# =========================================================

def _parse_route_items(value):
    """
    Convert a TOC cell containing point/track/CH items into a clean list.

    Supports the common EDRC TOC formats:
        50R,65R
        50BXTPR, 65BXTPR
        CH1, CH2
        one item
    """
    if value is None:
        return []

    text_value = str(value).strip()

    if not text_value:
        return []

    parts = re.split(r"[,;/\n]+", text_value)

    result = []
    for part in parts:
        item = str(part).strip()
        if item:
            result.append(item)

    # Some workbooks use multiple spaces between items instead of commas.
    if len(result) == 1 and "  " in result[0]:
        result = [
            item.strip()
            for item in re.split(r"\s{2,}", result[0])
            if item.strip()
        ]

    return result


# =========================================================
# TEST POINT OPERATION
# =========================================================
def test_point_operation():
    """
    EDRC-safe wrapper around the existing test engine.

    The original test logic is untouched. The wrapper only guarantees that
    the EDRC master's `running` flag is cleared when the test exits or fails.
    """
    global running
    running = True
    log("[EDRC] TEST ENGINE CALLBACK ENTERED")

    try:
        return _test_point_operation_original()
    except Exception as e:
        log(f"TEST ENGINE ERROR : {type(e).__name__} : {e}")
        try:
            messagebox.showerror(
                "Test Error",
                f"Emergency Point Key test failed.\\n\\n{e}"
            )
        except Exception:
            pass
        return False
    finally:
        running = False

        # Test is complete. Bring the automation/EDRC child window back so
        # the operator can see the final status/report before EDRC closes it.
        try:
            root.deiconify()
            root.state("zoomed")
            root.lift()
            root.attributes("-topmost", True)
            root.after(300, lambda: root.attributes("-topmost", False))
            root.update_idletasks()
            root.update()
            log("AUTOMATION / EDRC WINDOW RESTORED AFTER TEST")
        except Exception as e:
            log(f"EDRC WINDOW RESTORE ERROR : {e}")

        log("[EDRC] TEST ENGINE CALLBACK EXITED")
        log("[EDRC] RUNNING FLAG = False")


def _test_point_operation_original():
    """
    Run the POINT / CRANK HANDLE / TRACK LOCK / EMERGENCY POINT KEY test
    for ONE SELECTED TOC ROW / ROUTE ONLY.

    IMPORTANT:
        The operator must first select exactly one row in the TOC table.
        The test will NOT loop through the remaining routes.

    Sequence for the selected route:
        1. Initial point operation
        2. Crank Handle OUT (Transmit)
        3. Point operation must be possible
        4. Track DOWN
        5. Point operation must be blocked
        6. Emergency Point Key IN
        7. Emergency point operation must be possible
        8. Emergency Point Key OUT
        9. Crank Handle IN (Receive)
       10. Track UP
       11. Save result to Excel report
    """

    # ---------------------------------------------------------
    # BASIC VALIDATION
    # ---------------------------------------------------------
    if TOC_FILE is None:
        messagebox.showerror("Error", "Please Load TOC First.")
        return

    if not CONFIG_FILE or not os.path.isfile(CONFIG_FILE):
        messagebox.showerror(
            "Error",
            "Please Load HAH COORDINATES1.xlsx first."
        )
        return

    # ---------------------------------------------------------
    # GET EXACTLY ONE SELECTED TOC ROW
    # ---------------------------------------------------------
    selected_items = tree.selection()

    if not selected_items:
        messagebox.showwarning(
            "Select One Route",
            "Please select ONE route from the TOC table before starting."
        )
        log("START TEST BLOCKED : No TOC route selected")
        return

    if len(selected_items) > 1:
        messagebox.showwarning(
            "Select One Route",
            "Please select ONLY ONE route from the TOC table."
        )
        log("START TEST BLOCKED : More than one TOC row selected")
        return

    selected_item = selected_items[0]
    selected_values = tree.item(selected_item, "values")

    if not selected_values:
        messagebox.showerror(
            "TOC Selection Error",
            "Unable to read the selected TOC row."
        )
        return

    # Tree columns:
    # 0 = NO
    # 1 = SIGNAL
    # 2 = ROUTE
    # 3 = LOCK_ROUTE
    # 4 = POINT
    # 5 = CRANK_HANDLE
    selected_row_no = selected_values[0]
    selected_signal = str(selected_values[1]).strip()
    selected_route = str(selected_values[2]).strip()

    if not selected_route:
        messagebox.showerror(
            "TOC Selection Error",
            "The selected row does not contain a route."
        )
        return

    log("================================")
    log("SELECTED ROUTE TEST")
    log(f"TOC ROW      : {selected_row_no}")
    log(f"SIGNAL       : {selected_signal}")
    log(f"ROUTE        : {selected_route}")
    log("ONLY THIS ROUTE WILL BE TESTED")
    log("================================")

    # ---------------------------------------------------------
    # VALIDATE COORDINATE CONFIGURATION
    # ---------------------------------------------------------
    log(f"CONFIG FILE USED BY TEST : {CONFIG_FILE}")
    wb_config = _open_config()

    if wb_config is None:
        log("TEST BLOCKED : Unable to open HAH coordinate workbook")
        if not EDRC_UNATTENDED:
            messagebox.showerror(
                "Error",
                "Unable to open HAH coordinate workbook."
            )
        return

    # TRACK_CONFIG is no longer required here: track coordinates come
    # from the EDRC track-coordinate file (TRACK_COORDINATES_CAPTURED.xlsx)
    # when the coordinate workbook has no TRACK_CONFIG sheet.
    required = {"POINT", "CH", "SYSTEM_CONTROLS"}
    missing = required.difference(set(wb_config.sheetnames))
    wb_config.close()

    if missing:
        log(
            "TEST BLOCKED : Invalid HAH coordinate workbook. "
            "Missing: " + ", ".join(sorted(missing))
        )
        if not EDRC_UNATTENDED:
            messagebox.showerror(
                "Error",
                "Invalid HAH coordinate workbook.\n\nMissing:\n" +
                "\n".join(sorted(missing))
            )
        return

    if get_system_control("EMERGENCY POINT KEY") is None:
        log(
            "TEST BLOCKED : Emergency Point Key coordinate is missing "
            "from SYSTEM_CONTROLS"
        )
        if not EDRC_UNATTENDED:
            messagebox.showerror(
                "Error",
                "Emergency Point Key coordinate is missing from "
                "SYSTEM_CONTROLS."
            )
        return

    if not open_test_panel():
        return

    create_report()

    log("================================")
    log("POINT / CRANK HANDLE TEST STARTED")
    log(f"SELECTED ROUTE : {selected_route}")
    log(f"SELECTED SIGNAL : {selected_signal}")
    log("MODE : SINGLE ROUTE ONLY")
    log("WORKFLOW SOURCE : CRANKV2")
    log("COORDINATE SOURCE : HAH EXCEL MAPPING")
    log("================================")

    time.sleep(3)

    wb = None

    try:
        wb = load_workbook(
            TOC_FILE,
            data_only=True,
            read_only=True
        )

        if "TOC" not in wb.sheetnames:
            messagebox.showerror(
                "TOC Error",
                "TOC sheet not found."
            )
            return

        ws = wb["TOC"]

        # -----------------------------------------------------
        # USE THE ROW NUMBER SELECTED IN THE GUI.
        #
        # Excel data starts at row 2 because row 1 is the header.
        # This prevents the program from accidentally running every
        # route in the workbook.
        # -----------------------------------------------------
        # -----------------------------------------------------
        # READ THE EXACT SELECTED ROUTE BY HEADER NAME.
        # The GUI now shows only the first route for each signal, so the
        # old "selected row number + 1" positional lookup is not valid.
        # -----------------------------------------------------
        headers, excel_row_no, row = _find_selected_toc_row(
            ws,
            selected_signal,
            selected_route
        )

        if row is None:
            messagebox.showerror(
                "TOC Selection Error",
                "The selected first-route entry could not be found in the TOC.\n\n"
                f"Signal : {selected_signal}\n"
                f"Route  : {selected_route}"
            )
            log(
                f"TOC LOOKUP FAILED : Signal={selected_signal}, "
                f"Route={selected_route}"
            )
            return

        toc_signal = str(_toc_value(row, headers, "Signal") or "").strip()
        toc_route = str(_toc_value(row, headers, "Route") or "").strip()

        if (
            toc_signal.upper() != selected_signal.upper()
            or toc_route.upper() != selected_route.upper()
        ):
            messagebox.showerror(
                "Route Mismatch",
                "The selected GUI route does not match the TOC row.\n\n"
                f"GUI Signal : {selected_signal}\n"
                f"TOC Signal : {toc_signal}\n"
                f"GUI Route  : {selected_route}\n"
                f"TOC Route  : {toc_route}"
            )
            log(
                f"SAFETY STOP : GUI={selected_signal}/{selected_route} "
                f"TOC={toc_signal}/{toc_route}"
            )
            return

        log("--------------------------------")
        log("SELECTED FIRST ROUTE CONFIRMED")
        log(f"EXCEL ROW   : {excel_row_no}")
        log(f"SIGNAL      : {toc_signal}")
        log(f"ROUTE       : {toc_route}")
        log("--------------------------------")

        # NEW TOC: these values MUST come from their named columns.
        # Point_Track is the per-point track column for this test
        # (one track per POINT item, in the same order).
        # The general Track column is used only if Point_Track is empty.
        track_data = _toc_value(
            row,
            headers,
            "Point_Track", "POINT TRACK", "POINT_TRACKS"
        )

        if track_data is None or str(track_data).strip() == "":
            track_data = _toc_value(row, headers, "Track")
        point_data = _toc_value(row, headers, "Points", "Point")
        ch_data = _toc_value(row, headers, "CRANK_HANDLE", "Crank Handle", "CH")

        # Only the POINT column is mandatory. The TRACK column is
        # optional : when it is empty the point row is taken from the
        # POINT value itself (50R -> POINT 50, side A).
        if point_data is None or str(point_data).strip() == "":
            log(
                f"SELECTED ROUTE {selected_route} has no POINT data in the TOC"
            )
            if not EDRC_UNATTENDED:
                messagebox.showwarning(
                    "Route Data Missing",
                    f"Route {selected_route} has no POINT data."
                )
            return

        # -----------------------------------------------------
        # READ THE POINT / TRACK / CH LISTS FROM THE SELECTED TOC ROW
        # -----------------------------------------------------
        point_list = _parse_route_items(point_data)
        track_list = _parse_route_items(track_data)
        ch_list = _parse_route_items(ch_data)

        log("================================")
        log("SELECTED ROUTE DATA")
        log(f"POINT ITEMS : {point_list}")
        log(f"TRACK ITEMS : {track_list}")
        log(f"CH ITEMS    : {ch_list}")
        log("================================")

        if not track_list:
            # No TRACK column for this route - run with the POINT values
            # only; the point row/side comes from the POINT item.
            track_list = [None] * len(point_list)
            log(
                f"ROUTE {selected_route} has no TRACK items - using the "
                f"POINT values (side A) instead"
            )

        # -----------------------------------------------------
        # EMERGENCY POINT KEY TEST:
        # TEST ALL POINTS BELONGING TO THE SELECTED TOC ROUTE.
        #
        # IMPORTANT:
        #   - The selected ROUTE is the only route executed.
        #   - Every POINT/TRACK/CH item belonging to that route is tested.
        #   - Do NOT reduce the lists to the first/default point.
        #   - CONTROL_B is not required merely because a track is BXTPR;
        #     the existing point-operation logic uses the configured
        #     operation coordinate.
        # -----------------------------------------------------

        if not point_list:
            log(f"SELECTED ROUTE {selected_route} has no point items")
            return

        if any(t is not None for t in track_list) and len(track_list) < len(point_list):
            log(
                f"ROUTE DATA ERROR : {len(point_list)} point items but "
                f"only {len(track_list)} track items"
            )
            messagebox.showwarning(
                "Route Data Mismatch",
                f"Route {selected_route} has {len(point_list)} point items "
                f"but only {len(track_list)} track items."
            )
            return

        if len(ch_list) < len(point_list):
            log(
                f"ROUTE DATA ERROR : {len(point_list)} point items but "
                f"only {len(ch_list)} crank handles"
            )
            messagebox.showwarning(
                "Route Data Mismatch",
                f"Route {selected_route} has {len(point_list)} point items "
                f"but only {len(ch_list)} crank handles."
            )
            return

        # Keep the COMPLETE selected-route lists unchanged.
        # This is the critical difference from the previous version.
        log("================================")
        log("ONE ROUTE / ALL POINTS MODE")
        log(f"SELECTED ROUTE : {selected_route}")
        log(f"POINT COUNT    : {len(point_list)}")
        log(f"POINT ITEMS    : {point_list}")
        log(f"TRACK COUNT    : {len(track_list)}")
        log(f"TRACK ITEMS    : {track_list}")
        log(f"CH COUNT       : {len(ch_list)}")
        log(f"CH ITEMS       : {ch_list}")
        log("NO DEFAULT-POINT FILTER IS APPLIED")
        log("NO OTHER ROUTE WILL BE TESTED")
        log("================================")

        # -----------------------------------------------------
        # RUN ONLY THE SELECTED ROUTE
        # ALL POINTS/TRACKS/CHs FROM THAT ROUTE ARE PRESERVED.
        # -----------------------------------------------------
        log("================================")
        log(f"RUNNING ONLY ROUTE : {selected_route}")
        log(f"POINT ITEMS : {point_list}")
        log(f"TRACK ITEMS : {track_list}")
        log(f"CH ITEMS    : {ch_list}")
        log("================================")

        for point_item, track, ch in zip(
            point_list,
            track_list,
            ch_list
        ):

            if len(point_item) < 2:
                log(f"Invalid point item : {point_item}")
                continue

            point_state_text = point_item.strip().upper()
            expected_state = point_state_text[-1]

            point, point_side = resolve_point_from_track(
                track,
                fallback_point=point_state_text[:-1].strip()
            )

            if not point:
                log(
                    f"Unable to resolve POINT from TRACK : {track} "
                    f"(TOC POINT={point_item})"
                )
                continue

            log("--------------------------------")
            log(f"ROUTE         : {selected_route}")
            log(f"TOC POINT     : {point_item}")
            log(f"POINT ROW     : {point}")
            log(f"POINT SIDE    : {point_side}")
            log(f"TRACK         : {track}")
            log(f"EXPECTED STATE: {expected_state}")
            log(f"CRANK HANDLE  : {ch}")
            log("--------------------------------")

            get_point_control_coordinates(point)
            get_ch_coordinates(ch)

            if track:
                get_track_coordinate(track)

            # --------------------------------
            # 1. Initial Point Operation
            # --------------------------------
            if get_point_control_coordinates(point) is None:
                log(
                    f"POINT {point} CONFIGURATION NOT FOUND. "
                    f"Stopping selected route."
                )
                return

            # A point left FREE by an earlier run (crank handle still
            # keyed OUT) cannot be operated - recover it first.
            recover_point_from_free(point, ch)

            # Emergency Point Key test uses the configured default point's
            # normal/A operation coordinate only.  Do not require B control
            # merely because the route track contains BXTPR.
            if get_point_operation_coordinate(
                point,
                "A"
            ) is None:
                log(
                    f"POINT {point}{point_side} OPERATION COORDINATE "
                    f"NOT FOUND. Stopping selected route."
                )
                return

            if not operate_point(
                point,
                track,
                "A"
            ):
                log("INITIAL POINT OPERATION FAILED")
                return

            # --------------------------------
            # 2. Crank Handle OUT
            # --------------------------------
            if not operate_crank_handle(
                ch,
                "Transmit"
            ):
                log("CRANK HANDLE OUT FAILED")
                return

            # --------------------------------
            # 3. Point SHOULD change now
            # --------------------------------
            if operate_point(point, track, "A"):
                log(
                    "POINT OPERATION AFTER "
                    "CRANK HANDLE : PASS"
                )
                crank_result = "PASS"
            else:
                log(
                    "POINT OPERATION AFTER "
                    "CRANK HANDLE : FAIL"
                )
                crank_result = "FAIL"

            # --------------------------------
            # 4. Track DOWN   (only if the TOC gives a track)
            # 5. Point SHOULD NOT change while the track is down
            # --------------------------------
            if not (track and str(track).strip()):
                # No Track column in the TOC : derive the point's own
                # track (50 -> 50AXTPR) so TRACK DOWN still happens.
                log(
                    f"NO TRACK IN TOC FOR POINT {point} - looking up the "
                    f"point's own track"
                )
                track = auto_track_for_point(point, point_side)

            track_used = bool(track and str(track).strip())

            if not track_used:

                log(
                    f"NO TRACK FOUND FOR POINT {point} - TRACK DOWN AND "
                    f"POINT LOCK CHECK SKIPPED"
                )

                point_lock_result = "NOT TESTED"

            else:

                if not track_down(track):
                    log("TRACK DOWN FAILED")
                    return

                if operate_point(point, track, "A"):
                    log(
                        "POINT OPERATION AFTER "
                        "TRACK DOWN : FAIL"
                    )
                    point_lock_result = "FAIL"
                else:
                    log(
                        "POINT OPERATION AFTER "
                        "TRACK DOWN : PASS"
                    )
                    point_lock_result = "PASS"

            # --------------------------------
            # 6-8. Emergency Point Key sequence
            # --------------------------------
            emergency_result = "FAIL"

            if emergency_point_key_in():

                if emergency_point_operation(
                    point,
                    track,
                    default_point_mode=True
                ):
                    log(
                        "EMERGENCY POINT OPERATION : PASS"
                    )
                    emergency_result = "PASS"
                else:
                    log(
                        "EMERGENCY POINT OPERATION : FAIL"
                    )

                if emergency_point_key_out():
                    log(
                        "EMERGENCY POINT KEY OUT : PASS"
                    )
                else:
                    log(
                        "EMERGENCY POINT KEY OUT : FAIL"
                    )

            else:
                log(
                    "EMERGENCY POINT KEY IN FAILED"
                )

            # --------------------------------
            # 9. Crank Handle IN
            # --------------------------------
            if not operate_crank_handle(
                ch,
                "Receive"
            ):
                log("CRANK HANDLE IN FAILED")
                return

            # --------------------------------
            # 10. Track UP  (only if it was put down in step 4)
            # --------------------------------
            if track_used:
                if not track_up(track):
                    log("TRACK UP FAILED")
                    return

            # --------------------------------
            # 11. Save report
            # --------------------------------
            save_result(
                point=point,
                crank_handle=crank_result,
                point_lock=point_lock_result,
                emergency_operation=emergency_result
            )

        log("================================")
        log(f"SELECTED ROUTE {selected_route} COMPLETED")
        log("NO OTHER ROUTE WAS TESTED - ALL POINTS IN SELECTED ROUTE WERE TESTED")
        log("================================")

    except Exception as e:
        log(
            f"POINT TEST ERROR : "
            f"{type(e).__name__} : {e}"
        )
        messagebox.showerror(
            "Point Test Error",
            f"Point test stopped because of an error.\n\n{e}"
        )

    finally:
        if wb is not None:
            wb.close()

        try:
            finalize_report()
        except Exception as e:
            log(
                f"REPORT FINALIZE ERROR : {e}"
            )

        log(
            f"POINT / CRANK HANDLE TEST COMPLETED "
            f"FOR ROUTE : {selected_route}"
        )


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
    root.state("zoomed")
    root.lift()

    log("AUTOMATION / EDRC WINDOW RESTORED")
    log("STOPPED")


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
# FULL COORDINATE RECORDING MENU
# =========================================================
#
# NEW SIGNAL opens this menu after creating/selecting the HAH
# configuration. Every recorder writes directly into the active
# HAH workbook:
#
#   MAIN
#   SHUNT
#   CAL
#   POINT
#   CH
#   LC
#   TRACK_CONFIG
#   SYSTEM_CONTROLS
#
# Existing Point Indication / Point Control / Crank Handle /
# Emergency Point Key recorders are retained.
# =========================================================

selected_aspects = None


def select_signal_aspects(signal):
    global selected_aspects

    selected_aspects = None

    if capture_overlay is None:
        create_capture_overlay()

    overlay_step_label.config(text="SELECT NUMBER OF ASPECTS")
    overlay_hint_label.config(
        text=f"{signal}\n\nSelect 2, 3 or 4 aspects"
    )

    for widget in overlay_button_frame.winfo_children():
        widget.destroy()

    def choose(value):
        global selected_aspects
        selected_aspects = value

    for value in (2, 3, 4):
        tk.Button(
            overlay_button_frame,
            text=str(value),
            width=6,
            height=2,
            font=("Segoe UI", 15, "bold"),
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief="flat",
            command=lambda v=value: choose(v)
        ).pack(side="left", padx=8)

    capture_overlay.deiconify()
    capture_overlay.lift()
    capture_overlay.focus_force()

    while selected_aspects is None:
        root.update()
        time.sleep(0.02)

    for widget in overlay_button_frame.winfo_children():
        widget.destroy()

    capture_overlay.withdraw()
    return selected_aspects


def _record_name_list(label, storage):
    storage.clear()
    used = set()

    count = prompt_count(label)

    for i in range(count):
        name = prompt_name(label, i, count, used)
        used.add(name)
        storage.append(name)

    return storage


def build_signal_capture_list():
    return _record_name_list("Main Signals", loaded_signal_list)


def build_shunt_capture_list():
    return _record_name_list("Shunt Signals", loaded_shunt_list)


def build_calling_on_capture_list():
    return _record_name_list("Calling-On Signals", loaded_calling_on_list)


def build_lc_capture_list():
    return _record_name_list("LC Gates", loaded_lc_list)


def build_track_capture_list():
    return _record_name_list("Tracks", loaded_track_list)


def _upsert_row(ws, key, key_col=1):
    key = str(key).strip()

    for row in range(2, ws.max_row + 2):
        value = ws.cell(row=row, column=key_col).value

        if value in (None, ""):
            ws.cell(row=row, column=key_col).value = key
            return row

        if str(value).strip() == key:
            return row

    row = ws.max_row + 1
    ws.cell(row=row, column=key_col).value = key
    return row


def save_main_signal_coordinate(signal, aspects, coords):
    wb = load_workbook(CONFIG_FILE)

    if "MAIN" not in wb.sheetnames:
        wb.create_sheet("MAIN")

    ws = wb["MAIN"]

    headers = [
        "Signal", "Aspects",
        "Menu_X", "Menu_Y",
        "RED_X", "RED_Y",
        "YELLOW_X", "YELLOW_Y",
        "DOUBLE_YELLOW_X", "DOUBLE_YELLOW_Y",
        "GREEN_X", "GREEN_Y",
        "RouteInit_X", "RouteInit_Y"
    ]

    if ws.max_row == 1 and all(ws.cell(1, c).value is None for c in range(1, 15)):
        for c, value in enumerate(headers, 1):
            ws.cell(1, c).value = value

    row = _upsert_row(ws, signal)
    ws.cell(row, 2).value = aspects

    positions = {
        3: coords[0],   # menu
        5: coords[1],   # red
    }

    if aspects == 2:
        positions[11] = coords[2]       # green
    elif aspects == 3:
        positions[7] = coords[2]        # yellow
        positions[11] = coords[3]       # green
    else:
        positions[7] = coords[2]        # yellow
        positions[9] = coords[3]        # double yellow
        positions[11] = coords[4]       # green

    for col, point in positions.items():
        if point is not None:
            ws.cell(row, col).value = int(point[0])
            ws.cell(row, col + 1).value = int(point[1])

    wb.save(CONFIG_FILE)
    wb.close()
    log(f"{signal} saved successfully to MAIN")


def save_shunt_signal_coordinate(signal, coords):
    wb = load_workbook(CONFIG_FILE)

    if "SHUNT" not in wb.sheetnames:
        wb.create_sheet("SHUNT")

    ws = wb["SHUNT"]

    row = _upsert_row(ws, signal)

    for col, point in zip((2, 4, 6), coords):
        ws.cell(row, col).value = int(point[0])
        ws.cell(row, col + 1).value = int(point[1])

    wb.save(CONFIG_FILE)
    wb.close()
    log(f"{signal} saved successfully to SHUNT")


def save_calling_on_signal_coordinate(signal, coords):
    wb = load_workbook(CONFIG_FILE)

    if "CAL" not in wb.sheetnames:
        wb.create_sheet("CAL")

    ws = wb["CAL"]

    row = _upsert_row(ws, signal)

    for col, point in zip((2, 4), coords):
        ws.cell(row, col).value = int(point[0])
        ws.cell(row, col + 1).value = int(point[1])

    wb.save(CONFIG_FILE)
    wb.close()
    log(f"{signal} saved successfully to CAL")


def save_lc_gate_coordinate(gate, coords):
    wb = load_workbook(CONFIG_FILE)

    if "LC" not in wb.sheetnames:
        wb.create_sheet("LC")

    ws = wb["LC"]

    row = _upsert_row(ws, gate)

    # HAH LC layout: Gate, Menu_X/Y, Open_X/Y, Close_X/Y
    for col, point in zip((2, 4, 6), coords):
        ws.cell(row, col).value = int(point[0])
        ws.cell(row, col + 1).value = int(point[1])

    wb.save(CONFIG_FILE)
    wb.close()
    log(f"{gate} saved successfully to LC")


def save_track_coordinate(track, point):
    wb = load_workbook(CONFIG_FILE)

    if "TRACK_CONFIG" not in wb.sheetnames:
        ws = wb.create_sheet("TRACK_CONFIG")
    else:
        ws = wb["TRACK_CONFIG"]

    # TRACK_CONFIG is intentionally a simple 3-column map:
    # Track | X | Y
    found = False

    for row in range(1, ws.max_row + 1):
        if str(ws.cell(row, 1).value or "").strip() == str(track).strip():
            ws.cell(row, 2).value = int(point[0])
            ws.cell(row, 3).value = int(point[1])
            found = True
            break

    if not found:
        row = ws.max_row + 1
        ws.cell(row, 1).value = str(track).strip()
        ws.cell(row, 2).value = int(point[0])
        ws.cell(row, 3).value = int(point[1])

    wb.save(CONFIG_FILE)
    wb.close()
    log(f"{track} saved successfully to TRACK_CONFIG")


def record_main_signals():
    log("================================")
    log("MAIN SIGNAL COORDINATE RECORDER")
    log("================================")

    build_signal_capture_list()

    if not loaded_signal_list:
        log("No Main Signals To Record")
        return

    for index, signal in enumerate(loaded_signal_list, start=1):

        coords = [None] * 5
        aspects = 2

        steps = ["RECORD MENU"]
        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"MAIN SIGNAL {index} OF {len(loaded_signal_list)}\n\n"
                f"SIGNAL ID : {signal}",
                steps[step]
            )

            if action == "CANCEL":
                return

            if action == "UNDO":
                if step > 0:
                    coords[step - 1] = None
                    step -= 1
                continue

            coords[step] = result

            if step == 0:
                aspects = select_signal_aspects(signal)

                steps.append("RECORD RED")

                if aspects >= 3:
                    steps.append("RECORD YELLOW")

                if aspects == 4:
                    steps.append("RECORD DOUBLE YELLOW")

                steps.append("RECORD GREEN")

            step += 1

        save_main_signal_coordinate(signal, aspects, coords)
        log(f"{signal} GREEN : {coords[2] if aspects == 2 else coords[3] if aspects == 3 else coords[4]}")

    messagebox.showinfo(
        "Completed",
        "Main Signal coordinates saved successfully."
    )


def record_shunt_signals():
    log("=================================")
    log("SHUNT SIGNAL COORDINATE RECORDER")
    log("=================================")

    build_shunt_capture_list()

    if not loaded_shunt_list:
        log("No Shunt Signals To Record")
        return

    for index, signal in enumerate(loaded_shunt_list, start=1):

        coords = [None] * 3
        steps = ["RECORD MENU", "RECORD RED", "RECORD WHITE"]
        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"SHUNT SIGNAL {index} OF {len(loaded_shunt_list)}\n\n"
                f"SIGNAL ID : {signal}",
                steps[step]
            )

            if action == "CANCEL":
                return

            if action == "UNDO":
                if step > 0:
                    coords[step - 1] = None
                    step -= 1
                continue

            coords[step] = result
            step += 1

        save_shunt_signal_coordinate(signal, coords)

    messagebox.showinfo(
        "Completed",
        "Shunt Signal coordinates saved successfully."
    )


def record_calling_on_signals():
    log("=================================")
    log("CALLING-ON SIGNAL COORDINATE RECORDER")
    log("=================================")

    build_calling_on_capture_list()

    if not loaded_calling_on_list:
        log("No Calling-On Signals To Record")
        return

    for index, signal in enumerate(loaded_calling_on_list, start=1):

        coords = [None] * 2
        steps = ["RECORD MENU", "RECORD WHITE"]
        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"CALLING-ON SIGNAL {index} OF {len(loaded_calling_on_list)}\n\n"
                f"SIGNAL ID : {signal}",
                steps[step]
            )

            if action == "CANCEL":
                return

            if action == "UNDO":
                if step > 0:
                    coords[step - 1] = None
                    step -= 1
                continue

            coords[step] = result
            step += 1

        save_calling_on_signal_coordinate(signal, coords)

    messagebox.showinfo(
        "Completed",
        "Calling-On Signal coordinates saved successfully."
    )


def record_lc_gate_coordinate():
    log("================================")
    log("LC GATE COORDINATE RECORDER")
    log("================================")

    build_lc_capture_list()

    if not loaded_lc_list:
        log("No LC Gates To Record")
        return

    for index, gate in enumerate(loaded_lc_list, start=1):

        coords = [None] * 3
        steps = ["RECORD MENU", "RECORD OPEN", "RECORD CLOSE"]
        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"LC GATE {index} OF {len(loaded_lc_list)}\n\n"
                f"LC : {gate}",
                steps[step]
            )

            if action == "CANCEL":
                return

            if action == "UNDO":
                if step > 0:
                    coords[step - 1] = None
                    step -= 1
                continue

            coords[step] = result
            step += 1

        save_lc_gate_coordinate(gate, coords)

    messagebox.showinfo(
        "Completed",
        "LC Gate coordinates saved successfully."
    )


def record_track_coordinate():
    log("================================")
    log("TRACK COORDINATE RECORDER")
    log("================================")

    build_track_capture_list()

    if not loaded_track_list:
        log("No Tracks To Record")
        return

    for index, track in enumerate(loaded_track_list, start=1):

        action, result = capture_point(
            f"TRACK {index} OF {len(loaded_track_list)}\n\n"
            f"TRACK : {track}",
            "RECORD TRACK CENTER"
        )

        if action == "CANCEL":
            return

        if action == "UNDO":
            continue

        save_track_coordinate(track, result)

    messagebox.showinfo(
        "Completed",
        "Track coordinates saved successfully."
    )


def record_calling_on_track_coordinates():
    log("================================")
    log("CALLING-ON TRACK COORDINATE RECORDER")
    log("================================")

    # Calling-On tracks use the same TRACK_CONFIG map.
    record_track_coordinate()


def new_signal():
    """
    NEW SIGNAL:
      1. Create/select the HAH configuration.
      2. Open the full recording menu.
      3. User chooses exactly what to record.
    """
    save_config()

    if not CONFIG_FILE:
        return

    open_record_system_controls()



def open_record_system_controls():

    win = tk.Toplevel(root)
    win.title("Record System Controls")
    win.geometry("620x760")
    win.resizable(False, False)
    win.configure(bg="#e9edf2")
    win.transient(root)
    win.grab_set()

    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    win.geometry(
        f"620x760+{(screen_w - 620) // 2}+{(screen_h - 760) // 2}"
    )

    # Header
    header = tk.Frame(win, bg="#0f172a", height=105)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(
        header,
        text="RECORD SYSTEM CONTROLS",
        font=("Segoe UI", 18, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack(pady=(16, 0))

    tk.Label(
        header,
        text="Save panel coordinates directly into the active HAH configuration",
        font=("Segoe UI", 9),
        bg="#0f172a",
        fg="#94a3b8"
    ).pack(pady=(3, 8))

    # Scrollable body
    content = tk.Frame(win, bg="#e9edf2")
    content.pack(fill="both", expand=True, padx=12, pady=12)

    canvas = tk.Canvas(
        content,
        bg="#e9edf2",
        highlightthickness=0
    )

    scrollbar = ttk.Scrollbar(
        content,
        orient="vertical",
        command=canvas.yview
    )

    body = tk.Frame(canvas, bg="#e9edf2")

    window_id = canvas.create_window(
        (0, 0),
        window=body,
        anchor="nw"
    )

    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def update_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_body(event):
        canvas.itemconfigure(window_id, width=event.width)

    body.bind("<Configure>", update_scroll)
    canvas.bind("<Configure>", resize_body)

    def mouse_wheel(event):
        canvas.yview_scroll(int(-event.delta / 120), "units")

    canvas.bind("<MouseWheel>", mouse_wheel)
    body.bind("<MouseWheel>", mouse_wheel)

    def section(parent, title, subtitle):
        frame = tk.Frame(
            parent,
            bg="white",
            bd=1,
            relief="solid"
        )

        tk.Label(
            frame,
            text=title,
            anchor="w",
            font=("Segoe UI", 11, "bold"),
            bg="white",
            fg="#0f172a"
        ).pack(fill="x", padx=14, pady=(10, 0))

        tk.Label(
            frame,
            text=subtitle,
            anchor="w",
            font=("Segoe UI", 8),
            bg="white",
            fg="#64748b"
        ).pack(fill="x", padx=14, pady=(1, 8))

        return frame

    def action_button(parent, text, command, color):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=color,
            fg="white",
            activebackground=color,
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            font=("Segoe UI", 10, "bold"),
            width=20,
            height=2
        )

    # SIGNALS
    frame = section(
        body,
        "SIGNALS",
        "Record signal menu and indication coordinates"
    )
    frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 8))

    action_button(
        row, "Main Signal", record_main_signals, "#2563eb"
    ).pack(side="left", padx=5, pady=6)

    action_button(
        row, "Shunt Signal", record_shunt_signals, "#2563eb"
    ).pack(side="left", padx=5, pady=6)

    action_button(
        row, "Calling-On Signal", record_calling_on_signals, "#7c3aed"
    ).pack(side="left", padx=5, pady=6)

    # POINT & CH
    frame = section(
        body,
        "POINT & CRANK HANDLE",
        "Record point control, indications and crank-handle coordinates"
    )
    frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 4))

    action_button(
        row, "Point Indication",
        record_point_indication_coordinate,
        "#059669"
    ).pack(side="left", padx=5, pady=6)

    action_button(
        row, "Point Control",
        record_point_control_coordinate,
        "#059669"
    ).pack(side="left", padx=5, pady=6)

    row2 = tk.Frame(frame, bg="white")
    row2.pack(fill="x", padx=8, pady=(0, 8))

    action_button(
        row2, "Crank Handle",
        record_crank_handle_coordinate,
        "#059669"
    ).pack(side="left", padx=5, pady=6)

    # LC & TRACK
    frame = section(
        body,
        "LC GATE & TRACK",
        "Record level-crossing gate and track coordinates"
    )
    frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 8))

    action_button(
        row, "LC Gate",
        record_lc_gate_coordinate,
        "#ea580c"
    ).pack(side="left", padx=5, pady=6)

    action_button(
        row, "Track",
        record_track_coordinate,
        "#ea580c"
    ).pack(side="left", padx=5, pady=6)

    # CALLING-ON TRACK
    frame = section(
        body,
        "CALLING-ON TRACK",
        "Record track coordinates used by Calling-On operations"
    )
    frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 8))

    action_button(
        row,
        "Calling-On Track",
        record_calling_on_track_coordinates,
        "#ea580c"
    ).pack(side="left", padx=5, pady=6)

    # EMERGENCY KEY
    frame = section(
        body,
        "EMERGENCY POINT KEY",
        "Record the Emergency Point Key coordinate"
    )
    frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 8))

    action_button(
        row,
        "Emergency Point Key",
        record_emergency_point_key,
        "#dc2626"
    ).pack(side="left", padx=5, pady=6)

    tk.Label(
        body,
        text="Each recording is saved directly to the active HAH Excel workbook.",
        font=("Segoe UI", 9),
        bg="#e9edf2",
        fg="#64748b"
    ).pack(pady=(2, 12))

    tk.Button(
        win,
        text="CLOSE",
        command=win.destroy,
        bg="#0f172a",
        fg="white",
        activebackground="#1e293b",
        activeforeground="white",
        relief="flat",
        cursor="hand2",
        font=("Segoe UI", 10, "bold"),
        width=18,
        height=2
    ).pack(pady=(0, 14))

    win.bind("<Escape>", lambda e: win.destroy())



# =========================================================
# EDRC AUTO-RUNNER BRIDGE
# =========================================================
#
# The EDRC master launcher expects these EXACT entry points:
#
#     load_excel
#     load_universal_coordinates
#     start_run
#
# It also searches the GUI for buttons with those names.
# The original program only had:
#     IMPORT MASTER TOC
#     LOAD CONFIG
#     START SELECTED ROUTE
#
# Therefore the EDRC launcher could start this program but could
# not feed the TOC/coordinate files or start the test.
#
# These functions are intentionally small adapters. They do NOT
# replace the existing test engine.
# =========================================================

def _edrc_find_file(candidates):
    """Find an EDRC input file without opening a file dialog."""
    search_dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.getcwd(),
    ]

    # EDRC may provide an explicit path through an environment variable.
    for directory in search_dirs:
        for filename in candidates:
            candidate = os.path.abspath(os.path.join(directory, filename))
            if os.path.isfile(candidate):
                return candidate

    return None


def _edrc_load_toc_file(file_path):
    """Load the new TOC into the existing TOC Treeview."""
    global TOC_FILE

    if not file_path or not os.path.isfile(file_path):
        log(f"[EDRC] TOC file not found : {file_path}")
        return False

    wb = None

    try:
        wb = load_workbook(file_path, data_only=True, read_only=True)

        toc_sheet = next(
            (name for name in wb.sheetnames if str(name).strip().upper() == "TOC"),
            None
        )

        if toc_sheet is None:
            log("[EDRC] TOC sheet not found")
            return False

        ws = wb[toc_sheet]
        headers, first_rows = _first_route_rows(ws)

        required_columns = {
            "Signal": _toc_col(headers, "Signal"),
            "Route": _toc_col(headers, "Route"),
            "Track": _toc_col(headers, "Track"),
            "Points": _toc_col(headers, "Points", "Point"),
            "CRANK_HANDLE": _toc_col(headers, "CRANK_HANDLE", "Crank Handle", "CH"),
        }

        missing = [name for name, col in required_columns.items() if col is None]
        if missing:
            log("[EDRC] TOC missing required columns : " + ", ".join(missing))
            return False

        if not first_rows:
            log("[EDRC] TOC contains no first-route records")
            return False

        tree.delete(*tree.get_children())

        for index, (_excel_row_no, row) in enumerate(first_rows, start=1):
            signal = _toc_value(row, headers, "Signal") or ""
            route = _toc_value(row, headers, "Route") or ""
            lock_route = _toc_value(row, headers, "Lock_Route", "Lock Route", "Lock_Routes") or ""
            points = _toc_value(row, headers, "Points", "Point") or ""
            crank_handle = _toc_value(row, headers, "CRANK_HANDLE", "Crank Handle", "CH") or ""

            tree.insert(
                "",
                "end",
                values=(index, signal, route, lock_route, points, crank_handle)
            )

        TOC_FILE = os.path.abspath(file_path)

        items = tree.get_children()
        if items:
            tree.selection_set(items[0])
            tree.focus(items[0])
            tree.see(items[0])

        try:
            status_label.config(text="TOC LOADED", fg="#16a34a")
        except Exception:
            pass

        log("--------------------------------")
        log("[EDRC] TOC LOADED")
        log(f"[EDRC] TOC FILE : {TOC_FILE}")
        log(f"[EDRC] FIRST ROUTE PER SIGNAL : {len(first_rows)}")
        log("[EDRC] TOC FIELDS : Track / Points / CRANK_HANDLE read by header")
        log("--------------------------------")

        return True

    except Exception as e:
        log(f"[EDRC] TOC LOAD ERROR : {type(e).__name__} : {e}")
        return False

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def _edrc_load_coordinate_file(file_path):
    """Load the HAH coordinate workbook without opening a file dialog."""
    global CONFIG_FILE

    if not file_path or not os.path.isfile(file_path):
        log(f"[EDRC] Coordinate file not found : {file_path}")
        return False

    wb = None

    try:
        wb = load_workbook(
            file_path,
            data_only=True,
            read_only=True
        )

        sheets = {
            str(name).strip().upper(): name
            for name in wb.sheetnames
        }

        # Only POINT and CH must be in the coordinate workbook.
        # TRACK_CONFIG is optional - track coordinates come from the
        # EDRC track file (TRACK_COORDINATES_CAPTURED.xlsx).
        # SYSTEM_CONTROLS is checked separately below.
        required = [
            "POINT",
            "CH"
        ]

        missing = [
            name for name in required
            if name not in sheets
        ]

        log(
            "[EDRC] COORDINATE WORKBOOK : "
            + os.path.abspath(file_path)
        )
        log(
            "[EDRC] SHEETS FOUND : "
            + ", ".join(str(name) for name in wb.sheetnames)
        )

        if missing:
            log(
                "[EDRC] Coordinate workbook missing sheets : "
                + ", ".join(missing)
            )
            return False

        # TRACK_CONFIG is optional here. When it is missing or empty the
        # track coordinates are read from the EDRC track file instead.
        if "TRACK_CONFIG" in sheets:

            track_ws = wb[sheets["TRACK_CONFIG"]]
            valid_track_rows = 0

            for row in track_ws.iter_rows(values_only=True):
                if len(row) < 3:
                    continue

                name = str(row[0]).strip().upper() if row[0] is not None else ""

                if name == "TRACK":
                    continue

                if name and row[1] is not None and row[2] is not None:
                    valid_track_rows += 1

            if valid_track_rows == 0:
                log(
                    "[EDRC] TRACK_CONFIG empty - tracks will be read from "
                    "the EDRC track file"
                )
        else:
            log(
                "[EDRC] No TRACK_CONFIG sheet - tracks will be read from "
                "the EDRC track file"
            )

        # Emergency Point Key coordinate (SYSTEM_CONTROLS sheet).
        emergency_found = False

        if "SYSTEM_CONTROLS" in sheets:

            control_ws = wb[sheets["SYSTEM_CONTROLS"]]

            for row in control_ws.iter_rows(min_row=2, values_only=True):
                if not row:
                    continue

                name = _norm_id(row[0])

                if name == _norm_id("EMERGENCY POINT KEY"):
                    emergency_found = (
                        len(row) >= 3 and
                        row[1] is not None and
                        row[2] is not None
                    )
                    break

        if not emergency_found:
            log(
                "[EDRC] EMERGENCY POINT KEY coordinate not found in "
                "SYSTEM_CONTROLS - record it with the universal yard "
                "recorder (asked after the crank handles)"
            )
            return False

        CONFIG_FILE = os.path.abspath(file_path)

        try:
            status_label.config(
                text="CONFIG LOADED",
                fg="#16a34a"
            )
        except Exception:
            pass

        log("--------------------------------")
        log("[EDRC] COORDINATES LOADED")
        log(f"[EDRC] CONFIG FILE : {CONFIG_FILE}")
        log("--------------------------------")

        return True

    except Exception as e:
        log(
            f"[EDRC] COORDINATE LOAD ERROR : "
            f"{type(e).__name__} : {e}"
        )
        return False

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def load_excel(file_path=None):
    """
    EDRC entry point #1.

    Loads TOC.xlsx automatically. No file dialog is opened.

    The optional file_path is supported so the EDRC launcher can call:
        load_excel(path)
    or:
        load_excel()
    """
    if file_path is None:
        file_path = (
            os.environ.get("EDRC_LIST")
            or os.environ.get("EDRC_TOC_FILE")
        )

    if file_path:
        log("[EDRC] TOC SOURCE : EDRC_LIST")
    else:
        file_path = _edrc_find_file([
            "TOC.xlsx",
            "TOC.xlsm"
        ])
        log("[EDRC] TOC SOURCE : local fallback")

    log("[EDRC] load_excel() called")
    log(f"[EDRC] TOC FILE : {file_path}")

    return _edrc_load_toc_file(file_path)


def load_universal_coordinates(file_path=None):
    """
    EDRC entry point #2.

    Loads the HAH coordinate workbook automatically. Supports both
    HAH COORDINATES.xlsx and the existing HAH COORDINATES1.xlsx naming.
    """
    if file_path is None:
        file_path = (
            os.environ.get("EDRC_COORDS")
            or os.environ.get("EDRC_COORDINATES_FILE")
        )

    if file_path:
        log("[EDRC] COORDINATE SOURCE : EDRC_COORDS")
    else:
        file_path = _edrc_find_file([
            "HAH COORDINATES.xlsx",
            "HAH COORDINATES1.xlsx",
            "HAH COORDINATES.xlsm"
        ])
        log("[EDRC] COORDINATE SOURCE : local fallback")

    log("[EDRC] load_universal_coordinates() called")
    log(f"[EDRC] COORDINATE FILE : {file_path}")

    return _edrc_load_coordinate_file(file_path)


def start_run():
    """
    EDRC entry point #3.

    Starts the existing POINT/CRANK HANDLE/EMERGENCY POINT KEY
    test engine using the first loaded TOC row.

    The actual test logic remains inside test_point_operation().
    """
    global running

    log("[EDRC] start_run() called")

    if running:
        log("[EDRC] START IGNORED : test is already running")
        return True

    if not TOC_FILE:
        log("[EDRC] START BLOCKED : TOC not loaded")
        return False

    if not CONFIG_FILE or not os.path.isfile(CONFIG_FILE):
        log("[EDRC] START BLOCKED : coordinates not loaded")
        return False

    items = tree.get_children()

    if not items:
        log("[EDRC] START BLOCKED : TOC has no rows")
        return False

    # Walk-away mode: automatically use the first TOC route when
    # no operator has manually selected a route.
    selected = tree.selection()

    if not selected:
        tree.selection_set(items[0])
        tree.focus(items[0])
        tree.see(items[0])

    log("[EDRC] STARTING EXISTING TEST ENGINE")
    log("[EDRC] MODE : WALK-AWAY / FIRST TOC ROUTE")

    # IMPORTANT:
    # The EDRC master waits for this module-level flag to become True.
    # Set it BEFORE root.after() because the master checks it immediately
    # after pressing start_run().
    running = True
    log("[EDRC] RUNNING FLAG = True")
    log(f"[EDRC] START ACCEPTED : TOC={TOC_FILE}")
    log(f"[EDRC] START ACCEPTED : CONFIG={CONFIG_FILE}")
    log("[EDRC] QUEUING TEST ENGINE ON TK MAIN THREAD")

    # test_point_operation() uses Tk dialogs and GUI widgets, so it must
    # execute on Tk's main thread. after() schedules it safely.
    root.after(100, test_point_operation)

    return True


def _edrc_autoload_and_start():
    """
    Optional automatic fallback.

    If the EDRC launcher does not click the three bridge buttons, the
    program can still start from the standard project files.
    This is deliberately delayed until the Tk GUI is fully initialized.
    """
    try:
        toc_ok = load_excel()
        coord_ok = load_universal_coordinates()

        if toc_ok and coord_ok:
            start_run()
        else:
            log(
                "[EDRC] AUTO START NOT PERFORMED : "
                f"TOC={toc_ok}, COORDS={coord_ok}"
            )
    except Exception as e:
        log(f"[EDRC] AUTO START ERROR : {e}")


# =========================================================
# GUI
# =========================================================

root = tk.Tk()

create_capture_overlay()

capture_overlay.withdraw()
root.title(
    f"{APP_NAME}  |  Version {VERSION}"
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

# =========================================================
# LEFT LOGO
# =========================================================

left_logo_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

left_logo_frame.pack(
    side="left",
    padx=20,
    pady=10
)

try:

    img = Image.open("indian_railways.png")

    # Resize logo to fit the header
    img = img.resize((75, 75), Image.LANCZOS)

    railway_logo = ImageTk.PhotoImage(img)

    tk.Label(
        left_logo_frame,
        image=railway_logo,
        bg="#0f172a",
        bd=0
    ).pack()

except Exception as e:

    print(e)

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

    text="Testing for Track-locking of Points",

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

# =========================================================
# RIGHT LOGO
# =========================================================

right_logo_frame = tk.Frame(
    header_frame,
    bg="#0f172a"
)

right_logo_frame.pack(
    side="right",
    padx=20,
    pady=10
)

try:

    img = Image.open("company_logo.png")

    # Resize company logo
    img = img.resize((75, 75), Image.LANCZOS)

    company_logo = ImageTk.PhotoImage(img)

    tk.Label(
        right_logo_frame,
        image=company_logo,
        bg="#0f172a",
        bd=0
    ).pack()

except Exception as e:

    print(e)

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

    text="NEW CONFIG",

    command=new_signal,

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

    text="LOAD CONFIG",

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
# EDRC AUTO CONTROL BUTTONS
# =========================================================
#
# Keep these exact button captions. The EDRC master launcher finds
# controls by these names.
# =========================================================

edrc_frame = tk.Frame(
    left_panel,
    bg="#f1f5f9",
    bd=1,
    relief="solid"
)
edrc_frame.pack(
    fill="x",
    padx=10,
    pady=(5, 10)
)

tk.Label(
    edrc_frame,
    text="EDRC AUTO CONTROL",
    font=("Segoe UI", 9, "bold"),
    bg="#f1f5f9",
    fg="#334155"
).pack(pady=(6, 3))

tk.Button(
    edrc_frame,
    text="load_excel",
    command=load_excel,
    bg="#475569",
    fg="white",
    activebackground="#334155",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=2)

tk.Button(
    edrc_frame,
    text="load_universal_coordinates",
    command=load_universal_coordinates,
    bg="#475569",
    fg="white",
    activebackground="#334155",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=2)

tk.Button(
    edrc_frame,
    text="start_run",
    command=start_run,
    bg="#7c3aed",
    fg="white",
    activebackground="#6d28d9",
    activeforeground="white",
    relief="flat",
    cursor="hand2",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=(2, 6))

# =========================================================
# OTHER BUTTONS
# =========================================================
create_button(
    "IMPORT MASTER TOC",
    load_toc,
    "#2563eb"
).pack(pady=5)

create_button(
    "RECORD SYSTEM CONTROLS",
    open_record_system_controls,
    "#059669"
).pack(pady=5)
create_button(
    "START SELECTED ROUTE",
    test_point_operation,
    "#7c3aed"
).pack(pady=8)

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

    text="TOC",

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
        "POINT",
        "CRANK_HANDLE"
    ),

    show="headings",

    yscrollcommand=tree_scroll.set

)

tree_scroll.config(
    command=tree.yview
)

tree.heading("NO", text="NO")

tree.heading("SIGNAL", text="MAIN SIGNAL")

tree.heading("ROUTE", text="MAIN ROUTE")

tree.heading("LOCK_ROUTE", text="LOCK ROUTES")
tree.heading("POINT", text="POINT")
tree.heading("CRANK_HANDLE", text="CRANK HANDLE")

tree.column("POINT", width=200, anchor="center")
tree.column("CRANK_HANDLE", width=150, anchor="center")
tree.column("NO", width=80, anchor="center")

tree.column("SIGNAL", width=200, anchor="center")

tree.column("ROUTE", width=220, anchor="center")

tree.column("LOCK_ROUTE", width=500, anchor="w")

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

    text="SPACE = CAPTURE | BACKSPACE = UNDO | ESC = CANCEL",

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
