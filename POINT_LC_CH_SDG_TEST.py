import re
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
import subprocess
import pyautogui

import win32api
import win32con

import uiautomation as auto

from pywinauto import Desktop

from PIL import Image, ImageTk, ImageGrab
from openpyxl import Workbook
from openpyxl import load_workbook

from datetime import datetime


# =========================================================
# EMBEDDED REPORT MODULE
# PASS = NO COLOR
# FAIL = RED ONLY
# =========================================================

# EDRC supplies the exact session/program report folder through the
# EDRC_PROGRAM_DIR environment variable.  Keep Reports only as a manual-run
# fallback; never create/use it when launched by EDRC.
EDRC_PROGRAM_DIR = os.environ.get("EDRC_PROGRAM_DIR", "").strip()
# Report goes to the EDRC session folder when the launcher supplies
# one, otherwise to the working folder - never into a "Reports"
# sub-folder, where the launcher does not look.
REPORT_FOLDER = (
    EDRC_PROGRAM_DIR
    or os.environ.get("EDRC_REPORT_DIR", "").strip()
    or os.getcwd()
)
REPORT_FILE = None
PASS_COUNT = 0
FAIL_COUNT = 0

CURRENT_RESULT = {
    "signal": "",
    "route": "",
    "lc_gate": "PASS",
    "point": "PASS",
    "crank": "PASS",
    "failed_tests": [],
    "total_tests": 0
}

HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=12)

FAIL_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
FAIL_FONT = Font(color="9C0006", bold=False)

CENTER = Alignment(horizontal="center", vertical="center")
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

def create_report():
    global REPORT_FILE, PASS_COUNT, FAIL_COUNT, CURRENT_RESULT

    PASS_COUNT = 0
    FAIL_COUNT = 0

    CURRENT_RESULT = {
        "signal": "",
        "route": "",
        "lc_gate": "PASS",
        "point": "PASS",
        "crank": "PASS",
        "failed_tests": [],
        "failed_details": [],
        "total_tests": 0
    }

    os.makedirs(REPORT_FOLDER, exist_ok=True)

    REPORT_FILE = os.path.join(
        REPORT_FOLDER,
        datetime.now().strftime(
            "TEST FOR LC, POINT,CH & SDG Report_%Y-%m-%d_%H-%M-%S.xlsx"
        )
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Automation Report"

    summary_ws = wb.create_sheet("CONSOLIDATED SUMMARY")
    summary_ws.append(["Route", "Total Tests", "Final Result"])

    for cell in summary_ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    ws.merge_cells("A1:G1")
    title = ws["A1"]
    title.value = "TEST FOR LC, POINT, CH & SDG REPORT"
    title.font = Font(bold=True, color="FFFFFF", size=16)
    title.fill = HEADER_FILL
    title.alignment = CENTER
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    date_cell = ws["A2"]
    date_cell.value = (
        "Generated On : "
        + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )
    date_cell.font = Font(bold=True)
    date_cell.alignment = CENTER
    ws.row_dimensions[2].height = 22

    headers = [
        "Signal",
        "Route",
        "LC Gate",
        "Point",
        "Crank Handle",
        "Result",
        "Date & Time"
    ]

    ws.append(headers)

    for cell in ws[3]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    ws.freeze_panes = "A4"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 22

    wb.save(REPORT_FILE)
    wb.close()

    print("REPORT CREATED")
    print(REPORT_FILE)

    return REPORT_FILE


def save_result(signal, route, test, expected, actual, result, remarks=""):
    global REPORT_FILE, CURRENT_RESULT

    if REPORT_FILE is None:
        create_report()

    CURRENT_RESULT["signal"] = signal
    CURRENT_RESULT["route"] = route

    test_upper = str(test).upper()
    result = str(result).upper().strip()

    # Ignore recovery/receive entries in the consolidated route result.
    if "RECOVERY" in test_upper or "RECEIVE" in test_upper:
        return

    CURRENT_RESULT["total_tests"] += 1

    if "LC GATE" in test_upper:
        if result == "FAIL":
            CURRENT_RESULT["lc_gate"] = "FAIL"
            CURRENT_RESULT["failed_tests"].append(f"{test} FAILED")

    elif "POINT" in test_upper:
        if result == "FAIL":
            CURRENT_RESULT["point"] = "FAIL"
            CURRENT_RESULT["failed_tests"].append(f"{test} FAILED")

    elif "CRANK" in test_upper:
        if result == "FAIL":
            CURRENT_RESULT["crank"] = "FAIL"
            CURRENT_RESULT["failed_tests"].append(f"{test} FAILED")

    # Keep the exact failure information for the separate FAILED CASES sheet.
    if result == "FAIL":
        CURRENT_RESULT["failed_details"].append({
            "test": str(test),
            "expected": str(expected),
            "actual": str(actual),
            "remarks": str(remarks)
        })


def flush_route_result():
    global REPORT_FILE, CURRENT_RESULT

    if REPORT_FILE is None or CURRENT_RESULT["route"] == "":
        return

    wb = load_workbook(REPORT_FILE)

    ws = wb["Automation Report"]
    summary_ws = wb["CONSOLIDATED SUMMARY"]

    overall = "PASS"

    if (
        CURRENT_RESULT["lc_gate"] == "FAIL"
        or CURRENT_RESULT["point"] == "FAIL"
        or CURRENT_RESULT["crank"] == "FAIL"
    ):
        overall = "FAIL"

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws.append([
        CURRENT_RESULT["signal"],
        CURRENT_RESULT["route"],
        CURRENT_RESULT["lc_gate"],
        CURRENT_RESULT["point"],
        CURRENT_RESULT["crank"],
        overall,
        current_time
    ])

    row = ws.max_row

    # Normal/pass cells remain unfilled.
    for col in range(1, 8):
        cell = ws.cell(row=row, column=col)
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font(color="000000", bold=False)
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    # Only a failed route is red.
    if overall == "FAIL":
        for col in range(1, 8):
            cell = ws.cell(row=row, column=col)
            cell.fill = FAIL_FILL
            cell.font = FAIL_FONT

    if len(CURRENT_RESULT["failed_tests"]) == 0:
        final_result = "ALL TEST CASES PASSED"
    else:
        final_result = ", ".join(CURRENT_RESULT["failed_tests"])

    summary_ws.append([
        CURRENT_RESULT["route"],
        CURRENT_RESULT["total_tests"],
        final_result
    ])

    row = summary_ws.max_row

    for cell in summary_ws[row]:
        cell.fill = PatternFill(fill_type=None)
        cell.font = Font(color="000000", bold=False)
        cell.border = THIN_BORDER
        cell.alignment = CENTER

    # Only failed summary entries are red.
    if "PASSED" not in str(final_result).upper():
        result_cell = summary_ws.cell(row=row, column=3)
        result_cell.fill = FAIL_FILL
        result_cell.font = FAIL_FONT

    CURRENT_RESULT = {
        "signal": "",
        "route": "",
        "lc_gate": "PASS",
        "point": "PASS",
        "crank": "PASS",
        "failed_tests": [],
        "failed_details": [],
        "total_tests": 0
    }

    wb.save(REPORT_FILE)
    wb.close()


def _build_route_status_summary(wb, ws):
    """Add the requested route-count summary to the FIRST report sheet.

    Total routes are taken from the active TOC. Initiated routes are the
    unique main Signal+Route combinations that actually reached the report
    sheet. Any TOC main route missing from the report is listed as a
    non-initiated route.
    """
    global TOC_FILE

    toc_routes = []
    initiated_routes = []

    # Read the active TOC using its Signal/Route headers.
    try:
        if TOC_FILE and os.path.exists(TOC_FILE):
            toc_wb = load_workbook(TOC_FILE, data_only=True, read_only=True)
            if "TOC" in toc_wb.sheetnames:
                toc_ws = toc_wb["TOC"]
                headers = get_toc_header_map(toc_ws)
                signal_col = toc_col(headers, "Signal")
                route_col = toc_col(headers, "Route")

                if signal_col is not None and route_col is not None:
                    seen = set()
                    for row in toc_ws.iter_rows(min_row=2, values_only=True):
                        signal = str(row[signal_col] or "").strip()
                        route = str(row[route_col] or "").strip()
                        if not signal or not route:
                            continue
                        key = (signal.upper(), route.upper())
                        if key not in seen:
                            seen.add(key)
                            toc_routes.append((signal, route))
            toc_wb.close()
    except Exception as e:
        log(f"REPORT ROUTE SUMMARY TOC READ ERROR : {e}")

    # Read routes that were actually written to the first report sheet.
    try:
        seen = set()
        for row in ws.iter_rows(min_row=4, values_only=True):
            signal = str(row[0] or "").strip() if len(row) > 0 else ""
            route = str(row[1] or "").strip() if len(row) > 1 else ""
            if not signal or not route:
                continue
            key = (signal.upper(), route.upper())
            if key not in seen:
                seen.add(key)
                initiated_routes.append((signal, route))
    except Exception as e:
        log(f"REPORT ROUTE SUMMARY READ ERROR : {e}")

    toc_keys = {(s.upper(), r.upper()) for s, r in toc_routes}
    initiated_main = [item for item in initiated_routes if (item[0].upper(), item[1].upper()) in toc_keys]
    initiated_keys = {(s.upper(), r.upper()) for s, r in initiated_main}

    non_initiated = [
        f"{signal}/{route}"
        for signal, route in toc_routes
        if (signal.upper(), route.upper()) not in initiated_keys
    ]

    # Add the three requested lines below the existing report table.
    start_row = ws.max_row + 2

    summary_lines = [
        ("TOTAL NO. OF ROUTES", len(toc_routes)),
        ("TOTAL NO. OF INITIATED ROUTES", len(initiated_main)),
        ("NON-INITIATED ROUTE", ", ".join(non_initiated) if non_initiated else "NONE"),
    ]

    for offset, (label, value) in enumerate(summary_lines):
        row_no = start_row + offset
        ws.cell(row=row_no, column=1).value = label
        ws.cell(row=row_no, column=1).font = Font(bold=True)
        ws.cell(row=row_no, column=1).alignment = CENTER
        ws.cell(row=row_no, column=1).border = THIN_BORDER

        ws.merge_cells(start_row=row_no, start_column=2, end_row=row_no, end_column=7)
        value_cell = ws.cell(row=row_no, column=2)
        value_cell.value = value
        value_cell.alignment = CENTER
        value_cell.border = THIN_BORDER

    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 10, 28)
    log(f"TOTAL NO. OF ROUTES : {len(toc_routes)}")
    log(f"TOTAL NO. OF INITIATED ROUTES : {len(initiated_main)}")
    log(f"NON-INITIATED ROUTE : {', '.join(non_initiated) if non_initiated else 'NONE'}")


def finalize_report():
    global REPORT_FILE

    if REPORT_FILE is None:
        return

    try:
        wb = load_workbook(REPORT_FILE)
        ws = wb["Automation Report"]

        # Keep the summary on the FIRST sheet and remove any old failed-case
        # sheet if a report is reopened from an older workbook template.
        if "FAILED CASES" in wb.sheetnames:
            del wb["FAILED CASES"]

        _build_route_status_summary(wb, ws)
        wb.save(REPORT_FILE)
        wb.close()
    except Exception as e:
        log(f"REPORT FINAL SUMMARY ERROR : {e}")

    print("=" * 60)
    print("REPORT SAVED SUCCESSFULLY")
    print(REPORT_FILE)
    print("=" * 60)


# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = None

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

CURRENT_TEST = "ROUTE ENGINE"

TOC_FILE = None

SIMULATOR_RUNNING = False

PANEL_RUNNING = False
# =========================================================
# GLOBALS
# =========================================================

signals = {}
loaded_ch_list = []
auto_throw_data = []
loaded_signal_list = []
loaded_shunt_list = []
loaded_calling_on_list = []
loaded_point_list = []
loaded_lc_list = []
points = {}
running = False


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

def build_shunt_capture_list():

    global loaded_shunt_list

    loaded_shunt_list.clear()

    all_names = set()

    total = prompt_count("Shunt Signals")

    for i in range(total):

        name = prompt_name(
            "Shunt Signal",
            i,
            total,
            all_names
        )

        all_names.add(name)

        loaded_shunt_list.append(name)

    log(f"Total Shunt Signals : {len(loaded_shunt_list)}")

    for signal in loaded_shunt_list:
        log(f"SHUNT SIGNAL : {signal}")

    return loaded_shunt_list

def build_ch_capture_list():
    global loaded_ch_list

    loaded_ch_list.clear()

    all_names = set()

    n = prompt_count("Crank Handles")

    for i in range(n):
        name = prompt_name(
            "Point",
            i,
            n,
            all_names
        )

        all_names.add(name)

        loaded_ch_list.append(name)

    log(f"Total Crank Handles : {len(loaded_ch_list)}")

    for point in loaded_ch_list:
        log(f"POINT : {point}")

    return loaded_ch_list

def build_calling_on_capture_list():

    global loaded_calling_on_list

    loaded_calling_on_list.clear()

    all_names = set()

    total = prompt_count("Calling-On Signals")

    for i in range(total):

        name = prompt_name(
            "Calling-On Signal",
            i,
            total,
            all_names
        )

        all_names.add(name)

        loaded_calling_on_list.append(name)

    log(f"Total Calling-On Signals : {len(loaded_calling_on_list)}")

    for signal in loaded_calling_on_list:
        log(f"CALLING-ON SIGNAL : {signal}")

    return loaded_calling_on_list

def build_lc_capture_list():

    global loaded_lc_list

    loaded_lc_list.clear()

    all_names = set()

    n = prompt_count("LC Gates")

    for i in range(n):

        name = prompt_name(
            "LC Gate (ENTER EXACT TOC NAME)",
            i,
            n,
            all_names
        )

        if not name:
            continue

        name = str(name).strip()

        all_names.add(name)

        loaded_lc_list.append(name)

    log(f"Total LC Gates : {len(loaded_lc_list)}")

    for lc in loaded_lc_list:

        log(f"LC : {lc}")

    return loaded_lc_list

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

    global capture_overlay
    global overlay_progress_label
    global overlay_signal_label
    global overlay_step_label
    global overlay_hint_label
    global overlay_button_frame

    # Track/LC/Point/Crank/Calling-On recorders can call capture_point()
    # directly, so create the shared capture overlay if it does not exist.
    if capture_overlay is None or not capture_overlay.winfo_exists():
        create_capture_overlay()

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
    for widget in overlay_button_frame.winfo_children():
        widget.destroy()
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

    if capture_overlay is None:
        return

    if not capture_overlay.winfo_viewable():
        return

    capture_result = pyautogui.position()
    capture_action = "SPACE"


def undo_capture():

    global capture_action

    if capture_overlay is None:
        return

    if not capture_overlay.winfo_viewable():
        return

    capture_action = "UNDO"

def capture_cancel():

    global capture_action
    if capture_overlay is None:
        return

    if not capture_overlay.winfo_viewable():
        return

    capture_action = "CANCEL"
selected_aspects = None

def select_signal_aspects(signal):

    global selected_aspects

    selected_aspects = None

    overlay_step_label.config(text="SELECT NUMBER OF ASPECTS")

    overlay_hint_label.config(
        text=f"{signal}\n\nClick 2, 3 or 4"
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
            font=("Segoe UI", 16, "bold"),
            command=lambda v=value: choose(v)
        ).pack(side="left", padx=8)

    capture_overlay.deiconify()
    capture_overlay.lift()

    while selected_aspects is None:
        root.update()
        time.sleep(0.02)

    for widget in overlay_button_frame.winfo_children():
        widget.destroy()

    return selected_aspects
# =========================================================
# KEYBOARD
# =========================================================


keyboard.add_hotkey("space", capture_space)
keyboard.add_hotkey("backspace", undo_capture)
keyboard.add_hotkey("esc", capture_cancel)


def record_signal_coordinate():

    global capture_overlay

    if capture_overlay is None:
        create_capture_overlay()

    record_main_signals()
    record_shunt_signals()
    record_calling_on_signals()
    record_point_control_coordinate()
    record_crank_handle_coordinate()
    record_lc_gate_coordinate()


def build_signal_capture_list():

    global loaded_signal_list

    loaded_signal_list.clear()

    all_names = set()

    total = prompt_count("Main Signals")

    for i in range(total):

        name = prompt_name(
            "Main Signal",
            i,
            total,
            all_names
        )

        all_names.add(name)

        loaded_signal_list.append(name)

    log(f"Total Main Signals : {len(loaded_signal_list)}")

    for signal in loaded_signal_list:
        log(f"SIGNAL : {signal}")

    return loaded_signal_list

def new_signal():
    save_config()

    if CONFIG_FILE:
        record_signal_coordinate()

def existing_signal():

    load_config()

    if not CONFIG_FILE:
        return

    log("Existing Configuration Loaded")

    status_label.config(
        text="CONFIGURATION LOADED",
        fg="#16a34a"
    )


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

        steps = [
            "RECORD MENU"
        ]
        print(id(steps))

        step = 0
        print("Before:", steps)

        while step < len(steps):

            action, result = capture_point(

                f"MAIN SIGNAL {index} OF {len(loaded_signal_list)}\n\nSIGNAL ID : {signal}",

                steps[step]

            )

            if action == "CANCEL":
                return


            if action == "UNDO":

                if step > 0:
                    coords[step - 1] = None
                    step -= 1

                continue

            if step >= len(coords):
                print("STEP =", step)
                print("COORDS =", len(coords))
                print("STEPS =", steps)
                break

            coords[step] = result

            # MENU RECORDED
            if step == 0:

                aspects = select_signal_aspects(signal)

                steps.append("RECORD RED")

                if aspects >= 3:
                    steps.append("RECORD YELLOW")

                if aspects == 4:
                    steps.append("RECORD DOUBLE YELLOW")

                steps.append("RECORD GREEN")
                print("After:", steps)

            step += 1

        log(f"Final Coordinates : {coords}")

        save_main_signal_coordinate(
            signal,
            aspects,
            coords
        )

        if aspects == 2:
            green = coords[2]
        elif aspects == 3:
            green = coords[3]
        else:
            green = coords[4]

        log(f"{signal} GREEN : {green}")
    log("Main Signal Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Main Signal Coordinates Saved Successfully."
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

        steps = [
            "RECORD MENU",
            "RECORD RED",
            "RECORD WHITE"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(
                f"SHUNT SIGNAL {index} OF {len(loaded_shunt_list)}\n\nSIGNAL ID : {signal}",
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

        log(f"Final Coordinates : {coords}")

        save_shunt_signal_coordinate(
            signal,
            coords
        )

        log(f"{signal} MENU  : {coords[0]}")
        log(f"{signal} RED   : {coords[1]}")
        log(f"{signal} WHITE : {coords[2]}")

    log("Shunt Signal Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Shunt Signal Coordinates Saved Successfully."
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

        steps = [
            "RECORD MENU",
            "RECORD WHITE"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(

                f"CALLING-ON SIGNAL {index} OF {len(loaded_calling_on_list)}\n\nSIGNAL ID : {signal}",

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

        log(f"Final Coordinates : {coords}")

        save_calling_on_signal_coordinate(
            signal,
            coords
        )

        log(f"{signal} MENU  : {coords[0]}")
        log(f"{signal} WHITE : {coords[1]}")

        # Calling-On signals require their associated track coordinate as well.
        # Record it once and keep it in TRACK_CONFIG for wait_until_track_red().
        calling_tracks = [
            item["track"] for item in load_calling_on_track_entries()
            if item["signal"] == str(signal).strip().upper() and item["track"]
        ]
        for track in dict.fromkeys(calling_tracks):
            if track.strip().upper() in get_recorded_track_names():
                log(f"Calling-On Track already recorded : {track}")
                continue

            action, track_result = capture_point(
                f"CALLING-ON SIGNAL : {signal}\n\nTRACK : {track}",
                "RECORD TRACK CENTER"
            )

            if action == "CANCEL":
                return

            if action == "SPACE" and track_result is not None:
                tx, ty = track_result
                save_track_coordinate(track, tx, ty)
                log(f"Calling-On Track Saved : {track} -> ({tx}, {ty})")

    log("Calling-On Signal Coordinates Saved")

    messagebox.showinfo(
        "Completed",
        "Calling-On Signal Coordinates Saved Successfully."
    )

def save_main_signal_coordinate(signal, aspects, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["MAIN"]

    for row in range(2, sheet.max_row + 2):

        if sheet.cell(row=row, column=1).value in (None, ""):

            sheet.cell(row=row, column=1).value = signal
            break

        if str(sheet.cell(row=row, column=1).value).strip() == str(signal).strip():
            break

    sheet.cell(row=row, column=2).value = aspects

    # -------------------------
    # MENU
    # -------------------------

    sheet.cell(row=row, column=3).value = coords[0][0]
    sheet.cell(row=row, column=4).value = coords[0][1]

    # -------------------------
    # RED
    # -------------------------

    sheet.cell(row=row, column=5).value = coords[1][0]
    sheet.cell(row=row, column=6).value = coords[1][1]

    # -------------------------
    # GREEN POSITION
    # -------------------------

    if aspects == 2:

        green = coords[2]

    elif aspects == 3:

        sheet.cell(row=row, column=7).value = coords[2][0]
        sheet.cell(row=row, column=8).value = coords[2][1]

        green = coords[3]

    else:   # 4 Aspects

        sheet.cell(row=row, column=7).value = coords[2][0]
        sheet.cell(row=row, column=8).value = coords[2][1]

        sheet.cell(row=row, column=9).value = coords[3][0]
        sheet.cell(row=row, column=10).value = coords[3][1]

        green = coords[4]

    sheet.cell(row=row, column=11).value = green[0]
    sheet.cell(row=row, column=12).value = green[1]

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{signal} saved successfully.")

def save_shunt_signal_coordinate(signal, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["SHUNT"]

    for row in range(2, sheet.max_row + 2):

        if sheet.cell(row=row, column=1).value in (None, ""):
            sheet.cell(row=row, column=1).value = signal
            break

        if str(sheet.cell(row=row, column=1).value).strip() == str(signal).strip():
            break

    sheet.cell(row=row, column=2).value = coords[0][0]
    sheet.cell(row=row, column=3).value = coords[0][1]

    sheet.cell(row=row, column=4).value = coords[1][0]
    sheet.cell(row=row, column=5).value = coords[1][1]

    sheet.cell(row=row, column=6).value = coords[2][0]
    sheet.cell(row=row, column=7).value = coords[2][1]

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{signal} saved successfully.")

def pause_sleep(seconds):

    end_time = time.time() + seconds

    while True:

        pause_event.wait()

        remaining = end_time - time.time()

        if remaining <= 0:
            break

        time.sleep(min(0.1, remaining))


def save_calling_on_signal_coordinate(signal, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["CAL"]

    for row in range(2, sheet.max_row + 2):

        if sheet.cell(row=row, column=1).value in (None, ""):

            sheet.cell(row=row, column=1).value = signal
            break

        if str(sheet.cell(row=row, column=1).value).strip() == str(signal).strip():
            break

    sheet.cell(row=row, column=2).value = coords[0][0]
    sheet.cell(row=row, column=3).value = coords[0][1]

    sheet.cell(row=row, column=4).value = coords[1][0]
    sheet.cell(row=row, column=5).value = coords[1][1]

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{signal} saved successfully.")



# =========================================================
# PAUSE CONTROL
# =========================================================

paused = False

pause_event = threading.Event()

pause_event.set()


from tkinter import filedialog
from openpyxl import load_workbook

# =========================================================
# TOC HEADER MAP - NEW TL TOC COMPATIBILITY
# =========================================================
def _norm_toc_header(value):
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]+", "_", str(value).strip().upper()).strip("_")

def get_toc_header_map(ws):
    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return {_norm_toc_header(h): i for i, h in enumerate(headers) if _norm_toc_header(h)}

def toc_col(headers, *names):
    for name in names:
        key = _norm_toc_header(name)
        if key in headers:
            return headers[key]
    return None

def toc_value(row, headers, *names):
    index = toc_col(headers, *names)
    if index is None or index >= len(row):
        return None
    return row[index]

def load_toc(file=None, show_message=True):
    global TOC_FILE

    if file is None:
        file = filedialog.askopenfilename(
            title="Select TOC",
            filetypes=[("Excel Files", "*.xlsx")]
        )

    if not file:
        return

    if not os.path.exists(file):
        log(f"TOC file not found: {file}")
        return

    TOC_FILE = file

    tree.delete(*tree.get_children())

    wb = load_workbook(file, data_only=True)
    ws = wb["TOC"]
    headers = get_toc_header_map(ws)
    signal_col = toc_col(headers, "Signal")
    route_col = toc_col(headers, "Route")
    lock_col = toc_col(headers, "Lock_Route", "Lock Route", "Lock_Routes")

    if signal_col is None or route_col is None:
        wb.close()
        log("TOC LOAD ERROR : Signal/Route headers not found")
        if show_message:
            messagebox.showerror("TOC Error", "Signal and Route headers are required.")
        return

    count = 1

    for row in ws.iter_rows(min_row=2, values_only=True):

        signal = row[signal_col] if signal_col < len(row) else None
        route = row[route_col] if route_col < len(row) else None
        lock_routes = row[lock_col] if lock_col is not None and lock_col < len(row) else ""

        tree.insert(
            "",
            "end",
            values=(
                count,
                signal,
                route,
                lock_routes
            )
        )

        count += 1

    wb.close()

    log("TOC Loaded Successfully")
    log("TOC HEADER MAP : " + ", ".join(f"{k}={v+1}" for k, v in sorted(headers.items())))
    log(
        "TOC ACTIVE FIELDS : "
        f"Signal={signal_col+1 if signal_col is not None else 'MISSING'}, "
        f"Route={route_col+1 if route_col is not None else 'MISSING'}, "
        f"Lock_Route={lock_col+1 if lock_col is not None else 'MISSING'}, "
        f"Track={toc_col(headers, 'Track')+1 if toc_col(headers, 'Track') is not None else 'MISSING'}, "
        f"Points={toc_col(headers, 'Points', 'Point')+1 if toc_col(headers, 'Points', 'Point') is not None else 'MISSING'}, "
        f"44_LCP={toc_col(headers, '44_LCP', '44 LCP', 'LC_GATE', 'LC GATE', 'LC')+1 if toc_col(headers, '44_LCP', '44 LCP', 'LC_GATE', 'LC GATE', 'LC') is not None else 'MISSING'}"
    )
    if show_message:
        messagebox.showinfo(
            "TOC Loaded",
            f"{os.path.basename(file)} loaded successfully."
        )

    status_label.config(
        text="TOC LOADED",
        fg="#16a34a"
    )
# =========================================================
# COLORS
# =========================================================

BG = "#0f172a"

BTN = "#2563eb"

GREEN = "#22c55e"

RED = "#ef4444"

# =========================================================
# INTEGRATED CALLING-ON ENGINE
# =========================================================
# Calling-On track/bit-chart operations are embedded here so this file is self-contained.
# The main application's selected TOC, panel search and reporting remain authoritative.

# ==========================================================
# CALLING ON AUTOMATION ENGINE
# Version 2.0
# ==========================================================

import time
import pyautogui

from openpyxl import load_workbook
from pywinauto import Desktop

# ==========================================================
# FILES
# ==========================================================

TOC_FILE = None  # Uses the main application TOC_FILE selected by the user

# ==========================================================
# GLOBAL STATUS
# ==========================================================

LAST_ERROR = ""

BIT_CHART_TIMEOUT = 8

SEARCH_TIMEOUT = 5

SEARCH_RETRY = 0.25

# ==========================================================
# LOG
# ==========================================================

def log(msg):
    print(f"[CALLING-ON] {msg}")


# ==========================================================
# ERROR
# ==========================================================

def set_last_error(msg):
    global LAST_ERROR
    LAST_ERROR = str(msg)
    log(f"ERROR : {msg}")


def get_last_error():
    return LAST_ERROR


# ==========================================================
# RESULT
# ==========================================================

def success(step):

    return {
        "success": True,
        "step": step,
        "reason": ""
    }


def failed(step, reason):

    set_last_error(reason)

    return {
        "success": False,
        "step": step,
        "reason": reason
    }


# ==========================================================
# WAIT
# ==========================================================

def wait(seconds):

    end = time.time() + seconds

    while time.time() < end:
        time.sleep(0.05)

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

        return failed(
            "Click",
            f"{name} not found"
        )

    rect = control.rectangle()

    x = (rect.left + rect.right) // 2
    y = (rect.top + rect.bottom) // 2

    pyautogui.moveTo(
        x,
        y,
        duration=0.25
    )

    pyautogui.click()

    log(f"Clicked : {name}")

    return success("Click")
# ==========================================================
# LOAD CALLING-ON ROUTES
# Reads only the "toc" sheet
# ==========================================================

def load_calling_on_routes():

    routes = []

    try:

        if not TOC_FILE:
            log("TOC file not selected")
            return []

        wb = load_workbook(
            TOC_FILE,
            data_only=True
        )

        if "TOC" not in wb.sheetnames:

            log("toc sheet not found")

            wb.close()

            return []

        sheet = wb["TOC"]

        for row in sheet.iter_rows(
                min_row=2,
                values_only=True
        ):

            if row[0] is None:
                continue

            signal = str(row[0]).strip().upper()

            # Calling-On signals end with C
            if not signal.endswith("C"):
                continue

            headers = get_toc_header_map(sheet)
            route_value = toc_value(row, headers, "Route")
            track_value = toc_value(row, headers, "Track")
            if route_value is None:
                continue
            route = str(route_value).strip()
            track = str(track_value).strip() if track_value is not None else ""

            routes.append({

                "signal": signal,

                "route": route,

                "track": track

            })

        wb.close()

        log(f"Calling-On Routes Loaded : {len(routes)}")

        for item in routes:

            log(
                f"{item['signal']}  "
                f"{item['route']}  "
                f"{item['track']}"
            )

        return routes

    except Exception as e:

        log(f"TOC ERROR : {e}")

        return []


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

    set_last_error("Bit Chart Window Not Detected")

    return False

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

            wait(1)

        except Exception as e:
            log(f"Focus Error : {e}")

        pyautogui.hotkey("ctrl", "b")

        wait(2)
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

        wait(1)

        if not find_and_click("OK"):

            log("OK Button Not Found")

            continue

        wait(2)

        if verify_bit_chart():

            log("Bit Chart Open Success")

            return True

        log("Retry Opening Bit Chart")

    set_last_error("Unable To Open Bit Chart")

    return False


# ==========================================================
# CLOSE BIT CHART
# ==========================================================

def close_bit_chart():
    """toggle_track() already clicks Cancel, so the chart is normally
    closed by the time this runs. Only close it if a Cancel button is
    still on screen - otherwise return at once instead of searching
    every window for ~50 seconds."""

    log("--------------------------------")
    log("Closing Bit Chart")

    if not is_bit_chart_open():
        log("Bit Chart already closed")
        return True

    if not find_and_click("Cancel"):
        log("Bit Chart : Cancel not found - treating as closed")
        return True

    wait(1)

    log("Bit Chart Closed")

    return True


def is_bit_chart_open():
    """True if a Station bit chart window is still open on ANY screen."""
    try:
        for window in Desktop(backend="uia").windows():
            try:
                title = window.window_text().strip().upper()
            except Exception:
                continue

            if title.startswith("STATION") and "INDICATION" in title:
                log(f"BIT CHART WINDOW OPEN : {window.window_text()}")
                return True

    except Exception as e:
        log(f"BIT CHART CHECK ERROR : {e}")

    return False
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

    result = click_control(track)

    if not result["success"]:

        return result

    wait(1)

    # --------------------------------
    # Transmit
    # --------------------------------

    result = click_control("Transmit")

    if not result["success"]:

        return failed(
            "Transmit",
            "Transmit button not found"
        )

    log("Transmit Clicked")

    wait(2)

    # --------------------------------
    # Cancel
    # --------------------------------

    result = click_control("Cancel")

    if not result["success"]:

        return failed(
            "Cancel",
            "Cancel button not found"
        )

    log("Cancel Clicked")

    wait(1)

    return success("Toggle Track")


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

    result = toggle_track(track)

    if not result["success"]:

        log("--------------------------------")
        log("TRACK DOWN FAILED")
        log(f"STEP   : {result['step']}")
        log(f"REASON : {result['reason']}")
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

    result = toggle_track(track)

    if not result["success"]:

        log("--------------------------------")
        log("TRACK UP FAILED")
        log(f"STEP   : {result['step']}")
        log(f"REASON : {result['reason']}")
        log("--------------------------------")

        close_bit_chart()

        return False

    close_bit_chart()

    log("--------------------------------")
    log("TRACK UP COMPLETED")
    log("--------------------------------")

    return True



# =========================================================
# LOG
# =========================================================

LOG_FILE = os.path.join(
    REPORT_FOLDER if REPORT_FOLDER else os.getcwd(),
    f"POINT_LC_CH_SDG_LOG_{time.strftime('%Y%m%d_%H%M%S')}.txt"
)


def log(msg):

    line = f"[{time.strftime('%H:%M:%S')}] {msg}"

    # File + console first, so the complete log exists even when the
    # window is hidden or closed by the launcher.
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass

    try:
        print(line, flush=True)
    except Exception:
        pass

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

#def capture_point():

 #   global capture_waiting
  #  global captured_point

   # capture_waiting = True
    #captured_point = None

    #while capture_waiting:

     #   root.update_idletasks()
      #  pause_sleep(0.05)

    #return list(captured_point)
# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():

    global CONFIG_FILE

    file = filedialog.asksaveasfilename(
        title="Create / Open Configuration",
        defaultextension=".xlsx",
        initialfile="HAH COORDINATES.xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    CONFIG_FILE = file

    if os.path.exists(CONFIG_FILE):

        wb = load_workbook(CONFIG_FILE)
        log("Existing Configuration Opened")

    else:

        wb = Workbook()
        log("New Configuration Created")

    if "MAIN" not in wb.sheetnames:

        if wb.active.title == "Sheet":

            ws = wb.active
            ws.title = "MAIN"

        else:

            ws = wb.create_sheet("MAIN")

        ws.append([
            "Signal",
            "Aspects",
            "Menu_X",
            "Menu_Y",
            "RED_X",
            "RED_Y",
            "YELLOW_X",
            "YELLOW_Y",
            "DOUBLE_YELLOW_X",
            "DOUBLE_YELLOW_Y",
            "GREEN_X",
            "GREEN_Y",
            "RouteInit_X",
            "RouteInit_Y"
        ])
    if "SHUNT" not in wb.sheetnames:
        ws = wb.create_sheet("SHUNT")

        ws.append([
            "Signal",
            "Menu_X",
            "Menu_Y",
            "RED_X",
            "RED_Y",
            "WHITE_X",
            "WHITE_Y"
        ])
    if "CAL" not in wb.sheetnames:
        ws = wb.create_sheet("CAL")

        ws.append([
            "Signal",
            "Menu_X",
            "Menu_Y",
            "WHITE_X",
            "WHITE_Y"
        ])
    if "POINT" not in wb.sheetnames:
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
            "Free_Y"
        ])

    if "CH" not in wb.sheetnames:
        ws = wb.create_sheet("CH")

        ws.append([
            "Point",
            "Menu_X",
            "Menu_Y",
            "IN_X",
            "IN_Y",
            "OUT_X",
            "OUT_Y",
            "FREE_X",
            "FREE_Y"
        ])

    if "LC" not in wb.sheetnames:
        ws = wb.create_sheet("LC")

        ws.append([
            "Gate",
            "Menu_X",
            "Menu_Y",
            "Close_X",
            "Close_Y",
            "Open_X",
            "Open_Y"
        ])
    # ---------------------------------
    # SAVE CONFIGURATION
    # ---------------------------------

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"Configuration Ready : {CONFIG_FILE}")

    messagebox.showinfo(
        "Configuration",
        "Configuration is ready."
    )
# =========================================================
# LOAD CONFIG
# =========================================================

def load_config(file=None, show_message=True):

    global signals
    global points
    global CONFIG_FILE

    if file is None:
        file = filedialog.askopenfilename(
            title="Select Signal Configuration",
            filetypes=[("Excel Files", "*.xlsx")]
        )

    if not file:
        return

    if not os.path.exists(file):
        log(f"Configuration file not found: {file}")
        return

    CONFIG_FILE = file

    signals.clear()
    points.clear()

    if not os.path.exists(CONFIG_FILE):

        messagebox.showerror(
            "Error",
            "Configuration file not found."
        )

        return

    wb = load_workbook(CONFIG_FILE)

    # =====================================================
    # MAIN SIGNALS
    # =====================================================

    if "MAIN" in wb.sheetnames:

        ws = wb["MAIN"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            signal_name = str(row[0]).strip().upper()

            signals[signal_name] = {

                "type": "MAIN",

                "aspects": row[1],

                "menu": [row[2], row[3]],

                "RED": [row[4], row[5]],

                "YELLOW": [row[6], row[7]] if row[6] is not None else None,

                "DOUBLE_YELLOW": [row[8], row[9]] if row[8] is not None else None,

                "GREEN": [row[10], row[11]]

            }

    # =====================================================
    # SHUNT SIGNALS
    # =====================================================

    if "SHUNT" in wb.sheetnames:

        ws = wb["SHUNT"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            signal_name = str(row[0]).strip().upper()

            signals[signal_name] = {

                "type": "SHUNT",

                "menu": [row[1], row[2]],

                "RED": [row[3], row[4]],

                "WHITE": [row[5], row[6]]

            }

    # =====================================================
    # CALLING-ON SIGNALS
    # =====================================================

    if "CAL" in wb.sheetnames:

        ws = wb["CAL"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            signal_name = str(row[0]).strip().upper()

            signals[signal_name] = {

                "type": "CAL",

                "menu": [row[1], row[2]],

                "WHITE": [row[3], row[4]]

            }

    # =====================================================
    # POINTS
    # =====================================================

    if "POINT" in wb.sheetnames:

        ws = wb["POINT"]

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row[0]:
                continue

            point_name = str(row[0]).strip().upper()

            points[point_name] = {

                "menu": [row[1], row[2]],

                "normal": [row[3], row[4]],

                "reverse": [row[5], row[6]],

                "free": [row[7], row[8]]

            }

    wb.close()

    log("CONFIG LOADED")

    if show_message:
        messagebox.showinfo(
            "Configuration Loaded",
            f"{os.path.basename(CONFIG_FILE)} loaded successfully."
        )

    status_label.config(
        text="CONFIG LOADED",
        fg="#16a34a"
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
            row[5].value = n_x
            row[6].value = n_y

            row[7].value = r_x
            row[8].value = r_y



            updated = True
            break

    if not updated:
        sheet.append([
            point,
            None,
            None,
            None,
            None,
            n_x,
            n_y,
            r_x,
            r_y,

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

    for index, point in enumerate(loaded_point_list, start=1):

        coords = [None] * 4

        steps = [
            "RECORD CONTROL",
            "RECORD NORMAL",
            "RECORD REVERSE",
            "RECORD FREE"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(

                f"POINT {index} OF {len(loaded_point_list)}\n\nPOINT ID : {point}",

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

        log(f"Final Coordinates : {coords}")

        save_point_control_coordinate(
            point,
            coords
        )

        log(f"{point} CONTROL : {coords[0]}")
        log(f"{point} NORMAL  : {coords[1]}")
        log(f"{point} REVERSE : {coords[2]}")
        log(f"{point} FREE : {coords[3]}")

    log("Point Control Coordinates Saved")


def save_point_control_coordinate(point, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["POINT"]

    updated = False

    for row in sheet.iter_rows(min_row=2):

        if str(row[0].value).strip() == str(point).strip():

            row[1].value = coords[0][0]   # CONTROL X
            row[2].value = coords[0][1]   # CONTROL Y

            row[3].value = coords[1][0]   # NORMAL X
            row[4].value = coords[1][1]   # NORMAL Y

            row[5].value = coords[2][0]   # REVERSE X
            row[6].value = coords[2][1]   # REVERSE Y

            row[7].value = coords[3][0]   # FREE X
            row[8].value = coords[3][1]   # FREE Y

            updated = True
            break

    if not updated:

        sheet.append([
            point,

            coords[0][0],
            coords[0][1],

            coords[1][0],
            coords[1][1],

            coords[2][0],
            coords[2][1],

            coords[3][0],
            coords[3][1]
        ])

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{point} saved successfully.")


def get_point_indication_coordinates(point):

    log("Opening CONFIG_FILE...")

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "POINT" not in wb.sheetnames:

        log("POINT sheet not found")

        wb.close()

        return None

    sheet = wb["POINT"]

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip() == str(point).strip():

            coordinates = (
                int(row[5]),   # Normal_X
                int(row[6]),   # Normal_Y
                int(row[7]),   # Reverse_X
                int(row[8])    # Reverse_Y
            )

            wb.close()

            log(f"Coordinates Found : {coordinates}")

            return coordinates

    wb.close()

    log(f"Point {point} indication coordinates not found")

    return None


def get_lc_gate_coordinates(lc_name):
    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "LC" not in wb.sheetnames:
        log("LC_GATE sheet not found")

        wb.close()

        return None

    sheet = wb["LC"]

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() == str(lc_name).strip().upper():
            coordinates = {

                "control": (
                    int(row[1]),
                    int(row[2])
                ),

                "close": (
                    int(row[3]),
                    int(row[4])
                ),

                "open": (
                    int(row[5]),
                    int(row[6])
                )
            }

            wb.close()

            log(
                f"LC Gate {lc_name} Coordinates : "
                f"{coordinates}"
            )

            return coordinates

    # The TOC name (e.g. "44") may differ from the name recorded in the
    # LC sheet (e.g. "LC141"). If the sheet holds exactly ONE gate, use it.
    single = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            single.append(row)

    if len(single) == 1:

        row = single[0]

        coordinates = {
            "control": (int(row[1]), int(row[2])),
            "close": (int(row[3]), int(row[4])),
            "open": (int(row[5]), int(row[6]))
        }

        wb.close()

        log(
            f"LC Gate {lc_name} not listed - using the only gate in the "
            f"LC sheet : {row[0]} : {coordinates}"
        )

        return coordinates

    wb.close()

    log(f"LC Gate {lc_name} not found")

    return None

def detect_lc_gate_state(lc_name):
    """Read the LC gate indication lamps from the configured coordinates.

    LC sheet layout used by this program:
        A Gate
        B Menu_X
        C Menu_Y
        D Close_X
        E Close_Y
        F Open_X
        G Open_Y

    Panel indication:
        CLOSE = GREEN
        OPEN  = YELLOW
    """

    log("--------------------------------")
    log(f"CHECKING LC GATE : {lc_name}")

    coordinates = get_lc_gate_coordinates(lc_name)

    if coordinates is None:
        log(f"LC Gate coordinates not found : {lc_name}")
        return "UNKNOWN"

    close_x, close_y = coordinates["close"]
    open_x, open_y = coordinates["open"]

    log(f"CLOSE Coordinate : ({close_x}, {close_y})")
    log(f"OPEN Coordinate  : ({open_x}, {open_y})")

    try:
        screen = pyautogui.screenshot()
        close_rgb = screen.getpixel((close_x, close_y))
        open_rgb = screen.getpixel((open_x, open_y))
    except Exception as e:
        log(f"LC screenshot/read failed : {e}")
        return "UNKNOWN"

    log(f"CLOSE RGB : {close_rgb}")
    log(f"OPEN RGB  : {open_rgb}")

    # CLOSE indication = GREEN
    close_on = (
        close_rgb[1] >= 150 and
        close_rgb[1] > close_rgb[0] * 1.3 and
        close_rgb[1] > close_rgb[2] * 1.3
    )

    # OPEN indication = YELLOW
    open_on = (
        open_rgb[0] >= 150 and
        open_rgb[1] >= 150 and
        open_rgb[2] <= 130
    )

    log(f"CLOSE Lamp ON : {close_on}")
    log(f"OPEN Lamp ON  : {open_on}")

    if close_on and not open_on:
        log(f"LC Gate {lc_name} State : CLOSE")
        return "CLOSE"

    if open_on and not close_on:
        log(f"LC Gate {lc_name} State : OPEN")
        return "OPEN"

    if close_on and open_on:
        log(f"LC Gate {lc_name} State : INVALID (both lamps ON)")
        return "UNKNOWN"

    log(f"LC Gate {lc_name} State : UNKNOWN (both lamps OFF)")
    return "UNKNOWN"

def try_lc_gate_operation(signal, route, lc_name):

    log("--------------------------------")
    log(f"Trying LC Gate : {lc_name}")

    coordinates = get_lc_gate_coordinates(lc_name)

    if coordinates is None:

        log("LC Gate coordinates not found")

        return False

    control_x, control_y = coordinates["control"]

    # =====================================
    # CLICK LC GATE
    # =====================================

    safe_click(control_x, control_y)

    log(f"Clicked LC Gate : {lc_name}")

    time.sleep(2)

    # =====================================
    # CLICK TRANSMIT
    # =====================================

    log("Searching for Transmit...")

    if not find_and_click("Transmit"):

        log("Transmit button not found")

        save_result(
            signal,
            route,
            f"Locked LC Gate {lc_name}",
            "CLOSE",
            "Transmit Not Found",
            "PASS",
            "Transmit unavailable"
        )

        return True

    log("Transmit clicked")

    time.sleep(5)

    # =====================================
    # VERIFY LC GATE STATE
    # =====================================

    state = detect_lc_gate_state(lc_name)

    log(f"LC Gate State : {state}")

    if state == "CLOSE":

        log("LC Gate remained locked")

        save_result(
            signal,
            route,
            f"Locked LC Gate {lc_name}",
            "CLOSE",
            state,
            "PASS",
            "LC Gate remained locked"
        )

        return True

    log("LC Gate accepted Transmit")

    save_result(
        signal,
        route,
        f"Locked LC Gate {lc_name}",
        "CLOSE",
        state,
        "FAIL",
        "Transmit accepted"
    )

    return False


def verify_locked_lc_gates(signal, route):
    """Verify every LC gate associated with the current signal/route."""

    log("================================")
    log("LOCKED LC GATE VERIFICATION")
    log("================================")

    gates = load_lc_gates(signal, route)

    if not gates:
        log(f"No LC gates configured for {signal} / {route}")
        return True

    for lc_name in gates:
        if not try_lc_gate_operation(signal, route, lc_name):
            log(f"Locked LC Gate Test Failed : {lc_name}")
            return False

    log("================================")
    log("ALL LOCKED LC GATES VERIFIED")
    log("================================")

    return True


# =========================================================
# POINT LAMP DETECTION
# =========================================================

def is_normal_lamp_on(rgb):
    r, g, b = rgb
    return (
        g >= 120 and
        g > r + 40 and
        g > b + 40
    )


def is_reverse_lamp_on(rgb):
    r, g, b = rgb
    return (
        r >= 140 and
        g >= 140 and
        abs(r - g) <= 80 and
        b <= 130
    )


def detect_point_state(point):

    log("--------------------------------")
    log(f"CHECKING POINT {point} INDICATION")

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "POINT" not in wb.sheetnames:
        wb.close()
        log("POINT sheet not found")
        return None

    sheet = wb["POINT"]

    normal = None
    reverse = None
    free = None

    for row in sheet.iter_rows(min_row=2, values_only=True):

        if not row or row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(point).strip().upper():
            continue

        try:
            normal = (int(row[3]), int(row[4]))
            reverse = (int(row[5]), int(row[6]))
            free = (int(row[7]), int(row[8]))
        except (TypeError, ValueError):
            wb.close()
            log(f"Invalid point coordinates for Point {point}")
            return None

        break

    wb.close()

    if normal is None or reverse is None:
        log(f"Point {point} indication coordinates not found")
        return None

    try:
        screen = pyautogui.screenshot()
    except Exception as e:
        log(f"Point screenshot failed : {e}")
        return None

    log(f"Normal Coordinate : {normal}")
    log(f"Reverse Coordinate: {reverse}")
    log(f"Free Coordinate   : {free}")

    def search_lamp(x, y):

        green_found = False
        yellow_found = False

        best_green = (0, 0, 0)
        best_yellow = (0, 0, 0)

        # Search an 11x11 area around the recorded coordinate.
        # The exact recorded pixel can be black even when the lamp is ON.
        for dx in range(-5, 6):
            for dy in range(-5, 6):

                px = x + dx
                py = y + dy

                if px < 0 or py < 0:
                    continue

                try:
                    rgb = screen.getpixel((px, py))
                except Exception:
                    continue

                r, g, b = rgb

                # NORMAL indication = GREEN
                if (
                    g >= 120 and
                    g > r + 40 and
                    g > b + 40
                ):
                    green_found = True

                    if g > best_green[1]:
                        best_green = rgb

                # REVERSE indication = YELLOW
                if (
                    r >= 140 and
                    g >= 140 and
                    abs(r - g) <= 80 and
                    b <= 130
                ):
                    yellow_found = True

                    if (r + g) > (best_yellow[0] + best_yellow[1]):
                        best_yellow = rgb

        return green_found, yellow_found, best_green, best_yellow

    n_green, n_yellow, n_green_rgb, n_yellow_rgb = search_lamp(
        normal[0], normal[1]
    )

    r_green, r_yellow, r_green_rgb, r_yellow_rgb = search_lamp(
        reverse[0], reverse[1]
    )

    log(f"Normal Area Green : {n_green} {n_green_rgb}")
    log(f"Normal Area Yellow: {n_yellow} {n_yellow_rgb}")
    log(f"Reverse Area Green: {r_green} {r_green_rgb}")
    log(f"Reverse Area Yellow: {r_yellow} {r_yellow_rgb}")

    if n_green and not r_yellow:
        log(f"Point {point} Current State : N")
        return "N"

    if r_yellow and not n_green:
        log(f"Point {point} Current State : R")
        return "R"

    log(f"Point {point} Current State : UNKNOWN")
    return "UNKNOWN"


def get_expected_point_states(route):
    log("Reading toc sheet")

    try:
        wb = load_workbook(
            TOC_FILE,
            data_only=True,
            read_only=True
        )

        if "TOC" not in wb.sheetnames:
            log("toc sheet not found")
            wb.close()
            return None

        sheet = wb["TOC"]
        headers = get_toc_header_map(sheet)

        for row in sheet.iter_rows(min_row=2, values_only=True):
            route_value = toc_value(row, headers, "Route")
            points_value = toc_value(row, headers, "Points", "Point")
            if route_value is None:
                continue
            toc_route = str(route_value).strip()
            if toc_route == str(route).strip():
                expected = str(points_value).strip() if points_value not in (None, "") else ""

                wb.close()

                return expected

        wb.close()

        log(f"No point states found for route : {route}")

        return None

    except Exception as e:

        log(f"POINT VERIFY TOC ERROR : {e}")

        return None

def load_expected_points(route):

    expected = get_expected_point_states(route)

    if not expected:
        return []

    points = []

    for item in expected.split(","):

        item = item.strip()

        if len(item) < 2:
            continue

        point = item[:-1]
        state = item[-1].upper()

        points.append((point, state))

    return points
# =========================================================
# LOCKED POINT VERIFICATION
# =========================================================

def verify_locked_points(signal, route):

    log("================================")
    log("LOCKED POINT VERIFICATION")
    log("================================")

    expected_points = load_expected_points(route)

    if not expected_points:

        log("No locked points found")

        return True

    for point, expected_state in expected_points:

        log("--------------------------------")
        log(f"Locked Point : {point}")
        log(f"Expected State : {expected_state}")

        coordinates = get_point_control_coordinates(point)

        if coordinates is None:

            log(f"Point control coordinates not found : {point}")

            save_result(
                signal,
                route,
                f"Locked Point {point}",
                expected_state,
                "Coordinate Missing",
                "FAIL",
                "Point control coordinates missing"
            )

            return False

        control_x, control_y = coordinates["control"]

        log(f"Opening Point Control : {point}")

        safe_click(control_x, control_y)

        time.sleep(1)

        current_state = detect_point_state(point)

        log(f"Current State : {current_state}")

        if current_state != expected_state:

            log("Point has not reached expected state yet")

            time.sleep(3)

            current_state = detect_point_state(point)

            log(f"State After Wait : {current_state}")

            if current_state != expected_state:

                log("Point never reached expected state")

                return False

        operation = "Reverse" if expected_state == "N" else "Normal"

        log(f"Trying Locked Point Command : {operation}")

        if not try_point_operation(
            signal,
            route,
            point,
            expected_state,
            operation
        ):

            pyautogui.click(100, 100)

            return False

        pyautogui.click(100, 100)

        time.sleep(1)

        log(f"Locked Point Passed : {point}")

    log("================================")
    log("ALL LOCKED POINTS VERIFIED")
    log("================================")

    return True

def try_point_operation(signal, route, point, expected_state, operation):

    log("--------------------------------")
    log(f"Trying {operation} on Point {point}")

    operation_sent = False

    if find_and_click(operation):

        operation_sent = True

        log(f"{operation} Clicked")

        time.sleep(3)

    else:

        log(f"{operation} menu not found")

    actual_state = detect_point_state(point)

    if actual_state is None:

        save_result(
            signal,
            route,
            f"Locked Point {point}",
            expected_state,
            "Unknown",
            "FAIL",
            "Unable to detect point indication"
        )

        return False

    log(f"Expected State : {expected_state}")
    log(f"Actual State   : {actual_state}")

    # Locked point behaved correctly
    if operation_sent and actual_state == expected_state:

        save_result(
            signal,
            route,
            f"Locked Point {point}",
            expected_state,
            actual_state,
            "PASS",
            "Point remained locked"
        )

        return True

    log("--------------------------------")
    log("POINT LOCK TEST FAILED")
    log("STARTING RECOVERY")
    log("--------------------------------")

    if recover_point(
        signal,
        route,
        point,
        expected_state
    ):

        log("Point Recovery Successful")

        return True

    log("Point Recovery Failed")

    save_result(
        signal,
        route,
        f"Locked Point {point}",
        expected_state,
        actual_state,
        "FAIL",
        "Recovery Failed"
    )

    return False

# ==========================================================
# LOCKED CONDITION VERIFICATION
# ==========================================================
def verify_locked_conditions(signal, route):

    log("===============================")
    log("VERIFYING LOCKED CONDITIONS")
    log("===============================")

    if not verify_locked_points(signal, route):
        return False

    return True


def get_point_control_coordinates(point):

    log("Opening CONFIG_FILE...")

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "POINT" not in wb.sheetnames:
        wb.close()
        log("POINT sheet not found")
        return None

    sheet = wb["POINT"]

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(point).strip().upper():
            continue

        coordinates = {

            "control": (
                int(row[1]),
                int(row[2])
            ),

            "normal": (
                int(row[3]),
                int(row[4])
            ),

            "reverse": (
                int(row[5]),
                int(row[6])
            ),

            "free": (
                int(row[7]),
                int(row[8])
            )

        }

        wb.close()

        log(f"Point {point} Coordinates : {coordinates}")

        return coordinates

    wb.close()

    log(f"Point {point} control coordinates not found")

    return None

# =========================================================
# VERIFY POINTS
# =========================================================

def verify_points(signal, route):

    log("========== ENTER verify_points ==========")
    log(f"Route = {route}")

    log("================================")
    log("POINT LOCK VERIFICATION")
    log("================================")

    expected_points = load_expected_points(route)

    log(f"Points To Verify = {expected_points}")

    if not expected_points:

        log("No points found for this route.")

        return True, None, None

    first_failed_point = None
    first_failed_state = None

    # =====================================================
    # VERIFY EVERY POINT
    # =====================================================

    for point, route_state in expected_points:

        log("--------------------------------")
        log(f"VERIFYING POINT : {point}")
        log("--------------------------------")

        # ---------------------------------
        # READ CURRENT POINT STATE
        # ---------------------------------

        current_state = detect_point_state(point)

        log(
            f"Point {point} Current State : "
            f"{current_state}"
        )

        if current_state not in ("N", "R"):

            log(
                f"Point {point} has invalid/unknown "
                f"state : {current_state}"
            )

            save_result(
                signal,
                route,
                f"Point {point} Lock Test",
                "N/R",
                str(current_state),
                "FAIL",
                "Unable to determine current point state"
            )

            if first_failed_point is None:

                first_failed_point = point
                first_failed_state = current_state

            continue

        # =================================================
        # DECIDE OPPOSITE COMMAND
        # =================================================

        if current_state == "N":

            target_command = "Reverse"

        else:

            target_command = "Normal"

        log(
            f"Point {point} is {current_state}"
        )

        log(
            f"Trying opposite command : "
            f"{target_command}"
        )

        # =================================================
        # OPEN POINT CONTROL
        # =================================================

        coordinates = get_point_control_coordinates(point)

        if coordinates is None:

            log(
                f"Point control coordinates not found : "
                f"{point}"
            )

            save_result(
                signal,
                route,
                f"Point {point} Lock Test",
                current_state,
                "Coordinate Missing",
                "FAIL",
                "Point control coordinates not found"
            )

            if first_failed_point is None:

                first_failed_point = point
                first_failed_state = current_state

            continue

        control_x, control_y = coordinates["control"]

        log(
            f"Opening Point Control : {point}"
        )

        safe_click(
            control_x,
            control_y
        )

        time.sleep(1)

        # =================================================
        # TRY OPPOSITE POINT OPERATION
        # =================================================

        command_sent = find_and_click(
            target_command
        )

        if not command_sent:

            log(
                f"{target_command} command NOT found "
                f"for Point {point}"
            )

            # Close possible menu
            pyautogui.click(100, 100)

            save_result(
                signal,
                route,
                f"Point {point} Lock Test",
                current_state,
                "Command Not Found",
                "FAIL",
                f"{target_command} command not found"
            )

            if first_failed_point is None:

                first_failed_point = point
                first_failed_state = current_state

            continue

        log(
            f"{target_command} command clicked "
            f"for Point {point}"
        )

        # =================================================
        # WAIT FOR POINT RESPONSE
        # =================================================

        time.sleep(3)

        # =================================================
        # CHECK POINT AGAIN
        # =================================================

        new_state = detect_point_state(point)

        log(
            f"Point {point} State After "
            f"{target_command} : {new_state}"
        )

        # Close point menu
        pyautogui.click(100, 100)

        # =================================================
        # LOCK TEST RESULT
        # =================================================
        #
        # IMPORTANT:
        #
        # If the route is set, the point should be LOCKED.
        #
        # Therefore:
        #
        # N + Reverse attempt -> must remain N
        #
        # R + Normal attempt -> must remain R
        #
        # If the state changes, the point is NOT locked.
        # =================================================

        if new_state == current_state:

            log(
                f"Point {point} remained "
                f"{current_state}"
            )

            log(
                f"POINT {point} LOCK TEST : PASS"
            )

            save_result(
                signal,
                route,
                f"Point {point} Lock Test",
                current_state,
                new_state,
                "PASS",
                f"Point remained {current_state}; "
                f"{target_command} operation was blocked"
            )

        else:

            log(
                f"Point {point} CHANGED from "
                f"{current_state} to {new_state}"
            )

            log(
                f"POINT {point} LOCK TEST : FAIL"
            )

            save_result(
                signal,
                route,
                f"Point {point} Lock Test",
                current_state,
                new_state,
                "FAIL",
                f"Point changed after {target_command}"
            )

            if first_failed_point is None:

                first_failed_point = point
                first_failed_state = current_state

    # =====================================================
    # FINAL RESULT
    # =====================================================

    log("================================")
    log("POINT LOCK VERIFICATION COMPLETED")
    log("================================")

    if first_failed_point is not None:

        log(
            f"Point Lock Test FAILED : "
            f"{first_failed_point}"
        )

        return (
            False,
            first_failed_point,
            first_failed_state
        )

    log("BOTH POINTS LOCK TEST PASSED")

    return True, None, None

def verify_points_after_route(signal, route):
    expected_points = load_expected_points(route)

    for point, expected_state in expected_points:

        actual_state = detect_point_state(point)

        if actual_state == expected_state:

            save_result(
                signal,
                route,
                f"Point {point} Movement",
                expected_state,
                actual_state,
                "PASS",
                "Point moved correctly"
            )

        else:

            save_result(
                signal,
                route,
                f"Point {point} Movement",
                expected_state,
                actual_state,
                "FAIL",
                "Point did not move"
            )

            return False

    return True


def recover_point(signal, route, point, expected_state):

    log("================================")
    log("POINT RECOVERY STARTED")
    log(f"Point          : {point}")
    log(f"Expected State : {expected_state}")
    log("================================")

    current_state = detect_point_state(point)

    if current_state is None:

        log("Unable to detect current point state")

        return False

    log(f"Current State : {current_state}")

    coordinates = get_point_control_coordinates(point)

    if coordinates is None:

        log(f"Point control coordinates not found : {point}")

        return False

    control_x, control_y = coordinates["control"]

    safe_click(control_x, control_y)

    log(f"Opened Point Control : {point}")

    time.sleep(1)

    # ---------------------------------------
    # Decide which command to press
    # ---------------------------------------

    if expected_state == "N":
        target_command = "Normal"
    else:
        target_command = "Reverse"

    log(f"Looking for {target_command} command...")

    if find_and_click(target_command):

        log(f"{target_command} Clicked")

        time.sleep(5)

        new_state = detect_point_state(point)

        log(f"Point State After Operation : {new_state}")

        if new_state == expected_state:

            log("================================")
            log("DIRECT POINT RECOVERY SUCCESSFUL")
            log("================================")

            save_result(
                signal,
                route,
                f"Point {point} Recovery",
                expected_state,
                new_state,
                "PASS",
                "Direct Point Recovery Successful"
            )

            return True

        log("Direct Point Recovery Failed")

    else:

        log(f"{target_command} command not found")

    # ---------------------------------------
    # Crank Handle Recovery
    # ---------------------------------------

    log("================================")
    log("DIRECT POINT OPERATION FAILED")
    log("STARTING CRANK HANDLE RECOVERY")
    log("================================")

    recovered = recover_using_crank_handle(
        signal,
        route,
        point,
        expected_state
    )

    if not recovered:

        log("================================")
        log("CRANK HANDLE RECOVERY FAILED")
        log("================================")

        save_result(
            signal,
            route,
            f"Crank Handle Recovery - Point {point}",
            expected_state,
            "Recovery Failed",
            "FAIL",
            "Crank Handle Recovery Failed"
        )

        return False

    log("================================")
    log("CRANK HANDLE RECOVERY SUCCESSFUL")
    log("================================")

    save_result(
        signal,
        route,
        f"Crank Handle Recovery - Point {point}",
        expected_state,
        expected_state,
        "PASS",
        "Crank Handle Recovery Successful"
    )

    log("Waiting 5 seconds before retrying point...")

    time.sleep(5)

    # ---------------------------------------
    # Retry Point
    # ---------------------------------------

    safe_click(control_x, control_y)

    log(f"Opened Point Control : {point}")

    time.sleep(1)

    if not find_and_click(target_command):

        log(f"{target_command} command not found after Crank Handle")

        save_result(
            signal,
            route,
            f"Point {point} Recovery",
            expected_state,
            "Command Not Found",
            "FAIL",
            "Point command unavailable after Crank Handle"
        )

        return False

    log(f"{target_command} Clicked")

    time.sleep(5)

    new_state = detect_point_state(point)

    log(f"Point State After Recovery : {new_state}")

    if new_state == expected_state:

        log("================================")
        log("POINT RECOVERY SUCCESSFUL")
        log(f"Point {point} = {expected_state}")
        log("================================")

        save_result(
            signal,
            route,
            f"Point {point} Recovery",
            expected_state,
            new_state,
            "PASS",
            "Point recovered after Crank Handle"
        )

        return True

    log("================================")
    log("POINT RECOVERY FAILED")
    log("================================")

    save_result(
        signal,
        route,
        f"Point {point} Recovery",
        expected_state,
        new_state,
        "FAIL",
        "Point still not in expected state after Crank Handle"
    )

    return False


# =========================================================
# GET POINT CONTROL
# =========================================================

def get_point_control(point):

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "POINTS" not in wb.sheetnames:

        wb.close()

        return None

    sheet = wb["POINTS"]

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(point).strip().upper():
            continue

        wb.close()

        return (
            int(row[1]),
            int(row[2])
        )

    wb.close()

    return None
def recover_using_crank_handle(signal, route, point, expected_state):

    ch_id = resolve_crank_handle_config_id(point)
    log(f"CH CONFIG ID : {ch_id}")

    log("================================")
    log("CRANK HANDLE RECOVERY")
    log(f"Point : {point}")
    log(f"Expected State : {expected_state}")
    log("================================")

    # --------------------------------
    # GET CRANK HANDLE COORDINATE
    # --------------------------------

    log("Reading Crank Handle coordinate...")

    try:
        wb = load_workbook(
            CONFIG_FILE,
            data_only=True,
            read_only=True
        )

        log("CONFIG_FILE opened for Crank Handle")

        if "CH" not in wb.sheetnames:

            log("CRANK_HANDLE sheet not found")

            wb.close()

            return False

        sheet = wb["CH"]

        log("CRANK_HANDLE sheet found")

        crank_x = None
        crank_y = None

        for row in sheet.iter_rows(
                min_row=2,
                values_only=True
        ):

            log(f"Checking Crank Handle Row : {row}")

            if row[0] is None:
                continue

            if str(row[0]).strip().upper() == str(ch_id).strip().upper():

                crank_x = int(row[1])
                crank_y = int(row[2])

                break

        wb.close()

    except Exception as e:

        log(f"CRANK HANDLE READ ERROR : {e}")

        return False

    if crank_x is None or crank_y is None:

        log(
            f"Crank Handle coordinate "
            f"not found for Point {point}"
        )

        return False

    log(
        f"Crank Handle Coordinate : "
        f"({crank_x}, {crank_y})"
    )

    # --------------------------------
    # CLICK RED POINT IN CRANK HANDLE BOX
    # --------------------------------

    log(
        f"Moving mouse to Crank Handle "
        f"Point {point}"
    )

    safe_click(crank_x, crank_y)

    log("Mouse moved to Crank Handle point")



    log(
        f"Clicked Crank Handle Point : "
        f"{point}"
    )

    # Wait for crank handle popup
    log("Waiting 2 seconds for Crank Handle popup...")

    time.sleep(2)

    # --------------------------------
    # CLICK TRANSMIT
    # --------------------------------

    log("Searching for Transmit...")

    if not find_and_click("Transmit"):

        log("Transmit not found")

        return False

    log("Transmit clicked successfully")
    save_result(
        signal,
        route,
        f"Crank Handle Point {point}",
        "Transmit",
        "Transmit",
        "PASS",
        ""
    )

    # You said after Transmit the popup stays
    # and approximately 5 seconds should be allowed.

    log("Waiting 5 seconds after Transmit...")

    time.sleep(5)

    log("CRANK HANDLE OPERATION COMPLETED")

    return True

def resolve_crank_handle_config_id(point):
    """Resolve the CH-sheet identifier without changing POINT identifiers.

    Route/TOC point identifiers remain 50, 65, etc.
    Only the CH configuration lookup uses the configured CH names:
        50 -> CH1
        65 -> CH2

    Direct CH identifiers (CH1/CH2) are also accepted.
    """
    if point is None:
        return None

    value = str(point).strip().upper()

    # If the caller already supplied a CH identifier, use it directly.
    if value.startswith("CH"):
        return value

    # Current HAH configuration naming for the two crank handles.
    aliases = {
        "50": "CH1",
        "65": "CH2",
    }

    resolved = aliases.get(value, value)

    if resolved != value:
        log(f"CH CONFIG RESOLUTION : {value} -> {resolved}")

    return resolved


def try_crank_handle_operation(signal, route, point, expected_state):

    ch_id = resolve_crank_handle_config_id(point)

    log("--------------------------------")
    log(f"Trying Crank Handle : {point}")
    log(f"CH CONFIG ID       : {ch_id}")
    log(f"Expected Point State : {expected_state}")

    # ========================================
    # GET CRANK HANDLE COORDINATES
    # ========================================

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "CH" not in wb.sheetnames:

        wb.close()

        log("CH sheet not found")

        return False

    sheet = wb["CH"]

    crank_x = None
    crank_y = None

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(ch_id).strip().upper():
            continue

        try:

            crank_x = int(row[1])
            crank_y = int(row[2])

        except (TypeError, ValueError):

            wb.close()

            log(f"Invalid Crank Handle coordinates : {point}")

            return False

        break

    wb.close()

    if crank_x is None or crank_y is None:

        log(
            f"Crank Handle coordinate not found : {point}"
        )

        return False

    log(
        f"Crank Handle Coordinate : "
        f"({crank_x}, {crank_y})"
    )

    # ========================================
    # READ INITIAL CRANK HANDLE STATE
    # ========================================

    initial_state = get_crank_handle_state(ch_id)

    log(f"Initial Crank Handle State : {initial_state}")

    if initial_state not in ("IN", "OUT", "FREE"):

        log(
            f"Invalid Initial Crank Handle State : "
            f"{initial_state}"
        )

        save_result(
            signal,
            route,
            f"Locked Crank Handle {point}",
            "IN/OUT/FREE",
            str(initial_state),
            "FAIL",
            "Unable to determine initial crank handle state"
        )

        return False

    # ========================================
    # OPEN CRANK HANDLE CONTROL
    # ========================================

    log(
        f"Opening Crank Handle Control : "
        f"{point}"
    )

    safe_click(crank_x, crank_y)

    time.sleep(2)

    # ========================================
    # SEARCH TRANSMIT
    # ========================================

    log("Searching for Transmit...")

    if not find_and_click("Transmit"):

        log("Transmit not found")

        save_result(
            signal,
            route,
            f"Locked Crank Handle {point}",
            "Transmit",
            "Not Found",
            "FAIL",
            "Transmit button not found"
        )

        return False

    log("Transmit clicked successfully")

    time.sleep(5)

    # ========================================
    # VERIFY CRANK HANDLE AFTER TRANSMIT
    # ========================================

    current_state = get_crank_handle_state(ch_id)

    log(
        f"Crank Handle State After Transmit : "
        f"{current_state}"
    )

    # ========================================
    # UNKNOWN STATE = FAIL
    # ========================================

    if current_state not in ("IN", "OUT", "FREE"):

        log(
            f"Unable to determine Crank Handle state "
            f"after Transmit : {current_state}"
        )

        save_result(
            signal,
            route,
            f"Locked Crank Handle {point}",
            initial_state,
            str(current_state),
            "FAIL",
            "Unable to determine state after Transmit"
        )

        return False

    # ========================================
    # LOCKED HANDLE SHOULD NOT CHANGE
    # ========================================

    if current_state == initial_state:

        log("Crank Handle remained in same state")

        save_result(
            signal,
            route,
            f"Locked Crank Handle {point}",
            initial_state,
            current_state,
            "PASS",
            "Transmit Rejected - Crank Handle remained locked"
        )

        log("Crank Handle remained locked")

        return True

    # ========================================
    # HANDLE STATE CHANGED
    # ========================================

    log("--------------------------------")
    log("CRANK HANDLE STATE CHANGED")
    log("--------------------------------")

    log(f"Initial State : {initial_state}")
    log(f"Current State : {current_state}")

    save_result(
        signal,
        route,
        f"Locked Crank Handle {point}",
        initial_state,
        current_state,
        "FAIL",
        "Transmit accepted - Crank Handle state changed"
    )

    # ========================================
    # RECOVERY
    # ========================================

    log("--------------------------------")
    log("RECOVERING CRANK HANDLE")
    log("--------------------------------")

    if not recover_crank_handle(
            signal,
            route,
            ch_id
    ):

        log(
            f"Crank Handle Recovery Failed : "
            f"{point}"
        )

        return False

    time.sleep(2)

    recovered_state = get_crank_handle_state(ch_id)

    log(
        f"Crank Handle State After Recovery : "
        f"{recovered_state}"
    )

    # ========================================
    # VERIFY RECOVERY
    # ========================================

    if recovered_state in ("IN", "FREE"):

        log("--------------------------------")
        log("CRANK HANDLE RECOVERY SUCCESSFUL")
        log("--------------------------------")

        save_result(
            signal,
            route,
            f"Crank Handle {point} Recovery",
            "IN/FREE",
            recovered_state,
            "PASS",
            "Crank Handle recovered successfully"
        )

        return True

    log("--------------------------------")
    log("CRANK HANDLE RECOVERY FAILED")
    log("--------------------------------")

    save_result(
        signal,
        route,
        f"Crank Handle {point} Recovery",
        "IN/FREE",
        str(recovered_state),
        "FAIL",
        "Crank Handle not recovered"
    )

    return False

def verify_locked_crank_handle(signal, route):

    log("================================")
    log("LOCKED CRANK HANDLE VERIFICATION")
    log("================================")

    expected_points = load_expected_points(route)

    if not expected_points:

        log("No points found")

        return True

    verification_failed = False

    # ========================================
    # VERIFY ALL CRANK HANDLES
    # ========================================

    for point, expected_state in expected_points:

        log("--------------------------------")
        log(f"VERIFY CRANK HANDLE : {point}")
        log("--------------------------------")

        success = try_crank_handle_operation(
            signal,
            route,
            point,
            expected_state
        )

        if not success:

            log(
                f"Locked Crank Handle Failed : "
                f"{point}"
            )

            verification_failed = True

            # IMPORTANT:
            # Do NOT return here.
            # Continue checking the next crank handle.

            continue

        log(
            f"Locked Crank Handle Passed : "
            f"{point}"
        )

    # ========================================
    # FINAL RESULT
    # ========================================

    if verification_failed:

        log("================================")
        log("ONE OR MORE CRANK HANDLES FAILED")
        log("================================")

        return False

    log("================================")
    log("ALL CRANK HANDLES VERIFIED")
    log("================================")

    return True

# =========================================================
# LOAD ALL ROUTES OF A SIGNAL
# =========================================================

def load_routes_for_signal(signal_name):

    routes = []

    try:

        wb = load_workbook(TOC_FILE)

        sheet = wb["TOC"]

        for row in sheet.iter_rows(min_row=2, values_only=True):

            if row[0] is None:
                continue

            signal = str(row[0]).strip()
            route = str(row[1]).strip()

            if signal == str(signal_name):
                routes.append(route)

        log("--------------------------------")
        log(f"Signal : {signal_name}")
        log(f"Routes Found : {len(routes)}")

        for r in routes:
            log(f"   {r}")

        log("--------------------------------")

        return routes

    except Exception as e:

        log(f"TOC ERROR : {e}")

        return []

def load_first_test():
    try:

        wb = load_workbook(TOC_FILE)

        sheet = wb["TOC"]

        for row in sheet.iter_rows(min_row=2, values_only=True):

            if row[0] is None:
                continue

            signal = str(row[0]).strip()

            route = str(row[1]).strip()

            log("--------------------------------")

            log("FIRST TEST LOADED")

            log(f"Signal : {signal}")

            log(f"Route  : {route}")

            log("--------------------------------")

            return signal, route

        log("No test cases found.")

        return None, None

    except Exception as e:

        log(f"TOC ERROR : {e}")

        return None, None


from openpyxl import load_workbook



def release_route(signal, route):

    signal_x, signal_y = get_signal(signal)

    if signal_x is None:

        log(f"Signal {signal} coordinates not found.")

        save_result(
            signal,
            route,
            "Route Release",
            "Signal Coordinates",
            "Not Found",
            "FAIL",
            "Signal coordinates not found"
        )

        return False

    # ========================================
    # SIGNAL CANCEL
    # ========================================

    log("--------------------------------")
    log(f"Releasing Route : {route}")
    log(f"Signal Cancel : {signal}")
    log("--------------------------------")

    safe_click(
        signal_x,
        signal_y
    )

    time.sleep(1)

    if not find_and_click("Signal Cancel"):

        log(f"{signal} Signal Cancel Failed")

        save_result(
            signal,
            route,
            "Signal Cancel",
            "Successful",
            "Failed",
            "FAIL",
            "Signal Cancel command not found"
        )

        return False

    log(f"{signal} Signal Cancel Successful")

    save_result(
        signal,
        route,
        "Signal Cancel",
        "Successful",
        "Successful",
        "PASS",
        ""
    )

    # Give panel time to process Signal Cancel
    time.sleep(2)

    # ========================================
    # ROUTE RELEASE
    # ========================================

    log("--------------------------------")
    log(f"Opening Signal Menu Again : {signal}")
    log("--------------------------------")

    safe_click(
        signal_x,
        signal_y
    )

    time.sleep(1)

    if not find_and_click("Route Release"):

        log(f"{signal} Route Release Failed")

        save_result(
            signal,
            route,
            "Route Release",
            "Released",
            "Failed",
            "FAIL",
            "Route Release command not found"
        )

        return False

    log(f"{signal} Route Released")

    save_result(
        signal,
        route,
        "Route Release",
        "Released",
        "Released",
        "PASS",
        ""
    )

    # Allow route release to complete
    time.sleep(5)

    log("--------------------------------")
    log(f"Route Release Completed : {route}")
    log("--------------------------------")

    return True



# =========================================================
# LOAD POINT TEST
# =========================================================

def load_point_test():
    try:

        wb = load_workbook(TOC_FILE)

        sheet = wb["TOC"]
        headers = get_toc_header_map(sheet)

        for row in sheet.iter_rows(min_row=2, values_only=True):

            if row[0] is None:
                continue

            point_value = toc_value(row, headers, "Points", "Point")
            track_value = toc_value(row, headers, "Track")
            if point_value is None:
                continue
            point = str(point_value).strip()
            track = str(track_value).strip() if track_value is not None else ""

            log("--------------------------------")
            log("POINT TEST LOADED")
            log(f"Point : {point}")
            log(f"Track : {track}")
            log("--------------------------------")

            return point, track

        log("No Point Test Found")

        return None, None

    except Exception as e:

        log(f"POINT ERROR : {e}")

        return None, None


def get_signal(signal):

    try:

        wb = load_workbook(CONFIG_FILE, data_only=True)

        for sheet_name in ("MAIN", "SHUNT", "CAL"):

            if sheet_name not in wb.sheetnames:
                continue

            sheet = wb[sheet_name]

            for row in sheet.iter_rows(min_row=2, values_only=True):

                if row[0] is None:
                    continue

                if str(row[0]).strip().upper() != str(signal).strip().upper():
                    continue

                if sheet_name == "MAIN":

                    aspects = int(row[1])

                    x = int(row[2])
                    y = int(row[3])

                elif sheet_name == "SHUNT":

                    aspects = 2

                    x = int(row[1])
                    y = int(row[2])

                else:      # CAL

                    aspects = 2

                    x = int(row[1])
                    y = int(row[2])

                wb.close()

                log("--------------------------------")
                log(f"Signal Found : {signal}")
                log(f"Sheet        : {sheet_name}")
                log(f"Aspects      : {aspects}")
                log(f"Menu X       : {x}")
                log(f"Menu Y       : {y}")
                log("--------------------------------")

                return x, y

        wb.close()

        log(f"Signal {signal} not found")

        return None, None

    except Exception as e:

        log(f"GET SIGNAL ERROR : {e}")

        return None, None

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

def record_lc_gate_coordinate():

    log("================================")
    log("LC GATE COORDINATE RECORDER")
    log("================================")

    build_lc_capture_list()

    if not loaded_lc_list:
        log("No LC Gates To Record")
        return

    for index, lc in enumerate(loaded_lc_list, start=1):

        coords = [None] * 3

        steps = [
            "RECORD CONTROL",
            "RECORD CLOSE",
            "RECORD OPEN"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(

                f"LC GATE {index} OF {len(loaded_lc_list)}\n\nLC : {lc}",

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

        log(f"Final Coordinates : {coords}")

        save_lc_gate_coordinate(
            lc,
            coords
        )

        log(f"{lc} CONTROL : {coords[0]}")
        log(f"{lc} CLOSE   : {coords[1]}")
        log(f"{lc} OPEN    : {coords[2]}")

    log("LC Gate Coordinates Saved")


def save_lc_gate_coordinate(lc, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["LC"]

    updated = False

    for row in sheet.iter_rows(min_row=2):

        if str(row[0].value).strip() == str(lc).strip():

            row[1].value = coords[0][0]
            row[2].value = coords[0][1]

            row[3].value = coords[1][0]
            row[4].value = coords[1][1]

            row[5].value = coords[2][0]
            row[6].value = coords[2][1]

            updated = True
            break

    if not updated:

        sheet.append([

            lc,

            coords[0][0],
            coords[0][1],

            coords[1][0],
            coords[1][1],

            coords[2][0],
            coords[2][1]

        ])

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{lc} saved successfully.")


def load_calling_on_track_entries():
    """Load Calling-On signal -> track mappings from the selected TOC."""

    entries = []

    if not TOC_FILE or not os.path.exists(TOC_FILE):
        log("TOC file not available for Calling-On track recording")
        return entries

    wb = None

    try:
        wb = load_workbook(
            TOC_FILE,
            data_only=True,
            read_only=True
        )

        if "TOC" not in wb.sheetnames:
            log("TOC sheet not found for Calling-On track recording")
            return entries

        ws = wb["TOC"]

        # IMPORTANT:
        # Build the header map before reading the rows.
        headers = get_toc_header_map(ws)

        signal_col = toc_col(headers, "Signal")
        route_col = toc_col(headers, "Route")
        track_col = toc_col(headers, "Track")

        if signal_col is None or route_col is None:
            log("CALLING-ON TOC ERROR : Signal/Route header not found")
            return entries

        if track_col is None:
            log("CALLING-ON TOC ERROR : Track header not found")
            return entries

        for row in ws.iter_rows(
                min_row=2,
                values_only=True
        ):

            signal_value = (
                row[signal_col]
                if signal_col < len(row)
                else None
            )

            route_value = (
                row[route_col]
                if route_col < len(row)
                else None
            )

            track_value = (
                row[track_col]
                if track_col < len(row)
                else None
            )

            if signal_value is None or route_value is None:
                continue

            signal = str(signal_value).strip().upper()

            # Calling-On signals end with C.
            if not signal.endswith("C"):
                continue

            route = str(route_value).strip()

            track = (
                str(track_value).strip()
                if track_value is not None
                else ""
            )

            if not track:
                continue

            entries.append({
                "signal": signal,
                "route": route,
                "track": track
            })

        log(
            f"CALLING-ON TRACK ENTRIES LOADED : "
            f"{len(entries)}"
        )

        for item in entries:
            log(
                f"CALLING-ON : "
                f"{item['signal']} | "
                f"{item['route']} | "
                f"{item['track']}"
            )

        return entries

    except Exception as e:

        log(
            f"CALLING-ON TRACK TOC ERROR : "
            f"{type(e).__name__} : {e}"
        )

        return entries

    finally:

        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

def get_recorded_track_names():
    """Return track names already stored in TRACK_CONFIG."""
    names = set()
    try:
        if not CONFIG_FILE or not os.path.exists(CONFIG_FILE):
            return names

        wb = load_workbook(CONFIG_FILE, data_only=True, read_only=True)
        if "TRACK_CONFIG" not in wb.sheetnames:
            wb.close()
            return names

        ws = wb["TRACK_CONFIG"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] not in (None, ""):
                names.add(str(row[0]).strip().upper())
        wb.close()
    except Exception as e:
        log(f"TRACK CONFIG READ ERROR : {e}")

    return names


def save_track_coordinate(track, x, y):
    """Create/update TRACK_CONFIG in the active configuration workbook."""
    if not CONFIG_FILE:
        log("Configuration file is not selected")
        return False

    try:
        wb = load_workbook(CONFIG_FILE)

        if "TRACK_CONFIG" not in wb.sheetnames:
            ws = wb.create_sheet("TRACK_CONFIG")
            ws.append(["TOC_TRACK", "X", "Y"])
        else:
            ws = wb["TRACK_CONFIG"]

        track_key = str(track).strip().upper()
        updated = False

        for row in range(2, ws.max_row + 1):
            value = ws.cell(row, 1).value
            if value is not None and str(value).strip().upper() == track_key:
                ws.cell(row, 1).value = str(track).strip()
                ws.cell(row, 2).value = int(x)
                ws.cell(row, 3).value = int(y)
                updated = True
                break

        if not updated:
            ws.append([str(track).strip(), int(x), int(y)])

        wb.save(CONFIG_FILE)
        wb.close()

        log(f"TRACK CONFIG SAVED : {track} -> ({x}, {y})")
        return True

    except Exception as e:
        log(f"TRACK CONFIG SAVE ERROR : {e}")
        try:
            wb.close()
        except Exception:
            pass
        return False


def record_calling_on_track_coordinates():
    """Record every unique track referenced by Calling-On signals in the TOC."""
    if not TOC_FILE:
        messagebox.showerror("Calling-On Track", "Please Load TOC First.", parent=root)
        return

    if not CONFIG_FILE or not os.path.exists(CONFIG_FILE):
        messagebox.showerror("Calling-On Track", "Please create/load Configuration First.", parent=root)
        return

    entries = load_calling_on_track_entries()
    unique_tracks = []
    seen = set()

    for item in entries:
        key = item["track"].strip().upper()
        if key and key not in seen:
            seen.add(key)
            unique_tracks.append(item)

    if not unique_tracks:
        messagebox.showinfo(
            "Calling-On Track",
            "No Calling-On tracks were found in the loaded TOC.",
            parent=root
        )
        return

    already_recorded = get_recorded_track_names()
    pending = [item for item in unique_tracks if item["track"].strip().upper() not in already_recorded]

    if not pending:
        messagebox.showinfo(
            "Calling-On Track",
            "All Calling-On track coordinates are already saved in the configuration.",
            parent=root
        )
        return

    log("========================================")
    log("CALLING-ON TRACK COORDINATE RECORDER")
    log("========================================")

    for index, item in enumerate(pending, start=1):
        track = item["track"]
        signal = item["signal"]

        action, result = capture_point(
            f"CALLING-ON TRACK {index} OF {len(pending)}\n\n"
            f"SIGNAL : {signal}\nTRACK : {track}",
            "RECORD TRACK COORDINATE"
        )

        if action == "CANCEL":
            log("Calling-On track recording cancelled")
            return

        if action == "UNDO":
            log("Undo is not used for the first-level track list; continuing")
            continue

        if result is None:
            continue

        x, y = result
        save_track_coordinate(track, x, y)
        log(f"Calling-On Track Saved : {track} -> ({x}, {y})")

    messagebox.showinfo(
        "Completed",
        "Calling-On track coordinates saved successfully in the configuration.",
        parent=root
    )


def record_track_coordinate():

    if not CONFIG_FILE or not os.path.exists(CONFIG_FILE):
        messagebox.showerror(
            "Track Recorder",
            "Please create or load the configuration first.",
            parent=root
        )
        return

    track = simpledialog.askstring(
        "Track Coordinate",
        "Enter TOC Track Name\n\nExample : 1CXTPR",
        parent=root
    )

    if not track or not track.strip():
        return

    track = track.strip()

    log("================================")
    log("TRACK COORDINATE RECORDER")
    log("================================")
    log(f"Track : {track}")

    action, result = capture_point(
        f"TRACK COORDINATE\n\nTRACK : {track}",
        "RECORD TRACK CENTER"
    )

    if action != "SPACE" or result is None:
        log("Track coordinate recording cancelled")
        return

    x, y = result

    if save_track_coordinate(track, x, y):
        messagebox.showinfo(
            "Track Recorder",
            f"Track coordinate saved successfully.\n\n{track}\nX = {x}\nY = {y}",
            parent=root
        )


def _track_name_variants(track_name):
    """1CXTPR -> 1CXTPR / 1CXT ; 1CXT -> 1CXT / 1CXTPR."""
    name = str(track_name).strip().upper()

    variants = [name]

    if name.endswith("PR"):
        variants.append(name[:-2])
    else:
        variants.append(name + "PR")

    return variants


def _search_track_sheet(ws, track_name):

    wanted = _track_name_variants(track_name)

    for row in ws.iter_rows(min_row=1, values_only=True):

        if not row or row[0] is None:
            continue

        if str(row[0]).strip().upper() not in wanted:
            continue

        try:
            return int(float(row[1])), int(float(row[2]))
        except (TypeError, ValueError, IndexError):
            continue

    return None, None


def get_track_coordinate(track_name):
    """Track coordinates come from:
       1. TRACK_CONFIG in the coordinate file (if recorded there), or
       2. the TRACK COORDS file the EDRC launcher supplies
          (TRACK_COORDINATES_CAPTURED.xlsx).
       The name is matched with and without the PR suffix
       (1CXTPR <-> 1CXT)."""

    # 1. TRACK_CONFIG inside the coordinate file
    try:
        if CONFIG_FILE and os.path.exists(CONFIG_FILE):

            wb = load_workbook(CONFIG_FILE, data_only=True)

            if "TRACK_CONFIG" in wb.sheetnames:

                x, y = _search_track_sheet(
                    wb["TRACK_CONFIG"],
                    track_name
                )

                wb.close()

                if x is not None:
                    log(f"TRACK COORDINATE (TRACK_CONFIG) : {track_name} -> ({x},{y})")
                    return x, y
            else:
                wb.close()

    except Exception as e:
        log(f"TRACK_CONFIG READ ERROR : {e}")

    # 2. TRACK COORDS file from the launcher
    candidates = [
        os.environ.get("EDRC_TRACK_COORDS"),
        os.environ.get("EDRC_TRACKS"),
        os.environ.get("EDRC_TRACK_COORDINATES"),
        os.path.join(os.getcwd(), "TRACK_COORDINATES_CAPTURED.xlsx")
    ]

    for item in candidates:

        if not item:
            continue

        item = os.path.abspath(item)

        if not os.path.exists(item):
            continue

        try:
            wb = load_workbook(item, data_only=True)

            for sheet in wb.worksheets:

                x, y = _search_track_sheet(sheet, track_name)

                if x is not None:
                    wb.close()
                    log(
                        f"TRACK COORDINATE ({os.path.basename(item)}) : "
                        f"{track_name} -> ({x},{y})"
                    )
                    return x, y

            wb.close()

        except Exception as e:
            log(f"TRACK COORDINATE FILE ERROR : {item} : {e}")

    log(
        f"TRACK COORDINATE NOT FOUND : {track_name} "
        f"(checked TRACK_CONFIG and the TRACK COORDS file)"
    )

    return None, None
import time
import pyautogui

def wait_until_track_red(track, timeout=20):

    x, y = get_track_coordinate(track)

    if x is None or y is None:
        log(f"Track Coordinate Not Found : {track}")
        return False

    x = int(x)
    y = int(y)

    log(f"Waiting for {track} to become RED...")
    log(f"Coordinate : ({x}, {y})")

    start = time.time()

    while time.time() - start < timeout:

        image = pyautogui.screenshot()

        r, g, b = image.getpixel((x, y))

        log(f"{track} RGB = ({r}, {g}, {b})")

        if r > g * 2 and r > b * 2:
            log(f"{track} is RED")

            return True

        time.sleep(1)

    log(f"Timeout waiting for {track} to become RED")
    return False

def get_lc_gate_state(lc):
    """Compatibility wrapper. Use the single LC detector everywhere."""
    return detect_lc_gate_state(lc)

def verify_lc_gate_before_route(signal, route):

    log("================================")
    log("VERIFYING LC GATE")
    log("================================")

    gates = load_lc_gates(signal, route)

    if not gates:
        return True

    lc_name = gates[0]

    state = get_lc_gate_state(lc_name)

    # --------------------------------
    # Gate already safe
    # --------------------------------

    if state in ("CLOSE", "CLOSED"):

        log("LC Gate Already Closed")

        save_result(
            signal,
            route,
            f"LC Gate {lc_name}",
            "CLOSE",
            state,
            "PASS",
            "Gate already closed"
        )

        return True

    # --------------------------------
    # Gate is open
    # --------------------------------

    if state == "OPEN":

        log(f"LC Gate {lc_name} is OPEN")

        coord = get_lc_gate_coordinates(lc_name)

        if coord is None:

            log("LC Gate coordinates not found")

            return False

        control_x, control_y = coord["control"]

        safe_click(control_x, control_y)

        time.sleep(1)

        log("Searching Receive...")

        if not find_and_click("Receive"):

            log("Receive Button Not Found")

            save_result(
                signal,
                route,
                f"LC Gate {lc_name}",
                "Receive",
                "Not Found",
                "FAIL",
                "Receive button not found"
            )

            return False

        log("Receive Clicked")

        time.sleep(5)

        state = get_lc_gate_state(lc_name)

        if state in ("CLOSE", "CLOSED"):

            log("LC Gate Closed Successfully")

            save_result(
                signal,
                route,
                f"LC Gate {lc_name}",
                "CLOSE",
                state,
                "PASS",
                "Gate closed successfully"
            )

            return True

        log("LC Gate Closing Failed")

        save_result(
            signal,
            route,
            f"LC Gate {lc_name}",
            "CLOSE",
            state,
            "FAIL",
            "Gate failed to close"
        )

        return False

    # --------------------------------
    # Unknown state
    # --------------------------------

    log(f"Unknown LC Gate State : {state}")

    save_result(
        signal,
        route,
        f"LC Gate {lc_name}",
        "CLOSE",
        str(state),
        "FAIL",
        "Unknown state"
    )

    return False

def _find_existing_test_panel():
    """Find the already-running Hitachi Test Panel window."""
    try:
        desktop = Desktop(backend="uia")
        for window in desktop.windows():
            try:
                title = window.window_text().strip()
                if "Test Panel" in title:
                    return window
            except Exception:
                continue
    except Exception as e:
        log(f"TEST PANEL SEARCH ERROR : {e}")
    return None


def open_test_panel():
    """Use an already-open Hitachi Test Panel.

    Walk-away runs must NOT contain a Panel.exe path or start Panel.exe.
    The operator/system opens the panel before starting the EDRC run.
    This function only detects, focuses and maximizes the existing panel.
    """
    log("========================================")
    log("CHECKING EXISTING HITACHI TEST PANEL")
    log("========================================")

    panel = _find_existing_test_panel()

    if panel is None:
        log("Test Panel is NOT open")
        if os.environ.get("EDRC_UNATTENDED") != "1":
            messagebox.showerror(
                "Panel Error",
                "Hitachi Test Panel is not open.\n\n"
                "Open the Test Panel first, then start the test."
            )
        return False

    log("Existing Test Panel detected")

    try:
        panel.restore()
    except Exception:
        pass

    try:
        panel.set_focus()
        log("Existing Test Panel focused")
    except Exception as e:
        log(f"Could not focus existing Test Panel : {e}")

    try:
        panel.maximize()
    except Exception:
        pass

    log("Existing Test Panel ready - no Panel.exe launch performed")
    return True

def load_all_signals():

    try:

        wb = load_workbook(TOC_FILE)

        sheet = wb["TOC"]

        signals = []

        for row in sheet.iter_rows(min_row=2, values_only=True):

            if row[0] is None:
                continue

            signal = str(row[0]).strip()

            if signal not in signals:
                signals.append(signal)

        wb.close()

        log("--------------------------------")
        log("SIGNALS LOADED")
        for signal in signals:
            log(signal)
        log("--------------------------------")

        return signals

    except Exception as e:

        log(f"LOAD SIGNALS ERROR : {e}")

        return []

def load_lc_gates(signal, route):
    """Return every LC gate listed for the exact signal/route."""

    if not TOC_FILE or not os.path.exists(TOC_FILE):
        log("TOC file not available while loading LC gates")
        return []

    wb = load_workbook(
        TOC_FILE,
        data_only=True,
        read_only=True
    )

    try:
        if "TOC" not in wb.sheetnames:
            log("TOC sheet not found")
            return []

        ws = wb["TOC"]
        headers = get_toc_header_map(ws)
        signal_col = toc_col(headers, "Signal")
        route_col = toc_col(headers, "Route")
        lc_col = toc_col(headers, "44_LCP", "44 LCP", "LC_GATE", "LC GATE", "LC")

        if signal_col is None or route_col is None or lc_col is None:
            log("TOC ERROR : Signal/Route/44_LCP header not found")
            return []

        for row in ws.iter_rows(min_row=2, values_only=True):
            signal_value = row[signal_col] if signal_col < len(row) else None
            route_value = row[route_col] if route_col < len(row) else None
            if (str(signal_value).strip().upper() == str(signal).strip().upper()
                    and str(route_value).strip().upper() == str(route).strip().upper()):
                value = row[lc_col] if lc_col < len(row) else None

                if value is None or str(value).strip() == "":
                    return []

                return [
                    item.strip()
                    for item in str(value).split(",")
                    if item.strip()
                ]

        log(f"No LC gate entry found for {signal} / {route}")
        return []

    finally:
        wb.close()

# =========================================================
# RECORD CRANK HANDLE
# =========================================================

def record_crank_handle_coordinate():

    log("================================")
    log("CRANK HANDLE COORDINATE RECORDER")
    log("================================")

    build_ch_capture_list()

    if not loaded_ch_list:
        log("No Crank Handles To Record")
        return

    for index, ch in enumerate(loaded_ch_list, start=1):

        coords = [None] * 4

        steps = [
            "RECORD CONTROL",
            "RECORD IN",
            "RECORD OUT",
            "RECORD FREE"
        ]

        step = 0

        while step < len(steps):

            action, result = capture_point(

                f"CRANK HANDLE {index} OF {len(loaded_ch_list)}\n\nPOINT : {ch}",

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

        log(f"Final Coordinates : {coords}")

        save_crank_handle_coordinate(
            ch,
            coords
        )

        log(f"{ch} CONTROL : {coords[0]}")
        log(f"{ch} IN      : {coords[1]}")
        log(f"{ch} OUT     : {coords[2]}")
        log(f"{ch} FREE    : {coords[3]}")

    log("Crank Handle Coordinates Saved")

def save_crank_handle_coordinate(ch, coords):

    wb = load_workbook(CONFIG_FILE)

    sheet = wb["CH"]

    updated = False

    for row in sheet.iter_rows(min_row=2):

        if str(row[0].value).strip() == str(ch).strip():

            row[1].value = coords[0][0]
            row[2].value = coords[0][1]

            row[3].value = coords[1][0]
            row[4].value = coords[1][1]

            row[5].value = coords[2][0]
            row[6].value = coords[2][1]

            row[7].value = coords[3][0]
            row[8].value = coords[3][1]

            updated = True
            break

    if not updated:

        sheet.append([

            ch,

            coords[0][0],
            coords[0][1],

            coords[1][0],
            coords[1][1],

            coords[2][0],
            coords[2][1],

            coords[3][0],
            coords[3][1]

        ])

    wb.save(CONFIG_FILE)
    wb.close()

    log(f"{ch} saved successfully.")

# =========================================================
# GET CRANK HANDLE STATE
# =========================================================

def get_crank_handle_state(point):

    ch_id = resolve_crank_handle_config_id(point)

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

    ch_headers = [
        str(h).strip().upper() if h is not None else ""
        for h in next(
            sheet.iter_rows(min_row=1, max_row=1, values_only=True),
            ()
        )
    ]

    in_lamp = None
    out_lamp = None
    free_lamp = None

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(ch_id).strip().upper():
            continue

        try:

            # --------------------------------
            # CURRENT CH SHEET FORMAT
            # --------------------------------
            # A = CrankHandle
            # B = Menu_X
            # C = Menu_Y
            # D = IN_X
            # E = IN_Y
            # F = OUT_X
            # G = OUT_Y
            # H = FREE_X
            # I = FREE_Y
            # --------------------------------

            # Columns are located by HEADER NAME, so a CH sheet with
            # an extra ECH pair (Menu / IN / OUT / ECH / FREE) is read
            # correctly instead of taking ECH as FREE.
            def _pair(*names):
                for want in names:
                    for idx, head in enumerate(ch_headers):
                        if head == want:
                            return (int(row[idx]), int(row[idx + 1]))
                return None

            in_lamp = _pair("IN_X", "IN X", "INX") or (
                int(row[3]), int(row[4])
            )

            out_lamp = _pair("OUT_X", "OUT X", "OUTX") or (
                int(row[5]), int(row[6])
            )

            free_lamp = _pair("FREE_X", "FREE X", "FREEX") or (
                int(row[7]), int(row[8])
            )

        except (TypeError, ValueError, IndexError):

            wb.close()

            log(
                f"Invalid CH coordinates : {point}"
            )

            return None

        break

    wb.close()

    if in_lamp is None:

        log(
            f"Crank Handle {point} not found"
        )

        return None

    image = pyautogui.screenshot()

    in_rgb = image.getpixel(in_lamp)
    out_rgb = image.getpixel(out_lamp)
    free_rgb = image.getpixel(free_lamp)

    log("--------------------------------")
    log(f"CRANK HANDLE : {point}")
    log(f"CH CONFIG ID : {ch_id}")

    log(f"IN RGB   : {in_rgb}")
    log(f"OUT RGB  : {out_rgb}")
    log(f"FREE RGB : {free_rgb}")

    # --------------------------------
    # IN = Yellow
    # --------------------------------

    in_on = (
        in_rgb[0] > 180 and
        in_rgb[1] > 180 and
        in_rgb[2] < 120
    )

    # --------------------------------
    # OUT = Red
    # --------------------------------

    out_on = (
        out_rgb[0] > 180 and
        out_rgb[1] < 120 and
        out_rgb[2] < 120
    )

    # --------------------------------
    # FREE = Green
    # --------------------------------

    free_on = (
        free_rgb[1] > 180 and
        free_rgb[0] < 120 and
        free_rgb[2] < 120
    )

    log(f"IN    : {in_on}")
    log(f"OUT   : {out_on}")
    log(f"FREE  : {free_on}")

    # --------------------------------
    # DETERMINE STATE
    # --------------------------------

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

# =========================================================
# VERIFY CRANK HANDLE BEFORE ROUTE
# =========================================================

def verify_crank_handle_before_route(signal, route):

    expected_points = load_expected_points(route)

    if not expected_points:
        return True

    for point, _ in expected_points:

        log("================================")
        log(f"VERIFY CRANK HANDLE : {point}")
        log("================================")

        state = get_crank_handle_state(point)

        if state is None:

            log(f"Crank Handle Configuration Missing : {point}")

            save_result(
                signal,
                route,
                f"Crank Handle {point}",
                "Configured",
                "Missing",
                "FAIL",
                "Configuration Missing"
            )

            return False

        if state in ("IN", "FREE"):

            save_result(
                signal,
                route,
                f"Crank Handle {point}",
                "IN/FREE",
                state,
                "PASS",
                "Ready for Route"
            )

            continue

        if state == "OUT":

            log(f"Crank Handle {point} is OUT")

            save_result(
                signal,
                route,
                f"Crank Handle {point}",
                "IN",
                "OUT",
                "FAIL",
                "Recovering..."
            )

            if not recover_crank_handle(signal, route, point):

                log(f"Recovery Failed : {point}")

                return False

            state = get_crank_handle_state(point)

            if state not in ("IN", "FREE"):

                save_result(
                    signal,
                    route,
                    f"Crank Handle {point}",
                    "IN/FREE",
                    state,
                    "FAIL",
                    "Recovery Failed"
                )

                return False

            save_result(
                signal,
                route,
                f"Crank Handle {point}",
                "IN/FREE",
                state,
                "PASS",
                "Recovered Successfully"
            )

            continue

        log(f"Unknown Crank Handle State : {point}")

        save_result(
            signal,
            route,
            f"Crank Handle {point}",
            "IN/FREE",
            str(state),
            "FAIL",
            "Unknown State"
        )

        return False

    return True

# =========================================================
# RECOVER CRANK HANDLE
# =========================================================

def recover_crank_handle(signal, route, point):

    ch_id = resolve_crank_handle_config_id(point)
    log(f"CH CONFIG ID : {ch_id}")

    wb = load_workbook(
        CONFIG_FILE,
        data_only=True,
        read_only=True
    )

    if "CH" not in wb.sheetnames:

        wb.close()
        return False

    sheet = wb["CH"]

    control = None

    for row in sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if row[0] is None:
            continue

        if str(row[0]).strip().upper() != str(ch_id).strip().upper():
            continue

        control = (
            int(row[1]),
            int(row[2])
        )

        break

    wb.close()

    if control is None:

        log(f"Crank Handle {point} not found")

        return False

    log("--------------------------------")
    log(f"Recovering Crank Handle : {point}")
    log("--------------------------------")

    safe_click(control[0], control[1])

    time.sleep(1)

    if not find_and_click("Receive"):

        log("Receive button not found")

        save_result(
            signal,
            route,
            f"Crank Handle {point} Recovery",
            "Receive",
            "Not Found",
            "FAIL",
            "Receive button not found"
        )

        return False

    log("Receive Clicked")

    time.sleep(3)

    state = get_crank_handle_state(ch_id)

    if state in ("IN", "FREE"):

        save_result(
            signal,
            route,
            f"Crank Handle {point} Recovery",
            "IN/FREE",
            state,
            "PASS",
            "Recovery Successful"
        )

        log("Recovery Successful")

        return True

    save_result(
        signal,
        route,
        f"Crank Handle {point} Recovery",
        "IN/FREE",
        state,
        "FAIL",
        "Recovery Failed"
    )

    log("Recovery Failed")

    return False

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


# =========================================================
# EDRC WALK-AWAY WINDOW CONTROL
# =========================================================

def restore_edrc_window():
    """Restore this EDRC child window after the test has finished."""
    try:
        root.deiconify()
        root.state("zoomed")
        root.lift()
        root.attributes("-topmost", True)
        root.after(300, lambda: root.attributes("-topmost", False))
        root.update_idletasks()
        root.update()
        log("EDRC WINDOW RESTORED AFTER TEST")
    except Exception as e:
        log(f"EDRC WINDOW RESTORE ERROR : {e}")


def minimize_for_edrc_walkaway():
    """Hide the child EDRC GUI while the Test Panel is being operated."""
    if os.environ.get("EDRC_UNATTENDED") != "1":
        return

    try:
        root.iconify()
        root.update_idletasks()
        log("EDRC WINDOW MINIMIZED - WALK-AWAY PANEL OPERATION STARTING")
    except Exception as e:
        log(f"EDRC WINDOW MINIMIZE ERROR : {e}")


# =========================================================
# TEST ROUTE ENGINE
# =========================================================

def route_has_test_data(signal, route):
    """This program tests POINTS, LC GATE, CRANK HANDLE and the
    calling-on TRACK. If a TOC row has none of those columns filled
    (e.g. 8 / 8_L and 25 / 25_J), there is nothing to test for that
    route, so it is skipped instead of being set and released."""

    try:
        wb = load_workbook(TOC_FILE, data_only=True)
        ws = wb["TOC"]
        headers = get_toc_header_map(ws)

        for row in ws.iter_rows(min_row=2, values_only=True):

            signal_value = toc_value(row, headers, "Signal")
            route_value = toc_value(row, headers, "Route")

            if signal_value is None or route_value is None:
                continue

            if str(signal_value).strip().upper() != str(signal).strip().upper():
                continue

            if str(route_value).strip().upper() != str(route).strip().upper():
                continue

            fields = {
                "POINTS": toc_value(row, headers, "Points", "Point"),
                "CRANK_HANDLE": toc_value(
                    row, headers, "CRANK_HANDLE", "CRANK HANDLE", "CH"
                ),
                "44_LCP": toc_value(
                    row, headers,
                    "44_LCP", "44 LCP", "LC_GATE", "LC GATE", "LC"
                ),
                "TRACK": toc_value(row, headers, "Track")
            }

            wb.close()

            filled = [
                name
                for name, value in fields.items()
                if value is not None and str(value).strip() != ""
            ]

            if filled:
                log(f"{signal} / {route} : TOC DATA -> {', '.join(filled)}")
                return True

            log(
                f"{signal} / {route} : NO POINTS / CRANK HANDLE / LC / TRACK "
                f"IN TOC - ROUTE SKIPPED"
            )
            return False

        wb.close()

        log(f"{signal} / {route} : ROW NOT FOUND IN TOC - ROUTE SKIPPED")
        return False

    except Exception as e:
        log(f"TOC CHECK ERROR : {signal} / {route} : {e}")
        return True


def test_route_engine():
    if TOC_FILE is None:
        log("START FAILED : TOC NOT LOADED")
        return
    if not CONFIG_FILE or not os.path.exists(CONFIG_FILE):
        log("START FAILED : CONFIGURATION NOT LOADED")
        return
    # ========================================
    # OPEN HITACHI TEST PANEL
    # ========================================

    if not open_test_panel():
        return

    # EDRC walk-away: once the existing panel is confirmed, hide this
    # automation window so all mouse/keyboard activity is directed at the
    # Test Panel.
    minimize_for_edrc_walkaway()

    # ========================================
    # CREATE REPORT
    # ========================================

    create_report()

    log("========================================")
    log("ROUTE ENGINE STARTED")
    log("========================================")

    log("Starting automation in 3 seconds...")
    time.sleep(3)

    # ========================================
    # LOAD ALL SIGNALS
    # ========================================

    signals = load_all_signals()

    calling_on_data = load_calling_on_routes()

    for signal in signals:

        routes = load_routes_for_signal(signal)

        if not routes:
            continue

        signal_x, signal_y = get_signal(signal)

        if signal_x is None:
            log(f"Signal {signal} Coordinates Not Found")
            continue

        log("--------------------------------")
        log(f"Signal : {signal}")
        log(f"Routes Found : {len(routes)}")

        for r in routes:
            log(f"   {r}")

        # =====================================
        # FIND CALLING-ON DETAILS FOR SIGNAL
        # =====================================

        calling_on_routes = []
        track = None

        for item in calling_on_data:

            if item["signal"] == signal:

                calling_on_routes.append(item["route"])

                if track is None:
                    track = item["track"]

        log("--------------------------------")
        log(f"Calling-On Routes : {calling_on_routes}")
        log(f"Track : {track}")
        log("--------------------------------")

        # ====================================
        # MAIN ROUTES
        # ====================================

        # Skip Main Route execution for Calling-On signals
        if signal in [item["signal"] for item in calling_on_data]:
            log(f"{signal} is a Calling-On signal. Skipping Main Route block.")
        else:

            for route in routes:



                log("--------------------------------")
                log(f"Operating Route : {route}")

                # Nothing to test for this route in the TOC
                if not route_has_test_data(signal, route):
                    continue

                # This route must be completely finalized before
                # the automation moves to the next route.
                route_checks_passed = True
                route_selected = False

                # PRECHECKS
                # Skip Main Route logic for Calling-On signals
                if signal.endswith("C"):
                    log(f"{signal} is a Calling-On Signal")
                    break

                if not verify_crank_handle_before_route(signal, route):
                    continue

                # --------------------------------
                # LC PRE-CHECK
                # --------------------------------

                # Do not block route selection if the
                # LC indication is temporarily UNKNOWN.
                # The actual LC verification will be
                # performed again AFTER the route is set.

                if load_lc_gates(signal, route):

                    log(
                        "LC Gate present. "
                        "Route selection will continue."
                    )

                else:

                    log("No LC Gate associated with route")

                # --------------------------------
                # OPEN SIGNAL MENU
                # --------------------------------

                log(f"Moving mouse to Signal {signal}")

                safe_click(signal_x, signal_y)

                time.sleep(1)

                # --------------------------------
                # SELECT MAIN ROUTE
                # --------------------------------

                log(f"Looking for Route : {route}")

                if not find_and_click(route):
                    log(f"{route} NOT FOUND")

                    continue

                log(f"{route} Selected")
                route_selected = True

                # ========================================
                # WAIT AFTER ROUTE SET
                # ========================================

                time.sleep(5)

                # ========================================
                # VERIFY LC GATE AFTER ROUTE SET
                # ========================================

                log("================================")
                log("VERIFYING LC GATE AFTER ROUTE SET")
                log("================================")

                gates = load_lc_gates(
                    signal,
                    route
                )

                if gates:

                    if not verify_lc_gate_before_route(
                            signal,
                            route
                    ):
                        log("--------------------------------")
                        log("LC GATE VERIFICATION FAILED")
                        log("--------------------------------")

                        route_checks_passed = False

                    else:
                        log("LC Gate Verification PASS")

                else:

                    log("No LC Gate associated with route")

                # --------------------------------
                # VERIFY BOTH POINTS
                # --------------------------------

                passed, failed_point, expected_state = verify_points(
                    signal,
                    route
                )

                if not passed:

                    log("--------------------------------")
                    log("POINT LOCK VERIFICATION FAILED")
                    log("--------------------------------")

                    route_checks_passed = False

                else:

                    log("--------------------------------")
                    log("BOTH POINTS LOCK VERIFICATION PASSED")
                    log("--------------------------------")

                # --------------------------------
                # LOCKED LC GATE TEST
                # --------------------------------

                gates = load_lc_gates(signal, route)

                log(f"Route {route} -> LC Gates = {gates}")

                if len(gates) > 0:

                    if not verify_locked_lc_gates(signal, route):
                        log("LOCKED LC GATE TEST FAILED")
                        route_checks_passed = False

                else:

                    log("LC Gate Test Skipped")

                # --------------------------------
                # LOCKED CRANK HANDLE TEST
                # --------------------------------

                if not verify_locked_crank_handle(
                        signal,
                        route
                ):
                    log("--------------------------------")
                    log("LOCKED CRANK HANDLE TEST FAILED")
                    log("--------------------------------")
                    route_checks_passed = False

                else:
                    log("ALL CRANK HANDLES VERIFIED")

                # ========================================
                # FINALISE THIS ROUTE
                # ========================================
                #
                # IMPORTANT:
                # Signal Cancel + Route Release belongs to the
                # CURRENT route. It must happen before the loop
                # is allowed to start the next route.
                #
                # Do NOT use "continue" or "return" above this
                # point for a selected route.
                # ========================================

                if route_selected:

                    log("========================================")
                    log(f"FINALISING ROUTE : {route}")
                    log("========================================")

                    if not route_checks_passed:
                        log(
                            f"One or more verifications failed for "
                            f"{route}."
                        )
                        log(
                            "Proceeding with Signal Cancel + "
                            "Route Release for the current route."
                        )

                    if not release_route(signal, route):

                        log(
                            f"Signal Cancel / Route Release FAILED : "
                            f"{route}"
                        )

                    else:

                        log(
                            f"Signal Cancel + Route Release "
                            f"COMPLETED : {route}"
                        )

                    flush_route_result()

                    log(
                        f"Waiting 8 seconds before next route : "
                        f"{route}"
                    )

                    time.sleep(8)

                else:
                    log(
                        f"Route {route} was not selected. "
                        f"No Signal Cancel / Route Release required."
                    )

        # ========================================
        # CALLING-ON ROUTES
        # ========================================

        if calling_on_routes:

            log("========================================")
            log(f"CALLING-ON FOR SIGNAL {signal}")
            log("========================================")

            if track:

                if not track_down(track):
                    log("Track Down Failed")
                    continue

                log("Waiting for Track RED")

                if not wait_until_track_red(track):
                    log("Track did not become RED")

                    continue

                time.sleep(2)

            for calling_on_route in calling_on_routes:

                # Nothing to test for this calling-on route in the TOC
                if not route_has_test_data(signal, calling_on_route):
                    continue

                calling_signal = signal

                if not calling_signal:
                    log(
                        f"Calling-On Signal Not Found : "
                        f"{calling_on_route}"
                    )

                    continue

                log(
                    f"Calling-On Signal : "
                    f"{calling_signal}"
                )

                calling_route_checks_passed = True
                calling_route_selected = False

                signal_x, signal_y = get_signal(calling_signal)

                if signal_x is None:
                    log(f"Coordinates Not Found : {calling_signal}")
                    continue

                log(f"Moving to Calling-On Signal : {calling_signal}")
                # PRECHECKS

                if not verify_crank_handle_before_route(calling_signal, calling_on_route):
                    continue

                # --------------------------------
                # LC PRE-CHECK
                # --------------------------------

                if load_lc_gates(
                        calling_signal,
                        calling_on_route
                ):

                    log(
                        "Calling-On LC Gate present. "
                        "Route selection will continue."
                    )

                else:

                    log("No Calling-On LC Gate associated with route")
                safe_click(signal_x, signal_y)

                time.sleep(1)

                log(f"Selecting Calling-On Route : {calling_on_route}")

                if not find_and_click(calling_on_route):
                    log(f"{calling_on_route} NOT FOUND")
                    continue

                calling_route_selected = True

                time.sleep(5)

                # --------------------------------
                # VERIFY LC GATE AFTER CALLING-ON ROUTE SET
                # --------------------------------

                log("================================")
                log("VERIFYING LC GATE AFTER CALLING-ON ROUTE SET")
                log("================================")

                calling_gates_after_route = load_lc_gates(
                    calling_signal,
                    calling_on_route
                )

                if calling_gates_after_route:

                    if not verify_lc_gate_before_route(
                            calling_signal,
                            calling_on_route
                    ):
                        log("LC GATE VERIFICATION AFTER CALLING-ON ROUTE SET FAILED")
                        calling_route_checks_passed = False

                    else:
                        log("Calling-On LC Gate Verification After Route : PASS")

                else:

                    log("No Calling-On LC Gate associated with route")

                # --------------------------------
                # VERIFY BOTH POINTS
                # --------------------------------

                passed, failed_point, expected_state = verify_points(
                    calling_signal,
                    calling_on_route
                )

                if not passed:

                    log("Calling-On Verification FAILED")

                    recovery_success = recover_point(
                        calling_signal,
                        calling_on_route,
                        failed_point,
                        expected_state
                    )

                    if not recovery_success:
                        calling_route_checks_passed = False
                    else:
                        log(
                            "Calling-On point recovery completed. "
                            "Skipping duplicate point verification."
                        )

                log("Calling-On Verification PASS")

                # --------------------------------
                # LOCKED LC GATE TEST
                # --------------------------------

                gates = load_lc_gates(
                    calling_signal,
                    calling_on_route
                )

                if len(gates) > 0:

                    if not verify_locked_lc_gates(
                            calling_signal,
                            calling_on_route
                    ):
                        log("--------------------------------")
                        log("LOCKED LC GATE TEST FAILED")
                        log("--------------------------------")
                        calling_route_checks_passed = False

                else:

                    log("LC Gate Test Skipped")

                # --------------------------------
                # LOCKED CRANK HANDLE TEST
                # --------------------------------

                if not verify_locked_crank_handle(
                        calling_signal,
                        calling_on_route
                ):
                    log("--------------------------------")
                    log("LOCKED CRANK HANDLE TEST FAILED")
                    log("--------------------------------")
                    calling_route_checks_passed = False

                else:
                    log("LOCKED CRANK HANDLE TEST PASSED")

                # --------------------------------
                # FINALISE CURRENT CALLING-ON ROUTE
                # --------------------------------
                #
                # Signal Cancel + Route Release MUST be completed
                # for this selected route before the next
                # Calling-On route starts.
                # --------------------------------

                if calling_route_selected:

                    log("========================================")
                    log(
                        f"FINALISING CALLING-ON ROUTE : "
                        f"{calling_on_route}"
                    )
                    log("========================================")

                    if not calling_route_checks_passed:

                        log(
                            f"One or more verifications failed for "
                            f"{calling_on_route}."
                        )

                        log(
                            "Proceeding with Signal Cancel + "
                            "Route Release for the current route."
                        )

                    if not release_route(
                            calling_signal,
                            calling_on_route
                    ):

                        log(
                            "Calling-On Signal Cancel / "
                            "Route Release Failed : "
                            f"{calling_on_route}"
                        )

                    else:

                        log(
                            "Calling-On Signal Cancel + "
                            "Route Release Completed : "
                            f"{calling_on_route}"
                        )

                    flush_route_result()

                    time.sleep(3)

                else:

                    log(
                        f"Calling-On route {calling_on_route} "
                        f"was not selected. No release required."
                    )

            if track:

                log(f"Track Up : {track}")

                if not track_up(track):
                    log("Track Up Failed")

                time.sleep(2)



    # ========================================
    # FINISH
    # ========================================

    log("========================================")
    log("ALL ROUTES COMPLETED")
    log("========================================")

    log("EXITING test_route_engine()")

    finalize_report()

    if REPORT_FILE and os.path.exists(REPORT_FILE):
        log(f"REPORT READY : {os.path.abspath(REPORT_FILE)}")

        # Manual runs may still open the report for inspection.  EDRC
        # walk-away runs must leave Excel closed so the master can collect
        # the workbook cleanly.
        if os.environ.get("EDRC_UNATTENDED") != "1":
            try:
                os.startfile(os.path.abspath(REPORT_FILE))
            except Exception as e:
                log(f"FAILED TO OPEN REPORT : {e}")
    else:
        log("Report file not found")

    # In EDRC walk-away mode restore the child EDRC window after the test.
    if os.environ.get("EDRC_UNATTENDED") == "1":
        try:
            root.after(200, restore_edrc_window)
        except Exception as e:
            log(f"EDRC WINDOW RESTORE SCHEDULE ERROR : {e}")

    return


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


def open_record_system_controls():
    """Scrollable grouped UI for all coordinate/configuration recording tools."""

    win = tk.Toplevel(root)
    win.title("Record System Controls")
    win.geometry("680x650")
    win.minsize(620, 500)
    win.configure(bg="#e9edf2")
    win.transient(root)
    win.grab_set()

    # =====================================================
    # FIXED HEADER
    # =====================================================

    header = tk.Frame(win, bg="#0f172a", height=78)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(
        header,
        text="RECORD SYSTEM CONTROLS",
        font=("Segoe UI", 18, "bold"),
        bg="#0f172a",
        fg="white"
    ).pack(pady=(14, 0))

    tk.Label(
        header,
        text="Save panel coordinates directly into the active configuration",
        font=("Segoe UI", 9),
        bg="#0f172a",
        fg="#94a3b8"
    ).pack(pady=(2, 8))

    # =====================================================
    # SCROLLABLE CONTENT AREA
    # =====================================================

    content_area = tk.Frame(win, bg="#e9edf2")
    content_area.pack(fill="both", expand=True, padx=10, pady=(10, 4))

    canvas = tk.Canvas(
        content_area,
        bg="#e9edf2",
        highlightthickness=0,
        borderwidth=0
    )

    scrollbar = ttk.Scrollbar(
        content_area,
        orient="vertical",
        command=canvas.yview
    )

    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    body = tk.Frame(canvas, bg="#e9edf2")

    body_window = canvas.create_window(
        (0, 0),
        window=body,
        anchor="nw"
    )

    def update_scroll_region(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def resize_body(event):
        canvas.itemconfigure(body_window, width=event.width)

    body.bind("<Configure>", update_scroll_region)
    canvas.bind("<Configure>", resize_body)

    def mouse_wheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind("<MouseWheel>", mouse_wheel)
    body.bind("<MouseWheel>", mouse_wheel)

    # =====================================================
    # SECTION / BUTTON HELPERS
    # =====================================================

    def section(parent, title, subtitle):
        frame = tk.Frame(parent, bg="white", bd=1, relief="solid")

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

    def action_button(parent, text, command, color="#2563eb"):
        btn = tk.Button(
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
            width=22,
            height=2
        )
        btn.pack(side="left", padx=6, pady=8)
        return btn

    # =====================================================
    # SIGNALS
    # =====================================================

    signalling = section(
        body,
        "SIGNALS",
        "Record signal menu and indication coordinates"
    )
    signalling.pack(fill="x", pady=(0, 10))

    row = tk.Frame(signalling, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 4))

    action_button(row, "Main Signal", record_main_signals, "#2563eb")
    action_button(row, "Shunt Signal", record_shunt_signals, "#2563eb")
    action_button(row, "Calling-On Signal", record_calling_on_signals, "#7c3aed")

    # =====================================================
    # POINT / CRANK HANDLE
    # =====================================================

    point_frame = section(
        body,
        "POINT & CRANK HANDLE",
        "Record control, indication and crank-handle coordinates"
    )
    point_frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(point_frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 4))

    action_button(row, "Point Indication", record_point_indication_coordinate, "#059669")
    action_button(row, "Point Control", record_point_control_coordinate, "#059669")

    row2 = tk.Frame(point_frame, bg="white")
    row2.pack(fill="x", padx=8, pady=(0, 4))

    action_button(row2, "Crank Handle", record_crank_handle_coordinate, "#059669")

    # =====================================================
    # LC GATE & TRACK
    # =====================================================

    system_frame = section(
        body,
        "LC GATE & TRACK",
        "Record LC gate and track coordinates"
    )
    system_frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(system_frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 4))

    action_button(row, "LC Gate", record_lc_gate_coordinate, "#ea580c")
    action_button(row, "Track", record_track_coordinate, "#ea580c")

    # =====================================================
    # CALLING-ON TRACKS
    # =====================================================

    calling_track_frame = section(
        body,
        "CALLING-ON TRACKS",
        "Record track coordinates used by Calling-On signals"
    )
    calling_track_frame.pack(fill="x", pady=(0, 10))

    row = tk.Frame(calling_track_frame, bg="white")
    row.pack(fill="x", padx=8, pady=(0, 4))

    action_button(
        row,
        "Calling-On Tracks",
        record_calling_on_track_coordinates,
        "#ea580c"
    )

    # =====================================================
    # EXTRA SPACE AT BOTTOM
    # =====================================================

    tk.Frame(body, height=20, bg="#e9edf2").pack()

    # =====================================================
    # FIXED FOOTER
    # =====================================================

    footer = tk.Frame(win, bg="#e9edf2")
    footer.pack(fill="x", padx=18, pady=(4, 14))

    tk.Button(
        footer,
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
    ).pack()

    canvas.yview_moveto(0)


# =========================================================
# EDRC RUNNER COMPATIBILITY
# =========================================================
# The external EDRC runner looks for these three button names.
# They intentionally use the project's standard files:
#     TOC.xlsx
#     HAH COORDINATES.xlsx
#
# No file-selection dialog is used in walk-away mode.

def load_excel():
    """
    EDRC runner entry point.

    IMPORTANT:
    In walk-away mode the EDRC master already selects the TOC and passes
    its exact path to this program through EDRC_LIST (and EDRC_TOC).
    Never force the local/old TOC.xlsx when EDRC has supplied another file.
    """

    base_dir = os.path.dirname(os.path.abspath(__file__))

    # EDRC master -> exact TOC selected by the operator.
    # EDRC_LIST is the file actually handed to the program's LIST input.
    # EDRC_TOC is retained as an additional compatibility fallback.
    toc_path = (
        os.environ.get("EDRC_LIST")
        or os.environ.get("EDRC_TOC")
    )

    if toc_path:
        toc_path = os.path.abspath(toc_path)

        if os.path.exists(toc_path):
            log(
                f"EDRC: loading SELECTED TOC -> "
                f"{toc_path}"
            )
            load_toc(toc_path, show_message=False)

            if TOC_FILE:
                log(
                    f"EDRC: TOC READY -> "
                    f"{TOC_FILE}"
                )
            else:
                log("EDRC: SELECTED TOC LOAD FAILED")

            return

        log(
            f"EDRC: SELECTED TOC NOT FOUND -> "
            f"{toc_path}"
        )

    # Manual-run fallback only.
    # This is deliberately NOT used when EDRC supplied a valid TOC.
    fallback = os.path.join(base_dir, "TOC.xlsx")

    if os.path.exists(fallback):
        log(
            f"MANUAL MODE: loading fallback TOC -> "
            f"{fallback}"
        )
        load_toc(fallback, show_message=False)

        if TOC_FILE:
            log(
                f"TOC READY -> "
                f"{TOC_FILE}"
            )
        else:
            log("TOC LOAD FAILED")
    else:
        log("TOC LOAD FAILED: no EDRC TOC and no local TOC.xlsx")


def load_universal_coordinates():
    """EDRC runner entry point: load the coordinate file.

    The EDRC master passes the COORDS file the operator selected
    (e.g. "E2E COORDINATES .xlsx") through EDRC_COORDS / EDRC_CONFIG.
    Use that first; only fall back to a local HAH COORDINATES.xlsx
    when the master supplied nothing.
    """

    base_dir = os.path.dirname(os.path.abspath(__file__))

    candidates = [
        os.environ.get("EDRC_COORDS"),
        os.environ.get("EDRC_CONFIG"),
        os.environ.get("EDRC_COORDINATES"),
        os.environ.get("EDRC_UNIVERSAL_COORDINATES"),
        os.path.join(base_dir, "HAH COORDINATES.xlsx"),
        os.path.join(os.getcwd(), "HAH COORDINATES.xlsx")
    ]

    config_path = None

    for item in candidates:

        if not item:
            continue

        item = os.path.abspath(item)

        if os.path.exists(item):
            config_path = item
            break

    if config_path is None:
        log(
            "EDRC: coordinate file not found. Checked : "
            + " | ".join(str(c) for c in candidates if c)
        )
        return

    log(f"EDRC: loading COORDINATES -> {config_path}")
    load_config(config_path, show_message=False)
    log(f"EDRC: COORDINATES READY -> {CONFIG_FILE}")


def start_automation():
    """EDRC runner entry point: start the existing route engine.

    The engine runs in a BACKGROUND THREAD. Running it on the Tk main
    thread froze the window, so the log (written through root.after)
    only appeared at the very end and the launcher believed the
    program had never started.
    """

    global running

    log("EDRC: START AUTOMATION")

    def worker():
        global running
        try:
            test_route_engine()
        except Exception as e:
            log(
                f"EDRC: TEST ENGINE ERROR : "
                f"{type(e).__name__} : {e}"
            )
        finally:
            running = False

    running = True

    threading.Thread(
        target=worker,
        daemon=True
    ).start()

    return True


# =========================================================
# GUI
# =========================================================

root = tk.Tk()

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

    text="TEST FOR POINT/CRANK HANDLE",

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

    text="EXISTING CONFIG",

    command=existing_signal,

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
    load_toc,
    "#2563eb"
).pack(pady=5)

create_button(
    "RECORD SYSTEM CONTROLS",
    open_record_system_controls,
    "#059669"
).pack(pady=5)
create_button(
    "START TESTING",
    test_route_engine,
    "#7c3aed"
).pack(pady=8)

# =========================================================
# EDRC WALK-AWAY CONTROLS
# =========================================================
# These buttons are deliberately named exactly as the EDRC
# launcher expects. They are also useful for manual testing.

edrc_frame = tk.LabelFrame(
    left_panel,
    text="EDRC WALK-AWAY",
    bg="white",
    fg="#475569",
    font=("Segoe UI", 9, "bold"),
    bd=1,
    relief="solid"
)
edrc_frame.pack(fill="x", padx=12, pady=(4, 10))

tk.Button(
    edrc_frame,
    text="load_excel",
    command=load_excel,
    bg="#0ea5e9",
    fg="white",
    activebackground="#0ea5e9",
    activeforeground="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=(8, 3))

tk.Button(
    edrc_frame,
    text="load_universal_coordinates",
    command=load_universal_coordinates,
    bg="#0ea5e9",
    fg="white",
    activebackground="#0ea5e9",
    activeforeground="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=3)

tk.Button(
    edrc_frame,
    text="start_automation",
    command=start_automation,
    bg="#7c3aed",
    fg="white",
    activebackground="#7c3aed",
    activeforeground="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    width=24,
    height=1
).pack(pady=(3, 8))

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

        "LOCK_ROUTE"

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
# RUN
# =========================================================

root.mainloop()
