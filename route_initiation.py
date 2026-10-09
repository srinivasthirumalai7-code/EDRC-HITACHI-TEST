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

# =========================================================

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

REPORT_FILE = "Route_Intiation.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

auto_throw_data = []
loaded_signal_list = []

running = False

capture_queue = []

capture_index = 0

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
# REPORT DATA STORAGE
# =========================================================

report_rows = []

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

def load_universal_coordinates():
    """Loads the SAME multi-sheet 'complete yard' coordinate file produced
    by the Signal Clearance program (sheets: MAIN, SHUNT, CAL, POINT, CH,
    LC - point machines/crank handles/LC gates are ignored here since this
    program only tests signal routes). Only Menu + Route Initiation
    indicator are needed for route initiation testing, so everything else
    in those sheets (aspect lamps, snapshot paths, etc.) is skipped.

    Columns are found by matching their header text in row 1 - "Signal" /
    "Menu_X" / "Menu_Y" / "RouteInit_X" / "RouteInit_Y" - wherever they
    sit, regardless of any other columns present or their order. Same
    approach as the TOC loader below.

    Independent of TOC - can be loaded before or after Import TOC / Route
    Intiation Test, in any order. Each run REPLACES the current signals
    dict with what's found in the file, same as loading any other config."""

    global signals

    file = filedialog.askopenfilename(
        title="Select Universal Yard Coordinates",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)

    def parse_int(value):
        """Blank cells, empty strings, or bad values become None instead
        of raising - so one incomplete row doesn't abort the whole load."""

        if value is None:
            return None

        if isinstance(value, str) and value.strip() == "":
            return None

        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    def parse_point(x_val, y_val):
        """Only returns a point if BOTH coordinates are valid numbers."""

        x = parse_int(x_val)
        y = parse_int(y_val)

        if x is None or y is None:
            return None

        return [x, y]

    def read_sheet_by_header(ws, sig_type):
        """Reads Signal / Menu_X / Menu_Y / RouteInit_X / RouteInit_Y from
        one sheet by header name. Returns (loaded_count, skipped_count)."""

        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)

        if not header_row:
            log(f"{ws.title}: sheet has no header row - skipped")
            return 0, 0

        header_index = {}

        for idx, val in enumerate(header_row):
            if val is None:
                continue
            header_index[str(val).strip().upper()] = idx

        signal_col = header_index.get("SIGNAL")
        menu_x_col = header_index.get("MENU_X")
        menu_y_col = header_index.get("MENU_Y")
        ri_x_col = header_index.get("ROUTEINIT_X")
        ri_y_col = header_index.get("ROUTEINIT_Y")

        required = {
            "Signal": signal_col, "Menu_X": menu_x_col, "Menu_Y": menu_y_col,
            "RouteInit_X": ri_x_col, "RouteInit_Y": ri_y_col
        }
        missing = [name for name, col in required.items() if col is None]

        if missing:
            log(f"{ws.title}: missing column header(s) {missing} - sheet skipped")
            return 0, 0

        def cell(row, idx):
            return row[idx] if len(row) > idx else None

        loaded = 0
        skipped = 0

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row or cell(row, signal_col) is None:
                continue

            signal = str(cell(row, signal_col)).strip().upper()

            menu = parse_point(cell(row, menu_x_col), cell(row, menu_y_col))
            route_init = parse_point(cell(row, ri_x_col), cell(row, ri_y_col))

            if not menu or not route_init:
                skipped += 1
                continue

            signals[signal] = {
                "type": sig_type,
                "menu": menu,
                "ASPECT": route_init
            }
            loaded += 1

        return loaded, skipped

    signals = {}

    loaded_count = 0
    skipped_incomplete = 0

    if "MAIN" in wb.sheetnames:
        l, s = read_sheet_by_header(wb["MAIN"], "MAIN")
        loaded_count += l
        skipped_incomplete += s

    if "SHUNT" in wb.sheetnames:
        l, s = read_sheet_by_header(wb["SHUNT"], "SHUNT")
        loaded_count += l
        skipped_incomplete += s

    if "CAL" in wb.sheetnames:
        l, s = read_sheet_by_header(wb["CAL"], "CALLING_ON")
        loaded_count += l
        skipped_incomplete += s

    log(f"UNIVERSAL YARD COORDINATES LOADED : {loaded_count} signal(s)")

    if skipped_incomplete:
        log(
            f"{skipped_incomplete} signal(s) in the file were skipped - "
            f"missing Menu or Route Init coordinate"
        )

def load_auto_throw_points():

    global loaded_signal_list
    global auto_throw_data

    file = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)

    target_sheet = None

    for sheet in wb.sheetnames:

        if sheet.strip().upper() == "TOC":

            target_sheet = wb[sheet]
            break

    if target_sheet is None:

        messagebox.showerror(
            "ERROR",
            "TOC SHEET NOT FOUND"
        )
        return

    auto_throw_data.clear()
    loaded_signal_list.clear()

    # ===========================================
    # FIND COLUMN NUMBERS FROM HEADER
    # ===========================================

    headers = {}

    for i, cell in enumerate(
            target_sheet.iter_rows(
                min_row=1,
                max_row=1,
                values_only=True
            ).__next__()
    ):

        if cell:
            headers[str(cell).strip().upper()] = i

    if "SIGNAL" not in headers:

        messagebox.showerror(
            "ERROR",
            "SIGNAL COLUMN NOT FOUND"
        )
        return

    if "ROUTE" not in headers:

        messagebox.showerror(
            "ERROR",
            "ROUTE COLUMN NOT FOUND"
        )
        return

    signal_col = headers["SIGNAL"]
    route_col = headers["ROUTE"]
    track_col = headers.get("TRACK", None)

    unique_signals = set()

    # ===========================================
    # READ DATA
    # ===========================================

    for row in target_sheet.iter_rows(
            min_row=2,
            values_only=True
    ):

        if not row:
            continue

        if row[signal_col] is None:
            continue

        signal = str(row[signal_col]).strip().upper()

        route = ""

        if row[route_col] is not None:
            route = str(row[route_col]).strip()

        track = ""

        if track_col is not None:

            if track_col < len(row):

                if row[track_col] is not None:

                    track = str(row[track_col]).strip().upper()

        auto_throw_data.append({

            "signal": signal,

            "route": route,

            "track": track

        })

        unique_signals.add(signal)

        if signal not in loaded_signal_list:
            loaded_signal_list.append(signal)

    log("TOC LOADED")

    log(f"SIGNALS FOUND : {sorted(unique_signals)}")

    log("TOC DATA STORED")

    # ===========================================
    # UPDATE TREEVIEW
    # ===========================================

    tree.delete(*tree.get_children())

    for i, row in enumerate(auto_throw_data, start=1):

        tree.insert(
            "",
            "end",
            values=(
                i,
                row["signal"],
                row["route"],
                row["track"]
            )
        )

    log("NOW CLICK LOAD YARD COORDINATES TO LOAD SIGNAL COORDINATES")


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

# =========================================================
# AVG COLOR
# =========================================================

def is_route_set(point):

    screenshot = pyautogui.screenshot()

    x,y = point

    r,g,b = screenshot.getpixel((x,y))

    return r>150 and g>150
# =========================================================
# CHECK WHETHER ROUTE IS INITIATED
# =========================================================

def route_initiated(signal):

    signal = signal.strip().upper()

    if signal not in signals:
        return False

    aspect = signals[signal].get("ASPECT")

    if not aspect:
        return False

    return is_route_set(aspect)
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

    pause_sleep(1)

    pyautogui.hotkey("ctrl", "b")

    log("CTRL+B Pressed")

    pause_sleep(3)

    found = find_and_click(
        "50051",
        control_type="Text"
    )

    if not found:

        log("50051 Not Found")

        return False

    pause_sleep(1)

    ok = find_and_click(
        "OK",
        control_type="Button"
    )

    if not ok:

        log("OK Button Not Found")

        return False

    pause_sleep(3)

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

    pause_sleep(1)

    transmit_found = find_and_click(
        "TRANSMIT",
        control_type="Button"
    )

    if not transmit_found:

        log("TRANSMIT Not Found")

        return False

    pause_sleep(1)

    cancel_found = find_and_click(
        "CANCEL",
        control_type="Button"
    )

    if not cancel_found:

        log("CANCEL Not Found")

        return False

    log(f"{track} DROP/UP SUCCESS")

    pause_sleep(5)

    return True

# =========================================================
# CREATE ADVANCED REPORT
# =========================================================

def create_report():

    wb = Workbook()

    ws = wb.active

    ws.title = "ROUTE INITIATION REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title_cell = ws["A1"]

    title_cell.value = "ROUTE INITIATION AUTOMATION REPORT"

    title_cell.font = Font(
        bold=True,
        size=18,
        color="FFFFFF"
    )

    title_cell.fill = PatternFill(
        "solid",
        fgColor="1E293B"
    )

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

        "MAIN ROUTE",
        "TRACK",
        "RESULT",
        "DATE & TIME"

    ]

    ws.append(headers)

    header_fill = PatternFill(
        "solid",
        fgColor="2563EB"
    )

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
            color="FFFFFF"
        )

        cell.fill = header_fill

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

            row["MAIN_ROUTE"],
            row["LOCK_ROUTE"],
            row["RESULT"],
            row["DATE & TIME"]

        ])

    # =====================================================
    # ROW FORMATTING
    # =====================================================

    pass_fill = PatternFill(
        "solid",
        fgColor="DCFCE7"
    )

    fail_fill = PatternFill(
        "solid",
        fgColor="FEE2E2"
    )

    for row in ws.iter_rows(min_row=4):

        result_cell = row[2]

        if result_cell.value == "PASS":

            fill = pass_fill

        else:

            fill = fail_fill

        for cell in row:

            cell.fill = fill

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
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

            except:
                pass

        adjusted = max_length + 5

        ws.column_dimensions[column_letter].width = adjusted

    # =====================================================
    # SUMMARY
    # =====================================================

    total = len(report_rows)

    passed = len([
        x for x in report_rows
        if x["RESULT"] == "PASS"
    ])

    failed = len([
        x for x in report_rows
        if x["RESULT"] == "FAIL"
    ])

    start = ws.max_row + 3

    ws[f"A{start}"] = "TOTAL TESTS"
    ws[f"B{start}"] = total

    ws[f"A{start+1}"] = "PASSED"
    ws[f"B{start+1}"] = passed

    ws[f"A{start+2}"] = "FAILED"
    ws[f"B{start+2}"] = failed

    for r in range(start, start + 3):

        ws[f"A{r}"].font = Font(bold=True)
    # =====================================================
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
        color="FFFFFF"
    )

    title.fill = PatternFill(
        "solid",
        fgColor="1E293B"
    )

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    summary_ws.row_dimensions[1].height = 35

    # =====================================================
    # HEADERS
    # =====================================================

    headers = [

        "MAIN ROUTE",
        "TOTAL TESTS",
        "FINAL RESULT"

    ]

    summary_ws.append([])
    summary_ws.append(headers)

    header_fill = PatternFill(
        "solid",
        fgColor="2563EB"
    )

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
            color="FFFFFF"
        )

        cell.fill = header_fill

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

        sig = row["MAIN_ROUTE"]

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

    green_fill = PatternFill(
        "solid",
        fgColor="DCFCE7"
    )

    red_fill = PatternFill(
        "solid",
        fgColor="FEE2E2"
    )

    for sig, data in signal_summary.items():

        total = data["total"]

        failed = data["failed"]

        if failed == 0:

            final_result = "ALL TEST CASES PASSED"

            fill = green_fill

        else:

            failed_routes = []

            for r in report_rows:

                if (
                        r["MAIN_ROUTE"] == sig
                        and r["RESULT"] == "FAIL"
                ):
                    failed_routes.append(
                        f'{r["LOCK_ROUTE"]} NOT PASSED'
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
            cell.fill = fill

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
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

            except:
                pass

        adjusted = max_length + 5

        summary_ws.column_dimensions[
            column_letter
        ].width = adjusted

    # =====================================================
    # SAVE
    # =====================================================

    wb.save(REPORT_FILE)

    log(f"REPORT SAVED : {REPORT_FILE}")
def cancel_and_release(signal):

    if signal not in signals:
        return False

    # -----------------------------
    # SIGNAL CANCEL
    # -----------------------------
    pyautogui.click(
        signals[signal]["menu"][0],
        signals[signal]["menu"][1]
    )

    pause_sleep(1)

    if not find_and_click("Signal Cancel"):
        log("Signal Cancel Not Found")
        return False

    log(f"{signal} Signal Cancel Done")

    pause_sleep(2)

    # -----------------------------
    # ROUTE RELEASE
    # -----------------------------
    pyautogui.click(
        signals[signal]["menu"][0],
        signals[signal]["menu"][1]
    )

    pause_sleep(1)

    if not find_and_click("Route Release"):
        log("Route Release Not Found")
        return False

    log(f"{signal} Route Release Done")

    pause_sleep(8)

    return True
def run_auto_throw_engine():

    global running
    global report_rows

    report_rows.clear()

    log("STARTED")

    for row in auto_throw_data:

        if not running:
            break

        signal = row["signal"].strip().upper()

        route = row["route"].strip()

        if signal not in signals:

            log(f"{signal} NOT RECORDED")

            continue

        track = row.get("track", "").strip().upper()

        # -------------------------------------------------
        # CALLING ON : TRACK DOWN FIRST
        # -------------------------------------------------
        if signals[signal]["type"] == "CALLING_ON":

            if track:

                log(f"{signal} : TRACK DOWN -> {track}")

                drop_track(track)

                pause_sleep(2)

            else:

                log(f"{signal} : NO TRACK FOUND")

        # -------------------------------------------------
        # NOW SET ROUTE
        # -------------------------------------------------

        pyautogui.click(
            signals[signal]["menu"][0],
            signals[signal]["menu"][1]
        )

        pause_sleep(1)

        find_and_click(route)

        pause_sleep(7)

        if route_initiated(signal):

            result = "PASS"

            log(f"{signal} Route Initiated Successfully")

            cancel_and_release(signal)

            # ----------------------------------------
            # TRACK UP AFTER ROUTE RELEASE
            # ----------------------------------------
            if signals[signal]["type"] == "CALLING_ON":

                if track:
                    log(f"{signal} : TRACK UP -> {track}")

                    drop_track(track)

                    pause_sleep(2)

        else:

            result = "FAIL"

        report_rows.append({

            "MAIN_ROUTE":signal,

            "LOCK_ROUTE":route,

            "RESULT":result,

            "DATE & TIME":datetime.now().strftime(
                "%d-%m-%Y %H:%M:%S"
            )

        })

        log(f"{signal} -> {result}")

    create_report()
    status_label.config(
        text="COMPLETED",
        fg="#2563eb"
    )

    log("AUTOMATION COMPLETED")

    try:
        os.startfile(REPORT_FILE)
    except:
        pass
    running=False
# =========================================================
# START
# =========================================================

def start_automation():

    global running

    if not signals:

        log("LOAD CONFIG FIRST")
        return

    if not auto_throw_data:

        log("LOAD TOC FIRST")
        return

    running = True

    pause_event.set()

    status_label.config(
        text="RUNNING",
        fg="#16a34a"
    )

    log("AUTOMATION STARTING IN 2 SECONDS")

    root.iconify()

    def delayed_start():

        pause_sleep(2)

        run_auto_throw_engine()

    threading.Thread(
        target=delayed_start,
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
    "ROUTE INITIATION TESTING AUTOMATION"
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

    text="ROUTE INITIATION AUTOMATION SYSTEM",

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
# OTHER BUTTONS
# =========================================================

create_button(
    "ROUTE INTIATION TEST",
    load_auto_throw_points,
    "#2563eb"
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

    text="ROUTE INITIATION DETAILS",

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

tree.heading("LOCK_ROUTE", text="TRACK CIRCUITS")

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

    text="SPACE KEY = CAPTURE COORDINATES | ROUTE INTIATION AUTOMATION SYSTEM",

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
