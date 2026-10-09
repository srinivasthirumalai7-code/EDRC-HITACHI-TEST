# =========================================================
# AUTO THROW OF POINTS
# =========================================================
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

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "AUTO_THROW_OF_POINTS.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

auto_throw_data = []
loaded_signal_list = []
loaded_point_list = []
points = {}
running = False

capture_queue = []

capture_index = 0

capture_mode = None

current_signal = None

current_step = 0

last_capture_time = 0

capture_waiting = False

captured_point = None

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
# =========================================================
# SAVE CONFIG
# =========================================================

def load_auto_throw_points():
    global loaded_signal_list
    global loaded_point_list
    global auto_throw_data

    file = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)

    # =====================================================
    # Look for a sheet named "TOC" - same convention as every
    # other program. Falls back to the active sheet so a plain
    # single-tab workbook still works.
    # =====================================================

    sheet_map = {name.strip().upper(): name for name in wb.sheetnames}

    if "TOC" in sheet_map:
        target_sheet = wb[sheet_map["TOC"]]
    else:
        target_sheet = wb.active

    # =====================================================
    # HEADER-DRIVEN COLUMN LOOKUP
    # Finds Signal / Route / Points / Track by header text in row 1,
    # wherever they sit and in any order.
    # =====================================================

    header_row = next(target_sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)

    if not header_row:
        messagebox.showerror("ERROR", "Sheet has no header row")
        return

    header_index = {}
    for idx, val in enumerate(header_row):
        if val is not None:
            header_index[str(val).strip().upper()] = idx

    def find_col(*aliases):
        for a in aliases:
            if a in header_index:
                return header_index[a]
        return None

    signal_col = find_col("SIGNAL")
    route_col = find_col("ROUTE")
    points_col = find_col("POINTS", "POINT")
    track_col = find_col("TRACK")

    if signal_col is None or route_col is None:
        messagebox.showerror(
            "ERROR",
            f"Could not find 'Signal' and/or 'Route' column headers.\n"
            f"Headers found: {list(header_index.keys())}"
        )
        return

    def cell(row, idx):
        return row[idx] if idx is not None and len(row) > idx else None

    auto_throw_data.clear()

    unique_signals = set()

    unique_points = set()

    for row in target_sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row:
            continue

        if cell(row, signal_col) is None:
            continue

        signal = str(cell(row, signal_col)).strip().upper()

        route = str(cell(row, route_col) or "").strip()

        point_text = ""

        track = ""

        if cell(row, points_col):
            point_text = str(cell(row, points_col)).strip().upper()

        if cell(row, track_col):
            track = str(cell(row, track_col)).strip().upper()

        auto_throw_data.append({

            "signal": signal,

            "route": route,

            "points": point_text,

            "track": track

        })

        unique_signals.add(signal)

        if point_text:

            for p in point_text.split(","):

                p = p.strip()

                if not p:
                    continue

                point_no = ""

                for ch in p:

                    if ch.isdigit():

                        point_no += ch

                    else:

                        break

                if point_no:

                    unique_points.add(point_no)

    log("AUTO THROW OF POINTS LOADED")

    log(f"SIGNALS FOUND : {sorted(unique_signals)}")

    log(f"POINTS FOUND : {sorted(unique_points)}")
    loaded_signal_list = []

    for row in auto_throw_data:

        signal = row["signal"]

        if signal not in loaded_signal_list:
            loaded_signal_list.append(signal)
    loaded_point_list = sorted(unique_points)

    log("AUTO THROW DATA STORED")
    log("NOW CLICK LOAD YARD COORDINATES TO LOAD SIGNAL/POINT COORDINATES")
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

    except:
        pass

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
# SHUNT STATUS
# =========================================================

def shunt_route_set(signal):

    try:

        x, y = signals[signal]["ASPECT"]

        img = pyautogui.screenshot(
            region=(x - 10, y - 10, 20, 20)
        )

        gray = img.convert("L")

        bright = []

        for py in range(gray.height):

            for px in range(gray.width):

                if gray.getpixel((px, py)) > 180:

                    bright.append((px, py))

        log(f"BRIGHT PIXELS = {len(bright)}")

        if len(bright) < 10:
            return 0

        top = [p for p in bright if p[1] < 10]

        bottom = [p for p in bright if p[1] >= 10]

        if not top or not bottom:
            return 0

        top_x = sum(p[0] for p in top) / len(top)

        bottom_x = sum(p[0] for p in bottom) / len(bottom)

        diff = abs(top_x - bottom_x)

        log(f"TOP X = {top_x}")
        log(f"BOTTOM X = {bottom_x}")
        log(f"DIFF = {diff}")
        log(f"SCREENSHOT AREA = {x},{y}")

        img.save("shunt_test.png")

        return diff

    except Exception as e:

        log(f"SHUNT CHECK ERROR : {e}")

        return 0

# =========================================================
# FIND AND CLICK UI
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

    ws.title = "AUTO THROW REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title_cell = ws["A1"]

    title_cell.value = "AUTO THROW AUTOMATION REPORT"

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
        "LOCK ROUTE",
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
def run_auto_throw_engine():

    global running
    global report_rows

    report_rows.clear()


    log("AUTO THROW AUTOMATION STARTED")

    for row in auto_throw_data:

        if not running:
            break

        signal_name = str(row["signal"]).strip()
        route_name = str(row["route"]).strip()
        points_text = str(row["points"]).strip()
        track_name = str(row.get("track", "")).strip()

        log(f"PROCESSING : {signal_name} -> {route_name}")
        before_diff = 0

        if "SH" in signal_name.upper():
            before_diff = shunt_route_set(signal_name)

            log(f"BEFORE DIFF = {before_diff}")
        try:
            before_diff = 0

            if "SH" in signal_name.upper():
                before_diff = shunt_route_set(signal_name)

                log(f"BEFORE DIFF = {before_diff}")
            # -----------------------------------
            # SET POINTS
            # -----------------------------------

            if points_text:

                point_list = [
                    x.strip()
                    for x in points_text.split(",")
                    if x.strip()
                ]

                for point_entry in point_list:

                    point_no = ''.join(
                        c for c in point_entry
                        if c.isdigit()
                    )

                    state = point_entry[-1].upper()

                    if point_no not in points:
                        log(f"POINT {point_no} NOT RECORDED")
                        continue

                    pyautogui.click(
                        points[point_no][0],
                        points[point_no][1]
                    )

                    pause_sleep(3)

                    if state == "N":
                        find_and_click("Reverse")

                    elif state == "R":
                        find_and_click("Normal")

                    pause_sleep(5)

            # -----------------------------------
            # CALLING ON TRACK DOWN
            # -----------------------------------

            is_calling_on = (
                    signal_name.endswith("C")
                    and not signal_name.endswith("SH")
            )

            if is_calling_on and track_name:

                log(f"DROPPING TRACK : {track_name}")

                drop_track(track_name)

                pause_sleep(2)

            # -----------------------------------
            # SET ROUTE
            # -----------------------------------

            if signal_name not in signals:

                log(f"{signal_name} NOT RECORDED")
                continue

            pyautogui.click(
                signals[signal_name]["menu"][0],
                signals[signal_name]["menu"][1]
            )

            pause_sleep(1)

            find_and_click(route_name)

            log("ROUTE SET - WAITING 8 SECONDS FOR SIGNAL TO CLEAR")

            for i in range(8):

                if not running:
                    break

                pause_event.wait()

                pause_sleep(1)

                log(f"WAITING {i + 1}/8")

            # -----------------------------------
            # CHECK RESULT
            # -----------------------------------

            passed = False
            # CALLING ON
            if is_calling_on:

                try:

                    yellow = signals[signal_name]["YELLOW"]

                    if is_yellow(yellow):
                        passed = True

                except:
                    pass
            # SHUNT
            elif "SH" in signal_name.upper():

                after_diff = shunt_route_set(signal_name)

                log(f"AFTER DIFF = {after_diff}")

                log(f"BEFORE DIFF = {before_diff}")

                # Route should change the indication

                if abs(after_diff - before_diff) > 0.3:
                    passed = True

            # MAIN
            else:

                try:

                    if is_signal_changed(signal_name):
                        passed = True

                except:
                    pass

            result = "PASS" if passed else "FAIL"

            report_rows.append(
                {
                    "MAIN_ROUTE": signal_name,
                    "LOCK_ROUTE": route_name,
                    "RESULT": result,
                    "DATE & TIME": datetime.now().strftime(
                        "%d-%m-%Y %H:%M:%S"
                    )
                }
            )

            log(
                f"{signal_name} {route_name} -> {result}"
            )

            # -----------------------------------
            # CANCEL + ROUTE RELEASE
            # -----------------------------------

            if passed:

                log(f"{signal_name} PASS -> SIGNAL CANCEL")

                pyautogui.click(
                    signals[signal_name]["menu"][0],
                    signals[signal_name]["menu"][1]
                )

                pause_sleep(1)

                find_and_click("Signal Cancel")

                pause_sleep(2)

                log(f"{signal_name} ROUTE RELEASE")

                pyautogui.click(
                    signals[signal_name]["menu"][0],
                    signals[signal_name]["menu"][1]
                )

                pause_sleep(1)

                find_and_click("Route Release")

                pause_sleep(2)

                log("WAITING 7 SECONDS AFTER ROUTE RELEASE")

                for i in range(7):

                    if not running:
                        break

                    pause_event.wait()

                    pause_sleep(1)

                    log(f"POST RELEASE WAIT {i + 1}/7")

            # -----------------------------------
            # TRACK UP
            # -----------------------------------

            if is_calling_on and track_name:

                log(f"RESTORING TRACK : {track_name}")

                drop_track(track_name)

                pause_sleep(2)

        except Exception as e:

            log(
                f"ERROR IN {signal_name} : {e}"
            )
        if not running:
            break

    log("AUTO THROW AUTOMATION COMPLETED")
    if report_rows:

        create_report()

        try:
            os.startfile(REPORT_FILE)
        except:
            pass

    log("REPORT GENERATED")
    running = False
# =========================================================
# START
# =========================================================

def start_automation():

    global running

    if not signals:

        log("LOAD CONFIG FIRST")
        return

    if not auto_throw_data:

        log("LOAD AUTO THROW OF POINTS FIRST")
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


def load_universal_coordinates():
    """Loads the multi-sheet universal yard coordinate file (MAIN / SHUNT
    / CAL / POINT) produced by the Universal Yard Coordinate Recorder.
    Columns found by header name, wherever they sit.

    - MAIN: Menu + RED/YELLOW/DOUBLE_YELLOW/GREEN + Aspects
    - SHUNT: Menu + Indicator (used here as the single ASPECT point)
    - CAL:  Menu + Yellow
    - POINT: Menu coordinate only (this program clicks the point then
      uses the Normal/Reverse menu items - it doesn't need the
      Normal/Reverse/Free aspect coordinates the Signal Clearance
      program records)"""

    global signals, points

    if not auto_throw_data:
        messagebox.showerror("ERROR", "LOAD AUTO THROW OF POINTS FIRST")
        return

    file = filedialog.askopenfilename(
        title="Select Universal Yard Coordinates",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not file:
        return

    wb = load_workbook(file)

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

    def header_index_of(ws):
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            return {}
        idx = {}
        for i, val in enumerate(header_row):
            if val is not None:
                idx[str(val).strip().upper()] = i
        return idx

    def cell(row, idx):
        return row[idx] if idx is not None and len(row) > idx else None

    loaded_main = loaded_shunt = loaded_cal = loaded_points = 0

    # ---- MAIN ----
    if "MAIN" in wb.sheetnames:
        ws = wb["MAIN"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                signals[sig] = {
                    "type": "MAIN",
                    "aspects": parse_int(cell(row, hi.get("ASPECTS"))) or 3,
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "RED": parse_point(cell(row, hi.get("RED_X")), cell(row, hi.get("RED_Y"))),
                    "YELLOW": parse_point(cell(row, hi.get("YELLOW_X")), cell(row, hi.get("YELLOW_Y"))),
                    "DOUBLE_YELLOW": parse_point(cell(row, hi.get("DOUBLE_YELLOW_X")), cell(row, hi.get("DOUBLE_YELLOW_Y"))),
                    "GREEN": parse_point(cell(row, hi.get("GREEN_X")), cell(row, hi.get("GREEN_Y")))
                }
                loaded_main += 1

    # ---- SHUNT ----
    if "SHUNT" in wb.sheetnames:
        ws = wb["SHUNT"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "ASPECT": parse_point(cell(row, hi.get("INDICATOR_X")), cell(row, hi.get("INDICATOR_Y")))
                }
                loaded_shunt += 1

    # ---- CALLING-ON ----
    if "CAL" in wb.sheetnames:
        ws = wb["CAL"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                signals[sig] = {
                    "type": "CALLING_ON",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "YELLOW": parse_point(cell(row, hi.get("YELLOW_X")), cell(row, hi.get("YELLOW_Y")))
                }
                loaded_cal += 1

    # ---- POINT ----
    if "POINT" in wb.sheetnames:
        ws = wb["POINT"]
        hi = header_index_of(ws)
        name_col = hi.get("POINT")
        if name_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, name_col) is None:
                    continue
                name = str(cell(row, name_col)).strip().upper()
                menu = parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y")))
                if menu:
                    points[name] = menu
                    loaded_points += 1

    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_shunt} SHUNT, {loaded_cal} CALLING-ON, {loaded_points} POINT(S)"
    )

# =========================================================
# GUI
# =========================================================

root = tk.Tk()

root.title(
    "AUTO THROW OF POINTS TESTING AUTOMATION"
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

    text="AUTO THROW OF POINTS AUTOMATION SYSTEM",

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
    "AUTO THROW OF POINTS",
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

    text="AUTO THROW OF POINTS DETAILS",

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
