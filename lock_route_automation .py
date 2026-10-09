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


import uiautomation as auto

from pywinauto import Desktop

from openpyxl import Workbook
from openpyxl import load_workbook

import numpy as np
import cv2

orb = cv2.ORB_create(nfeatures=500)

from datetime import datetime

# =========================================================
# FILES
# =========================================================

CONFIG_FILE = "SIGNAL_CONFIG.xlsx"

REPORT_FILE = "LOCK_ROUTE_REPORT.xlsx"

# =========================================================
# GLOBALS
# =========================================================

signals = {}

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
CONFIG_FILE = ""

config_file_path = ""
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
# SAVE CONFIG
# =========================================================

def load_lock_routes():

    global lock_routes_data

    file = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx")]
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

    ws = wb[sheet_map["TOC"]] if "TOC" in sheet_map else wb.active

    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)

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
    lock_route_col = find_col("LOCK_ROUTE", "LOCK ROUTE")
    track_col = find_col("TRACK_LOCK_ROUTE", "TRACK LOCK ROUTE", "TRACK")

    if signal_col is None or route_col is None or lock_route_col is None:
        messagebox.showerror(
            "ERROR",
            f"Could not find 'Signal'/'Route'/'Lock_Route' column headers.\n"
            f"Headers found: {list(header_index.keys())}"
        )
        return

    def cell(row, idx):
        return row[idx] if idx is not None and len(row) > idx else None

    lock_routes_data = []

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row:
            continue

        if cell(row, signal_col) is None:
            continue

        signal = str(cell(row, signal_col)).replace(".0", "").strip().upper()

        route = str(cell(row, route_col) or "").replace(".0", "").strip().upper()

        lock_route = str(cell(row, lock_route_col) or "").replace(".0", "").strip().upper()

        track_data = ""

        if cell(row, track_col):
            track_data = str(cell(row, track_col)).strip().upper()

        lock_routes_data.append({

            "signal": signal,

            "route": route,

            "lock_route": lock_route,

            "track_data": track_data

        })

    refresh_table()

    log("LOCK ROUTES LOADED")


def load_universal_coordinates():
    """Loads the multi-sheet universal yard coordinate file (MAIN / SHUNT
    / CAL). Also loads each signal's Route Initiation Indicator
    (RouteInit_X/Y), which is now the primary check used to decide
    LOCKED vs NOT LOCKED for every signal type. SHUNT uses the latest
    single-indicator + reference-snapshot technique (ORB feature
    matching), matching the Universal Yard Coordinate Recorder."""

    global signals

    if not lock_routes_data:
        messagebox.showwarning("WARNING", "Please load TOC/RCC file first.")
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

    loaded_main = loaded_shunt = loaded_cal = 0

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
                    "GREEN": parse_point(cell(row, hi.get("GREEN_X")), cell(row, hi.get("GREEN_Y"))),
                    "ROUTE_INIT": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                }
                loaded_main += 1

    # ---- SHUNT: latest technique - ONE indicator point + reference
    # snapshot, compared via ORB feature matching (same as the Universal
    # Yard Coordinate Recorder / Signal Clearance program), instead of
    # the old three-point C1/C2/C3 brightness check ----
    if "SHUNT" in wb.sheetnames:
        ws = wb["SHUNT"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                snap_val = cell(row, hi.get("SNAPSHOT_PATH"))
                snapshot_path = str(snap_val).strip() if snap_val else None
                signals[sig] = {
                    "type": "SHUNT",
                    "menu": parse_point(cell(row, hi.get("MENU_X")), cell(row, hi.get("MENU_Y"))),
                    "indicator": parse_point(cell(row, hi.get("INDICATOR_X")), cell(row, hi.get("INDICATOR_Y"))),
                    "snapshot": snapshot_path,
                    "ROUTE_INIT": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                }
                loaded_shunt += 1

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
                    "YELLOW": parse_point(cell(row, hi.get("YELLOW_X")), cell(row, hi.get("YELLOW_Y"))),
                    "ROUTE_INIT": parse_point(cell(row, hi.get("ROUTEINIT_X")), cell(row, hi.get("ROUTEINIT_Y")))
                }
                loaded_cal += 1

    log(
        f"UNIVERSAL YARD COORDINATES LOADED : "
        f"{loaded_main} MAIN, {loaded_shunt} SHUNT (single indicator + snapshot), {loaded_cal} CALLING-ON"
    )

# =========================================================
# TABLE
# =========================================================

def refresh_table():

    tree.delete(*tree.get_children())

    for i, row in enumerate(
        lock_routes_data,
        start=1
    ):

        tree.insert(
            "",
            "end",
            values=(

                i,

                row["signal"],

                row["route"],

                row["lock_route"]

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


def wait_for_stable_yellow(point, max_wait=8.0, poll_interval=0.3, stable_reads=3):
    """Polls a coordinate's color, waiting for it to settle into a
    steady YELLOW (an indicator can blink for a moment before it locks
    in) or until max_wait seconds pass, whichever happens first -
    instead of a flat sleep. Returns True if it ended up steadily
    yellow, False otherwise."""

    if point is None:
        return False

    start = time.time()
    consecutive = 0

    while time.time() - start < max_wait:

        pause_event.wait()

        if is_yellow(point):
            consecutive += 1
        else:
            consecutive = 0

        if consecutive >= stable_reads:
            return True

        time.sleep(poll_interval)

    # Timed out - go with whatever the last reading was
    return is_yellow(point)

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

def shunt_compare_by_features(signal):
    """ORB feature match between the shunt indicator right now vs. its
    reference snapshot - same technique the Universal Yard Coordinate
    Recorder and Signal Clearance program use. Returns a match score
    (0-100, higher = more similar to the reference = less change), or
    None if this signal has no indicator/snapshot recorded."""

    info = signals.get(signal, {})
    point = info.get("indicator")
    snapshot_path = info.get("snapshot")

    if not point or not snapshot_path or not os.path.exists(snapshot_path):
        return None

    try:

        initial_img = cv2.imread(snapshot_path)

        if initial_img is None:
            return None

        screenshot = pyautogui.screenshot()

        x, y = point

        current_pil = screenshot.crop((x - 50, y - 50, x + 50, y + 50))

        current_img = cv2.cvtColor(np.array(current_pil), cv2.COLOR_RGB2BGR)

        initial_gray = cv2.cvtColor(initial_img, cv2.COLOR_BGR2GRAY)
        current_gray = cv2.cvtColor(current_img, cv2.COLOR_BGR2GRAY)

        kp1, des1 = orb.detectAndCompute(initial_gray, None)
        kp2, des2 = orb.detectAndCompute(current_gray, None)

        if des1 is None or des2 is None:
            return None

        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)

        matches = bf.match(des1, des2)

        if not matches:
            return 0

        good_matches = [m for m in matches if m.distance < 50]

        match_score = (
            (len(good_matches) / max(len(kp1), len(kp2))) * 100
            if max(len(kp1), len(kp2)) > 0 else 0
        )

        log(f"{signal} SHUNT Feature Match: {match_score:.2f}%")

        return match_score

    except Exception as e:

        log(f"{signal} Shunt Feature Compare Error: {e}")

        return None

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


def resolve_signal(signal_name, signal_type):
    """detect_signal_type() only distinguishes 'C' (calling-on) from
    everything else (defaulting to MAIN) - it has no idea a plain number
    like '9' in Lock_Route text ('9-A') actually means shunt signal
    '9SH', since the coordinate file (and every other program) records
    shunt signals WITH their 'SH' suffix.

    If the plain name isn't a known signal but name+'SH' is a known
    SHUNT signal, use that instead. The menu click still uses the
    original lock-route text unchanged - only which coordinates get
    looked up (and reported) changes."""

    if signal_name in signals:
        return signal_name, signal_type

    shunt_name = signal_name + "SH"

    if shunt_name in signals and signals[shunt_name].get("type") == "SHUNT":
        return shunt_name, "SHUNT"

    return signal_name, signal_type

# =========================================================
# GET TRACK FOR LOCK ROUTE
# =========================================================

def get_track_for_route(lock_route, track_data, is_main_calling_on=False):

    routes = [

        x.strip().upper()

        for x in lock_route.split(",")

    ]

    tracks = [

        x.strip().upper()

        for x in track_data.split(",")

    ]

    mapping = {}

    # =====================================================
    # FIX: Skip first track if main signal is CALLING_ON
    # =====================================================
    # If main is CALLING_ON:
    #   tracks[0] = main track (drop before set, up after release)
    #   tracks[1+] = lock route tracks
    # If main is MAIN:
    #   tracks[0+] = lock route tracks
    # =====================================================

    track_index = 1 if is_main_calling_on else 0

    for route in routes:

        signal_name, signal_type = detect_signal_type(route)

        # ONLY CALLING ON LOCK ROUTES GET TRACK

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

    ws.title = "LOCK ROUTE REPORT"

    # =====================================================
    # TITLE
    # =====================================================

    ws.merge_cells("A1:D1")

    title_cell = ws["A1"]

    title_cell.value = "LOCK ROUTE AUTOMATION REPORT"

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
# =========================================================
# RUN ENGINE
# =========================================================

def run_engine():

    global running

    root.iconify()

    time.sleep(2)

    for row in lock_routes_data:
        pause_event.wait()

        if not running:
            return

        main_signal = row["signal"]

        log(f"EXCEL MAIN SIGNAL = {main_signal}")

        log(f"AVAILABLE CONFIG SIGNALS = {list(signals.keys())}")

        main_route = row["route"]

        lock_routes = row["lock_route"]

        track_data = row["track_data"]

        # =================================================
        # DETECT CALLING_ON AND DROP MAIN TRACK
        # =================================================

        main_signal_clean = str(main_signal).replace(".0", "").strip().upper()
        
        is_calling_on = main_signal_clean.endswith("C")

        # =================================================
        # GET TRACK MAPPING (after detecting CALLING_ON)
        # =================================================

        track_mapping = get_track_for_route(
            lock_routes,
            track_data,
            is_calling_on
        )
        
        main_track = None
        
        if is_calling_on and track_data:
            main_track_list = [t.strip() for t in track_data.replace(",", " ").split()]
            if main_track_list:
                main_track = main_track_list[0]
                log(f"MAIN SIGNAL IS CALLING_ON - DROPPING TRACK: {main_track}")
                drop_track(main_track)
                time.sleep(3)

        # =================================================
        # SET MAIN ROUTE
        # =================================================

        log(
            f"SETTING {main_signal_clean} {main_route}"
        )

        main_signal = main_signal_clean

        if main_signal not in signals:
            log(f"{main_signal} NOT FOUND IN CONFIG")

            continue

        click(signals[main_signal]["menu"])

        time.sleep(1)

        click_menu_item(
            main_route.replace("-", "_")
        )

        # Close the dropdown - it can visually sit on top of the aspect
        # indicator and corrupt the color read otherwise.
        pyautogui.press("esc")

        time.sleep(1)

        main_route_init_point = signals[main_signal].get("ROUTE_INIT")

        if main_route_init_point:

            log(f"{main_signal} : waiting up to 8s for its Route Initiation Indicator to settle...")

            main_route_initiated = wait_for_stable_yellow(main_route_init_point, max_wait=8.0)

        else:

            time.sleep(4)

            main_route_initiated = is_signal_changed(main_signal)

        if main_route_initiated:
            log(f"{main_signal} Main Route Initiated")
        else:
            log(f"{main_signal} Main Route did NOT initiate")

        # =================================================
        # LOCK ROUTES
        # =================================================

        routes = [

            x.strip()

            for x in lock_routes.split(",")

        ]

        for lr in routes:
            pause_event.wait()

            if not running:
                return

            signal_name, detected_type = detect_signal_type(lr)

            signal_name, detected_type = resolve_signal(signal_name, detected_type)

            if signal_name not in signals:
                log(f"{signal_name} NOT FOUND")

                continue

            sig_info = signals[signal_name]

            sig_type = detected_type

            log(
                f"{signal_name} -> {sig_type}"
            )

            # =============================================
            # DROP TRACK (CALLING-ON ONLY, BEFORE SETTING ROUTE)
            # =============================================
            current_track = track_mapping.get(lr)

            if sig_type == "CALLING_ON" and current_track:

                log(f"DROPPING TRACK : {current_track}")

                drop_track(current_track)

                time.sleep(3)

            # =============================================
            # SET THE LOCK ROUTE (same click sequence for all types)
            # =============================================

            click(sig_info["menu"])

            time.sleep(1)

            click_menu_item(
                lr.replace("-", "_")
            )

            # Close the dropdown - it can visually sit on top of the
            # aspect indicator and corrupt the color read otherwise.
            pyautogui.press("esc")

            time.sleep(1)

            # =============================================
            # CHECK RESULT
            # Primary check: Route Initiation Indicator, waited on until
            # it settles (blinking can happen briefly before it locks
            # in) or 8 seconds pass, whichever is first. If this signal
            # has no Route Init point recorded, fall back to the old
            # per-type check instead.
            # =============================================

            route_init_point = sig_info.get("ROUTE_INIT")

            if route_init_point:

                log(f"{signal_name} : waiting up to 8s for Route Initiation Indicator to settle...")

                route_initiated = wait_for_stable_yellow(route_init_point, max_wait=8.0)

                if route_initiated:

                    result = "FAIL"
                    status = "NOT LOCKED"
                    need_cancel = True

                    log(f"{signal_name} Route Initiation Indicator settled YELLOW -> route accepted -> NOT LOCKED")

                else:

                    result = "PASS"
                    status = "LOCKED"
                    need_cancel = False

                    log(f"{signal_name} Route Initiation Indicator did not settle yellow -> LOCKED")

            else:

                log(f"{signal_name} has no Route Init coordinate recorded - using fallback check")

                time.sleep(5)

                if sig_type == "CALLING_ON":

                    yellow = is_yellow(sig_info.get("YELLOW"))

                    if yellow:
                        result = "FAIL"; status = "NOT LOCKED"; need_cancel = True
                    else:
                        result = "PASS"; status = "LOCKED"; need_cancel = False

                elif sig_type == "MAIN":

                    changed = is_signal_changed(signal_name)

                    if changed:
                        result = "FAIL"; status = "NOT LOCKED"; need_cancel = True
                    else:
                        result = "PASS"; status = "LOCKED"; need_cancel = False

                elif sig_type == "SHUNT":

                    match_score = shunt_compare_by_features(signal_name)

                    if match_score is None:
                        result = "FAIL"; status = "UNKNOWN (no snapshot)"; need_cancel = True
                    elif match_score < 70:
                        result = "FAIL"; status = "NOT LOCKED"; need_cancel = True
                    else:
                        result = "PASS"; status = "LOCKED"; need_cancel = False

                else:

                    result = "FAIL"; status = "UNKNOWN TYPE"; need_cancel = False

            # =============================================
            # DO CANCEL
            # =============================================

            if need_cancel:

                log(f"{signal_name} SIGNAL SET -> CANCELLING")

                # SIGNAL CANCEL
                click(sig_info["menu"])

                time.sleep(1)

                click_menu_item("Signal Cancel")

                log(f"{signal_name} SIGNAL CANCEL DONE")

                time.sleep(2)

                # ROUTE RELEASE
                click(sig_info["menu"])

                time.sleep(1)

                click_menu_item("Route Release")

                log(f"{signal_name} ROUTE RELEASE DONE")

                time.sleep(7)
            # =============================================
            # TRACK UP AGAIN
            # =============================================

            if sig_type == "CALLING_ON":

                # =========================================
                # CLOSE MENU BEFORE TRACK UP
                # =========================================

                pyautogui.press("esc")

                time.sleep(1)

                # =========================================
                # TRACK UP
                # =========================================

                if current_track:
                    log(f"UP TRACK : {current_track}")

                    drop_track(current_track)

                    time.sleep(3)

                # =========================================
                # EXTRA ESC AFTER TRACK UP
                # =========================================

                pyautogui.press("esc")

                time.sleep(1)


            else:

                log(f"{signal_name} NOT SET -> PRESSING ESC")

                pyautogui.press("esc")

                time.sleep(1)
            report_rows.append({

                "MAIN_SIGNAL": main_signal,

                "MAIN_ROUTE": main_route,

                "LOCK_ROUTE": lr,

                "TYPE": sig_type,

                "STATUS": status,

                "RESULT": result,

                "DATE & TIME": datetime.now().strftime("%d-%m-%Y %H:%M:%S")

            })

            log(
                f"{lr} -> {result}"
            )
        # =============================================
        # ALL LOCK ROUTES COMPLETED
        # =============================================

        log(
            f"{main_signal} LOCK ROUTES COMPLETED"
        )

        if main_route_initiated:

            # =============================================
            # SIGNAL CANCEL
            # =============================================

            click(signals[main_signal]["menu"])

            time.sleep(1)

            click_menu_item("Signal Cancel")

            log(f"{main_signal} SIGNAL CANCEL DONE")

            time.sleep(2)

            # =============================================
            # OPEN MENU AGAIN
            # =============================================

            click(signals[main_signal]["menu"])

            time.sleep(1)

            # =============================================
            # ROUTE RELEASE
            # =============================================

            click_menu_item("Route Release")

            log(f"{main_signal} ROUTE RELEASE DONE")

            time.sleep(6)

        else:

            log(f"{main_signal} Main Route was never initiated - skipping Signal Cancel/Route Release")

        # =============================================
        # UP MAIN TRACK IF CALLING_ON (unconditional - if it was
        # dropped, it needs to come back up regardless)
        # =============================================

        if is_calling_on and main_track:
            log(f"UPPING MAIN TRACK: {main_track}")
            drop_track(main_track)
            time.sleep(3)

    running = False

    status_label.config(
        text="COMPLETED",
        fg="#16a34a"
    )

    create_report()

    root.deiconify()

    log("AUTOMATION COMPLETED")

    try:
        os.startfile(REPORT_FILE)

    except:
        pass

# =========================================================
# START
# =========================================================

def start_automation():

    global running

    if not signals:

        log("LOAD CONFIG FIRST")

        return

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
    "LOCK ROUTES TESTING AUTOMATION"
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

    text="LOCK ROUTE AUTOMATION SYSTEM",

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
    "LOCK ROUTES(TOC/RCC)",
    load_lock_routes,
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

    text="LOCK ROUTE DETAILS",

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
