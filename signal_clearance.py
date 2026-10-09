# =========================================================
# SIGNAL CLEARANCE AUTOMATION SYSTEM (MERGED)
# Combines: MAIN SIGNAL CLEARANCE + CALLING ON CLEARANCE + SHUNT CLEARANCE
#
# TESTING ORDER:
#   For each MAIN signal (in TOC order):
#       1. Run MAIN signal clearance (all its lock routes)
#       2. Run its CALLING-ON clearance, if one exists (e.g. Signal "1" -> "1C")
#   After ALL main + calling-on signals are done:
#       3. Run SHUNT signal clearance for ALL shunt signals (ORB feature detection)
#
# COORDINATE CAPTURE: a SINGLE recording pass and a SINGLE save/load file
# cover MAIN + CALLING-ON + SHUNT together - there is no separate capture
# flow per signal type. Each signal's type is auto-detected from its name
# (SH suffix -> SHUNT, C suffix -> CALLING-ON, otherwise -> MAIN), so the
# operator just walks the panel once, signal by signal.
# =========================================================

import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import win32api
import win32con

import time
import threading
import os

from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

import uiautomation as auto
import pyautogui

from pywinauto import Desktop

from PIL import Image, ImageTk

import numpy as np
import cv2

# =========================================================
# FOLDERS
# =========================================================
SNAPSHOTS_FOLDER = "snapshots"
DEBUG_FOLDER = "debug_captures"

if not os.path.exists(SNAPSHOTS_FOLDER):
    os.makedirs(SNAPSHOTS_FOLDER)

if not os.path.exists(DEBUG_FOLDER):
    os.makedirs(DEBUG_FOLDER)

REPORT_FILE = None

# =========================================================
# ORB DETECTOR (SHUNT)
# =========================================================
orb = cv2.ORB_create(nfeatures=500)

# =========================================================
# GLOBALS - MAIN SIGNAL
# =========================================================
main_data = {}
main_signal_routes_map = {}
main_signal_order = []
main_signal_aspects = {}

# =========================================================
# GLOBALS - CALLING ON
# =========================================================
cal_excel_data = []
cal_signal_data = {}
cal_signal_order = []

# =========================================================
# GLOBALS - SHUNT
# =========================================================
shunt_data = {}
shunt_signal_routes_map = {}
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
# GLOBALS - SHARED CAPTURE STATE
# =========================================================
capture_module = None          # "MASTER" while a recording pass is active, else None
capture_index = 0
capture_list = []
record_stage = None
current_aspects = 2

# =========================================================
# GLOBALS - SHARED RUN STATE
# =========================================================
running = False

# =========================================================
# COLORS
# =========================================================
BG = "#0f172a"
BTN = "#2563eb"
ACCENT = "#22c55e"
RED = "#ef4444"


# =========================================================================
# =========================================================================
#  SHARED UTILITIES  (identical across all three original tools)
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


def click(point):

    if point is None:

        log("Missing Coordinate")
        return False

    x, y = point

    win32api.SetCursorPos((x, y))

    time.sleep(0.4)

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


def click_menu_item(name):

    try:

        item = auto.MenuItemControl(
            searchDepth=15,
            Name=name
        )

        if item.Exists(3):

            item.Click()

            log(f"Clicked Menu: {name}")

            return True

        log(f"Menu Not Found: {name}")

        return False

    except Exception as e:

        log(f"UIA ERROR: {e}")

        return False


def get_avg_color(x, y, screenshot):

    pixels = []

    for dx in range(-2, 3):

        for dy in range(-2, 3):

            pixels.append(
                screenshot.getpixel((x + dx, y + dy))
            )

    r = sum(p[0] for p in pixels) // len(pixels)
    g = sum(p[1] for p in pixels) // len(pixels)
    b = sum(p[2] for p in pixels) // len(pixels)

    return r, g, b


def is_route_initiated(point):
    """Checks the Route Initiation Indicator: YELLOW means the route was
    actually initiated in the interlocking, regardless of whether the
    signal itself cleared. Signal Cancel / Route Release are gated on
    this - not on the pass/fail result - so nothing is left locked."""

    if point is None:
        return False

    screenshot = pyautogui.screenshot()

    x, y = point

    r, g, b = get_avg_color(x, y, screenshot)

    return r > 150 and g > 150


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
#  UNIFIED REPORT
#  Every route tested (MAIN + CALLING-ON + SHUNT, pass and fail alike)
#  goes into one "ALL RESULTS" sheet. A second "CONSOLIDATED" sheet rolls
#  each signal up into one row (routes tested / passed / failed / overall).
# =========================================================================
# =========================================================================

report_rows = []       # each entry: dict with signal/type/route/track/detail/route_init/result/time


def create_report():

    global wb_report, REPORT_FILE, report_rows

    REPORT_FILE = (
        "Signal_Clearance_Report_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".xlsx"
    )

    wb_report = Workbook()

    report_rows = []


def record_result(signal, sig_type, route, track, detail, route_initiated, result):

    report_rows.append({
        "signal": signal,
        "type": sig_type,
        "route": route,
        "track": track or "",
        "detail": detail,
        "route_initiated": "YES" if route_initiated else "NO",
        "result": result,
        "time": datetime.now().strftime("%H:%M:%S")
    })


def write_report_sheets():
    """(Re)builds ALL RESULTS and CONSOLIDATED from report_rows. Safe to
    call more than once - e.g. once on STOP, again if resumed and
    finished normally - since it always rebuilds from scratch."""

    for name in list(wb_report.sheetnames):
        del wb_report[name]

    header_fill = PatternFill("solid", fgColor="1E293B")
    header_font = Font(color="FFFFFF", bold=True)
    pass_fill = PatternFill("solid", fgColor="C6EFCE")
    fail_fill = PatternFill("solid", fgColor="FFC7CE")

    # =====================================================
    # ALL RESULTS - every route, pass and fail together
    # =====================================================
    ws_all = wb_report.create_sheet("ALL RESULTS")

    ws_all.append(["Signal", "Type", "Route", "Track", "Detail", "Route Initiated", "Result", "Time"])

    for cell in ws_all[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row in report_rows:

        ws_all.append([
            row["signal"], row["type"], row["route"], row["track"],
            row["detail"], row["route_initiated"], row["result"], row["time"]
        ])

        r = ws_all.max_row
        fill = pass_fill if row["result"] == "PASS" else fail_fill

        for cell in ws_all[r]:
            cell.fill = fill

    autosize_columns(ws_all)

    # =====================================================
    # CONSOLIDATED - one row per signal
    # =====================================================
    ws_cons = wb_report.create_sheet("CONSOLIDATED")

    ws_cons.append(["Signal", "Type", "Routes Tested", "Passed", "Failed", "Overall Result"])

    for cell in ws_cons[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    per_signal = {}
    order = []

    for row in report_rows:

        key = (row["signal"], row["type"])

        if key not in per_signal:
            per_signal[key] = {"total": 0, "pass": 0, "fail": 0}
            order.append(key)

        per_signal[key]["total"] += 1

        if row["result"] == "PASS":
            per_signal[key]["pass"] += 1
        else:
            per_signal[key]["fail"] += 1

    total_pass = 0
    total_fail = 0
    total_routes = 0

    for key in order:

        signal, sig_type = key
        stats = per_signal[key]
        overall = "PASS" if stats["fail"] == 0 else "FAIL"

        ws_cons.append([signal, sig_type, stats["total"], stats["pass"], stats["fail"], overall])

        r = ws_cons.max_row
        fill = pass_fill if overall == "PASS" else fail_fill

        for cell in ws_cons[r]:
            cell.fill = fill

        total_pass += stats["pass"]
        total_fail += stats["fail"]
        total_routes += stats["total"]

    ws_cons.append([])
    ws_cons.append(["TOTAL SIGNALS", len(order)])
    ws_cons.append(["TOTAL ROUTES TESTED", total_routes])
    ws_cons.append(["TOTAL PASSED", total_pass])
    ws_cons.append(["TOTAL FAILED", total_fail])

    for r in range(ws_cons.max_row - 3, ws_cons.max_row + 1):
        ws_cons.cell(row=r, column=1).font = Font(bold=True)

    autosize_columns(ws_cons)


# =========================================================================
# =========================================================================
#  MASTER TOC IMPORT  (SINGLE SHEET)
#
#  Signal, Route, and Track columns are found by matching their header
#  text in row 1 - "Signal" / "Route" / "Track" - wherever they sit and
#  regardless of any other columns (Lock_Route, notes, etc.) mixed in
#  between them. No fixed column order is required.
#
#  Signal type is detected from the signal name itself, priority order:
#     1) Ends with "SH"  -> SHUNT     (Track column ignored)
#     2) Ends with "C"   -> CALLING-ON (Track column REQUIRED)
#     3) Otherwise        -> MAIN      (Track column ignored)
#
#  Looks for a sheet literally named "TOC"; if none exists, falls back to
#  whichever sheet is active (so a plain single-tab workbook still works).
# =========================================================================
# =========================================================================

def detect_signal_type(signal):
    """Detection priority: SH suffix -> C suffix -> otherwise MAIN."""

    if signal.endswith("SH"):
        return "SHUNT"

    if signal.endswith("C"):
        return "CAL"

    return "MAIN"


def import_master_toc():

    global main_data, main_signal_routes_map, main_signal_order
    global cal_excel_data, cal_signal_data, cal_signal_order
    global shunt_data, shunt_signal_routes_map, shunt_signal_order

    file = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx")]
    )

    if not file:
        return

    try:

        wb = load_workbook(file)

        sheet_map = {name.strip().upper(): name for name in wb.sheetnames}

        ws = wb[sheet_map["TOC"]] if "TOC" in sheet_map else wb.active

        # =============================================
        # HEADER-DRIVEN COLUMN LOOKUP
        # Finds Signal / Route / Track by their header text in row 1,
        # wherever they sit and however many other columns (Lock_Route,
        # notes, etc.) are mixed in around them - no fixed column order
        # required.
        # =============================================
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)

        if not header_row:
            log("TOC Import Error: Sheet has no header row")
            return

        header_index = {}

        for idx, cell_value in enumerate(header_row):

            if cell_value is None:
                continue

            key = str(cell_value).strip().upper()
            header_index[key] = idx

        signal_col = header_index.get("SIGNAL")
        route_col = header_index.get("ROUTE")
        track_col = header_index.get("TRACK")

        if signal_col is None or route_col is None:
            log(
                "TOC Import Error: Could not find 'Signal' and/or 'Route' "
                f"column headers. Headers found: {list(header_index.keys())}"
            )
            return

        if track_col is None:
            log("TOC Import Warning: No 'Track' column header found - "
                "Calling-On rows will be skipped unless one is added.")

        main_data = {}
        main_signal_routes_map = {}

        cal_excel_data = []
        cal_signal_data = {}
        cal_signal_order = []

        shunt_data = {}
        shunt_signal_routes_map = {}

        skipped = 0

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row or len(row) <= max(signal_col, route_col):
                continue

            if row[signal_col] is None or row[route_col] is None:
                continue

            signal = str(row[signal_col]).strip().upper()
            route = str(row[route_col]).strip()

            track = None

            if track_col is not None and len(row) > track_col and row[track_col] is not None:
                track = str(row[track_col]).strip().upper()

            signal_type = detect_signal_type(signal)

            # =============================================
            # SHUNT
            # =============================================
            if signal_type == "SHUNT":

                shunt_signal_routes_map.setdefault(signal, []).append(route)

                if signal not in shunt_data:

                    shunt_data[signal] = {
                        "menu_coordinate": None,
                        "state_indicator": None,
                        "initial_snapshot": None,
                        "route_init": None
                    }

            # =============================================
            # CALLING-ON
            # =============================================
            elif signal_type == "CAL":

                if not track:

                    log(f"{signal} {route} - Missing Track, Row Skipped")
                    skipped += 1
                    continue

                cal_excel_data.append({
                    "SIGNAL": signal,
                    "ROUTE": route,
                    "TRACK": track
                })

                if signal not in cal_signal_data:

                    cal_signal_data[signal] = {
                        "menu": None,
                        "yellow": None,
                        "route_init": None
                    }

                    cal_signal_order.append(signal)

            # =============================================
            # MAIN
            # =============================================
            else:

                main_signal_routes_map.setdefault(signal, []).append(route)

                if signal not in main_data:

                    main_data[signal] = {
                        "open_menu": None,
                        "RED": None,
                        "YELLOW": None,
                        "DOUBLE_YELLOW": None,
                        "GREEN": None,
                        "ROUTE_INIT": None
                    }

        main_signal_order = list(main_signal_routes_map.keys())
        shunt_signal_order = list(shunt_signal_routes_map.keys())

        main_refresh()
        cal_refresh_table()
        shunt_refresh()

        log("MASTER TOC IMPORT COMPLETE (SINGLE SHEET)")
        log(f"MAIN Signals Loaded : {len(main_signal_order)}")
        log(f"CALLING-ON Signals Loaded : {len(cal_signal_order)}")
        log(f"SHUNT Signals Loaded : {len(shunt_signal_order)}")

        if skipped:
            log(f"{skipped} Row(s) Skipped (Calling-On rows missing Track)")

    except Exception as e:

        log(f"Master TOC Import Error: {e}")


# =========================================================================
# =========================================================================
#  SHARED CAPTURE DISPATCHER
# =========================================================================
# =========================================================================

def main_refresh():

    main_tree.delete(*main_tree.get_children())

    for i, s in enumerate(main_signal_order, start=1):

        main_tree.insert(
            "", "end",
            values=(i, s, len(main_signal_routes_map.get(s, [])))
        )


def main_detect_color(r, g, b):

    if r > 150 and g < 120:
        return "RED"
    elif g > 150 and r < 120:
        return "GREEN"
    elif r > 150 and g > 150:
        return "YELLOW"

    return "UNKNOWN"


def main_get_signal_state(signal):
    """Samples every captured aspect coordinate and returns whichever one
    is BOTH a genuine color match AND the brightest. Checking coordinates
    in a fixed RED->YELLOW->DOUBLE_YELLOW->GREEN order and stopping at the
    first match is unreliable: an unlit lamp can still show a faint
    residual glow that clears the color threshold, so a real YELLOW aspect
    could get reported as RED just because RED is checked first. Comparing
    brightness across all matching coordinates avoids that."""

    screenshot = pyautogui.screenshot()

    check_order = ["RED", "YELLOW", "DOUBLE_YELLOW", "GREEN"]

    candidates = []

    for aspect in check_order:

        point = main_data[signal][aspect]

        if point is None:
            continue

        x, y = point

        r, g, b = get_avg_color(x, y, screenshot)

        detected = main_detect_color(r, g, b)

        if aspect == "DOUBLE_YELLOW":
            match = detected == "YELLOW"
        elif aspect == "GREEN":
            # For a 2-aspect signal, this coordinate is simply "the proceed
            # lamp" - it may render as GREEN or as YELLOW depending on the
            # signal type, so accept either rather than requiring GREEN.
            match = detected in ("GREEN", "YELLOW")
        else:
            match = detected == aspect

        log(f"{signal} {aspect} Coordinate -> RGB({r},{g},{b}) Detected: {detected}")

        if match:
            candidates.append((aspect, r + g + b))

    if not candidates:
        return "UNKNOWN"

    candidates.sort(key=lambda c: c[1], reverse=True)

    return candidates[0][0]


def main_convert_state_text(state):

    if state == "RED":
        return "SIGNAL IN DANGER"

    elif state in ["GREEN", "YELLOW", "DOUBLE_YELLOW"]:
        return "SIGNAL IS CLEAR"

    return state


def run_main_signal(signal):
    """Runs clearance testing for every lock route of one MAIN signal."""

    if signal not in main_data:
        log(f"{signal} Not Found In Coordinate File")
        return

    signal_point = main_data[signal]["open_menu"]

    if signal_point is None:
        log(f"{signal} Missing Coordinate")
        return

    route_init_point = main_data[signal].get("ROUTE_INIT")

    routes = main_signal_routes_map.get(signal, [])

    for route in routes:

        if not running:
            return

        initial_state = main_get_signal_state(signal)
        log(f"{signal} Initial State: {initial_state}")

        click(signal_point)
        time.sleep(1.5)

        log(f"{signal} Selecting Route {route}")

        ok = click_menu_item(route)

        if not ok:
            log(f"{route} FAILED")
            record_result(signal, "MAIN", route, "", "Menu item not found", False, "FAIL")
            continue

        log(f"{route} SET")

        time.sleep(5)

        final_state = main_get_signal_state(signal)
        log(f"{signal} Final State: {final_state}")

        if initial_state == "RED" and final_state in ["GREEN", "YELLOW", "DOUBLE_YELLOW"]:
            result = "PASS"
        else:
            result = "FAIL"

        log(f"{signal} {route} -> {result}")

        initial_text = main_convert_state_text(initial_state)
        final_text = main_convert_state_text(final_state)
        detail = f"{initial_text} -> {final_text}"

        # =============================================
        # ROUTE INITIATION CHECK
        # Cancel/Release runs on this, NOT on pass/fail -
        # a route that was initiated must be released
        # whether or not the signal itself cleared.
        # =============================================
        route_initiated = is_route_initiated(route_init_point)
        log(f"{signal} Route Initiation Indicator: {'YELLOW' if route_initiated else 'NOT YELLOW'}")

        record_result(signal, "MAIN", route, "", detail, route_initiated, result)

        if route_initiated:

            click(signal_point)
            time.sleep(1)
            click_menu_item("Signal Cancel")
            log("Signal Cancel Done")
            time.sleep(4)

            click(signal_point)
            time.sleep(1)
            click_menu_item("Route Release")
            log("Route Release Done")
            time.sleep(8)

        else:

            log(f"{signal} {route} - Route Not Initiated, Skipping Cancel/Release")

    log(f"{signal} MAIN Completed")


# =========================================================================
# =========================================================================
#  CALLING ON MODULE
# =========================================================================
# =========================================================================

def cal_refresh_table():

    cal_tree.delete(*cal_tree.get_children())

    for i, row in enumerate(cal_excel_data, start=1):

        cal_tree.insert(
            "", "end",
            values=(i, row["SIGNAL"], row["ROUTE"], row["TRACK"])
        )


def cal_find_and_click(name, control_type=None):

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

                            current_type = str(item.element_info.control_type)

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


def cal_open_bit_chart():

    pyautogui.click(500, 500)
    time.sleep(1)

    pyautogui.hotkey("ctrl", "b")
    log("CTRL+B Pressed")

    time.sleep(3)

    found = cal_find_and_click("50051", control_type="Text")

    if not found:
        log("50051 Not Found")
        return False

    time.sleep(1)

    ok = cal_find_and_click("OK", control_type="Button")

    if not ok:
        log("OK Button Not Found")
        return False

    time.sleep(3)

    return True


def cal_drop_track(track):

    log(f"Dropping Track : {track}")

    if not cal_open_bit_chart():
        return False

    if not cal_find_and_click(track):
        log(f"{track} Not Found")
        return False

    time.sleep(1)

    if not cal_find_and_click("TRANSMIT", control_type="Button"):
        log("TRANSMIT Not Found")
        return False

    time.sleep(1)

    if not cal_find_and_click("CANCEL", control_type="Button"):
        log("CANCEL Not Found")
        return False

    log(f"{track} Dropped Successfully")

    time.sleep(5)

    return True


def cal_restore_track(track):

    log(f"Restoring Track : {track}")

    if not cal_open_bit_chart():
        return False

    if not cal_find_and_click(track):
        log(f"{track} Not Found")
        return False

    time.sleep(1)

    if not cal_find_and_click("TRANSMIT", control_type="Button"):
        log("TRANSMIT Not Found")
        return False

    time.sleep(1)

    if not cal_find_and_click("CANCEL", control_type="Button"):
        log("CANCEL Not Found")
        return False

    log(f"{track} Restored Successfully")

    time.sleep(5)

    return True


def cal_detect_yellow(r, g, b):

    if r > 150 and g > 150:
        return True

    return False


def cal_get_signal_status(signal):

    screenshot = pyautogui.screenshot()

    x, y = cal_signal_data[signal]["yellow"]

    r, g, b = get_avg_color(x, y, screenshot)

    return "YELLOW" if cal_detect_yellow(r, g, b) else "NO ASPECT"


def cal_operate_signal(signal, route):

    menu_point = cal_signal_data[signal]["menu"]

    click(menu_point)
    time.sleep(1)

    ok = click_menu_item(route)

    if not ok:
        return False

    log(f"{signal} {route} Selected")

    time.sleep(8)

    return True


def cal_signal_cancel(signal):

    menu_point = cal_signal_data[signal]["menu"]

    click(menu_point)
    time.sleep(1)

    click_menu_item("Signal Cancel")
    log("Signal Cancel Done")

    time.sleep(3)


def cal_route_release(signal):

    menu_point = cal_signal_data[signal]["menu"]

    click(menu_point)
    time.sleep(1)

    click_menu_item("Route Release")
    log("Route Release Done")

    time.sleep(5)


def get_calling_on_signal_for(main_signal):
    """Finds the CALLING-ON signal linked to a MAIN signal (e.g. '1' -> '1C')."""

    candidate = (main_signal + "C").upper()

    for s in cal_signal_order:

        if s.upper() == candidate:
            return s

    return None


def get_calling_on_rows(signal):

    return [r for r in cal_excel_data if r["SIGNAL"] == signal]


def run_calling_on_signal(signal):
    """Runs clearance testing for every route of one CALLING-ON signal."""

    rows = get_calling_on_rows(signal)

    if not rows:
        log(f"{signal} No CALLING-ON Routes Found")
        return

    route_init_point = cal_signal_data.get(signal, {}).get("route_init")

    for row in rows:

        if not running:
            return

        route = row["ROUTE"]
        track = row["TRACK"]

        log(f"Processing : {signal} | {route} | {track}")

        track_ok = cal_drop_track(track)

        if not track_ok:
            log(f"Track Failed : {track}")
            record_result(signal, "CALLING-ON", route, track, "Track drop failed", False, "FAIL")
            continue

        route_ok = cal_operate_signal(signal, route)

        if not route_ok:
            log(f"{route} Failed")
            record_result(signal, "CALLING-ON", route, track, "Route selection failed", False, "FAIL")
            continue

        signal_status = cal_get_signal_status(signal)
        log(f"Signal Status : {signal_status}")

        result = "PASS" if signal_status == "YELLOW" else "FAIL"

        # =============================================
        # ROUTE INITIATION CHECK
        # Cancel/Release runs on this, NOT on pass/fail -
        # a route that was initiated must be released
        # whether or not the signal itself cleared.
        # =============================================
        route_initiated = is_route_initiated(route_init_point)
        log(f"{signal} Route Initiation Indicator: {'YELLOW' if route_initiated else 'NOT YELLOW'}")

        record_result(signal, "CALLING-ON", route, track, f"Signal Status: {signal_status}", route_initiated, result)

        log(f"{signal} {route} -> {result}")

        if route_initiated:
            cal_signal_cancel(signal)
            cal_route_release(signal)
        else:
            log(f"{signal} {route} - Route Not Initiated, Skipping Cancel/Release")

        cal_restore_track(track)

        time.sleep(3)

    log(f"{signal} CALLING-ON Completed")


# =========================================================================
# =========================================================================
#  SHUNT MODULE
# =========================================================================
# =========================================================================

def shunt_refresh():

    shunt_tree.delete(*shunt_tree.get_children())

    for i, s in enumerate(shunt_signal_order, start=1):

        shunt_tree.insert(
            "", "end",
            values=(i, s, len(shunt_signal_routes_map.get(s, [])))
        )


def shunt_compare_by_features(signal, initial_path):
    """Compares images using ORB feature detection."""

    try:

        point = shunt_data[signal]["state_indicator"]

        if point is None:
            return 0, "No coordinate"

        initial_img = cv2.imread(initial_path)

        if initial_img is None:
            return 0, "Cannot load initial"

        screenshot = pyautogui.screenshot()

        x, y = point

        current_pil = screenshot.crop((x - 50, y - 50, x + 50, y + 50))

        current_img = cv2.cvtColor(np.array(current_pil), cv2.COLOR_RGB2BGR)

        initial_gray = cv2.cvtColor(initial_img, cv2.COLOR_BGR2GRAY)
        current_gray = cv2.cvtColor(current_img, cv2.COLOR_BGR2GRAY)

        kp1, des1 = orb.detectAndCompute(initial_gray, None)
        kp2, des2 = orb.detectAndCompute(current_gray, None)

        if des1 is None or des2 is None:
            return 0, "Cannot detect features"

        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)

        matches = bf.match(des1, des2)
        matches = sorted(matches, key=lambda m: m.distance)

        if len(matches) > 0:

            good_matches = [m for m in matches if m.distance < 50]

            match_score = (
                (len(good_matches) / max(len(kp1), len(kp2))) * 100
                if max(len(kp1), len(kp2)) > 0 else 0
            )

            log(f"{signal} Features: Initial={len(kp1)}, Current={len(kp2)}, Matches={len(matches)}, Good={len(good_matches)}")
            log(f"{signal} Match Score: {match_score:.2f}%")

        else:

            match_score = 0
            log(f"{signal} No matches found")

        try:

            if len(matches) > 3:

                matched_img = cv2.drawMatches(
                    initial_gray, kp1, current_gray, kp2, matches[:10], None,
                    flags=cv2.DrawMatchesFlags_NOT_DRAW_SINGLE_POINTS
                )

                debug_file = os.path.join(
                    DEBUG_FOLDER,
                    f"{signal}_features_{datetime.now().strftime('%H%M%S')}.png"
                )

                cv2.imwrite(debug_file, matched_img)

            else:

                debug_file = os.path.join(
                    DEBUG_FOLDER,
                    f"{signal}_current_{datetime.now().strftime('%H%M%S')}.png"
                )

                cv2.imwrite(debug_file, current_img)

        except Exception as e:

            log(f"Debug save error: {e}")

        return match_score, f"Features matched: {len(matches)}"

    except Exception as e:

        log(f"Feature Comparison Error: {e}")
        return 0, str(e)


def run_shunt_signal(signal):
    """Runs clearance testing for every route of one SHUNT signal."""

    if signal not in shunt_data:
        log(f"{signal} Not Found")
        return

    menu_point = shunt_data[signal]["menu_coordinate"]

    if menu_point is None:
        log(f"{signal} Missing Menu Coordinate")
        return

    initial_snapshot = shunt_data[signal]["initial_snapshot"]

    if initial_snapshot is None:
        log(f"{signal} Missing Initial Snapshot")
        return

    route_init_point = shunt_data[signal].get("route_init")

    routes = shunt_signal_routes_map.get(signal, [])

    for route in routes:

        if not running:
            return

        log(f"Initial snapshot ready: {initial_snapshot}")

        click(menu_point)
        time.sleep(1.5)

        log(f"Selecting Route: {route}")

        ok = click_menu_item(route)

        if not ok:

            log(f"{route} FAILED - Menu Item Not Found")

            record_result(signal, "SHUNT", route, "", "Menu item not found", False, "MENU_ERROR")

            continue

        log(f"Route {route} SET Command Sent")

        # =============================================
        # MANDATORY 7 SECOND WAIT
        # Give the panel time to react before checking
        # anything - both the Route Initiation Indicator
        # and the feature-match comparison read the panel
        # AFTER this wait, never before it.
        # =============================================
        log("Waiting 7 seconds after setting route...")
        time.sleep(7)

        # =============================================
        # ROUTE INITIATION CHECK (checked first)
        # This alone decides whether Signal Cancel / Route
        # Release run - NOT the pass/fail result below.
        # If the indicator never goes YELLOW, nothing was
        # actually locked, so cleanup is not required -
        # regardless of whether the signal itself cleared.
        # =============================================
        route_initiated = is_route_initiated(route_init_point)
        log(f"{signal} Route Initiation Indicator: {'YELLOW' if route_initiated else 'NOT YELLOW'}")

        log("Comparing features with initial snapshot...")

        match_score, details = shunt_compare_by_features(signal, initial_snapshot)

        log(f"{signal} -> {route} | Feature Match: {match_score:.2f}% | {details}")

        if match_score < 70:

            result = "PASS"
            log(f"Feature change detected ({match_score:.2f}%) -> PASS")

        elif match_score > 85:

            result = "FAIL"
            log(f"No significant feature change ({match_score:.2f}%) -> FAIL")

        else:

            result = "PASS"
            log(f"Moderate feature change ({match_score:.2f}%) -> PASS (benefit)")

        record_result(signal, "SHUNT", route, "", f"Feature Match: {match_score:.2f}%", route_initiated, result)

        if route_initiated:

            log(f"Clearance: {signal} -> {route}")

            click(menu_point)
            time.sleep(1)
            log("Sending: Signal Cancel")
            click_menu_item("Signal Cancel")
            time.sleep(3)

            click(menu_point)
            time.sleep(1)
            log("Sending: Route Release")
            click_menu_item("Route Release")
            time.sleep(3)

            log(f"Clearance Completed: {signal} -> {route}")

        else:

            log(
                f"Skip Clearance: {signal} -> {route} - "
                f"Route Initiation Indicator Not Yellow (Result: {result}) - "
                f"Signal Cancel / Route Release Not Required"
            )

    log(f"{signal} SHUNT Completed")


# =========================================================================
# =========================================================================
#  UNIFIED COORDINATE CAPTURE
#  One recording pass, one save file, one load file - covers MAIN,
#  CALLING-ON, and SHUNT signals together. Each signal's TYPE (and
#  therefore which sub-steps it needs) is worked out automatically from
#  its name via detect_signal_type() - no separate flows to run.
# =========================================================================
# =========================================================================

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

        main_refresh()
        cal_refresh_table()
        shunt_refresh()

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


def start_run():

    global running

    if not main_signal_order and not shunt_signal_order:
        log("Import TOC First")
        return

    create_report()

    running = True

    status_label.config(text="RUNNING", fg="#16a34a")

    root.iconify()

    threading.Thread(target=run_engine, daemon=True).start()


def run_engine():

    global running

    time.sleep(2)

    log("=== SIGNAL CLEARANCE AUTOMATION STARTED ===")

    # =====================================================
    # PHASE 1 : MAIN SIGNAL -> ITS CALLING-ON, FOR EACH SIGNAL
    # =====================================================
    for signal in main_signal_order:

        if not running:
            write_report_sheets()
            wb_report.save(REPORT_FILE)
            root.deiconify()
            return

        log(f"--- Processing MAIN Signal {signal} ---")

        run_main_signal(signal)

        cal_signal = get_calling_on_signal_for(signal)

        if cal_signal:

            log(f"--- Processing CALLING-ON {cal_signal} (for {signal}) ---")

            run_calling_on_signal(cal_signal)

        else:

            log(f"No CALLING-ON Found For {signal}")

    # =====================================================
    # PHASE 2 : ALL SHUNT SIGNALS
    # =====================================================
    if shunt_signal_order:

        log("=== STARTING SHUNT SIGNAL CLEARANCE ===")

        for signal in shunt_signal_order:

            if not running:
                write_report_sheets()
                wb_report.save(REPORT_FILE)
                root.deiconify()
                return

            log(f"--- Processing SHUNT Signal {signal} ---")

            run_shunt_signal(signal)

    # =====================================================
    # SAVE REPORT
    # =====================================================
    write_report_sheets()

    wb_report.save(REPORT_FILE)

    running = False

    status_label.config(text="COMPLETED", fg="#16a34a")

    root.deiconify()

    log("ALL SIGNALS COMPLETED")

    os.startfile(REPORT_FILE)


def stop_run():

    global running

    running = False

    status_label.config(text="STOPPED", fg="#dc2626")

    root.deiconify()

    log("STOPPED")


# =========================================================================
# =========================================================================
#  GUI
# =========================================================================
# =========================================================================

root = tk.Tk()

root.title("SIGNAL CLEARANCE AUTOMATION SYSTEM")

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

style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[14, 6])

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
    title_frame, text="SIGNAL CLEARANCE AUTOMATION SYSTEM",
    font=("Segoe UI", 22, "bold"), bg="#0f172a", fg="white"
).pack(pady=(15, 0))

tk.Label(
    title_frame, text="MAIN -> CALLING-ON  (per signal)   |   SHUNT (all, at end)",
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


def create_button(parent, text, command, color, width=26, height=2, font_size=11):

    return tk.Button(
        parent, text=text, command=command,
        bg=color, fg="white",
        activebackground=color, activeforeground="white",
        relief="flat", cursor="hand2",
        font=("Segoe UI", font_size, "bold"),
        width=width, height=height
    )


# =========================================================
# SIGNAL CONFIG SECTION
# =========================================================

signal_frame = tk.Frame(left_panel, bg="white")
signal_frame.pack(pady=10)

# =========================================================
# DUMMY MAIN BUTTON
# =========================================================

dummy_btn = tk.Button(
    signal_frame,
    text="CAPTURE SIGNALLING GEARS",
    bg="#2563eb", fg="white",
    activebackground="#2563eb", activeforeground="white",
    relief="flat",
    font=("Segoe UI", 11, "bold"),
    width=26, height=2
)

dummy_btn.pack(pady=(0, 8))

# =========================================================
# LOAD YARD COORDINATES BUTTON
# =========================================================

create_button(
    left_panel, "LOAD YARD COORDINATES", master_load_config, "#7c3aed"
).pack(pady=8)

# =========================================================
# OTHER BUTTONS
# =========================================================

create_button(
    left_panel, "IMPORT MASTER TOC", import_master_toc, "#2563eb"
).pack(pady=8)

create_button(left_panel, "START TESTING", start_run, "#16a34a").pack(pady=20)
create_button(left_panel, "STOP TESTING", stop_run, "#dc2626").pack(pady=8)

status_frame = tk.Frame(left_panel, bg="#f8fafc", bd=1, relief="solid")
status_frame.pack(fill="x", padx=15, pady=25)

tk.Label(
    status_frame, text="SYSTEM STATUS",
    font=("Segoe UI", 11, "bold"), bg="#f8fafc", fg="#0f172a"
).pack(pady=10)

status_label = tk.Label(
    status_frame, text="READY",
    font=("Segoe UI", 16, "bold"), bg="#f8fafc", fg="#16a34a"
)
status_label.pack(pady=(0, 15))

# =========================================================
# RIGHT PANEL - NOTEBOOK + LOG
# =========================================================

right_panel = tk.Frame(main_frame, bg="#e9edf2")
right_panel.pack(side="left", fill="both", expand=True)

notebook = ttk.Notebook(right_panel)
notebook.pack(fill="both", expand=True)

# --- MAIN TAB ---
main_tab = tk.Frame(notebook, bg="white")
notebook.add(main_tab, text="MAIN SIGNALS")

main_tree_scroll = ttk.Scrollbar(main_tab)
main_tree_scroll.pack(side="right", fill="y")

main_tree = ttk.Treeview(
    main_tab, columns=("NO", "SIGNAL", "ROUTES"),
    show="headings", yscrollcommand=main_tree_scroll.set
)
main_tree.heading("NO", text="NO")
main_tree.heading("SIGNAL", text="SIGNAL")
main_tree.heading("ROUTES", text="TOTAL ROUTES")
main_tree.column("NO", width=80, anchor="center")
main_tree.column("SIGNAL", width=250, anchor="center")
main_tree.column("ROUTES", width=200, anchor="center")
main_tree.pack(fill="both", expand=True)
main_tree_scroll.config(command=main_tree.yview)

# --- CALLING ON TAB ---
cal_tab = tk.Frame(notebook, bg="white")
notebook.add(cal_tab, text="CALLING-ON SIGNALS")

cal_tree_scroll = ttk.Scrollbar(cal_tab)
cal_tree_scroll.pack(side="right", fill="y")

cal_tree = ttk.Treeview(
    cal_tab, columns=("NO", "SIGNAL", "ROUTE", "TRACK"),
    show="headings", yscrollcommand=cal_tree_scroll.set
)
cal_tree.heading("NO", text="NO")
cal_tree.heading("SIGNAL", text="SIGNAL")
cal_tree.heading("ROUTE", text="ROUTE")
cal_tree.heading("TRACK", text="TRACK")
cal_tree.column("NO", width=80, anchor="center")
cal_tree.column("SIGNAL", width=220, anchor="center")
cal_tree.column("ROUTE", width=300, anchor="center")
cal_tree.column("TRACK", width=220, anchor="center")
cal_tree.pack(fill="both", expand=True)
cal_tree_scroll.config(command=cal_tree.yview)

# --- SHUNT TAB ---
shunt_tab = tk.Frame(notebook, bg="white")
notebook.add(shunt_tab, text="SHUNT SIGNALS")

shunt_tree_scroll = ttk.Scrollbar(shunt_tab)
shunt_tree_scroll.pack(side="right", fill="y")

shunt_tree = ttk.Treeview(
    shunt_tab, columns=("NO", "SIGNAL", "ROUTES"),
    show="headings", yscrollcommand=shunt_tree_scroll.set
)
shunt_tree.heading("NO", text="NO")
shunt_tree.heading("SIGNAL", text="SIGNAL")
shunt_tree.heading("ROUTES", text="TOTAL ROUTES")
shunt_tree.column("NO", width=80, anchor="center")
shunt_tree.column("SIGNAL", width=250, anchor="center")
shunt_tree.column("ROUTES", width=200, anchor="center")
shunt_tree.pack(fill="both", expand=True)
shunt_tree_scroll.config(command=shunt_tree.yview)

# =========================================================
# LOG
# =========================================================

log_title = tk.Label(
    right_panel, text="LIVE OPERATION LOG",
    font=("Segoe UI", 16, "bold"), bg="#e9edf2", fg="#0f172a"
)
log_title.pack(anchor="w", pady=(15, 10))

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
    text="SPACE = CAPTURE  |  BACKSPACE = UNDO LAST  |  ORDER: MAIN -> CALLING-ON (per signal) -> SHUNT (all, at end)",
    bg="#0f172a", fg="white", font=("Segoe UI", 10)
)
footer.pack(fill="x")

# =========================================================
# RUN
# =========================================================

root.mainloop()
