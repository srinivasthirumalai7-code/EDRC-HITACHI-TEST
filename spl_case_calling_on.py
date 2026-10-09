# =========================================================
# SPL CASE CALLING ON SIGNAL AUTOMATION SYSTEM
# =========================================================

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import filedialog, ttk
import threading
import time
import os
import win32api
import win32con
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment

from pynput import keyboard

import uiautomation as auto

import pyautogui

from pywinauto import Desktop

from datetime import datetime

# =========================================================
# REPORT FILE
# =========================================================


# =========================================================
# GLOBALS
# =========================================================

excel_data = []

signal_data = {}

main_signal_data = {}

signal_order = []

main_signal_order = []

capture_index = 0

capture_list = []


running = False

last_capture_time = 0
# =========================================================
# CREATE REPORT
# =========================================================

def create_report():

    global wb_report
    global ws_report
    global ws_fail
    global REPORT_FILE

    REPORT_FILE = (
        "Calling_On_Report_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".xlsx"
    )

    wb_report = Workbook()

    ws_report = wb_report.active
    ws_report.title = "ALL RESULTS"

    ws_fail = wb_report.create_sheet(
        title="FAILED ROUTES"
    )

    headers = [

        "MAIN_SIGNAL",

        "MAIN_ROUTE",

        "CALLING_ON_SIGNAL",

        "CALLING_ON_ROUTE",

        "TRACK",

        "STATUS",

        "RESULT",

        "TIME"
    ]

    ws_report.append(headers)
    ws_fail.append(headers)

    header_fill = PatternFill(
        "solid",
        fgColor="1E293B"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    for ws in [ws_report, ws_fail]:

        for cell in ws[1]:

            cell.fill = header_fill

            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center"
            )
# =========================================================
# COLORS
# =========================================================

BG = "#0f172a"

BTN = "#2563eb"

GREEN = "#22c55e"

RED = "#ef4444"

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

        log("Coordinate Missing")

        return False

    x, y = point

    win32api.SetCursorPos((x, y))

    time.sleep(0.3)

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

    time.sleep(0.3)

    return True

# =========================================================
# LOAD EXCEL
# =========================================================

def load_excel():
    global excel_data
    global signal_data
    global signal_order
    global main_signal_data
    global main_signal_order
    file = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx")]
    )

    if not file:
        return

    try:

        wb = load_workbook(file)

        # =================================================
        # Look for a sheet named "SPL_CASE_CALLING_ON" - this
        # program's data shape (Main Signal/Route + Calling-On
        # Signal/Route + Track) is different from the standard
        # TOC schema, so it gets its own dedicated sheet. Falls
        # back to the active sheet so a plain single-tab
        # workbook still works.
        # =================================================

        sheet_map = {name.strip().upper(): name for name in wb.sheetnames}

        ws = wb[sheet_map["SPL_CASE_CALLING_ON"]] if "SPL_CASE_CALLING_ON" in sheet_map else wb.active

        # =================================================
        # HEADER-DRIVEN COLUMN LOOKUP
        # Finds each column by its header text in row 1, wherever it
        # sits and regardless of other columns mixed in. Accepts a
        # couple of reasonable header spellings per column.
        # =================================================

        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)

        if not header_row:
            log("TOC Import Error: Sheet has no header row")
            return

        header_occurrences = {}
        for idx, val in enumerate(header_row):
            if val is None:
                continue
            key = str(val).strip().upper()
            header_occurrences.setdefault(key, []).append(idx)

        def find_col(*aliases):
            for a in aliases:
                if a in header_occurrences:
                    return header_occurrences[a][0]
            return None

        main_signal_col = find_col("MAIN_SIGNAL", "MAIN SIGNAL")
        signal_col = find_col("CALLING_ON_SIGNAL", "CALLING ON SIGNAL", "CALLING_ON", "CALLING ON", "SIGNAL")
        track_col = find_col("TRACK")

        # Explicit distinct route headers, if your sheet has them
        main_route_col = find_col("MAIN_ROUTE", "MAIN ROUTE")
        route_col = find_col("CALLING_ON_ROUTE", "CALLING ON ROUTE")

        # Your sheet instead has "Route" appearing TWICE (same header
        # text for both Main Route and Calling-On Route) - a plain
        # header->column lookup can only keep one of them. Resolve the
        # duplicates by position instead: whichever "Route" column sits
        # BEFORE the Calling-On Signal column is the Main Route, whichever
        # sits AFTER it is the Calling-On Route.
        route_occurrences = header_occurrences.get("ROUTE", [])

        used_single_route_column = False

        if main_route_col is None:
            before = [i for i in route_occurrences if signal_col is not None and i < signal_col]
            if before:
                main_route_col = before[-1]
            elif route_occurrences:
                main_route_col = route_occurrences[0]
                used_single_route_column = True

        if route_col is None:
            after = [i for i in route_occurrences if signal_col is not None and i > signal_col]
            if after:
                route_col = after[0]
            elif len(route_occurrences) > 1:
                route_col = route_occurrences[-1]
            elif route_occurrences:
                route_col = route_occurrences[0]
                used_single_route_column = True

        required = {
            "MAIN_SIGNAL": main_signal_col, "MAIN_ROUTE": main_route_col,
            "CALLING_ON_SIGNAL": signal_col, "CALLING_ON_ROUTE": route_col, "TRACK": track_col
        }
        missing = [name for name, col in required.items() if col is None]

        if missing:
            log(
                f"TOC Import Error: Could not find column header(s) {missing}. "
                f"Headers found: {list(header_occurrences.keys())}"
            )
            return

        excel_data = []

        signal_data = {}

        signal_order = []
        main_signal_data = {}

        main_signal_order = []

        def cell(row, idx):
            return row[idx] if idx is not None and len(row) > idx else None

        for row in ws.iter_rows(min_row=2, values_only=True):

            if not row:
                continue

            if cell(row, main_signal_col) is None:
                continue

            if cell(row, main_route_col) is None:
                continue

            if cell(row, signal_col) is None:
                continue

            main_signal = str(cell(row, main_signal_col)).strip().upper()

            main_route = str(cell(row, main_route_col)).strip()

            signal = str(cell(row, signal_col)).strip().upper()

            route = str(cell(row, route_col) or "").strip()

            track = str(cell(row, track_col) or "").strip().upper()

            excel_data.append({

                "MAIN_SIGNAL": main_signal,

                "MAIN_ROUTE": main_route,

                "SIGNAL": signal,

                "ROUTE": route,

                "TRACK": track

            })

            # Calling On Signal

            if signal not in signal_data:
                signal_data[signal] = {

                    "menu": None,

                    "yellow": None

                }

                signal_order.append(signal)

            # Main Signal

            if main_signal not in main_signal_data:
                main_signal_data[main_signal] = {

                    "menu": None,

                    "red": None,

                    "yellow": None,

                    "green": None,

                    "aspect_count": None
                }

                main_signal_order.append(main_signal)

        refresh_table()

        log("SPL_CALLING_ON Sheet Loaded Successfully")

        if used_single_route_column:
            log("NOTE: Only one 'Route' column found - using it as both Main Route and Calling-On Route")

    except Exception as e:

        log(f"Excel Error : {e}")

# =========================================================
# REFRESH TABLE
# =========================================================

def refresh_table():

    tree.delete(*tree.get_children())

    for i, row in enumerate(excel_data, start=1):
        tree.insert(
            "",
            "end",
            values=(

                i,

                row["MAIN_SIGNAL"],

                row["MAIN_ROUTE"],

                row["SIGNAL"],

                row["ROUTE"],

                row["TRACK"]
            )
        )

# =========================================================
# FIND & CLICK UI
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
# CLICK MENU ITEM
# =========================================================

def click_menu_item(name):

    try:

        item = auto.MenuItemControl(
            searchDepth=15,
            Name=name
        )

        if item.Exists(3):

            item.Click()

            log(f"Clicked Menu : {name}")

            return True

        else:

            log(f"Menu Not Found : {name}")

            return False

    except Exception as e:

        log(f"UI Error : {e}")

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

    log(f"Dropping Track : {track}")

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

    log(f"{track} Dropped Successfully")

    time.sleep(5)

    return True
# =========================================================
# RESTORE TRACK
# =========================================================

def restore_track(track):

    log(f"Restoring Track : {track}")

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

    log(f"{track} Restored Successfully")

    time.sleep(5)

    return True
# =========================================================
# GET AVG COLOR
# =========================================================

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

# =========================================================
# DETECT YELLOW
# =========================================================

def detect_yellow(r, g, b):

    if r > 150 and g > 150:

        return True

    return False

# =========================================================
# GET SIGNAL STATUS
# =========================================================

def get_signal_status(signal):

    screenshot = pyautogui.screenshot()

    x, y = signal_data[signal]["yellow"]

    r, g, b = get_avg_color(
        x,
        y,
        screenshot
    )

    yellow = detect_yellow(r, g, b)

    if yellow:

        return "YELLOW"

    return "NO ASPECT"

# =========================================================
# OPERATE SIGNAL
# =========================================================

def operate_signal(signal, route):

    menu_point = signal_data[signal]["menu"]

    click(menu_point)

    time.sleep(1)

    ok = click_menu_item(route)

    if not ok:

        pyautogui.press("esc")

        log(f"{signal} {route} Menu Closed")

        return False

    log(f"{signal} {route} Selected")

    time.sleep(8)

    status = get_signal_status(signal)

    if status != "YELLOW":

        pyautogui.press("esc")

        log(f"{signal} {route} NOT SET -> ESC Pressed")

        return False

    return True

# =========================================================
# SIGNAL CANCEL
# =========================================================

def signal_cancel(signal):

    menu_point = signal_data[signal]["menu"]

    click(menu_point)

    time.sleep(1)

    click_menu_item("Signal Cancel")

    log("Signal Cancel Done")

    time.sleep(3)

# =========================================================
# ROUTE RELEASE
# =========================================================

def route_release(signal):

    menu_point = signal_data[signal]["menu"]

    click(menu_point)

    time.sleep(1)

    click_menu_item("Route Release")

    log("Route Release Done")

    time.sleep(5)

def get_main_signal_status(main_signal):

    try:

        screenshot = pyautogui.screenshot()

        data = main_signal_data[main_signal]

        # GREEN CHECK

        if data["green"]:

            x, y = data["green"]

            r, g, b = screenshot.getpixel((x, y))

            if g > 150:

                return "GREEN"

        # YELLOW CHECK

        if data["yellow"]:

            x, y = data["yellow"]

            r, g, b = screenshot.getpixel((x, y))

            if r > 150 and g > 150:

                return "YELLOW"

        return "RED"

    except:

        return "RED"

def set_main_route(main_signal, route):

    try:

        menu = main_signal_data[main_signal]["menu"]

        pyautogui.click(menu[0], menu[1])

        time.sleep(1)

        if click_menu_item(route):

            log(f"{main_signal} -> {route} SET")

            return True

        log(f"{main_signal} -> {route} NOT FOUND")

        return False

    except Exception as e:

        log(f"Set Main Route Error : {e}")

        return False

def verify_main_route(main_signal):

    try:

        status = get_main_signal_status(main_signal)

        if status in ["GREEN", "YELLOW"]:

            log(f"{main_signal} Route Set Verified ({status})")

            return True

        log(f"{main_signal} Route Not Set")

        return False

    except Exception as e:

        log(f"Verify Main Route Error : {e}")

        return False

def cancel_main_signal(main_signal):

    try:

        menu = main_signal_data[main_signal]["menu"]

        pyautogui.click(menu[0], menu[1])

        time.sleep(1)

        if click_menu_item("Signal Cancel"):

            log(f"{main_signal} Signal Cancel Done")

            return True

        return False

    except Exception as e:

        log(f"Signal Cancel Error : {e}")

        return False

def verify_red(main_signal):

    try:

        status = get_main_signal_status(main_signal)

        if status == "RED":

            log(f"{main_signal} Returned To RED")

            return True

        log(f"{main_signal} Not Returned To RED")

        return False

    except Exception as e:

        log(f"Verify RED Error : {e}")

        return False

# =========================================================
# START
# =========================================================

def start_run():

    global running

    if not excel_data:
        log("Load Excel First")
        return

    create_report()

    running = True
    threading.Thread(
        target=run_engine,
        daemon=True
    ).start()


import re

def expected_calling_on(main_route, co_route):

    try:

        main_suffix = main_route.split("_")[-1].strip().upper()
        co_suffix = co_route.split("_")[-1].strip().upper()

        main_letter = re.match(r"[A-Z]+", main_suffix)
        co_letter = re.match(r"[A-Z]+", co_suffix)

        if not main_letter or not co_letter:
            return False

        return main_letter.group() == co_letter.group()

    except:
        return False

def release_main_route(main_signal):

    try:

        menu = main_signal_data[main_signal]["menu"]

        click(menu)

        time.sleep(1)

        click_menu_item("Route Release")

        log(f"{main_signal} Main Route Released")

        time.sleep(5)

        return True

    except Exception as e:

        log(f"Main Route Release Error : {e}")

        return False
# =========================================================
# MAIN ENGINE
# =========================================================

def run_engine():

    global running

    root.iconify()

    time.sleep(2)

    log("AUTOMATION STARTED")

    for row_index, row in enumerate(excel_data):

        if not running:
            break

        signal = row["SIGNAL"]
        co_route = row["ROUTE"]
        track = row["TRACK"]
        main_signal = row["MAIN_SIGNAL"]
        main_route = row["MAIN_ROUTE"]

        log(
            f"Processing : {main_signal} {main_route} -> "
            f"{signal} {co_route} | Track {track}"
        )

        # =================================================
        # SET MAIN ROUTE
        # =================================================

        if not set_main_route(main_signal, main_route):
            log("Main Route Failed")
            continue

        time.sleep(5)

        if not verify_main_route(main_signal):
            log("Main Route Not Verified")
            continue

        # =================================================
        # SIGNAL CANCEL MAIN
        # =================================================

        cancel_main_signal(main_signal)

        time.sleep(3)

        verify_red(main_signal)

        time.sleep(2)

        # =================================================
        # DROP TRACK
        # =================================================

        track_ok = drop_track(track)

        if not track_ok:
            log(f"Track Failed : {track}")
            release_main_route(main_signal)
            time.sleep(3)
            continue

        # =================================================
        # TEST CALLING-ON ROUTE
        # =================================================

        log(f"Testing : {co_route}")

        route_ok = operate_signal(signal, co_route)

        if not route_ok:

            log(f"{co_route} Failed")

            signal_status = "NO ASPECT"

        else:

            signal_status = get_signal_status(signal)

            if signal_status == "YELLOW":
                signal_cancel(signal)
                route_release(signal)
                time.sleep(3)

        # =================================================
        # TRACK UP
        # =================================================

        restore_track(track)

        time.sleep(3)

        # =================================================
        # RELEASE MAIN ROUTE
        # Only needed if the calling-on route never set - its own
        # Signal Cancel + Route Release already releases the main
        # route too when it DID set, so a separate main release here
        # would be redundant (and wrong) in that case.
        # =================================================

        if signal_status != "YELLOW":

            release_main_route(main_signal)

            time.sleep(3)

        else:

            log(f"{main_signal} route already released via {signal} Route Release")

        # =================================================
        # RECORD RESULT
        # Every row is a combination that's expected to work - if the
        # calling-on route set (YELLOW), that's a PASS; if it never
        # set (NO ASPECT), that's a FAIL.
        # =================================================

        result = "PASS" if signal_status == "YELLOW" else "FAIL"

        now = datetime.now().strftime("%H:%M:%S")

        row_data = [

            main_signal,

            main_route,

            signal,

            co_route,

            track,

            signal_status,

            result,

            now

        ]

        ws_report.append(row_data)

        current_row = ws_report.max_row

        if result == "PASS":

            fill = PatternFill(
                "solid",
                fgColor="C6EFCE"
            )

        else:

            fill = PatternFill(
                "solid",
                fgColor="FFC7CE"
            )

            ws_fail.append(row_data)

        for cell in ws_report[current_row]:
            cell.fill = fill

        log(
            f"{signal} {co_route}"
            f" -> {result}"
        )

        time.sleep(3)

    running = False

    root.deiconify()

    log("AUTOMATION COMPLETED")
    for ws in [ws_report, ws_fail]:

        for column in ws.columns:

            max_length = 0

            column_letter = column[0].column_letter

            for cell in column:

                try:

                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
                except:
                    pass

            ws.column_dimensions[
                column_letter
            ].width = max_length + 5

    wb_report.save(REPORT_FILE)
    os.startfile(REPORT_FILE)

# =========================================================
# STOP
# =========================================================

def stop_run():

    global running

    running = False

    log("STOPPED")
def load_universal_coordinates():
    """Loads the multi-sheet universal yard coordinate file (MAIN / CAL
    sheets) produced by the Universal Yard Coordinate Recorder. Pulls
    Menu/RED/YELLOW/GREEN for MAIN signals and Menu/Yellow for
    CALLING-ON signals - columns found by header name, wherever they
    sit. DOUBLE_YELLOW is ignored (this program only tracks 3 MAIN
    aspects: RED/YELLOW/GREEN)."""

    global main_signal_data, signal_data

    if not excel_data:
        messagebox.showwarning(
            "Load Excel First",
            "Please load the SPL_CASE_CALLING_ON Excel file first."
        )
        log("Please Load Excel First")
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

    loaded_main = 0
    loaded_cal = 0

    # ---- MAIN ----
    if "MAIN" in wb.sheetnames:
        ws = wb["MAIN"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        aspects_col = hi.get("ASPECTS")
        mx, my = hi.get("MENU_X"), hi.get("MENU_Y")
        rx, ry = hi.get("RED_X"), hi.get("RED_Y")
        yx, yy = hi.get("YELLOW_X"), hi.get("YELLOW_Y")
        gx, gy = hi.get("GREEN_X"), hi.get("GREEN_Y")

        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                main_signal_data[sig] = {
                    "menu": parse_point(cell(row, mx), cell(row, my)),
                    "red": parse_point(cell(row, rx), cell(row, ry)),
                    "yellow": parse_point(cell(row, yx), cell(row, yy)),
                    "green": parse_point(cell(row, gx), cell(row, gy)),
                    "aspect_count": parse_int(cell(row, aspects_col)) or 3
                }
                loaded_main += 1

    # ---- CALLING-ON ----
    if "CAL" in wb.sheetnames:
        ws = wb["CAL"]
        hi = header_index_of(ws)
        sig_col = hi.get("SIGNAL")
        mx, my = hi.get("MENU_X"), hi.get("MENU_Y")
        yx, yy = hi.get("YELLOW_X"), hi.get("YELLOW_Y")

        if sig_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or cell(row, sig_col) is None:
                    continue
                sig = str(cell(row, sig_col)).strip().upper()
                signal_data[sig] = {
                    "menu": parse_point(cell(row, mx), cell(row, my)),
                    "yellow": parse_point(cell(row, yx), cell(row, yy))
                }
                loaded_cal += 1

    log(f"UNIVERSAL YARD COORDINATES LOADED : {loaded_main} MAIN, {loaded_cal} CALLING-ON")
# =========================================================
# GUI
# =========================================================

root = tk.Tk()

root.title(
    "SPECIAL CASE CALLING ON SIGNAL AUTOMATION"
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

    text="SPECIAL CASE CALLING ON SIGNAL AUTOMATION",

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
    "TOC/RCC",
    load_excel,
    "#2563eb"
).pack(pady=8)

create_button(
    "LOAD YARD COORDINATES",
    load_universal_coordinates,
    "#7c3aed"
).pack(pady=8)

create_button(
    "START TESTING",
    start_run,
    "#16a34a"
).pack(pady=20)

create_button(
    "STOP TESTING",
    stop_run,
    "#dc2626"
).pack(pady=8)

# =========================================================
# STATUS BOX
# =========================================================



# ADD STATUS FRAME HERE
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
).pack(pady=10)

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

    text="CALLING ON ROUTE DETAILS",

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
        "MAIN_SIGNAL",
        "MAIN_ROUTE",
        "SIGNAL",
        "ROUTE",
        "TRACK"
    ),
    show="headings",
    yscrollcommand=tree_scroll.set
)

tree.heading("NO", text="NO")

tree.heading(
    "MAIN_SIGNAL",
    text="MAIN SIGNAL"
)

tree.heading(
    "MAIN_ROUTE",
    text="MAIN ROUTE"
)

tree.heading(
    "SIGNAL",
    text="CALLING ON"
)

tree.heading(
    "ROUTE",
    text="CALLING ON ROUTE"
)

tree.heading(
    "TRACK",
    text="TRACK"
)

tree.column("NO", width=80, anchor="center")
tree.column("SIGNAL", width=220, anchor="center")
tree.column("ROUTE", width=300, anchor="center")
tree.column("TRACK", width=220, anchor="center")

tree.pack(fill="both", expand=True)

tree_scroll.config(command=tree.yview)

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

    text="SPACE KEY = CAPTURE COORDINATES | SPECIAL CASE CALLING ON SIGNAL AUTOMATION",

    bg="#0f172a",

    fg="white",

    font=("Segoe UI", 10)

)

footer.pack(
    fill="x"
)

# =========================================================
# RUN
# =========================================================

root.mainloop()
