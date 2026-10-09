HITACHI C-FAT STREAMLIT FRONTEND — MULTI-TEST SEQUENTIAL QUEUE

FILES
- streamlit_app.py                         Streamlit UI and multi-select queue launcher
- run_queue.py                             Sequential process supervisor
- POINT_LC_CH_SDG_TEST.py                  Existing Point / LC / CH / SDG backend (preserved)
- EMERGENCY_POINT_KEY_TEST.py               Emergency Point Key backend
- CASCADING.py                              Cascading backend
- signal_clearance.py                      Signal Clearance backend
- EDRC_AUTOMATION_SUITE(1)_OEM_LOGOS(1).py  Existing EDRC suite (AUTOFEED_SOURCE used by helper)

INSTALL
1. Put all project files in one folder.
2. Install packages required by the backend programs and Streamlit in the same Python environment.
3. Open a terminal in the project folder and run:
       python -m streamlit run streamlit_app.py
4. Upload the TOC workbook and the coordinate/configuration workbook.
5. Select one or more tests. They execute in the displayed selection order.
6. Confirm that the environment is an isolated simulator/C-FAT test panel, then choose Run selected tests.
7. Use Refresh status to update queue progress, logs, and reports. Each test gets its own subfolder under EDRC_REPORTS/HITACHI/SESSION_*/.

VERIFIED DRIVER SEQUENCES
- POINT_LC_CH_SDG_TEST: load_excel -> load_universal_coordinates -> start_automation
- EMERGENCY_POINT_KEY_TEST: load_excel -> load_universal_coordinates -> start_run
- CASCADING: load_cascading_routes -> load_universal_coordinates -> start_automation
- signal_clearance.py: import_master_toc -> master_load_config -> start_run

WORKBOOK NOTES
The same TOC and coordinate workbook are passed to each selected test. Ensure they meet every selected backend's sheet/header requirements:
- Emergency Point Key needs POINT and CH sheets plus SYSTEM_CONTROLS with a recorded EMERGENCY POINT KEY coordinate; TRACK_CONFIG is optional when external track coordinates are available.
- Signal Clearance expects a complete yard coordinate workbook with applicable MAIN, SHUNT, CAL and point/CH/LC sheets.
- Cascading expects a cascading TOC and a universal yard coordinate workbook.
- Point/LC/CH/SDG expects its existing supported TOC and universal coordinate workbook.

REPORTS AND LOGS
- Original backend report creation and report layout logic are not rewritten by the Streamlit frontend.
- Each backend runs from its own program subfolder so relative report paths remain isolated per test.
- Console logs are saved as console.log inside each test subfolder.

SAFETY
Run only in an isolated test/simulator environment. Do not connect this automation to live railway signalling equipment. This package has been statically checked; it has not been run against any railway panel or hardware.


RAIL HEADER MOTION FIX (2026-10-09)
- The train is now animated using explicit left-position keyframes from 100% to -100% of its own width.
- Removed the reduced-motion CSS override that could leave the train stationary.
- If the old header still appears, the running app/deployment is still serving the previous streamlit_app.py. Replace the deployed project files and restart/redeploy; then hard-refresh the browser (Ctrl+Shift+R).

Header alignment update (2026-10-09): The moving EDRC train asset now has a transparent exterior background and is confined to the bottom rail band. The left E2E Rail/signal and right NOVA images remain fixed above the train layer; the title and subtitle occupy a separate upper zone.

UI fixes (2026-10-09): no test is preselected on initial page load; choose tests manually before starting the queue. The header uses a separately positioned railway signal that cycles red, yellow, and green, so it no longer overlaps the E2E Rail logo.


HEADER POSITION UPDATE (2026-10-09)
- Increased Streamlit's top content padding so the complete rail header is visible below the top app chrome.
- Header assets and animation logic are otherwise unchanged.


HEADER UPDATE v12 (2026-10-09)
- Removed the (already hidden) dashed-track element and its CSS rule entirely; no dashed lines remain.
- Train enlarged slightly: max width 500px -> 560px (responsive: min(560px,54vw); 86vw on narrow screens), still fully visible and below the heading/subtitle.
- Signal fix: removed redundant animation-delay values that double-shifted the keyframes; lamps now cycle Red -> Yellow -> Green -> Red.


HEADER + BACKGROUND UPDATE v13 (2026-10-09)
- Page background is now assets/rail_background.jpg (user-supplied "South India Rail Push" image, title text removed, upscaled/sharpened). Keep the assets/ folder next to streamlit_app.py; if it is missing the app falls back to the old dark gradient.
- Header re-laid out like the reference video: heading row on top; below it the E2E Rail logo, a black hooded 3-lamp signal and the NOVA logo stay fixed. The train runs right to left inside a lane clipped between the signal and NOVA, so it emerges from behind NOVA and disappears behind the signal.
- Signal cycles Green -> Yellow -> Red, 0.5 s per lamp (1.5 s loop), as in the video.
- Content area gets a translucent dark panel and white headings for legibility over the photo.
- Top padding raised to 5rem so Streamlit's fixed toolbar no longer covers the top of the header.
