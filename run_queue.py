"""Run selected Hitachi C-FAT tests sequentially through the EDRC auto-feed helper."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent

DRIVERS = {
    "POINT_LC_CH_SDG_TEST": {
        "script": "POINT_LC_CH_SDG_TEST.py",
        "list": "load_excel",
        "coords": "load_universal_coordinates",
        "start": "start_automation",
    },
    "EMERGENCY_POINT_KEY_TEST": {
        "script": "EMERGENCY_POINT_KEY_TEST.py",
        "list": "load_excel",
        "coords": "load_universal_coordinates",
        "start": "start_run",
    },
    "CASCADING": {
        "script": "CASCADING.py",
        "list": "load_cascading_routes",
        "coords": "load_universal_coordinates",
        "start": "start_automation",
    },
    "SIGNAL_CLEARANCE": {
        "script": "signal_clearance.py",
        "list": "import_master_toc",
        "coords": "master_load_config",
        "start": "start_run",
    },
}


def write_status(path: Path, **values) -> None:
    current = {}
    if path.exists():
        try:
            current = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    current.update(values)
    current["updated_at"] = datetime.now().isoformat(timespec="seconds")
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(current, indent=2), encoding="utf-8")
    temporary.replace(path)


def main(config_path: str) -> int:
    config = json.loads(Path(config_path).read_text(encoding="utf-8"))
    tests = config["tests"]
    toc = str(Path(config["toc"]).resolve())
    coords = str(Path(config["coords"]).resolve())
    session_dir = Path(config["session_dir"]).resolve()
    status_path = Path(config["status_path"]).resolve()
    suite_path = APP_DIR / "EDRC_AUTOMATION_SUITE(1)_OEM_LOGOS(1).py"
    helper_path = APP_DIR / "_edrc_autofeed.py"

    # Extract the existing helper embedded in the suite without importing/running the suite UI.
    import ast
    tree = ast.parse(suite_path.read_text(encoding="utf-8-sig"), filename=str(suite_path))
    helper_source = None
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "AUTOFEED_SOURCE"
            for target in node.targets
        ):
            candidate = ast.literal_eval(node.value)
            if isinstance(candidate, str) and "_begin_autorun" in candidate:
                helper_source = candidate
                break
    if not helper_source:
        write_status(status_path, state="Failed", error="AUTOFEED_SOURCE not found in suite.")
        return 2

    # Preserve the existing diagnostic enhancement from the original frontend.
    old_exec = '''    with open(target, "r", encoding="utf-8") as _fh:
        _code = compile(_fh.read(), target, "exec")

    exec(_code, _G)
'''
    new_exec = '''    _say("HELPER BOOTED")
    _say("TARGET SCRIPT: %s" % target)
    _say("PYTHON: %s" % sys.executable)
    _say("WORKING DIRECTORY: %s" % os.getcwd())
    _say("PROGRAM KEY: %s | UNATTENDED: %s" % (_program_key, _UNATTENDED))
    _say("DRIVER PROFILE: %s" % (str(_profile),))
    try:
        with open(target, "r", encoding="utf-8") as _fh:
            _code = compile(_fh.read(), target, "exec")
        _say("BACKEND SOURCE READ; EXECUTING BACKEND")
        exec(_code, _G)
        _say("BACKEND MAINLOOP EXITED")
    except BaseException as exc:
        import traceback
        _say("FATAL BACKEND STARTUP ERROR: %s: %s" % (type(exc).__name__, exc))
        traceback.print_exc()
        raise
'''
    if old_exec in helper_source:
        helper_source = helper_source.replace(old_exec, new_exec, 1)
    helper_path.write_text(helper_source, encoding="utf-8")

    results = []
    write_status(status_path, state="Running", selected_tests=tests, current_test=None,
                 completed=0, total=len(tests), results=results, error=None)
    for index, test_key in enumerate(tests, start=1):
        spec = DRIVERS.get(test_key)
        if spec is None:
            results.append({"test": test_key, "status": "Skipped", "reason": "No verified driver profile."})
            write_status(status_path, results=results, completed=index, current_test=None)
            continue

        backend = APP_DIR / spec["script"]
        if not backend.is_file():
            results.append({"test": test_key, "status": "Failed", "reason": f"Missing source: {backend.name}"})
            write_status(status_path, results=results, completed=index, current_test=test_key)
            continue

        program_dir = session_dir / f"{index:02d}_{test_key}"
        program_dir.mkdir(parents=True, exist_ok=True)
        log_path = program_dir / "console.log"
        env = os.environ.copy()
        env.update({
            "EDRC_OEM": "HITACHI",
            "EDRC_SESSION_DIR": str(session_dir),
            "EDRC_PROGRAM_DIR": str(program_dir),
            "EDRC_PROGRAM": test_key,
            "EDRC_UNATTENDED": "1",
            "EDRC_LIST": toc,
            "EDRC_TOC": toc,
            "EDRC_TOC_FILE": toc,
            "EDRC_COORDS": coords,
            "EDRC_DRIVE_LIST": spec["list"],
            "EDRC_DRIVE_COORDS": spec["coords"],
            "EDRC_DRIVE_START": spec["start"],
        })
        for key in ("EDRC_DRIVE_CONFIG", "EDRC_DRIVE_TRACK_COORDS"):
            env.pop(key, None)

        write_status(status_path, state="Running", current_test=test_key,
                     current_index=index, current_log=str(log_path), current_program_dir=str(program_dir),
                     completed=index - 1, results=results)
        with log_path.open("w", encoding="utf-8", errors="replace") as log_handle:
            log_handle.write(
                f"[STREAMLIT] Sequential queue test {index}/{len(tests)}: {test_key}\n"
                f"[STREAMLIT] Backend: {backend}\n[STREAMLIT] TOC: {toc}\n"
                f"[STREAMLIT] Coordinates: {coords}\n"
            )
            log_handle.flush()
            try:
                child = subprocess.Popen(
                    [sys.executable, "-u", str(helper_path), str(backend)],
                    cwd=str(program_dir), env=env, stdout=log_handle,
                    stderr=subprocess.STDOUT,
                    creationflags=(subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0),
                )
                write_status(status_path, child_pid=child.pid)
                while child.poll() is None:
                    time.sleep(0.5)
                    # Stop request is shared through the queue status file.
                    try:
                        latest = json.loads(status_path.read_text(encoding="utf-8"))
                    except Exception:
                        latest = {}
                    if latest.get("stop_requested"):
                        if os.name == "nt":
                            subprocess.run(["taskkill", "/PID", str(child.pid), "/T", "/F"],
                                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                        else:
                            child.terminate()
                        child.wait(timeout=8)
                        break
                code = child.returncode
            except Exception as exc:
                code = 99
                log_handle.write(f"[STREAMLIT] Launcher error: {type(exc).__name__}: {exc}\n")

        try:
            latest = json.loads(status_path.read_text(encoding="utf-8"))
        except Exception:
            latest = {}
        if latest.get("stop_requested"):
            results.append({"test": test_key, "status": "Stopped", "exit_code": code,
                            "report_dir": str(program_dir)})
            write_status(status_path, state="Stopped", current_test=None,
                         completed=index, results=results)
            return 1

        result_state = "Completed" if code == 0 else "Failed"
        results.append({"test": test_key, "status": result_state, "exit_code": code,
                        "report_dir": str(program_dir)})
        write_status(status_path, state="Running", current_test=None,
                     completed=index, results=results)

    failures = [r for r in results if r["status"] != "Completed"]
    write_status(status_path, state="Completed with failures" if failures else "Completed",
                 current_test=None, completed=len(tests), results=results, child_pid=None)
    return 1 if failures else 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: run_queue.py <queue-config.json>")
    raise SystemExit(main(sys.argv[1]))
