from __future__ import annotations

import os
import queue
import shutil
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import time
from typing import Callable
import tkinter as tk
from tkinter import filedialog
import webbrowser
import webview

from flask import Flask, jsonify, render_template, request

DEFAULT_COLUMN_LABEL = "K"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

# Office constants used through COM
MSO_SHAPE_PICTURE = 13
MSO_SHAPE_LINKED_PICTURE = 11
MSO_FALSE = 0
XL_MOVE_AND_SIZE = 1
RPC_E_CALL_REJECTED = -2147418111
MIN_COLUMN_OVERLAP_RATIO = 0.25
MIN_CELL_SIZE_POINTS = 2.0

_GEN_PY_REPAIR_LOCK = threading.Lock()
_GEN_PY_REPAIRED = False
_OPEN_EXCEL_LOCK = threading.Lock()
_OPEN_EXCEL_APP = None


def resolve_resource_path(name: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / name
    return Path(__file__).resolve().with_name(name)


@dataclass
class WorkbookResult:
    file_name: str
    target_column: str
    backup_path: str
    resized_images: int
    pictures_found: int
    errors: int


@dataclass
class WorkbookTask:
    workbook_path: Path
    target_column_label: str
    target_column_index: int


def normalize_column_label(raw_value: str) -> str:
    label = raw_value.strip().upper()
    if not label:
        raise ValueError("Cot khong duoc de trong")
    if not label.isalpha():
        raise ValueError("Cot chi duoc gom cac ky tu A-Z")
    if len(label) > 3:
        raise ValueError("Cot qua dai. Vi du hop le: K, AA, XFD")
    return label


def column_label_to_index(raw_value: str) -> int:
    label = normalize_column_label(raw_value)
    value = 0
    for char in label:
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def is_excel_candidate(path: Path) -> bool:
    return (
        path.suffix.lower() in SUPPORTED_EXTENSIONS
        and not path.name.startswith("~$")
        and ".backup_" not in path.stem
    )


def create_backup_file(workbook_path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = workbook_path.with_name(
        f"{workbook_path.stem}.backup_{timestamp}{workbook_path.suffix}"
    )
    counter = 1
    while candidate.exists():
        candidate = workbook_path.with_name(
            f"{workbook_path.stem}.backup_{timestamp}_{counter}{workbook_path.suffix}"
        )
        counter += 1

    shutil.copy2(workbook_path, candidate)
    return candidate


def is_picture_shape(shape) -> bool:
    try:
        shape_type = int(shape.Type)
    except Exception:
        return False
    return shape_type in (MSO_SHAPE_PICTURE, MSO_SHAPE_LINKED_PICTURE)


def shape_column_overlap_ratio(shape, column_left: float, column_right: float) -> float:
    shape_left = float(com_retry(lambda: shape.Left))
    shape_width = float(com_retry(lambda: shape.Width))
    if shape_width <= 0:
        return 0.0

    shape_right = shape_left + shape_width
    overlap_width = min(shape_right, column_right) - max(shape_left, column_left)
    if overlap_width <= 0:
        return 0.0

    return overlap_width / shape_width


def select_target_row_by_vertical_overlap(
    worksheet,
    target_column_index: int,
    top_row: int,
    bottom_row: int,
    shape_top: float,
    shape_bottom: float,
) -> int:
    best_row = max(1, top_row)
    best_overlap = -1.0

    for row in range(max(1, top_row), max(1, bottom_row) + 1):
        row_cell = com_retry(lambda row=row: worksheet.Cells(row, target_column_index))
        row_top = float(com_retry(lambda row_cell=row_cell: row_cell.Top))
        row_height = float(com_retry(lambda row_cell=row_cell: row_cell.Height))
        if row_height <= 0:
            continue

        row_bottom = row_top + row_height
        overlap = min(shape_bottom, row_bottom) - max(shape_top, row_top)
        if overlap > best_overlap:
            best_overlap = overlap
            best_row = row

    return best_row


def close_running_excel_instances(logger: Callable[[str], None]) -> int:
    closed_instances = 0

    try:
        import pythoncom
        import win32com.client as win32
        from win32com.client import dynamic as win32_dynamic
    except Exception:
        return 0

    # Repeatedly attach to active Excel, save/close workbooks, then quit.
    for _ in range(20):
        try:
            excel_app = com_retry(
                lambda: win32_dynamic.Dispatch(pythoncom.GetActiveObject("Excel.Application"))
            )
        except Exception:
            break

        try:
            wb_count = int(com_retry(lambda: excel_app.Workbooks.Count))
            if wb_count > 0:
                logger(f"[SYS] Saving and closing {wb_count} workbook(s) in running Excel...")

            for index in range(wb_count, 0, -1):
                workbook = com_retry(lambda index=index: excel_app.Workbooks(index))
                workbook_name = str(com_retry(lambda workbook=workbook: workbook.Name))

                try:
                    com_retry(lambda workbook=workbook: workbook.Save())
                except Exception:
                    # Continue closing even if one workbook cannot be saved.
                    pass

                com_retry(lambda workbook=workbook: workbook.Close(SaveChanges=True))
                logger(f"[SYS] Closed: {workbook_name}")

            com_retry(lambda: excel_app.Quit())
            closed_instances += 1
            time.sleep(0.15)
        except Exception as exc:
            logger(f"[WARN] Failed while closing running Excel instance: {exc}")
            break

    # Fallback: attach using Dispatch for cases where ActiveObject is unavailable.
    if closed_instances == 0:
        try:
            excel_app = com_retry(lambda: win32.Dispatch("Excel.Application"))
            wb_count = int(com_retry(lambda: excel_app.Workbooks.Count))

            if wb_count > 0:
                logger(f"[SYS] Saving and closing {wb_count} workbook(s) in running Excel...")
                for index in range(wb_count, 0, -1):
                    workbook = com_retry(lambda index=index: excel_app.Workbooks(index))
                    workbook_name = str(com_retry(lambda workbook=workbook: workbook.Name))

                    try:
                        com_retry(lambda workbook=workbook: workbook.Save())
                    except Exception:
                        pass

                    com_retry(lambda workbook=workbook: workbook.Close(SaveChanges=True))
                    logger(f"[SYS] Closed: {workbook_name}")

                com_retry(lambda: excel_app.Quit())
                closed_instances = 1
            else:
                # If Dispatch created a blank instance, close it immediately.
                com_retry(lambda: excel_app.Quit())
        except Exception as exc:
            logger(f"[WARN] Dispatch fallback could not close running Excel: {exc}")

    return closed_instances


def close_managed_excel_instance(logger: Callable[[str], None]) -> int:
    global _OPEN_EXCEL_APP

    with _OPEN_EXCEL_LOCK:
        if _OPEN_EXCEL_APP is None:
            return 0
        excel_app = _OPEN_EXCEL_APP
        _OPEN_EXCEL_APP = None

    try:
        wb_count = int(
            com_retry(
                lambda: excel_app.Workbooks.Count,
                attempts=60,
                delay_seconds=0.25,
            )
        )
    except Exception:
        wb_count = 0

    if wb_count > 0:
        logger(f"[SYS] Saving and closing {wb_count} workbook(s) from previous session...")

    for index in range(wb_count, 0, -1):
        try:
            workbook = com_retry(
                lambda index=index: excel_app.Workbooks(index),
                attempts=60,
                delay_seconds=0.25,
            )
            workbook_name = str(
                com_retry(
                    lambda workbook=workbook: workbook.Name,
                    attempts=60,
                    delay_seconds=0.25,
                )
            )

            try:
                com_retry(
                    lambda workbook=workbook: workbook.Save(),
                    attempts=60,
                    delay_seconds=0.25,
                )
            except Exception:
                pass

            com_retry(
                lambda workbook=workbook: workbook.Close(SaveChanges=True),
                attempts=60,
                delay_seconds=0.25,
            )
            logger(f"[SYS] Closed: {workbook_name}")
        except Exception as exc:
            logger(f"[WARN] Failed closing workbook from previous session: {exc}")

    try:
        com_retry(
            lambda: excel_app.Quit(),
            attempts=60,
            delay_seconds=0.25,
        )
    except Exception:
        pass

    return 1


def is_broken_gen_py_error(exc: Exception) -> bool:
    message = str(exc)
    return "win32com.gen_py" in message and "has no attribute" in message


def repair_win32com_gen_py_cache() -> bool:
    global _GEN_PY_REPAIRED

    with _GEN_PY_REPAIR_LOCK:
        if _GEN_PY_REPAIRED:
            return False

        try:
            import win32com
            from win32com.client import gencache
        except Exception:
            return False

        candidate_paths: list[Path] = []
        try:
            candidate_paths.append(Path(gencache.GetGeneratePath()))
        except Exception:
            pass

        try:
            candidate_paths.append(Path(win32com.__gen_path__))
        except Exception:
            pass

        seen: set[Path] = set()
        for candidate in candidate_paths:
            try:
                resolved = candidate.resolve()
            except Exception:
                resolved = candidate

            if resolved in seen:
                continue
            seen.add(resolved)

            if candidate.exists():
                shutil.rmtree(candidate, ignore_errors=True)

        try:
            gencache.is_readonly = False
        except Exception:
            pass

        try:
            gencache.Rebuild()
        except Exception:
            return False

        _GEN_PY_REPAIRED = True
        return True


def com_retry(action, attempts: int = 30, delay_seconds: float = 0.2):
    """Retry transient Excel COM busy errors and recover broken gen_py cache once."""
    last_error = None
    for _ in range(attempts):
        try:
            return action()
        except Exception as exc:
            if is_broken_gen_py_error(exc):
                if repair_win32com_gen_py_cache():
                    continue

            hresult = exc.args[0] if getattr(exc, "args", None) else None
            if hresult == RPC_E_CALL_REJECTED:
                last_error = exc
                time.sleep(delay_seconds)
                continue
            raise

    if last_error is not None:
        raise last_error

    raise RuntimeError("COM action failed without exception details")


def create_excel_application(prefer_new_instance: bool = True):
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "Thieu thu vien pywin32. Cai dat bang lenh: pip install pywin32"
        ) from exc

    primary_dispatch = win32.DispatchEx if prefer_new_instance else win32.Dispatch
    try:
        return com_retry(lambda: primary_dispatch("Excel.Application"))
    except Exception as exc:
        if not is_broken_gen_py_error(exc):
            raise

        # Last-resort late-binding path in case generated wrappers still fail.
        from win32com.client import dynamic as win32_dynamic

        return com_retry(lambda: win32_dynamic.Dispatch("Excel.Application"))


def fit_images_in_column(
    excel_app,
    workbook_path: Path,
    target_column_index: int,
    target_column_label: str,
    make_backup: bool,
) -> WorkbookResult:
    backup_path = "(skip backup)"
    if make_backup:
        backup_path = str(create_backup_file(workbook_path))

    workbook = None
    resized_images = 0
    pictures_found = 0
    errors = 0

    try:
        workbook = com_retry(
            lambda: excel_app.Workbooks.Open(
                str(workbook_path), UpdateLinks=0, ReadOnly=False
            )
        )

        for worksheet in com_retry(lambda: workbook.Worksheets):
            probe_cell = com_retry(lambda: worksheet.Cells(1, target_column_index))
            target_column_left = float(com_retry(lambda: probe_cell.Left))
            target_column_right = target_column_left + float(com_retry(lambda: probe_cell.Width))

            shape_count = int(com_retry(lambda: worksheet.Shapes.Count))
            for index in range(1, shape_count + 1):
                shape = com_retry(lambda: worksheet.Shapes.Item(index))
                if not is_picture_shape(shape):
                    continue

                try:
                    top_left_cell = com_retry(lambda: shape.TopLeftCell)
                    bottom_right_cell = com_retry(lambda: shape.BottomRightCell)

                    top_col = int(top_left_cell.Column)
                    bottom_col = int(bottom_right_cell.Column)
                    min_col = min(top_col, bottom_col)
                    max_col = max(top_col, bottom_col)

                    # Match by column span first, then fallback to real overlap check.
                    overlap_ratio = shape_column_overlap_ratio(
                        shape,
                        target_column_left,
                        target_column_right,
                    )
                    spans_target_column = min_col <= target_column_index <= max_col
                    if not spans_target_column and overlap_ratio < MIN_COLUMN_OVERLAP_RATIO:
                        continue

                    pictures_found += 1

                    shape_top = float(com_retry(lambda: shape.Top))
                    shape_height = float(com_retry(lambda: shape.Height))
                    shape_bottom = shape_top + max(0.0, shape_height)

                    top_row = int(top_left_cell.Row)
                    bottom_row = int(bottom_right_cell.Row)
                    if bottom_row < top_row:
                        bottom_row = top_row

                    # Choose the row with the largest vertical overlap to avoid row jumping.
                    target_row = select_target_row_by_vertical_overlap(
                        worksheet,
                        target_column_index,
                        top_row,
                        bottom_row,
                        shape_top,
                        shape_bottom,
                    )

                    target_cell = com_retry(
                        lambda: worksheet.Cells(target_row, target_column_index)
                    )

                    cell_left = float(com_retry(lambda: target_cell.Left))
                    cell_top = float(com_retry(lambda: target_cell.Top))
                    cell_width = float(com_retry(lambda: target_cell.Width))
                    cell_height = float(com_retry(lambda: target_cell.Height))

                    if cell_width < MIN_CELL_SIZE_POINTS or cell_height < MIN_CELL_SIZE_POINTS:
                        errors += 1
                        continue

                    # Fill the whole target cell.
                    com_retry(lambda: setattr(shape, "LockAspectRatio", MSO_FALSE))
                    com_retry(lambda: setattr(shape, "Placement", XL_MOVE_AND_SIZE))
                    com_retry(lambda: setattr(shape, "Left", cell_left))
                    com_retry(lambda: setattr(shape, "Top", cell_top))
                    com_retry(lambda: setattr(shape, "Width", cell_width))
                    com_retry(lambda: setattr(shape, "Height", cell_height))
                    resized_images += 1
                except Exception:
                    errors += 1

        com_retry(lambda: workbook.Save())
    finally:
        if workbook is not None:
            com_retry(lambda: workbook.Close(SaveChanges=False))

    return WorkbookResult(
        file_name=workbook_path.name,
        target_column=target_column_label,
        backup_path=backup_path,
        resized_images=resized_images,
        pictures_found=pictures_found,
        errors=errors,
    )


def process_workbooks(
    tasks: list[WorkbookTask],
    make_backup: bool,
    logger: Callable[[str], None],
) -> list[WorkbookResult]:
    excel_app = create_excel_application(prefer_new_instance=True)

    def set_app_property(name: str, value) -> None:
        try:
            com_retry(lambda: setattr(excel_app, name, value))
        except Exception:
            pass

    set_app_property("Visible", False)
    set_app_property("DisplayAlerts", False)
    set_app_property("ScreenUpdating", False)
    set_app_property("EnableEvents", False)

    results: list[WorkbookResult] = []
    try:
        for task in tasks:
            excel_file = task.workbook_path

            if not excel_file.exists():
                logger(
                    f"[{task.target_column_label}] {excel_file.name}: bo qua vi file khong ton tai"
                )
                continue

            if not is_excel_candidate(excel_file):
                logger(
                    f"[{task.target_column_label}] {excel_file.name}: bo qua vi khong phai file Excel hop le"
                )
                continue

            try:
                result = fit_images_in_column(
                    excel_app=excel_app,
                    workbook_path=excel_file,
                    target_column_index=task.target_column_index,
                    target_column_label=task.target_column_label,
                    make_backup=make_backup,
                )
                results.append(result)
                logger(
                    f"[{result.target_column}] {result.file_name}: resize={result.resized_images}, "
                    f"pictures={result.pictures_found}, errors={result.errors}"
                )
                logger(f"  backup: {result.backup_path}")
            except Exception as exc:
                logger(f"[{task.target_column_label}] {excel_file.name}: loi -> {exc}")
                results.append(
                    WorkbookResult(
                        file_name=excel_file.name,
                        target_column=task.target_column_label,
                        backup_path="(failed)",
                        resized_images=0,
                        pictures_found=0,
                        errors=1,
                    )
                )
    finally:
        com_retry(lambda: excel_app.Quit())

    return results


def summarize_results(results: list[WorkbookResult]) -> tuple[int, int, int]:
    total_resized = sum(r.resized_images for r in results)
    total_pictures = sum(r.pictures_found for r in results)
    total_errors = sum(r.errors for r in results)
    return total_resized, total_pictures, total_errors


# --- FLASK APP AND ROUTES ---

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    template_folder = os.path.join(sys._MEIPASS, "templates")
    static_folder = os.path.join(sys._MEIPASS, "static")
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)
log_queue = queue.Queue()


def web_logger(message: str):
    log_queue.put(message)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/select_files", methods=["POST"])
def api_select_files():
    # Use standard tkinter dialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    file_paths = filedialog.askopenfilenames(
        title="Chon file Excel",
        filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xls"), ("All files", "*.*")],
    )
    root.destroy()
    return jsonify({"files": list(file_paths)})


@app.route("/api/run", methods=["POST"])
def api_run():
    data = request.json
    make_backup = data.get("make_backup", True)
    raw_tasks = data.get("tasks", [])

    tasks = []
    for item in raw_tasks:
        try:
            path = Path(item["path"])
            col_label = normalize_column_label(item["column"])
            col_idx = column_label_to_index(col_label)
            tasks.append(WorkbookTask(path, col_label, col_idx))
        except Exception as exc:
            web_logger(f"Invalid column format for {item['path']}: {exc}")

    if not tasks:
        return jsonify({"status": "error", "message": "No valid tasks provided."})

    def run_worker():
        try:
            web_logger("Processing started...")
            results = process_workbooks(tasks, make_backup, web_logger)
            tr, tp, te = summarize_results(results)
            web_logger(f"DONE. Resized: {tr}, Pictures: {tp}, Errors: {te}")
        except Exception as exc:
            web_logger(f"Critical error: {exc}")

    thread = threading.Thread(target=run_worker, daemon=True)
    thread.start()
    # Keep request behavior unchanged: wait for completion so frontend reads final logs.
    thread.join()

    return jsonify(
        {
            "status": "success",
            "total_resized": 0,
            "total_pictures": 0,
            "total_errors": 0,
        }
    )


@app.route("/api/logs", methods=["GET"])
def api_logs():
    logs = []
    while not log_queue.empty():
        try:
            logs.append(log_queue.get_nowait())
        except queue.Empty:
            break
    return jsonify({"logs": logs})


@app.route("/api/open_file", methods=["POST"])
def api_open_file():
    data = request.json
    file_path = data.get("path")
    if not file_path or not os.path.exists(file_path):
        return jsonify({"status": "error", "message": "File path invalid or missing."})

    def open_worker():
        try:
            closed_managed = close_managed_excel_instance(web_logger)
            if closed_managed > 0:
                web_logger("[SYS] Closed previous Excel session opened by this tool.")

            closed_count = close_running_excel_instances(web_logger)
            if closed_count > 0:
                web_logger(f"[SYS] Closed {closed_count} running Excel instance(s).")

            excel_app = create_excel_application(prefer_new_instance=True)
        except RuntimeError as exc:
            web_logger(f"[ERROR] {exc}")
            return
        except Exception as exc:
            web_logger(f"[WARN] Failed to pre-close existing Excel session: {exc}")
            try:
                excel_app = create_excel_application(prefer_new_instance=True)
            except RuntimeError as inner_exc:
                web_logger(f"[ERROR] {inner_exc}")
                return

        try:
            web_logger(f"Launching workbook: {os.path.basename(file_path)}...")
            excel_app.Workbooks.Open(file_path)
            excel_app.Visible = True

            global _OPEN_EXCEL_APP
            with _OPEN_EXCEL_LOCK:
                _OPEN_EXCEL_APP = excel_app

            web_logger("[SYS] Workbook is now open & visible.")

        except Exception as exc:
            web_logger(f"[ERROR] COM Exception: {exc}")
            try:
                com_retry(lambda: excel_app.Quit())
            except Exception:
                pass

    threading.Thread(target=open_worker, daemon=True).start()
    return jsonify({"status": "success", "message": "Dispatched open command."})


if __name__ == "__main__":
    # Start the native desktop window wrapping the Flask app
    window = webview.create_window(
        "NEXUS COMMAND PRO", app, width=1280, height=800, background_color="#010405"
    )
    webview.start()
