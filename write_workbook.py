"""
Ahsan Mushtaq — workbook writer (importer v1).

Takes the user's CURRENT EV Master Sheet (.xlsx) and appends parsed chat
rows, adding the two new columns Ahsan approved:
  - "Time"        inserted after Date        (col C)
  - "Remote/Local" inserted after Remarks    (inside the Survey Report group)

Group-header merges are rebuilt, header styles copied, existing data and
formatting untouched. Unparseable-but-data-bearing messages go on a
"Review" sheet — never guessed into rows.

The caller's file is never modified in place: pass a working copy path.
"""

import copy
import re
from datetime import datetime, date
import openpyxl
from openpyxl.utils import get_column_letter

# Final layout after the two column inserts (1-based)
COL = {
    "sr": 1, "date": 2, "time": 3, "client_name": 4, "address": 5,
    "city": 6, "client_cell": 7, "multiline_poc": 8, "vehicle": 9,
    "priority": 10, "status": 11, "charger_detail": 12,
    "surveyor_name": 13, "power_m": 14, "ground_m": 15, "bore": 16,
    "surveyor_exp": 17, "remarks": 18, "remote_local": 19,
    "team_name": 20, "hardware_cost": 21, "fuel_exp": 22,
    "team_exp": 23, "rider_exp": 24, "team_name2": 25, "exp2": 26,
}
HEADER_ROW, GROUP_ROW, FIRST_DATA_ROW = 3, 2, 4

# Group merges in the FINAL layout
FINAL_GROUPS = [
    ("D2:L2", "Client Information"),
    ("M2:T2", "Survey Report"),
    ("U2:X2", "Installation Detail"),
    ("Y2:Z2", "Complaint Cost"),
]

REVIEW_SHEET = "Review"


def _check_template(ws):
    h1 = ws.cell(HEADER_ROW, 1).value
    h2 = ws.cell(HEADER_ROW, 2).value
    if str(h1).strip() != "Sr. No" or str(h2).strip() != "Date":
        raise ValueError(
            "This does not look like the EV Master Sheet "
            f"(row 3 = {h1!r}, {h2!r}; expected 'Sr. No', 'Date').")


def _copy_style(src, dst):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy.copy(src.protection)


def _ensure_new_columns(ws):
    """Insert Time (col 3) and Remote/Local (after Remarks) once."""
    headers = [str(ws.cell(HEADER_ROW, c).value or "").strip()
               for c in range(1, ws.max_column + 1)]
    has_time = "Time" in headers
    has_rl = "Remote/Local" in headers
    if has_time and has_rl:
        return  # already upgraded (idempotent)

    # drop old merges, rebuild at the end
    for rng in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(rng))

    if not has_time:
        ws.insert_cols(COL["time"])
        ws.cell(HEADER_ROW, COL["time"]).value = "Time"
        _copy_style(ws.cell(HEADER_ROW, COL["date"]),
                    ws.cell(HEADER_ROW, COL["time"]))
        ws.column_dimensions[get_column_letter(COL["time"])].width = 9

    # Remarks sits at COL["remarks"] after the Time insert; Remote/Local
    # goes right after it (only if not already there from a past run)
    headers = [str(ws.cell(HEADER_ROW, c).value or "").strip()
               for c in range(1, ws.max_column + 1)]
    if "Remote/Local" not in headers:
        # find Remarks dynamically — robust to any prior state
        rcol = headers.index("Remarks") + 1
        ws.insert_cols(rcol + 1)
        ws.cell(HEADER_ROW, rcol + 1).value = "Remote/Local"
        _copy_style(ws.cell(HEADER_ROW, rcol),
                    ws.cell(HEADER_ROW, rcol + 1))
        ws.column_dimensions[get_column_letter(rcol + 1)].width = 12

    for rng, label in FINAL_GROUPS:
        ws.merge_cells(rng)
        cell_ref = re.match(r"([A-Z]+)(\d+)", rng).group(1) + "2"
        ws[cell_ref].value = label


def _last_data_row(ws):
    """Last row holding real data: a Date or a Client Name.

    (The template pre-fills Sr. No below the data — those bare numbers
    do not count as data and get overwritten by new rows.)
    """
    last = FIRST_DATA_ROW - 1
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if ws.cell(r, COL["date"]).value not in (None, "") or \
           ws.cell(r, COL["client_name"]).value not in (None, ""):
            last = r
    return last


def _write_review_sheet(wb, review):
    if REVIEW_SHEET in wb.sheetnames:
        wb.remove(wb[REVIEW_SHEET])
    ws = wb.create_sheet(REVIEW_SHEET)
    headers = ["Date", "Time", "Sender", "Message", "Needs your eye — reason"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = openpyxl.styles.Font(bold=True)
    for r, item in enumerate(review, 2):
        ws.cell(r, 1, item["date"])
        ws.cell(r, 2, item["time"])
        ws.cell(r, 3, item["sender"])
        ws.cell(r, 4, item["message"])
        ws.cell(r, 5, item["reason"])
    widths = [12, 10, 20, 80, 45]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w


def append_rows(master_copy_path, rows, review, out_path,
               date_format="DMY"):
    """Append parsed rows to a working copy of the master sheet.

    Returns (rows_written, review_count).
    """
    wb = openpyxl.load_workbook(master_copy_path)
    ws = wb.active
    _check_template(ws)
    _ensure_new_columns(ws)

    date_fmt = "DD/MM/YYYY" if date_format == "DMY" else "MM/DD/YYYY"
    last = _last_data_row(ws)
    sr = 0
    for r in range(FIRST_DATA_ROW, last + 1):
        v = ws.cell(r, COL["sr"]).value
        if isinstance(v, (int, float)) and v > sr:
            sr = int(v)
    first_new = last + 1
    written = 0
    for row in rows:
        last += 1
        sr += 1
        ws.cell(last, COL["sr"], sr)
        dcell = ws.cell(last, COL["date"])
        if isinstance(row["date"], date):
            dcell.value = datetime(row["date"].year, row["date"].month,
                                   row["date"].day)
            dcell.number_format = date_fmt
        else:
            dcell.value = ""
        ws.cell(last, COL["time"], row["time"])
        ws.cell(last, COL["client_name"], row["client_name"])
        ws.cell(last, COL["address"], row["address"])
        ws.cell(last, COL["city"], row["city"])
        ws.cell(last, COL["client_cell"], row["client_cell"])
        ws.cell(last, COL["vehicle"], row["vehicle"])
        ws.cell(last, COL["status"], row["status"])
        if row["power_m"] is not None:
            ws.cell(last, COL["power_m"], row["power_m"])
        if row["ground_m"] is not None:
            ws.cell(last, COL["ground_m"], row["ground_m"])
        ws.cell(last, COL["bore"], row["bore"])
        ws.cell(last, COL["remarks"], row["remarks"])
        ws.cell(last, COL["remote_local"], row["remote_local"])
        written += 1

    # clear stale pre-filled Sr. No below the new data (bare numbers whose
    # row has no date and no client — template scaffolding, not data)
    for r in range(last + 1, ws.max_row + 1):
        if ws.cell(r, COL["date"]).value in (None, "") and \
           ws.cell(r, COL["client_name"]).value in (None, ""):
            if ws.cell(r, COL["sr"]).value not in (None, ""):
                ws.cell(r, COL["sr"]).value = None
        else:
            break

    _write_review_sheet(wb, review)
    wb.save(out_path)
    return written, len(review)
