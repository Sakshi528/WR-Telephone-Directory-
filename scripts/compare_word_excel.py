"""
Compares station rosters in the Word phone directory against the KMP source
workbook and reports, per station: employees in Word, employees in Excel,
missing employees, and missing control rooms.

Read-only: does not modify the Word file, the Excel file, the database, or
any import script. Writes reports/comparison_report.csv.

Word source: uploads/Western Region Phone Directory 2025 Main_Telephone.docx
Excel source: uploads/WR_DB_Ready_Final.xlsx

--- Word document structure notes (why the parsing below looks the way it does) ---

Most stations follow a simple pattern: a Heading paragraph, optionally an
address paragraph, then one table whose header row contains "Name" and
"Designation" columns, followed by employee rows and a "Control Room" row.

A few sections instead bundle MANY stations into a single large Word table,
using embedded "<N>. Station Name" rows as sub-station markers (e.g. the
"Black Start Facilitated Stations" table contains ~25 stations this way,
including the only listing for stations like Indira Sagar, Bargi, UKAI,
Kawas and Gandhar's black-start rosters). This script splits on those
numbered marker rows within a table, not just on Word headings.

Two stations (NTPC Kawas, NTPC Gandhar) appear twice in the document — once
under their own heading, once inside the Black Start table — so Word-side
results are grouped by resolved Excel organization identity (not by raw
heading text) to avoid double counting.

One section ("Synopsis of Important Telephone Numbers") is a front-matter
quick-reference that duplicates people covered in full elsewhere; it doesn't
match any real Excel organization name, so it naturally reports as a single
unmatched row rather than corrupting real per-station counts.

Known limitation: a couple of tables (e.g. the Maharashtra/MSETCL
transmission table) group multiple sub-offices using plain repeated-text
title blocks instead of numbered markers, and the header row itself repeats
mid-table as a Word print-pagination artifact. Those are treated as one
combined station rather than split further.
"""

import csv
import os
import re
import sys
from collections import Counter

import docx
import openpyxl
from docx.table import Table
from docx.text.paragraph import Paragraph

WORD_PATH = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"
EXCEL_PATH = "uploads/WR_DB_Ready_Final.xlsx"
OUTPUT_CSV = "reports/comparison_report.csv"

NUMBERED_TITLE_RE = re.compile(r"^\d+\.\s*[A-Za-z]")
LEADING_NUMBER_RE = re.compile(r"^\d+(\.\d+)*\.?\s*")


def clean(v):
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


def strip_leading_number(text):
    return LEADING_NUMBER_RE.sub("", text).strip()


# ── Excel loading ────────────────────────────────────────────────────────────

def load_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True)

    def data_rows(name):
        return list(wb[name].iter_rows(values_only=True))[1:]

    org_name_by_id = {}
    org_id_by_name_lower = {}
    for oid, name, _addr, _state in data_rows("organizations"):
        name = clean(name)
        if oid is None or not name:
            continue
        org_name_by_id[oid] = name
        org_id_by_name_lower[name.lower()] = oid

    suborg_name_by_id = {}
    suborg_parent_by_id = {}
    suborg_id_by_name_lower = {}
    for sid, parent_oid, name, _addr, _state in data_rows("sub_organizations"):
        name = strip_leading_number(clean(name))
        if sid is None or not name:
            continue
        suborg_name_by_id[sid] = name
        suborg_parent_by_id[sid] = parent_oid
        suborg_id_by_name_lower[name.lower()] = sid

    emp_by_suborg = {}
    emp_by_org_direct = {}
    for emp_id, org_id, suborg_id, name, _designation in data_rows("employees"):
        name = clean(name)
        if not name:
            continue
        if suborg_id is not None:
            emp_by_suborg.setdefault(suborg_id, []).append(name)
        elif org_id is not None:
            emp_by_org_direct.setdefault(org_id, []).append(name)

    cr_by_suborg = Counter()
    cr_by_org_direct = Counter()
    for row in data_rows("control_rooms"):
        _rid, org_id, suborg_id = row[0], row[1], row[2]
        if suborg_id is not None:
            cr_by_suborg[suborg_id] += 1
        elif org_id is not None:
            cr_by_org_direct[org_id] += 1

    return {
        "org_name_by_id": org_name_by_id,
        "org_id_by_name_lower": org_id_by_name_lower,
        "suborg_name_by_id": suborg_name_by_id,
        "suborg_id_by_name_lower": suborg_id_by_name_lower,
        "emp_by_suborg": emp_by_suborg,
        "emp_by_org_direct": emp_by_org_direct,
        "cr_by_suborg": cr_by_suborg,
        "cr_by_org_direct": cr_by_org_direct,
    }


# ── Word parsing ─────────────────────────────────────────────────────────────

def iter_block_items(doc):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def find_header_columns(rows_cells):
    """Return (name_col, designation_col) from the first row that looks like
    a roster header, or (None, None) if this table isn't a roster table."""
    for cells in rows_cells:
        name_col = next((i for i, c in enumerate(cells) if "name" in c.lower()), None)
        desig_col = next((i for i, c in enumerate(cells) if "designat" in c.lower()), None)
        if name_col is not None and desig_col is not None:
            return name_col, desig_col
    return None, None


def is_header_row(cells):
    lower = [c.lower() for c in cells]
    return any("name" in c for c in lower) and any("designat" in c for c in lower)


def parse_table_segments(table, fallback_name):
    """Split one Word table into (name, employee_set, has_control_room) segments."""
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    name_col, desig_col = find_header_columns(rows_cells)
    if name_col is None:
        return []  # not a roster table (reference/appendix table) — skip

    segments = []
    current_name = None
    current_employees = {}
    current_has_cr = False

    def flush():
        if current_employees or current_has_cr:
            segments.append((
                current_name or fallback_name,
                set(current_employees.values()),
                current_has_cr,
            ))

    for cells in rows_cells:
        if not cells:
            continue
        first_cell = cells[0]

        if NUMBERED_TITLE_RE.match(first_cell):
            flush()
            current_name = strip_leading_number(first_cell)
            current_employees = {}
            current_has_cr = False
            continue

        if is_header_row(cells):
            continue  # roster header, or a repeated print-pagination header

        if not any(cells):
            continue  # blank row

        name_cell = cells[name_col] if name_col < len(cells) else ""
        desig_cell = cells[desig_col] if desig_col < len(cells) else ""

        if not name_cell:
            continue

        if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
            current_has_cr = True
            continue

        if name_cell.lower() == desig_cell.lower():
            continue  # repeated title-block noise row (e.g. "GUJARAT" / "GUJARAT")

        current_employees[name_cell.lower()] = name_cell

    flush()
    return segments


def parse_word_segments(path):
    doc = docx.Document(path)
    current_heading = None
    all_segments = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue
        all_segments.extend(parse_table_segments(item, current_heading))
    return all_segments


# ── Matching Word segments to Excel organizations ───────────────────────────

def _best_substring_match(key_lower, name_lower_to_id):
    """Fallback for names like 'Kawas' that should match 'NTPC Kawas'. Picks
    the known name with the smallest length difference among those where one
    contains the other, e.g. prefers 'ntpc kawas' over the much longer
    'ntpc kawas solar pv project', which also contains 'kawas'."""
    if len(key_lower) < 4:
        return None
    best = None
    best_diff = None
    for known_lower, ident in name_lower_to_id.items():
        if known_lower in key_lower or key_lower in known_lower:
            diff = abs(len(known_lower) - len(key_lower))
            if best_diff is None or diff < best_diff:
                best, best_diff = ident, diff
    return best


def match_segment(name, excel):
    key_lower = strip_leading_number(name).lower()

    if key_lower in excel["suborg_id_by_name_lower"]:
        sid = excel["suborg_id_by_name_lower"][key_lower]
        return ("suborg", sid, excel["suborg_name_by_id"][sid])
    if key_lower in excel["org_id_by_name_lower"]:
        oid = excel["org_id_by_name_lower"][key_lower]
        return ("org", oid, excel["org_name_by_id"][oid])

    sid = _best_substring_match(key_lower, excel["suborg_id_by_name_lower"])
    if sid is not None:
        return ("suborg", sid, excel["suborg_name_by_id"][sid])
    oid = _best_substring_match(key_lower, excel["org_id_by_name_lower"])
    if oid is not None:
        return ("org", oid, excel["org_name_by_id"][oid])

    return ("none", key_lower, name)


def group_segments(segments, excel):
    groups = {}
    for name, employees, has_cr in segments:
        match_type, match_id, display_name = match_segment(name, excel)
        key = (match_type, match_id)
        g = groups.setdefault(key, {
            "display_name": display_name,
            "employees": {},
            "has_cr": False,
            "match_type": match_type,
            "match_id": match_id,
        })
        for emp in employees:
            g["employees"][emp.lower()] = emp
        g["has_cr"] = g["has_cr"] or has_cr
    return groups


def build_report_rows(groups, excel):
    rows = []
    for (match_type, match_id), g in groups.items():
        word_names = g["employees"]  # lower -> original

        if match_type == "suborg":
            excel_names = excel["emp_by_suborg"].get(match_id, [])
            excel_has_cr = excel["cr_by_suborg"].get(match_id, 0) > 0
        elif match_type == "org":
            excel_names = excel["emp_by_org_direct"].get(match_id, [])
            excel_has_cr = excel["cr_by_org_direct"].get(match_id, 0) > 0
        else:
            excel_names = []
            excel_has_cr = False

        excel_names_lower = {n.lower() for n in excel_names}
        missing = sorted(
            orig for lower, orig in word_names.items() if lower not in excel_names_lower
        )
        missing_cr = "Yes" if g["has_cr"] and not excel_has_cr else "No"

        station_name = g["display_name"]
        if match_type == "none":
            station_name = f"{station_name} (not found in Excel)"

        rows.append({
            "Station": station_name,
            "Employees in Word": len(word_names),
            "Employees in Excel": len(excel_names),
            "Missing Employees": "; ".join(missing) if missing else "None",
            "Missing Control Rooms": missing_cr,
        })

    rows.sort(key=lambda r: r["Station"].lower())
    return rows


def run(word_path, excel_path, output_csv):
    excel = load_excel(excel_path)
    segments = parse_word_segments(word_path)
    groups = group_segments(segments, excel)
    rows = build_report_rows(groups, excel)

    os.makedirs(os.path.dirname(output_csv), exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "Station", "Employees in Word", "Employees in Excel",
            "Missing Employees", "Missing Control Rooms",
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"Stations compared: {len(rows)}")
    print(f"Report written to: {output_csv}")


if __name__ == "__main__":
    run(WORD_PATH, EXCEL_PATH, OUTPUT_CSV)
