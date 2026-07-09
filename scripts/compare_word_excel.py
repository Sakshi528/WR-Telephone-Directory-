"""Compares Word directory rosters against the KMP source workbook and
reports missing employees / control rooms per station. Read-only.
Writes reports/comparison_report.csv."""

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

# matches both "switchyard" and "switch yard" (Word uses both spellings)
SWITCHYARD_RE = re.compile(r"switch\s*yard", re.IGNORECASE)

# content-based (not column-position) classifiers for phone/email cell values
PHONE_RE = re.compile(r"\d{5,}")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")


def clean(v):
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


def strip_leading_number(text):
    return LEADING_NUMBER_RE.sub("", text).strip()


def dedup_preserve_order(seq):
    """Preserves first-seen order -- merged Word cells repeat their value
    across every grid cell they span, so the same value can appear 2-4
    times in one row's raw cell list."""
    seen = set()
    out = []
    for v in seq:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out


def classify_row_content(cells):
    """Splits a Word row's cells into (phones, emails, others) by what each
    value looks like, ignoring column position -- safe even when a row's
    merged-cell layout doesn't match its table header's column order."""
    phones, emails, others = [], [], []
    for raw in cells:
        v = clean(raw)
        if not v:
            continue
        if EMAIL_RE.search(v):
            emails.append(v)
        elif PHONE_RE.search(v):
            phones.append(v)
        else:
            others.append(v)
    return dedup_preserve_order(phones), dedup_preserve_order(emails), dedup_preserve_order(others)


def normalize_phone_set(text):
    """Digit runs of length >= 5 extracted from a phone string, as a
    frozenset -- tolerant of differing separators/formatting between
    sources (e.g. Word's "07323-284542/ 284719" vs the database's joined
    "07323-284542 / 284719")."""
    return frozenset(PHONE_RE.findall(text or ""))


def normalize_email_set(text):
    """Lowercased email addresses extracted from a string, as a frozenset."""
    return frozenset(e.lower() for e in EMAIL_RE.findall(text or ""))


# ─── EXCEL LOADING ───────────────────────────────────────────────────────────

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


# ─── WORD PARSING ────────────────────────────────────────────────────────────

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
            continue

        name_cell = cells[name_col] if name_col < len(cells) else ""
        desig_cell = cells[desig_col] if desig_col < len(cells) else ""

        if not name_cell:
            continue

        if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
            current_has_cr = True
            continue

        if name_cell.lower() == desig_cell.lower():
            continue  # e.g. "GUJARAT" / "GUJARAT" title-block noise row

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


# ─── MATCHING WORD SEGMENTS TO EXCEL ORGANIZATIONS ──────────────────────────

def _best_substring_match(key_lower, name_lower_to_id):
    """Fallback for names like 'Kawas' that should match 'NTPC Kawas' --
    picks the closest-length known name containing (or contained by) it."""
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
