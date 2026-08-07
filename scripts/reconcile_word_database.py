"""Reconciles the PostgreSQL database against the Word phone directory, per
station: employees, control rooms, and switchyards. Read-only throughout.

Matches Word stations against the database's Organization table (a flat
name namespace with parent_id, not the workbook's two-tier org/suborg
split), so segmentation is extended locally with switchyard detection
rather than reusing compare_word_excel.parse_table_segments as-is.

Control rooms and switchyards are compared as actual DirectoryNumber
records (name, phone, email), not as a per-station present/absent flag --
a station is only "Perfect" when every record matches by name, normalized
phone digits, and email, with no missing/extra/duplicate records on either
side. Employee reconciliation is unchanged.

Writes reports/directory_reconciliation_report.csv and reports/perfect_stations.csv.

Usage:
  python scripts/reconcile_word_database.py
"""

import csv
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx
from docx.text.paragraph import Paragraph

from scripts.compare_word_excel import (
    clean,
    strip_leading_number,
    iter_block_items,
    find_header_columns,
    is_header_row,
    NUMBERED_TITLE_RE,
    SWITCHYARD_RE,
    classify_row_content,
    normalize_phone_set,
    normalize_email_set,
    _best_substring_match,
)

from app import app
from models.organization import Organization
from models.employee import Employee
from models.directory_number import DirectoryNumber

WORD_PATH = "uploads/Western Region Phone Directory 2026 Main_TD.docx"
OUTPUT_CSV = "reports/directory_reconciliation_report.csv"
PERFECT_STATIONS_CSV = "reports/perfect_stations.csv"


# ─── WORD PARSING ────────────────────────────────────────────────────────────

def parse_table_segments_ext(table, fallback_name):
    """Same segmentation as compare_word_excel.parse_table_segments, extended
    to capture actual control-room/switchyard record data (label, phones,
    emails) instead of only a boolean flag. Returns (name, employee_set,
    cr_records, sw_records) tuples, where each record is
    {"label", "phones", "emails"}."""
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    name_col, desig_col = find_header_columns(rows_cells)
    if name_col is None:
        return []  # not a roster table

    segments = []
    current_name = None
    current_employees = {}
    current_cr = []
    current_sw = []

    def flush():
        if current_employees or current_cr or current_sw:
            segments.append((
                current_name or fallback_name,
                set(current_employees.values()),
                list(current_cr),
                list(current_sw),
            ))

    def build_record(cells, label):
        label_lower = label.lower()
        phones, emails, _others = classify_row_content(
            [c for c in cells if clean(c).lower() != label_lower]
        )
        return {"label": label, "phones": phones, "emails": emails}

    def pick_label(name_cell, desig_cell, bare_words):
        # prefers whichever cell is NOT just the bare category word (e.g.
        # "SLDC DNH" / "Control room" -> "SLDC DNH", not the generic term),
        # so distinctly-named rows aren't collapsed into a shared, useless
        # label that then falsely mismatches the database's specific name
        name_bare = name_cell.strip().lower() in bare_words
        desig_bare = desig_cell.strip().lower() in bare_words
        if desig_bare and not name_bare and name_cell.strip():
            return name_cell
        if name_bare and not desig_bare and desig_cell.strip():
            return desig_cell
        return name_cell if name_cell.strip() else desig_cell

    for cells in rows_cells:
        if not cells:
            continue
        first_cell = cells[0]

        if NUMBERED_TITLE_RE.match(first_cell):
            flush()
            current_name = strip_leading_number(first_cell)
            current_employees = {}
            current_cr = []
            current_sw = []
            continue

        if is_header_row(cells):
            continue

        if not any(cells):
            continue

        name_cell = cells[name_col] if name_col < len(cells) else ""
        desig_cell = cells[desig_col] if desig_col < len(cells) else ""

        if not name_cell:
            continue

        if SWITCHYARD_RE.search(name_cell) or SWITCHYARD_RE.search(desig_cell):
            label = pick_label(name_cell, desig_cell, {"switchyard", "switch yard"})
            current_sw.append(build_record(cells, label))
            continue

        if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
            label = pick_label(name_cell, desig_cell, {"control room", "control rooms"})
            current_cr.append(build_record(cells, label))
            continue

        if name_cell.lower() == desig_cell.lower():
            continue  # repeated title-block noise row

        current_employees[name_cell.lower()] = name_cell

    flush()
    return segments


def parse_word(path):
    doc = docx.Document(path)
    current_heading = None
    all_segments = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue
        all_segments.extend(parse_table_segments_ext(item, current_heading))
    return all_segments


# ─── DATABASE LOADING ────────────────────────────────────────────────────────

def load_db():
    org_id_by_name_lower = {}
    org_name_by_id = {}
    for o in Organization.query.all():
        org_id_by_name_lower[o.organization_name.strip().lower()] = o.id
        org_name_by_id[o.id] = o.organization_name

    emp_by_org = {}
    for e in Employee.query.all():
        if e.organization_id is not None:
            emp_by_org.setdefault(e.organization_id, []).append(e.employee_name)

    cr_by_org = defaultdict(list)
    sw_by_org = defaultdict(list)
    for dn in DirectoryNumber.query.all():
        if dn.organization_id is None:
            continue
        cat = (dn.category or "").strip().lower()
        record = {"name": dn.name, "phone": dn.phone_number, "email": dn.email}
        if cat == "control room":
            cr_by_org[dn.organization_id].append(record)
        elif cat == "switchyard":
            sw_by_org[dn.organization_id].append(record)

    return {
        "org_id_by_name_lower": org_id_by_name_lower,
        "org_name_by_id": org_name_by_id,
        "emp_by_org": emp_by_org,
        "cr_by_org": cr_by_org,
        "sw_by_org": sw_by_org,
    }


# ─── MATCHING WORD SEGMENTS TO DATABASE ORGANIZATIONS ───────────────────────

def match_to_db(name, db):
    key_lower = strip_leading_number(name).lower()
    if key_lower in db["org_id_by_name_lower"]:
        oid = db["org_id_by_name_lower"][key_lower]
        return oid, db["org_name_by_id"][oid]

    oid = _best_substring_match(key_lower, db["org_id_by_name_lower"])
    if oid is not None:
        return oid, db["org_name_by_id"][oid]

    return None, name


def _merge_record(by_label, rec):
    key = rec["label"].strip().lower()
    existing = by_label.get(key)
    if existing:
        existing["phones"] = existing["phones"] | set(rec["phones"])
        existing["emails"] = existing["emails"] | set(rec["emails"])
    else:
        by_label[key] = {
            "label": rec["label"],
            "phones": set(rec["phones"]),
            "emails": set(rec["emails"]),
        }


def group_segments(segments, db):
    groups = {}
    for name, employees, cr_records, sw_records in segments:
        org_id, display_name = match_to_db(name, db)
        key = org_id if org_id is not None else ("unmatched", strip_leading_number(name).lower())
        g = groups.setdefault(key, {
            "display_name": display_name,
            "org_id": org_id,
            "employees": {},
            "cr_by_label": {},
            "sw_by_label": {},
        })
        for emp in employees:
            g["employees"][emp.lower()] = emp
        for rec in cr_records:
            _merge_record(g["cr_by_label"], rec)
        for rec in sw_records:
            _merge_record(g["sw_by_label"], rec)
    return groups


# ─── CONTROL ROOM / SWITCHYARD RECORD COMPARISON ────────────────────────────

def compare_directory_records(station, category, cat_key, word_by_label, db_records, totals):
    """Compares Word's control-room/switchyard records for one station
    against the database's DirectoryNumber rows for the same category:
    record count and names (Missing/Extra), duplicate names in the
    database, and -- for names present on both sides -- phone numbers
    (normalized to digit runs) and email addresses. Returns (rows,
    has_issue)."""
    rows = []
    has_issue = False

    db_by_name = defaultdict(list)
    for rec in db_records:
        db_by_name[(rec["name"] or "").strip().lower()].append(rec)

    word_names = set(word_by_label.keys())
    db_names = set(db_by_name.keys())

    for name in sorted(word_names - db_names):
        rows.append({
            "Station": station, "Category": category, "Issue": "Missing",
            "Name": word_by_label[name]["label"],
            "Details": f"Word lists a {category.lower()} here; none found in database",
        })
        totals[f"missing_{cat_key}"] += 1
        has_issue = True

    for name in sorted(db_names - word_names):
        rows.append({
            "Station": station, "Category": category, "Issue": "Extra",
            "Name": db_by_name[name][0]["name"],
            "Details": f"Database has a {category.lower()} here; none found in Word directory",
        })
        totals[f"extra_{cat_key}"] += 1
        has_issue = True

    for name in sorted(db_names):
        records = db_by_name[name]
        if len(records) > 1:
            rows.append({
                "Station": station, "Category": category, "Issue": "Duplicate",
                "Name": records[0]["name"],
                "Details": f"{len(records)} {category.lower()} records share this name in the database",
            })
            totals[f"duplicate_{cat_key}"] += 1
            has_issue = True

    for name in sorted(word_names & db_names):
        word_rec = word_by_label[name]
        db_records_for_name = db_by_name[name]

        word_phones = frozenset().union(*(normalize_phone_set(p) for p in word_rec["phones"])) \
            if word_rec["phones"] else frozenset()
        db_phones = normalize_phone_set(" / ".join(r["phone"] or "" for r in db_records_for_name))

        if word_phones != db_phones:
            rows.append({
                "Station": station, "Category": category, "Issue": "Phone Mismatch",
                "Name": word_rec["label"],
                "Details": (f"Word phone digits {sorted(word_phones)} do not match database "
                            f"phone digits {sorted(db_phones)}"),
            })
            totals[f"phone_mismatch_{cat_key}"] += 1
            has_issue = True

        word_emails = frozenset().union(*(normalize_email_set(e) for e in word_rec["emails"])) \
            if word_rec["emails"] else frozenset()
        db_emails = normalize_email_set(" / ".join(r["email"] or "" for r in db_records_for_name))

        if word_emails != db_emails:
            rows.append({
                "Station": station, "Category": category, "Issue": "Email Mismatch",
                "Name": word_rec["label"],
                "Details": (f"Word email(s) {sorted(word_emails)} do not match database "
                            f"email(s) {sorted(db_emails)}"),
            })
            totals[f"email_mismatch_{cat_key}"] += 1
            has_issue = True

    return rows, has_issue


# ─── RECONCILIATION ──────────────────────────────────────────────────────────

def run():
    with app.app_context():
        db_data = load_db()

    segments = parse_word(WORD_PATH)
    all_groups = group_segments(segments, db_data)

    # unmatched Word sections (front-matter summaries, vendor/liaison lists)
    # aren't real stations in the database, so there's nothing to reconcile
    skipped_non_directory = [g["display_name"] for g in all_groups.values() if g["org_id"] is None]
    groups = {k: g for k, g in all_groups.items() if g["org_id"] is not None}

    rows = []
    perfect_station_names = []
    stations_checked = 0
    perfect_stations = 0
    stations_with_issues = 0
    totals = {
        "missing_emp": 0, "extra_emp": 0,
        "missing_cr": 0, "extra_cr": 0, "duplicate_cr": 0,
        "phone_mismatch_cr": 0, "email_mismatch_cr": 0,
        "missing_sw": 0, "extra_sw": 0, "duplicate_sw": 0,
        "phone_mismatch_sw": 0, "email_mismatch_sw": 0,
    }

    for g in groups.values():
        org_id = g["org_id"]
        station = g["display_name"]
        stations_checked += 1
        station_has_issue = False

        word_names = g["employees"]  # lower -> original
        db_names = db_data["emp_by_org"].get(org_id, [])
        db_names_lower = {n.strip().lower() for n in db_names if n}

        missing_emps = sorted(orig for lower, orig in word_names.items() if lower not in db_names_lower)
        extra_emps = sorted(n for n in db_names if n and n.strip().lower() not in word_names)

        for name in missing_emps:
            rows.append({
                "Station": station, "Category": "Employee", "Issue": "Missing",
                "Name": name, "Details": "In Word directory but not found in database",
            })
            totals["missing_emp"] += 1
            station_has_issue = True

        for name in extra_emps:
            rows.append({
                "Station": station, "Category": "Employee", "Issue": "Extra",
                "Name": name, "Details": "In database but not found in Word directory",
            })
            totals["extra_emp"] += 1
            station_has_issue = True

        cr_rows, cr_issue = compare_directory_records(
            station, "Control Room", "cr", g["cr_by_label"],
            db_data["cr_by_org"].get(org_id, []), totals,
        )
        rows.extend(cr_rows)
        station_has_issue = station_has_issue or cr_issue

        sw_rows, sw_issue = compare_directory_records(
            station, "Switchyard", "sw", g["sw_by_label"],
            db_data["sw_by_org"].get(org_id, []), totals,
        )
        rows.extend(sw_rows)
        station_has_issue = station_has_issue or sw_issue

        if station_has_issue:
            stations_with_issues += 1
        else:
            perfect_stations += 1
            perfect_station_names.append(station)

    rows.sort(key=lambda r: (r["Station"].lower(), r["Category"], r["Issue"]))

    os.makedirs(os.path.dirname(OUTPUT_CSV), exist_ok=True)
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Station", "Category", "Issue", "Name", "Details"])
        writer.writeheader()
        writer.writerows(rows)

    os.makedirs(os.path.dirname(PERFECT_STATIONS_CSV), exist_ok=True)
    with open(PERFECT_STATIONS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Station"])
        for name in perfect_station_names:
            writer.writerow([name])

    print("=" * 70)
    print("  Word <-> Database Reconciliation Summary")
    print("=" * 70)
    print(f"  Non-directory sections ignored : {len(skipped_non_directory)}")
    for name in skipped_non_directory:
        print(f"    - {name}")
    print(f"  Stations checked       : {stations_checked}")
    print(f"  Perfect stations       : {perfect_stations}")
    print(f"  Stations with issues   : {stations_with_issues}")
    print(f"  Missing employees      : {totals['missing_emp']}")
    print(f"  Extra employees        : {totals['extra_emp']}")
    print(f"  Missing control rooms  : {totals['missing_cr']}")
    print(f"  Extra control rooms    : {totals['extra_cr']}")
    print(f"  Duplicate control rooms: {totals['duplicate_cr']}")
    print(f"  Control room phone mismatches : {totals['phone_mismatch_cr']}")
    print(f"  Control room email mismatches : {totals['email_mismatch_cr']}")
    print(f"  Missing switchyards    : {totals['missing_sw']}")
    print(f"  Extra switchyards      : {totals['extra_sw']}")
    print(f"  Duplicate switchyards  : {totals['duplicate_sw']}")
    print(f"  Switchyard phone mismatches   : {totals['phone_mismatch_sw']}")
    print(f"  Switchyard email mismatches   : {totals['email_mismatch_sw']}")
    print("=" * 70)
    print(f"  Report written to: {OUTPUT_CSV}")
    print(f"  Perfect stations written to: {PERFECT_STATIONS_CSV}")


if __name__ == "__main__":
    run()
