"""Direct station-by-station field diff against the Word directory --
treats Word as literal source of truth and flags every place the app's
text doesn't match it, not just the previously-checked issue types
(duplicated_phone/address_in_name/numbered_prefix/possible_duplicate_row).
Read-only -- writes reports/full_word_vs_app_field_diff.csv only.

Why a new parser instead of reusing reconcile_word_database.parse_word:
that function (and compare_word_excel.classify_row_content underneath it)
only ever extracts a name_col/desig_col pair for employees, and collapses
control-room/switchyard rows' remaining cells into deduped phone/email
TOKEN SETS -- it was built for presence/absence + digit-set comparison,
not a raw-text field-by-field diff. It also never captures the
organization heading text itself (the gap that missed "NLDC And RLDCs"
staying generic after the 5 individual rows were renamed) or per-employee
phone/mobile/email at all (Employee rows in Word DO have Telephone (O)/
Telephone (R)/Mobile/Email columns in most tables -- confirmed by
inspection -- but no existing script ever reads them).

Word table headers are NOT uniform -- 32 distinct header-row variants
exist across this one document (different S.No./Sl.No./Sr.No. spellings,
merged cells that duplicate a header's text across 2 grid columns, some
tables missing phone columns entirely, one with "Activity" instead of a
serial number, etc.). So columns are mapped by HEADER KEYWORD per table,
not by fixed position -- see FIELD_KEYWORDS.

Comparison rules (deliberately conservative about claiming "no match"):
  - organization name: Word section heading vs Organization.organization_name,
    compared case/whitespace-normalized.
  - employees: matched to Employee rows for that org by normalized name.
    Only fields Word's OWN table actually had a column for are compared
    (a table with no phone column contributes no phone mismatch --
    "column absent" is not "value absent").
  - control room / switchyard: matched to DirectoryNumber rows by
    normalized label. DirectoryNumber has one combined phone_number field
    (no office/mobile split), so a Word field is scored as matching if its
    digits (for phone-ish fields) or lowercased text (for email) appear
    somewhere within the app's combined field -- not exact equality --
    since the app is allowed to combine what Word split into columns.
  - address: Word's roster tables have no address column -- checked
    whether a free paragraph between the section heading and its table
    looks address-like; if none exists, reported as
    "no_word_address_data", not a false mismatch.
  - organizations with no Word match at all, and Word sections with no
    org match at all, are flagged as their own rows rather than skipped.

Usage:
  python scripts/full_word_vs_app_field_diff.py
"""

import csv
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

from app import app
from models.organization import Organization
from models.employee import Employee
from models.directory_number import DirectoryNumber

from scripts.compare_word_excel import (
    clean, strip_leading_number, iter_block_items, is_header_row,
    NUMBERED_TITLE_RE, SWITCHYARD_RE, PHONE_RE, EMAIL_RE, _best_substring_match,
)
from scripts.reconcile_word_database import looks_like_known_org

WORD_PATH = "uploads/Western Region Phone Directory 2026 Main_TD.docx"
OUT_CSV = "reports/full_word_vs_app_field_diff.csv"

FIELD_KEYWORDS = [
    ("email", ("email", "mail id", "e-mail")),
    ("mobile", ("mobile",)),
    ("residence_phone", ("telephone (r", "telephone(r", "telephone(r )", "residence")),
    ("office_phone", ("telephone (o", "telephone(o", "telephone no", "phone no", "office", "telephone")),
    ("designation", ("designat",)),
    ("department", ("department",)),
    ("name", ("name",)),
]


def map_header_columns(header_cells):
    """header keyword -> first matching column index (a merged header cell
    duplicates identical text across its grid columns, so 'first match' is
    safe -- not a silent data loss, see module docstring)."""
    field_by_col = {}
    for idx, cell in enumerate(header_cells):
        cell_lower = cell.lower()
        for field, keywords in FIELD_KEYWORDS:
            if field in field_by_col.values():
                continue
            if any(kw in cell_lower for kw in keywords):
                field_by_col.setdefault(idx, field)
                break
    # field -> column index (first occurrence)
    col_by_field = {}
    for idx, field in field_by_col.items():
        col_by_field.setdefault(field, idx)
    return col_by_field


def find_header_columns_generic(rows_cells):
    for cells in rows_cells:
        lower = [c.lower() for c in cells]
        if any("name" in c for c in lower) and any("designat" in c for c in lower):
            return map_header_columns(cells)
    return None


def parse_table_records(table, fallback_name, org_id_by_name_lower):
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    col_by_field = find_header_columns_generic(rows_cells)
    if col_by_field is None or "name" not in col_by_field:
        return []

    name_col = col_by_field["name"]
    desig_col = col_by_field.get("designation")

    segments = []
    current_name = None
    current_name_raw = None  # the untouched title text, e.g. "1. Indira Sagar" -- None
    # when this segment's name came from the outer document heading instead
    # of an in-table numbered title (see parse_word_with_address, which
    # falls back to the outer heading only in that case)
    current_employees = []
    current_cr = []
    current_sw = []

    def flush():
        if current_employees or current_cr or current_sw:
            segments.append((current_name or fallback_name, current_name_raw, list(current_employees), list(current_cr), list(current_sw)))

    for i, cells in enumerate(rows_cells):
        if not cells:
            continue
        first_cell = cells[0]

        if NUMBERED_TITLE_RE.match(first_cell):
            flush()
            current_name = strip_leading_number(first_cell)
            current_name_raw = strip_leading_number(first_cell)  # title text, leading "N." stripped same as the station name itself (that numbering is section-list numbering, not part of the name -- consistent with every other script in this repo)
            current_employees, current_cr, current_sw = [], [], []
            continue

        # "banner" title row: every non-empty cell repeats the same text
        # (a merged title cell spanning the full row width), no leading
        # number -- e.g. "POWER GRID CORPORATION OF INDIA LIMITED" repeated
        # across all 7 columns. NUMBERED_TITLE_RE alone misses these
        # entirely (confirmed: neither this script's first draft nor the
        # existing compare_word_excel.parse_table_segments /
        # reconcile_word_database.parse_table_segments_ext handle this
        # pattern -- they only check NUMBERED_TITLE_RE too, so multiple
        # real stations sharing one table silently got merged under
        # whichever numbered/outer heading came before them, in every
        # Word-parsing script in this repo, not just this one). Confirmed
        # via lookahead (next row is a column-header row) rather than on
        # sight alone, since decorative quote banners ("Talent wins
        # games...") use the exact same all-cells-identical shape and
        # would otherwise be misread as station titles.
        non_empty = [c for c in cells if c]
        if len(non_empty) >= 2 and len(set(non_empty)) == 1 and not is_header_row(cells):
            lookahead = rows_cells[i + 1:i + 3]
            # a lookahead header row alone is NOT enough on its own -- also
            # confirmed for this same shape on an address-only subtitle row
            # AND a decorative quote sitting mid-table between two real
            # employees of the SAME station (see reconcile_word_database.
            # looks_like_known_org's docstring for the concrete examples,
            # found and fixed there first) -- both would otherwise silently
            # steal every following row into a pseudo-segment matching no
            # real organization, orphaning that station's real data.
            if any(is_header_row(r) for r in lookahead if r) and looks_like_known_org(non_empty[0], org_id_by_name_lower):
                flush()
                title_text = non_empty[0]
                current_name = strip_leading_number(title_text)
                current_name_raw = strip_leading_number(title_text)
                current_employees, current_cr, current_sw = [], [], []
            continue  # banner row itself is never a data row either way

        if is_header_row(cells):
            continue
        if not any(cells):
            continue

        name_cell = cells[name_col] if name_col < len(cells) else ""
        desig_cell = cells[desig_col] if desig_col is not None and desig_col < len(cells) else ""
        if not name_cell:
            continue
        if name_cell.lower() == desig_cell.lower() and desig_cell:
            continue  # title-block noise row

        record = {}
        for field, col in col_by_field.items():
            if col < len(cells):
                record[field] = cells[col]

        if SWITCHYARD_RE.search(name_cell) or SWITCHYARD_RE.search(desig_cell):
            label = name_cell if name_cell.strip().lower() not in ("switchyard", "switch yard") else desig_cell
            record["label"] = label or name_cell
            current_sw.append(record)
            continue
        if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
            label = name_cell if name_cell.strip().lower() not in ("control room", "control rooms") else desig_cell
            record["label"] = label or name_cell
            current_cr.append(record)
            continue

        current_employees.append(record)

    flush()
    return segments


def parse_word_with_address(path, org_id_by_name_lower):
    """Same heading-tracking as compare_word_excel.parse_word_segments, but
    also captures the first non-empty free paragraph directly between a
    heading and its table (a candidate address line), and returns raw
    (un-stripped) heading text alongside the stripped station name."""
    doc = docx.Document(path)
    current_heading_raw = None
    current_heading = None
    pending_address = None
    all_segments = []
    heading_raw_by_name = {}
    address_by_name = {}

    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            text = clean(item.text)
            if item.style.name.startswith("Heading") and text:
                current_heading_raw = text
                current_heading = strip_leading_number(text)
                pending_address = None
            elif text and current_heading and pending_address is None:
                pending_address = text
            continue

        segments = parse_table_records(item, current_heading, org_id_by_name_lower)
        for name, name_raw, employees, cr, sw in segments:
            heading_raw_by_name.setdefault(name, name_raw or current_heading_raw or name)
            if pending_address:
                address_by_name.setdefault(name, pending_address)
        all_segments.extend((name, employees, cr, sw) for name, _raw, employees, cr, sw in segments)

    return all_segments, heading_raw_by_name, address_by_name


# ─── matching Word segments to Organization rows ────────────────────────

def match_to_org(name, org_id_by_name_lower):
    key_lower = strip_leading_number(name).lower()
    if key_lower in org_id_by_name_lower:
        return org_id_by_name_lower[key_lower]
    return _best_substring_match(key_lower, org_id_by_name_lower)


def normalize_text(v):
    return re.sub(r"\s+", " ", (v or "").strip().lower())


def normalize_digits(v):
    return "".join(PHONE_RE.findall(v or ""))


def phone_field_matches(word_value, app_value):
    word_digits = normalize_digits(word_value)
    if not word_digits:
        return normalize_text(word_value) == "" or normalize_text(word_value) in normalize_text(app_value)
    return word_digits in normalize_digits(app_value)


def email_field_matches(word_value, app_value):
    return normalize_text(word_value) in normalize_text(app_value) if word_value else normalize_text(app_value) == ""


def group_segments(segments, org_id_by_name_lower):
    groups = defaultdict(lambda: {"employees": [], "cr": [], "sw": [], "names_seen": []})
    unmatched_word_sections = []
    for name, employees, cr_records, sw_records in segments:
        org_id = match_to_org(name, org_id_by_name_lower)
        if org_id is None:
            if employees or cr_records or sw_records:
                unmatched_word_sections.append(name)
            continue
        g = groups[org_id]
        g["employees"].extend(employees)
        g["cr"].extend(cr_records)
        g["sw"].extend(sw_records)
        g["names_seen"].append(name)
    return groups, unmatched_word_sections


# ─── diff builders ──────────────────────────────────────────────────────

def diff_org_name(rows, org, word_heading_raw):
    if word_heading_raw is None:
        return
    # strip a leading section number ("14.3. NTPC Mauda", "17.5 Sterlite...")
    # before comparing -- same numbered_prefix artifact already established
    # as inert document-section-numbering, not real name content, earlier
    # this session (13 DirectoryNumber.name rows stripped on that basis).
    # word_value still shows the untouched raw text either way.
    match = normalize_text(strip_leading_number(word_heading_raw)) == normalize_text(org.organization_name)
    rows.append({
        "organization": org.organization_name, "field": "organization_name",
        "word_value": word_heading_raw, "app_value": org.organization_name,
        "match": "yes" if match else "no",
    })


def diff_address(rows, org, word_address):
    if word_address is None:
        rows.append({
            "organization": org.organization_name, "field": "address",
            "word_value": "(no address paragraph found in Word for this section)",
            "app_value": org.address or "", "match": "no_word_address_data",
        })
        return
    match = normalize_text(word_address) == normalize_text(org.address)
    rows.append({
        "organization": org.organization_name, "field": "address",
        "word_value": word_address, "app_value": org.address or "",
        "match": "yes" if match else "no",
    })


def diff_employees(rows, org, word_employees, app_employees):
    word_by_name = {}
    for rec in word_employees:
        n = normalize_text(rec.get("name", ""))
        if n:
            word_by_name.setdefault(n, rec)
    app_by_name = {normalize_text(e.employee_name): e for e in app_employees if e.employee_name}

    for name_key, rec in word_by_name.items():
        emp = app_by_name.get(name_key)
        if emp is None:
            rows.append({
                "organization": org.organization_name, "field": "(employee row)",
                "word_value": rec.get("name", ""), "app_value": "(no matching employee found in app)",
                "match": "word_only",
            })
            continue
        field_pairs = [
            ("designation", rec.get("designation"), emp.designation),
            ("office_phone", rec.get("office_phone"), emp.office_phone),
            ("mobile", rec.get("mobile"), emp.mobile_phone),
            ("email", rec.get("email"), emp.email),
        ]
        for field, word_val, app_val in field_pairs:
            if field not in rec:
                continue  # this Word table had no column for this field -- not comparable
            is_phone = field in ("office_phone", "mobile")
            matched = phone_field_matches(word_val, app_val) if is_phone else \
                (email_field_matches(word_val, app_val) if field == "email" else normalize_text(word_val) == normalize_text(app_val))
            rows.append({
                "organization": org.organization_name, "field": f"employee[{emp.employee_name}].{field}",
                "word_value": word_val or "", "app_value": app_val or "",
                "match": "yes" if matched else "no",
            })

    for name_key, emp in app_by_name.items():
        if name_key not in word_by_name:
            rows.append({
                "organization": org.organization_name, "field": "(employee row)",
                "word_value": "(not found in Word for this section)", "app_value": emp.employee_name,
                "match": "app_only",
            })


def diff_cr_sw(rows, org, word_records, app_records, category_label):
    word_by_label = {}
    for rec in word_records:
        n = normalize_text(rec.get("label", ""))
        if n:
            word_by_label.setdefault(n, rec)
    app_by_label = {normalize_text(dn.name): dn for dn in app_records if dn.name}

    for label_key, rec in word_by_label.items():
        dn = app_by_label.get(label_key)
        if dn is None:
            rows.append({
                "organization": org.organization_name, "field": f"({category_label} row)",
                "word_value": rec.get("label", ""), "app_value": "(no matching record found in app)",
                "match": "word_only",
            })
            continue
        for field in ("office_phone", "mobile", "email"):
            if field not in rec:
                continue
            word_val = rec.get(field)
            app_val = dn.phone_number if field in ("office_phone", "mobile") else dn.email
            matched = phone_field_matches(word_val, app_val) if field != "email" else email_field_matches(word_val, app_val)
            rows.append({
                "organization": org.organization_name, "field": f"{category_label}[{dn.name}].{field}",
                "word_value": word_val or "", "app_value": app_val or "",
                "match": "yes" if matched else "no",
            })

    for label_key, dn in app_by_label.items():
        if label_key not in word_by_label:
            rows.append({
                "organization": org.organization_name, "field": f"({category_label} row)",
                "word_value": "(not found in Word for this section)", "app_value": dn.name,
                "match": "app_only",
            })


def main():
    with app.app_context():
        orgs = Organization.query.all()
        org_id_by_name_lower = {o.organization_name.strip().lower(): o.id for o in orgs}
        org_by_id = {o.id: o for o in orgs}

    segments, heading_raw_by_name, address_by_name = parse_word_with_address(WORD_PATH, org_id_by_name_lower)

    with app.app_context():
        groups, unmatched_word_sections = group_segments(segments, org_id_by_name_lower)

        all_employees = Employee.query.all()
        emp_by_org = defaultdict(list)
        for e in all_employees:
            if e.organization_id:
                emp_by_org[e.organization_id].append(e)

        all_numbers = DirectoryNumber.query.all()
        cr_by_org = defaultdict(list)
        sw_by_org = defaultdict(list)
        for dn in all_numbers:
            if dn.organization_id is None:
                continue
            cat = (dn.category or "").lower()
            if "switchyard" in cat:
                sw_by_org[dn.organization_id].append(dn)
            elif "control room" in cat:
                cr_by_org[dn.organization_id].append(dn)

        rows = []
        orgs_checked = 0
        orgs_with_word_match = 0
        orgs_without_word_match = []

        for org in orgs:
            orgs_checked += 1
            g = groups.get(org.id)
            if g is None:
                orgs_without_word_match.append(org.organization_name)
                rows.append({
                    "organization": org.organization_name, "field": "(organization)",
                    "word_value": "(no Word section matched this organization)",
                    "app_value": org.organization_name, "match": "app_only",
                })
                continue

            orgs_with_word_match += 1
            word_name = g["names_seen"][0] if g["names_seen"] else None
            diff_org_name(rows, org, heading_raw_by_name.get(word_name) if word_name else None)
            diff_address(rows, org, address_by_name.get(word_name) if word_name else None)
            diff_employees(rows, org, g["employees"], emp_by_org.get(org.id, []))
            diff_cr_sw(rows, org, g["cr"], cr_by_org.get(org.id, []), "control_room")
            diff_cr_sw(rows, org, g["sw"], sw_by_org.get(org.id, []), "switch_yard")

        for name in unmatched_word_sections:
            rows.append({
                "organization": name, "field": "(organization)",
                "word_value": name, "app_value": "(no organization in the app matched this Word section)",
                "match": "word_only",
            })

    fieldnames = ["organization", "field", "word_value", "app_value", "match"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    from collections import Counter
    match_counts = Counter(r["match"] for r in rows)

    print("=" * 70)
    print("  Full Word <-> App Field Diff -- counts (nothing changed)")
    print("=" * 70)
    print(f"  Organizations checked          : {orgs_checked}")
    print(f"  Organizations with a Word match : {orgs_with_word_match}")
    print(f"  Organizations with NO Word match: {len(orgs_without_word_match)}")
    print(f"  Word sections with no app org   : {len(unmatched_word_sections)}")
    print(f"  Total field rows written        : {len(rows)}  -> {OUT_CSV}")
    for k, v in match_counts.items():
        print(f"    - {k:<20}: {v}")
    print("=" * 70)


if __name__ == "__main__":
    main()
