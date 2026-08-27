"""Full data-cleanliness sweep across ALL rows of directory_numbers (not just
previously-flagged stations), requested as Phase 1/2/3 of a review pass
before any fix script is written. Read-only throughout -- writes proposal
CSVs and a console summary only, changes nothing.

Schema note: directory_numbers has a single combined `phone_number` text
field (no separate office_phone/mobile -- see scan_data_quality.py's
docstring, still true), no `station` column (station = the linked
Organization's name via organization_id), and no address/notes column to
split an embedded address into. Category ("control_room" / "switch_yard" /
"employee") is derived from the free-text `category` column using the exact
same rule routes/user_routes.py already uses to bucket the telephone
directory page (switchyard / control room / everything else) -- not the
switch_yard/control_room boolean columns, which the migration 014 comment
says are additive and don't drive this grouping.

PHASE 1 -- reports/full_data_quality_audit.csv
  One row per (directory_number id, field) that has at least one issue,
  issue_type(s) semicolon-joined when more than one applies to that field,
  plus one row per id for possible_duplicate_row (field="(row)"):
    - address_in_name (field=name): comma-list / parenthetical / PIN-code /
      "Backup" / "State Load Despatch Centre" / "SLDC" / "DISCOM" / city
      hints, same style of signal-scan as scan_data_quality.py's
      scan_address_in_name, extended with a parenthetical-address pattern
      and DISCOM/SLDC hints per review feedback.
    - duplicated_phone (field=phone_number): same repeated-segment logic as
      scan_data_quality.py's scan_duplicated_phone.
    - numbered_prefix (field=name): name has a leading "N." / "N " pattern;
      cross-checked against the Word directory for the same station -- see
      build_word_label_index() -- and marked confirmed/unconfirmed/
      no-word-data accordingly in the notes column.
    - garbled_field (field=phone_number): non-empty phone_number with no
      digits at all (placeholder text only, e.g. "Control Room", "SCADA
      VOIP/ PSS VOIP:").
    - possible_duplicate_row: >1 row sharing the same (organization_id,
      category, normalized name) after normalizing punctuation/case/
      whitespace and folding "State SLDC" -> "SLDC" (per review feedback --
      these are the same entity, not a real difference).

PHASE 2 -- reports/word_vs_app_cleanliness.csv
  One row per station that Word and the app both have data for (any Word
  segment that matched an Organization). Re-parses the Word doc's own raw
  table-row text directly (not through compare_word_excel's already-deduped
  phone/email token lists, since the raw text is what the duplicated-phone
  and garbled-field checks need to see) and applies the same signal
  detectors to Word's own labels/phone cells, then reports whether the app
  side has any Phase-1 issue for that station and whether Word's own text
  already carries the same kind of issue -- so we can tell whether dirty
  data originated in Word or was introduced during import/editing. Best-
  effort: Word table layouts vary, so a station with no clear Word issue
  found simply reports False rather than being skipped.

PHASE 3 -- printed summary only: total rows scanned, count per issue type,
and how many distinct directory_number ids have more than one issue type
stacked.

Usage:
  python scripts/full_data_quality_audit.py
"""

import csv
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx
from docx.text.paragraph import Paragraph

from app import app
from models.directory_number import DirectoryNumber

from scripts.compare_word_excel import (
    clean,
    strip_leading_number,
    iter_block_items,
    find_header_columns,
    is_header_row,
    NUMBERED_TITLE_RE,
    SWITCHYARD_RE,
    LEADING_NUMBER_RE,
    PHONE_RE,
)
from scripts.reconcile_word_database import (
    parse_word, load_db as load_reconcile_db, match_to_db, looks_like_known_org, WORD_PATH,
)

AUDIT_CSV = "reports/full_data_quality_audit.csv"
CLEANLINESS_CSV = "reports/word_vs_app_cleanliness.csv"

SEGMENT_SPLIT_RE = re.compile(r"[/;,]")
PIN_CODE_RE = re.compile(r"-\s*\d{6}\b")
PAREN_ADDRESS_RE = re.compile(r"\([^()]{10,}\)")

CITY_STATE_HINTS = (
    "maharashtra", "gujarat", "chhattisgarh", "madhya pradesh", "goa", "raipur",
    "bhopal", "mumbai", "ahmedabad", "vadodara", "surat", "indore", "bhilai",
    "raigad", "nagpur", "panaji", "danganiya",
)

# Full org-name phrases only -- NOT bare "sldc" or "discom": those are
# legitimate abbreviations that show up constantly in genuine control-room
# labels ("SLDC-Control room", "SLDC DNH"), so treating them as an
# address_in_name signal on their own produced false positives (verified:
# rows 7681/7682/7686/7689/7690/7692 all flagged solely for containing
# "sldc" with no other signal, none of them actually address text).
ORG_TEXT_HINTS = ("state load despatch centre", "load despatch centre")

STATE_SLDC_RE = re.compile(r"\bstate\s+sldc\b", re.IGNORECASE)


def normalize_dup_key(name):
    """Punctuation/case/whitespace-normalized, with 'State SLDC' folded to
    'SLDC' (per review feedback -- same entity, not a real naming
    difference) so duplicate-row detection isn't fooled by that variant."""
    name = STATE_SLDC_RE.sub("SLDC", name or "")
    name = name.lower()
    name = re.sub(r"[^a-z0-9]+", " ", name)
    return re.sub(r"\s+", " ", name).strip()


def classify_category(category_text):
    cat = (category_text or "").lower()
    if "switchyard" in cat or "switch yard" in cat:
        return "switch_yard"
    if "control room" in cat:
        return "control_room"
    return "employee"


def address_in_name_signals(name):
    name = name or ""
    lower = name.lower()
    signals = []
    if "address:" in lower:
        signals.append("'Address:' marker")
    if PIN_CODE_RE.search(name):
        signals.append("PIN-code-like pattern")
    if PAREN_ADDRESS_RE.search(name):
        signals.append("long parenthetical (address-like)")
    if name.count(",") >= 2:
        signals.append(f"{name.count(',')} commas (address-list structure)")
    if "backup" in lower:
        signals.append("'Backup'")
    hit_org = [h for h in ORG_TEXT_HINTS if h in lower]
    if hit_org:
        signals.append("org-text hint(s): " + ", ".join(hit_org))
    hit_cities = [c for c in CITY_STATE_HINTS if c in lower]
    if hit_cities:
        signals.append("city/state name(s): " + ", ".join(hit_cities))
    if len(name) > 80:
        signals.append(f"unusually long name ({len(name)} chars)")
    return signals


def duplicated_phone_signal(value):
    if not value:
        return None
    segments = [s.strip() for s in SEGMENT_SPLIT_RE.split(value) if s.strip()]
    if len(segments) < 2:
        return None
    counts = Counter(segments)
    repeat_count = max(counts.values())
    if repeat_count >= 2:
        return repeat_count
    return None


def garbled_field_signal(value):
    if not value or not value.strip():
        return False
    return re.search(r"\d", value) is None


# ─── Word cross-reference (numbered_prefix confirmation) ────────────────────

def build_word_label_index():
    """org_id -> set of lowercased raw labels appearing in Word for that
    station (employee names + control-room/switchyard labels), used only to
    confirm/refute numbered_prefix flags."""
    with app.app_context():
        db_data = load_reconcile_db()

    segments = parse_word(WORD_PATH, db_data["org_id_by_name_lower"])
    index = defaultdict(set)
    for name, employees, cr_records, sw_records in segments:
        org_id, _display = match_to_db(name, db_data)
        if org_id is None:
            continue
        for emp in employees:
            index[org_id].add(emp.strip().lower())
        for rec in cr_records:
            index[org_id].add(rec["label"].strip().lower())
        for rec in sw_records:
            index[org_id].add(rec["label"].strip().lower())
    return index


def numbered_prefix_check(name, org_id, word_index):
    stripped = LEADING_NUMBER_RE.sub("", name).strip()
    if stripped == name.strip():
        return None  # no leading-number pattern at all

    labels = word_index.get(org_id)
    if not labels:
        return "numbered_prefix (no Word data for this station -- unconfirmed)"

    stripped_hit = stripped.lower() in labels
    full_hit = name.strip().lower() in labels
    if stripped_hit and not full_hit:
        return "numbered_prefix (confirmed -- Word has this entry without the number)"
    if full_hit:
        return None  # Word also carries the same numbered form -- not an artifact
    return "numbered_prefix (Word has neither form for this station -- unconfirmed)"


# ─── PHASE 1 ──────────────────────────────────────────────────────────────

def run_phase1(word_index):
    with app.app_context():
        rows = DirectoryNumber.query.all()
        row_data = [
            {
                "id": dn.id,
                "station": dn.organization_obj.organization_name if dn.organization_obj else (dn.organization or ""),
                "organization_id": dn.organization_id,
                "category": classify_category(dn.category),
                "name": dn.name or "",
                "phone_number": dn.phone_number or "",
            }
            for dn in rows
        ]

    out_rows = []

    # name-field issues: address_in_name, numbered_prefix
    for r in row_data:
        issues, notes = [], []
        sig = address_in_name_signals(r["name"])
        if sig:
            issues.append("address_in_name")
            notes.append("address_in_name signals: " + "; ".join(sig))
        np_note = numbered_prefix_check(r["name"], r["organization_id"], word_index)
        if np_note:
            issues.append("numbered_prefix")
            notes.append(np_note)
        if issues:
            out_rows.append({
                "id": r["id"], "station": r["station"], "category": r["category"],
                "field": "name", "issue_type(s)": "; ".join(issues),
                "raw_value": r["name"], "matched_duplicate_id": "",
                "notes": " | ".join(notes),
            })

    # phone_number-field issues: duplicated_phone, garbled_field
    for r in row_data:
        issues, notes = [], []
        repeat_count = duplicated_phone_signal(r["phone_number"])
        if repeat_count:
            issues.append("duplicated_phone")
            notes.append(f"segment repeated {repeat_count}x")
        if garbled_field_signal(r["phone_number"]):
            issues.append("garbled_field")
            notes.append("no digits at all -- placeholder text only")
        if issues:
            out_rows.append({
                "id": r["id"], "station": r["station"], "category": r["category"],
                "field": "phone_number", "issue_type(s)": "; ".join(issues),
                "raw_value": r["phone_number"], "matched_duplicate_id": "",
                "notes": " | ".join(notes),
            })

    # possible_duplicate_row: same (org, category, normalized name), >1 row
    groups = defaultdict(list)
    for r in row_data:
        key = (r["organization_id"], r["category"], normalize_dup_key(r["name"]))
        if key[2]:  # skip empty-name rows -- nothing meaningful to group on
            groups[key].append(r)

    dup_row_count = 0
    for key, members in groups.items():
        if len(members) < 2:
            continue
        dup_row_count += len(members)
        ids = [m["id"] for m in members]
        for m in members:
            others = [str(i) for i in ids if i != m["id"]]
            out_rows.append({
                "id": m["id"], "station": m["station"], "category": m["category"],
                "field": "(row)", "issue_type(s)": "possible_duplicate_row",
                "raw_value": m["name"], "matched_duplicate_id": ", ".join(others),
                "notes": f"{len(members)} rows share this normalized name for this station+category",
            })

    fieldnames = ["id", "station", "category", "field", "issue_type(s)", "raw_value", "matched_duplicate_id", "notes"]
    os.makedirs("reports", exist_ok=True)
    with open(AUDIT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(out_rows)

    return row_data, out_rows, dup_row_count


# ─── PHASE 2 ──────────────────────────────────────────────────────────────

def parse_word_raw_rows(path, org_id_by_name_lower):
    """Re-walks the Word doc directly (not via compare_word_excel's already
    phone/email-deduped structures) so duplicated-phone and garbled-field
    checks see the actual raw cell text, matching what a human would read.
    Yields (station_name, kind, label, raw_row_text) per roster row, where
    kind is 'employee' / 'control_room' / 'switch_yard'.

    Also detects "banner" title rows (every non-empty cell repeats the same
    text, no leading number -- e.g. a station name spanning all 7 columns)
    as station boundaries, same as reconcile_word_database.py and
    full_word_vs_app_field_diff.py now do -- ported here for the same
    reason: NUMBERED_TITLE_RE alone misses 16 real stations (Power Grid,
    Grid Controller, WRLDC's own State SLDCs, NTPC, etc.), silently merging
    their rows into whichever numbered/outer heading preceded them. Guarded
    by looks_like_known_org (reused from reconcile_word_database.py, not
    reimplemented) since the same all-cells-identical shape also matches
    address-only subtitle rows and decorative quotes sitting mid-table
    between two real employees of the SAME station -- both confirmed by
    direct inspection to otherwise silently steal rows into a bogus
    pseudo-station matching nothing real."""
    doc = docx.Document(path)
    current_heading = None
    current_station = None

    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue

        table = item
        rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
        name_col, desig_col = find_header_columns(rows_cells)
        if name_col is None:
            continue

        station = current_heading
        for i, cells in enumerate(rows_cells):
            if not cells:
                continue
            first_cell = cells[0]
            if NUMBERED_TITLE_RE.match(first_cell):
                station = strip_leading_number(first_cell)
                continue

            non_empty = [c for c in cells if c]
            if len(non_empty) >= 2 and len(set(non_empty)) == 1 and not is_header_row(cells):
                lookahead = rows_cells[i + 1:i + 3]
                if any(is_header_row(r) for r in lookahead if r) and looks_like_known_org(non_empty[0], org_id_by_name_lower):
                    station = strip_leading_number(non_empty[0])
                continue

            if is_header_row(cells):
                continue
            if not any(cells):
                continue

            name_cell = cells[name_col] if name_col < len(cells) else ""
            desig_cell = cells[desig_col] if desig_col < len(cells) else ""
            if not name_cell:
                continue

            raw_row_text = " | ".join(c for c in cells if c)

            if SWITCHYARD_RE.search(name_cell) or SWITCHYARD_RE.search(desig_cell):
                label = name_cell if name_cell.strip().lower() not in ("switchyard", "switch yard") else desig_cell
                yield station, "switch_yard", label or name_cell, raw_row_text
                continue
            if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
                label = name_cell if name_cell.strip().lower() not in ("control room", "control rooms") else desig_cell
                yield station, "control_room", label or name_cell, raw_row_text
                continue
            if name_cell.lower() == desig_cell.lower():
                continue

            yield station, "employee", name_cell, raw_row_text


def word_raw_signals(label, raw_row_text):
    issues = []
    if address_in_name_signals(label):
        issues.append("address_in_name")
    phone_tokens = PHONE_RE.findall(raw_row_text)
    if phone_tokens:
        counts = Counter(phone_tokens)
        if max(counts.values()) >= 2:
            issues.append("duplicated_phone")
    if LEADING_NUMBER_RE.sub("", label).strip() != label.strip():
        issues.append("numbered_prefix")
    non_label_text = raw_row_text.replace(label, "", 1)
    if non_label_text.strip() and not PHONE_RE.search(non_label_text):
        issues.append("garbled_field")
    return issues


def run_phase2(row_data, phase1_out_rows):
    app_issues_by_org = defaultdict(set)
    id_to_org = {r["id"]: r["organization_id"] for r in row_data}
    for r in phase1_out_rows:
        org_id = id_to_org.get(r["id"])
        for issue in r["issue_type(s)"].split("; "):
            app_issues_by_org[org_id].add(issue)

    with app.app_context():
        db_data = load_reconcile_db()
    org_name_by_id = db_data["org_name_by_id"]

    word_issues_by_org = defaultdict(set)
    matched_org_ids = set()
    for station, kind, label, raw_row_text in parse_word_raw_rows(WORD_PATH, db_data["org_id_by_name_lower"]):
        if not station:
            continue
        org_id, _display = match_to_db(station, db_data)
        if org_id is None:
            continue
        matched_org_ids.add(org_id)
        for issue in word_raw_signals(label, raw_row_text):
            word_issues_by_org[org_id].add(issue)

    out_rows = []
    for org_id in sorted(matched_org_ids, key=lambda oid: org_name_by_id.get(oid, "")):
        app_issues = app_issues_by_org.get(org_id, set())
        word_issues = word_issues_by_org.get(org_id, set())
        if not app_issues and not word_issues:
            continue
        both = app_issues & word_issues
        origin = (
            "originated_in_word" if both and app_issues <= word_issues
            else "introduced_during_import" if app_issues and not word_issues
            else "mixed" if app_issues and word_issues
            else "word_only_not_in_app"
        )
        out_rows.append({
            "station": org_name_by_id.get(org_id, str(org_id)),
            "app_has_issue": bool(app_issues),
            "app_issue_types": "; ".join(sorted(app_issues)),
            "word_has_issue": bool(word_issues),
            "word_issue_types": "; ".join(sorted(word_issues)),
            "likely_origin": origin,
        })

    fieldnames = ["station", "app_has_issue", "app_issue_types", "word_has_issue", "word_issue_types", "likely_origin"]
    with open(CLEANLINESS_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(out_rows)

    return out_rows


# ─── PHASE 3 ──────────────────────────────────────────────────────────────

def print_summary(row_data, phase1_out_rows, dup_row_count, phase2_out_rows):
    issue_counts = Counter()
    ids_with_issue_types = defaultdict(set)
    for r in phase1_out_rows:
        for issue in r["issue_type(s)"].split("; "):
            issue_counts[issue] += 1
            ids_with_issue_types[r["id"]].add(issue)

    stacked = sum(1 for issues in ids_with_issue_types.values() if len(issues) > 1)

    print("=" * 70)
    print("  Full Data Quality Audit -- Phase 3 Summary (nothing changed)")
    print("=" * 70)
    print(f"  Total directory_numbers rows scanned : {len(row_data)}")
    print(f"  Rows with >=1 issue                  : {len(ids_with_issue_types)}")
    print("  Issue-type counts (row-level occurrences):")
    for issue, count in sorted(issue_counts.items()):
        print(f"    - {issue:<22}: {count}")
    print(f"  possible_duplicate_row involved rows : {dup_row_count}")
    print(f"  Rows with >1 issue type stacked       : {stacked}")
    print("-" * 70)
    print(f"  Phase 1 written to: {AUDIT_CSV}")
    print(f"  Stations compared Word<->app cleanliness: {len(phase2_out_rows)}")
    origin_counts = Counter(r["likely_origin"] for r in phase2_out_rows)
    for origin, count in sorted(origin_counts.items()):
        print(f"    - {origin:<28}: {count}")
    print(f"  Phase 2 written to: {CLEANLINESS_CSV}")
    print("=" * 70)


def main():
    word_index = build_word_label_index()
    row_data, phase1_out_rows, dup_row_count = run_phase1(word_index)
    phase2_out_rows = run_phase2(row_data, phase1_out_rows)
    print_summary(row_data, phase1_out_rows, dup_row_count, phase2_out_rows)


if __name__ == "__main__":
    main()
