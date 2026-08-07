"""Validates uploads/WR_DB_Ready_Final_Verified_TD_UPDATED.xlsx (soon to be
imported into PostgreSQL) against uploads/Western Region Phone Directory
2026 Main_TD.docx, the official source of truth. Read-only on both inputs.
Writes Validation_Report.xlsx at the repo root.

Reuses the Word-parsing primitives already built for the existing
compare_word_excel.py / reconcile_word_database.py pipeline rather than
re-solving Word table parsing from scratch.

Usage:
  python scripts/generate_validation_report.py
"""

import os
import sys
from collections import defaultdict, Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import docx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from docx.text.paragraph import Paragraph

from scripts.compare_word_excel import (
    clean, strip_leading_number, dedup_preserve_order,
    classify_row_content, normalize_phone_set, normalize_email_set,
    iter_block_items, find_header_columns, is_header_row,
    NUMBERED_TITLE_RE, SWITCHYARD_RE,
    match_segment, _best_substring_match,
)
from scripts.reconcile_word_database import _merge_record

WORD_PATH = "uploads/Western Region Phone Directory 2026 Main_TD.docx"
# Was 'uploads/WR_DB_Ready_Final_Verified_TD_UPDATED.xlsx' -- an
# intermediate snapshot from an earlier validation round (880 employees /
# 214 control rooms). Repointed at the current canonical master workbook
# (985 employees / 339 control rooms) so this report reflects what's
# actually being corrected now.
EXCEL_PATH = "WR_DB_Ready_Final_Verified_v3.xlsx"
OUTPUT_XLSX = "Validation_Report.xlsx"

SYNOPSIS_HEADING = "Synopsis of Important Telephone Numbers"

# heading text -> Emergency_Services.service_type ("Empaneled Hospitals" is
# special-cased to the Hospitals sheet instead)
EMERGENCY_SUBHEADINGS = {
    "Police Stations": "Police Station",
    "Fire Stations": "Fire Station",
    "Empaneled Hospitals": None,
    "Blood Bank": "Blood Bank",
    "Forest Department": "Forest Department",
    "Disaster Management Control Centre": "Disaster Management Control Centre",
    "Bomb Disposal Squads": "Bomb Disposal Squad",
    "CORONA Helpline": "CORONA Helpline",
}

# heading text -> Reference_Sections.section (best-effort category)
REFERENCE_HEADINGS = [
    "RRAS Provider in Western Region",
    "Disaster Management Contact Details",
    "Nodal Officers in Western Region",
    "Contact Details of LDCs for Disaster Management in Power Sector",
    "Hotline Nos- Orange Directory",
    "Important Vendors in WRLDC",
]

DISCREPANCY_HEADER = [
    "Category", "Source (Word)", "Current Value (Excel)",
    "Expected Value (Word)", "Issue Description", "Suggested Correction",
]


CR_BARE_WORDS = {"control room", "control rooms"}
SW_BARE_WORDS = {"switchyard", "switch yard"}


def pick_best_label(cell_a, cell_b, bare_words):
    """Prefers whichever of two candidate label strings is NOT just the bare
    generic category word, so two differently-worded cells for the same
    record (e.g. cr_label='GRID CONTROLLER OF INDIA LIMITED' / designation=
    'SLDC-Control room') resolve to the more specific one. Mirrors the
    pick_label heuristic already used for the Word side in
    reconcile_word_database.parse_table_segments_ext."""
    a_bare = cell_a.strip().lower() in bare_words
    b_bare = cell_b.strip().lower() in bare_words
    if b_bare and not a_bare and cell_a.strip():
        return cell_a
    if a_bare and not b_bare and cell_b.strip():
        return cell_b
    return cell_a if cell_a.strip() else cell_b


def merge_excel_records_by_label(items):
    """Multiple Excel rows can resolve to the same effective label (e.g.
    several generically-labeled 'Control Room' rows under one organization)
    -- union their phone/email text rather than silently dropping all but
    the first, mirroring reconcile_word_database._merge_record on the Word
    side."""
    merged = {}
    for item in items:
        key = item["label"].strip().lower()
        if key in merged:
            merged[key]["phones_raw"] = (merged[key]["phones_raw"] + " " + item["phones_raw"]).strip()
            merged[key]["emails_raw"] = (merged[key]["emails_raw"] + " " + item["emails_raw"]).strip()
        else:
            merged[key] = dict(item)
    return list(merged.values())


def make_row(category, source_word, current_excel, expected_word, issue, suggestion):
    return {
        "Category": category,
        "Source (Word)": source_word,
        "Current Value (Excel)": current_excel,
        "Expected Value (Word)": expected_word,
        "Issue Description": issue,
        "Suggested Correction": suggestion,
    }


# ─── EXCEL LOADING ───────────────────────────────────────────────────────────

def load_excel(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    def rows(sheet):
        it = wb[sheet].iter_rows(values_only=True)
        next(it)  # header
        return list(it)

    organizations = []
    org_name_by_id, org_id_by_name_lower = {}, {}
    for org_id, name, address, state in rows("organizations"):
        name = clean(name)
        if org_id is None or not name:
            continue
        organizations.append({"org_id": org_id, "name": name, "address": clean(address), "state": clean(state)})
        org_name_by_id[org_id] = name
        org_id_by_name_lower.setdefault(name.lower(), org_id)

    sub_organizations = []
    suborg_name_by_id, suborg_parent_by_id, suborg_id_by_name_lower = {}, {}, {}
    for sid, org_id, name, address, state in rows("sub_organizations"):
        name = clean(name)
        if sid is None or not name:
            continue
        sub_organizations.append({"suborg_id": sid, "org_id": org_id, "name": name, "address": clean(address), "state": clean(state)})
        suborg_name_by_id[sid] = name
        suborg_parent_by_id[sid] = org_id
        suborg_id_by_name_lower.setdefault(strip_leading_number(name).lower(), sid)

    employees = []
    employees_by_id = {}
    for emp_id, org_id, suborg_id, department, name, designation in rows("employees"):
        name = clean(name)
        if not name:
            continue
        rec = {
            "employee_id": emp_id, "org_id": org_id, "suborg_id": suborg_id,
            "department": clean(department), "name": name, "designation": clean(designation),
        }
        employees.append(rec)
        employees_by_id[emp_id] = rec

    kmp, kmp_by_id = [], {}
    for kmp_id, org_id, name, designation, office, residence, mobile, email in rows("KMP"):
        name = clean(name)
        if not name:
            continue
        rec = {
            "kmp_id": kmp_id, "org_id": org_id, "name": name, "designation": clean(designation),
            "office_phone": office or "", "residence_phone": residence or "",
            "mobile_phone": mobile or "", "email": email or "",
        }
        kmp.append(rec)
        kmp_by_id[kmp_id] = rec

    utility_heads, uh_by_id = [], {}
    for uh_id, org_id, name, designation, office, residence, mobile, email in rows("Utility_Heads"):
        name = clean(name)
        if not name:
            continue
        rec = {
            "utility_head_id": uh_id, "org_id": org_id, "name": name, "designation": clean(designation),
            "office_phone": office or "", "residence_phone": residence or "",
            "mobile_phone": mobile or "", "email": email or "",
        }
        utility_heads.append(rec)
        uh_by_id[uh_id] = rec

    control_rooms, cr_by_id = [], {}
    for cr_id, org_id, suborg_id, _oname, _sname, label, designation, ph1, ph2, mob1, mob2, email1 in rows("control_rooms"):
        label = clean(label)
        designation = clean(designation)
        # cr_label is often just the org/station banner repeated (e.g. "GRID
        # CONTROLLER OF INDIA LIMITED"), not a distinguishing name -- prefer
        # designation when it's the more specific of the two, same heuristic
        # Word's own pick_label uses.
        effective_label = pick_best_label(label, designation, CR_BARE_WORDS)
        if not effective_label:
            continue
        rec = {
            "id": cr_id, "org_id": org_id, "suborg_id": suborg_id,
            "label": effective_label, "designation": designation,
            "phones_raw": " ".join(str(x) for x in (ph1, ph2, mob1, mob2) if x),
            "emails_raw": str(email1 or ""),
        }
        control_rooms.append(rec)
        cr_by_id[cr_id] = rec

    switchyards, sw_by_id = [], {}
    for sw_id, org_id, suborg_id, _oname, _sname, label, ph1, mob1, email1 in rows("switchyards"):
        label = clean(label)
        if not label:
            continue
        rec = {
            "id": sw_id, "org_id": org_id, "suborg_id": suborg_id, "label": label,
            "phones_raw": " ".join(str(x) for x in (ph1, mob1) if x),
            "emails_raw": str(email1 or ""),
        }
        switchyards.append(rec)
        sw_by_id[sw_id] = rec

    hospitals, hosp_by_id = [], {}
    for hid, org_id, name, contact, opd, emergency, address, website in rows("Hospitals"):
        name = clean(name)
        if not name:
            continue
        rec = {
            "hospital_id": hid, "org_id": org_id, "name": name,
            "contact_number": contact or "", "opd_number": opd or "",
            "emergency_number": emergency or "", "address": clean(address), "website": website or "",
        }
        hospitals.append(rec)
        hosp_by_id[hid] = rec

    emergency_services = []
    for sid, service_type, name, contact, website in rows("Emergency_Services"):
        name = clean(name)
        if not name:
            continue
        emergency_services.append({
            "service_id": sid, "service_type": clean(service_type), "name": name,
            "contact_number": contact or "", "website": website or "",
        })

    reference_sections = []
    for category, section, sub_context, name, designation, ph_o, ph_r, mobile, landline, email, contact in rows("Reference_Sections"):
        name = clean(name)
        if not name:
            continue
        reference_sections.append({
            "category": clean(category), "section": clean(section), "sub_context": clean(sub_context),
            "name": name, "designation": clean(designation),
            "phones_raw": " ".join(str(x) for x in (ph_o, ph_r, mobile, landline, contact) if x),
            "email": email or "",
        })

    phones_by_ref = defaultdict(list)
    for _pid, ref_type, ref_id, _ptype, number in rows("phone_numbers"):
        if ref_type and ref_id is not None and number:
            phones_by_ref[(ref_type, ref_id)].append(str(number))

    emails_by_ref = defaultdict(list)
    for _eid, ref_type, ref_id, email in rows("email_addresses"):
        if ref_type and ref_id is not None and email:
            emails_by_ref[(ref_type, ref_id)].append(str(email))

    # employees carry no inline contact columns -- their only source is the
    # normalized phone_numbers/email_addresses tables
    for e in employees:
        e["phones_raw"] = " ".join(phones_by_ref.get(("employee", e["employee_id"]), []))
        e["emails_raw"] = " ".join(emails_by_ref.get(("employee", e["employee_id"]), []))

    return {
        "organizations": organizations, "org_name_by_id": org_name_by_id, "org_id_by_name_lower": org_id_by_name_lower,
        "sub_organizations": sub_organizations, "suborg_name_by_id": suborg_name_by_id,
        "suborg_parent_by_id": suborg_parent_by_id, "suborg_id_by_name_lower": suborg_id_by_name_lower,
        "employees": employees, "employees_by_id": employees_by_id,
        "kmp": kmp, "kmp_by_id": kmp_by_id,
        "utility_heads": utility_heads, "uh_by_id": uh_by_id,
        "control_rooms": control_rooms, "cr_by_id": cr_by_id,
        "switchyards": switchyards, "sw_by_id": sw_by_id,
        "hospitals": hospitals, "hosp_by_id": hosp_by_id,
        "emergency_services": emergency_services,
        "reference_sections": reference_sections,
        "phones_by_ref": phones_by_ref, "emails_by_ref": emails_by_ref,
    }


# ─── WORD PARSING: STATION TABLES (employees / control rooms / switchyards) ─

def parse_station_table(table, fallback_name):
    """Same segmentation as reconcile_word_database.parse_table_segments_ext,
    extended to also capture per-employee designation/phone/email -- that
    function only keeps employee names, since the existing pipeline never
    needed employee-level contact diffing. Returns (name, employee_records,
    cr_records, sw_records) segments, where an employee_record is
    {"name", "designation", "phones", "emails"}."""
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    name_col, desig_col = find_header_columns(rows_cells)
    if name_col is None:
        return []

    segments = []
    current_name = None
    current_employees = {}
    current_cr = []
    current_sw = []

    def flush():
        if current_employees or current_cr or current_sw:
            segments.append((
                current_name or fallback_name,
                list(current_employees.values()),
                list(current_cr),
                list(current_sw),
            ))

    def build_record(cells, label, exclude_lower):
        phones, emails, _others = classify_row_content(
            [c for c in cells if clean(c).lower() not in exclude_lower]
        )
        return {"label": label, "phones": phones, "emails": emails}

    def pick_label(name_cell, desig_cell, bare_words):
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
            current_employees, current_cr, current_sw = {}, [], []
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
            current_sw.append(build_record(cells, label, {label.lower()}))
            continue

        if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
            label = pick_label(name_cell, desig_cell, {"control room", "control rooms"})
            current_cr.append(build_record(cells, label, {label.lower()}))
            continue

        if name_cell.lower() == desig_cell.lower():
            continue  # repeated title-block noise row

        phones, emails, _others = classify_row_content(
            [c for c in cells if clean(c).lower() not in {name_cell.lower(), desig_cell.lower()}]
        )
        current_employees[name_cell.lower()] = {
            "name": name_cell, "designation": desig_cell, "phones": phones, "emails": emails,
        }

    flush()
    return segments


# ─── WORD PARSING: SYNOPSIS (KMP / Utility Head source) ─────────────────────

def parse_synopsis_table(table):
    """"Synopsis of Important Telephone Numbers": organization-name banner
    rows (every cell in the row repeats the same value -- a fully merged
    row), a Name/Designation/.../Email header row, then person rows. Word
    doesn't literally label anyone KMP or Utility Head; both sheets are
    curated subsets of this same pool, so every person row is returned and
    matched against each Excel sheet separately."""
    people = []
    current_org = None
    for row in table.rows:
        raw_cells = [clean(c.text) for c in row.cells]
        non_blank = dedup_preserve_order([c for c in raw_cells if c])
        if not non_blank:
            continue
        if len(non_blank) == 1:
            current_org = non_blank[0]
            continue
        if is_header_row(raw_cells):
            continue

        name = raw_cells[0] if raw_cells else ""
        designation = raw_cells[1] if len(raw_cells) > 1 else ""
        if not name or name.lower() == designation.lower():
            continue

        phones, emails, _others = classify_row_content(raw_cells[2:])
        people.append({
            "org_banner": current_org or "",
            "name": name, "designation": designation,
            "phones": phones, "emails": emails,
        })
    return people


# ─── WORD PARSING: EMERGENCY CONTACT NUMBERS (Hospitals / Emergency_Services) ─

def parse_simple_contact_table(table):
    """Generic Sl.No/Name/Contact.../Website table used throughout the
    Emergency Contact Numbers back-matter -- maps columns by header text
    rather than position."""
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    if not rows_cells:
        return []
    header_lower = [h.lower() for h in rows_cells[0]]

    def find_col(*keywords):
        for i, h in enumerate(header_lower):
            if any(k in h for k in keywords):
                return i
        return None

    name_col = find_col("name", "hospital", "station")
    contact_col = find_col("contact")
    opd_col = find_col("opd")
    emergency_col = find_col("emergency")
    website_col = find_col("website")
    if name_col is None:
        return []

    def cell(cells, col):
        return cells[col] if col is not None and col < len(cells) else ""

    results = []
    for cells in rows_cells[1:]:
        if not any(cells):
            continue
        name = cell(cells, name_col)
        if not name or name.lower() in ("sl. no.", "sl.no."):
            continue
        results.append({
            "name": name, "contact_number": cell(cells, contact_col),
            "opd_number": cell(cells, opd_col), "emergency_number": cell(cells, emergency_col),
            "website": cell(cells, website_col),
        })
    return results


# ─── WORD PARSING: REFERENCE_SECTIONS BACK-MATTER (best effort) ─────────────

def parse_reference_table(table):
    """Best-effort parse of a Reference_Sections back-matter table (RRAS,
    Disaster Management, Nodal Officers, Hotlines, Vendors) -- these are
    heterogeneous in shape, so this only proceeds when a name-like header
    column can be confidently located; otherwise every row in this table is
    routed to Needs_Manual_Verification rather than guessed. Returns
    (rows, confident)."""
    rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
    if not rows_cells:
        return [], True
    header_lower = [h.lower() for h in rows_cells[0]]

    def find_col(*keywords):
        for i, h in enumerate(header_lower):
            if any(k in h for k in keywords):
                return i
        return None

    name_col = find_col("name", "provider", "vendor", "office", "station", "officer")
    if name_col is None:
        return [], False

    results = []
    for cells in rows_cells[1:]:
        if not any(cells):
            continue
        name = cells[name_col] if name_col < len(cells) else ""
        if not name:
            continue
        phones, emails, _others = classify_row_content([c for i, c in enumerate(cells) if i != name_col])
        results.append({"name": name, "phones": phones, "emails": emails})
    return results, True


# ─── WORD DOCUMENT WALK ──────────────────────────────────────────────────────

def parse_word(path):
    doc = docx.Document(path)
    current_heading = None
    current_emergency_sub = None
    current_reference_heading = None

    all_station_segments = []
    word_kmp_candidates = []
    word_hospitals = []
    word_emergency_services = []
    word_reference_rows = defaultdict(list)
    unresolved_reference_sections = set()
    totals = Counter()

    reference_headings_stripped = {strip_leading_number(h): h for h in REFERENCE_HEADINGS}

    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                heading_text = strip_leading_number(clean(item.text))
                current_heading = heading_text
                current_emergency_sub = heading_text if heading_text in EMERGENCY_SUBHEADINGS else None
                current_reference_heading = reference_headings_stripped.get(heading_text)
            continue

        # item is a Table
        if current_heading == SYNOPSIS_HEADING:
            people = parse_synopsis_table(item)
            word_kmp_candidates.extend(people)
            totals["synopsis"] += len(people)
            continue

        if current_emergency_sub is not None:
            found = parse_simple_contact_table(item)
            if current_emergency_sub == "Empaneled Hospitals":
                word_hospitals.extend(found)
            else:
                service_type = EMERGENCY_SUBHEADINGS[current_emergency_sub]
                for r in found:
                    r["service_type"] = service_type
                word_emergency_services.extend(found)
            totals["emergency"] += len(found)
            continue

        if current_reference_heading is not None:
            found, confident = parse_reference_table(item)
            if confident:
                word_reference_rows[current_reference_heading].extend(found)
                totals["reference_confident"] += len(found)
            else:
                unresolved_reference_sections.add(current_reference_heading)
                totals["reference_unconfident"] += 1
            continue

        segments = parse_station_table(item, current_heading)
        all_station_segments.extend(segments)
        for _name, employees, cr_records, sw_records in segments:
            totals["station"] += len(employees) + len(cr_records) + len(sw_records)

    return {
        "all_station_segments": all_station_segments,
        "word_kmp_candidates": word_kmp_candidates,
        "word_hospitals": word_hospitals,
        "word_emergency_services": word_emergency_services,
        "word_reference_rows": word_reference_rows,
        "unresolved_reference_sections": unresolved_reference_sections,
        "totals": totals,
    }


# ─── MATCHING WORD STATIONS TO EXCEL ORG/SUBORG ─────────────────────────────

def group_word_segments(all_segments, excel):
    """Two-tier org/suborg matching (compare_word_excel.match_segment) +
    three-way employee/control-room/switchyard record carrying -- neither
    existing group_segments variant in the pipeline supports both at once."""
    groups = {}
    for name, employees, cr_records, sw_records in all_segments:
        match_type, match_id, display_name = match_segment(name, excel)
        key = (match_type, match_id)
        g = groups.setdefault(key, {
            "display_name": display_name, "match_type": match_type, "match_id": match_id,
            "employees": {}, "cr_by_label": {}, "sw_by_label": {},
        })
        for emp in employees:
            g["employees"][emp["name"].lower()] = emp
        for rec in cr_records:
            _merge_record(g["cr_by_label"], rec)
        for rec in sw_records:
            _merge_record(g["sw_by_label"], rec)
    return groups


# ─── GENERIC ENTITY COMPARATOR ───────────────────────────────────────────────

def compare_entity_set(
    category, word_items, excel_items, key_fn_word, key_fn_excel,
    designation_fn_word=None, designation_fn_excel=None,
    phones_word_fn=None, phones_excel_fn=None,
    emails_word_fn=None, emails_excel_fn=None,
    label_word_fn=None, label_excel_fn=None,
    report_word_missing=True,
):
    """Generic set-diff + field-diff comparator shared by KMP, Utility
    Heads, Control Rooms, Switchyards, Hospitals, Emergency Services, and
    Reference Sections -- all follow the same shape (a flat list of Word
    records and a flat list of Excel records within one already-matched
    scope, matched by a normalized key, then diffed field by field)."""
    word_by_key = {}
    for item in word_items:
        k = key_fn_word(item)
        if k:
            word_by_key.setdefault(k, item)

    excel_by_key = {}
    for item in excel_items:
        k = key_fn_excel(item)
        if k:
            excel_by_key.setdefault(k, item)

    rows_missing, rows_desig, rows_phone, rows_email, rows_extra = [], [], [], [], []

    for k, w in word_by_key.items():
        label = label_word_fn(w) if label_word_fn else str(k)
        e = excel_by_key.get(k)
        if e is None:
            if report_word_missing:
                rows_missing.append(make_row(
                    category, label, "Not present", label,
                    f"{category} listed in Word but not found in Excel",
                    f"Add '{label}' to the workbook",
                ))
            continue

        if designation_fn_word and designation_fn_excel:
            wd, ed = designation_fn_word(w), designation_fn_excel(e)
            if wd and ed and wd.lower() != ed.lower():
                rows_desig.append(make_row(
                    category, f"{label} — {wd}", ed, wd,
                    "Designation in Excel does not match Word",
                    f"Update designation to '{wd}'",
                ))

        if phones_word_fn and phones_excel_fn:
            wp, ep = phones_word_fn(w), phones_excel_fn(e)
            if wp and ep and wp != ep:
                rows_phone.append(make_row(
                    category, label, ", ".join(sorted(ep)), ", ".join(sorted(wp)),
                    "Phone number(s) do not match between Word and Excel",
                    f"Update phone number(s) to {sorted(wp)}",
                ))
            elif wp and not ep:
                rows_phone.append(make_row(
                    category, label, "(missing)", ", ".join(sorted(wp)),
                    "Phone number present in Word but missing in Excel",
                    f"Add phone number(s) {sorted(wp)}",
                ))

        if emails_word_fn and emails_excel_fn:
            we, ee = emails_word_fn(w), emails_excel_fn(e)
            if we and ee and we != ee:
                rows_email.append(make_row(
                    category, label, ", ".join(sorted(ee)), ", ".join(sorted(we)),
                    "Email does not match between Word and Excel",
                    f"Update email to {sorted(we)}",
                ))
            elif we and not ee:
                rows_email.append(make_row(
                    category, label, "(missing)", ", ".join(sorted(we)),
                    "Email present in Word but missing in Excel",
                    f"Add email {sorted(we)}",
                ))

    for k, e in excel_by_key.items():
        if k not in word_by_key:
            label = label_excel_fn(e) if label_excel_fn else str(k)
            rows_extra.append(make_row(
                category, "Not found in Word", label, "(not present)",
                f"{category} present in Excel but not found anywhere in the Word directory",
                "Manually verify — may be outdated, or a name-matching mismatch",
            ))

    return {
        "missing": rows_missing, "designation": rows_desig,
        "phone": rows_phone, "email": rows_email, "extra_needs_manual": rows_extra,
    }


def word_phone_set(item):
    return frozenset().union(*(normalize_phone_set(p) for p in item["phones"])) if item["phones"] else frozenset()


def word_email_set(item):
    return frozenset().union(*(normalize_email_set(e) for e in item["emails"])) if item["emails"] else frozenset()


# ─── ORGANIZATIONS ───────────────────────────────────────────────────────────

def compare_organizations(groups, excel):
    matched_org_ids, matched_suborg_ids, unresolved_stations = set(), set(), []

    for (match_type, match_id), g in groups.items():
        if match_type == "org":
            matched_org_ids.add(match_id)
        elif match_type == "suborg":
            matched_suborg_ids.add(match_id)
            parent = excel["suborg_parent_by_id"].get(match_id)
            if parent is not None:
                matched_org_ids.add(parent)
        else:
            unresolved_stations.append(g["display_name"])

    rows_missing = [
        make_row(
            "Organization", name, "Not found", name,
            "Organization/station named in Word could not be matched to any organization or "
            "sub-organization in Excel",
            f"Add '{name}' to organizations/sub_organizations, or verify the naming matches",
        )
        for name in sorted(set(unresolved_stations))
    ]

    rows_extra = []
    for o in excel["organizations"]:
        has_direct = o["org_id"] in matched_org_ids
        has_via_suborg = any(
            s["org_id"] == o["org_id"] and s["suborg_id"] in matched_suborg_ids
            for s in excel["sub_organizations"]
        )
        if not has_direct and not has_via_suborg:
            rows_extra.append(make_row(
                "Organization", "Not found in Word", o["name"], "(not present)",
                "Organization present in Excel but not referenced anywhere in the Word directory",
                "Manually verify — may be outdated or a naming mismatch",
            ))

    return {"missing": rows_missing, "extra_needs_manual": rows_extra}


# ─── EMPLOYEES ────────────────────────────────────────────────────────────────

def compare_employees(groups, excel):
    emp_by_suborg, emp_by_org_direct = defaultdict(list), defaultdict(list)
    for e in excel["employees"]:
        (emp_by_suborg[e["suborg_id"]] if e["suborg_id"] is not None else emp_by_org_direct[e["org_id"]]).append(e)

    word_station_meta = {}
    word_name_index = defaultdict(set)
    for (match_type, match_id), g in groups.items():
        word_station_meta[g["display_name"]] = (match_type, match_id)
        for lower in g["employees"]:
            word_name_index[lower].add(g["display_name"])

    all_missing, all_desig, all_phone, all_email = [], [], [], []
    matched_excel_ids = set()

    for (match_type, match_id), g in groups.items():
        if match_type == "none":
            continue
        excel_list = emp_by_suborg.get(match_id, []) if match_type == "suborg" else emp_by_org_direct.get(match_id, [])
        station = g["display_name"]

        result = compare_entity_set(
            f"Employee ({station})", list(g["employees"].values()), excel_list,
            key_fn_word=lambda w: w["name"].lower(), key_fn_excel=lambda e: e["name"].lower(),
            designation_fn_word=lambda w: w["designation"], designation_fn_excel=lambda e: e["designation"],
            phones_word_fn=word_phone_set, phones_excel_fn=lambda e: normalize_phone_set(e.get("phones_raw", "")),
            emails_word_fn=word_email_set, emails_excel_fn=lambda e: normalize_email_set(e.get("emails_raw", "")),
            label_word_fn=lambda w, station=station: f"{station}: {w['name']}",
            label_excel_fn=lambda e, station=station: f"{station}: {e['name']}",
        )
        all_missing.extend(result["missing"])
        all_desig.extend(result["designation"])
        all_phone.extend(result["phone"])
        all_email.extend(result["email"])

        word_names_lower = set(g["employees"].keys())
        for e in excel_list:
            if e["name"].lower() in word_names_lower:
                matched_excel_ids.add(e["employee_id"])

    rows_wrong_org, rows_wrong_suborg, rows_needs_manual = [], [], []
    for e in excel["employees"]:
        if e["employee_id"] in matched_excel_ids:
            continue
        stations = sorted(word_name_index.get(e["name"].lower(), []))
        if len(stations) == 1:
            station = stations[0]
            mtype, mid = word_station_meta.get(station, ("none", None))
            # Word is sometimes less granular than Excel -- one flat heading
            # covering several Excel sub-organizations under the same
            # parent. That's not a wrong mapping, just missing sub-org
            # detail in the source document, so only flag a genuine
            # mismatch: the matched org itself differs, or Word resolves to
            # a *different* specific sub-organization than Excel recorded.
            if mtype == "org" and mid == e["org_id"]:
                continue
            if mtype == "suborg" and mid == e["suborg_id"]:
                continue
            current_org_name = excel["org_name_by_id"].get(e["org_id"], "?")
            current_suborg_name = excel["suborg_name_by_id"].get(e["suborg_id"], "—") if e["suborg_id"] else "—"
            row = make_row(
                "Employee", f"{station}: {e['name']}",
                f"{current_org_name} / {current_suborg_name} (org_id={e['org_id']}, suborg_id={e['suborg_id']})", station,
                "Employee's Word station does not match their Excel organization mapping",
                f"Re-map '{e['name']}' to '{station}'",
            )
            (rows_wrong_suborg if mtype == "suborg" else rows_wrong_org).append(row)
        elif len(stations) == 0:
            rows_needs_manual.append(make_row(
                "Employee", "(not found)", e["name"], "(unknown)",
                "Employee in Excel not found anywhere in the Word directory",
                "Manually verify against the Word directory",
            ))
        else:
            rows_needs_manual.append(make_row(
                "Employee", "; ".join(stations), e["name"], "(ambiguous)",
                f"Name appears under multiple Word stations ({', '.join(stations)}); "
                "cannot confidently determine the correct mapping",
                "Manually verify which station this employee belongs to",
            ))

    return {
        "missing": all_missing, "wrong_org": rows_wrong_org, "wrong_suborg": rows_wrong_suborg,
        "designation": all_desig, "phone": all_phone, "email": all_email, "needs_manual": rows_needs_manual,
    }


# ─── KMP / UTILITY HEADS (Synopsis-sourced, curated subsets) ────────────────

def compare_kmp(word_kmp_candidates, excel_kmp):
    return compare_entity_set(
        "KMP", word_kmp_candidates, excel_kmp,
        key_fn_word=lambda w: w["name"].lower(), key_fn_excel=lambda e: e["name"].lower(),
        designation_fn_word=lambda w: w["designation"], designation_fn_excel=lambda e: e["designation"],
        phones_word_fn=word_phone_set,
        phones_excel_fn=lambda e: normalize_phone_set(" ".join([e["office_phone"], e["residence_phone"], e["mobile_phone"]])),
        emails_word_fn=word_email_set, emails_excel_fn=lambda e: normalize_email_set(e["email"]),
        label_word_fn=lambda w: f"{w['org_banner']}: {w['name']} — {w['designation']}",
        label_excel_fn=lambda e: f"{e['name']} — {e['designation']}",
        report_word_missing=False,  # KMP is a curated subset of Synopsis, not "everyone listed"
    )


def word_employees_by_org(groups, excel):
    """Aggregates every matched Word station's employees under the
    organization they ultimately belong to (a suborg's employees roll up to
    its parent org), for Utility Head lookup -- Excel's Utility_Heads sheet
    is keyed by organization_id only, not suborg."""
    by_org = defaultdict(list)
    for (match_type, match_id), g in groups.items():
        if match_type == "org":
            org_id = match_id
        elif match_type == "suborg":
            org_id = excel["suborg_parent_by_id"].get(match_id)
        else:
            continue
        if org_id is not None:
            by_org[org_id].extend(g["employees"].values())
    return by_org


def compare_utility_heads(groups, excel):
    """Utility Heads turned out NOT to be sourced from the Synopsis table
    (verified: Excel's org_id=1 utility head "Manoj Kumar Agrawal" /
    "Executive Director" matches NLDC and RLDCs' own roster table, not the
    Synopsis section) -- each org's utility head is drawn from that same
    org's own Word roster, matching this app's live "Utility Head" concept
    (the senior-most person per organization)."""
    by_org = word_employees_by_org(groups, excel)
    rows_desig, rows_phone, rows_email, rows_extra = [], [], [], []

    for uh in excel["utility_heads"]:
        candidates = by_org.get(uh["org_id"], [])
        match = next((c for c in candidates if c["name"].lower() == uh["name"].lower()), None)
        label = f"{uh['name']} — {uh['designation']}"

        if match is None:
            rows_extra.append(make_row(
                "Utility Head", "Not found in Word", label, "(not present)",
                "Utility Head present in Excel but not found in this organization's Word roster",
                "Manually verify — may be outdated, or a name/organization mismatch",
            ))
            continue

        if match["designation"] and uh["designation"] and match["designation"].lower() != uh["designation"].lower():
            rows_desig.append(make_row(
                "Utility Head", f"{label} — Word: {match['designation']}", uh["designation"], match["designation"],
                "Designation in Excel does not match Word", f"Update designation to '{match['designation']}'",
            ))

        wp, ep = word_phone_set(match), normalize_phone_set(" ".join([uh["office_phone"], uh["residence_phone"], uh["mobile_phone"]]))
        if wp and ep and wp != ep:
            rows_phone.append(make_row(
                "Utility Head", label, ", ".join(sorted(ep)), ", ".join(sorted(wp)),
                "Phone number(s) do not match between Word and Excel", f"Update phone number(s) to {sorted(wp)}",
            ))
        elif wp and not ep:
            rows_phone.append(make_row(
                "Utility Head", label, "(missing)", ", ".join(sorted(wp)),
                "Phone number present in Word but missing in Excel", f"Add phone number(s) {sorted(wp)}",
            ))

        we, ee = word_email_set(match), normalize_email_set(uh["email"])
        if we and ee and we != ee:
            rows_email.append(make_row(
                "Utility Head", label, ", ".join(sorted(ee)), ", ".join(sorted(we)),
                "Email does not match between Word and Excel", f"Update email to {sorted(we)}",
            ))
        elif we and not ee:
            rows_email.append(make_row(
                "Utility Head", label, "(missing)", ", ".join(sorted(we)),
                "Email present in Word but missing in Excel", f"Add email {sorted(we)}",
            ))

    return {"missing": [], "designation": rows_desig, "phone": rows_phone, "email": rows_email, "extra_needs_manual": rows_extra}


# ─── CONTROL ROOMS / SWITCHYARDS ─────────────────────────────────────────────

def compare_control_rooms_switchyards(groups, excel):
    cr_by_suborg, cr_by_org = defaultdict(list), defaultdict(list)
    for c in excel["control_rooms"]:
        (cr_by_suborg[c["suborg_id"]] if c["suborg_id"] is not None else cr_by_org[c["org_id"]]).append(c)
    sw_by_suborg, sw_by_org = defaultdict(list), defaultdict(list)
    for s in excel["switchyards"]:
        (sw_by_suborg[s["suborg_id"]] if s["suborg_id"] is not None else sw_by_org[s["org_id"]]).append(s)

    cr_out = {"missing": [], "phone": [], "email": [], "needs_manual": []}
    sw_out = {"missing": [], "phone": [], "email": [], "needs_manual": []}

    for (match_type, match_id), g in groups.items():
        if match_type == "none":
            continue
        station = g["display_name"]
        excel_cr = merge_excel_records_by_label(
            cr_by_suborg.get(match_id, []) if match_type == "suborg" else cr_by_org.get(match_id, [])
        )
        excel_sw = merge_excel_records_by_label(
            sw_by_suborg.get(match_id, []) if match_type == "suborg" else sw_by_org.get(match_id, [])
        )

        cr_result = compare_entity_set(
            f"Control Room ({station})", list(g["cr_by_label"].values()), excel_cr,
            key_fn_word=lambda w: w["label"].strip().lower(), key_fn_excel=lambda e: e["label"].strip().lower(),
            phones_word_fn=lambda w: normalize_phone_set(" ".join(w["phones"])),
            phones_excel_fn=lambda e: normalize_phone_set(e["phones_raw"]),
            emails_word_fn=lambda w: normalize_email_set(" ".join(w["emails"])),
            emails_excel_fn=lambda e: normalize_email_set(e["emails_raw"]),
            label_word_fn=lambda w, station=station: f"{station}: {w['label']}",
            label_excel_fn=lambda e, station=station: f"{station}: {e['label']}",
        )
        for k in ("missing", "phone", "email"):
            cr_out[k].extend(cr_result[k])
        cr_out["needs_manual"].extend(cr_result["extra_needs_manual"])

        sw_result = compare_entity_set(
            f"Switchyard ({station})", list(g["sw_by_label"].values()), excel_sw,
            key_fn_word=lambda w: w["label"].strip().lower(), key_fn_excel=lambda e: e["label"].strip().lower(),
            phones_word_fn=lambda w: normalize_phone_set(" ".join(w["phones"])),
            phones_excel_fn=lambda e: normalize_phone_set(e["phones_raw"]),
            emails_word_fn=lambda w: normalize_email_set(" ".join(w["emails"])),
            emails_excel_fn=lambda e: normalize_email_set(e["emails_raw"]),
            label_word_fn=lambda w, station=station: f"{station}: {w['label']}",
            label_excel_fn=lambda e, station=station: f"{station}: {e['label']}",
        )
        for k in ("missing", "phone", "email"):
            sw_out[k].extend(sw_result[k])
        sw_out["needs_manual"].extend(sw_result["extra_needs_manual"])

    return cr_out, sw_out


# ─── HOSPITALS / EMERGENCY SERVICES ──────────────────────────────────────────

def compare_hospitals(word_hospitals, excel_hospitals):
    def word_phones(w):
        return normalize_phone_set(" ".join([w["contact_number"], w["opd_number"], w["emergency_number"]]))

    def excel_phones(e):
        return normalize_phone_set(" ".join([e["contact_number"], e["opd_number"], e["emergency_number"]]))

    return compare_entity_set(
        "Hospital", word_hospitals, excel_hospitals,
        key_fn_word=lambda w: w["name"].lower(), key_fn_excel=lambda e: e["name"].lower(),
        phones_word_fn=word_phones, phones_excel_fn=excel_phones,
        label_word_fn=lambda w: w["name"], label_excel_fn=lambda e: e["name"],
    )


def compare_emergency_services(word_emergency_services, excel_emergency_services):
    return compare_entity_set(
        "Emergency Service", word_emergency_services, excel_emergency_services,
        key_fn_word=lambda w: (w.get("service_type", "") or "").lower() + "|" + w["name"].lower(),
        key_fn_excel=lambda e: (e["service_type"] or "").lower() + "|" + e["name"].lower(),
        phones_word_fn=lambda w: normalize_phone_set(w["contact_number"]),
        phones_excel_fn=lambda e: normalize_phone_set(e["contact_number"]),
        label_word_fn=lambda w: f"{w.get('service_type', '')}: {w['name']}",
        label_excel_fn=lambda e: f"{e['service_type']}: {e['name']}",
    )


# ─── REFERENCE SECTIONS (best effort) ────────────────────────────────────────

def compare_reference_sections(word_reference_rows, unresolved_reference_sections, excel_reference_sections):
    rows_missing, rows_needs_manual = [], []

    excel_by_section = defaultdict(list)
    for r in excel_reference_sections:
        excel_by_section[r["section"]].append(r)

    all_sections = set(REFERENCE_HEADINGS) | set(excel_by_section.keys())
    for heading_text in sorted(all_sections):
        excel_rows = excel_by_section.get(heading_text, [])

        if heading_text in unresolved_reference_sections or heading_text not in word_reference_rows:
            for e in excel_rows:
                rows_needs_manual.append(make_row(
                    "Reference Section", heading_text, e["name"],
                    "(could not confidently parse this section's table structure from Word)",
                    "This back-matter section's Word table shape could not be automatically "
                    "matched with confidence",
                    "Manually verify against the Word directory",
                ))
            continue

        result = compare_entity_set(
            f"Reference ({heading_text})", word_reference_rows[heading_text], excel_rows,
            key_fn_word=lambda w: w["name"].lower(), key_fn_excel=lambda e: e["name"].lower(),
            phones_word_fn=word_phone_set, phones_excel_fn=lambda e: normalize_phone_set(e["phones_raw"]),
            emails_word_fn=word_email_set, emails_excel_fn=lambda e: normalize_email_set(e["email"]),
            label_word_fn=lambda w, h=heading_text: f"{h}: {w['name']}",
            label_excel_fn=lambda e, h=heading_text: f"{h}: {e['name']}",
        )
        rows_missing.extend(result["missing"])
        # low-confidence category throughout -- downgrade everything else to
        # manual verification rather than asserting a hard mismatch
        rows_needs_manual.extend(result["extra_needs_manual"])
        rows_needs_manual.extend(result["phone"])
        rows_needs_manual.extend(result["email"])

    return {"missing": rows_missing, "needs_manual": rows_needs_manual}


# ─── DUPLICATES ───────────────────────────────────────────────────────────────

def describe_ref(ref, excel):
    ref_type, ref_id = ref
    lookup = {
        "employee": excel["employees_by_id"], "kmp": excel["kmp_by_id"],
        "utility_head": excel["uh_by_id"], "control_room": excel["cr_by_id"],
        "switchyard": excel["sw_by_id"], "hospital": excel["hosp_by_id"],
    }
    rec = lookup.get(ref_type, {}).get(ref_id)
    if rec is None:
        return f"{ref_type}#{ref_id}"
    return f"{ref_type}:{rec.get('name') or rec.get('label') or '?'}"


def find_duplicates(excel):
    rows = []

    seen = defaultdict(list)
    for o in excel["organizations"]:
        seen[o["name"].lower()].append(o)
    for items in seen.values():
        if len(items) > 1:
            rows.append(make_row(
                "Organization", "-", "org_id: " + ", ".join(str(i["org_id"]) for i in items), items[0]["name"],
                f"{len(items)} organizations share the name '{items[0]['name']}'",
                "Merge or rename the duplicate organization records",
            ))

    seen = defaultdict(list)
    for e in excel["employees"]:
        seen[(e["name"].lower(), e["org_id"])].append(e)
    for (_name_lower, org_id), items in seen.items():
        if len(items) > 1:
            rows.append(make_row(
                "Employee", "-", f"{len(items)} rows in org_id={org_id}", items[0]["name"],
                f"{len(items)} employee records share the name '{items[0]['name']}' in the same organization",
                "Verify whether these are duplicate entries or distinct people sharing a name",
            ))

    for entity_name, items_list in (("Control Room", excel["control_rooms"]), ("Switchyard", excel["switchyards"])):
        seen = defaultdict(list)
        for c in items_list:
            seen[(c["label"].lower(), c["org_id"], c["suborg_id"])].append(c)
        for items in seen.values():
            if len(items) > 1:
                rows.append(make_row(
                    entity_name, "-", f"{len(items)} rows", items[0]["label"],
                    f"{len(items)} {entity_name.lower()} records share the same name in the same organization",
                    "Merge or verify these duplicate records",
                ))

    # Indian STD/area codes are 2-6 digits and legitimately repeat across
    # every phone number in the same region (e.g. "02624-244120" contributes
    # the bare digit-run "02624") -- only digit runs long enough to be a
    # complete, individually-identifying number (a 10-digit mobile, or an
    # STD code + local number combined) are checked for genuine duplication.
    phone_owners = defaultdict(set)
    for (ref_type, ref_id), numbers in excel["phones_by_ref"].items():
        for n in numbers:
            for d in normalize_phone_set(n):
                if len(d) >= 8:
                    phone_owners[d].add((ref_type, ref_id))
    for digits, owners in phone_owners.items():
        if len(owners) > 1:
            rows.append(make_row(
                "Phone Number", "-", digits, "; ".join(sorted(describe_ref(o, excel) for o in owners)),
                f"Phone number {digits} is shared by {len(owners)} different records",
                "Verify whether this is a genuinely shared line or a data entry error",
            ))

    email_owners = defaultdict(set)
    for (ref_type, ref_id), emails in excel["emails_by_ref"].items():
        for em in emails:
            key = em.strip().lower()
            if key:
                email_owners[key].add((ref_type, ref_id))
    for email_key, owners in email_owners.items():
        if len(owners) > 1:
            rows.append(make_row(
                "Email Address", "-", email_key, "; ".join(sorted(describe_ref(o, excel) for o in owners)),
                f"Email address {email_key} is shared by {len(owners)} different records",
                "Verify whether this is a genuinely shared mailbox or a data entry error",
            ))

    return rows


# ─── HIERARCHY ────────────────────────────────────────────────────────────────

def check_hierarchy(excel):
    rows = []

    for e in excel["employees"]:
        if e["suborg_id"] is None:
            continue
        parent = excel["suborg_parent_by_id"].get(e["suborg_id"])
        if parent is not None and e["org_id"] is not None and parent != e["org_id"]:
            rows.append(make_row(
                "Hierarchy", "-", f"employee={e['name']} org_id={e['org_id']} suborg_id={e['suborg_id']}",
                f"suborg {e['suborg_id']}'s parent organization is org_id={parent}",
                "Employee's org_id does not match their sub-organization's actual parent organization",
                f"Correct org_id to {parent} (or fix the sub-organization mapping)",
            ))

    for entity_name, items in (("Control Room", excel["control_rooms"]), ("Switchyard", excel["switchyards"])):
        for c in items:
            if c["suborg_id"] is None:
                continue
            parent = excel["suborg_parent_by_id"].get(c["suborg_id"])
            if parent is not None and c["org_id"] is not None and parent != c["org_id"]:
                rows.append(make_row(
                    "Hierarchy", "-", f"{entity_name}={c['label']} org_id={c['org_id']} suborg_id={c['suborg_id']}",
                    f"suborg {c['suborg_id']}'s parent organization is org_id={parent}",
                    f"{entity_name}'s org_id does not match their sub-organization's actual parent organization",
                    f"Correct org_id to {parent}",
                ))

    for s in excel["sub_organizations"]:
        if s["org_id"] not in excel["org_name_by_id"]:
            rows.append(make_row(
                "Hierarchy", "-", f"suborg={s['name']} org_id={s['org_id']}", "(a valid org_id is expected)",
                "Sub-organization references an org_id that does not exist in the organizations sheet",
                "Fix the sub-organization's org_id foreign key",
            ))

    return rows


# ─── SUMMARY ──────────────────────────────────────────────────────────────────

def compute_summary(excel, r, word_totals):
    totals = {
        "Organizations": len(excel["organizations"]),
        "Sub Organizations": len(excel["sub_organizations"]),
        "Employees": len(excel["employees"]),
        "KMP": len(excel["kmp"]),
        "Utility Heads": len(excel["utility_heads"]),
        "Control Rooms": len(excel["control_rooms"]),
        "Switchyards": len(excel["switchyards"]),
        "Hospitals": len(excel["hospitals"]),
        "Emergency Services": len(excel["emergency_services"]),
    }
    total_excel_records = sum(totals.values())

    incorrect = (
        len(r["employees"]["wrong_org"]) + len(r["employees"]["wrong_suborg"])
        + len(r["employees"]["designation"]) + len(r["employees"]["phone"]) + len(r["employees"]["email"])
        + len(r["kmp"]["designation"]) + len(r["kmp"]["phone"]) + len(r["kmp"]["email"])
        + len(r["utility_heads"]["designation"]) + len(r["utility_heads"]["phone"]) + len(r["utility_heads"]["email"])
        + len(r["cr"]["phone"]) + len(r["cr"]["email"])
        + len(r["sw"]["phone"]) + len(r["sw"]["email"])
        + len(r["hospitals"]["phone"])
        + len(r["emergency"]["phone"])
    )
    missing = (
        len(r["organizations"]["missing"]) + len(r["employees"]["missing"])
        + len(r["cr"]["missing"]) + len(r["sw"]["missing"])
        + len(r["hospitals"]["missing"]) + len(r["emergency"]["missing"])
        + len(r["reference"]["missing"])
    )
    extra_and_manual = (
        len(r["organizations"]["extra_needs_manual"]) + len(r["employees"]["needs_manual"])
        + len(r["kmp"]["extra_needs_manual"]) + len(r["utility_heads"]["extra_needs_manual"])
        + len(r["cr"]["needs_manual"]) + len(r["sw"]["needs_manual"])
        + len(r["hospitals"]["extra_needs_manual"]) + len(r["emergency"]["extra_needs_manual"])
        + len(r["reference"]["needs_manual"])
    )
    duplicates = len(r["duplicates"])
    hierarchy_errors = len(r["hierarchy"])

    correct = max(total_excel_records - incorrect - missing - extra_and_manual, 0)
    denom = correct + incorrect + extra_and_manual
    accuracy_pct = round(100 * correct / denom, 2) if denom else 0.0

    word_total_records = word_totals["station"] + word_totals["synopsis"] + word_totals["emergency"]
    parsing_confidence_denom = word_total_records + word_totals["reference_confident"] + word_totals["reference_unconfident"]
    parsing_confidence_pct = (
        round(100 * (word_total_records + word_totals["reference_confident"]) / parsing_confidence_denom, 2)
        if parsing_confidence_denom else 100.0
    )

    return {
        "totals": totals, "total_excel_records": total_excel_records,
        "correct": correct, "incorrect": incorrect, "missing": missing,
        "extra_and_manual": extra_and_manual, "duplicates": duplicates, "hierarchy_errors": hierarchy_errors,
        "accuracy_pct": accuracy_pct, "parsing_confidence_pct": parsing_confidence_pct,
    }


# ─── REPORT WRITER ────────────────────────────────────────────────────────────

def write_discrepancy_sheet(wb, name, rows):
    ws = wb.create_sheet(name)
    ws.append(DISCREPANCY_HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F6F5C")
    for row in rows:
        ws.append([row.get(h, "") for h in DISCREPANCY_HEADER])
    widths = [22, 42, 30, 30, 46, 40]
    for col, w in zip("ABCDEF", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    return ws


def write_summary_sheet(wb, summary):
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = "Telephone Directory Validation Report"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Word source: {WORD_PATH}"
    ws["A3"] = f"Excel workbook: {EXCEL_PATH}"

    r = 5
    ws.cell(row=r, column=1, value="Category").font = Font(bold=True)
    ws.cell(row=r, column=2, value="Total in Excel").font = Font(bold=True)
    r += 1
    for label, value in summary["totals"].items():
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=value)
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="Final Summary").font = Font(bold=True, size=12)
    r += 1
    metrics = [
        ("Total Excel Records Checked", summary["total_excel_records"]),
        ("Total Correct Records", summary["correct"]),
        ("Total Incorrect Records", summary["incorrect"]),
        ("Total Missing Records (in Word, not in Excel)", summary["missing"]),
        ("Total Extra / Needs-Manual-Verification Records", summary["extra_and_manual"]),
        ("Total Duplicate Records", summary["duplicates"]),
        ("Total Hierarchy Errors", summary["hierarchy_errors"]),
    ]
    for label, value in metrics:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=value)
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="Overall Data Accuracy %").font = Font(bold=True)
    ws.cell(row=r, column=2, value=summary["accuracy_pct"])
    ws.cell(row=r, column=3, value="= Correct / (Correct + Incorrect + Extra) — of what's in Excel, how much matches Word").font = Font(italic=True, size=9)
    r += 1
    ws.cell(row=r, column=1, value="Parsing Confidence %").font = Font(bold=True)
    ws.cell(row=r, column=2, value=summary["parsing_confidence_pct"])
    ws.cell(row=r, column=3, value="= share of Word source records the automated parser could structure with confidence "
                                    "(the rest were routed to Needs_Manual_Verification by design)").font = Font(italic=True, size=9)
    r += 2

    confidence_note = (
        "Confidence tiers: HIGH — Organizations, Employees, Control Rooms, Switchyards, Hospitals, "
        "Police/Fire/Blood-Bank/etc. emergency services (structured, header-driven Word tables). "
        "MEDIUM — KMP, Utility Heads: verified against the Synopsis table's factual content; Word does not "
        "explicitly label who counts as KMP or Utility Head, so only entries already in the Excel sheets are "
        "fact-checked, not the selection itself. LOW / BEST EFFORT — Reference_Sections (RRAS, Disaster "
        "Management, Nodal Officers, Hotlines, Vendors): heterogeneous free-text back matter; low-confidence "
        "rows are routed to Needs_Manual_Verification rather than guessed."
    )
    ws.cell(row=r, column=1, value=confidence_note).alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 90

    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 70
    return ws


def run():
    word_mtime_before = os.path.getmtime(WORD_PATH)
    excel_mtime_before = os.path.getmtime(EXCEL_PATH)

    print("Loading Excel workbook...")
    excel = load_excel(EXCEL_PATH)

    print("Parsing Word document...")
    word = parse_word(WORD_PATH)

    print("Matching Word stations to Excel organizations...")
    groups = group_word_segments(word["all_station_segments"], excel)

    print("Comparing organizations...")
    r_orgs = compare_organizations(groups, excel)

    print("Comparing employees...")
    r_employees = compare_employees(groups, excel)

    print("Comparing KMP / Utility Heads...")
    r_kmp = compare_kmp(word["word_kmp_candidates"], excel["kmp"])
    r_uh = compare_utility_heads(groups, excel)

    print("Comparing control rooms / switchyards...")
    r_cr, r_sw = compare_control_rooms_switchyards(groups, excel)

    print("Comparing hospitals / emergency services...")
    r_hosp = compare_hospitals(word["word_hospitals"], excel["hospitals"])
    r_emerg = compare_emergency_services(word["word_emergency_services"], excel["emergency_services"])

    print("Comparing reference sections (best effort)...")
    r_ref = compare_reference_sections(
        word["word_reference_rows"], word["unresolved_reference_sections"], excel["reference_sections"]
    )

    print("Finding duplicates...")
    r_dupes = find_duplicates(excel)

    print("Checking hierarchy...")
    r_hier = check_hierarchy(excel)

    results = {
        "organizations": r_orgs, "employees": r_employees, "kmp": r_kmp, "utility_heads": r_uh,
        "cr": r_cr, "sw": r_sw, "hospitals": r_hosp, "emergency": r_emerg, "reference": r_ref,
        "duplicates": r_dupes, "hierarchy": r_hier,
    }
    summary = compute_summary(excel, results, word["totals"])

    print("Writing report...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    write_summary_sheet(wb, summary)
    write_discrepancy_sheet(wb, "Missing_Organizations", r_orgs["missing"])
    write_discrepancy_sheet(wb, "Missing_Employees", r_employees["missing"])
    write_discrepancy_sheet(wb, "Wrong_Organization_Mapping", r_employees["wrong_org"])
    write_discrepancy_sheet(wb, "Wrong_SubOrganization_Mapping", r_employees["wrong_suborg"])
    write_discrepancy_sheet(wb, "Wrong_Designation", r_employees["designation"] + r_kmp["designation"] + r_uh["designation"])
    write_discrepancy_sheet(
        wb, "Phone_Mismatches",
        r_employees["phone"] + r_kmp["phone"] + r_uh["phone"] + r_cr["phone"] + r_sw["phone"] + r_hosp["phone"] + r_emerg["phone"],
    )
    write_discrepancy_sheet(
        wb, "Email_Mismatches",
        r_employees["email"] + r_kmp["email"] + r_uh["email"] + r_cr["email"] + r_sw["email"],
    )
    write_discrepancy_sheet(wb, "Missing_KMP", r_kmp["missing"])
    write_discrepancy_sheet(
        wb, "Utility_Head_Errors",
        r_uh["missing"] + r_uh["designation"] + r_uh["phone"] + r_uh["email"] + r_uh["extra_needs_manual"],
    )
    write_discrepancy_sheet(wb, "Control_Room_Errors", r_cr["missing"] + r_cr["phone"] + r_cr["email"])
    write_discrepancy_sheet(wb, "Switchyard_Errors", r_sw["missing"] + r_sw["phone"] + r_sw["email"])
    write_discrepancy_sheet(wb, "Duplicate_Records", r_dupes)
    write_discrepancy_sheet(wb, "Hierarchy_Errors", r_hier)
    write_discrepancy_sheet(
        wb, "Needs_Manual_Verification",
        r_orgs["extra_needs_manual"] + r_employees["needs_manual"]
        + r_kmp["extra_needs_manual"] + r_uh["extra_needs_manual"]
        + r_cr["needs_manual"] + r_sw["needs_manual"]
        + r_hosp["extra_needs_manual"] + r_emerg["extra_needs_manual"]
        + r_ref["missing"] + r_ref["needs_manual"],
    )

    wb.save(OUTPUT_XLSX)

    assert os.path.getmtime(WORD_PATH) == word_mtime_before, "Word source file was modified!"
    assert os.path.getmtime(EXCEL_PATH) == excel_mtime_before, "Excel source file was modified!"

    print("=" * 70)
    print("  Validation Report Summary")
    print("=" * 70)
    for label, value in summary["totals"].items():
        print(f"  {label:<22}: {value}")
    print(f"  Total Excel Records Checked : {summary['total_excel_records']}")
    print(f"  Correct                     : {summary['correct']}")
    print(f"  Incorrect                   : {summary['incorrect']}")
    print(f"  Missing                     : {summary['missing']}")
    print(f"  Extra / Needs Manual        : {summary['extra_and_manual']}")
    print(f"  Duplicates                  : {summary['duplicates']}")
    print(f"  Hierarchy errors            : {summary['hierarchy_errors']}")
    print(f"  Overall Data Accuracy %     : {summary['accuracy_pct']}")
    print(f"  Parsing Confidence %        : {summary['parsing_confidence_pct']}")
    print("=" * 70)
    print(f"  Report written to: {OUTPUT_XLSX}")


if __name__ == "__main__":
    run()
