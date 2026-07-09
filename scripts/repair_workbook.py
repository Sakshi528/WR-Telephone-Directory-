"""Repairs uploads/WR_DB_Ready_Final.xlsx into uploads/WR_DB_Ready_Final_Verified.xlsx.
The source is never opened in write mode -- only the copy is edited, and
every change is logged to reports/repair_change_log.csv.
"""

import csv
import os
import shutil
import sys
from collections import defaultdict

import openpyxl
import docx
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.compare_word_excel import (
    clean,
    strip_leading_number,
    iter_block_items,
    find_header_columns,
    is_header_row,
    NUMBERED_TITLE_RE,
    SWITCHYARD_RE,
    PHONE_RE,
    EMAIL_RE,
    dedup_preserve_order as _dedup,
    classify_row_content as _classify_row_content,
)

SOURCE_PATH = "uploads/WR_DB_Ready_Final.xlsx"
OUTPUT_PATH = "uploads/WR_DB_Ready_Final_Verified.xlsx"
CHANGE_LOG_PATH = "reports/repair_change_log.csv"
WORD_PATH = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"
DECISIONS_CSV = "config/manual_reconciliation_decisions.csv"

EMP_COL_ID, EMP_COL_ORG, EMP_COL_SUBORG, EMP_COL_NAME, EMP_COL_DESIG = 1, 2, 3, 4, 5
CR_COL_ID, CR_COL_ORG, CR_COL_SUBORG, CR_COL_ORG_NAME, CR_COL_SUBORG_NAME = 1, 2, 3, 4, 5


def parse_word_employee_roles(path):
    """Returns dict: (name.lower(), designation.lower()) -> set of distinct
    station names Word lists that exact person+role under."""
    doc = docx.Document(path)
    lookup = defaultdict(set)

    def process_table(table, fallback_name):
        rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
        name_col, desig_col = find_header_columns(rows_cells)
        if name_col is None:
            return  # not a roster table

        current_name = None
        for cells in rows_cells:
            if not cells:
                continue
            first_cell = cells[0]

            if NUMBERED_TITLE_RE.match(first_cell):
                current_name = strip_leading_number(first_cell)
                continue

            if is_header_row(cells):
                continue

            if not any(cells):
                continue

            name_cell = cells[name_col] if name_col < len(cells) else ""
            desig_cell = cells[desig_col] if desig_col < len(cells) else ""

            if not name_cell:
                continue
            if "control room" in name_cell.lower() or "control room" in desig_cell.lower():
                continue
            if SWITCHYARD_RE.search(name_cell) or SWITCHYARD_RE.search(desig_cell):
                continue
            if name_cell.lower() == desig_cell.lower():
                continue  # repeated title-block noise row

            station = current_name or fallback_name
            if station:
                lookup[(name_cell.lower(), desig_cell.lower())].add(station)

    current_heading = None
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue
        process_table(item, current_heading)

    return lookup


def repair_employee_suborg_mappings(wb, change_log):
    ws = wb["employees"]

    so_ws = wb["sub_organizations"]
    suborg_id_by_name_lower = {}
    for sid, _parent_oid, name, _addr, _state in so_ws.iter_rows(min_row=2, values_only=True):
        if sid is not None and name:
            suborg_id_by_name_lower[strip_leading_number(clean(name)).lower()] = sid

    word_lookup = parse_word_employee_roles(WORD_PATH)

    entries = []
    for row_idx in range(2, ws.max_row + 1):
        name = clean(ws.cell(row_idx, EMP_COL_NAME).value)
        if not name:
            continue
        entries.append({
            "row": row_idx,
            "employee_id": ws.cell(row_idx, EMP_COL_ID).value,
            "suborg_id": ws.cell(row_idx, EMP_COL_SUBORG).value,
            "name": name,
            "designation": clean(ws.cell(row_idx, EMP_COL_DESIG).value),
        })

    by_key = defaultdict(list)
    for e in entries:
        by_key[(e["name"].lower(), e["designation"].lower())].append(e)

    repaired = 0
    manual_review = 0

    for key, group in sorted(by_key.items()):
        stations = word_lookup.get(key)
        if not stations:
            continue  # this role doesn't appear in Word at all -- out of scope

        if len(stations) > 1:
            station_list = ", ".join(sorted(stations))
            for e in group:
                manual_review += 1
                change_log.append({
                    "Sheet": "employees",
                    "Action": "Manual review required",
                    "Details": (f"employee_id={e['employee_id']} name='{e['name']}' "
                                f"designation='{e['designation']}' current_suborg_id={e['suborg_id']} -- "
                                f"Word lists this exact name+designation at {len(stations)} different "
                                f"stations ({station_list}); can't determine which one this row means"),
                })
            continue

        station_name = next(iter(stations))
        resolved_sid = suborg_id_by_name_lower.get(station_name.lower())

        if resolved_sid is None:
            for e in group:
                manual_review += 1
                change_log.append({
                    "Sheet": "employees",
                    "Action": "Manual review required",
                    "Details": (f"employee_id={e['employee_id']} name='{e['name']}' "
                                f"designation='{e['designation']}' current_suborg_id={e['suborg_id']} -- "
                                f"Word shows this role at station '{station_name}', which has no exact "
                                f"match in sub_organizations"),
                })
            continue

        claimed_by = {e["suborg_id"] for e in group if e["suborg_id"] == resolved_sid}
        for e in group:
            if e["suborg_id"] == resolved_sid:
                continue  # already correct

            if claimed_by:
                manual_review += 1
                change_log.append({
                    "Sheet": "employees",
                    "Action": "Manual review required",
                    "Details": (f"employee_id={e['employee_id']} name='{e['name']}' "
                                f"designation='{e['designation']}' current_suborg_id={e['suborg_id']} -- "
                                f"Word shows this role belongs to '{station_name}' (suborg_id={resolved_sid}), "
                                f"but another row for the same name+designation already holds that link -- "
                                f"likely a duplicate row, not a distinct assignment; left unchanged"),
                })
                continue

            old_sid = e["suborg_id"]
            ws.cell(e["row"], EMP_COL_SUBORG).value = resolved_sid
            change_log.append({
                "Sheet": "employees",
                "Action": "Corrected suborg_id (Word-verified)",
                "Details": (f"employee_id={e['employee_id']} name='{e['name']}' "
                            f"designation='{e['designation']}' old_suborg_id={old_sid} "
                            f"new_suborg_id={resolved_sid} station='{station_name}' "
                            f"reason='Word directory uniquely and exactly places this person at this "
                            f"station, which differs from the suborg_id previously stored'"),
            })
            repaired += 1
            claimed_by.add(resolved_sid)

    return repaired, manual_review


def _parse_suborg_id(value):
    if value is None or value == "" or value == "None":
        return None
    return int(value)


def apply_manual_decisions(wb, change_log):
    """Applies human-approved rows from config/manual_reconciliation_decisions.csv.
    Only decision=APPROVED rows are applied; REJECTED/SKIPPED are left alone."""
    if not os.path.exists(DECISIONS_CSV):
        return 0, 0

    with open(DECISIONS_CSV, newline="", encoding="utf-8") as f:
        decisions = list(csv.DictReader(f))

    approved = [d for d in decisions if d["decision"] == "APPROVED"]
    if not approved:
        return 0, 0

    emp_ws = wb["employees"]
    cr_ws = wb["control_rooms"]
    sw_ws = wb["switchyards"]

    applied = 0
    not_found = 0
    rows_to_delete = defaultdict(list)  # sheet name -> [row_idx, ...]

    def find_employee_rows(name, designation, suborg_id):
        matches = []
        for row_idx in range(2, emp_ws.max_row + 1):
            r_name = clean(emp_ws.cell(row_idx, EMP_COL_NAME).value)
            r_desig = clean(emp_ws.cell(row_idx, EMP_COL_DESIG).value)
            if r_name.lower() == name.lower() and r_desig.lower() == designation.lower():
                if emp_ws.cell(row_idx, EMP_COL_SUBORG).value == suborg_id:
                    matches.append(row_idx)
        return matches

    def find_directory_rows(ws, suborg_id):
        matches = []
        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row_idx, CR_COL_SUBORG).value == suborg_id:
                matches.append(row_idx)
        return matches

    for d in approved:
        record_type = d["record_type"]
        name = clean(d["record_name"])
        designation = clean(d["designation"])
        old_sid = _parse_suborg_id(d["old_suborg_id"])
        new_sid = _parse_suborg_id(d["new_suborg_id"])
        action = d["action"]
        note = f"approved_by={d['approved_by']} on {d['date']}; reason: {d['reason']}"

        if action == "ignore":
            change_log.append({
                "Sheet": record_type.lower().replace(" ", "_"),
                "Action": "Manual decision applied: ignore (no change)",
                "Details": f"record_name='{name}' designation='{designation}' -- {note}",
            })
            continue

        if record_type == "Employee" and action == "move_suborg":
            matches = find_employee_rows(name, designation, old_sid)
            if not matches:
                not_found += 1
                change_log.append({
                    "Sheet": "employees",
                    "Action": "Manual decision skipped (row not found)",
                    "Details": (f"record_name='{name}' designation='{designation}' action={action} "
                                f"old_suborg_id={old_sid} new_suborg_id={new_sid} -- no matching "
                                f"workbook row found; data may have changed since approval"),
                })
                continue
            for row_idx in matches:
                emp_ws.cell(row_idx, EMP_COL_SUBORG).value = new_sid
                applied += 1
            change_log.append({
                "Sheet": "employees",
                "Action": "Manual decision applied: moved suborg_id",
                "Details": (f"record_name='{name}' designation='{designation}' "
                            f"old_suborg_id={old_sid} new_suborg_id={new_sid} "
                            f"({len(matches)} row(s)) -- {note}"),
            })

        elif record_type == "Employee" and action == "merge_duplicate":
            matches = find_employee_rows(name, designation, old_sid)
            if len(matches) < 2:
                not_found += 1
                change_log.append({
                    "Sheet": "employees",
                    "Action": "Manual decision skipped (duplicate rows not found)",
                    "Details": (f"record_name='{name}' designation='{designation}' suborg_id={old_sid} "
                                f"-- expected 2+ matching rows, found {len(matches)}; data may have "
                                f"changed since approval"),
                })
                continue
            keep, remove = matches[0], matches[1:]
            for row_idx in remove:
                rows_to_delete["employees"].append(row_idx)
                applied += 1
            change_log.append({
                "Sheet": "employees",
                "Action": "Manual decision applied: merged duplicate",
                "Details": (f"record_name='{name}' designation='{designation}' suborg_id={old_sid} -- "
                            f"kept 1 row, removed {len(remove)} duplicate row(s) -- {note}"),
            })

        elif record_type in ("Control Room", "Switchyard") and action in ("move_control_room", "move_switchyard"):
            ws = cr_ws if record_type == "Control Room" else sw_ws
            sheet_name = "control_rooms" if record_type == "Control Room" else "switchyards"
            matches = find_directory_rows(ws, old_sid)
            if not matches:
                not_found += 1
                change_log.append({
                    "Sheet": sheet_name,
                    "Action": "Manual decision skipped (row not found)",
                    "Details": (f"record_name='{name}' action={action} old_suborg_id={old_sid} "
                                f"new_suborg_id={new_sid} -- no matching workbook row found; data may "
                                f"have changed since approval"),
                })
                continue
            for row_idx in matches:
                ws.cell(row_idx, CR_COL_SUBORG).value = new_sid
                applied += 1
            change_log.append({
                "Sheet": sheet_name,
                "Action": "Manual decision applied: moved suborg_id",
                "Details": (f"record_name='{name}' old_suborg_id={old_sid} new_suborg_id={new_sid} "
                            f"({len(matches)} row(s)) -- {note}"),
            })

    for row_idx in sorted(rows_to_delete["employees"], reverse=True):
        emp_ws.delete_rows(row_idx, 1)

    return applied, not_found


def repair_employees(wb, change_log):
    ws = wb["employees"]

    so_ws = wb["sub_organizations"]
    valid_suborg_ids = {row[0] for row in so_ws.iter_rows(min_row=2, values_only=True) if row[0] is not None}

    entries = []
    skipped = 0
    for row_idx in range(2, ws.max_row + 1):
        name = clean(ws.cell(row_idx, EMP_COL_NAME).value)
        if not name:
            skipped += 1
            continue
        entries.append({
            "row": row_idx,
            "employee_id": ws.cell(row_idx, EMP_COL_ID).value,
            "org_id": ws.cell(row_idx, EMP_COL_ORG).value,
            "suborg_id": ws.cell(row_idx, EMP_COL_SUBORG).value,
            "name": name,
            "designation": clean(ws.cell(row_idx, EMP_COL_DESIG).value),
        })

    by_name_designation = {}
    for e in entries:
        key = (e["name"].lower(), e["designation"].lower())
        by_name_designation.setdefault(key, []).append(e)

    repaired = 0
    manual_review = 0
    rows_to_delete = []

    for (name_lower, designation_lower), group in by_name_designation.items():
        if len(group) < 2:
            continue  # no duplication at all -- nothing to consider

        valid_rows = [e for e in group if e["suborg_id"] is not None and e["suborg_id"] in valid_suborg_ids]
        null_rows = [e for e in group if e["suborg_id"] is None]

        if not null_rows:
            continue  # duplicated name+designation, but no orphan rows to repair

        if len(valid_rows) == 1 and len(null_rows) == 1:
            orphan = null_rows[0]
            valid = valid_rows[0]
            change_log.append({
                "Sheet": "employees",
                "Action": "Removed orphan duplicate employee row",
                "Details": (f"employee_id={orphan['employee_id']} name='{orphan['name']}' "
                            f"designation='{orphan['designation']}' org_id={orphan['org_id']} "
                            f"suborg_id=NULL removed; kept valid row employee_id={valid['employee_id']}"),
            })
            rows_to_delete.append(orphan["row"])
            repaired += 1
            continue

        # some employees are legitimately assigned to many stations, so only
        # a 1:1 valid/orphan pairing is safe to treat as a true duplicate
        if not valid_rows:
            reason = f"name+designation repeats {len(group)}x but no valid suborg_id exists among them"
        elif len(valid_rows) > 1:
            reason = (f"name+designation repeats {len(group)}x with {len(valid_rows)} valid suborg_id "
                      f"rows (employee_id={', '.join(str(v['employee_id']) for v in valid_rows)}) -- "
                      f"likely a shared role assigned to multiple stations, not a duplicate")
        else:
            reason = (f"name+designation repeats {len(group)}x with 1 valid suborg_id row "
                      f"(employee_id={valid_rows[0]['employee_id']}) but {len(null_rows)} orphan rows -- "
                      f"can't tell which, if any, orphan corresponds to that valid row")

        for orphan in null_rows:
            manual_review += 1
            change_log.append({
                "Sheet": "employees",
                "Action": "Manual review required",
                "Details": (f"employee_id={orphan['employee_id']} name='{orphan['name']}' "
                            f"designation='{orphan['designation']}' suborg_id=NULL -- {reason}"),
            })

    for row_idx in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_idx, 1)

    return repaired, skipped, manual_review


def parse_word_control_room_rows(path):
    """Returns dict: station name -> list of control-room row dicts (label,
    phones, emails), for every row any roster table in Word labels as a
    control room. Phone/email values are classified by content
    (_classify_row_content), not column position -- a merged-cell row can
    have a different cell count than its own table's header, which once
    caused the literal text "Control Room" to be read as a phone number."""
    doc = docx.Document(path)
    stations = defaultdict(list)

    def process_table(table, fallback_name):
        rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
        name_col, desig_col = find_header_columns(rows_cells)
        if name_col is None:
            return  # not a roster table

        current_name = None
        for cells in rows_cells:
            if not cells:
                continue
            first_cell = cells[0]

            if NUMBERED_TITLE_RE.match(first_cell):
                current_name = strip_leading_number(first_cell)
                continue

            if is_header_row(cells):
                continue

            if not any(cells):
                continue

            name_cell = cells[name_col] if name_col < len(cells) else ""
            desig_cell = cells[desig_col] if desig_col is not None and desig_col < len(cells) else ""
            if not name_cell:
                continue
            if "control room" not in name_cell.lower() and "control room" not in desig_cell.lower():
                continue

            station = current_name or fallback_name
            if not station:
                continue

            # e.g. a "Power House" / "Control Room" row -- label is whichever
            # cell actually says "Control Room"
            label = name_cell if "control room" in name_cell.lower() else desig_cell
            label_lower = label.lower()

            phones, emails, _others = _classify_row_content(
                [c for c in cells if clean(c).lower() != label_lower]
            )

            stations[station].append({
                "label": label,
                "phones": phones,
                "emails": emails,
            })

    current_heading = None
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue
        process_table(item, current_heading)

    return stations


def add_missing_control_rooms(wb, change_log):
    """Adds a control_rooms row for every (station, control-room label) pair
    Word has that the sheet has no row for at all -- unlike
    repair_control_rooms above, which only fixes suborg_id on existing rows.
    Matched first against sub_organizations, then (if that fails) against
    organizations, since some stations have no sub-organization breakdown
    and are legitimately parent-level. Exact name match only. A
    (suborg_id/org_id, label) pair already present is never re-added, so a
    second, distinct control room at the same station is still added."""
    ws = wb["control_rooms"]
    so_ws = wb["sub_organizations"]
    orgs_ws = wb["organizations"]

    suborg_by_name_lower = {}
    for sid, parent_oid, name, _addr, state in so_ws.iter_rows(min_row=2, values_only=True):
        if sid is not None and name:
            suborg_by_name_lower[strip_leading_number(clean(name)).lower()] = (sid, parent_oid, clean(state))

    org_name_by_id = {}
    org_by_name_lower = {}
    for oid, name, _addr, state in orgs_ws.iter_rows(min_row=2, values_only=True):
        if oid is not None and name:
            org_name_by_id[oid] = clean(name)
            org_by_name_lower[strip_leading_number(clean(name)).lower()] = (oid, clean(state))

    existing_labels = set()      # (suborg_id, label.lower())
    existing_org_labels = set()  # (org_id, label.lower()) for suborg_id-less rows
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        rid, org_id, suborg_id = row[0], row[1], row[2]
        label = row[5]
        if rid and rid > max_id:
            max_id = rid
        if suborg_id is not None:
            existing_labels.add((suborg_id, clean(label).lower()))
        elif org_id is not None:
            existing_org_labels.add((org_id, clean(label).lower()))

    word_stations = parse_word_control_room_rows(WORD_PATH)

    added = 0
    no_match = 0
    already_present = 0

    for station, cr_rows in sorted(word_stations.items()):
        key = strip_leading_number(clean(station)).lower()
        match = suborg_by_name_lower.get(key)
        org_only_match = None if match is not None else org_by_name_lower.get(key)

        if match is None and org_only_match is None:
            no_match += len(cr_rows)
            for cr in cr_rows:
                change_log.append({
                    "Sheet": "control_rooms",
                    "Action": "Manual review required",
                    "Details": (f"station='{station}' label='{cr['label']}' -- Word lists a control "
                                f"room here, but this station name has no exact match in "
                                f"sub_organizations or organizations -- left unchanged"),
                })
            continue

        if match is not None:
            sid, parent_oid, state = match
            parent_name = org_name_by_id.get(parent_oid, "")
        else:
            sid = None
            parent_oid, state = org_only_match
            parent_name = org_name_by_id.get(parent_oid, "")

        for cr in cr_rows:
            label_key = (sid, clean(cr["label"]).lower()) if sid is not None else None
            org_label_key = (parent_oid, clean(cr["label"]).lower())
            if (label_key is not None and label_key in existing_labels) or (
                sid is None and org_label_key in existing_org_labels
            ):
                already_present += 1
                continue

            max_id += 1
            phones = (cr["phones"] + [None] * 4)[:4]
            emails = (cr["emails"] + [None] * 2)[:2]
            sub_org_field = station if sid is not None else None

            ws.append([
                max_id, parent_oid, sid, parent_name, sub_org_field,
                cr["label"], None, None, state,
                phones[0], phones[1], phones[2], phones[3],
                None, None, None, emails[0], emails[1],
            ])
            if sid is not None:
                existing_labels.add(label_key)
            else:
                existing_org_labels.add(org_label_key)
            added += 1
            change_log.append({
                "Sheet": "control_rooms",
                "Action": "Added control room row (Word-verified, was missing entirely)",
                "Details": (f"station='{station}' org_id={parent_oid} suborg_id={sid} "
                            f"label='{cr['label']}' phones={cr['phones']} emails={cr['emails']} -- Word "
                            f"lists this control room but no row for this station existed in the "
                            f"control_rooms sheet at all"),
            })

    return added, no_match, already_present


def build_word_directory_row_index(path):
    """Indexes every data row of every roster-shaped table in Word as
    (station, cells), so a row can be relocated to its correct station by
    searching for content it still contains (e.g. a phone number) -- needed
    when the row's own org/suborg columns are what's unreliable."""
    doc = docx.Document(path)
    index = []

    def process_table(table, fallback_name):
        rows_cells = [[clean(c.text) for c in r.cells] for r in table.rows]
        name_col, _desig_col = find_header_columns(rows_cells)
        if name_col is None:
            return

        current_name = None
        for cells in rows_cells:
            if not cells:
                continue
            first_cell = cells[0]
            if NUMBERED_TITLE_RE.match(first_cell):
                current_name = strip_leading_number(first_cell)
                continue
            if is_header_row(cells):
                continue
            if not any(cells):
                continue
            station = current_name or fallback_name
            if station:
                index.append((station, cells))

    current_heading = None
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if item.style.name.startswith("Heading") and clean(item.text):
                current_heading = strip_leading_number(clean(item.text))
            continue
        process_table(item, current_heading)

    return index


def repair_shifted_directory_rows(wb, change_log, sheet_name):
    """Repairs control_rooms/switchyards rows whose columns were shifted by
    an inconsistent merged-cell layout in the source Word table (e.g.
    id=10 'ACB (India) Limited', where the phone ended up in cr_label and
    the label ended up in the state column). Detected purely by content
    shape -- never by station name or row id -- then repaired by using the
    row's own still-present value as a content anchor to find that row in
    Word and re-derive station, label, phone(s) and email(s) from there.
    An ambiguous or unmatched anchor is left unchanged and logged for
    manual review instead of guessed at."""
    ws = wb[sheet_name]
    so_ws = wb["sub_organizations"]
    orgs_ws = wb["organizations"]

    suborg_by_name_lower = {}
    for sid, parent_oid, name, _addr, state in so_ws.iter_rows(min_row=2, values_only=True):
        if sid is not None and name:
            suborg_by_name_lower[strip_leading_number(clean(name)).lower()] = (sid, parent_oid, clean(state))

    org_name_by_id = {}
    org_by_name_lower = {}
    for oid, name, _addr, state in orgs_ws.iter_rows(min_row=2, values_only=True):
        if oid is not None and name:
            org_name_by_id[oid] = clean(name)
            org_by_name_lower[strip_leading_number(clean(name)).lower()] = (oid, clean(state))

    word_index = None  # built lazily -- most rows in either sheet are fine

    repaired = 0
    manual_review = 0

    for row_idx in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row_idx, c).value for c in range(1, 19)]
        rid, cr_label = row_vals[0], row_vals[5]
        phones, fax, mobiles, emails = row_vals[9:13], row_vals[13], row_vals[14:16], row_vals[16:18]

        label_text = clean(cr_label)
        has_phone_cols = any(phones) or fax or any(mobiles)
        has_email_cols = any(emails)

        # two shapes of the same problem: the label is phone-shaped with no
        # real phone/email data, or a phone/email column holds the wrong type
        anchor = None
        if label_text and PHONE_RE.search(label_text) and not EMAIL_RE.search(label_text) \
                and not has_phone_cols and not has_email_cols:
            anchor = label_text
        else:
            for v in [*phones, fax, *mobiles]:
                if v and EMAIL_RE.search(clean(v)):
                    anchor = clean(v)
                    break
            if anchor is None:
                for v in emails:
                    if v and PHONE_RE.search(clean(v)) and not EMAIL_RE.search(clean(v)):
                        anchor = clean(v)
                        break

        if anchor is None:
            continue

        if word_index is None:
            word_index = build_word_directory_row_index(WORD_PATH)

        matches = [
            (station, cells) for station, cells in word_index
            if any(anchor == clean(c) or (len(anchor) >= 6 and anchor in clean(c)) for c in cells)
        ]
        # collapse to distinct stations before judging ambiguity -- a
        # repeated merged separator row can produce multiple raw matches
        distinct_stations = _dedup([m[0] for m in matches])

        if len(distinct_stations) != 1:
            manual_review += 1
            change_log.append({
                "Sheet": sheet_name,
                "Action": "Manual review required",
                "Details": (f"id={rid} cr_label='{cr_label}' looks shifted (a phone/email value was "
                            f"found in the wrong column -- anchor='{anchor}') -- its own content was "
                            f"found in Word at {len(distinct_stations)} distinct station(s) "
                            f"({', '.join(distinct_stations) or 'none'}); can't uniquely relocate it"),
            })
            continue

        station = distinct_stations[0]
        match_cells = next(cells for st, cells in matches if st == station)
        phones_found, emails_found, others_found = _classify_row_content(match_cells)
        label = max(others_found, key=len) if others_found else label_text

        key = strip_leading_number(clean(station)).lower()
        suborg_match = suborg_by_name_lower.get(key)
        org_match = None if suborg_match is not None else org_by_name_lower.get(key)

        if suborg_match is None and org_match is None:
            manual_review += 1
            change_log.append({
                "Sheet": sheet_name,
                "Action": "Manual review required",
                "Details": (f"id={rid} cr_label='{cr_label}' -- relocated via Word to station "
                            f"'{station}', but that name has no exact match in sub_organizations or "
                            f"organizations -- left unchanged"),
            })
            continue

        if suborg_match is not None:
            sid, parent_oid, state = suborg_match
            sub_org_field = station
        else:
            sid = None
            parent_oid, state = org_match
            sub_org_field = None
        parent_name = org_name_by_id.get(parent_oid, "")

        old_label, old_addr, old_state = cr_label, row_vals[7], row_vals[8]
        phones_padded = (phones_found + [None] * 4)[:4]
        emails_padded = (emails_found + [None] * 2)[:2]

        ws.cell(row_idx, 2).value = parent_oid
        ws.cell(row_idx, 3).value = sid
        ws.cell(row_idx, 4).value = parent_name
        ws.cell(row_idx, 5).value = sub_org_field
        ws.cell(row_idx, 6).value = label
        ws.cell(row_idx, 7).value = None
        ws.cell(row_idx, 8).value = None
        ws.cell(row_idx, 9).value = state
        ws.cell(row_idx, 10).value = phones_padded[0]
        ws.cell(row_idx, 11).value = phones_padded[1]
        ws.cell(row_idx, 12).value = phones_padded[2]
        ws.cell(row_idx, 13).value = phones_padded[3]
        ws.cell(row_idx, 14).value = None
        ws.cell(row_idx, 15).value = None
        ws.cell(row_idx, 16).value = None
        ws.cell(row_idx, 17).value = emails_padded[0]
        ws.cell(row_idx, 18).value = emails_padded[1]

        repaired += 1
        change_log.append({
            "Sheet": sheet_name,
            "Action": "Repaired shifted row (Word-verified relocation)",
            "Details": (f"id={rid} old_label='{old_label}' old_address='{old_addr}' "
                        f"old_state='{old_state}' -- relocated via its own content to station "
                        f"'{station}', org_id={parent_oid} suborg_id={sid} -- rewrote to "
                        f"label='{label}' phones={phones_found} emails={emails_found}"),
        })

    return repaired, manual_review


def repair_control_rooms(wb, change_log):
    ws = wb["control_rooms"]

    so_ws = wb["sub_organizations"]
    suborg_id_by_name_lower = {}
    for sid, _parent_oid, name, _addr, _state in so_ws.iter_rows(min_row=2, values_only=True):
        if sid is not None and name:
            suborg_id_by_name_lower[clean(name).lower()] = sid

    repaired = 0
    skipped = 0
    manual_review = 0

    for row_idx in range(2, ws.max_row + 1):
        suborg_cell = ws.cell(row_idx, CR_COL_SUBORG)
        if suborg_cell.value is not None:
            continue  # already linked, nothing to do

        cr_id = ws.cell(row_idx, CR_COL_ID).value
        sub_org_name = clean(ws.cell(row_idx, CR_COL_SUBORG_NAME).value)

        if not sub_org_name:
            skipped += 1
            continue

        matched_sid = suborg_id_by_name_lower.get(sub_org_name.lower())
        if matched_sid is None:
            manual_review += 1
            change_log.append({
                "Sheet": "control_rooms",
                "Action": "Manual review required",
                "Details": (f"id={cr_id} sub_organization_name='{sub_org_name}' has no exact match in "
                            f"sub_organizations -- left unchanged"),
            })
            continue

        suborg_cell.value = matched_sid
        repaired += 1
        change_log.append({
            "Sheet": "control_rooms",
            "Action": "Populated suborg_id",
            "Details": f"id={cr_id} sub_organization_name='{sub_org_name}' -> suborg_id={matched_sid}",
        })

    return repaired, skipped, manual_review


def run():
    shutil.copy(SOURCE_PATH, OUTPUT_PATH)  # original is never opened in write mode

    wb = openpyxl.load_workbook(OUTPUT_PATH)
    change_log = []

    # human-approved decisions take precedence over the automatic repairs below
    manual_applied, manual_not_found = apply_manual_decisions(wb, change_log)

    # suborg_id correction runs before duplicate-removal so a fixed row is
    # evaluated as a normal linked row rather than as an orphan
    suborg_repaired, suborg_review = repair_employee_suborg_mappings(wb, change_log)
    emp_repaired, emp_skipped, emp_review = repair_employees(wb, change_log)
    cr_repaired, cr_skipped, cr_review = repair_control_rooms(wb, change_log)

    # runs after repair_control_rooms so its suborg_id fills are already
    # reflected when checking for duplicates here
    cr_added, cr_added_no_match, cr_added_duplicate = add_missing_control_rooms(wb, change_log)

    # runs last: relocating a shifted row needs its own still-shifted content
    cr_shift_repaired, cr_shift_review = repair_shifted_directory_rows(wb, change_log, "control_rooms")
    sw_shift_repaired, sw_shift_review = repair_shifted_directory_rows(wb, change_log, "switchyards")

    wb.save(OUTPUT_PATH)

    os.makedirs(os.path.dirname(CHANGE_LOG_PATH), exist_ok=True)
    with open(CHANGE_LOG_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Sheet", "Action", "Details"])
        writer.writeheader()
        writer.writerows(change_log)

    employees_repaired = manual_applied + suborg_repaired + emp_repaired
    control_rooms_repaired = cr_repaired
    rows_skipped = emp_skipped + cr_skipped
    rows_manual_review = (
        suborg_review + emp_review + cr_review + cr_added_no_match
        + cr_shift_review + sw_shift_review
    )

    print(f"Repaired workbook written to: {OUTPUT_PATH}")
    print(f"Change log written to: {CHANGE_LOG_PATH}")
    print()
    print(f"Employees repaired: {employees_repaired}")
    print(f"  - approved manual decisions applied: {manual_applied}")
    print(f"  - suborg_id corrected against Word directory: {suborg_repaired}")
    print(f"  - orphan duplicates removed: {emp_repaired}")
    print(f"Control rooms repaired: {control_rooms_repaired}")
    print(f"Control rooms added (were missing entirely, Word-verified): {cr_added}")
    print(f"  - already present, skipped: {cr_added_duplicate}")
    print(f"  - no exact sub_organizations match, needs manual review: {cr_added_no_match}")
    print(f"Control rooms with shifted columns repaired: {cr_shift_repaired}")
    print(f"Switchyards with shifted columns repaired: {sw_shift_repaired}")
    print(f"Rows skipped: {rows_skipped}")
    print(f"Rows requiring manual review: {rows_manual_review}")
    if manual_not_found:
        print(f"Manual decisions skipped (workbook row no longer matches): {manual_not_found}")


if __name__ == "__main__":
    run()
