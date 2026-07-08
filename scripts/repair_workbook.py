"""
Repairs uploads/WR_DB_Ready_Final.xlsx into a new, corrected copy:
uploads/WR_DB_Ready_Final_Verified.xlsx

The original workbook is never opened in write mode -- it is copied first,
and only the copy is edited. Every change is logged to
reports/repair_change_log.csv.

Repair 1 (employees): remove an orphan duplicate employee row only when the
match is unambiguous. A blank-suborg_id row is auto-deleted only when ALL of
the following hold for its (name, designation) group:

  1. same employee name
  2. same designation
  3. exactly ONE row in the group has a valid suborg_id
  4. exactly ONE row in the group has suborg_id blank (the orphan itself)
  5. i.e. there are NOT multiple valid rows for that name+designation

If a name+designation group has multiple valid (linked) rows, or multiple
orphan (blank-suborg_id) rows, nothing in that group is deleted -- some
employees are legitimately assigned to many stations (e.g. Alapan Bera,
Jitendra Ranawat each appear as the same role across dozens of stations),
and a 1:1 valid/orphan pairing is the only shape that's safe to assume is a
true accidental duplicate rather than a distinct station assignment that
just never got its suborg_id filled in. Every orphan row in an ambiguous
group is instead logged to the manual review report. No employee name is
ever changed.

Repair 2 (control_rooms): where suborg_id is blank but sub_organization_name
is an exact (case/whitespace-insensitive) match for a real sub_organizations
name, fill in that suborg_id. No fuzzy/substring matching is used here --
anything short of an exact match is left alone and logged for manual review.

Usage:
  python scripts/repair_workbook.py
"""

import csv
import os
import re
import shutil

import openpyxl

SOURCE_PATH = "uploads/WR_DB_Ready_Final.xlsx"
OUTPUT_PATH = "uploads/WR_DB_Ready_Final_Verified.xlsx"
CHANGE_LOG_PATH = "reports/repair_change_log.csv"

EMP_COL_ID, EMP_COL_ORG, EMP_COL_SUBORG, EMP_COL_NAME, EMP_COL_DESIG = 1, 2, 3, 4, 5
CR_COL_ID, CR_COL_ORG, CR_COL_SUBORG, CR_COL_ORG_NAME, CR_COL_SUBORG_NAME = 1, 2, 3, 4, 5


def clean(v):
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


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

        # Ambiguous: multiple valid rows, multiple orphan rows, or both -- some
        # employees are legitimately assigned to many stations, so a 1:1
        # valid/orphan pairing is required before deleting anything. Anything
        # else is left alone and logged for manual review instead of guessed at.
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

    emp_repaired, emp_skipped, emp_review = repair_employees(wb, change_log)
    cr_repaired, cr_skipped, cr_review = repair_control_rooms(wb, change_log)

    wb.save(OUTPUT_PATH)

    os.makedirs(os.path.dirname(CHANGE_LOG_PATH), exist_ok=True)
    with open(CHANGE_LOG_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Sheet", "Action", "Details"])
        writer.writeheader()
        writer.writerows(change_log)

    employees_repaired = emp_repaired
    control_rooms_repaired = cr_repaired
    rows_skipped = emp_skipped + cr_skipped
    rows_manual_review = emp_review + cr_review

    print(f"Repaired workbook written to: {OUTPUT_PATH}")
    print(f"Change log written to: {CHANGE_LOG_PATH}")
    print()
    print(f"Employees repaired: {employees_repaired}")
    print(f"Control rooms repaired: {control_rooms_repaired}")
    print(f"Rows skipped: {rows_skipped}")
    print(f"Rows requiring manual review: {rows_manual_review}")


if __name__ == "__main__":
    run()
