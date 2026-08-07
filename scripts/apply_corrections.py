"""Applies verified corrections derived from the Word<->Excel validation
analysis (the same analysis behind Validation_Report.xlsx) to a copy of
uploads/WR_DB_Ready_Final_Verified_TD_UPDATED.xlsx. Read-only on both
original inputs. Writes WR_DB_Ready_Final_Verified_v2.xlsx and
Correction_Log.xlsx.

Only evidence-backed, unambiguous corrections are applied automatically.
In particular: a value that's present in Word but blank in Excel is a safe
fill-in, but a value that's present on BOTH sides yet differs is only
auto-applied when the target field is single-purpose (designation, a lone
email). Multi-slot phone fields (Office/Residence/Mobile/Fax/... or
Contact/OPD/Emergency) are never overwritten on a mismatch, because Word's
own phone extraction doesn't preserve which slot a number belongs to --
guessing which column to overwrite would be exactly the kind of silent
data-modification this script is designed to avoid. Those cases are
routed to manual verification instead, same as any ambiguous match.

Usage:
  python scripts/apply_corrections.py
"""

import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from scripts.generate_validation_report import (
    WORD_PATH, EXCEL_PATH,
    load_excel, parse_word, group_word_segments,
    word_phone_set, word_email_set, word_employees_by_org,
    normalize_phone_set, normalize_email_set,
    merge_excel_records_by_label,
    compare_organizations, compare_hospitals, compare_emergency_services,
    compare_reference_sections,
    make_row,
)

V2_XLSX = "WR_DB_Ready_Final_Verified_v2.xlsx"
LOG_XLSX = "Correction_Log.xlsx"

CORRECTION_HEADER = [
    "Sheet Name", "Record ID", "Field Name", "Previous Value", "New Value",
    "Reason", "Source (Word Section/Table)", "Confidence Level",
]

COLS = {
    "employees": {"employee_id": 1, "org_id": 2, "suborg_id": 3, "department": 4, "name": 5, "designation": 6},
    "KMP": {"kmp_id": 1, "org_id": 2, "name": 3, "designation": 4, "office_phone": 5, "residence_phone": 6, "mobile_phone": 7, "email": 8},
    "Utility_Heads": {"utility_head_id": 1, "org_id": 2, "name": 3, "designation": 4, "office_phone": 5, "residence_phone": 6, "mobile_phone": 7, "email": 8},
    "control_rooms": {"id": 1, "org_id": 2, "suborg_id": 3, "organization_name": 4, "sub_organization_name": 5, "cr_label": 6, "designation": 7, "phone_1": 8, "phone_2": 9, "mobile_1": 10, "mobile_2": 11, "email_1": 12},
    "switchyards": {"id": 1, "org_id": 2, "suborg_id": 3, "organization_name": 4, "sub_organization_name": 5, "cr_label": 6, "phone_1": 7, "mobile_1": 8, "email_1": 9},
    "Hospitals": {"hospital_id": 1, "organization_id": 2, "hospital_name": 3, "contact_number": 4, "opd_number": 5, "emergency_number": 6, "address": 7, "website": 8},
    "Emergency_Services": {"service_id": 1, "service_type": 2, "name": 3, "contact_number": 4, "website": 5},
    "phone_numbers": {"phone_id": 1, "reference_type": 2, "reference_id": 3, "phone_type": 4, "phone_number": 5},
    "email_addresses": {"email_id": 1, "reference_type": 2, "reference_id": 3, "email": 4},
}


def new_correction(sheet, record_id, field, previous, new, reason, source, confidence, action="update", **extra):
    c = {
        "sheet": sheet, "record_id": record_id, "field": field,
        "previous_value": previous, "new_value": new, "reason": reason,
        "source_word": source, "confidence": confidence, "action": action,
    }
    c.update(extra)
    return c


def phone_correction(sheet, record_id, current_set, word_set, station, field_label):
    if word_set and not current_set:
        return new_correction(
            sheet, record_id, field_label, "(none)", ", ".join(sorted(word_set)),
            "Phone number present in Word but missing in Excel", station, "High",
        )
    if word_set and current_set and word_set != current_set:
        return new_correction(
            sheet, record_id, field_label, ", ".join(sorted(current_set)), ", ".join(sorted(word_set)),
            "Phone number(s) differ from Word, but which specific field (Office/Residence/Mobile/etc.) "
            "changed can't be determined with confidence -- left for manual verification rather than "
            "guessing which column to overwrite",
            station, "Low", action="manual",
        )
    return None


def email_correction(sheet, record_id, current_set, word_set, station, field_label="email"):
    if word_set and not current_set:
        return new_correction(
            sheet, record_id, field_label, "(none)", ", ".join(sorted(word_set)),
            "Email present in Word but missing in Excel", station, "High",
        )
    if word_set and current_set and word_set != current_set:
        if len(current_set) > 1:
            return new_correction(
                sheet, record_id, field_label, ", ".join(sorted(current_set)), ", ".join(sorted(word_set)),
                "Email differs from Word and Excel already has multiple emails on file -- left for manual "
                "verification rather than guessing which to replace",
                station, "Low", action="manual",
            )
        return new_correction(
            sheet, record_id, field_label, ", ".join(sorted(current_set)), ", ".join(sorted(word_set)),
            "Email differs from Word", station, "High",
        )
    return None


def compute_next_ids(excel_path):
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    def max_id(sheet):
        ids = [row[0] for row in wb[sheet].iter_rows(min_row=2, values_only=True) if row[0] is not None]
        return max(ids) if ids else 0

    return {
        "employees": max_id("employees"),
        "control_rooms": max_id("control_rooms"),
        "switchyards": max_id("switchyards"),
        "phone_numbers": max_id("phone_numbers"),
        "email_addresses": max_id("email_addresses"),
    }


def load_mobile_numbers(excel_path):
    """Employees' 'phones_raw' (from load_excel) mixes every phone type
    together; requirement 7's duplicate check specifically wants Mobile,
    so this re-reads phone_numbers filtered to phone_type == 'Mobile'."""
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    mobile_by_employee = defaultdict(set)
    for _pid, ref_type, ref_id, ptype, number in wb["phone_numbers"].iter_rows(min_row=2, values_only=True):
        if ref_type == "employee" and (ptype or "").strip().lower() == "mobile" and number:
            mobile_by_employee[ref_id] |= normalize_phone_set(str(number))
    return mobile_by_employee


# ─── EMPLOYEES: field corrections + missing-record inserts ──────────────────

def collect_employee_corrections(groups, excel):
    emp_by_suborg, emp_by_org_direct = defaultdict(list), defaultdict(list)
    for e in excel["employees"]:
        (emp_by_suborg[e["suborg_id"]] if e["suborg_id"] is not None else emp_by_org_direct[e["org_id"]]).append(e)

    corrections = []
    matched_excel_ids = set()

    for (match_type, match_id), g in groups.items():
        if match_type == "none":
            continue
        excel_list = emp_by_suborg.get(match_id, []) if match_type == "suborg" else emp_by_org_direct.get(match_id, [])
        excel_by_lower = {e["name"].lower(): e for e in excel_list}
        station = g["display_name"]
        org_id = match_id if match_type == "org" else excel["suborg_parent_by_id"].get(match_id)
        suborg_id = match_id if match_type == "suborg" else None

        for lower, w in g["employees"].items():
            e = excel_by_lower.get(lower)
            if e is None:
                corrections.append(new_correction(
                    "employees", None, "(new row)", None,
                    {"org_id": org_id, "suborg_id": suborg_id, "department": None,
                     "name": w["name"], "designation": w["designation"],
                     "phones": w["phones"], "emails": w["emails"]},
                    f"Employee listed in Word under '{station}' but missing from Excel",
                    station, "High", action="insert",
                ))
                continue

            matched_excel_ids.add(e["employee_id"])

            if w["designation"] and e["designation"] and w["designation"].lower() != e["designation"].lower():
                corrections.append(new_correction(
                    "employees", e["employee_id"], "designation", e["designation"], w["designation"],
                    "Designation differs from Word", station, "High",
                ))

            wp, ep = word_phone_set(w), normalize_phone_set(e.get("phones_raw", ""))
            c = phone_correction("employees/phone_numbers", e["employee_id"], ep, wp, station, "phone")
            if c:
                c["action"] = "insert_phone" if c["action"] == "update" else c["action"]
                corrections.append(c)

            we, ee = word_email_set(w), normalize_email_set(e.get("emails_raw", ""))
            c = email_correction("employees/email_addresses", e["employee_id"], ee, we, station, "email")
            if c:
                c["action"] = "insert_email" if c["action"] == "update" else c["action"]
                corrections.append(c)

    return corrections, matched_excel_ids


def collect_org_mapping_corrections(groups, excel, matched_excel_ids):
    word_station_meta = {}
    word_name_index = defaultdict(set)
    for (match_type, match_id), g in groups.items():
        word_station_meta[g["display_name"]] = (match_type, match_id)
        for lower in g["employees"]:
            word_name_index[lower].add(g["display_name"])

    corrections = []
    remapped_ids = set()
    for e in excel["employees"]:
        if e["employee_id"] in matched_excel_ids:
            continue
        stations = sorted(word_name_index.get(e["name"].lower(), []))
        if len(stations) != 1:
            continue
        station = stations[0]
        mtype, mid = word_station_meta.get(station, ("none", None))
        if mtype == "org" and mid == e["org_id"]:
            continue
        if mtype == "suborg" and mid == e["suborg_id"]:
            continue

        remapped_ids.add(e["employee_id"])
        if mtype == "org":
            corrections.append(new_correction(
                "employees", e["employee_id"], "org_id", e["org_id"], mid,
                f"Employee's Word station '{station}' maps to a different organization than Excel records",
                station, "High",
            ))
            if e["suborg_id"] is not None:
                corrections.append(new_correction(
                    "employees", e["employee_id"], "suborg_id", e["suborg_id"], None,
                    f"Clearing sub-organization since Word only identifies the parent organization '{station}'",
                    station, "Medium",
                ))
        elif mtype == "suborg":
            corrections.append(new_correction(
                "employees", e["employee_id"], "suborg_id", e["suborg_id"], mid,
                f"Employee's Word station '{station}' maps to a different sub-organization than Excel records",
                station, "High",
            ))
            new_org = excel["suborg_parent_by_id"].get(mid)
            if new_org is not None and new_org != e["org_id"]:
                corrections.append(new_correction(
                    "employees", e["employee_id"], "org_id", e["org_id"], new_org,
                    f"Updating parent organization to match the corrected sub-organization '{station}'",
                    station, "High",
                ))

    return corrections, remapped_ids, word_name_index


# ─── KMP / UTILITY HEADS ──────────────────────────────────────────────────────

def collect_kmp_corrections(word_kmp_candidates, excel_kmp):
    word_by_name = {}
    for w in word_kmp_candidates:
        word_by_name.setdefault(w["name"].lower(), w)

    corrections = []
    for e in excel_kmp:
        w = word_by_name.get(e["name"].lower())
        if w is None:
            continue
        station = w.get("org_banner") or "Synopsis of Important Telephone Numbers"

        if w["designation"] and e["designation"] and w["designation"].lower() != e["designation"].lower():
            corrections.append(new_correction(
                "KMP", e["kmp_id"], "designation", e["designation"], w["designation"],
                "Designation differs from Word", station, "High",
            ))

        ep = normalize_phone_set(" ".join([e["office_phone"], e["residence_phone"], e["mobile_phone"]]))
        c = phone_correction("KMP", e["kmp_id"], ep, word_phone_set(w), station, "office_phone")
        if c:
            corrections.append(c)

        ee = normalize_email_set(e["email"])
        c = email_correction("KMP", e["kmp_id"], ee, word_email_set(w), station, "email")
        if c:
            corrections.append(c)

    return corrections


def collect_utility_head_corrections(groups, excel):
    by_org = word_employees_by_org(groups, excel)
    corrections = []
    for uh in excel["utility_heads"]:
        candidates = by_org.get(uh["org_id"], [])
        w = next((c for c in candidates if c["name"].lower() == uh["name"].lower()), None)
        if w is None:
            continue
        station = excel["org_name_by_id"].get(uh["org_id"], "?")

        if w["designation"] and uh["designation"] and w["designation"].lower() != uh["designation"].lower():
            corrections.append(new_correction(
                "Utility_Heads", uh["utility_head_id"], "designation", uh["designation"], w["designation"],
                "Designation differs from Word", station, "High",
            ))

        ep = normalize_phone_set(" ".join([uh["office_phone"], uh["residence_phone"], uh["mobile_phone"]]))
        c = phone_correction("Utility_Heads", uh["utility_head_id"], ep, word_phone_set(w), station, "office_phone")
        if c:
            corrections.append(c)

        ee = normalize_email_set(uh["email"])
        c = email_correction("Utility_Heads", uh["utility_head_id"], ee, word_email_set(w), station, "email")
        if c:
            corrections.append(c)

    return corrections


# ─── CONTROL ROOMS / SWITCHYARDS ─────────────────────────────────────────────

def collect_cr_sw_corrections(groups, excel):
    cr_by_suborg, cr_by_org = defaultdict(list), defaultdict(list)
    for c in excel["control_rooms"]:
        (cr_by_suborg[c["suborg_id"]] if c["suborg_id"] is not None else cr_by_org[c["org_id"]]).append(c)
    sw_by_suborg, sw_by_org = defaultdict(list), defaultdict(list)
    for s in excel["switchyards"]:
        (sw_by_suborg[s["suborg_id"]] if s["suborg_id"] is not None else sw_by_org[s["org_id"]]).append(s)

    corrections = []
    for (match_type, match_id), g in groups.items():
        if match_type == "none":
            continue
        station = g["display_name"]
        org_id = match_id if match_type == "org" else excel["suborg_parent_by_id"].get(match_id)
        suborg_id = match_id if match_type == "suborg" else None

        excel_cr = merge_excel_records_by_label(
            cr_by_suborg.get(match_id, []) if match_type == "suborg" else cr_by_org.get(match_id, [])
        )
        excel_cr_by_label = {c["label"].strip().lower(): c for c in excel_cr}
        for label_lower, w in g["cr_by_label"].items():
            e = excel_cr_by_label.get(label_lower)
            wp = normalize_phone_set(" ".join(w["phones"]))
            we = normalize_email_set(" ".join(w["emails"]))
            if e is None:
                corrections.append(new_correction(
                    "control_rooms", None, "(new row)", None,
                    {"org_id": org_id, "suborg_id": suborg_id, "organization_name": station,
                     "sub_organization_name": "", "cr_label": w["label"], "designation": "",
                     # raw phone/email text preserved (not the exploded digit-only
                     # normalized set) so formatting like "07323-284542/ 284719" survives
                     "phone_1": ", ".join(w["phones"]), "email_1": ", ".join(w["emails"])},
                    f"Control room listed in Word under '{station}' but missing from Excel",
                    station, "High", action="insert",
                ))
                continue
            c = phone_correction("control_rooms", e["id"], normalize_phone_set(e["phones_raw"]), wp, station, "phone_1")
            if c:
                corrections.append(c)
            c = email_correction("control_rooms", e["id"], normalize_email_set(e["emails_raw"]), we, station, "email_1")
            if c:
                corrections.append(c)

        excel_sw = merge_excel_records_by_label(
            sw_by_suborg.get(match_id, []) if match_type == "suborg" else sw_by_org.get(match_id, [])
        )
        excel_sw_by_label = {s["label"].strip().lower(): s for s in excel_sw}
        for label_lower, w in g["sw_by_label"].items():
            e = excel_sw_by_label.get(label_lower)
            wp = normalize_phone_set(" ".join(w["phones"]))
            we = normalize_email_set(" ".join(w["emails"]))
            if e is None:
                corrections.append(new_correction(
                    "switchyards", None, "(new row)", None,
                    {"org_id": org_id, "suborg_id": suborg_id, "organization_name": station,
                     "sub_organization_name": "", "cr_label": w["label"],
                     "phone_1": ", ".join(w["phones"]), "email_1": ", ".join(w["emails"])},
                    f"Switchyard listed in Word under '{station}' but missing from Excel",
                    station, "High", action="insert",
                ))
                continue
            c = phone_correction("switchyards", e["id"], normalize_phone_set(e["phones_raw"]), wp, station, "phone_1")
            if c:
                corrections.append(c)
            c = email_correction("switchyards", e["id"], normalize_email_set(e["emails_raw"]), we, station, "email_1")
            if c:
                corrections.append(c)

    return corrections


# ─── HOSPITALS / EMERGENCY SERVICES ──────────────────────────────────────────

def collect_hospital_corrections(word_hospitals, excel_hospitals):
    word_by_name = {}
    for w in word_hospitals:
        word_by_name.setdefault(w["name"].lower(), w)

    corrections = []
    for e in excel_hospitals:
        w = word_by_name.get(e["name"].lower())
        if w is None:
            continue
        wp = normalize_phone_set(" ".join([w["contact_number"], w["opd_number"], w["emergency_number"]]))
        ep = normalize_phone_set(" ".join([e["contact_number"], e["opd_number"], e["emergency_number"]]))
        if wp and not ep:
            corrections.append(new_correction(
                "Hospitals", e["hospital_id"], "contact_number", "(none)", ", ".join(sorted(wp)),
                "Contact/OPD/Emergency numbers present in Word but missing in Excel",
                "Empaneled Hospitals", "High",
            ))
        elif wp and ep and wp != ep:
            corrections.append(new_correction(
                "Hospitals", e["hospital_id"], "contact_number/opd_number/emergency_number",
                ", ".join(sorted(ep)), ", ".join(sorted(wp)),
                "Numbers differ from Word, but which of contact/OPD/emergency changed can't be determined "
                "with confidence -- left for manual verification",
                "Empaneled Hospitals", "Low", action="manual",
            ))

    return corrections


def collect_emergency_corrections(word_emergency_services, excel_emergency_services):
    word_by_key = {}
    for w in word_emergency_services:
        word_by_key.setdefault(((w.get("service_type") or "").lower(), w["name"].lower()), w)

    corrections = []
    for e in excel_emergency_services:
        w = word_by_key.get(((e["service_type"] or "").lower(), e["name"].lower()))
        if w is None:
            continue
        c = phone_correction(
            "Emergency_Services", e["service_id"],
            normalize_phone_set(e["contact_number"]), normalize_phone_set(w["contact_number"]),
            e["service_type"], "contact_number",
        )
        if c:
            corrections.append(c)

    return corrections


# ─── DUPLICATES (exact-match deletion only) ──────────────────────────────────

def collect_employee_duplicate_deletions(excel, mobile_by_employee):
    seen = defaultdict(list)
    for e in excel["employees"]:
        key = (
            e["name"].strip().lower(), e["org_id"], e["suborg_id"],
            (e["designation"] or "").strip().lower(),
            frozenset(mobile_by_employee.get(e["employee_id"], set())),
            normalize_email_set(e.get("emails_raw", "")),
        )
        seen[key].append(e)

    corrections, deleted_ids = [], set()
    for items in seen.values():
        if len(items) > 1:
            items_sorted = sorted(items, key=lambda x: x["employee_id"])
            keep = items_sorted[0]
            for dup in items_sorted[1:]:
                deleted_ids.add(dup["employee_id"])
                corrections.append(new_correction(
                    "employees", dup["employee_id"], "(entire row)", "duplicate row", "deleted",
                    f"Exact duplicate of employee_id={keep['employee_id']} -- identical name, organization, "
                    "sub-organization, designation, mobile, and email",
                    "-", "High", action="delete",
                ))
    return corrections, deleted_ids


def collect_cr_sw_duplicate_deletions(excel):
    corrections = []
    for entity_name, sheet_name, items_list in (
        ("Control Room", "control_rooms", excel["control_rooms"]),
        ("Switchyard", "switchyards", excel["switchyards"]),
    ):
        seen = defaultdict(list)
        for c in items_list:
            key = (
                c["label"].strip().lower(), c["org_id"], c["suborg_id"],
                normalize_phone_set(c["phones_raw"]), normalize_email_set(c["emails_raw"]),
            )
            seen[key].append(c)
        for items in seen.values():
            if len(items) > 1:
                items_sorted = sorted(items, key=lambda x: x["id"])
                keep = items_sorted[0]
                for dup in items_sorted[1:]:
                    corrections.append(new_correction(
                        sheet_name, dup["id"], "(entire row)", "duplicate row", "deleted",
                        f"Exact duplicate of {entity_name.lower()} id={keep['id']} -- identical label, "
                        "organization, sub-organization, phone, and email",
                        "-", "High", action="delete",
                    ))
    return corrections


# ─── MANUAL VERIFICATION (everything not confidently auto-applied) ──────────

def collect_employee_manual_rows(excel, word_name_index, resolved_ids):
    rows = []
    for e in excel["employees"]:
        if e["employee_id"] in resolved_ids:
            continue
        stations = sorted(word_name_index.get(e["name"].lower(), []))
        if len(stations) == 0:
            rows.append(make_row(
                "Employee", "(not found)", e["name"], "(unknown)",
                "Employee in Excel not found anywhere in the Word directory",
                "Manually verify against the Word directory",
            ))
        elif len(stations) > 1:
            rows.append(make_row(
                "Employee", "; ".join(stations), e["name"], "(ambiguous)",
                f"Name appears under multiple Word stations ({', '.join(stations)}); "
                "cannot confidently determine the correct mapping",
                "Manually verify which station this employee belongs to",
            ))
    return rows


def collect_near_duplicate_manual_rows(excel, deleted_employee_ids):
    rows = []
    seen = defaultdict(list)
    for e in excel["employees"]:
        seen[(e["name"].strip().lower(), e["org_id"])].append(e)
    for items in seen.values():
        survivors = [i["employee_id"] for i in items if i["employee_id"] not in deleted_employee_ids]
        if len(survivors) > 1:
            rows.append(make_row(
                "Employee", "-", f"employee_ids {survivors} in org_id={items[0]['org_id']}", items[0]["name"],
                f"{len(survivors)} employee records share the name '{items[0]['name']}' in the same "
                "organization but differ in designation/mobile/email -- not identical enough to auto-merge",
                "Manually verify whether these are duplicate entries or distinct people sharing a name",
            ))
    return rows


# ─── APPLY TO WORKBOOK COPY ───────────────────────────────────────────────────

def find_row(ws, id_col, record_id):
    for row in ws.iter_rows(min_row=2):
        if row[id_col - 1].value == record_id:
            return row
    return None


def apply_corrections(wb, corrections, next_ids):
    applied = []
    for c in corrections:
        sheet = c["sheet"]
        action = c["action"]

        if action == "manual":
            continue  # not applied -- goes to the log as a manual-verification row only

        if sheet == "employees/phone_numbers":
            ws_emp, ws_phone = wb["employees"], wb["phone_numbers"]
            if action == "insert_phone":
                for num in c["new_value"].split(", "):
                    next_ids["phone_numbers"] += 1
                    ws_phone.append([next_ids["phone_numbers"], "employee", c["record_id"], "Contact", num])
                applied.append(c)
            continue

        if sheet == "employees/email_addresses":
            ws_email = wb["email_addresses"]
            if action == "insert_email":
                for em in c["new_value"].split(", "):
                    next_ids["email_addresses"] += 1
                    ws_email.append([next_ids["email_addresses"], "employee", c["record_id"], em])
                applied.append(c)
            continue

        if action == "insert":
            col_map = COLS[sheet]
            next_ids[sheet] += 1
            new_id = next_ids[sheet]
            row_data = c["new_value"]
            ws = wb[sheet]

            if sheet == "employees":
                new_row = [None] * len(col_map)
                new_row[col_map["employee_id"] - 1] = new_id
                new_row[col_map["org_id"] - 1] = row_data["org_id"]
                new_row[col_map["suborg_id"] - 1] = row_data["suborg_id"]
                new_row[col_map["department"] - 1] = row_data["department"]
                new_row[col_map["name"] - 1] = row_data["name"]
                new_row[col_map["designation"] - 1] = row_data["designation"]
                ws.append(new_row)
                ws_phone, ws_email = wb["phone_numbers"], wb["email_addresses"]
                for num in row_data["phones"]:
                    next_ids["phone_numbers"] += 1
                    ws_phone.append([next_ids["phone_numbers"], "employee", new_id, "Contact", num])
                for em in row_data["emails"]:
                    next_ids["email_addresses"] += 1
                    ws_email.append([next_ids["email_addresses"], "employee", new_id, em])
            elif sheet in ("control_rooms", "switchyards"):
                new_row = [None] * len(col_map)
                new_row[col_map["id"] - 1] = new_id
                for field in ("org_id", "suborg_id", "organization_name", "sub_organization_name", "cr_label",
                              "phone_1", "email_1"):
                    if field in col_map and field in row_data:
                        new_row[col_map[field] - 1] = row_data[field]
                if sheet == "control_rooms" and "designation" in row_data:
                    new_row[col_map["designation"] - 1] = row_data["designation"]
                ws.append(new_row)

            c["record_id"] = new_id
            applied.append(c)
            continue

        if action == "delete":
            ws = wb[sheet]
            id_col = COLS[sheet].get("id") or COLS[sheet].get("employee_id")
            row = find_row(ws, id_col, c["record_id"])
            if row is not None:
                ws.delete_rows(row[0].row, 1)
                applied.append(c)
            continue

        # plain field update
        id_col_name = list(COLS[sheet].keys())[0]
        id_col = COLS[sheet][id_col_name]
        field = c["field"]
        if field not in COLS[sheet]:
            continue
        row = find_row(ws := wb[sheet], id_col, c["record_id"])
        if row is None:
            continue
        row[COLS[sheet][field] - 1].value = c["new_value"]
        applied.append(c)

    return applied


# ─── REPORT WRITERS ───────────────────────────────────────────────────────────

def format_value(v):
    if isinstance(v, (list, set, frozenset)):
        return ", ".join(sorted(v))
    if isinstance(v, dict):
        return str(v)
    return v


def write_log_sheet(wb, name, entries, status_suffix):
    ws = wb.create_sheet(name)
    ws.append(CORRECTION_HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F6F5C")
    for e in entries:
        confidence = e.get("confidence", "-")
        ws.append([
            e.get("sheet", ""), e.get("record_id", ""), e.get("field", ""),
            format_value(e.get("previous_value")), format_value(e.get("new_value")),
            e.get("reason", ""), e.get("source_word", ""), f"{confidence} ({status_suffix})",
        ])
    widths = [22, 12, 20, 30, 30, 55, 34, 22]
    for col, w in zip("ABCDEFGH", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def write_manual_rows(wb, name, rows):
    ws = wb.create_sheet(name)
    header = ["Category", "Source (Word)", "Current Value (Excel)", "Expected Value (Word)",
              "Issue Description", "Suggested Correction"]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="B9702F")
    for r in rows:
        ws.append([r.get(h, "") for h in header])
    widths = [22, 40, 34, 30, 46, 40]
    for col, w in zip("ABCDEF", widths):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


def write_summary_sheet(wb, summary):
    ws = wb.create_sheet("Summary", 0)
    ws["A1"] = "Import Readiness Summary"
    ws["A1"].font = Font(bold=True, size=14)
    r = 3
    for label, value in summary.items():
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws.cell(row=r, column=2, value=value)
        r += 1
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 60


def run():
    word_mtime_before = os.path.getmtime(WORD_PATH)
    excel_mtime_before = os.path.getmtime(EXCEL_PATH)

    print("Loading Excel workbook and parsing Word document...")
    excel = load_excel(EXCEL_PATH)
    word = parse_word(WORD_PATH)
    groups = group_word_segments(word["all_station_segments"], excel)
    next_ids = compute_next_ids(EXCEL_PATH)
    mobile_by_employee = load_mobile_numbers(EXCEL_PATH)

    print("Collecting evidence-backed corrections...")
    emp_corrections, matched_excel_ids = collect_employee_corrections(groups, excel)
    org_map_corrections, remapped_ids, word_name_index = collect_org_mapping_corrections(groups, excel, matched_excel_ids)
    kmp_corrections = collect_kmp_corrections(word["word_kmp_candidates"], excel["kmp"])
    uh_corrections = collect_utility_head_corrections(groups, excel)
    crsw_corrections = collect_cr_sw_corrections(groups, excel)
    hosp_corrections = collect_hospital_corrections(word["word_hospitals"], excel["hospitals"])
    emerg_corrections = collect_emergency_corrections(word["word_emergency_services"], excel["emergency_services"])

    emp_dup_corrections, deleted_employee_ids = collect_employee_duplicate_deletions(excel, mobile_by_employee)
    crsw_dup_corrections = collect_cr_sw_duplicate_deletions(excel)

    all_corrections = (
        emp_corrections + org_map_corrections + kmp_corrections + uh_corrections
        + crsw_corrections + hosp_corrections + emerg_corrections
        + emp_dup_corrections + crsw_dup_corrections
    )
    auto_apply = [c for c in all_corrections if c["action"] != "manual"]
    deferred = [c for c in all_corrections if c["action"] == "manual"]

    print("Collecting manual-verification items (not auto-applied)...")
    resolved_ids = matched_excel_ids | remapped_ids
    r_orgs = compare_organizations(groups, excel)
    manual_rows = (
        r_orgs["missing"] + r_orgs["extra_needs_manual"]
        + collect_employee_manual_rows(excel, word_name_index, resolved_ids)
        + collect_near_duplicate_manual_rows(excel, deleted_employee_ids)
        + compare_hospitals(word["word_hospitals"], excel["hospitals"])["extra_needs_manual"]
        + compare_emergency_services(word["word_emergency_services"], excel["emergency_services"])["extra_needs_manual"]
        + compare_reference_sections(
            word["word_reference_rows"], word["unresolved_reference_sections"], excel["reference_sections"]
        )["needs_manual"]
    )

    print("Applying corrections to a workbook copy...")
    wb_v2 = openpyxl.load_workbook(EXCEL_PATH)
    applied = apply_corrections(wb_v2, auto_apply, next_ids)
    wb_v2.save(V2_XLSX)

    print("Writing correction log...")
    total_reviewed = (
        len(excel["organizations"]) + len(excel["sub_organizations"]) + len(excel["employees"])
        + len(excel["kmp"]) + len(excel["utility_heads"]) + len(excel["control_rooms"])
        + len(excel["switchyards"]) + len(excel["hospitals"]) + len(excel["emergency_services"])
    )
    total_corrections_applied = len(applied)
    total_manual = len(deferred) + len(manual_rows)
    remaining_critical = len(manual_rows)  # unresolved orgs/mappings/duplicates; excludes routine field-level manual items
    readiness_pct = round(100 * (total_reviewed - remaining_critical) / total_reviewed, 2) if total_reviewed else 0.0

    summary = {
        "Total Records Reviewed": total_reviewed,
        "Total Corrections Applied": total_corrections_applied,
        "Total Records Left for Manual Verification": total_manual,
        "Remaining Critical Issues": remaining_critical,
        "Overall Import Readiness %": readiness_pct,
        "Readiness Formula": "(Total Reviewed - Remaining Critical Issues) / Total Reviewed x 100",
        "Note on 'Source (Word Page/Table)'": (
            ".docx files do not store page numbers (pagination is a rendering-time concept, not part of the "
            "document XML) -- the Source column instead names the Word section/table heading the evidence "
            "came from."
        ),
        "Note on methodology": (
            "Corrections were derived by re-running the same parsing/matching analysis that produced "
            "Validation_Report.xlsx (not by re-parsing that report's prose text), so every correction traces "
            "back to a real record ID rather than a name-matched guess."
        ),
    }

    wb_log = openpyxl.Workbook()
    wb_log.remove(wb_log.active)
    write_summary_sheet(wb_log, summary)
    write_log_sheet(wb_log, "Applied_Corrections", applied, "Applied")
    write_log_sheet(wb_log, "Deferred_Field_Corrections", deferred, "Manual Verification Required")
    write_manual_rows(wb_log, "Needs_Manual_Verification", manual_rows)
    wb_log.save(LOG_XLSX)

    assert os.path.getmtime(WORD_PATH) == word_mtime_before, "Word source file was modified!"
    assert os.path.getmtime(EXCEL_PATH) == excel_mtime_before, "Excel source file was modified!"

    print("=" * 70)
    print("  Import Readiness Summary")
    print("=" * 70)
    for label, value in summary.items():
        print(f"  {label}: {value}")
    print("=" * 70)
    print(f"  Corrected workbook: {V2_XLSX}")
    print(f"  Correction log:     {LOG_XLSX}")


if __name__ == "__main__":
    run()
