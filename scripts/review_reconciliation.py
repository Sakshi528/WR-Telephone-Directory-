"""Interactive, human-approved reconciliation review.

Reads reports/directory_reconciliation_report.csv and cross-references the
Word directory, the live database, and the verified workbook to generate
repair suggestions. Every suggestion is presented with an explicit [Y/n]
prompt; decisions are recorded in config/manual_reconciliation_decisions.csv,
the only thing this script writes. Already-decided (APPROVED/REJECTED)
combinations are never re-presented; SKIPPED ones may reappear later.

Usage:
  python scripts/review_reconciliation.py
"""

import csv
import os
import sys
from collections import defaultdict

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.compare_word_excel import clean, strip_leading_number
from scripts.repair_workbook import parse_word_employee_roles

from app import app
from models.organization import Organization
from models.employee import Employee
from models.directory_number import DirectoryNumber

RECONCILIATION_CSV = "reports/directory_reconciliation_report.csv"
WORKBOOK_PATH = "uploads/WR_DB_Ready_Final_Verified.xlsx"
WORD_PATH = "uploads/Western Region Phone Directory 2025 Main_Telephone.docx"
DECISIONS_CSV = "config/manual_reconciliation_decisions.csv"

DECISIONS_FIELDS = [
    "record_type", "record_name", "designation",
    "old_suborg_id", "new_suborg_id", "action",
    "reason", "approved_by", "date", "decision",
]

EMP_COL_ID, EMP_COL_ORG, EMP_COL_SUBORG, EMP_COL_NAME, EMP_COL_DESIG = 1, 2, 3, 4, 5
CR_COL_ID, CR_COL_ORG, CR_COL_SUBORG, CR_COL_ORG_NAME, CR_COL_SUBORG_NAME, CR_COL_LABEL = 1, 2, 3, 4, 5, 6


def _normalize_sid(value):
    # fresh suggestions carry an int/None; values read back from CSV are strings
    if value is None or value == "" or value == "None":
        return ""
    return str(int(value))


def decision_key(row):
    return (row["record_type"], row["record_name"].lower(), (row["designation"] or "").lower(),
            row["action"], _normalize_sid(row["old_suborg_id"]), _normalize_sid(row["new_suborg_id"]))


def load_existing_decisions():
    if not os.path.exists(DECISIONS_CSV):
        return {}
    with open(DECISIONS_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    return {decision_key(r): r for r in rows}


def append_decision(row):
    file_exists = os.path.exists(DECISIONS_CSV)
    os.makedirs(os.path.dirname(DECISIONS_CSV), exist_ok=True)
    with open(DECISIONS_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=DECISIONS_FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


# ─── LOAD RECONCILIATION REPORT ──────────────────────────────────────────────

def load_report():
    with open(RECONCILIATION_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


# ─── WORKBOOK LOOKUPS ────────────────────────────────────────────────────────

def load_workbook_lookups():
    wb = openpyxl.load_workbook(WORKBOOK_PATH, read_only=True)

    so_ws = wb["sub_organizations"]
    suborg_id_by_name_lower = {}
    for sid, _parent_oid, name, _addr, _state in so_ws.iter_rows(min_row=2, values_only=True):
        if sid is not None and name:
            suborg_id_by_name_lower[strip_leading_number(clean(name)).lower()] = sid

    emp_ws = wb["employees"]
    emp_rows_by_key = defaultdict(list)
    for row in emp_ws.iter_rows(min_row=2, values_only=True):
        emp_id, org_id, suborg_id, name, designation = row
        name = clean(name)
        if not name:
            continue
        key = (name.lower(), clean(designation).lower())
        emp_rows_by_key[key].append({"employee_id": emp_id, "suborg_id": suborg_id})

    cr_ws = wb["control_rooms"]
    cr_rows = list(cr_ws.iter_rows(min_row=2, values_only=True))

    return {
        "suborg_id_by_name_lower": suborg_id_by_name_lower,
        "emp_rows_by_key": emp_rows_by_key,
        "cr_rows": cr_rows,
    }


def resolve_station_to_suborg_id(station_name, wb_lookups):
    key = strip_leading_number(clean(station_name)).lower()
    return wb_lookups["suborg_id_by_name_lower"].get(key)


# ─── DATABASE LOOKUPS ────────────────────────────────────────────────────────

def load_db_lookups():
    with app.app_context():
        org_id_by_name_lower = {}
        for o in Organization.query.all():
            org_id_by_name_lower[o.organization_name.strip().lower()] = o.id

        employees_by_org = defaultdict(list)
        for e in Employee.query.all():
            if e.organization_id is not None:
                employees_by_org[e.organization_id].append(
                    {"name": e.employee_name, "designation": e.designation or ""}
                )

        cr_by_org = defaultdict(list)
        sw_by_org = defaultdict(list)
        for dn in DirectoryNumber.query.all():
            if dn.organization_id is None:
                continue
            cat = (dn.category or "").strip().lower()
            if cat == "control room":
                cr_by_org[dn.organization_id].append({"name": dn.name, "phone": dn.phone_number or ""})
            elif cat == "switchyard":
                sw_by_org[dn.organization_id].append({"name": dn.name, "phone": dn.phone_number or ""})

    return {
        "org_id_by_name_lower": org_id_by_name_lower,
        "employees_by_org": employees_by_org,
        "cr_by_org": cr_by_org,
        "sw_by_org": sw_by_org,
    }


# ─── WORD LOOKUPS ────────────────────────────────────────────────────────────

def load_word_lookups():
    role_lookup = parse_word_employee_roles(WORD_PATH)  # (name_l, desig_l) -> {stations}
    employees_by_station = defaultdict(set)
    for (name_l, desig_l), stations in role_lookup.items():
        for station in stations:
            employees_by_station[station].add((name_l, desig_l))
    return role_lookup, employees_by_station


# ─── SUGGESTION GENERATION: EMPLOYEES ────────────────────────────────────────

def build_employee_suggestions(report_rows, word_role_lookup, db_lookups, wb_lookups):
    missing_stations = defaultdict(list)  # name_lower -> [station, ...]
    extra_stations = defaultdict(list)
    display_name = {}

    for r in report_rows:
        if r["Category"] != "Employee":
            continue
        name_lower = r["Name"].strip().lower()
        display_name[name_lower] = r["Name"].strip()
        if r["Issue"] == "Missing":
            missing_stations[name_lower].append(r["Station"])
        elif r["Issue"] == "Extra":
            extra_stations[name_lower].append(r["Station"])

    suggestions = []
    all_names = set(missing_stations) | set(extra_stations)

    for name_lower in sorted(all_names):
        missing_list = missing_stations.get(name_lower, [])
        extra_list = extra_stations.get(name_lower, [])
        name = display_name[name_lower]

        if len(missing_list) == 1 and len(extra_list) == 1:
            from_station, to_station = extra_list[0], missing_list[0]
            designation = _find_designation(name, from_station, extra_list, db_lookups)
            sugg = _build_move_suggestion(
                name, designation, from_station, to_station, wb_lookups,
                confidence="HIGH",
                reason="Appears as Missing at exactly one station and Extra at exactly one "
                       "station -- exact Word match.",
            )
            if sugg:
                suggestions.append(sugg)
            continue

        if missing_list and extra_list:
            # Ambiguous: more than one candidate on at least one side.
            designation = _find_designation(name, extra_list[0], extra_list, db_lookups)
            suggestions.append({
                "record_type": "Employee", "record_name": name, "designation": designation,
                "action": "move_suborg", "old_suborg_id": None, "new_suborg_id": None,
                "confidence": "MEDIUM",
                "reason": (f"Missing at {len(missing_list)} station(s) ({', '.join(missing_list)}) "
                           f"and Extra at {len(extra_list)} station(s) ({', '.join(extra_list)}) -- "
                           f"can't determine which pairing is correct without guessing."),
                "candidates": {"missing": missing_list, "extra": extra_list},
            })
            continue

        if missing_list and not extra_list:
            suggestions.append({
                "record_type": "Employee", "record_name": name, "designation": "",
                "action": "ignore", "old_suborg_id": None, "new_suborg_id": None,
                "confidence": "LOW",
                "reason": (f"Missing at {len(missing_list)} station(s) but has no Extra record "
                           f"anywhere -- nothing to move from; this employee may simply not be in "
                           f"the source workbook at all. Not auto-fixable."),
            })
            continue

        if extra_list and not missing_list:
            suggestions.append({
                "record_type": "Employee", "record_name": name, "designation": "",
                "action": "ignore", "old_suborg_id": None, "new_suborg_id": None,
                "confidence": "LOW",
                "reason": (f"Extra at {len(extra_list)} station(s) ({', '.join(extra_list)}) but not "
                           f"Missing anywhere -- may be a legitimate assignment Word doesn't list, or "
                           f"a genuine data error. Needs manual investigation."),
            })

    # Merge-duplicate detection: same name appearing 2+ times in the DB at the same station.
    with app.app_context():
        for org_id, emps in db_lookups["employees_by_org"].items():
            by_name = defaultdict(list)
            for e in emps:
                by_name[e["name"].strip().lower()].append(e)
            org = Organization.query.get(org_id)
            for name_lower, dupes in by_name.items():
                if len(dupes) < 2:
                    continue
                if not org:
                    continue
                designations = {d["designation"].strip().lower() for d in dupes}
                if len(designations) != 1:
                    # different designations -- may be two people sharing a name
                    suggestions.append({
                        "record_type": "Employee", "record_name": dupes[0]["name"],
                        "designation": " / ".join(sorted({d["designation"] for d in dupes})),
                        "action": "merge_duplicate", "old_suborg_id": None, "new_suborg_id": None,
                        "confidence": "LOW",
                        "reason": (f"Appears {len(dupes)} times in the database at station "
                                   f"'{org.organization_name}' with different designations -- may be "
                                   f"different people sharing a name, not a true duplicate."),
                    })
                    continue

                shared_sid = resolve_station_to_suborg_id(org.organization_name, wb_lookups)
                confidence = "HIGH" if shared_sid is not None else "MEDIUM"
                suggestions.append({
                    "record_type": "Employee", "record_name": dupes[0]["name"],
                    "designation": dupes[0]["designation"],
                    "action": "merge_duplicate", "old_suborg_id": shared_sid, "new_suborg_id": shared_sid,
                    "confidence": confidence,
                    "reason": (f"Appears {len(dupes)} times in the database at station "
                               f"'{org.organization_name}' with identical designation -- looks like a "
                               f"true duplicate row. Approving removes all but one workbook row at "
                               f"suborg_id={shared_sid} for this name+designation."),
                })

    return suggestions


def _find_designation(name, station, candidate_stations, db_lookups):
    org_id = db_lookups["org_id_by_name_lower"].get(strip_leading_number(clean(station)).lower())
    if org_id is not None:
        for e in db_lookups["employees_by_org"].get(org_id, []):
            if e["name"].strip().lower() == name.lower():
                return e["designation"]
    return ""


def _build_move_suggestion(name, designation, from_station, to_station, wb_lookups, confidence, reason):
    new_sid = resolve_station_to_suborg_id(to_station, wb_lookups)
    old_sid = resolve_station_to_suborg_id(from_station, wb_lookups)

    if new_sid is None:
        return {
            "record_type": "Employee", "record_name": name, "designation": designation,
            "action": "move_suborg", "old_suborg_id": old_sid, "new_suborg_id": None,
            "confidence": "LOW",
            "reason": (f"Word shows this person belongs at '{to_station}', but that name has no "
                       f"exact match in the workbook's sub_organizations sheet -- cannot generate "
                       f"a concrete suborg_id to move to."),
        }

    key = (name.lower(), designation.lower())
    matching_rows = [r for r in wb_lookups["emp_rows_by_key"].get(key, []) if r["suborg_id"] == old_sid]
    if not matching_rows:
        return {
            "record_type": "Employee", "record_name": name, "designation": designation,
            "action": "move_suborg", "old_suborg_id": old_sid, "new_suborg_id": new_sid,
            "confidence": "LOW",
            "reason": (f"Expected a workbook row for '{name}' / '{designation}' with suborg_id="
                       f"{old_sid} (station '{from_station}') but none was found -- can't confirm "
                       f"the source row to move."),
        }

    return {
        "record_type": "Employee", "record_name": name, "designation": designation,
        "action": "move_suborg", "old_suborg_id": old_sid, "new_suborg_id": new_sid,
        "confidence": confidence,
        "reason": f"{reason} Move from '{from_station}' (suborg_id={old_sid}) to '{to_station}' (suborg_id={new_sid}).",
    }


# ─── SUGGESTION GENERATION: CONTROL ROOMS / SWITCHYARDS ─────────────────────

def build_directory_suggestions(report_rows, wb_lookups, category):
    missing = [r["Station"] for r in report_rows if r["Category"] == category and r["Issue"] == "Missing"]
    extra = [r["Station"] for r in report_rows if r["Category"] == category and r["Issue"] == "Extra"]

    suggestions = []
    used_extra = set()

    for station in missing:
        best_match = None
        station_key = strip_leading_number(clean(station)).lower()
        for candidate in extra:
            if candidate in used_extra:
                continue
            cand_key = strip_leading_number(clean(candidate)).lower()
            if station_key in cand_key or cand_key in station_key:
                best_match = candidate
                break

        record_name = f"{category} at {station}"
        new_sid = resolve_station_to_suborg_id(station, wb_lookups)

        if best_match is None:
            suggestions.append({
                "record_type": category, "record_name": record_name, "designation": "",
                "action": f"move_{'control_room' if category == 'Control Room' else 'switchyard'}",
                "old_suborg_id": None, "new_suborg_id": new_sid,
                "confidence": "LOW",
                "reason": (f"Word lists a {category.lower()} at '{station}' with no record in the "
                           f"database, and no plausible matching 'Extra' {category.lower()} was found "
                           f"by name overlap. Needs manual investigation -- no source record to move."),
            })
            continue

        used_extra.add(best_match)
        old_sid = resolve_station_to_suborg_id(best_match, wb_lookups)
        suggestions.append({
            "record_type": category, "record_name": record_name, "designation": "",
            "action": f"move_{'control_room' if category == 'Control Room' else 'switchyard'}",
            "old_suborg_id": old_sid, "new_suborg_id": new_sid,
            "confidence": "LOW",
            "reason": (f"Partial name overlap between missing station '{station}' and extra station "
                       f"'{best_match}' -- not an exact match, so this is only a possible candidate, "
                       f"not a confirmed one. Review carefully before approving."),
        })

    return suggestions


# ─── INTERACTIVE REVIEW ──────────────────────────────────────────────────────

def prompt_decision(suggestion, approved_by):
    print("-" * 60)
    print(f"{suggestion['record_type']}:")
    print(f"  {suggestion['record_name']}" + (f" ({suggestion['designation']})" if suggestion["designation"] else ""))
    print()
    print(f"Action: {suggestion['action']}")
    if suggestion.get("old_suborg_id") is not None or suggestion.get("new_suborg_id") is not None:
        print(f"  old_suborg_id: {suggestion['old_suborg_id']}")
        print(f"  new_suborg_id: {suggestion['new_suborg_id']}")
    print()
    print(f"Reason: {suggestion['reason']}")
    print()
    print(f"Confidence: {suggestion['confidence']}")
    print()

    while True:
        resp = input("Approve? [y/N/s=skip/q=quit] ").strip().lower()
        if resp in ("y", "yes"):
            return "APPROVED"
        if resp in ("n", "no", ""):
            return "REJECTED"
        if resp in ("s", "skip"):
            return "SKIPPED"
        if resp in ("q", "quit"):
            return "QUIT"
        print("Please answer y, n, s, or q.")


def run():
    import getpass
    from datetime import date

    approved_by = os.environ.get("RECONCILE_REVIEWER") or getpass.getuser()

    report_rows = load_report()
    wb_lookups = load_workbook_lookups()
    db_lookups = load_db_lookups()
    word_role_lookup, _word_employees_by_station = load_word_lookups()
    existing_decisions = load_existing_decisions()

    all_suggestions = []
    all_suggestions.extend(build_employee_suggestions(report_rows, word_role_lookup, db_lookups, wb_lookups))
    all_suggestions.extend(build_directory_suggestions(report_rows, wb_lookups, "Control Room"))
    all_suggestions.extend(build_directory_suggestions(report_rows, wb_lookups, "Switchyard"))

    stations_by_suggestion = defaultdict(list)
    for s in all_suggestions:
        key = decision_key({
            "record_type": s["record_type"],
            "record_name": s["record_name"],
            "designation": s["designation"],
            "action": s["action"],
            "old_suborg_id": s["old_suborg_id"],
            "new_suborg_id": s["new_suborg_id"],
        })
        existing = existing_decisions.get(key)
        if existing and existing["decision"] in ("APPROVED", "REJECTED"):
            continue  # never re-presented
        stations_by_suggestion[s.get("record_name")].append(s)

    total = sum(len(v) for v in stations_by_suggestion.values())
    print(f"\n{total} suggestion(s) to review.\n")

    reviewed = 0
    for name, suggestions in stations_by_suggestion.items():
        for s in suggestions:
            reviewed += 1
            print(f"\n[{reviewed}/{total}]")
            decision = prompt_decision(s, approved_by)
            if decision == "QUIT":
                print("Stopping -- progress saved, already-decided items won't be re-asked.")
                return
            append_decision({
                "record_type": s["record_type"],
                "record_name": s["record_name"],
                "designation": s["designation"],
                "old_suborg_id": s["old_suborg_id"],
                "new_suborg_id": s["new_suborg_id"],
                "action": s["action"],
                "reason": s["reason"],
                "approved_by": approved_by,
                "date": date.today().isoformat(),
                "decision": decision,
            })

    print("\nReview complete.")


if __name__ == "__main__":
    run()
