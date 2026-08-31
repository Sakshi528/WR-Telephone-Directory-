"""Applies office_phone backfills from a batch-c review chunk CSV (schema:
organization,employee,field,current_value,raw_word_value,proposed_value,
status) to the live Employee table, and appends to
reports/session_changes_audit_20260831.csv -- same read-before-write /
audit-trail pattern as apply_session_fixes_20260831.py, extended to handle
a whole CSV of rows instead of a hardcoded list.

Matching: organization -> Organization.organization_name (exact),
employee -> Employee.employee_name (exact) within that org. Safe because
these strings were both written verbatim from live Organization/Employee
rows when the proposal CSV was built (see build_batch_c_employee_fields.py
/ full_word_vs_app_field_diff.py) -- not fuzzy-matched here.

A row is skipped (not applied, not counted as an error) if its employee
name is in EXCLUDE_EMPLOYEES for its org -- used for known placeholder
values like Word's own "__".

Refuses to touch a row if: the org isn't found, the employee name doesn't
resolve to exactly one Employee row in that org, or office_phone is
already non-blank (never overwrites existing data) -- these are reported
as skipped/problem rows, not applied silently.

Usage:
  python scripts/apply_batch_c_officephone.py <chunk.csv>              # dry run
  python scripts/apply_batch_c_officephone.py <chunk.csv> --apply      # commit + audit
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.employee import Employee

AUDIT_CSV = "reports/session_changes_audit_20260831.csv"
SESSION_DATE = "2026-08-31"

EXCLUDE_EMPLOYEES = {
    ("Dadra And Nagar Haveli", "DR. ARUN T."),  # Word's own "__" placeholder, not real data
}


def main():
    if len(sys.argv) < 2:
        print("Usage: python scripts/apply_batch_c_officephone.py <chunk.csv> [--apply]")
        sys.exit(1)
    chunk_path = sys.argv[1]
    apply = "--apply" in sys.argv

    with open(chunk_path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    with app.app_context():
        orgs = {o.organization_name: o.id for o in Organization.query.all()}
        all_emps = Employee.query.all()
        by_org_name = {}
        for e in all_emps:
            by_org_name.setdefault(e.organization_id, {}).setdefault(e.employee_name, []).append(e)

        applied, skipped_excluded, problems = [], [], []

        for r in rows:
            key = (r["organization"], r["employee"])
            if key in EXCLUDE_EMPLOYEES:
                skipped_excluded.append(r)
                continue

            org_id = orgs.get(r["organization"])
            if org_id is None:
                problems.append(("NO ORG MATCH", r))
                continue
            candidates = by_org_name.get(org_id, {}).get(r["employee"], [])
            if len(candidates) != 1:
                problems.append((f"{len(candidates)} EMPLOYEE MATCHES", r))
                continue
            emp = candidates[0]
            if emp.office_phone:
                problems.append(("ALREADY NON-BLANK, NOT OVERWRITTEN", r))
                continue

            old_value = emp.office_phone
            new_value = r["proposed_value"]
            if apply:
                emp.office_phone = new_value

            applied.append({
                "table": "employees", "row_id": emp.id, "field": "office_phone",
                "old_value": old_value or "", "new_value": new_value,
                "reason": f"batch c office_phone backfill, {r['organization']} / {r['employee']} "
                          f"-- Word value {r['raw_word_value']!r} applied as-is.",
                "timestamp": SESSION_DATE,
            })

        if apply:
            db.session.commit()

    if apply and applied:
        file_exists = os.path.exists(AUDIT_CSV)
        with open(AUDIT_CSV, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"])
            if not file_exists:
                w.writeheader()
            w.writerows(applied)

    print("=" * 70)
    print(f"  {'APPLIED' if apply else 'DRY RUN -- nothing written to DB'}: {chunk_path}")
    print("=" * 70)
    print(f"  Rows in chunk           : {len(rows)}")
    print(f"  Excluded (known bad)    : {len(skipped_excluded)}")
    for r in skipped_excluded:
        print(f"    - {r['organization']} / {r['employee']} (raw={r['raw_word_value']!r})")
    print(f"  Applied                 : {len(applied)}")
    print(f"  Problems (not applied)  : {len(problems)}")
    for kind, r in problems:
        print(f"    - [{kind}] {r['organization']} / {r['employee']}")
    print(f"  Audit trail -> {AUDIT_CSV}" if apply else "")
    print("=" * 70)


if __name__ == "__main__":
    main()
