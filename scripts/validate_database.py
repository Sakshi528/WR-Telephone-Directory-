"""
Validates the current state of the PostgreSQL database.

Read-only: never inserts, updates, or deletes anything.

Checks:
  - organization / sub-organization / employee / control room / switchyard counts
  - duplicate employee IDs (primary-key integrity sanity check)
  - duplicate employee phone numbers (data-quality warning, not treated as an error --
    a shared office/mobile number can be legitimate)
  - orphan foreign keys (Employee.department_id, EmergencyContact.employee_id)
  - invalid suborg_id (Organization.parent_id referencing a non-existent organization --
    this schema represents sub-organizations as Organization rows with parent_id set,
    there is no separate suborganizations table)
  - invalid org_id (Employee.organization_id / DirectoryNumber.organization_id
    referencing a non-existent organization)

Usage:
  python scripts/validate_database.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.organization import Organization
from models.employee import Employee
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact


class Issue:
    def __init__(self, severity, check, message):
        self.severity = severity  # "ERROR" | "WARNING"
        self.check = check
        self.message = message


def run():
    sep = "=" * 70
    print(sep)
    print("  Database Validation Report")
    print(sep)

    issues = []

    with app.app_context():

        org_count = Organization.query.count()
        suborg_count = Organization.query.filter(Organization.parent_id.isnot(None)).count()
        employee_count = Employee.query.count()
        control_room_count = DirectoryNumber.query.filter_by(category="Control Room").count()
        switchyard_count = DirectoryNumber.query.filter_by(category="Switchyard").count()

        print("\n  Counts:")
        print(f"    Organizations     : {org_count}")
        print(f"    Sub-organizations : {suborg_count}")
        print(f"    Employees         : {employee_count}")
        print(f"    Control rooms     : {control_room_count}")
        print(f"    Switchyards       : {switchyard_count}")

        # ── Duplicate employee IDs (PK integrity sanity check) ───────────────
        dup_ids = (
            db.session.query(Employee.id, db.func.count(Employee.id))
            .group_by(Employee.id)
            .having(db.func.count(Employee.id) > 1)
            .all()
        )
        for emp_id, count in dup_ids:
            issues.append(Issue("ERROR", "duplicate employee IDs",
                                 f"employee id={emp_id} appears {count} times"))

        # ── Duplicate phone numbers (employee mobile_phone) ──────────────────
        dup_phones = (
            db.session.query(Employee.mobile_phone, db.func.count(Employee.id))
            .filter(Employee.mobile_phone.isnot(None), Employee.mobile_phone != "")
            .group_by(Employee.mobile_phone)
            .having(db.func.count(Employee.id) > 1)
            .all()
        )
        for phone, count in dup_phones:
            issues.append(Issue("WARNING", "duplicate phone numbers",
                                 f"mobile_phone='{phone}' is shared by {count} employees"))

        # ── Orphan foreign keys ──────────────────────────────────────────────
        dept_ids = {d.id for d in Department.query.all()}
        for e in Employee.query.filter(Employee.department_id.isnot(None)).all():
            if e.department_id not in dept_ids:
                issues.append(Issue("ERROR", "orphan foreign keys",
                                     f"employee id={e.id} name='{e.employee_name}' has "
                                     f"department_id={e.department_id} which does not exist"))

        employee_ids = {e.id for e in Employee.query.all()}
        for ec in EmergencyContact.query.all():
            if ec.employee_id not in employee_ids:
                issues.append(Issue("ERROR", "orphan foreign keys",
                                     f"emergency_contact id={ec.id} references employee_id="
                                     f"{ec.employee_id} which does not exist"))

        # ── Invalid suborg_id (Organization.parent_id) ───────────────────────
        org_ids = {o.id for o in Organization.query.all()}
        for o in Organization.query.filter(Organization.parent_id.isnot(None)).all():
            if o.parent_id not in org_ids:
                issues.append(Issue("ERROR", "invalid suborg_id",
                                     f"organization id={o.id} name='{o.organization_name}' has "
                                     f"parent_id={o.parent_id} which does not exist"))

        # ── Invalid org_id (Employee / DirectoryNumber -> organizations) ─────
        for e in Employee.query.filter(Employee.organization_id.isnot(None)).all():
            if e.organization_id not in org_ids:
                issues.append(Issue("ERROR", "invalid org_id",
                                     f"employee id={e.id} name='{e.employee_name}' has "
                                     f"organization_id={e.organization_id} which does not exist"))

        for dn in DirectoryNumber.query.filter(DirectoryNumber.organization_id.isnot(None)).all():
            if dn.organization_id not in org_ids:
                issues.append(Issue("ERROR", "invalid org_id",
                                     f"directory_number id={dn.id} name='{dn.name}' has "
                                     f"organization_id={dn.organization_id} which does not exist"))

    errors = [i for i in issues if i.severity == "ERROR"]
    warnings = [i for i in issues if i.severity == "WARNING"]

    for severity, group in (("ERROR", errors), ("WARNING", warnings)):
        print(f"\n  {severity}S ({len(group)}):")
        if not group:
            print("    (none)")
            continue
        by_check = {}
        for issue in group:
            by_check.setdefault(issue.check, []).append(issue.message)
        for check, messages in by_check.items():
            print(f"\n    [{check}] {len(messages)} {severity.lower()}(s)")
            for msg in messages[:50]:
                print(f"      - {msg}")
            if len(messages) > 50:
                print(f"      ... and {len(messages) - 50} more")

    print(f"\n{sep}")
    print(f"  TOTAL: {len(errors)} error(s), {len(warnings)} warning(s)")
    print(sep)

    return len(errors) == 0


if __name__ == "__main__":
    ok = run()
    sys.exit(0 if ok else 1)
