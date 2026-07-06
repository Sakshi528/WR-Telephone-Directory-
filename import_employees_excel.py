"""
Replaces WRLDC employees with the current list from 'uploads/employee list.xlsx'.

Steps:
  1. Deletes all employees currently in org "Western Region Load Despatch Centre" (id=11)
     and their emergency contacts.
  2. Imports 98 employees from the Excel file into the same org.
  3. Creates any departments that don't already exist.

Usage:
  python import_employees_excel.py          # dry run
  python import_employees_excel.py --apply  # commit changes
"""

import sys
import openpyxl
from app import app
from models import db
from models.employee import Employee
from models.organization import Organization
from models.department import Department
from models.emergency_contact import EmergencyContact

EXCEL_PATH = "uploads/employee list.xlsx"
WRLDC_ORG_ID = 5309  # WRLDC


def run(apply: bool) -> None:
    sep = "=" * 65
    print(sep)
    print("  WRLDC Employee Import")
    print("  DRY RUN" if not apply else "  APPLY MODE â€” changes will be committed")
    print(sep)

    with app.app_context():

        # â”€â”€ Verify org exists â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        wrldc = db.session.get(Organization, WRLDC_ORG_ID)
        if not wrldc:
            print(f"  ERROR: Organization id={WRLDC_ORG_ID} not found. Aborting.")
            return
        print(f"\n  Target org: [{WRLDC_ORG_ID}] {wrldc.organization_name}")

        # â”€â”€ Count records to be deleted â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        existing = Employee.query.filter_by(organization_id=WRLDC_ORG_ID).all()
        print(f"\n  Employees to DELETE: {len(existing)}")
        for emp in existing:
            print(f"    - {emp.employee_name} | {emp.designation}")

        # â”€â”€ Read Excel â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))[1:]   # skip header
        rows = [r for r in rows if r[1]]                  # skip blank names

        print(f"\n  Employees to IMPORT: {len(rows)}")

        # â”€â”€ Department lookup / creation plan â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        dept_names = sorted(set(str(r[4]).strip() for r in rows if r[4]))
        existing_depts = {
            d.department_name.strip().lower(): d
            for d in Department.query.all()
        }
        new_depts = [n for n in dept_names if n.lower() not in existing_depts]
        if new_depts:
            print(f"\n  New departments to CREATE: {new_depts}")
        else:
            print(f"\n  All {len(dept_names)} departments already exist.")

        if not apply:
            print(f"\n  [Dry run complete â€” rerun with --apply to commit]")
            print(sep)
            return

        # â”€â”€ DELETE existing WRLDC employees â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        emp_ids = [e.id for e in existing]
        if emp_ids:
            EmergencyContact.query.filter(
                EmergencyContact.employee_id.in_(emp_ids)
            ).delete(synchronize_session=False)
            Employee.query.filter_by(organization_id=WRLDC_ORG_ID).delete(
                synchronize_session=False
            )
            print(f"\n  Deleted {len(existing)} existing employees and their emergency contacts.")

        # â”€â”€ Create missing departments â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        dept_map = dict(existing_depts)   # name.lower() â†’ Department obj
        for name in new_depts:
            dept = Department(department_name=name)
            db.session.add(dept)
            db.session.flush()
            dept_map[name.lower()] = dept
            print(f"  Created department: {name}")

        # â”€â”€ INSERT new employees â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        imported = 0
        for row in rows:
            empno, name, level, designation, dept_name, region, location, mobile, email = row

            name        = str(name).strip() if name else ""
            designation = str(designation).strip() if designation else ""
            dept_name   = str(dept_name).strip() if dept_name else ""
            region      = str(region).strip() if region else ""
            location    = str(location).strip() if location else ""
            mobile      = str(int(mobile)) if isinstance(mobile, (int, float)) else str(mobile or "").strip()
            email       = str(email).strip() if email else ""

            if not name:
                continue

            dept_obj = dept_map.get(dept_name.lower())

            emp = Employee(
                employee_name   = name,
                designation     = designation,
                organization_id = WRLDC_ORG_ID,
                department_id   = dept_obj.id if dept_obj else None,
                region          = region,
                location        = location,
                mobile_phone    = mobile or None,
                email           = email or None,
            )
            db.session.add(emp)
            imported += 1

        db.session.commit()
        print(f"\n  Imported {imported} employees into '{wrldc.organization_name}'.")
        print(sep)


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
