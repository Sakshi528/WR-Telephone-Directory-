"""Replaces WRLDC employees with the current list from 'uploads/employee
list.xlsx'. Looks up (or creates) the WRLDC org, deletes its existing
employees and emergency contacts, then imports the Excel rows and any new
departments. Delete + insert run inside one transaction, rolled back whole
on failure.

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
from models.import_batch import ImportBatch
from utils.logger import get_logger

logger = get_logger(__name__)

EXCEL_PATH = "uploads/employee list.xlsx"
WRLDC_ORG_NAME = "Western Region Load Despatch Centre"


def run(apply: bool) -> None:
    logger.info("=" * 65)
    logger.info("WRLDC Employee Import")
    logger.info("DRY RUN" if not apply else "APPLY MODE — changes will be committed")
    logger.info("=" * 65)

    with app.app_context():

        wrldc = Organization.query.filter_by(organization_name=WRLDC_ORG_NAME).first()
        if wrldc:
            logger.info("Target org: [%d] %s", wrldc.id, wrldc.organization_name)
            existing = Employee.query.filter_by(organization_id=wrldc.id).all()
        else:
            logger.info("Target org '%s' not found — will be created on --apply.", WRLDC_ORG_NAME)
            existing = []

        logger.info("Employees to DELETE: %d", len(existing))
        for emp in existing:
            logger.info("  - %s | %s", emp.employee_name, emp.designation)

        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))[1:]   # skip header
        rows = [r for r in rows if r[1]]                  # skip blank names

        logger.info("Employees to IMPORT: %d", len(rows))

        dept_names = sorted(set(str(r[4]).strip() for r in rows if r[4]))
        existing_depts = {
            d.department_name.strip().lower(): d
            for d in Department.query.all()
        }
        new_depts = [n for n in dept_names if n.lower() not in existing_depts]
        if new_depts:
            logger.info("New departments to CREATE: %s", new_depts)
        else:
            logger.info("All %d departments already exist.", len(dept_names))

        if not apply:
            logger.info("Dry run complete — rerun with --apply to commit")
            return

        try:
            if wrldc is None:
                wrldc = Organization(organization_name=WRLDC_ORG_NAME, parent_id=None)
                db.session.add(wrldc)
                db.session.flush()
                logger.info("Created organization: [%d] %s", wrldc.id, wrldc.organization_name)

            emp_ids = [e.id for e in existing]
            if emp_ids:
                EmergencyContact.query.filter(
                    EmergencyContact.employee_id.in_(emp_ids)
                ).delete(synchronize_session=False)
                Employee.query.filter_by(organization_id=wrldc.id).delete(
                    synchronize_session=False
                )
                logger.info(
                    "Deleted %d existing employees and their emergency contacts.",
                    len(existing),
                )

            dept_map = dict(existing_depts)   # name.lower() → Department obj
            for name in new_depts:
                dept = Department(department_name=name)
                db.session.add(dept)
                db.session.flush()
                dept_map[name.lower()] = dept
                logger.info("Created department: %s", name)

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
                    organization_id = wrldc.id,
                    department_id   = dept_obj.id if dept_obj else None,
                    region          = region,
                    location        = location,
                    mobile_phone    = mobile or None,
                    email           = email or None,
                    is_kmp          = False,   # WRLDC staff are never KMP — explicit, not relying on the model default
                )
                db.session.add(emp)
                imported += 1

            db.session.commit()
            logger.info("Imported %d employees into '%s'.", imported, wrldc.organization_name)
            logger.info("=" * 65)

            # First persisted record of this run -- Archive > Import History.
            # This pipeline is delete-then-reinsert, not a row-level diff,
            # so "updated" isn't a meaningful count here.
            db.session.add(ImportBatch(
                workbook_name=EXCEL_PATH, mode="APPLIED", status="SUCCESS",
                inserted_count=imported, updated_count=0, skipped_count=0,
                summary=f"Imported {imported} employees into '{wrldc.organization_name}' "
                        f"({len(existing)} previous records deleted first).",
            ))
            db.session.commit()

        except Exception as exc:
            logger.exception("Import failed — rolling back all changes.")
            db.session.rollback()
            db.session.add(ImportBatch(
                workbook_name=EXCEL_PATH, mode="APPLIED", status="FAILED",
                summary=f"Import failed: {exc}",
            ))
            db.session.commit()
            raise


if __name__ == "__main__":
    run(apply="--apply" in sys.argv)
