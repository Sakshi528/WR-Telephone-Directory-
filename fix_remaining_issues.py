"""Fixes the two remaining data issues found during audit."""
from app import app
from models import db
from models.employee import Employee
from models.directory_number import DirectoryNumber

with app.app_context():

    # 1. Fix Kokilaben Hospital — missing org in DirectoryNumber
    dn = db.session.get(DirectoryNumber, 220)
    if dn and not dn.organization:
        dn.organization = "Hospitals"
        print(f"Fixed: id=220 '{dn.name}' → org set to 'Hospitals'")
    else:
        print("id=220 already has org or not found")

    # 2. Remove duplicate S.S.Patel at MP SLDC — keep the one with more data
    dupes = Employee.query.filter(
        Employee.employee_name == "S.S.Patel"
    ).all()
    if len(dupes) > 1:
        # Keep the one that has more fields filled in
        dupes.sort(key=lambda e: sum([
            bool(e.designation), bool(e.office_phone),
            bool(e.mobile_phone), bool(e.email)
        ]), reverse=True)
        for emp in dupes:
            print(f"  id={emp.id} | {emp.employee_name} | {emp.designation} | "
                  f"office={emp.office_phone} | mobile={emp.mobile_phone}")
        to_delete = dupes[1:]
        for emp in to_delete:
            db.session.delete(emp)
            print(f"  Deleted duplicate id={emp.id}")
    else:
        print("No S.S.Patel duplicate found")

    db.session.commit()
    print("\nDone.")
