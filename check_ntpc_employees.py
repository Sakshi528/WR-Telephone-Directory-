"""Check if the 2 NTPC RED employees are still in the database and where they landed."""
from app import app
from models import db
from models.employee import Employee
from models.organization import Organization

with app.app_context():
    for name in ["Subrata Mandal", "Ashwini Kumar Tripathy"]:
        emp = Employee.query.filter(Employee.employee_name.ilike(f"%{name}%")).first()
        if emp:
            org = db.session.get(Organization, emp.organization_id) if emp.organization_id else None
            print(f"FOUND: {emp.employee_name} | {emp.designation}")
            print(f"  Org ID : {emp.organization_id}")
            print(f"  Org Name: {org.organization_name if org else '(organization deleted / orphaned)'}")
        else:
            print(f"NOT FOUND: {name}")
        print()
