"""Quick check — see current state of the organizations we tried to merge."""
from app import app
from models import db
from models.organization import Organization
from models.employee import Employee

TARGET_IDS = [4, 5, 6, 7, 8, 9, 27, 43, 46, 47, 56, 48, 52, 53, 62, 63, 11]

with app.app_context():
    print(f"Total organizations in DB: {Organization.query.count()}")
    print()
    print(f"{'ID':<6} {'Found?':<8} Name")
    print("-" * 60)
    for org_id in sorted(TARGET_IDS):
        org = db.session.get(Organization, org_id)
        if org:
            emp_count = Employee.query.filter_by(organization_id=org_id).count()
            print(f"{org_id:<6} YES      {org.organization_name}  ({emp_count} emp)")
        else:
            print(f"{org_id:<6} MISSING")
