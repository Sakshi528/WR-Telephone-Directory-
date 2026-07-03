"""Quick check — shows employees in the ambiguous org records."""
from app import app
from models import db
from models.organization import Organization
from models.employee import Employee

with app.app_context():
    for org_id, label in [(9, "NTPC"), (181, "Adani Power Limited")]:
        org = db.session.get(Organization, org_id)
        emps = Employee.query.filter_by(organization_id=org_id).all()
        print(f"\n[{org_id}] {org.organization_name if org else '(not found)'} — {len(emps)} employee(s)")
        for e in emps:
            print(f"  - {e.employee_name} | {e.designation} | {e.location or '—'}")
