"""Live organization statistics -- computed on demand rather than stored,
since a materialized organization_statistics table would need a refresh job
and can go stale at this data scale. Mirrors the pattern already used by
routes/admin_routes.py:dashboard for its own live counts.
"""

from sqlalchemy import func

from models import db
from models.employee import Employee
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from models.update_request import UpdateRequest

EMPLOYEE_STATUSES = [
    "ACTIVE", "INACTIVE", "TRANSFERRED", "RETIRED", "DEPUTATION", "RESIGNED",
]


def get_organization_statistics(org_id):
    status_counts = dict(
        db.session.query(Employee.status, func.count(Employee.id))
        .filter(Employee.organization_id == org_id)
        .group_by(Employee.status)
        .all()
    )

    total_employees = sum(status_counts.values())

    return {
        "total_employees": total_employees,
        "by_status": {status: status_counts.get(status, 0) for status in EMPLOYEE_STATUSES},
        "departments": Department.query
            .join(Employee, Employee.department_id == Department.id)
            .filter(Employee.organization_id == org_id)
            .distinct()
            .count(),
        "control_rooms": DirectoryNumber.query
            .filter(DirectoryNumber.organization_id == org_id)
            .count(),
        "emergency_contacts": EmergencyContact.query
            .join(Employee, EmergencyContact.employee_id == Employee.id)
            .filter(Employee.organization_id == org_id)
            .count(),
        "pending_requests": UpdateRequest.query
            .join(Employee, UpdateRequest.employee_id == Employee.id)
            .filter(Employee.organization_id == org_id, UpdateRequest.status == "Pending")
            .count(),
    }
