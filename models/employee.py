from models import db

EMPLOYEE_STATUSES = [
    "ACTIVE", "INACTIVE", "TRANSFERRED", "RETIRED", "DEPUTATION", "RESIGNED",
    "CONTRACT_ENDED", "DECEASED",
]


class Employee(db.Model):
    __tablename__ = "employees"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    employee_name = db.Column(
        db.String(255),
        nullable=False
    )

    designation = db.Column(
        db.String(255)
    )

    organization_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id")
    )

    department_id = db.Column(
        db.Integer,
        db.ForeignKey("departments.id")
    )

    location = db.Column(
        db.String(255)
    )
    region = db.Column(
        db.String(255)
    )

    office_phone = db.Column(
        db.Text
    )
    
    residence_phone = db.Column(
        db.Text
    )

    mobile_phone = db.Column(
        db.Text
    )

    email = db.Column(
        db.String(255)
    )

    is_utility_head = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )

    is_kmp = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )

    status = db.Column(
        db.String(20),
        nullable=False,
        default="ACTIVE"
    )

    status_changed_at = db.Column(
        db.DateTime
    )

    organization = db.relationship(
        "Organization",
        lazy=True
    )

    department = db.relationship(
        "Department",
        lazy=True
    )
