from models import db


class EmployeeStatusHistory(db.Model):
    """Despite the name, this is the general per-employee change log the
    Archive module's timeline is built from: a row is written whenever
    status, organization, or designation changes, populating only the
    fields that actually changed (the other pair stays NULL). One reused
    table drives the Status Timeline, Organization History, and
    Designation History sections of an Archive employee profile."""

    __tablename__ = "employee_status_history"

    id = db.Column(db.Integer, primary_key=True)

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id", ondelete="CASCADE"),
        nullable=False
    )

    old_status = db.Column(db.String(20))

    new_status = db.Column(db.String(20), nullable=False)

    old_organization_id = db.Column(db.Integer, db.ForeignKey("organizations.id"), nullable=True)

    new_organization_id = db.Column(db.Integer, db.ForeignKey("organizations.id"), nullable=True)

    old_designation = db.Column(db.String(255))

    new_designation = db.Column(db.String(255))

    reason = db.Column(db.Text)

    changed_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    changed_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    employee = db.relationship("Employee", lazy=True, foreign_keys=[employee_id])

    changed_by_user = db.relationship("User", lazy=True)

    old_organization = db.relationship("Organization", lazy=True, foreign_keys=[old_organization_id])

    new_organization = db.relationship("Organization", lazy=True, foreign_keys=[new_organization_id])
