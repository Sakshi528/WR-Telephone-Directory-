from models import db


class AdministrativeHeadAssistant(db.Model):
    """A Personal Assistant/Secretary attached to an Administrative Head.
    One head may have several; at most one is_primary=True per head
    (enforced by a partial unique index). If employee_id is set, contact
    info is read live from that Employee record (see the resolved_*
    properties) rather than duplicated here -- employee_id stays NULL for
    assistants who aren't in the employee directory, same reasoning as
    EmergencyContact."""

    __tablename__ = "administrative_head_assistants"

    id = db.Column(db.Integer, primary_key=True)

    administrative_head_id = db.Column(
        db.Integer,
        db.ForeignKey("administrative_heads.id", ondelete="CASCADE"),
        nullable=False
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id", ondelete="SET NULL"),
        nullable=True
    )

    designation = db.Column(db.String(100), nullable=False)  # PA / PS / Executive Assistant / ...

    name = db.Column(db.String(255))  # used only when employee_id IS NULL

    office_phone = db.Column(db.String(50))

    mobile = db.Column(db.String(50))

    email = db.Column(db.String(255))

    is_primary = db.Column(db.Boolean, nullable=False, default=False)

    created_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    employee = db.relationship("Employee", lazy=True)

    @property
    def resolved_name(self):
        return self.employee.employee_name if self.employee_id and self.employee else (self.name or "")

    @property
    def resolved_office_phone(self):
        return self.employee.office_phone if self.employee_id and self.employee else self.office_phone

    @property
    def resolved_mobile(self):
        return self.employee.mobile_phone if self.employee_id and self.employee else self.mobile

    @property
    def resolved_email(self):
        return self.employee.email if self.employee_id and self.employee else self.email
