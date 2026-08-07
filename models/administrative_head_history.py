from models import db


class AdministrativeHeadHistory(db.Model):
    """Archive of ended Administrative Head / KMP assignments -- a row is
    moved here from administrative_heads when the assignment ends (see
    services/head_service.end_role)."""

    __tablename__ = "administrative_head_history"

    id = db.Column(db.Integer, primary_key=True)

    organization_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id"),
        nullable=True
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=True
    )

    role_category = db.Column(db.String(20), nullable=False)

    role_title = db.Column(db.String(150), nullable=False)

    effective_from = db.Column(db.Date, nullable=False)

    effective_to = db.Column(db.Date, nullable=False)

    reason = db.Column(db.Text)

    replacement_employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=True
    )

    remarks = db.Column(db.Text)

    service_type_id = db.Column(
        db.Integer,
        db.ForeignKey("service_types.id"),
        nullable=True
    )

    # Snapshot of the head's identity/status at the time the role ended --
    # same shape as the live administrative_heads columns they were copied
    # from (see services/head_service.end_role).
    name = db.Column(db.String(255))

    office_phone = db.Column(db.String(50))

    mobile_phone = db.Column(db.String(50))

    email = db.Column(db.String(255))

    office_address = db.Column(db.Text)

    status = db.Column(db.String(20))

    moved_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    organization = db.relationship(
        "Organization",
        lazy=True,
        foreign_keys=[organization_id]
    )

    employee = db.relationship(
        "Employee",
        lazy=True,
        foreign_keys=[employee_id]
    )

    replacement_employee = db.relationship(
        "Employee",
        lazy=True,
        foreign_keys=[replacement_employee_id]
    )

    service_type = db.relationship("ServiceType", lazy=True)

    @property
    def resolved_name(self):
        return self.employee.employee_name if self.employee_id and self.employee else (self.name or "")

    @property
    def resolved_office_phone(self):
        return self.employee.office_phone if self.employee_id and self.employee else self.office_phone

    @property
    def resolved_mobile_phone(self):
        return self.employee.mobile_phone if self.employee_id and self.employee else self.mobile_phone

    @property
    def resolved_email(self):
        return self.employee.email if self.employee_id and self.employee else self.email
