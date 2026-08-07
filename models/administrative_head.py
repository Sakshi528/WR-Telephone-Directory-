from models import db


class AdministrativeHead(db.Model):
    """Current organization role assignments -- Administrative Heads and KMP
    share this table, distinguished by role_category. Only one ACTIVE
    ADMINISTRATIVE_HEAD row may exist per (organization_id, role_title)
    (enforced by a partial unique index); KMP allows multiple concurrent
    holders of the same role_title.

    employee_id is optional: a head doesn't have to already be an Employee
    (e.g. a Principal Secretary or Collector who isn't in the directory).
    When employee_id is set, identity/contact info is read live from that
    Employee via the resolved_* properties; when it's NULL, the head's own
    name/office_phone/mobile_phone/email/office_address columns are used
    instead -- same pattern as AdministrativeHeadAssistant. A CHECK
    constraint (chk_administrative_head_identity) requires at least one of
    employee_id/name to be set."""

    __tablename__ = "administrative_heads"

    id = db.Column(db.Integer, primary_key=True)

    organization_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id", ondelete="CASCADE"),
        nullable=False
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id", ondelete="CASCADE"),
        nullable=True
    )

    role_category = db.Column(db.String(20), nullable=False)  # ADMINISTRATIVE_HEAD | KMP

    role_title = db.Column(db.String(150), nullable=False)

    service_type_id = db.Column(
        db.Integer,
        db.ForeignKey("service_types.id", ondelete="SET NULL"),
        nullable=True
    )

    effective_from = db.Column(db.Date, nullable=False)

    # ACTIVE | ON_LEAVE | VACANT | INACTIVE -- distinct from ending the role
    # (which archives to AdministrativeHeadHistory): a VACANT/ON_LEAVE head
    # keeps their tenure record but drops out of Email Distribution by default.
    status = db.Column(db.String(20), nullable=False, default="ACTIVE")

    remarks = db.Column(db.Text)

    # Standalone identity/contact -- used only when employee_id IS NULL.
    name = db.Column(db.String(255))

    office_phone = db.Column(db.String(50))

    mobile_phone = db.Column(db.String(50))

    email = db.Column(db.String(255))

    office_address = db.Column(db.Text)

    created_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    organization = db.relationship("Organization", lazy=True)

    employee = db.relationship("Employee", lazy=True)

    service_type = db.relationship("ServiceType", lazy=True)

    assistants = db.relationship(
        "AdministrativeHeadAssistant",
        backref="administrative_head",
        lazy=True,
        cascade="all, delete-orphan",
        order_by="AdministrativeHeadAssistant.is_primary.desc()"
    )

    @property
    def resolved_name(self):
        return self.employee.employee_name if self.employee_id and self.employee else (self.name or "")

    @property
    def resolved_designation(self):
        if self.employee_id and self.employee:
            return self.employee.designation or self.role_title
        return self.role_title

    @property
    def resolved_office_phone(self):
        return self.employee.office_phone if self.employee_id and self.employee else self.office_phone

    @property
    def resolved_mobile_phone(self):
        return self.employee.mobile_phone if self.employee_id and self.employee else self.mobile_phone

    @property
    def resolved_email(self):
        return self.employee.email if self.employee_id and self.employee else self.email

    @property
    def resolved_office_address(self):
        if self.office_address:
            return self.office_address
        return self.organization.address if self.organization else None
