from models import db


class EmailGroup(db.Model):
    """A saved, named distribution group. STATIC groups hold an explicit
    employee membership list (GroupMember rows). DYNAMIC groups instead hold
    filter criteria (EmailGroupFilter rows) and are resolved live via
    services/email_distribution_service.resolve_dynamic_group -- their
    membership always reflects the current database, never a snapshot."""

    __tablename__ = "email_groups"

    id = db.Column(db.Integer, primary_key=True)

    name = db.Column(db.String(150), nullable=False, unique=True)

    description = db.Column(db.Text)

    list_type = db.Column(db.String(10), nullable=False, default="STATIC")  # STATIC | DYNAMIC

    created_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    members = db.relationship(
        "GroupMember",
        backref="group",
        lazy=True,
        cascade="all, delete-orphan"
    )

    filters = db.relationship(
        "EmailGroupFilter",
        backref="group",
        lazy=True,
        cascade="all, delete-orphan"
    )


class GroupMember(db.Model):
    __tablename__ = "group_members"
    __table_args__ = (
        db.UniqueConstraint("group_id", "employee_id", name="uq_group_member"),
    )

    id = db.Column(db.Integer, primary_key=True)

    group_id = db.Column(
        db.Integer,
        db.ForeignKey("email_groups.id", ondelete="CASCADE"),
        nullable=False
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id", ondelete="CASCADE"),
        nullable=False
    )

    added_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    employee = db.relationship("Employee", lazy=True)
