from models import db


class AuditLog(db.Model):
    """Append-only change log for the new enterprise modules. No route
    should ever UPDATE or DELETE a row here -- see services/audit_service."""

    __tablename__ = "audit_logs"

    id = db.Column(db.Integer, primary_key=True)

    module = db.Column(db.String(50), nullable=False)

    record_type = db.Column(db.String(50), nullable=False)

    record_id = db.Column(db.Integer)

    action = db.Column(db.String(20), nullable=False)

    field_name = db.Column(db.String(100))

    old_value = db.Column(db.Text)

    new_value = db.Column(db.Text)

    changed_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    changed_by_label = db.Column(db.String(150))

    reason = db.Column(db.Text)

    ip_address = db.Column(db.String(64))

    browser = db.Column(db.String(100))

    operating_system = db.Column(db.String(100))

    session_id = db.Column(db.String(64))

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )
