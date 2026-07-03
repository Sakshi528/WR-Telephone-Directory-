from models import db


class UpdateRequest(db.Model):
    __tablename__ = "update_requests"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id")
    )

    directory_number_id = db.Column(
        db.Integer,
        db.ForeignKey("directory_numbers.id")
    )

    requested_mobile = db.Column(
        db.String(50)
    )

    requested_office_phone = db.Column(
        db.String(100)
    )

    requested_email = db.Column(
        db.String(255)
    )

    requested_directory_number = db.Column(
        db.String(100)
    )

    reason = db.Column(db.Text)

    # New generic fields (v2 — replaces specific phone columns for new submissions)
    requested_by = db.Column(db.String(255))
    department = db.Column(db.String(255))
    contact_number = db.Column(db.String(50))
    field_name = db.Column(db.String(100))
    new_value = db.Column(db.Text)

    request_type = db.Column(
        db.String(20),
        default="employee"
    )

    status = db.Column(
        db.String(20),
        default="Pending"
    )

    request_date = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    user = db.relationship(
        "User",
        back_populates="update_requests"
    )

    employee = db.relationship(
        "Employee"
    )

    directory_number = db.relationship(
        "DirectoryNumber"
    )
