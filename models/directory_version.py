from models import db


class DirectoryVersion(db.Model):
    """One row per generated/regenerated official snapshot of the telephone
    directory. No employee data is duplicated here -- see
    services/directory_version_service.py for the snapshot query and
    services/pdf_generator.py / excel_generator.py for the files this row
    points at. Regeneration never mutates an existing row or its files; it
    inserts a new one with a revision-suffixed version_number."""

    __tablename__ = "directory_versions"

    id = db.Column(db.Integer, primary_key=True)

    version_number = db.Column(db.String(20), nullable=False, unique=True)
    version_name = db.Column(db.String(255), nullable=False)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)

    generated_on = db.Column(db.DateTime, server_default=db.func.now())
    generated_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    employee_count = db.Column(db.Integer, nullable=False, default=0)
    organization_count = db.Column(db.Integer, nullable=False, default=0)

    pdf_filename = db.Column(db.String(255))
    pdf_path = db.Column(db.String(500))
    excel_filename = db.Column(db.String(255))
    excel_path = db.Column(db.String(500))

    status = db.Column(db.String(20), nullable=False, default="Published")
    remarks = db.Column(db.Text)

    created_at = db.Column(db.DateTime, server_default=db.func.now())
    updated_at = db.Column(db.DateTime, server_default=db.func.now(), onupdate=db.func.now())

    generated_by_user = db.relationship("User", lazy=True)
