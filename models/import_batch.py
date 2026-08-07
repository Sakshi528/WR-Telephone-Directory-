from models import db


class ImportBatch(db.Model):
    """One row per run of import_from_excel_db.py / import_employees_excel.py
    -- the first persisted record of an import run (previously log-file
    only). Written at the end of each script's run() using counters those
    scripts already compute. rollback_available is always False for now;
    undoing an already-committed import isn't implemented."""

    __tablename__ = "import_batches"

    id = db.Column(db.Integer, primary_key=True)

    workbook_name = db.Column(db.String(255), nullable=False)

    imported_by = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=True
    )

    import_date = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    mode = db.Column(db.String(20), nullable=False)  # DRY_RUN | APPLIED

    inserted_count = db.Column(db.Integer, nullable=False, default=0)

    updated_count = db.Column(db.Integer, nullable=False, default=0)

    skipped_count = db.Column(db.Integer, nullable=False, default=0)

    status = db.Column(db.String(20), nullable=False)  # SUCCESS | FAILED

    rollback_available = db.Column(db.Boolean, nullable=False, default=False)

    summary = db.Column(db.Text)

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )

    imported_by_user = db.relationship("User", lazy=True)
