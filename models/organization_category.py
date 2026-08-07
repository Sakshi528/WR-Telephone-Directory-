from models import db


class OrganizationCategory(db.Model):
    """Normalized classification for an Organization (Transmission Utility,
    Generator, SLDC, etc.) -- lets Email Distribution Lists build category
    based mailing lists without free-text matching."""

    __tablename__ = "organization_categories"

    id = db.Column(db.Integer, primary_key=True)

    category_name = db.Column(db.String(100), nullable=False, unique=True)

    description = db.Column(db.Text)

    # Drives whether the cascading Organization Type dropdown shows a State
    # step for this type (e.g. State SLDC/STU/DISCOM) or goes straight to
    # Organization (e.g. CTU/RLDC).
    is_state_based = db.Column(db.Boolean, nullable=False, default=False)

    organizations = db.relationship(
        "Organization",
        backref="category",
        lazy=True
    )
