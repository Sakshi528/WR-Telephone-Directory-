from models import db


class OrganizationSubcategory(db.Model):
    """A curated, per-Category list of Subcategory values (e.g. under
    Category "Generation Company": "ISTS Connected RE Generators In
    Western Region", "Other Generating Stations In Western Region").
    Organization.region stores the free-text subcategory an org actually
    carries -- this table exists purely so the Add/Edit Organization form
    can suggest previously-used values instead of everyone retyping them,
    and so an admin can pre-register a new subcategory before it's ever
    used, without needing a code change/migration."""

    __tablename__ = "organization_subcategories"

    id = db.Column(db.Integer, primary_key=True)

    category_id = db.Column(
        db.Integer,
        db.ForeignKey("organization_categories.id", ondelete="CASCADE"),
        nullable=False
    )

    subcategory_name = db.Column(db.String(255), nullable=False)

    category = db.relationship("OrganizationCategory", backref="subcategories", lazy=True)

    __table_args__ = (
        db.UniqueConstraint("category_id", "subcategory_name", name="uq_subcategory_per_category"),
    )
