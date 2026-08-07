from models import db


class EmailGroupFilter(db.Model):
    """One filter criterion for a DYNAMIC EmailGroup. Rows sharing the same
    filter_type are ORed together; different filter_types are ANDed --
    e.g. (category=Transmission Utility OR category=SLDC) AND status=ACTIVE.
    Exactly one of organization_id/category_id/role_value/status_value is
    populated per row, matching the filter_type."""

    __tablename__ = "email_group_filters"

    id = db.Column(db.Integer, primary_key=True)

    group_id = db.Column(
        db.Integer,
        db.ForeignKey("email_groups.id", ondelete="CASCADE"),
        nullable=False
    )

    filter_type = db.Column(db.String(30), nullable=False)  # ORGANIZATION | ORGANIZATION_CATEGORY | ROLE | STATUS

    organization_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id", ondelete="CASCADE"),
        nullable=True
    )

    category_id = db.Column(
        db.Integer,
        db.ForeignKey("organization_categories.id", ondelete="CASCADE"),
        nullable=True
    )

    role_value = db.Column(db.String(30))  # UTILITY_HEAD | ADMINISTRATIVE_HEAD | KMP

    status_value = db.Column(db.String(20))  # ACTIVE | INACTIVE | TRANSFERRED | RETIRED | DEPUTATION | RESIGNED

    organization = db.relationship("Organization", lazy=True)

    category = db.relationship("OrganizationCategory", lazy=True)
