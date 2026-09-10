from models import db

class Organization(db.Model):

    __tablename__ = "organizations"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    organization_name = db.Column(
        db.String(500),
        unique=True
    )

    region = db.Column(
        db.String(255)
    )

    # Indian state, distinct from `region` above (which holds the parent
    # import-grouping label, e.g. "Nuclear Power Stations in Western
    # Region"). Drives the State level of the Type -> State -> Organization
    # cascading dropdowns for state-based Organization Types.
    state = db.Column(
        db.String(100)
    )

    address = db.Column(
        db.Text
    )

    parent_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id", ondelete="SET NULL"),
        nullable=True
    )

    category_id = db.Column(
        db.Integer,
        db.ForeignKey("organization_categories.id", ondelete="SET NULL"),
        nullable=True
    )

    # When True, this organization is explicitly marked as having no Utility
    # Head -- overrides utils.designation_rank.resolve_utility_head, which
    # would otherwise always auto-resolve someone from the employee list.
    utility_head_excluded = db.Column(
        db.Boolean,
        nullable=False,
        server_default=db.text("false")
    )

    children = db.relationship(
        "Organization",
        backref=db.backref("parent", remote_side="Organization.id"),
        lazy="dynamic",
        foreign_keys="Organization.parent_id"
    )
