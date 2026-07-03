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

    address = db.Column(
        db.Text
    )

    parent_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id", ondelete="SET NULL"),
        nullable=True
    )

    children = db.relationship(
        "Organization",
        backref=db.backref("parent", remote_side="Organization.id"),
        lazy="dynamic",
        foreign_keys="Organization.parent_id"
    )
