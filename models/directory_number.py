from models import db


class DirectoryNumber(db.Model):
    __tablename__ = "directory_numbers"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    name = db.Column(
        db.String(255),
        nullable=False
    )

    phone_number = db.Column(
        db.Text
    )
    
    email = db.Column(
        db.Text
    )
    
    organization = db.Column(
        db.String(255)
    )

    category = db.Column(
        db.String(100)
    )

    switch_yard = db.Column(
        db.Boolean,
        nullable=False,
        server_default=db.text("false")
    )

    control_room = db.Column(
        db.Boolean,
        nullable=False,
        server_default=db.text("false")
    )

    ip_address = db.Column(
        db.String(45)
    )

    organization_id = db.Column(
        db.Integer,
        db.ForeignKey("organizations.id", ondelete="SET NULL"),
        nullable=True
    )

    organization_obj = db.relationship(
        "Organization",
        lazy=True,
        foreign_keys="DirectoryNumber.organization_id"
    )