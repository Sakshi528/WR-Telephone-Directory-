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
