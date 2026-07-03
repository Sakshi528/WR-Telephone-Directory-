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