from flask_login import UserMixin

from models import db


class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    username = db.Column(
        db.String(100),
        nullable=False
    )

    email = db.Column(
        db.String(255),
        unique=True,
        nullable=False
    )

    password = db.Column(
        db.String(255),
        nullable=False
    )

    role = db.Column(
        db.String(20),
        default="user"
    )

    status = db.Column(
        db.String(20),
        default="active"
    )

    created_at = db.Column(
        db.DateTime,
        server_default=db.func.now()
    )
    

    def is_admin(self):
        return self.role == "admin"
    
    update_requests = db.relationship(
        "UpdateRequest",
        back_populates="user",
        foreign_keys="UpdateRequest.user_id",
        lazy=True
    )
