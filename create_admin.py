from werkzeug.security import generate_password_hash

from app import app
from models import db
from models.user import User


with app.app_context():

    password_hash = generate_password_hash(
        "admin123"
    )

    existing_admin = User.query.filter_by(
        email="admin@gmail.com"
    ).first()

    if existing_admin:
        existing_admin.username = "admin"
        existing_admin.password = password_hash
        existing_admin.role = "admin"
        existing_admin.status = "active"

        db.session.commit()

        print("Admin reset successfully")

    else:

        admin = User(
            username="admin",
            email="admin@gmail.com",
            password=password_hash,
            role="admin"
        )

        db.session.add(admin)

        db.session.commit()

        print("Admin Created Successfully")
