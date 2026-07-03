from models import db


class EmergencyContact(db.Model):
    __tablename__ = "emergency_contacts"

    id = db.Column(db.Integer, primary_key=True)

    employee_id = db.Column(
        db.Integer,
        db.ForeignKey("employees.id"),
        nullable=False
    )

    contact_name = db.Column(db.String(255), nullable=False)

    relation = db.Column(db.String(100))

    phone = db.Column(db.String(50))

    employee = db.relationship("Employee", lazy=True)
