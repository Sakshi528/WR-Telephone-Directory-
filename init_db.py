from app import app
from models import db
from models.user import User  # noqa: F401
from models.employee import Employee  # noqa: F401
from models.organization import Organization  # noqa: F401
from models.department import Department  # noqa: F401
from models.directory_number import DirectoryNumber  # noqa: F401
from models.emergency_contact import EmergencyContact  # noqa: F401
from models.update_request import UpdateRequest  # noqa: F401

with app.app_context():
    db.create_all()
    print("Tables created")
