from models import db


class ServiceType(db.Model):
    """Normalized civil-service cadre for an Administrative Head (IAS, IPS,
    IFS, State Civil Service, etc.) -- lets Email Distribution Lists and the
    Administrative Heads filters work without free-text matching."""

    __tablename__ = "service_types"

    id = db.Column(db.Integer, primary_key=True)

    service_type_name = db.Column(db.String(100), nullable=False, unique=True)

    description = db.Column(db.Text)
