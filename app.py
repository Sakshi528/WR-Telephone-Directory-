import os

from flask import Flask

from flask_login import LoginManager

from config import Config
from models import db
from models.user import User
from models.emergency_contact import EmergencyContact  # noqa: F401 — registers model
from models.employee_status_history import EmployeeStatusHistory  # noqa: F401 — registers model
from models.administrative_head import AdministrativeHead  # noqa: F401 — registers model
from models.administrative_head_history import AdministrativeHeadHistory  # noqa: F401 — registers model
from models.administrative_head_assistant import AdministrativeHeadAssistant  # noqa: F401 — registers model
from models.service_type import ServiceType  # noqa: F401 — registers model
from models.email_group import EmailGroup, GroupMember  # noqa: F401 — registers model
from models.email_group_filter import EmailGroupFilter  # noqa: F401 — registers model
from models.organization_category import OrganizationCategory  # noqa: F401 — registers model
from models.audit_log import AuditLog  # noqa: F401 — registers model
from models.import_batch import ImportBatch  # noqa: F401 — registers model
from models.directory_version import DirectoryVersion  # noqa: F401 — registers model

from routes.auth_routes import auth_bp
from routes.user_routes import user_bp
from routes.admin_routes import admin_bp


app = Flask(__name__)
app.config.from_object(Config)
app.config["SESSION_PERMANENT"] = True

db.init_app(app)

login_manager = LoginManager()
login_manager.login_view = "auth.admin_login"
login_manager.login_message = ""
login_manager.init_app(app)


@login_manager.unauthorized_handler
def unauthorized():
    from flask import redirect
    return redirect("/admin-login")


@app.errorhandler(401)
def handle_401(_e):
    from flask import redirect
    return redirect("/admin-login")


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


app.register_blueprint(auth_bp)
app.register_blueprint(user_bp)
app.register_blueprint(admin_bp)


@app.context_processor
def inject_utility_head_ids():
    from utils.designation_rank import compute_utility_head_ids
    return {"utility_head_ids": compute_utility_head_ids()}


if __name__ == "__main__":
    # debug=True exposes Werkzeug's interactive debugger console over the
    # network -- anyone who can reach an unhandled exception gets a Python
    # REPL on the server. Default False (LAN-safe); opt in locally with
    # FLASK_DEBUG=1 for a dev machine that isn't reachable by anyone else.
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=8000, debug=debug)
