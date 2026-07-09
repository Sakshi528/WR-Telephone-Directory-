from flask import Flask

from flask_login import LoginManager

from config import Config
from models import db
from models.user import User
from models.emergency_contact import EmergencyContact  # noqa: F401 — registers model

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


@app.route("/count")
def count():
    from models.employee import Employee
    from models.organization import Organization
    from models.department import Department
    from models.directory_number import DirectoryNumber

    return (
        f"Employees: {Employee.query.count()}<br>"
        f"Organizations: {Organization.query.count()}<br>"
        f"Departments: {Department.query.count()}<br>"
        f"Directory Numbers: {DirectoryNumber.query.count()}"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
