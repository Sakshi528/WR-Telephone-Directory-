from flask import Blueprint, render_template, request, redirect, url_for, flash

from flask_login import login_user, logout_user, login_required, current_user

from werkzeug.security import check_password_hash

from models.user import User


auth_bp = Blueprint("auth", __name__)


@auth_bp.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    if current_user.is_authenticated:
        if current_user.role == "admin":
            return redirect(url_for("admin.dashboard"))
        logout_user()
        flash("Please login with an admin account", "warning")

    if request.method == "POST":
        email    = request.form.get("email")
        password = request.form.get("password")

        admin = User.query.filter_by(email=email, role="admin").first()

        if admin and check_password_hash(admin.password, password):
            if admin.status != "active":
                flash("Admin account is inactive", "danger")
                return redirect(url_for("auth.admin_login"))
            login_user(admin)
            return redirect(url_for("admin.dashboard"))

        flash("Invalid Admin Credentials", "danger")

    return render_template("admin_login.html")


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("user.home"))
