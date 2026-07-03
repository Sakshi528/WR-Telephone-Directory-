from functools import wraps

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, flash
)
from flask_login import login_required, current_user
from sqlalchemy import or_

from models import db
from models.employee import Employee
from models.organization import Organization
from models.department import Department
from models.update_request import UpdateRequest
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact


admin_bp = Blueprint("admin", __name__)


# ─── AUTHORIZATION DECORATOR ─────────────────────────────────────────────────

def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not (current_user.is_authenticated and current_user.role == "admin"):
            flash("Admin login required", "warning")
            return redirect(url_for("auth.admin_login"))
        return f(*args, **kwargs)
    return decorated


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def optional_int(value):
    return int(value) if value else None


# ─── DASHBOARD ───────────────────────────────────────────────────────────────

@admin_bp.route("/admin/dashboard")
@admin_required
def dashboard():
    return render_template(
        "admin_dashboard.html",
        total_employees=Employee.query.count(),
        total_organizations=Organization.query.count(),
        total_departments=Department.query.count(),
        pending_requests=UpdateRequest.query.filter_by(status="Pending").count(),
    )


# ─── EMPLOYEES ───────────────────────────────────────────────────────────────

@admin_bp.route("/admin/employees")
@admin_required
def manage_employees():
    keyword = request.args.get("keyword", "")
    if keyword:
        employees = Employee.query.outerjoin(Organization).outerjoin(Department).filter(
            or_(
                Employee.employee_name.ilike(f"%{keyword}%"),
                Employee.designation.ilike(f"%{keyword}%"),
                Employee.email.ilike(f"%{keyword}%"),
                Employee.mobile_phone.ilike(f"%{keyword}%"),
                Organization.organization_name.ilike(f"%{keyword}%"),
                Department.department_name.ilike(f"%{keyword}%"),
            )
        ).order_by(Employee.employee_name).all()
    else:
        employees = Employee.query.order_by(Employee.employee_name).all()

    return render_template("manage_employees.html", employees=employees, keyword=keyword)


@admin_bp.route("/admin/employees/add", methods=["GET", "POST"])
@admin_required
def add_employee():
    organizations = Organization.query.order_by(Organization.organization_name).all()
    departments   = Department.query.order_by(Department.department_name).all()

    if request.method == "POST":
        name   = request.form.get("employee_name", "").strip()
        desig  = request.form.get("designation", "").strip()
        org_id = request.form.get("organization_id", "").strip()
        email  = request.form.get("email", "").strip()
        mobile = request.form.get("mobile_phone", "").strip()

        if not name:
            flash("Employee name is required.", "danger")
            return redirect(request.url)
        if not desig:
            flash("Designation is required.", "danger")
            return redirect(request.url)
        if not org_id:
            flash("Organization is required.", "danger")
            return redirect(request.url)
        if email and "@" not in email:
            flash("Invalid email address.", "danger")
            return redirect(request.url)
        if mobile and not all(c.isdigit() or c in " +-,()" for c in mobile):
            flash("Mobile number contains invalid characters.", "danger")
            return redirect(request.url)

        existing = Employee.query.filter(
            Employee.employee_name.ilike(name),
            Employee.organization_id == int(org_id),
        ).first()
        if existing:
            flash(
                f"An employee named '{name}' already exists in this organization.",
                "danger",
            )
            return redirect(request.url)

        employee = Employee(
            employee_name   = name,
            designation     = desig,
            organization_id = int(org_id),
            department_id   = optional_int(request.form.get("department_id")),
            location        = request.form.get("location"),
            region          = request.form.get("region"),
            office_phone    = request.form.get("office_phone"),
            residence_phone = request.form.get("residence_phone"),
            mobile_phone    = mobile,
            email           = email,
            is_utility_head = bool(request.form.get("is_utility_head")),
        )
        db.session.add(employee)
        db.session.commit()
        flash("Employee added successfully.", "success")
        return redirect(url_for("admin.manage_employees"))

    return render_template(
        "add_employee.html",
        organizations=organizations,
        departments=departments,
    )


@admin_bp.route("/admin/employees/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_employee(id):
    employee      = db.session.get(Employee, id)
    organizations = Organization.query.order_by(Organization.organization_name).all()
    departments   = Department.query.order_by(Department.department_name).all()

    if not employee:
        flash("Employee not found", "danger")
        return redirect(url_for("admin.manage_employees"))

    if request.method == "POST":
        name   = request.form.get("employee_name", "").strip()
        desig  = request.form.get("designation", "").strip()
        org_id = request.form.get("organization_id", "").strip()
        email  = request.form.get("email", "").strip()
        mobile = request.form.get("mobile_phone", "").strip()

        if not name:
            flash("Employee name is required.", "danger")
            return redirect(request.url)
        if not desig:
            flash("Designation is required.", "danger")
            return redirect(request.url)
        if not org_id:
            flash("Organization is required.", "danger")
            return redirect(request.url)
        if email and "@" not in email:
            flash("Invalid email address.", "danger")
            return redirect(request.url)
        if mobile and not all(c.isdigit() or c in " +-,()" for c in mobile):
            flash("Mobile number contains invalid characters.", "danger")
            return redirect(request.url)

        duplicate = Employee.query.filter(
            Employee.employee_name.ilike(name),
            Employee.organization_id == int(org_id),
            Employee.id != id,
        ).first()
        if duplicate:
            flash(
                f"An employee named '{name}' already exists in this organization.",
                "danger",
            )
            return redirect(request.url)

        employee.employee_name   = name
        employee.designation     = desig
        employee.organization_id = int(org_id)
        employee.department_id   = optional_int(request.form.get("department_id"))
        employee.location        = request.form.get("location")
        employee.region          = request.form.get("region")
        employee.office_phone    = request.form.get("office_phone")
        employee.residence_phone = request.form.get("residence_phone")
        employee.mobile_phone    = mobile
        employee.email           = email
        employee.is_utility_head = bool(request.form.get("is_utility_head"))

        # Save emergency contact
        ec_name     = request.form.get("ec_name", "").strip()
        ec_relation = request.form.get("ec_relation", "").strip()
        ec_phone    = request.form.get("ec_phone", "").strip()
        if ec_name or ec_relation or ec_phone:
            ec = EmergencyContact.query.filter_by(employee_id=employee.id).first()
            if not ec:
                ec = EmergencyContact(employee_id=employee.id)
                db.session.add(ec)
            ec.contact_name = ec_name or None
            ec.relation     = ec_relation or None
            ec.phone        = ec_phone or None

        db.session.commit()
        flash("Employee updated successfully.", "success")
        return redirect(url_for("admin.manage_employees"))

    ec = EmergencyContact.query.filter_by(employee_id=employee.id).first()
    return render_template(
        "edit_employee.html",
        employee=employee,
        organizations=organizations,
        departments=departments,
        ec=ec,
    )


@admin_bp.route("/admin/employees/delete/<int:id>", methods=["POST"])
@admin_required
def delete_employee(id):
    employee = db.session.get(Employee, id)
    if employee:
        db.session.delete(employee)
        db.session.commit()
        flash("Employee deleted", "danger")
    return redirect(url_for("admin.manage_employees"))


# ─── ORGANIZATIONS ────────────────────────────────────────────────────────────

@admin_bp.route("/admin/organizations")
@admin_required
def manage_organizations():
    keyword = request.args.get("keyword", "")
    if keyword:
        orgs = Organization.query.filter(
            or_(
                Organization.organization_name.ilike(f"%{keyword}%"),
                Organization.region.ilike(f"%{keyword}%"),
                Organization.address.ilike(f"%{keyword}%"),
            )
        ).order_by(Organization.organization_name).all()
    else:
        orgs = Organization.query.order_by(Organization.organization_name).all()
    return render_template("manage_organizations.html", organizations=orgs, keyword=keyword)


@admin_bp.route("/admin/organizations/add", methods=["GET", "POST"])
@admin_required
def add_organization():
    if request.method == "POST":
        name    = request.form.get("organization_name", "").strip()
        region  = request.form.get("region", "").strip()
        address = request.form.get("address", "").strip()

        if not name:
            flash("Organization name is required.", "danger")
            return redirect(request.url)
        if not region:
            flash("Region is required.", "danger")
            return redirect(request.url)
        if not address:
            flash("Address is required.", "danger")
            return redirect(request.url)
        if Organization.query.filter_by(organization_name=name).first():
            flash("An organization with this name already exists.", "danger")
            return redirect(request.url)

        org = Organization(
            organization_name=name,
            region=region,
            address=address,
        )
        db.session.add(org)
        db.session.commit()
        flash("Organization added successfully.", "success")
        return redirect(url_for("admin.manage_organizations"))
    return render_template("add_organization.html")


@admin_bp.route("/admin/organizations/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_organization(id):
    org = db.session.get(Organization, id)
    if not org:
        flash("Organization not found", "danger")
        return redirect(url_for("admin.manage_organizations"))

    if request.method == "POST":
        name    = request.form.get("organization_name", "").strip()
        region  = request.form.get("region", "").strip()
        address = request.form.get("address", "").strip()

        if not name:
            flash("Organization name is required.", "danger")
            return redirect(request.url)
        if not region:
            flash("Region is required.", "danger")
            return redirect(request.url)
        if not address:
            flash("Address is required.", "danger")
            return redirect(request.url)

        duplicate = Organization.query.filter(
            Organization.organization_name == name,
            Organization.id != id,
        ).first()
        if duplicate:
            flash("An organization with this name already exists.", "danger")
            return redirect(request.url)

        org.organization_name = name
        org.region            = region
        org.address           = address
        db.session.commit()
        flash("Organization updated successfully.", "success")
        return redirect(url_for("admin.manage_organizations"))

    return render_template("edit_organization.html", organization=org)


@admin_bp.route("/admin/organizations/delete/<int:id>", methods=["POST"])
@admin_required
def delete_organization(id):
    org = db.session.get(Organization, id)
    if org:
        db.session.delete(org)
        db.session.commit()
        flash("Organization deleted", "danger")
    return redirect(url_for("admin.manage_organizations"))


# ─── EMERGENCY CONTACTS ───────────────────────────────────────────────────────

@admin_bp.route("/admin/emergency-contacts")
@admin_required
def manage_emergency_contacts():
    keyword = request.args.get("keyword", "")
    if keyword:
        contacts = (
            EmergencyContact.query
            .join(Employee)
            .filter(
                or_(
                    Employee.employee_name.ilike(f"%{keyword}%"),
                    EmergencyContact.contact_name.ilike(f"%{keyword}%"),
                    EmergencyContact.relation.ilike(f"%{keyword}%"),
                    EmergencyContact.phone.ilike(f"%{keyword}%"),
                )
            )
            .order_by(Employee.employee_name)
            .all()
        )
    else:
        contacts = (
            EmergencyContact.query
            .join(Employee)
            .order_by(Employee.employee_name)
            .all()
        )
    return render_template(
        "manage_emergency_contacts.html", contacts=contacts, keyword=keyword
    )


@admin_bp.route("/admin/emergency-contacts/add", methods=["GET", "POST"])
@admin_required
def add_emergency_contact():
    employees = Employee.query.order_by(Employee.employee_name).all()
    if request.method == "POST":
        contact = EmergencyContact(
            employee_id  = int(request.form["employee_id"]),
            contact_name = request.form["contact_name"],
            relation     = request.form.get("relation"),
            phone        = request.form.get("phone"),
        )
        db.session.add(contact)
        db.session.commit()
        flash("Emergency contact added", "success")
        return redirect(url_for("admin.manage_emergency_contacts"))
    return render_template("add_emergency_contact.html", employees=employees)


@admin_bp.route("/admin/emergency-contacts/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_emergency_contact(id):
    contact   = db.session.get(EmergencyContact, id)
    employees = Employee.query.order_by(Employee.employee_name).all()

    if not contact:
        flash("Contact not found", "danger")
        return redirect(url_for("admin.manage_emergency_contacts"))

    if request.method == "POST":
        contact.employee_id  = int(request.form["employee_id"])
        contact.contact_name = request.form["contact_name"]
        contact.relation     = request.form.get("relation")
        contact.phone        = request.form.get("phone")
        db.session.commit()
        flash("Emergency contact updated", "success")
        return redirect(url_for("admin.manage_emergency_contacts"))

    return render_template(
        "edit_emergency_contact.html", contact=contact, employees=employees
    )


@admin_bp.route("/admin/emergency-contacts/delete/<int:id>", methods=["POST"])
@admin_required
def delete_emergency_contact(id):
    contact = db.session.get(EmergencyContact, id)
    if contact:
        db.session.delete(contact)
        db.session.commit()
        flash("Emergency contact deleted", "danger")
    return redirect(url_for("admin.manage_emergency_contacts"))


# ─── UPDATE REQUESTS ─────────────────────────────────────────────────────────

@admin_bp.route("/admin/requests")
@admin_required
def manage_requests():
    reqs = UpdateRequest.query.order_by(
        UpdateRequest.request_date.desc()
    ).all()
    return render_template("manage_requests.html", requests=reqs)


@admin_bp.route("/admin/requests/approve/<int:id>")
@admin_required
def approve_request(id):
    update_request = db.session.get(UpdateRequest, id)
    if not update_request:
        flash("Request not found", "danger")
        return redirect(url_for("admin.manage_requests"))

    if update_request.request_type == "directory":
        directory_number = db.session.get(DirectoryNumber, update_request.directory_number_id)
        if not directory_number:
            flash("Directory number not found", "danger")
            return redirect(url_for("admin.manage_requests"))
        # New schema: field_name / new_value
        if update_request.field_name and update_request.new_value:
            field = update_request.field_name.lower().replace(" ", "_")
            if field == "phone_number":
                directory_number.phone_number = update_request.new_value
            elif field == "email":
                directory_number.email = update_request.new_value
        # Legacy schema
        elif update_request.requested_directory_number:
            directory_number.phone_number = update_request.requested_directory_number
    else:
        employee = db.session.get(Employee, update_request.employee_id)
        if not employee:
            flash("Employee not found", "danger")
            return redirect(url_for("admin.manage_requests"))
        # New schema: field_name / new_value
        if update_request.field_name and update_request.new_value:
            field = update_request.field_name.lower().replace(" ", "_")
            field_map = {
                "mobile": "mobile_phone",
                "mobile_phone": "mobile_phone",
                "office_phone": "office_phone",
                "office": "office_phone",
                "email": "email",
                "residence_phone": "residence_phone",
                "residence": "residence_phone",
                "designation": "designation",
                "address": "address",
            }
            attr = field_map.get(field)
            if attr:
                setattr(employee, attr, update_request.new_value)
            elif field in ("emergency_contact_name", "emergency_contact_relation",
                           "emergency_contact_phone"):
                ec = EmergencyContact.query.filter_by(
                    employee_id=employee.id
                ).first()
                if not ec:
                    ec = EmergencyContact(employee_id=employee.id)
                    db.session.add(ec)
                if field == "emergency_contact_name":
                    ec.contact_name = update_request.new_value
                elif field == "emergency_contact_relation":
                    ec.relation = update_request.new_value
                elif field == "emergency_contact_phone":
                    ec.phone = update_request.new_value
        # Legacy schema
        else:
            if update_request.requested_mobile:
                employee.mobile_phone = update_request.requested_mobile
            if update_request.requested_office_phone:
                employee.office_phone = update_request.requested_office_phone
            if update_request.requested_email:
                employee.email = update_request.requested_email

    update_request.status = "Approved"
    db.session.commit()
    flash("Request approved", "success")
    return redirect(url_for("admin.manage_requests"))


@admin_bp.route("/admin/requests/reject/<int:id>")
@admin_required
def reject_request(id):
    update_request = db.session.get(UpdateRequest, id)
    if update_request:
        update_request.status = "Rejected"
        db.session.commit()
    flash("Request rejected", "warning")
    return redirect(url_for("admin.manage_requests"))


# ─── UTILITY HEAD TOGGLE ─────────────────────────────────────────────────────

@admin_bp.route("/admin/employees/<int:id>/toggle-utility-head", methods=["POST"])
@admin_required
def toggle_utility_head(id):
    emp = db.session.get(Employee, id)
    if emp:
        emp.is_utility_head = not emp.is_utility_head
        db.session.commit()
        state = "marked as" if emp.is_utility_head else "removed from"
        flash(f"{emp.employee_name} {state} Utility Head.", "success")
    return redirect(request.referrer or url_for("user.utility_heads"))


# ─── DIRECTORY NUMBERS ────────────────────────────────────────────────────────

@admin_bp.route("/admin/directory-numbers")
@admin_required
def directory_numbers():
    keyword = request.args.get("keyword", "").strip()
    q = DirectoryNumber.query
    if keyword:
        like = f"%{keyword}%"
        q = q.filter(
            or_(
                DirectoryNumber.name.ilike(like),
                DirectoryNumber.phone_number.ilike(like),
                DirectoryNumber.organization.ilike(like),
                DirectoryNumber.category.ilike(like),
            )
        )
    numbers = q.order_by(DirectoryNumber.organization, DirectoryNumber.name).all()
    return render_template("directory_numbers.html", numbers=numbers, keyword=keyword)


@admin_bp.route("/admin/directory-numbers/add", methods=["GET", "POST"])
@admin_required
def add_directory_number():
    if request.method == "POST":
        number = DirectoryNumber(
            name         = request.form["name"],
            phone_number = request.form.get("phone_number"),
            email        = request.form.get("email"),
            organization = request.form.get("organization"),
            category     = request.form.get("category"),
        )
        db.session.add(number)
        db.session.commit()
        flash("Directory number added", "success")
        return redirect(url_for("admin.directory_numbers"))
    return render_template("add_directory_number.html")


@admin_bp.route("/admin/directory-numbers/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_directory_number(id):
    number = db.session.get(DirectoryNumber, id)
    if not number:
        flash("Directory number not found", "danger")
        return redirect(url_for("admin.directory_numbers"))

    if request.method == "POST":
        number.name         = request.form["name"]
        number.phone_number = request.form.get("phone_number")
        number.email        = request.form.get("email")
        number.organization = request.form.get("organization")
        number.category     = request.form.get("category")
        db.session.commit()
        flash("Directory number updated", "success")
        return redirect(url_for("admin.directory_numbers"))

    return render_template("edit_directory_number.html", number=number)


@admin_bp.route("/admin/directory-numbers/delete/<int:id>", methods=["POST"])
@admin_required
def delete_directory_number(id):
    number = db.session.get(DirectoryNumber, id)
    if number:
        db.session.delete(number)
        db.session.commit()
        flash("Directory number deleted", "danger")
    return redirect(url_for("admin.directory_numbers"))
