import io
from datetime import datetime
from functools import wraps

from flask import (
    Blueprint, render_template, request,
    redirect, url_for, flash, jsonify, send_file
)
from flask_login import login_required, current_user
from sqlalchemy import or_, func

from models import db
from models.employee import Employee, EMPLOYEE_STATUSES
from models.organization import Organization
from models.organization_category import OrganizationCategory
from models.department import Department
from models.update_request import UpdateRequest
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from models.employee_status_history import EmployeeStatusHistory
from models.email_group import EmailGroup
from models.email_group_filter import EmailGroupFilter
from models.administrative_head import AdministrativeHead
from models.administrative_head_assistant import AdministrativeHeadAssistant
from models.service_type import ServiceType
from models.directory_version import DirectoryVersion
from models.import_batch import ImportBatch
from utils.designation_rank import resolve_utility_head
from services.audit_service import log_audit_event
from services.email_distribution_service import resolve_dynamic_group, dedupe_and_sort, to_contact_row
from services import head_service
from services import directory_version_service
from services.verification_workbook_service import (
    generate_verification_workbook, WORKBOOK_FILENAME,
)

def _default_category_id():
    """organizations.category_id is NOT NULL -- any organization left
    'Uncategorized' on the add/edit form falls back to the seeded 'Others'
    category rather than failing the insert."""
    other = OrganizationCategory.query.filter_by(category_name="Others").first()
    return other.id if other else None


ROLE_CHOICES = [
    ("UTILITY_HEAD", "Utility Head"),
    ("ADMINISTRATIVE_HEAD", "Administrative Head"),
    ("KMP", "KMP"),
]
STATUS_CHOICES = ["ACTIVE", "INACTIVE", "TRANSFERRED", "RETIRED", "DEPUTATION", "RESIGNED"]


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


def organization_type_mismatch(organization_id, organization_type_id):
    """True if the submitted Organization Type doesn't match the selected
    Organization's actual category -- guards against a tampered/stale
    cascading-dropdown selection (JS disabled, DOM edited, cached page).
    The JSON endpoint behind the dropdown only ever returns organizations
    that already match the requested type, so a real mismatch here only
    happens on a bypassed client."""
    if not organization_type_id or not organization_id:
        return False
    org = db.session.get(Organization, organization_id)
    return bool(org) and org.category_id != organization_type_id


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


@admin_bp.route("/admin/employees/suggest")
@admin_required
def employees_suggest():
    """Id-returning employee picker, used by static/js/autocomplete.js in
    opt-in (data-target-id-field) mode -- /suggestions and
    /utility-heads/suggest only return name strings, not ids, so neither
    can back a real employee_id foreign key picker (e.g. Administrative
    Head / PA-PS assignment forms)."""
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify([])

    employees = (
        Employee.query.outerjoin(Organization)
        .filter(Employee.employee_name.ilike(f"%{q}%"))
        .order_by(Employee.employee_name)
        .limit(10)
        .all()
    )
    return jsonify([
        {
            "label": e.employee_name,
            "value": e.id,
            "category": e.organization.organization_name if e.organization else (e.designation or ""),
        }
        for e in employees
    ])


@admin_bp.route("/admin/employees/add", methods=["GET", "POST"])
@admin_required
def add_employee():
    categories  = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    departments = Department.query.order_by(Department.department_name).all()

    if request.method == "POST":
        name   = request.form.get("employee_name", "").strip()
        desig  = request.form.get("designation", "").strip()
        org_id = request.form.get("organization_id", "").strip()
        org_type_id = optional_int(request.form.get("organization_type_id"))
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
        if organization_type_mismatch(optional_int(org_id), org_type_id):
            flash("Selected organization does not belong to the selected organization type.", "danger")
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
            is_kmp          = bool(request.form.get("is_kmp")),
        )
        db.session.add(employee)
        db.session.commit()
        flash("Employee added successfully.", "success")
        if request.form.get("next") == "directory":
            return redirect(url_for("user.directory"))
        return redirect(url_for("admin.manage_employees"))

    return render_template(
        "add_employee.html",
        categories=categories,
        departments=departments,
    )


@admin_bp.route("/admin/employees/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_employee(id):
    employee    = db.session.get(Employee, id)
    categories  = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    departments = Department.query.order_by(Department.department_name).all()

    if not employee:
        flash("Employee not found", "danger")
        return redirect(url_for("admin.manage_employees"))

    if request.method == "POST":
        name   = request.form.get("employee_name", "").strip()
        desig  = request.form.get("designation", "").strip()
        org_id = request.form.get("organization_id", "").strip()
        org_type_id = optional_int(request.form.get("organization_type_id"))
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
        if organization_type_mismatch(optional_int(org_id), org_type_id):
            flash("Selected organization does not belong to the selected organization type.", "danger")
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

        old_organization_id = employee.organization_id
        old_designation = employee.designation

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
        employee.is_kmp          = bool(request.form.get("is_kmp"))

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

        # Organization/designation changes feed the Archive profile's
        # timeline the same way status changes already do -- see
        # models/employee_status_history.py
        org_changed = old_organization_id != employee.organization_id
        desig_changed = old_designation != desig
        if org_changed or desig_changed:
            db.session.add(EmployeeStatusHistory(
                employee_id=employee.id,
                old_status=employee.status,
                new_status=employee.status,
                old_organization_id=old_organization_id if org_changed else None,
                new_organization_id=employee.organization_id if org_changed else None,
                old_designation=old_designation if desig_changed else None,
                new_designation=desig if desig_changed else None,
                changed_by=current_user.id if current_user.is_authenticated else None,
            ))

        db.session.commit()
        flash("Employee updated successfully.", "success")
        return redirect(url_for("admin.manage_employees"))

    ec = EmergencyContact.query.filter_by(employee_id=employee.id).first()
    return render_template(
        "edit_employee.html",
        employee=employee,
        categories=categories,
        departments=departments,
        ec=ec,
    )


@admin_bp.route("/admin/employees/delete/<int:id>", methods=["POST"])
@admin_required
def delete_employee(id):
    employee = db.session.get(Employee, id)
    if employee:
        log_audit_event(
            module="employees", record_type="Employee", record_id=employee.id,
            action="DELETE", field_name="employee_name",
            old_value=employee.employee_name, new_value=None,
        )
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

    status_rows = (
        db.session.query(Employee.organization_id, Employee.status, func.count(Employee.id))
        .group_by(Employee.organization_id, Employee.status)
        .all()
    )
    status_by_org = {}
    for org_id, status, count in status_rows:
        status_by_org.setdefault(org_id, {})[status] = count

    return render_template(
        "manage_organizations.html", organizations=orgs, keyword=keyword,
        status_by_org=status_by_org,
    )


@admin_bp.route("/admin/organizations/add", methods=["GET", "POST"])
@admin_required
def add_organization():
    if request.method == "POST":
        name        = request.form.get("organization_name", "").strip()
        region      = request.form.get("region", "").strip()
        address     = request.form.get("address", "").strip()
        category_id = optional_int(request.form.get("category_id")) or _default_category_id()

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
            category_id=category_id,
        )
        db.session.add(org)
        db.session.flush()
        if category_id:
            log_audit_event(
                module="organizations", record_type="Organization", record_id=org.id,
                action="CATEGORY_CHANGE", field_name="category_id",
                old_value=None, new_value=str(category_id),
            )
        db.session.commit()
        flash("Organization added successfully.", "success")
        return redirect(url_for("admin.manage_organizations"))
    categories = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    return render_template("add_organization.html", categories=categories)


@admin_bp.route("/admin/organizations/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_organization(id):
    org = db.session.get(Organization, id)
    if not org:
        flash("Organization not found", "danger")
        return redirect(url_for("admin.manage_organizations"))

    if request.method == "POST":
        name        = request.form.get("organization_name", "").strip()
        region      = request.form.get("region", "").strip()
        address     = request.form.get("address", "").strip()
        category_id = optional_int(request.form.get("category_id")) or _default_category_id()

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

        if category_id != org.category_id:
            log_audit_event(
                module="organizations", record_type="Organization", record_id=org.id,
                action="CATEGORY_CHANGE", field_name="category_id",
                old_value=str(org.category_id) if org.category_id else None,
                new_value=str(category_id) if category_id else None,
            )

        org.organization_name = name
        org.region            = region
        org.address           = address
        org.category_id       = category_id
        db.session.commit()
        flash("Organization updated successfully.", "success")
        return redirect(url_for("admin.manage_organizations"))

    categories = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    return render_template("edit_organization.html", organization=org, categories=categories)


@admin_bp.route("/admin/organizations/delete/<int:id>", methods=["POST"])
@admin_required
def delete_organization(id):
    org = db.session.get(Organization, id)
    if org:
        # Deleting an organization that still has dependents doesn't fail --
        # the DB's ON DELETE rules silently orphan employees/directory
        # numbers/child organizations (organization_id/parent_id -> NULL)
        # and CASCADE-delete any Administrative Head tied to it. Blocking
        # here forces a deliberate reassignment instead of an invisible
        # data loss with no confirmation and no audit trail.
        emp_count = Employee.query.filter_by(organization_id=org.id).count()
        dn_count = DirectoryNumber.query.filter_by(organization_id=org.id).count()
        child_count = Organization.query.filter_by(parent_id=org.id).count()
        ah_count = AdministrativeHead.query.filter_by(organization_id=org.id).count()

        if emp_count or dn_count or child_count or ah_count:
            parts = []
            if emp_count:
                parts.append(f"{emp_count} employee(s)")
            if dn_count:
                parts.append(f"{dn_count} control room/switchyard record(s)")
            if child_count:
                parts.append(f"{child_count} sub-organization(s)")
            if ah_count:
                parts.append(f"{ah_count} administrative head(s)")
            flash(
                "Cannot delete '" + org.organization_name + "' -- it still has " +
                ", ".join(parts) + ". Reassign or remove those first.",
                "danger",
            )
            return redirect(url_for("admin.manage_organizations"))

        log_audit_event(
            module="organizations", record_type="Organization", record_id=org.id,
            action="DELETE", field_name="organization_name",
            old_value=org.organization_name, new_value=None,
        )
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

    if update_request.request_type == "administrative_head":
        head = db.session.get(AdministrativeHead, update_request.administrative_head_id)
        if not head:
            flash("Administrative head not found", "danger")
            return redirect(url_for("admin.manage_requests"))
        # Contact fields are resolved from the linked Employee record when
        # one exists (see AdministrativeHead.resolved_office_phone etc.) --
        # writing to the head's own standalone field in that case would be
        # silently ignored everywhere the resolved_* value is displayed, so
        # route those three fields to the linked employee instead.
        contact_fields = {"office_phone", "office", "mobile_phone", "mobile", "email"}
        field_map = {
            "role_title": "role_title", "designation": "role_title",
            "office_phone": "office_phone", "office": "office_phone",
            "mobile_phone": "mobile_phone", "mobile": "mobile_phone",
            "email": "email",
            "office_address": "office_address", "address": "office_address",
        }
        if update_request.field_name and update_request.new_value:
            key = update_request.field_name.lower().replace(" ", "_")
            attr = field_map.get(key)
            if attr:
                target = head.employee if (key in contact_fields and head.employee_id and head.employee) else head
                setattr(target, attr, update_request.new_value)
    elif update_request.request_type == "directory":
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
            if field == "status":
                # Not a plain attribute set -- status changes need the same
                # history/audit/auto-archive side effects as the admin's
                # direct status dropdown (see _apply_employee_status_change).
                ok, message = _apply_employee_status_change(
                    employee, update_request.new_value.strip().upper(),
                    reason=f"Requested by {update_request.requested_by or 'anonymous'}"
                    + (f": {update_request.reason}" if update_request.reason else ""),
                )
                if not ok:
                    flash(message, "danger")
                    return redirect(url_for("admin.manage_requests"))
            elif attr:
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
        org = db.session.get(Organization, emp.organization_id) if emp.organization_id else None

        if emp.is_utility_head:
            if org is not None:
                # Only one manually-designated head per organization -- clear
                # any other employee's flag in the same org so resolution
                # (utils.designation_rank.resolve_utility_head) stays
                # unambiguous, and un-exclude the org since this is an
                # explicit choice of head.
                Employee.query.filter(
                    Employee.organization_id == emp.organization_id,
                    Employee.id != emp.id,
                ).update({"is_utility_head": False})
                org.utility_head_excluded = False
        elif org is not None:
            # "Remove as Utility Head" was clicked. If nobody else in the
            # organization would take over -- e.g. a single-employee org, or
            # this employee was also the most senior by designation -- auto-
            # resolution would just re-pick this same person, making the
            # removal a no-op. In that case, explicitly mark the
            # organization as having no Utility Head at all.
            siblings = (
                Employee.query
                .filter_by(organization_id=emp.organization_id, status="ACTIVE")
                .order_by(Employee.id)
                .all()
            )
            new_head = resolve_utility_head(siblings)
            org.utility_head_excluded = new_head is None or new_head.id == emp.id

        db.session.commit()
        state = "marked as" if emp.is_utility_head else "removed from"
        flash(f"{emp.employee_name} {state} Utility Head.", "success")
    return redirect(request.referrer or url_for("user.utility_heads"))


@admin_bp.route("/admin/employees/<int:id>/toggle-kmp", methods=["POST"])
@admin_required
def toggle_kmp(id):
    emp = db.session.get(Employee, id)
    if emp:
        old_value = emp.is_kmp
        emp.is_kmp = not emp.is_kmp
        log_audit_event(
            module="employees", record_type="Employee", record_id=emp.id,
            action="UPDATE", field_name="is_kmp",
            old_value=str(old_value), new_value=str(emp.is_kmp),
        )
        db.session.commit()
        state = "added to" if emp.is_kmp else "removed from"
        flash(f"{emp.employee_name} {state} the Telephone Directory (KMP).", "success")
    return redirect(request.referrer or url_for("admin.manage_employees"))


# ─── EMPLOYEE STATUS ──────────────────────────────────────────────────────────
# EMPLOYEE_STATUSES now lives in models/employee.py -- shared with
# user_routes.py's public "Status" update-request field, not duplicated.


def _apply_employee_status_change(employee, new_status, reason=None):
    """Core status-change logic, shared by the admin's direct status dropdown
    (change_employee_status below) and an approved "Status" Update Request
    (approve_request) -- both need the same EmployeeStatusHistory logging,
    audit entry, and Administrative Head auto-archival, not just a raw
    employee.status assignment. Returns (ok, message)."""
    if new_status not in EMPLOYEE_STATUSES:
        return False, "Invalid status."

    old_status = employee.status
    if new_status == old_status:
        return False, "Status unchanged."

    employee.status = new_status
    employee.status_changed_at = datetime.utcnow()

    db.session.add(EmployeeStatusHistory(
        employee_id=employee.id,
        old_status=old_status,
        new_status=new_status,
        reason=reason or None,
        changed_by=current_user.id if current_user.is_authenticated else None,
    ))
    log_audit_event(
        module="employee_status", record_type="Employee", record_id=employee.id,
        action="STATUS_CHANGE", field_name="status",
        old_value=old_status, new_value=new_status, reason=reason or None,
    )

    # Auto-archive: a role doesn't stay "current" once the person holding it
    # is no longer an active employee. Reuses the exact function the manual
    # "End Role" button already calls -- same code path, not a new one --
    # so AdministrativeHeadHistory (the existing archive for ended roles)
    # stays complete without requiring a separate manual step.
    if new_status != "ACTIVE":
        for head in AdministrativeHead.query.filter_by(employee_id=employee.id).all():
            head_service.end_role(
                head, reason=f"Employee status changed to {new_status}",
            )

    return True, f"{employee.employee_name}'s status changed to {new_status}."


@admin_bp.route("/admin/employees/<int:id>/status", methods=["POST"])
@admin_required
def change_employee_status(id):
    employee = db.session.get(Employee, id)
    if not employee:
        flash("Employee not found", "danger")
        return redirect(url_for("admin.manage_employees"))

    new_status = request.form.get("new_status", "").strip()
    reason = request.form.get("reason", "").strip()

    ok, message = _apply_employee_status_change(employee, new_status, reason)
    if ok:
        db.session.commit()
        flash(message, "success")
    else:
        db.session.rollback()
        flash(message, "warning" if message == "Status unchanged." else "danger")
    return redirect(request.referrer or url_for("user.employee_details", id=employee.id))


@admin_bp.route("/admin/employees/<int:id>/status-history")
@admin_required
def employee_status_history(id):
    employee = db.session.get(Employee, id)
    if not employee:
        flash("Employee not found", "danger")
        return redirect(url_for("admin.manage_employees"))

    history = (
        EmployeeStatusHistory.query
        .filter_by(employee_id=id)
        .order_by(EmployeeStatusHistory.changed_at.desc())
        .all()
    )

    return render_template("employee_status_history.html", employee=employee, history=history)


# ─── DIRECTORY VERSIONS (ADMIN) ────────────────────────────────────────────────

@admin_bp.route("/admin/directory-versions/generate", methods=["POST"])
@admin_required
def generate_directory_version():
    month = request.form.get("month", type=int)
    year = request.form.get("year", type=int)
    remarks = request.form.get("remarks", "").strip()

    if not month or not year:
        flash("Month and Year are required.", "danger")
        return redirect(url_for("user.directory_versions_list"))

    try:
        version = directory_version_service.generate_version(month, year, remarks, current_user)
        db.session.commit()
        flash(f"Directory Version {version.version_number} generated.", "success")
    except ValueError as e:
        db.session.rollback()
        flash(str(e), "danger")
    return redirect(url_for("user.directory_versions_list"))


@admin_bp.route("/admin/directory-versions/<int:id>/regenerate", methods=["POST"])
@admin_required
def regenerate_directory_version(id):
    version = db.session.get(DirectoryVersion, id)
    if not version:
        flash("Directory Version not found.", "danger")
        return redirect(url_for("user.directory_versions_list"))

    remarks = request.form.get("remarks", version.remarks or "").strip()
    new_version = directory_version_service.regenerate_version(
        version.year, version.month, remarks, current_user,
    )
    db.session.commit()
    flash(f"Directory Version {new_version.version_number} generated (regeneration).", "success")
    return redirect(url_for("user.directory_versions_list"))


@admin_bp.route("/admin/directory-versions/<int:id>/delete", methods=["POST"])
@admin_required
def delete_directory_version(id):
    version = db.session.get(DirectoryVersion, id)
    if not version:
        flash("Directory Version not found.", "danger")
        return redirect(url_for("user.directory_versions_list"))

    version_number = version.version_number
    directory_version_service.delete_version(version, current_user)
    db.session.commit()
    flash(f"Directory Version {version_number} deleted.", "success")
    return redirect(url_for("user.directory_versions_list"))


@admin_bp.route("/admin/directory-versions/<int:id>/remarks", methods=["POST"])
@admin_required
def update_directory_version_remarks(id):
    version = db.session.get(DirectoryVersion, id)
    if not version:
        flash("Directory Version not found.", "danger")
        return redirect(url_for("user.directory_versions_list"))

    remarks = request.form.get("remarks", "").strip()
    directory_version_service.update_remarks(version, remarks, current_user)
    db.session.commit()
    flash("Remarks updated.", "success")
    return redirect(url_for("user.directory_versions_list"))


# ─── IMPORT HISTORY (ADMIN ONLY) ──────────────────────────────────────────────
# Internal administration tool -- not part of the public UI. Model,
# logging, and schema are unchanged; only the route path and access
# control moved here from routes/user_routes.py.

@admin_bp.route("/admin/import-history")
@admin_required
def import_history():
    page = request.args.get("page", 1, type=int)
    per_page = 50
    query = ImportBatch.query.order_by(ImportBatch.import_date.desc())
    total = query.count()
    batches = query.offset((page - 1) * per_page).limit(per_page).all()
    return render_template(
        "import_history.html", batches=batches, total=total, page=page, per_page=per_page,
    )


# ─── EXPORT VERIFICATION WORKBOOK (ADMIN ONLY) ────────────────────────────────
# Generated entirely from live PostgreSQL data (never the source import
# workbook) so another team can verify exactly what the application
# currently displays before LAN deployment -- see
# services/verification_workbook_service.py for the per-sheet queries.

@admin_bp.route("/admin/export-verification-workbook")
@admin_required
def export_verification_workbook():
    workbook_bytes = generate_verification_workbook(generated_by=current_user.username)
    buf = io.BytesIO(workbook_bytes)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=WORKBOOK_FILENAME,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
        org_id = request.form.get("organization_id", type=int)
        org = db.session.get(Organization, org_id) if org_id else None
        number = DirectoryNumber(
            name            = request.form["name"],
            phone_number    = request.form.get("phone_number"),
            email           = request.form.get("email"),
            organization    = org.organization_name if org else None,
            organization_id = org.id if org else None,
            category        = request.form.get("category"),
            switch_yard     = "switch_yard" in request.form,
            control_room    = "control_room" in request.form,
            ip_address      = request.form.get("ip_address") or None,
        )
        db.session.add(number)
        db.session.commit()
        flash("Directory number added", "success")
        return redirect(url_for("admin.directory_numbers"))

    # Dashboard "Add Control Room" / "Add Switch Yard" shortcuts land here
    # with ?type=control_room|switch_yard to pre-fill the Category field.
    prefill_type = request.args.get("type")
    prefill = {
        "category": "Control Room" if prefill_type == "control_room"
                    else "Switchyard" if prefill_type == "switch_yard"
                    else "",
    }
    organizations = Organization.query.order_by(Organization.organization_name).all()
    return render_template("add_directory_number.html", prefill=prefill, organizations=organizations)


@admin_bp.route("/admin/directory-numbers/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_directory_number(id):
    number = db.session.get(DirectoryNumber, id)
    if not number:
        flash("Directory number not found", "danger")
        return redirect(url_for("admin.directory_numbers"))

    if request.method == "POST":
        org_id = request.form.get("organization_id", type=int)
        org = db.session.get(Organization, org_id) if org_id else None
        number.name            = request.form["name"]
        number.phone_number    = request.form.get("phone_number")
        number.email           = request.form.get("email")
        number.organization    = org.organization_name if org else None
        number.organization_id = org.id if org else None
        number.category        = request.form.get("category")
        # ip_address/switch_yard/control_room have no inputs on this form
        # today -- only touch them if a future form actually submits
        # "ip_address", so this route can't silently wipe values some
        # other form (or a direct DB edit) set.
        if "ip_address" in request.form:
            number.ip_address   = request.form.get("ip_address") or None
            number.switch_yard  = "switch_yard" in request.form
            number.control_room = "control_room" in request.form
        db.session.commit()
        flash("Directory number updated", "success")
        return redirect(url_for("admin.directory_numbers"))

    organizations = Organization.query.order_by(Organization.organization_name).all()
    return render_template("edit_directory_number.html", number=number, organizations=organizations)


@admin_bp.route("/admin/directory-numbers/delete/<int:id>", methods=["POST"])
@admin_required
def delete_directory_number(id):
    number = db.session.get(DirectoryNumber, id)
    if number:
        db.session.delete(number)
        db.session.commit()
        flash("Directory number deleted", "danger")
    return redirect(url_for("admin.directory_numbers"))


# ─── ADMINISTRATIVE HEADS ─────────────────────────────────────────────────────

def _parse_date(value):
    return datetime.strptime(value, "%Y-%m-%d").date() if value else None


@admin_bp.route("/admin/administrative-heads/add", methods=["GET", "POST"])
@admin_required
def add_administrative_head():
    if request.method == "POST":
        employee_id = optional_int(request.form.get("employee_id"))
        name = request.form.get("name", "").strip()
        office_phone = request.form.get("office_phone", "").strip()
        mobile_phone = request.form.get("mobile_phone", "").strip()
        email = request.form.get("email", "").strip()
        office_address = request.form.get("office_address", "").strip()
        organization_id = optional_int(request.form.get("organization_id"))
        organization_type_id = optional_int(request.form.get("organization_type_id"))
        role_title = request.form.get("role_title", "").strip()
        service_type_id = optional_int(request.form.get("service_type_id"))
        status = request.form.get("status", "ACTIVE").strip() or "ACTIVE"
        effective_from = _parse_date(request.form.get("effective_from"))
        remarks = request.form.get("remarks", "").strip()

        if employee_id and not db.session.get(Employee, employee_id):
            flash("Select a valid employee.", "danger")
            return redirect(request.url)
        if not employee_id and not name:
            flash("Either link an existing employee or enter a name for this head.", "danger")
            return redirect(request.url)
        if not organization_id:
            flash("Organization is required.", "danger")
            return redirect(request.url)
        if organization_type_mismatch(organization_id, organization_type_id):
            flash("Selected organization does not belong to the selected organization type.", "danger")
            return redirect(request.url)
        if not role_title:
            flash("Designation (role title) is required.", "danger")
            return redirect(request.url)

        head_service.assign_administrative_head(
            employee_id=employee_id, organization_id=organization_id,
            role_title=role_title, service_type_id=service_type_id,
            status=status, effective_from=effective_from, remarks=remarks,
            name=name or None, office_phone=office_phone or None,
            mobile_phone=mobile_phone or None, email=email or None,
            office_address=office_address or None,
        )
        db.session.commit()
        flash("Administrative Head assigned.", "success")
        return redirect(url_for("user.administrative_heads"))

    return render_template(
        "add_administrative_head.html",
        categories=OrganizationCategory.query.order_by(OrganizationCategory.category_name).all(),
        service_types=ServiceType.query.order_by(ServiceType.service_type_name).all(),
    )


@admin_bp.route("/admin/administrative-heads/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_administrative_head(id):
    head = db.session.get(AdministrativeHead, id)
    if not head or head.role_category != "ADMINISTRATIVE_HEAD":
        flash("Administrative Head not found", "danger")
        return redirect(url_for("user.administrative_heads"))

    if request.method == "POST":
        employee_id = optional_int(request.form.get("employee_id"))
        name = request.form.get("name", "").strip()
        office_phone = request.form.get("office_phone", "").strip()
        mobile_phone = request.form.get("mobile_phone", "").strip()
        email = request.form.get("email", "").strip()
        office_address = request.form.get("office_address", "").strip()
        organization_id = optional_int(request.form.get("organization_id"))
        organization_type_id = optional_int(request.form.get("organization_type_id"))
        role_title = request.form.get("role_title", "").strip()
        service_type_id = optional_int(request.form.get("service_type_id"))
        status = request.form.get("status", "ACTIVE").strip() or "ACTIVE"
        remarks = request.form.get("remarks", "").strip()

        if employee_id and not db.session.get(Employee, employee_id):
            flash("Select a valid employee.", "danger")
            return redirect(request.url)
        if not employee_id and not name:
            flash("Either link an existing employee or enter a name for this head.", "danger")
            return redirect(request.url)
        if not organization_id:
            flash("Organization is required.", "danger")
            return redirect(request.url)
        if organization_type_mismatch(organization_id, organization_type_id):
            flash("Selected organization does not belong to the selected organization type.", "danger")
            return redirect(request.url)
        if not role_title:
            flash("Designation (role title) is required.", "danger")
            return redirect(request.url)

        head_service.update_administrative_head(
            head, organization_id, role_title, service_type_id, remarks,
            employee_id=employee_id, name=name or None,
            office_phone=office_phone or None, mobile_phone=mobile_phone or None,
            email=email or None, office_address=office_address or None, status=status,
        )
        db.session.commit()
        flash("Administrative Head updated.", "success")
        return redirect(url_for("user.administrative_heads"))

    return render_template(
        "edit_administrative_head.html", head=head,
        categories=OrganizationCategory.query.order_by(OrganizationCategory.category_name).all(),
        service_types=ServiceType.query.order_by(ServiceType.service_type_name).all(),
    )


@admin_bp.route("/admin/administrative-heads/end/<int:id>", methods=["POST"])
@admin_required
def end_administrative_head(id):
    head = db.session.get(AdministrativeHead, id)
    if not head or head.role_category != "ADMINISTRATIVE_HEAD":
        flash("Administrative Head not found", "danger")
        return redirect(url_for("user.administrative_heads"))

    reason = request.form.get("reason", "").strip()
    if not reason:
        flash("A reason is required to end this role.", "danger")
        return redirect(request.referrer or url_for("user.administrative_heads"))

    replacement_employee_id = optional_int(request.form.get("replacement_employee_id"))
    head_service.end_role(head, reason=reason, replacement_employee_id=replacement_employee_id)
    db.session.commit()
    flash("Administrative Head role ended.", "warning")
    return redirect(url_for("user.administrative_heads"))


# ─── ADMINISTRATIVE HEAD ASSISTANTS (PA/PS) ──────────────────────────────────

@admin_bp.route("/admin/administrative-heads/<int:head_id>/assistants/add", methods=["GET", "POST"])
@admin_required
def add_administrative_head_assistant(head_id):
    head = db.session.get(AdministrativeHead, head_id)
    if not head:
        flash("Administrative Head not found", "danger")
        return redirect(url_for("user.administrative_heads"))

    if request.method == "POST":
        designation = request.form.get("designation", "").strip()
        employee_id = optional_int(request.form.get("employee_id"))
        name = request.form.get("name", "").strip()
        office_phone = request.form.get("office_phone", "").strip()
        mobile = request.form.get("mobile", "").strip()
        email = request.form.get("email", "").strip()
        is_primary = bool(request.form.get("is_primary"))

        if not designation:
            flash("Designation (PA/PS/...) is required.", "danger")
            return redirect(request.url)
        if not employee_id and not name:
            flash("Either pick an existing employee or enter a name.", "danger")
            return redirect(request.url)

        head_service.add_assistant(
            administrative_head_id=head.id, designation=designation,
            employee_id=employee_id or None, name=name or None,
            office_phone=office_phone or None, mobile=mobile or None,
            email=email or None, is_primary=is_primary, created_by=current_user.id,
        )
        db.session.commit()
        flash("Assistant added.", "success")
        return redirect(url_for("user.administrative_head_details", id=head.id))

    return render_template("administrative_head_assistant_form.html", head=head, assistant=None)


@admin_bp.route("/admin/administrative-heads/<int:head_id>/assistants/edit/<int:assistant_id>", methods=["GET", "POST"])
@admin_required
def edit_administrative_head_assistant(head_id, assistant_id):
    head = db.session.get(AdministrativeHead, head_id)
    assistant = db.session.get(AdministrativeHeadAssistant, assistant_id)
    if not head or not assistant or assistant.administrative_head_id != head.id:
        flash("Assistant not found", "danger")
        return redirect(url_for("user.administrative_heads"))

    if request.method == "POST":
        designation = request.form.get("designation", "").strip()
        employee_id = optional_int(request.form.get("employee_id"))
        name = request.form.get("name", "").strip()
        office_phone = request.form.get("office_phone", "").strip()
        mobile = request.form.get("mobile", "").strip()
        email = request.form.get("email", "").strip()

        if not designation:
            flash("Designation (PA/PS/...) is required.", "danger")
            return redirect(request.url)
        if not employee_id and not name:
            flash("Either pick an existing employee or enter a name.", "danger")
            return redirect(request.url)

        head_service.update_assistant(
            assistant, designation=designation, employee_id=employee_id or None,
            name=name or None, office_phone=office_phone or None,
            mobile=mobile or None, email=email or None,
        )
        db.session.commit()
        flash("Assistant updated.", "success")
        return redirect(url_for("user.administrative_head_details", id=head.id))

    return render_template("administrative_head_assistant_form.html", head=head, assistant=assistant)


@admin_bp.route("/admin/administrative-heads/<int:head_id>/assistants/delete/<int:assistant_id>", methods=["POST"])
@admin_required
def delete_administrative_head_assistant(head_id, assistant_id):
    assistant = db.session.get(AdministrativeHeadAssistant, assistant_id)
    if assistant and assistant.administrative_head_id == head_id:
        head_service.remove_assistant(assistant)
        db.session.commit()
        flash("Assistant removed.", "danger")
    return redirect(url_for("user.administrative_head_details", id=head_id))


@admin_bp.route("/admin/administrative-heads/<int:head_id>/assistants/<int:assistant_id>/set-primary", methods=["POST"])
@admin_required
def set_primary_administrative_head_assistant(head_id, assistant_id):
    head_service.set_primary_assistant(head_id, assistant_id)
    db.session.commit()
    flash("Primary assistant updated.", "success")
    return redirect(url_for("user.administrative_head_details", id=head_id))


# ─── EMAIL DISTRIBUTION LISTS (CUSTOM GROUPS) ────────────────────────────────

def _build_filter_summary(group):
    parts = []
    org_names = [f.organization.organization_name for f in group.filters if f.filter_type == "ORGANIZATION"]
    if org_names:
        parts.append("Organization: " + " OR ".join(org_names))
    cat_names = [f.category.category_name for f in group.filters if f.filter_type == "ORGANIZATION_CATEGORY"]
    if cat_names:
        parts.append("Category: " + " OR ".join(cat_names))
    role_names = [f.role_value for f in group.filters if f.filter_type == "ROLE"]
    if role_names:
        parts.append("Role: " + " OR ".join(role_names))
    status_names = [f.status_value for f in group.filters if f.filter_type == "STATUS"]
    if status_names:
        parts.append("Status: " + " OR ".join(status_names))
    summary = " AND ".join(parts) if parts else "(no filters)"

    include_count = sum(1 for f in group.filters if f.filter_type == "INCLUDE_EMPLOYEE")
    exclude_count = sum(1 for f in group.filters if f.filter_type == "EXCLUDE_EMPLOYEE")
    if include_count:
        summary += f" + {include_count} manually added"
    if exclude_count:
        summary += f" − {exclude_count} manually removed"
    return summary


def _apply_filters_from_form(group):
    """Replaces group.filters with the criteria submitted by the add/edit
    custom-group form -- simpler and less error-prone than diffing the old
    and new filter sets."""
    group.filters = []

    for oid in request.form.getlist("organization_ids"):
        group.filters.append(EmailGroupFilter(filter_type="ORGANIZATION", organization_id=int(oid)))
    for cid in request.form.getlist("category_ids"):
        group.filters.append(EmailGroupFilter(filter_type="ORGANIZATION_CATEGORY", category_id=int(cid)))
    for role in request.form.getlist("roles"):
        group.filters.append(EmailGroupFilter(filter_type="ROLE", role_value=role))
    for status in request.form.getlist("statuses"):
        group.filters.append(EmailGroupFilter(filter_type="STATUS", status_value=status))


@admin_bp.route("/admin/email-groups")
@admin_required
def manage_email_groups():
    groups = EmailGroup.query.filter_by(list_type="DYNAMIC").order_by(EmailGroup.name).all()
    summaries = []
    for group in groups:
        rows = resolve_dynamic_group(group)
        emails = dedupe_and_sort(rows)
        summaries.append({
            "group": group,
            "filter_summary": _build_filter_summary(group),
            "member_count": len(rows),
            "email_count": len(emails),
        })
    return render_template("manage_email_groups.html", summaries=summaries)


@admin_bp.route("/admin/email-groups/<int:id>/members")
@admin_required
def email_group_members(id):
    group = db.session.get(EmailGroup, id)
    if not group:
        flash("Email group not found", "danger")
        return redirect(url_for("admin.manage_email_groups"))

    rows = resolve_dynamic_group(group)
    contacts = [to_contact_row(r) for r in rows]
    contacts.sort(key=lambda c: (c["organization"], c["name"]))

    include_ids = {f.employee_id for f in group.filters if f.filter_type == "INCLUDE_EMPLOYEE"}
    for c in contacts:
        c["manually_added"] = c["record_type"] == "employee" and c["record_id"] in include_ids

    return render_template(
        "email_group_members.html", group=group, contacts=contacts,
        filter_summary=_build_filter_summary(group),
    )


@admin_bp.route("/admin/email-groups/<int:id>/members/add", methods=["POST"])
@admin_required
def add_email_group_member(id):
    group = db.session.get(EmailGroup, id)
    if not group:
        flash("Email group not found", "danger")
        return redirect(url_for("admin.manage_email_groups"))

    employee_id = optional_int(request.form.get("employee_id"))
    employee = db.session.get(Employee, employee_id) if employee_id else None
    if not employee:
        flash("Select a valid employee to add.", "danger")
        return redirect(url_for("admin.email_group_members", id=id))

    # Adding someone back cancels a prior manual exclusion of them.
    group.filters = [
        f for f in group.filters
        if not (f.filter_type in ("INCLUDE_EMPLOYEE", "EXCLUDE_EMPLOYEE") and f.employee_id == employee_id)
    ]
    group.filters.append(EmailGroupFilter(filter_type="INCLUDE_EMPLOYEE", employee_id=employee_id))
    db.session.commit()
    flash(f"{employee.employee_name} added to {group.name}.", "success")
    return redirect(url_for("admin.email_group_members", id=id))


@admin_bp.route("/admin/email-groups/<int:id>/members/remove", methods=["POST"])
@admin_required
def remove_email_group_member(id):
    group = db.session.get(EmailGroup, id)
    if not group:
        flash("Email group not found", "danger")
        return redirect(url_for("admin.manage_email_groups"))

    employee_id = optional_int(request.form.get("employee_id"))
    if not employee_id:
        flash("Invalid member.", "danger")
        return redirect(url_for("admin.email_group_members", id=id))

    # Removing someone cancels a prior manual inclusion, then (if they'd
    # still match the group's filters) explicitly excludes them so removal
    # sticks regardless of how they got in.
    group.filters = [
        f for f in group.filters
        if not (f.filter_type in ("INCLUDE_EMPLOYEE", "EXCLUDE_EMPLOYEE") and f.employee_id == employee_id)
    ]
    group.filters.append(EmailGroupFilter(filter_type="EXCLUDE_EMPLOYEE", employee_id=employee_id))
    db.session.commit()
    flash("Member removed.", "danger")
    return redirect(url_for("admin.email_group_members", id=id))


@admin_bp.route("/admin/email-groups/add", methods=["GET", "POST"])
@admin_required
def add_email_group():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        description = request.form.get("description", "").strip()

        if not name:
            flash("Group name is required.", "danger")
            return redirect(request.url)
        if EmailGroup.query.filter_by(name=name).first():
            flash("A group with this name already exists.", "danger")
            return redirect(request.url)

        has_any_filter = any(
            request.form.getlist(key)
            for key in ("organization_ids", "category_ids", "roles", "statuses")
        )
        if not has_any_filter:
            flash("Select at least one filter for the group.", "danger")
            return redirect(request.url)

        group = EmailGroup(
            name=name, description=description, list_type="DYNAMIC",
            created_by=current_user.id,
        )
        _apply_filters_from_form(group)
        db.session.add(group)
        db.session.flush()
        log_audit_event(
            module="email_groups", record_type="EmailGroup", record_id=group.id,
            action="CREATE", field_name="name", old_value=None, new_value=name,
        )
        db.session.commit()
        flash("Email group created.", "success")
        return redirect(url_for("admin.manage_email_groups"))

    categories = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    organizations = Organization.query.order_by(Organization.organization_name).all()
    return render_template(
        "email_group_form.html", group=None,
        categories=categories, organizations=organizations,
        roles=ROLE_CHOICES, statuses=STATUS_CHOICES,
    )


@admin_bp.route("/admin/email-groups/edit/<int:id>", methods=["GET", "POST"])
@admin_required
def edit_email_group(id):
    group = db.session.get(EmailGroup, id)
    if not group or group.list_type != "DYNAMIC":
        flash("Email group not found", "danger")
        return redirect(url_for("admin.manage_email_groups"))

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        description = request.form.get("description", "").strip()

        if not name:
            flash("Group name is required.", "danger")
            return redirect(request.url)

        duplicate = EmailGroup.query.filter(
            EmailGroup.name == name, EmailGroup.id != id,
        ).first()
        if duplicate:
            flash("A group with this name already exists.", "danger")
            return redirect(request.url)

        has_any_filter = any(
            request.form.getlist(key)
            for key in ("organization_ids", "category_ids", "roles", "statuses")
        )
        if not has_any_filter:
            flash("Select at least one filter for the group.", "danger")
            return redirect(request.url)

        group.name = name
        group.description = description
        _apply_filters_from_form(group)
        log_audit_event(
            module="email_groups", record_type="EmailGroup", record_id=group.id,
            action="UPDATE", field_name="filters", old_value=None, new_value=None,
        )
        db.session.commit()
        flash("Email group updated.", "success")
        return redirect(url_for("admin.manage_email_groups"))

    categories = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    organizations = Organization.query.order_by(Organization.organization_name).all()
    return render_template(
        "email_group_form.html", group=group,
        categories=categories, organizations=organizations,
        roles=ROLE_CHOICES, statuses=STATUS_CHOICES,
    )


@admin_bp.route("/admin/email-groups/delete/<int:id>", methods=["POST"])
@admin_required
def delete_email_group(id):
    group = db.session.get(EmailGroup, id)
    if group:
        log_audit_event(
            module="email_groups", record_type="EmailGroup", record_id=group.id,
            action="DELETE", field_name="name", old_value=group.name, new_value=None,
        )
        db.session.delete(group)
        db.session.commit()
        flash("Email group deleted", "danger")
    return redirect(url_for("admin.manage_email_groups"))
