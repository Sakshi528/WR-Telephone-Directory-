import io
import re
from datetime import datetime

from flask import (
    Blueprint, jsonify, render_template, request,
    redirect, url_for, flash, send_file
)
from flask_login import current_user
from sqlalchemy import or_

from models import db
from models.employee import Employee, EMPLOYEE_STATUSES
from models.organization import Organization
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from models.update_request import UpdateRequest
from models.administrative_head import AdministrativeHead
from models.organization_category import OrganizationCategory
from models.service_type import ServiceType
from models.directory_version import DirectoryVersion
from models.email_group import EmailGroup
from utils.designation_rank import resolve_utility_head, compute_utility_head_ids
from utils.section_order import section_sort_key
from utils.user_agent import parse_user_agent
from services import head_service
from services import directory_version_service
from services.email_distribution_service import (
    STANDARD_GROUPS, STANDARD_GROUPS_BY_SLUG, CONTACT_TYPE_CHOICES, CATEGORY_SLUGS,
    dedupe_and_sort, build_mailto, MAILTO_SAFE_LIMIT,
    resolve_category_contacts, to_contact_row, summarize_contacts, filter_contact_rows,
    resolve_dynamic_group,
)


user_bp = Blueprint("user", __name__)


def _request_metadata():
    """IP/browser/OS for the submitting request -- same approach as
    services/audit_service.py's _request_metadata (see
    utils/user_agent.py: Werkzeug 3.x no longer parses User-Agent strings
    on its own), duplicated locally rather than imported since
    update_requests has no session_id column (only audit_logs does) and
    the two capture slightly different fields."""
    ip_address = request.headers.get("X-Forwarded-For", request.remote_addr)
    browser, operating_system = parse_user_agent(request.headers.get("User-Agent"))
    return ip_address, browser, operating_system


# ─── SEARCH FILTERS ──────────────────────────────────────────────────────────

def employee_search_filter(keyword):
    filters = [
        Employee.employee_name.ilike(f"%{keyword}%"),
        Employee.designation.ilike(f"%{keyword}%"),
        Employee.region.ilike(f"%{keyword}%"),
        Organization.organization_name.ilike(f"%{keyword}%"),
        Organization.region.ilike(f"%{keyword}%"),
        Organization.address.ilike(f"%{keyword}%"),
        Department.department_name.ilike(f"%{keyword}%"),
        Employee.location.ilike(f"%{keyword}%"),
        Employee.office_phone.ilike(f"%{keyword}%"),
        Employee.mobile_phone.ilike(f"%{keyword}%"),
    ]
    # Only search email when the keyword looks like an email address,
    # otherwise "ntpc" matches every @ntpc.co.in employee across all companies.
    if "@" in keyword:
        filters.append(Employee.email.ilike(f"%{keyword}%"))
    return or_(*filters)


def directory_number_search_filter(keyword):
    from sqlalchemy import select as sa_select
    # match dir numbers whose org string belongs to an org in the searched region
    region_org_names = sa_select(Organization.organization_name).where(
        Organization.region.ilike(f"%{keyword}%")
    )
    filters = [
        DirectoryNumber.name.ilike(f"%{keyword}%"),
        DirectoryNumber.phone_number.ilike(f"%{keyword}%"),
        DirectoryNumber.category.ilike(f"%{keyword}%"),
        DirectoryNumber.organization.ilike(f"%{keyword}%"),
        DirectoryNumber.organization.in_(region_org_names),
    ]
    if "@" in keyword:
        filters.append(DirectoryNumber.email.ilike(f"%{keyword}%"))
    return or_(*filters)


def get_matched_organizations(keyword):
    return Organization.query.filter(
        or_(
            Organization.organization_name.ilike(f"%{keyword}%"),
            Organization.region.ilike(f"%{keyword}%"),
            Organization.address.ilike(f"%{keyword}%"),
        )
    ).all()


# ─── TELEPHONE DIRECTORY CATEGORY CARDS (landing page) ───────────────────────
# Presentational only -- short blurbs shown on the landing cards, not stored
# in the database (organization_categories.description is unused elsewhere
# and left untouched). Keyed by category_name so a category with no blurb
# here just shows no description rather than erroring.

TELEPHONE_DIRECTORY_CATEGORY_DESCRIPTIONS = {
    "RLDC": "Regional and National Load Despatch Centres",
    "State SLDC": "State Load Despatch Centres",
    "CPSU": "Central Public Sector Undertakings",
    "DISCOM": "State power distribution companies",
    "Generation Company": "Power generation companies and utilities",
    "Thermal": "Thermal power stations",
    "Hydel": "Hydroelectric power stations",
    "Nuclear": "Nuclear power stations",
    "RE Generators": "Renewable energy generating stations",
    "Transmission Utility": "Power transmission utilities",
    "QCA": "Qualified Coordinating Agencies",
    "Others": "Organizations not yet assigned a specific type",
}


def _compute_telephone_directory_categories():
    """One card per OrganizationCategory that has at least one organization
    -- organization count and contact count (active KMP employees) both
    reuse the same is_kmp/status filters telephone_directory() itself uses,
    so the numbers on a card match what its filtered view will show."""
    categories = OrganizationCategory.query.order_by(OrganizationCategory.category_name).all()
    cards = []
    for cat in categories:
        org_count = Organization.query.filter_by(category_id=cat.id).count()
        if org_count == 0:
            continue
        contact_count = (
            Employee.query
            .join(Organization, Employee.organization_id == Organization.id)
            .filter(Organization.category_id == cat.id)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .filter(Employee.status == "ACTIVE")
            .count()
        )
        cards.append({
            "id": cat.id,
            "name": cat.category_name,
            "description": TELEPHONE_DIRECTORY_CATEGORY_DESCRIPTIONS.get(cat.category_name, ""),
            "organization_count": org_count,
            "contact_count": contact_count,
        })
    return cards


# ─── SUGGESTION HELPERS ──────────────────────────────────────────────────────

def add_suggestion(suggestions, seen, label, value, category):
    if not label or not value:
        return
    key = value.strip().lower()
    if key in seen:
        return
    seen.add(key)
    suggestions.append({"label": label, "value": value, "category": category})


def ranked_items(items, keyword, label_getter):
    keyword = keyword.lower()
    return sorted(
        items,
        key=lambda item: (
            not label_getter(item).lower().startswith(keyword),
            label_getter(item).lower(),
        ),
    )


# ─── HOME ────────────────────────────────────────────────────────────────────

@user_bp.route("/")
def home():
    # If already logged-in admin, redirect to dashboard
    if current_user.is_authenticated and current_user.role == "admin":
        return redirect(url_for("admin.dashboard"))

    if request.args.get("expired"):
        flash("Your session expired due to inactivity.", "warning")

    total_employees = Employee.query.filter_by(is_kmp=False).count()
    total_organizations = Organization.query.count()

    # Last updated = most recently created employee record (proxy for data freshness)
    last_emp = Employee.query.order_by(Employee.id.desc()).first()
    last_updated = last_emp.id if last_emp else None   # placeholder; real date needs created_at

    return render_template(
        "index.html",
        total_employees=total_employees,
        total_organizations=total_organizations,
        last_updated=datetime.today().strftime("%B %Y"),
    )


@user_bp.route("/session-expired")
def session_expired():
    flash("Your session expired due to inactivity.", "warning")
    return redirect(url_for("user.home"))


# ─── EMPLOYEE DIRECTORY (WRLDC ONLY) ─────────────────────────────────────────

@user_bp.route("/directory")
def directory():
    keyword = request.args.get("keyword", "")

    query = Employee.query.outerjoin(Organization).outerjoin(Department)
    query = query.filter(Employee.is_kmp == False)  # noqa: E712
    # Non-ACTIVE employees (retired/transferred/etc.) don't appear in the
    # live directory.
    query = query.filter(Employee.status == "ACTIVE")

    if keyword:
        query = query.filter(employee_search_filter(keyword))

    employees = query.order_by(Employee.employee_name).all()

    # Build emergency-contact lookup: employee_id → first EmergencyContact
    emp_ids = [e.id for e in employees]
    ec_map = {}
    if emp_ids:
        contacts = EmergencyContact.query.filter(
            EmergencyContact.employee_id.in_(emp_ids)
        ).all()
        for c in contacts:
            if c.employee_id not in ec_map:
                ec_map[c.employee_id] = c

    return render_template(
        "directory.html",
        employees=employees,
        keyword=keyword,
        ec_map=ec_map,
    )


@user_bp.route("/employee/<int:id>")
def employee_details(id):
    employee = db.session.get(Employee, id) or _abort404()
    return render_template("employee_details.html", employee=employee)


# ─── UPDATE REQUESTS (ANONYMOUS) ─────────────────────────────────────────────

@user_bp.route("/employee/request-update/<int:id>", methods=["GET", "POST"])
def request_update(id):
    employee = db.session.get(Employee, id) or _abort404()
    if request.method == "POST":
        ip_address, browser, operating_system = _request_metadata()
        update_request = UpdateRequest(
            user_id=None,
            employee_id=employee.id,
            request_type="employee",
            requested_by=request.form.get("requested_by"),
            department=request.form.get("department"),
            contact_number=request.form.get("contact_number"),
            field_name=request.form.get("field_name"),
            new_value=request.form.get("new_value"),
            reason=request.form.get("reason"),
            ip_address=ip_address, browser=browser, operating_system=operating_system,
        )
        db.session.add(update_request)
        db.session.commit()
        flash("Update request submitted. Thank you!", "success")
        return redirect(url_for("user.employee_details", id=employee.id))
    requester_ip, _browser, _os = _request_metadata()
    current_values = {
        "mobile_phone": employee.mobile_phone or "",
        "office_phone": employee.office_phone or "",
        "residence_phone": employee.residence_phone or "",
        "email": employee.email or "",
        "designation": employee.designation or "",
        "address": employee.location or "",
        "status": employee.status or "ACTIVE",
    }
    if not employee.is_kmp:
        emergency_contact = EmergencyContact.query.filter_by(employee_id=employee.id).first()
        current_values.update({
            "emergency_contact_name": emergency_contact.contact_name if emergency_contact else "",
            "emergency_contact_relation": emergency_contact.relation if emergency_contact else "",
            "emergency_contact_phone": emergency_contact.phone if emergency_contact else "",
        })
    return render_template(
        "update_request.html", employee=employee, requester_ip=requester_ip,
        current_values=current_values,
        employee_statuses=EMPLOYEE_STATUSES,
    )


@user_bp.route(
    "/directory-number/request-update/<int:id>", methods=["GET", "POST"]
)
def request_directory_update(id):
    number = db.session.get(DirectoryNumber, id) or _abort404()
    if request.method == "POST":
        ip_address, browser, operating_system = _request_metadata()
        update_request = UpdateRequest(
            user_id=None,
            directory_number_id=number.id,
            request_type="directory",
            requested_by=request.form.get("requested_by"),
            department=request.form.get("department"),
            contact_number=request.form.get("contact_number"),
            field_name=request.form.get("field_name"),
            new_value=request.form.get("new_value"),
            reason=request.form.get("reason"),
            ip_address=ip_address, browser=browser, operating_system=operating_system,
        )
        db.session.add(update_request)
        db.session.commit()
        flash("Directory correction request submitted. Thank you!", "success")
        return redirect(url_for("user.directory_number_details", id=number.id))
    requester_ip, _browser, _os = _request_metadata()
    current_values = {
        "phone_number": number.phone_number or "",
        "email": number.email or "",
        "name": number.name or "",
        "organization": number.organization or "",
    }
    return render_template(
        "directory_update_request.html", number=number, requester_ip=requester_ip,
        current_values=current_values,
    )


@user_bp.route(
    "/administrative-head/request-update/<int:id>", methods=["GET", "POST"]
)
def request_administrative_head_update(id):
    head = db.session.get(AdministrativeHead, id) or _abort404()
    if request.method == "POST":
        ip_address, browser, operating_system = _request_metadata()
        update_request = UpdateRequest(
            user_id=None,
            administrative_head_id=head.id,
            request_type="administrative_head",
            requested_by=request.form.get("requested_by"),
            department=request.form.get("department"),
            contact_number=request.form.get("contact_number"),
            field_name=request.form.get("field_name"),
            new_value=request.form.get("new_value"),
            reason=request.form.get("reason"),
            ip_address=ip_address, browser=browser, operating_system=operating_system,
        )
        db.session.add(update_request)
        db.session.commit()
        flash("Update request submitted. Thank you!", "success")
        return redirect(url_for("user.administrative_head_details", id=head.id))
    requester_ip, _browser, _os = _request_metadata()
    current_values = {
        "role_title": head.resolved_designation or "",
        "office_phone": head.resolved_office_phone or "",
        "mobile_phone": head.resolved_mobile_phone or "",
        "email": head.resolved_email or "",
        "office_address": head.resolved_office_address or "",
        "status": head.status or "ACTIVE",
    }
    return render_template(
        "administrative_head_update_request.html", head=head, requester_ip=requester_ip,
        current_values=current_values,
    )


# ─── TELEPHONE DIRECTORY (ALL ORGS) ──────────────────────────────────────────

@user_bp.route("/telephone-directory")
def telephone_directory():
    keyword = request.args.get("keyword", "")
    category_id = request.args.get("category_id", type=int)
    employees = []
    control_rooms = []
    other_numbers = []
    matched_organizations = []

    if not keyword and not category_id:
        # Landing page: category cards instead of the full flat listing --
        # same route, same "no keyword" branch structure, just a different
        # template mode. See _compute_telephone_directory_categories().
        return render_template(
            "telephone_directory.html",
            landing=True,
            categories=_compute_telephone_directory_categories(),
            grouped={}, other_numbers=[], matched_organizations=[],
            keyword="", employees=[], control_rooms=[], numbers=[],
            category_id=None, current_category=None,
        )

    if not keyword:
        # A category card was clicked -- same grouping logic as the old
        # "no keyword" branch, just scoped to that category's organizations.
        current_category = db.session.get(OrganizationCategory, category_id)
        category_org_ids = [
            o.id for o in Organization.query.filter_by(category_id=category_id).all()
        ]

        # Build org address lookup
        all_orgs = Organization.query.all()
        org_address = {o.organization_name: (o.address or "") for o in all_orgs}

        # Seed a card for every organization in this category up front --
        # previously an org only got a card once an employee/directory-number
        # loop below happened to touch it, so a correctly-categorized
        # organization with no employees or control rooms yet (e.g. many of
        # the private transmission SPVs, which the Word directory lists by
        # name/address only, with no roster) silently never appeared, even
        # though the category card's own count included it.
        category_orgs = Organization.query.filter(Organization.id.in_(category_org_ids)).all()
        grouped = {
            o.organization_name: {
                "employees": [], "control_rooms": [], "switchyards": [],
                "address": o.address or "",
            }
            for o in category_orgs
        }

        utility_head_ids = compute_utility_head_ids()
        all_employees = (
            Employee.query
            .outerjoin(Organization)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .filter(Employee.status == "ACTIVE")
            .filter(Employee.organization_id.in_(category_org_ids))
            .order_by(
                Organization.organization_name,
                Employee.employee_name,
            )
            .all()
        )
        # heads-first within each org group -- sort is stable, so the SQL
        # order above is preserved among ties
        all_employees.sort(key=lambda e: e.id not in utility_head_ids)
        all_numbers = (
            DirectoryNumber.query
            .filter(DirectoryNumber.organization != None)
            .filter(DirectoryNumber.organization_id.in_(category_org_ids))
            .order_by(DirectoryNumber.organization)
            .all()
        )

        for emp in all_employees:
            org_name = (
                emp.organization.organization_name if emp.organization else "Unknown"
            )
            if org_name not in grouped:
                grouped[org_name] = {
                    "employees": [],
                    "control_rooms": [],
                    "switchyards": [],
                    "address": emp.organization.address if emp.organization else "",
                }
            grouped[org_name]["employees"].append(emp)

        default_other = []
        for dn in all_numbers:
            org_name = dn.organization or "Unknown"
            if org_name not in grouped:
                grouped[org_name] = {
                    "employees": [],
                    "control_rooms": [],
                    "switchyards": [],
                    "address": org_address.get(org_name, ""),
                }
            if dn.category and "switchyard" in dn.category.lower():
                grouped[org_name]["switchyards"].append(dn)
            elif dn.category and "control room" in dn.category.lower():
                grouped[org_name]["control_rooms"].append(dn)
            else:
                default_other.append(dn)

        return render_template(
            "telephone_directory.html",
            landing=False,
            grouped=dict(sorted(grouped.items(), key=lambda kv: section_sort_key(kv[0]))),
            other_numbers=default_other,
            matched_organizations=[],
            keyword="",
            employees=all_employees,
            control_rooms=[],
            numbers=default_other,
            category_id=category_id,
            current_category=current_category,
        )

    matched_organizations = get_matched_organizations(keyword)
    matched_org_names = [o.organization_name for o in matched_organizations]
    matched_org_ids   = [o.id for o in matched_organizations]

    # Build subtree: matched orgs + their direct children.
    # Used for both the employee query and the CR/switchyard FK filter.
    subtree_ids = list(matched_org_ids)
    child_org_objects = []
    for org in matched_organizations:
        children = Organization.query.filter_by(parent_id=org.id).all()
        child_org_objects.extend(children)
        for child in children:
            if child.id not in subtree_ids:
                subtree_ids.append(child.id)

    utility_head_ids = compute_utility_head_ids()
    emp_query = (
        Employee.query.outerjoin(Organization).outerjoin(Department)
        .filter(Employee.is_kmp == True)  # noqa: E712
        .filter(Employee.status == "ACTIVE")
    )
    if subtree_ids:
        employees = emp_query.filter(
            or_(
                Employee.organization_id.in_(subtree_ids),
                employee_search_filter(keyword),
            )
        ).order_by(
            Organization.organization_name,
            Employee.employee_name,
        ).all()
    else:
        employees = emp_query.filter(
            employee_search_filter(keyword)
        ).order_by(
            Employee.employee_name,
        ).all()
    # heads-first within each org group -- sort is stable, so the SQL
    # order above is preserved among ties
    employees.sort(key=lambda e: e.id not in utility_head_ids)
    employees = employees[:200]

    cr_filters = [
        DirectoryNumber.name.ilike(f"%{keyword}%"),
        DirectoryNumber.phone_number.ilike(f"%{keyword}%"),
    ]
    if "@" in keyword:
        cr_filters.append(DirectoryNumber.email.ilike(f"%{keyword}%"))
    if subtree_ids:
        cr_filters.append(DirectoryNumber.organization_id.in_(subtree_ids))

    all_matched_numbers = DirectoryNumber.query.filter(or_(*cr_filters)).all()
    control_rooms = [
        n for n in all_matched_numbers
        if n.category and "control room" in n.category.lower()
    ]
    switchyards = [
        n for n in all_matched_numbers
        if n.category and "switchyard" in n.category.lower()
    ]
    other_numbers = [
        n for n in all_matched_numbers
        if not (n.category and "control room" in n.category.lower())
        and not (n.category and "switchyard" in n.category.lower())
    ]

    # Build org address lookup
    org_address = {o.organization_name: (o.address or "") for o in matched_organizations}

    grouped = {}
    for emp in employees:
        org_name = (
            emp.organization.organization_name if emp.organization else "Unknown"
        )
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": emp.organization.address if emp.organization else "",
            }
        grouped[org_name]["employees"].append(emp)

    for cr in control_rooms:
        org_name = cr.organization or "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": org_address.get(org_name, ""),
            }
        grouped[org_name]["control_rooms"].append(cr)

    for sw in switchyards:
        org_name = sw.organization or "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": org_address.get(org_name, ""),
            }
        grouped[org_name]["switchyards"].append(sw)

    # When a parent org matches the keyword, include ALL its child orgs (even empty),
    # so that e.g. "Black Start" shows every sub-org with its address.
    child_orgs = set()
    for child in child_org_objects:
        child_orgs.add(child.organization_name)
        if child.organization_name not in grouped:
            grouped[child.organization_name] = {
                "employees": [],
                "control_rooms": [],
                "switchyards": [],
                "address": child.address or "",
            }

    # Show orgs that have content, or that are explicitly sub-orgs of a matched parent.
    grouped_sorted = dict(sorted(
        (k, v) for k, v in grouped.items()
        if v["employees"] or v["control_rooms"] or v["switchyards"] or k in child_orgs
    ))

    return render_template(
        "telephone_directory.html",
        landing=False,
        grouped=grouped_sorted,
        other_numbers=other_numbers,
        matched_organizations=matched_organizations,
        keyword=keyword,
        employees=employees,
        control_rooms=control_rooms,
        numbers=other_numbers,
        category_id=None,
        current_category=None,
    )


@user_bp.route("/directory-number/<int:id>")
def directory_number_details(id):
    number = db.session.get(DirectoryNumber, id) or _abort404()
    return render_template("directory_number_details.html", number=number)


# ─── UNIVERSAL SEARCH ─────────────────────────────────────────────────────────

@user_bp.route("/search")
def universal_search():
    keyword = request.args.get("keyword", "")
    employees = []
    directory_numbers = []

    if keyword:
        employees = Employee.query.outerjoin(Organization).outerjoin(
            Department
        ).filter(employee_search_filter(keyword)).all()

        directory_numbers = DirectoryNumber.query.filter(
            directory_number_search_filter(keyword)
        ).all()

    return render_template(
        "universal_search.html",
        employees=employees,
        directory_numbers=directory_numbers,
        keyword=keyword,
    )


# ─── ORGANIZATION TYPE CASCADE API ───────────────────────────────────────────
# Backs static/js/org_type_cascade.js's Type -> State -> Organization
# dropdowns, used on Utility Heads/Administrative Heads filters and the
# Employee/Administrative Head add/edit forms.

@user_bp.route("/api/organization-types/<int:type_id>/states")
def api_organization_type_states(type_id):
    states = (
        db.session.query(Organization.state)
        .filter(Organization.category_id == type_id, Organization.state.isnot(None))
        .distinct().order_by(Organization.state).all()
    )
    return jsonify([{"id": s[0], "name": s[0]} for s in states])


@user_bp.route("/api/organizations")
def api_organizations():
    type_id = request.args.get("type_id", type=int)
    state = request.args.get("state", "").strip()
    query = Organization.query
    if type_id:
        query = query.filter(Organization.category_id == type_id)
    if state:
        query = query.filter(Organization.state == state)
    orgs = query.order_by(Organization.organization_name).all()
    return jsonify([{"id": o.id, "name": o.organization_name} for o in orgs])


# ─── UTILITY HEADS ────────────────────────────────────────────────────────────
# Resolved per organization via utils.designation_rank.resolve_utility_head:
# a manually-flagged employee (Employee.is_utility_head) wins if one exists,
# otherwise falls back to the most senior designation -- see that module for
# the ranking rules.

def _compute_utility_heads(keyword=None, category_id=None, state=None, organization_id=None, designation=None):
    org_query = Organization.query
    if organization_id:
        org_query = org_query.filter(Organization.id == organization_id)
    else:
        if category_id:
            org_query = org_query.filter(Organization.category_id == category_id)
        if state:
            org_query = org_query.filter(Organization.state == state)
    organizations = org_query.order_by(Organization.organization_name).all()

    heads = []
    for org in organizations:
        if org.utility_head_excluded:
            continue
        employees = (
            Employee.query
            .filter_by(organization_id=org.id)
            .filter(Employee.status == "ACTIVE")
            .order_by(Employee.id)
            .all()
        )
        head = resolve_utility_head(employees)
        if head:
            heads.append(head)

    # keyword ORs across organization name / head name / designation (matches
    # Administrative Heads' free-text search box) -- done in Python since
    # heads are only known after the per-org resolve_utility_head pass above.
    if keyword:
        kw = keyword.lower()
        heads = [
            h for h in heads
            if kw in (h.organization.organization_name if h.organization else "").lower()
            or kw in (h.employee_name or "").lower()
            or kw in (h.designation or "").lower()
        ]
    if designation:
        d = designation.lower()
        heads = [h for h in heads if d in (h.designation or "").lower()]
    return heads


@user_bp.route("/utility-heads")
def utility_heads():
    keyword = request.args.get("keyword", "").strip()
    category_id = request.args.get("category_id", type=int)
    state = request.args.get("state", "").strip()
    organization_id = request.args.get("organization_id", type=int)
    designation = request.args.get("designation", "").strip()
    heads = _compute_utility_heads(keyword or None, category_id, state or None, organization_id, designation or None)
    return render_template(
        "utility_heads.html", heads=heads, keyword=keyword,
        category_id=category_id, state=state, organization_id=organization_id,
        designation=designation,
        categories=OrganizationCategory.query.order_by(OrganizationCategory.category_name).all(),
        emails=dedupe_and_sort(heads), mailto_limit=MAILTO_SAFE_LIMIT,
    )


@user_bp.route("/utility-heads/suggest")
def utility_heads_suggest():
    q = request.args.get("q", "").strip()
    if len(q) < 1:
        return jsonify([])
    orgs = (
        Organization.query
        .join(Employee, Organization.id == Employee.organization_id)
        .filter(Organization.organization_name.ilike(f"%{q}%"))
        .with_entities(Organization.organization_name)
        .distinct()
        .order_by(Organization.organization_name)
        .limit(10)
        .all()
    )
    return jsonify([row[0] for row in orgs])


# ─── EMERGENCY CONTACTS ───────────────────────────────────────────────────────

@user_bp.route("/emergency-contacts")
def emergency_contacts():
    contacts = (
        EmergencyContact.query
        .join(Employee)
        .order_by(Employee.employee_name)
        .all()
    )
    return render_template("emergency_contacts.html", contacts=contacts)


# ─── AUTOCOMPLETE SUGGESTIONS ─────────────────────────────────────────────────

@user_bp.route("/suggestions")
def search_suggestions():
    keyword = request.args.get("q", "").strip()
    source  = request.args.get("source", "all")
    suggestions = []
    seen = set()

    if len(keyword) < 2:
        return jsonify(suggestions)

    if source in {"all", "employee"}:
        organizations = Organization.query.filter(
            or_(
                Organization.organization_name.ilike(f"%{keyword}%"),
                Organization.region.ilike(f"%{keyword}%"),
            )
        ).limit(20).all()

        for org in ranked_items(organizations, keyword, lambda o: o.organization_name)[:6]:
            add_suggestion(suggestions, seen,
                           org.organization_name, org.organization_name, "Organization")

        departments = Department.query.filter(
            Department.department_name.ilike(f"%{keyword}%")
        ).limit(12).all()

        for dept in ranked_items(departments, keyword, lambda d: d.department_name)[:4]:
            add_suggestion(suggestions, seen,
                           dept.department_name, dept.department_name, "Department")

        emps = Employee.query.outerjoin(Organization).outerjoin(
            Department
        ).filter(employee_search_filter(keyword)).limit(24).all()

        for emp in ranked_items(emps, keyword, lambda e: e.employee_name)[:8]:
            label = emp.employee_name
            if emp.designation:
                label = f"{label} - {emp.designation}"
            add_suggestion(suggestions, seen, label, emp.employee_name, "Employee")

    if source in {"all", "directory"}:
        numbers = DirectoryNumber.query.filter(
            directory_number_search_filter(keyword)
        ).limit(24).all()

        for number in ranked_items(numbers, keyword, lambda n: n.name)[:8]:
            label = number.name
            if number.phone_number:
                label = f"{label} - {number.phone_number}"
            add_suggestion(suggestions, seen, label, number.name,
                           number.category or "Directory")

    return jsonify(suggestions[:10])


# ─── EXPORT ───────────────────────────────────────────────────────────────────

@user_bp.route("/export/employees")
def export_employees():
    fmt      = request.args.get("format", "xlsx")
    keyword  = request.args.get("keyword", "")
    org_id   = request.args.get("organization", "")
    dept_id  = request.args.get("department", "")

    wrldc_ids = [
        o.id for o in Organization.query.filter(
            or_(
                Organization.organization_name.ilike("%WRLDC%"),
                Organization.organization_name.ilike("%Western Region Load Despatch%"),
                Organization.organization_name.ilike("%Western Region Load Dispatch%"),
            )
        ).all()
    ]

    query = Employee.query.outerjoin(Organization).outerjoin(Department)
    query = query.filter(Employee.organization_id.in_(wrldc_ids))
    query = query.filter(Employee.status == "ACTIVE")
    if keyword:
        query = query.filter(employee_search_filter(keyword))
    if org_id:
        query = query.filter(Employee.organization_id == org_id)
    if dept_id:
        query = query.filter(Employee.department_id == dept_id)

    employees = query.order_by(Employee.employee_name).all()

    if fmt == "docx":
        return _export_employees_docx(employees)
    return _export_employees_xlsx(employees)


@user_bp.route("/export/telephone")
def export_telephone():
    fmt     = request.args.get("format", "xlsx")
    keyword = request.args.get("keyword", "")

    # Reuse same query logic as telephone_directory view
    if keyword:
        matched_organizations = get_matched_organizations(keyword)
        matched_org_ids = [o.id for o in matched_organizations]

        emp_query = (
            Employee.query.outerjoin(Organization).outerjoin(Department)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .filter(Employee.status == "ACTIVE")
        )
        if matched_org_ids:
            employees = emp_query.filter(
                or_(
                    Employee.organization_id.in_(matched_org_ids),
                    employee_search_filter(keyword),
                )
            ).order_by(Organization.organization_name, Employee.employee_name).all()
        else:
            employees = emp_query.filter(
                employee_search_filter(keyword)
            ).order_by(Employee.employee_name).all()
    else:
        employees = (
            Employee.query.outerjoin(Organization)
            .filter(Employee.is_kmp == True)  # noqa: E712
            .filter(Employee.status == "ACTIVE")
            .order_by(Organization.organization_name, Employee.employee_name)
            .all()
        )

    if fmt == "docx":
        return _export_employees_docx(employees, title="Telephone Directory")
    return _export_employees_xlsx(employees, title="Telephone Directory")


@user_bp.route("/export/org-word/<path:org_name>")
def export_org_word(org_name):
    employees = (
        Employee.query
        .join(Organization)
        .filter(Organization.organization_name == org_name)
        .order_by(Employee.employee_name)
        .all()
    )
    control_rooms = (
        DirectoryNumber.query
        .filter(
            DirectoryNumber.organization.ilike(f"%{org_name}%"),
            DirectoryNumber.category == "Control Room",
        )
        .all()
    )
    org = Organization.query.filter_by(organization_name=org_name).first()
    return _export_org_docx(org_name, org.address if org else "", employees, control_rooms)


# ─── EXPORT HELPERS ───────────────────────────────────────────────────────────

# Fixes a docx export bug: long phone/email values in a Control Room table
# cell rendered overlapping the adjacent cell instead of wrapping. Root
# cause (confirmed by inspecting the generated XML) -- doc.add_table()
# leaves autofit's underlying <w:tblW type="auto" w="0"/> paired with a
# naive equal-split <w:tblGrid> (every column gets the same width
# regardless of what it holds), and phone/email values are long strings
# with few or no space characters, so many renderers have no wrap point
# inside a too-narrow column and the text overflows into the next one.

_BREAK_AFTER_RE = re.compile(r"([/;,@-])")


def _wrappable(text):
    """Inserts a zero-width space after /, ;, ,, -, @ separators so Word always
    has a wrap point inside a long unbroken phone/email string, regardless
    of column width -- defense-in-depth on top of _set_column_widths below,
    since even a generously-sized column can't wrap text with no break
    point at all (e.g. a long duplicated-phone string or a plain email
    address)."""
    if not text:
        return text
    return _BREAK_AFTER_RE.sub("\\1\u200b", text)


def _set_column_widths(table, widths_inches):
    """Disables autofit-to-window and pins each column to an explicit
    width, sized to what that column actually holds (Phone/Email columns
    get more room than Name) -- set on both the table's column AND every
    cell in it, since Word sometimes honors a cell's own tcW over the
    shared tblGrid otherwise."""
    from docx.shared import Inches
    from docx.oxml.ns import qn

    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = tblPr.makeelement(qn("w:tblLayout"), {qn("w:type"): "fixed"})
    tblPr.append(layout)

    for col_idx, width in enumerate(widths_inches):
        table.columns[col_idx].width = Inches(width)
        for row in table.rows:
            row.cells[col_idx].width = Inches(width)


def _export_employees_xlsx(employees, title="Employee Directory"):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    utility_head_ids = compute_utility_head_ids()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title[:31]

    headers = ["Name", "Designation", "Organization", "Department",
               "Office Phone", "Mobile", "Email", "Utility Head"]
    header_font   = Font(bold=True, color="FFFFFF")
    header_fill   = PatternFill("solid", fgColor="1a3a6b")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row, emp in enumerate(employees, 2):
        ws.cell(row=row, column=1, value=emp.employee_name)
        ws.cell(row=row, column=2, value=emp.designation or "")
        ws.cell(row=row, column=3,
                value=emp.organization.organization_name if emp.organization else "")
        ws.cell(row=row, column=4,
                value=emp.department.department_name if emp.department else "")
        ws.cell(row=row, column=5, value=emp.office_phone or "")
        ws.cell(row=row, column=6, value=emp.mobile_phone or "")
        ws.cell(row=row, column=7, value=emp.email or "")
        ws.cell(row=row, column=8, value="Yes" if emp.id in utility_head_ids else "")

    widths = [26, 24, 32, 18, 20, 16, 30, 12]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}{max(len(employees) + 1, 1)}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="employees.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _export_employees_docx(employees, title="Employee Directory"):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    utility_head_ids = compute_utility_head_ids()

    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)

    table = doc.add_table(rows=1, cols=7)
    table.style = "Light List Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(
        ["Name", "Designation", "Organization", "Office Phone", "Mobile", "Email", "Utility Head"]
    ):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for emp in employees:
        row = table.add_row().cells
        row[0].text = emp.employee_name or ""
        row[1].text = emp.designation or ""
        row[2].text = emp.organization.organization_name if emp.organization else ""
        row[3].text = _wrappable(emp.office_phone or "")
        row[4].text = _wrappable(emp.mobile_phone or "")
        row[5].text = _wrappable(emp.email or "")
        row[6].text = "Yes" if emp.id in utility_head_ids else ""

    _set_column_widths(table, [1.0, 0.9, 1.1, 0.8, 0.8, 1.2, 0.5])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="employees.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


def _export_org_docx(org_name, address, employees, control_rooms):
    from docx import Document
    from docx.shared import RGBColor

    doc = Document()
    heading = doc.add_heading(org_name, level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)
    if address:
        doc.add_paragraph(address)

    if employees:
        doc.add_heading("Employees", level=2)
        tbl = doc.add_table(rows=1, cols=5)
        tbl.style = "Light List Accent 1"
        hdr = tbl.rows[0].cells
        for i, h in enumerate(["Name", "Designation", "Office Phone", "Mobile", "Email"]):
            hdr[i].text = h
            hdr[i].paragraphs[0].runs[0].bold = True
        for emp in employees:
            r = tbl.add_row().cells
            r[0].text = emp.employee_name or ""
            r[1].text = emp.designation or ""
            r[2].text = _wrappable(emp.office_phone or "")
            r[3].text = _wrappable(emp.mobile_phone or "")
            r[4].text = _wrappable(emp.email or "")

        _set_column_widths(tbl, [1.3, 1.2, 1.1, 1.1, 1.8])

    if control_rooms:
        doc.add_heading("Control Room", level=2)
        tbl2 = doc.add_table(rows=1, cols=3)
        tbl2.style = "Light List Accent 1"
        hdr2 = tbl2.rows[0].cells
        for i, h in enumerate(["Name", "Phone Number", "Email"]):
            hdr2[i].text = h
            hdr2[i].paragraphs[0].runs[0].bold = True
        for cr in control_rooms:
            r = tbl2.add_row().cells
            r[0].text = cr.name or ""
            r[1].text = _wrappable(cr.phone_number or "")
            r[2].text = _wrappable(cr.email or "")

        _set_column_widths(tbl2, [1.7, 2.3, 2.5])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in org_name)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{safe_name}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@user_bp.route("/export/utility-heads")
def export_utility_heads():
    fmt = request.args.get("format", "xlsx")
    heads = _compute_utility_heads()
    if fmt == "docx":
        return _export_utility_heads_docx(heads)
    return _export_utility_heads_xlsx(heads)


def _export_utility_heads_xlsx(heads):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Utility Heads"

    headers = ["#", "Name", "Designation", "Organization", "Address",
               "Office Phone", "Mobile", "Email"]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a3a6b")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row, emp in enumerate(heads, 2):
        ws.cell(row=row, column=1, value=row - 1)
        ws.cell(row=row, column=2, value=emp.employee_name or "")
        ws.cell(row=row, column=3, value=emp.designation or "")
        ws.cell(row=row, column=4,
                value=emp.organization.organization_name if emp.organization else "")
        ws.cell(row=row, column=5,
                value=(emp.organization.address or "") if emp.organization else "")
        ws.cell(row=row, column=6, value=emp.office_phone or "")
        ws.cell(row=row, column=7, value=emp.mobile_phone or "")
        ws.cell(row=row, column=8, value=emp.email or "")

    col_widths = [5, 28, 30, 45, 55, 22, 18, 30]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(i)
        ].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="utility_heads.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _export_utility_heads_docx(heads):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    heading = doc.add_heading("Utility Heads – Western Region", level=1)
    heading.runs[0].font.color.rgb = RGBColor(0x1a, 0x3a, 0x6b)

    table = doc.add_table(rows=1, cols=7)
    table.style = "Light List Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(
        ["Name", "Designation", "Organization", "Address",
         "Office Phone", "Mobile", "Email"]
    ):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True

    for emp in heads:
        row = table.add_row().cells
        row[0].text = emp.employee_name or ""
        row[1].text = emp.designation or ""
        row[2].text = emp.organization.organization_name if emp.organization else ""
        row[3].text = (emp.organization.address or "") if emp.organization else ""
        row[4].text = emp.office_phone or ""
        row[5].text = emp.mobile_phone or ""
        row[6].text = emp.email or ""

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="utility_heads.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ─── ADMINISTRATIVE HEADS ────────────────────────────────────────────────────
# Senior government officials (IAS/IPS/IFS, Energy Secretaries, Collectors,
# ...) -- a distinct table-backed identity, never mixed with Utility Heads
# (Employee.is_utility_head / compute_utility_head_ids) or KMP (is_kmp).

ADMINISTRATIVE_HEAD_STATUSES = ["ACTIVE", "ON_LEAVE", "VACANT", "INACTIVE"]


def _get_administrative_heads_filters():
    return {
        "keyword": request.args.get("keyword", "").strip(),
        "category_id": request.args.get("category_id", type=int),
        "state": request.args.get("state", "").strip(),
        "organization_id": request.args.get("organization_id", type=int),
        "service_type_id": request.args.get("service_type_id", type=int),
        "designation": request.args.get("designation", "").strip(),
        "status": request.args.get("status", "").strip(),
    }


def _administrative_heads_filter_options():
    return {
        "categories": OrganizationCategory.query.order_by(OrganizationCategory.category_name).all(),
        "organizations": Organization.query.order_by(Organization.organization_name).all(),
        "service_types": ServiceType.query.order_by(ServiceType.service_type_name).all(),
        "statuses": ADMINISTRATIVE_HEAD_STATUSES,
    }


@user_bp.route("/administrative-heads")
def administrative_heads():
    filters = _get_administrative_heads_filters()
    heads = head_service.search_administrative_heads(
        keyword=filters["keyword"] or None,
        designation=filters["designation"] or None,
        category_id=filters["category_id"],
        state=filters["state"] or None,
        organization_id=filters["organization_id"],
        service_type_id=filters["service_type_id"],
        status=filters["status"] or None,
    )
    emails = dedupe_and_sort(heads)
    return render_template(
        "administrative_heads.html", heads=heads, filters=filters,
        emails=emails, mailto_limit=MAILTO_SAFE_LIMIT,
        **_administrative_heads_filter_options(),
    )


@user_bp.route("/administrative-heads/<int:id>")
def administrative_head_details(id):
    head = db.session.get(AdministrativeHead, id)
    if not head or head.role_category != "ADMINISTRATIVE_HEAD":
        _abort404()
    return render_template("administrative_head_details.html", head=head)


@user_bp.route("/export/administrative-heads")
def export_administrative_heads():
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    fmt = request.args.get("format", "xlsx")
    filters = _get_administrative_heads_filters()
    heads = head_service.search_administrative_heads(
        keyword=filters["keyword"] or None,
        designation=filters["designation"] or None,
        state=filters["state"] or None,
        organization_id=filters["organization_id"],
        service_type_id=filters["service_type_id"],
        status=filters["status"] or None,
    )

    headers = ["Organization", "Name", "Designation", "Service Type", "Status", "Email", "Mobile", "Office Phone"]

    def row_values(h):
        return [
            h.organization.organization_name if h.organization else "",
            h.resolved_name,
            h.role_title,
            h.service_type.service_type_name if h.service_type else "",
            h.status,
            h.resolved_email or "",
            h.resolved_mobile_phone or "",
            h.resolved_office_phone or "",
        ]

    if fmt == "csv":
        import csv
        from flask import Response
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(headers)
        for h in heads:
            writer.writerow(row_values(h))
        return Response(
            buf.getvalue(), mimetype="text/csv",
            headers={"Content-Disposition": "attachment; filename=administrative_heads.csv"},
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Administrative Heads"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a3a6b")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font, cell.fill, cell.alignment = header_font, header_fill, Alignment(horizontal="center")

    for row, head in enumerate(heads, 2):
        for col, value in enumerate(row_values(head), 1):
            ws.cell(row=row, column=col, value=value)

    for i, w in enumerate([35, 28, 28, 18, 14, 32, 18, 18], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name="administrative_heads.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─── EMAIL DISTRIBUTION LISTS ────────────────────────────────────────────────

def _get_standard_group_or_404(slug):
    group = STANDARD_GROUPS_BY_SLUG.get(slug)
    if not group:
        _abort404()
    return group


def _get_contact_type():
    contact_type = request.args.get("contact_type", "all")
    return contact_type if contact_type in dict(CONTACT_TYPE_CHOICES) else "all"


@user_bp.route("/email-lists")
def email_distribution_home():
    category_name_by_slug = {slug: category_name for slug, category_name, _ in CATEGORY_SLUGS}
    sections = {}
    for group in STANDARD_GROUPS:
        if group["section"] == "Role Based":
            continue  # reachable at /email-lists/<slug> directly, just not as a home-page card
        rows = group["resolver"]()
        is_org_category_card = group["slug"] in category_name_by_slug
        if is_org_category_card and not rows:
            continue  # empty category -- same "no card if nothing in it" rule the Telephone Directory landing page uses
        emails = dedupe_and_sort(rows)
        card = {
            "slug": group["slug"],
            "name": group["name"],
            "member_count": len(rows),
            "email_count": len(emails),
            "category_name": category_name_by_slug.get(group["slug"]),
        }
        sections.setdefault(group["section"], []).append(card)

    # Admin-created custom (DYNAMIC) groups -- e.g. "ADANI" -- previously
    # only reachable from the admin panel; surfaced here too since they
    # resolve from the same public Employee/Organization data every other
    # card on this page already exposes.
    custom_groups = EmailGroup.query.filter_by(list_type="DYNAMIC").order_by(EmailGroup.name).all()
    custom_cards = []
    for group in custom_groups:
        rows = resolve_dynamic_group(group)
        emails = dedupe_and_sort(rows)
        custom_cards.append({
            "id": group.id, "name": group.name,
            "member_count": len(rows), "email_count": len(emails),
        })
    if custom_cards:
        sections["Custom Groups"] = custom_cards

    return render_template("email_distribution_home.html", sections=sections)


@user_bp.route("/email-lists/custom/<int:id>")
def email_distribution_custom_group(id):
    group = db.session.get(EmailGroup, id)
    if not group or group.list_type != "DYNAMIC":
        _abort404()
    rows = resolve_dynamic_group(group)
    contact_rows = [to_contact_row(r) for r in rows]
    contact_rows.sort(key=lambda r: (r["organization"].lower(), r["name"].lower()))
    emails = dedupe_and_sort(rows)
    return render_template(
        "email_distribution_custom_group.html",
        group=group, contact_rows=contact_rows, emails=emails,
        mailto_limit=MAILTO_SAFE_LIMIT,
    )


@user_bp.route("/email-lists/<slug>")
def email_distribution_group(slug):
    group = _get_standard_group_or_404(slug)
    contact_type = _get_contact_type()
    rows = group["resolver"](contact_type=contact_type)
    emails = dedupe_and_sort(rows)
    contact_rows = [to_contact_row(r) for r in rows]
    contact_rows.sort(key=lambda r: (r["organization"].lower(), r["name"].lower()))
    summary = summarize_contacts(rows)
    return render_template(
        "email_distribution_group.html",
        group=group, rows=rows, emails=emails,
        contact_rows=contact_rows, summary=summary,
        mailto_limit=MAILTO_SAFE_LIMIT,
        contact_type=contact_type, contact_type_choices=CONTACT_TYPE_CHOICES,
    )


@user_bp.route("/email-lists/<slug>/emails.json")
def email_distribution_emails_json(slug):
    group = _get_standard_group_or_404(slug)
    contact_type = _get_contact_type()
    emails = dedupe_and_sort(group["resolver"](contact_type=contact_type))
    mailto = build_mailto(emails, subject=group["name"])
    return jsonify({
        "name": group["name"],
        "emails": emails,
        "email_count": len(emails),
        "mailto": mailto,
        "mailto_safe": len(mailto) <= MAILTO_SAFE_LIMIT,
        "mailto_limit": MAILTO_SAFE_LIMIT,
    })


@user_bp.route("/email-lists/<slug>/export.xlsx")
def email_distribution_export_xlsx(slug):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    group = _get_standard_group_or_404(slug)
    contact_type = _get_contact_type()
    emails = dedupe_and_sort(group["resolver"](contact_type=contact_type))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = group["name"][:31]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a3a6b")
    cell = ws.cell(row=1, column=1, value="Email")
    cell.font, cell.fill, cell.alignment = header_font, header_fill, Alignment(horizontal="center")

    for row, email in enumerate(emails, 2):
        ws.cell(row=row, column=1, value=email)
    ws.column_dimensions["A"].width = 40

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in group["name"])
    return send_file(
        buf, as_attachment=True,
        download_name=f"{safe_name}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@user_bp.route("/email-lists/<slug>/export.csv")
def email_distribution_export_csv(slug):
    import csv
    from flask import Response

    group = _get_standard_group_or_404(slug)
    contact_type = _get_contact_type()
    emails = dedupe_and_sort(group["resolver"](contact_type=contact_type))

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["Email"])
    for email in emails:
        writer.writerow([email])

    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in group["name"])
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={safe_name}.csv"},
    )


# ─── EMAIL DISTRIBUTION LISTS -- ADDRESS BOOK BROWSE VIEW ────────────────────
# Organization Category + Contact Type, chained -- see
# services/email_distribution_service.py:resolve_category_contacts.

BROWSE_CATEGORY_CHOICES = [(category_name, plural) for _, category_name, plural in CATEGORY_SLUGS]
BROWSE_CATEGORY_NAMES = {category_name for _, category_name, _ in CATEGORY_SLUGS}


def _get_browse_category():
    category = request.args.get("category", "")
    return category if category in BROWSE_CATEGORY_NAMES else None


def _get_browse_rows():
    """Returns (category_name_or_None, rows) -- rows is [] when no (valid)
    category has been chosen yet."""
    category = _get_browse_category()
    if not category:
        return None, []
    state = request.args.get("state", "").strip()
    return category, resolve_category_contacts(category, _get_contact_type(), state=state or None)


def _browse_contact_rows(q=""):
    category, rows = _get_browse_rows()
    contact_rows = [to_contact_row(r) for r in rows]
    contact_rows = filter_contact_rows(contact_rows, q)
    contact_rows.sort(key=lambda r: (r["organization"].lower(), r["name"].lower()))
    return category, rows, contact_rows


@user_bp.route("/email-lists/browse")
def email_distribution_browse():
    # Category is fixed by whichever card on the Email Distribution Lists
    # home page was clicked (?category=<name>) -- there is no other entry
    # point to this page, so the category is shown as read-only context
    # here rather than as another dropdown to re-pick.
    categories_by_name = {c.category_name: c for c in OrganizationCategory.query.all()}
    plural_by_name = {category_name: plural for _, category_name, plural in CATEGORY_SLUGS}

    category_arg = _get_browse_category()
    category = None
    if category_arg and category_arg in categories_by_name:
        category = {
            "name": category_arg,
            "label": plural_by_name.get(category_arg, category_arg),
            "id": categories_by_name[category_arg].id,
            "is_state_based": categories_by_name[category_arg].is_state_based,
        }

    return render_template(
        "email_distribution_browse.html",
        category=category,
        contact_type_choices=CONTACT_TYPE_CHOICES,
        mailto_limit=MAILTO_SAFE_LIMIT,
    )


@user_bp.route("/email-lists/browse/data.json")
def email_distribution_browse_data_json():
    category, rows = _get_browse_rows()
    if not category:
        return jsonify({"summary": None, "rows": []})

    # Summary reflects the full category + contact-type result, not the
    # live search box -- search only narrows the grid client-side.
    summary = summarize_contacts(rows)
    contact_rows = [to_contact_row(r) for r in rows]
    contact_rows.sort(key=lambda r: (r["organization"].lower(), r["name"].lower()))
    contact_rows = filter_contact_rows(contact_rows, request.args.get("q", ""))

    return jsonify({"summary": summary, "rows": contact_rows})


@user_bp.route("/email-lists/browse/export.xlsx")
def email_distribution_browse_export_xlsx():
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    category, _rows, contact_rows = _browse_contact_rows(request.args.get("q", ""))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contacts"

    headers = ["Organization", "Name", "Role / Designation", "Email", "Mobile"]
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a3a6b")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font, cell.fill, cell.alignment = header_font, header_fill, Alignment(horizontal="center")

    for row, contact in enumerate(contact_rows, 2):
        ws.cell(row=row, column=1, value=contact["organization"])
        ws.cell(row=row, column=2, value=contact["name"])
        ws.cell(row=row, column=3, value=contact["designation"])
        ws.cell(row=row, column=4, value=contact["email"])
        ws.cell(row=row, column=5, value=contact["mobile"])

    for i, w in enumerate([35, 28, 28, 32, 18], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in (category or "contacts"))
    return send_file(
        buf, as_attachment=True,
        download_name=f"{safe_name}_address_book.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@user_bp.route("/email-lists/browse/export.csv")
def email_distribution_browse_export_csv():
    import csv
    from flask import Response

    category, _rows, contact_rows = _browse_contact_rows(request.args.get("q", ""))

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["Organization", "Name", "Role / Designation", "Email", "Mobile"])
    for c in contact_rows:
        writer.writerow([c["organization"], c["name"], c["designation"], c["email"], c["mobile"]])

    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in (category or "contacts"))
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={safe_name}_address_book.csv"},
    )


# ─── DIRECTORY VERSIONS ────────────────────────────────────────────────────────
# Official monthly snapshots of the telephone directory (PDF + Excel), stored
# on disk with metadata in directory_versions -- see
# services/directory_version_service.py. No employee data is duplicated here.

@user_bp.route("/directory-versions")
def directory_versions_list():
    f = {
        "month": request.args.get("month", type=int),
        "year": request.args.get("year", type=int),
        "keyword": request.args.get("keyword", "").strip(),
        "page": request.args.get("page", 1, type=int),
    }
    versions, total, page, per_page = directory_version_service.search_versions(
        month=f["month"], year=f["year"], keyword=f["keyword"] or None, page=f["page"],
    )
    return render_template(
        "directory_versions_list.html", versions=versions, filters=f,
        total=total, page=page, per_page=per_page,
        stats=directory_version_service.summary_stats(),
        years=directory_version_service.distinct_years(),
    )


@user_bp.route("/directory-versions/<int:id>/preview")
def directory_version_preview(id):
    version = db.session.get(DirectoryVersion, id) or _abort404()
    return send_file(version.pdf_path, mimetype="application/pdf",
                      as_attachment=False, download_name=version.pdf_filename)


@user_bp.route("/directory-versions/<int:id>/download/pdf")
def directory_version_download_pdf(id):
    version = db.session.get(DirectoryVersion, id) or _abort404()
    return send_file(version.pdf_path, mimetype="application/pdf",
                      as_attachment=True, download_name=version.pdf_filename)


@user_bp.route("/directory-versions/<int:id>/download/excel")
def directory_version_download_excel(id):
    version = db.session.get(DirectoryVersion, id) or _abort404()
    return send_file(
        version.excel_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=version.excel_filename,
    )


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def _abort404():
    from flask import abort
    abort(404)
