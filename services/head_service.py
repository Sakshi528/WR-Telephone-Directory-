"""Administrative Head assignment lifecycle + PA/PS assistant management --
the module models/administrative_head.py and
models/administrative_head_history.py already anticipated by name in their
docstrings. Every mutation here calls services/audit_service.log_audit_event
with module="administrative_heads".
"""

from datetime import date

from models import db
from models.employee import Employee
from models.organization import Organization
from models.administrative_head import AdministrativeHead
from models.administrative_head_history import AdministrativeHeadHistory
from models.administrative_head_assistant import AdministrativeHeadAssistant
from services.audit_service import log_audit_event


# ─── ADMINISTRATIVE HEAD ASSIGNMENT ──────────────────────────────────────────

def assign_administrative_head(
    organization_id, role_title, employee_id=None, name=None,
    office_phone=None, mobile_phone=None, email=None, office_address=None,
    service_type_id=None, status="ACTIVE", effective_from=None, remarks=None,
):
    """employee_id or name (or both) must be given -- enforced again here,
    not just by the DB CHECK constraint, so callers get a clear ValueError
    instead of an IntegrityError from a half-built request."""
    if not employee_id and not (name or "").strip():
        raise ValueError("Either employee_id or a name is required.")

    head = AdministrativeHead(
        employee_id=employee_id,
        organization_id=organization_id,
        role_category="ADMINISTRATIVE_HEAD",
        role_title=role_title,
        service_type_id=service_type_id,
        status=status or "ACTIVE",
        effective_from=effective_from or date.today(),
        remarks=remarks,
        name=None if employee_id else (name or None),
        office_phone=None if employee_id else office_phone,
        mobile_phone=None if employee_id else mobile_phone,
        email=None if employee_id else email,
        office_address=office_address,
    )
    db.session.add(head)
    db.session.flush()
    log_audit_event(
        module="administrative_heads", record_type="AdministrativeHead", record_id=head.id,
        action="CREATE", field_name="role_title", old_value=None, new_value=role_title,
    )
    return head


def update_administrative_head(
    head, organization_id, role_title, service_type_id, remarks,
    employee_id=None, name=None, office_phone=None, mobile_phone=None,
    email=None, office_address=None, status=None,
):
    if not employee_id and not (name or "").strip():
        raise ValueError("Either employee_id or a name is required.")

    if role_title != head.role_title:
        log_audit_event(
            module="administrative_heads", record_type="AdministrativeHead", record_id=head.id,
            action="UPDATE", field_name="role_title", old_value=head.role_title, new_value=role_title,
        )
    if status and status != head.status:
        log_audit_event(
            module="administrative_heads", record_type="AdministrativeHead", record_id=head.id,
            action="UPDATE", field_name="status", old_value=head.status, new_value=status,
        )
        head.status = status

    head.organization_id = organization_id
    head.role_title = role_title
    head.service_type_id = service_type_id
    head.remarks = remarks
    head.employee_id = employee_id
    head.name = None if employee_id else (name or None)
    head.office_phone = None if employee_id else office_phone
    head.mobile_phone = None if employee_id else mobile_phone
    head.email = None if employee_id else email
    head.office_address = office_address
    return head


def end_role(head, reason, effective_to=None, replacement_employee_id=None):
    """Archives head to AdministrativeHeadHistory (matching the table's own
    docstring), then deletes it -- and its assistants, via cascade -- from
    the live table. Assistants aren't carried into history; they belong to
    the live assignment, not a point-in-time record of who held the role."""
    history = AdministrativeHeadHistory(
        organization_id=head.organization_id,
        employee_id=head.employee_id,
        role_category=head.role_category,
        role_title=head.role_title,
        effective_from=head.effective_from,
        effective_to=effective_to or date.today(),
        reason=reason,
        replacement_employee_id=replacement_employee_id,
        remarks=head.remarks,
        name=head.name,
        office_phone=head.office_phone,
        mobile_phone=head.mobile_phone,
        email=head.email,
        office_address=head.office_address,
        status=head.status,
        service_type_id=head.service_type_id,
    )
    db.session.add(history)
    log_audit_event(
        module="administrative_heads", record_type="AdministrativeHead", record_id=head.id,
        action="END_ROLE", field_name="role_title", old_value=head.role_title, new_value=None,
        reason=reason,
    )
    db.session.delete(head)
    return history


# ─── PA/PS ASSISTANTS ─────────────────────────────────────────────────────────

def _clear_primary(administrative_head_id):
    AdministrativeHeadAssistant.query.filter_by(
        administrative_head_id=administrative_head_id, is_primary=True
    ).update({"is_primary": False})


def add_assistant(
    administrative_head_id, designation, employee_id=None, name=None,
    office_phone=None, mobile=None, email=None, is_primary=False, created_by=None,
):
    if is_primary:
        _clear_primary(administrative_head_id)

    assistant = AdministrativeHeadAssistant(
        administrative_head_id=administrative_head_id,
        employee_id=employee_id,
        designation=designation,
        name=name,
        office_phone=office_phone,
        mobile=mobile,
        email=email,
        is_primary=is_primary,
        created_by=created_by,
    )
    db.session.add(assistant)
    db.session.flush()
    log_audit_event(
        module="administrative_heads", record_type="AdministrativeHeadAssistant", record_id=assistant.id,
        action="CREATE", field_name="designation", old_value=None, new_value=designation,
    )
    return assistant


def update_assistant(assistant, designation, employee_id=None, name=None,
                      office_phone=None, mobile=None, email=None):
    log_audit_event(
        module="administrative_heads", record_type="AdministrativeHeadAssistant", record_id=assistant.id,
        action="UPDATE", field_name="designation", old_value=assistant.designation, new_value=designation,
    )
    assistant.employee_id = employee_id
    assistant.designation = designation
    assistant.name = name
    assistant.office_phone = office_phone
    assistant.mobile = mobile
    assistant.email = email
    return assistant


def remove_assistant(assistant):
    log_audit_event(
        module="administrative_heads", record_type="AdministrativeHeadAssistant", record_id=assistant.id,
        action="DELETE", field_name="designation", old_value=assistant.designation, new_value=None,
    )
    db.session.delete(assistant)


def set_primary_assistant(administrative_head_id, assistant_id):
    _clear_primary(administrative_head_id)
    assistant = db.session.get(AdministrativeHeadAssistant, assistant_id)
    if assistant and assistant.administrative_head_id == administrative_head_id:
        assistant.is_primary = True
        log_audit_event(
            module="administrative_heads", record_type="AdministrativeHeadAssistant", record_id=assistant.id,
            action="SET_PRIMARY", field_name="is_primary", old_value="False", new_value="True",
        )
    return assistant


# ─── SEARCH / FILTER ──────────────────────────────────────────────────────────

def search_administrative_heads(
    keyword=None, designation=None, category_id=None,
    state=None, organization_id=None, service_type_id=None, status=None,
    subcategory=None,
):
    """Query builder behind both the list view and its search box.
    Organization Type/State/Organization/Service Type/Designation/Status
    are structured filters, ANDed together via SQL. `keyword` is the free-text search box
    -- it ORs across Administrative Head name, assistant name, and
    designation, done in Python since the result set is small and assistant
    matching needs each head's resolved assistant names anyway.

    Outer-joins Employee (a head may not have one -- see
    models/administrative_head.py) and reads names via resolved_name, not
    a direct .employee.employee_name access.

    A head linked to an employee who is no longer ACTIVE is excluded (the
    change_employee_status auto-archive hook should already have moved them
    to administrative_head_history, but this is a defensive second check,
    not the primary mechanism)."""
    query = (
        AdministrativeHead.query
        .outerjoin(Employee, AdministrativeHead.employee_id == Employee.id)
        .outerjoin(Organization, AdministrativeHead.organization_id == Organization.id)
        .filter(AdministrativeHead.role_category == "ADMINISTRATIVE_HEAD")
        .filter(db.or_(AdministrativeHead.employee_id.is_(None), Employee.status == "ACTIVE"))
    )

    if organization_id:
        query = query.filter(AdministrativeHead.organization_id == organization_id)
    if service_type_id:
        query = query.filter(AdministrativeHead.service_type_id == service_type_id)
    if status:
        query = query.filter(AdministrativeHead.status == status)
    if category_id:
        query = query.filter(Organization.category_id == category_id)
    if state:
        query = query.filter(Organization.state == state)
    if subcategory:
        query = query.filter(Organization.region == subcategory)
    if designation:
        query = query.filter(AdministrativeHead.role_title.ilike(f"%{designation}%"))

    heads = query.all()

    if keyword:
        kw = keyword.lower()
        heads = [
            h for h in heads
            if kw in (h.resolved_name or "").lower()
            or kw in (h.role_title or "").lower()
            or any(kw in (a.resolved_name or "").lower() for a in h.assistants)
        ]

    heads.sort(key=lambda h: (h.resolved_name or "").lower())
    return heads
