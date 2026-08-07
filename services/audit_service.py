"""Append-only audit logging for the enterprise modules (Administrative
Heads/KMP, Email Groups, Employee status, Update Request review). Called
explicitly at each mutation point -- not a generic model/session hook -- so
each call site controls exactly what's logged and why.

Usage:
    from services.audit_service import log_audit_event
    log_audit_event(
        module="employee_status", record_type="Employee", record_id=emp.id,
        action="STATUS_CHANGE", field_name="status",
        old_value="ACTIVE", new_value="RETIRED", reason=reason,
    )
"""

import uuid

from flask import has_request_context, request, session
from flask_login import current_user

from models import db
from models.audit_log import AuditLog
from utils.user_agent import parse_user_agent


def _request_metadata():
    """IP/browser/OS/session-id for the current request, or all-None
    outside a request context (e.g. a CLI script). Browser/OS: Werkzeug 3.x
    removed its bundled User-Agent parser (request.user_agent.browser/
    .platform are always None now), so utils.user_agent.parse_user_agent
    does this instead. Session ID: this app uses Flask's client-side
    cookie session with no server-side session store, so there's no
    natural stable ID to read -- one is generated on first use and kept in
    the session dict for the rest of that browser session, same pattern
    any app without a server-side session backend uses for audit trails."""
    if not has_request_context():
        return None, None, None, None

    ip_address = request.headers.get("X-Forwarded-For", request.remote_addr)
    browser, operating_system = parse_user_agent(request.headers.get("User-Agent"))

    session_id = session.get("_audit_sid")
    if not session_id:
        session_id = uuid.uuid4().hex
        session["_audit_sid"] = session_id

    return ip_address, browser, operating_system, session_id


def log_audit_event(
    module, record_type, record_id, action,
    field_name=None, old_value=None, new_value=None, reason=None
):
    changed_by = None
    changed_by_label = None
    if current_user and current_user.is_authenticated:
        changed_by = current_user.id
        changed_by_label = current_user.username or current_user.email

    ip_address, browser, operating_system, session_id = _request_metadata()

    entry = AuditLog(
        module=module,
        record_type=record_type,
        record_id=record_id,
        action=action,
        field_name=field_name,
        old_value=old_value,
        new_value=new_value,
        changed_by=changed_by,
        changed_by_label=changed_by_label,
        reason=reason,
        ip_address=ip_address,
        browser=browser,
        operating_system=operating_system,
        session_id=session_id,
    )
    db.session.add(entry)
    return entry
