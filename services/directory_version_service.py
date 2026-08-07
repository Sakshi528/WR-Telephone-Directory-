"""Query + generation layer for the Directory Versions module -- official
monthly snapshots of the telephone directory (PDF + Excel), stored on disk
with metadata in directory_versions. No employee data is duplicated; every
call re-reads the live tables fresh. Regeneration never mutates an existing
row or its files -- see next_version_number.
"""

import os

from flask import current_app

from models import db
from models.employee import Employee
from models.organization import Organization
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from models.directory_version import DirectoryVersion
from services import head_service
from services.audit_service import log_audit_event
from services.pdf_generator import generate_pdf
from services.excel_generator import generate_excel
from utils.designation_rank import compute_utility_head_ids
from utils.logger import get_logger

logger = get_logger(__name__)

DEFAULT_PAGE_SIZE = 25

MONTH_NAMES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _paginate(query, page, per_page):
    page = max(int(page or 1), 1)
    per_page = max(int(per_page or DEFAULT_PAGE_SIZE), 1)
    total = query.count()
    items = query.offset((page - 1) * per_page).limit(per_page).all()
    return items, total, page, per_page


def next_version_number(year, month):
    """'YYYY.MM' if no row exists yet for that (year, month); otherwise
    'YYYY.MM.N' for the next unused revision -- used by regenerate_version
    so a prior snapshot's row/files are never touched."""
    base = f"{year}.{month:02d}"
    existing = {
        v.version_number for v in
        DirectoryVersion.query.filter_by(year=year, month=month).all()
    }
    if base not in existing:
        return base
    n = 2
    while f"{base}.{n}" in existing:
        n += 1
    return f"{base}.{n}"


def build_directory_snapshot():
    """Re-queries the same grouping shape as
    routes.user_routes.telephone_directory()'s no-keyword branch (org-wise
    employees + control rooms/switchyards + address), plus the extra
    role-based sections a directory document needs. Deliberately
    re-implemented rather than imported -- that route is out of scope for
    this module and must not risk regression."""
    all_orgs = Organization.query.all()
    org_address = {o.organization_name: (o.address or "") for o in all_orgs}

    utility_head_ids = compute_utility_head_ids()
    all_employees = (
        Employee.query
        .outerjoin(Organization)
        .filter(Employee.is_kmp == True)  # noqa: E712
        .filter(Employee.status == "ACTIVE")
        .order_by(Organization.organization_name, Employee.employee_name)
        .all()
    )
    all_employees.sort(key=lambda e: e.id not in utility_head_ids)

    all_numbers = (
        DirectoryNumber.query
        .filter(DirectoryNumber.organization != None)  # noqa: E711
        .order_by(DirectoryNumber.organization)
        .all()
    )

    grouped = {}
    for emp in all_employees:
        org_name = emp.organization.organization_name if emp.organization else "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [], "control_rooms": [], "switchyards": [],
                "address": emp.organization.address if emp.organization else "",
            }
        grouped[org_name]["employees"].append(emp)

    other_numbers = []
    for dn in all_numbers:
        org_name = dn.organization or "Unknown"
        if org_name not in grouped:
            grouped[org_name] = {
                "employees": [], "control_rooms": [], "switchyards": [],
                "address": org_address.get(org_name, ""),
            }
        if dn.category and "switchyard" in dn.category.lower():
            grouped[org_name]["switchyards"].append(dn)
        elif dn.category and "control room" in dn.category.lower():
            grouped[org_name]["control_rooms"].append(dn)
        else:
            other_numbers.append(dn)

    utility_heads = (
        Employee.query.outerjoin(Organization)
        .filter(Employee.id.in_(utility_head_ids))
        .order_by(Organization.organization_name)
        .all()
    )
    administrative_heads = head_service.search_administrative_heads(status="ACTIVE")
    kmp = (
        Employee.query.outerjoin(Organization)
        .filter(Employee.is_kmp == True, Employee.status == "ACTIVE")  # noqa: E712
        .order_by(Organization.organization_name, Employee.employee_name)
        .all()
    )
    emergency_contacts = (
        EmergencyContact.query.join(Employee)
        .filter(Employee.status == "ACTIVE")
        .order_by(Employee.employee_name)
        .all()
    )

    return {
        "grouped": dict(sorted(grouped.items())),
        "other_numbers": other_numbers,
        "utility_heads": utility_heads,
        "administrative_heads": administrative_heads,
        "kmp": kmp,
        "emergency_contacts": emergency_contacts,
        "employee_count": Employee.query.filter(Employee.status == "ACTIVE").count(),
        "organization_count": Organization.query.count(),
    }


def _write_version_files(version_number, year, month, pdf_bytes, excel_bytes):
    folder = os.path.join(current_app.config["DIRECTORY_VERSIONS_STORAGE"], str(year), f"{month:02d}")
    os.makedirs(folder, exist_ok=True)

    stem = f"WR_Telephone_Directory_{version_number.replace('.', '_')}"
    pdf_filename, excel_filename = f"{stem}.pdf", f"{stem}.xlsx"
    pdf_path = os.path.join(folder, pdf_filename)
    excel_path = os.path.join(folder, excel_filename)

    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
    with open(excel_path, "wb") as f:
        f.write(excel_bytes)

    return pdf_filename, pdf_path, excel_filename, excel_path


def _create_version(version_number, month, year, remarks, created_by, action):
    version_name = f"WR Telephone Directory - {MONTH_NAMES[month]} {year}"
    version = DirectoryVersion(
        version_number=version_number, version_name=version_name,
        month=month, year=year, remarks=remarks or None,
        generated_by=created_by.id if created_by else None,
    )

    snapshot = build_directory_snapshot()
    version.employee_count = snapshot["employee_count"]
    version.organization_count = snapshot["organization_count"]

    pdf_bytes = generate_pdf(version, snapshot)
    excel_bytes = generate_excel(version, snapshot)
    pdf_filename, pdf_path, excel_filename, excel_path = _write_version_files(
        version_number, year, month, pdf_bytes, excel_bytes,
    )
    version.pdf_filename, version.pdf_path = pdf_filename, pdf_path
    version.excel_filename, version.excel_path = excel_filename, excel_path

    db.session.add(version)
    db.session.flush()  # assigns version.id for the audit log record_id

    log_audit_event(
        module="directory_versions", record_type="DirectoryVersion", record_id=version.id,
        action=action, field_name="version_number",
        old_value=None, new_value=version.version_number, reason=remarks or None,
    )
    logger.info("%s directory version %s (%s employees, %s organizations)",
                action, version.version_number, version.employee_count, version.organization_count)
    return version


def generate_version(month, year, remarks, created_by):
    """Raises ValueError if a version already exists for (year, month) --
    use regenerate_version for that case instead."""
    existing = DirectoryVersion.query.filter_by(year=year, month=month).first()
    if existing:
        raise ValueError(
            f"A Directory Version already exists for {MONTH_NAMES[month]} {year} "
            f"({existing.version_number}). Use Regenerate instead."
        )
    return _create_version(f"{year}.{month:02d}", month, year, remarks, created_by, "GENERATE")


def regenerate_version(year, month, remarks, created_by):
    """Always allowed -- inserts a new row with a revision-suffixed
    version_number; the prior row and its files are never touched."""
    version_number = next_version_number(year, month)
    return _create_version(version_number, month, year, remarks, created_by, "REGENERATE")


def delete_version(version, deleted_by):
    for path in (version.pdf_path, version.excel_path):
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                logger.warning("Could not remove file %s for version %s", path, version.version_number)
        elif path:
            logger.warning("File %s for version %s was already missing", path, version.version_number)

    log_audit_event(
        module="directory_versions", record_type="DirectoryVersion", record_id=version.id,
        action="DELETE", field_name="version_number",
        old_value=version.version_number, new_value=None,
    )
    db.session.delete(version)


def update_remarks(version, remarks, updated_by):
    old_remarks = version.remarks
    version.remarks = remarks or None
    log_audit_event(
        module="directory_versions", record_type="DirectoryVersion", record_id=version.id,
        action="UPDATE", field_name="remarks",
        old_value=old_remarks, new_value=version.remarks,
    )


def search_versions(month=None, year=None, keyword=None, page=1, per_page=DEFAULT_PAGE_SIZE):
    query = DirectoryVersion.query
    if month:
        query = query.filter(DirectoryVersion.month == month)
    if year:
        query = query.filter(DirectoryVersion.year == year)
    if keyword:
        kw = f"%{keyword}%"
        query = query.filter(db.or_(
            DirectoryVersion.version_number.ilike(kw),
            DirectoryVersion.version_name.ilike(kw),
            DirectoryVersion.remarks.ilike(kw),
        ))
    query = query.order_by(DirectoryVersion.generated_on.desc())
    return _paginate(query, page, per_page)


def distinct_years():
    rows = db.session.query(DirectoryVersion.year).distinct().order_by(DirectoryVersion.year.desc()).all()
    return [r[0] for r in rows]


def summary_stats():
    total = DirectoryVersion.query.count()
    latest = DirectoryVersion.query.order_by(DirectoryVersion.generated_on.desc()).first()
    return {
        "total_versions": total,
        "latest_version_number": latest.version_number if latest else "-",
        "last_generated_on": latest.generated_on if latest else None,
        "published_versions": DirectoryVersion.query.filter_by(status="Published").count(),
    }
