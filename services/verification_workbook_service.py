"""Query + Excel-generation layer for the admin-only Export Verification
Workbook feature. Every sheet re-reads live PostgreSQL data at export time
-- the source Excel workbook used during import is never touched -- and
each sheet's query/ordering deliberately mirrors the corresponding page's
own default (unfiltered) view, so the export is a faithful snapshot of
what a user currently sees in the application:

  Organizations         -> templates/manage_organizations.html
  Employees              -> routes.user_routes.directory() (Employee Directory, is_kmp=False)
  Telephone Directory    -> routes.user_routes.telephone_directory() no-keyword branch
  Utility Heads           -> routes.user_routes._compute_utility_heads() with no filters
  Administrative Heads    -> services.head_service.search_administrative_heads() with no filters
  Control Rooms/Switchyards -> DirectoryNumber split by category, admin ordering

Queries are re-implemented here rather than imported from routes/ -- same
defensive rationale as services.directory_version_service.build_directory_snapshot:
those routes are out of scope for this module and must not risk regression.
"""

import io
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from models.employee import Employee
from models.organization import Organization
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from services import head_service
from utils.designation_rank import compute_utility_head_ids, resolve_utility_head

WORKBOOK_FILENAME = "WRLDC_Telephone_Directory_Verification.xlsx"

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1a3a6b")
TITLE_FONT = Font(bold=True, size=14, color="1a3a6b")
LABEL_FONT = Font(bold=True)


def _write_sheet(wb, title, headers, rows, col_widths=None):
    ws = wb.create_sheet(title[:31])
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    for r, row_values in enumerate(rows, 2):
        for c, value in enumerate(row_values, 1):
            ws.cell(row=r, column=c, value=value)
    widths = col_widths or [22] * len(headers)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    last_row = max(len(rows) + 1, 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    return ws


# ─── ORGANIZATIONS ────────────────────────────────────────────────────────────

def _collect_organizations():
    orgs = Organization.query.order_by(Organization.organization_name).all()

    status_rows = (
        Employee.query
        .with_entities(Employee.organization_id, Employee.status)
        .all()
    )
    status_by_org = {}
    for org_id, status in status_rows:
        counts = status_by_org.setdefault(org_id, {})
        counts[status] = counts.get(status, 0) + 1

    rows = []
    for i, org in enumerate(orgs, 1):
        counts = status_by_org.get(org.id, {})
        total = sum(counts.values())
        active = counts.get("ACTIVE", 0)
        rows.append([
            i,
            org.organization_name,
            org.region or "",
            org.category.category_name if org.category else "Uncategorized",
            org.address or "",
            total,
            active,
            total - active,
        ])
    return rows


# ─── EMPLOYEES (Employee Directory: internal WRLDC staff, is_kmp=False) ──────

def _collect_employees():
    employees = (
        Employee.query
        .filter(Employee.is_kmp == False)  # noqa: E712
        .filter(Employee.status == "ACTIVE")
        .order_by(Employee.employee_name)
        .all()
    )

    emp_ids = [e.id for e in employees]
    ec_map = {}
    if emp_ids:
        contacts = EmergencyContact.query.filter(
            EmergencyContact.employee_id.in_(emp_ids)
        ).all()
        for c in contacts:
            if c.employee_id not in ec_map:
                ec_map[c.employee_id] = c

    rows = []
    for i, e in enumerate(employees, 1):
        ec = ec_map.get(e.id)
        rows.append([
            i,
            e.employee_name,
            e.designation or "",
            e.office_phone or "",
            e.mobile_phone or "",
            e.email or "",
            ec.contact_name if ec else "",
            ec.relation if ec else "",
            ec.phone if ec else "",
        ])
    return rows


# ─── TELEPHONE DIRECTORY (all orgs, KMP employees + control rooms/switchyards) ─

def _collect_telephone_directory():
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

    rows = []
    for org_name, group in sorted(grouped.items()):
        address = group["address"]
        for emp in group["employees"]:
            rows.append([
                org_name, address, "Employee", emp.employee_name,
                emp.designation or "", emp.office_phone or "", emp.mobile_phone or "",
                emp.email or "", "Yes" if emp.id in utility_head_ids else "",
            ])
        for cr in group["control_rooms"]:
            rows.append([
                org_name, address, "Control Room", cr.name,
                "", cr.phone_number or "", "", cr.email or "", "",
            ])
        for sw in group["switchyards"]:
            rows.append([
                org_name, address, "Switchyard", sw.name,
                "", sw.phone_number or "", "", sw.email or "", "",
            ])

    for dn in other_numbers:
        rows.append([
            "(Other / Unassigned Numbers)", "", dn.category or "Directory Number",
            dn.name, "", dn.phone_number or "", "", dn.email or "", "",
        ])

    return rows


# ─── UTILITY HEADS (mirrors user_routes._compute_utility_heads(), no filters) ─

def _collect_utility_heads():
    organizations = Organization.query.order_by(Organization.organization_name).all()

    rows = []
    i = 0
    for org in organizations:
        employees = (
            Employee.query
            .filter_by(organization_id=org.id)
            .filter(Employee.status == "ACTIVE")
            .order_by(Employee.id)
            .all()
        )
        head = resolve_utility_head(employees)
        if head:
            i += 1
            rows.append([
                i, head.employee_name or "", head.designation or "",
                org.organization_name,
                org.address or "",
                head.office_phone or "", head.mobile_phone or "", head.email or "",
            ])
    return rows


# ─── ADMINISTRATIVE HEADS (mirrors head_service.search_administrative_heads(), no filters) ─

def _collect_administrative_heads():
    heads = head_service.search_administrative_heads()
    rows = []
    for h in heads:
        rows.append([
            h.organization.organization_name if h.organization else "",
            h.resolved_name,
            h.role_title,
            h.service_type.service_type_name if h.service_type else "",
            h.status,
            h.resolved_email or "",
            h.resolved_mobile_phone or "",
            h.resolved_office_phone or "",
        ])
    return rows


# ─── CONTROL ROOMS / SWITCHYARDS (mirrors admin.directory_numbers ordering) ──

def _collect_directory_numbers(category_keyword):
    numbers = (
        DirectoryNumber.query
        .order_by(DirectoryNumber.organization, DirectoryNumber.name)
        .all()
    )
    rows = []
    i = 0
    for n in numbers:
        if n.category and category_keyword in n.category.lower():
            i += 1
            rows.append([
                i, n.name, n.organization or "", n.phone_number or "",
                n.email or "", n.category or "",
            ])
    return rows


# ─── SUMMARY ──────────────────────────────────────────────────────────────────

def _write_summary_sheet(wb, counts, generated_at, generated_by):
    ws = wb.create_sheet("Summary", 0)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 40

    ws.cell(row=1, column=1, value="WRLDC Telephone Directory").font = TITLE_FONT
    ws.cell(row=2, column=1, value="Verification Workbook").font = Font(bold=True, size=12, color="444444")

    r = 4
    ws.cell(row=r, column=1, value="Source").font = LABEL_FONT
    ws.cell(row=r, column=2, value="Live PostgreSQL Database (telephone_directory)")
    r += 1
    ws.cell(row=r, column=1, value="Exported At").font = LABEL_FONT
    ws.cell(row=r, column=2, value=generated_at.strftime("%Y-%m-%d %H:%M:%S"))
    r += 1
    ws.cell(row=r, column=1, value="Exported By").font = LABEL_FONT
    ws.cell(row=r, column=2, value=generated_by or "Admin")
    r += 2

    ws.cell(row=r, column=1, value="Record Counts").font = Font(bold=True, size=12, color="1a3a6b")
    r += 1
    header_row = r
    ws.cell(row=header_row, column=1, value="Sheet").font = HEADER_FONT
    ws.cell(row=header_row, column=1).fill = HEADER_FILL
    ws.cell(row=header_row, column=2, value="Records").font = HEADER_FONT
    ws.cell(row=header_row, column=2).fill = HEADER_FILL
    r += 1
    for label, count in counts:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=count)
        r += 1

    return ws


# ─── PUBLIC ENTRY POINT ───────────────────────────────────────────────────────

def generate_verification_workbook(generated_by=None):
    """Builds the full 8-sheet verification workbook from live DB data and
    returns raw .xlsx bytes. Sheet order matches the order the data appears
    in the application (Organizations first, Summary last on-screen but
    kept as the first tab for at-a-glance record counts)."""
    org_rows = _collect_organizations()
    employee_rows = _collect_employees()
    telephone_rows = _collect_telephone_directory()
    utility_head_rows = _collect_utility_heads()
    admin_head_rows = _collect_administrative_heads()
    control_room_rows = _collect_directory_numbers("control room")
    switchyard_rows = _collect_directory_numbers("switchyard")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _write_sheet(
        wb, "Organizations",
        ["Sr No.", "Organization Name", "Region", "Category", "Address",
         "Total Employees", "Active Employees", "Inactive Employees"],
        org_rows,
        [8, 45, 30, 26, 50, 14, 14, 14],
    )
    _write_sheet(
        wb, "Employees",
        ["Sr No.", "Name", "Designation", "Office Phone", "Mobile", "Email",
         "Emergency Contact Person", "Relation", "Emergency Contact No."],
        employee_rows,
        [8, 26, 26, 18, 16, 30, 26, 16, 18],
    )
    _write_sheet(
        wb, "Telephone Directory",
        ["Organization", "Address", "Record Type", "Name", "Designation",
         "Office Phone", "Mobile", "Email", "Utility Head"],
        telephone_rows,
        [45, 50, 14, 28, 26, 18, 16, 30, 12],
    )
    _write_sheet(
        wb, "Utility Heads",
        ["#", "Name", "Designation", "Organization", "Address",
         "Office Phone", "Mobile", "Email"],
        utility_head_rows,
        [5, 28, 30, 45, 50, 18, 16, 30],
    )
    _write_sheet(
        wb, "Administrative Heads",
        ["Organization", "Name", "Designation", "Service Type", "Status",
         "Email", "Mobile", "Office Phone"],
        admin_head_rows,
        [45, 28, 32, 18, 14, 32, 16, 16],
    )
    _write_sheet(
        wb, "Control Rooms",
        ["Sr No.", "Name", "Organization", "Phone Number", "Email", "Category"],
        control_room_rows,
        [8, 30, 45, 22, 30, 20],
    )
    _write_sheet(
        wb, "Switchyards",
        ["Sr No.", "Name", "Organization", "Phone Number", "Email", "Category"],
        switchyard_rows,
        [8, 30, 45, 22, 30, 20],
    )

    generated_at = datetime.now()
    counts = [
        ("Organizations", len(org_rows)),
        ("Employees", len(employee_rows)),
        ("Telephone Directory", len(telephone_rows)),
        ("Utility Heads", len(utility_head_rows)),
        ("Administrative Heads", len(admin_head_rows)),
        ("Control Rooms", len(control_room_rows)),
        ("Switchyards", len(switchyard_rows)),
    ]
    # create_sheet(..., 0) inside _write_summary_sheet already inserts the
    # Summary tab at the front, ahead of the 7 data sheets built above.
    _write_summary_sheet(wb, counts, generated_at, generated_by)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
