"""Live resolution of Email Distribution List membership. Nothing here is
ever stored -- every call queries current Employee/Organization/
DirectoryNumber rows and returns a fresh result, mirroring org_stats.py's
"compute live, don't materialize" philosophy. No email address is ever
written to any table; STANDARD_GROUPS and the resolvers below are the single
source of truth for both the home page counts and each group's detail view.

Data-availability note: the Emergency section originally scoped for this
module ("Hospitals", "Emergency Services") has no email-bearing source table
today -- emergency_contacts holds personal next-of-kin phone numbers per
employee, not an institutional hospital/emergency-services directory, and no
DirectoryNumber.category values for those exist yet (only "Control Room" and
"Switchyard" do). Those two are exposed instead; Hospitals/Emergency Services
can be added as one more STANDARD_GROUPS entry once such data is imported.
"""

from types import SimpleNamespace
from urllib.parse import quote

from models.employee import Employee
from models.organization import Organization
from models.organization_category import OrganizationCategory
from models.directory_number import DirectoryNumber
from models.administrative_head import AdministrativeHead
from models.administrative_head_assistant import AdministrativeHeadAssistant
from utils.designation_rank import resolve_utility_head, compute_utility_head_ids

MAILTO_SAFE_LIMIT = 1800


# ─── STANDARD GROUP RESOLVERS ────────────────────────────────────────────────

# contact_type slug -> DirectoryNumber.category keyword(s) it matches
CONTACT_TYPE_DIRECTORY_KEYWORDS = {
    "control_room": ["control room"],
    "switchyard": ["switchyard"],
}

CONTACT_TYPE_CHOICES = [
    ("all", "All Contacts"),
    ("utility_head", "Utility Heads"),
    ("administrative_head", "Administrative Heads"),
    ("kmp", "KMP"),
    ("control_room", "Control Room Only"),
    ("switchyard", "Switchyard Only"),
]

_EMPLOYEE_CONTACT_TYPES = ("all", "employees", "utility_head", "administrative_head", "kmp")


def resolve_category_contacts(category_name, contact_type="all", active_only=True, state=None):
    """Organization-based group, with an optional sub-filter narrowing it to
    just one contact type -- e.g. "only Control Room emails in Transmission
    Utilities", or "only Utility Heads in RE Generators". "all" (the default,
    used for home-page counts) combines every employee with any Control
    Room/Switchyard directory numbers belonging to organizations in this
    category. Utility Head/Administrative Head/KMP scoping reuses the exact
    same role definitions as resolve_dynamic_group's ROLE handling -- just
    intersected with this category's organizations instead of every
    organization. `state` further narrows to one Organization.state value,
    for state-based categories (State SLDC/STU/DISCOM)."""
    if contact_type not in dict(CONTACT_TYPE_CHOICES):
        contact_type = "all"

    org_query = (
        Organization.query
        .join(OrganizationCategory, Organization.category_id == OrganizationCategory.id)
        .filter(OrganizationCategory.category_name == category_name)
    )
    if state:
        org_query = org_query.filter(Organization.state == state)
    org_ids = [org.id for org in org_query.all()]

    rows = []
    if contact_type in _EMPLOYEE_CONTACT_TYPES:
        emp_query = Employee.query.filter(Employee.organization_id.in_(org_ids))
        if active_only:
            emp_query = emp_query.filter(Employee.status == "ACTIVE")
        employees = emp_query.order_by(Employee.employee_name).all()

        if contact_type == "utility_head":
            head_ids = compute_utility_head_ids()
            employees = [e for e in employees if e.id in head_ids]
        elif contact_type == "administrative_head":
            admin_head_emp_ids = {
                ah.employee_id for ah in
                AdministrativeHead.query.filter_by(role_category="ADMINISTRATIVE_HEAD", status="ACTIVE").all()
                if ah.employee_id
            }
            employees = [e for e in employees if e.id in admin_head_emp_ids]
        elif contact_type == "kmp":
            employees = [e for e in employees if e.is_kmp]

        rows.extend(employees)

    if contact_type in ("all", "administrative_head"):
        # Standalone (non-Employee) Administrative Heads have no Employee row
        # to be found through the block above -- included here so they aren't
        # silently invisible to Email Distribution.
        rows.extend(AdministrativeHead.query.filter(
            AdministrativeHead.role_category == "ADMINISTRATIVE_HEAD",
            AdministrativeHead.employee_id.is_(None),
            AdministrativeHead.organization_id.in_(org_ids),
            AdministrativeHead.status == "ACTIVE",
        ).all())

    if contact_type in ("all", "control_room", "switchyard"):
        keywords = (
            CONTACT_TYPE_DIRECTORY_KEYWORDS[contact_type] if contact_type != "all"
            else [kw for kws in CONTACT_TYPE_DIRECTORY_KEYWORDS.values() for kw in kws]
        )
        dn_rows = DirectoryNumber.query.filter(
            DirectoryNumber.organization_id.in_(org_ids),
            DirectoryNumber.category.isnot(None),
        ).all()
        rows.extend([r for r in dn_rows if any(k in r.category.lower() for k in keywords)])

    return rows


def resolve_utility_heads():
    """Same definition as routes/user_routes.py:_compute_utility_heads --
    one auto-resolved head per organization, never the manually-set
    Employee.is_utility_head flag. Archived (non-ACTIVE) employees are
    never eligible, same as the live-page computation."""
    heads = []
    for org in Organization.query.order_by(Organization.organization_name).all():
        employees = (
            Employee.query.filter_by(organization_id=org.id)
            .filter(Employee.status == "ACTIVE")
            .order_by(Employee.id).all()
        )
        head = resolve_utility_head(employees)
        if head:
            heads.append(head)
    return heads


def resolve_administrative_heads():
    """Returns AdministrativeHead rows directly (not Employee rows) so a
    standalone head (employee_id IS NULL) is included -- an Employee-table
    join would silently drop them. Callers read identity/contact via the
    resolved_* properties, which to_contact_row/dedupe_and_sort already do.

    Excludes a head linked to an employee who is no longer ACTIVE (the
    change_employee_status auto-archive hook should already have moved
    them to administrative_head_history -- this is a defensive check, not
    the primary mechanism)."""
    heads = AdministrativeHead.query.filter_by(
        role_category="ADMINISTRATIVE_HEAD", status="ACTIVE"
    ).all()
    heads = [h for h in heads if h.employee_id is None or (h.employee and h.employee.status == "ACTIVE")]
    heads.sort(key=lambda h: (h.resolved_name or "").lower())
    return heads


def resolve_administrative_head_contacts(mode="official"):
    """Administrative Heads split by Official / PA-PS / both -- PA/PS rows
    are wrapped in a bare SimpleNamespace(email=...) so the existing
    dedupe_and_sort contract (getattr(row, "email", ...)) needs no change.
    AdministrativeHeadAssistant.resolved_email already prefers the linked
    Employee's email when the assistant is one, so nothing is duplicated."""
    rows = []
    if mode in ("official", "both"):
        rows.extend(resolve_administrative_heads())
    if mode in ("pa", "both"):
        assistants = (
            AdministrativeHeadAssistant.query
            .join(AdministrativeHead, AdministrativeHeadAssistant.administrative_head_id == AdministrativeHead.id)
            .filter(AdministrativeHead.role_category == "ADMINISTRATIVE_HEAD", AdministrativeHead.status == "ACTIVE")
            .all()
        )
        rows.extend(SimpleNamespace(email=a.resolved_email) for a in assistants)
    return rows


def resolve_kmp():
    return (
        Employee.query
        .filter(Employee.is_kmp == True)  # noqa: E712
        .filter(Employee.status == "ACTIVE")
        .order_by(Employee.employee_name)
        .all()
    )


def resolve_directory_category_group(keywords):
    """DirectoryNumber rows whose free-text category contains any of the
    given keywords -- same case-insensitive substring match already used by
    routes/user_routes.py:telephone_directory for Control Room/Switchyard
    grouping, reused here rather than re-deciding the match rule."""
    rows = DirectoryNumber.query.filter(DirectoryNumber.category.isnot(None)).all()
    keywords = [k.lower() for k in keywords]
    return [r for r in rows if any(k in r.category.lower() for k in keywords)]


CATEGORY_SLUGS = [
    ("transmission-utility", "Transmission Utility", "Transmission Utilities"),
    ("rldc", "RLDC", "RLDCs"),
    ("state-sldc", "State SLDC", "State SLDCs"),
    ("cpsu", "CPSU", "CPSUs"),
    ("re-generators", "RE Generators", "RE Generators"),
    ("discom", "DISCOM", "DISCOMs"),
    ("generation-company", "Generation Company", "Generation Companies"),
    ("thermal", "Thermal", "Thermal"),
    ("hydel", "Hydel", "Hydel"),
    ("nuclear", "Nuclear", "Nuclear"),
    ("qca", "QCA", "QCAs"),
    ("others", "Others", "Others"),
]

STANDARD_GROUPS = [
    {
        "slug": slug, "name": plural, "section": "Organization Based",
        "contact_type_filter": True,
        "resolver": (lambda contact_type="all", cat=category_name: resolve_category_contacts(cat, contact_type)),
    }
    for slug, category_name, plural in CATEGORY_SLUGS
] + [
    {"slug": "utility-heads", "name": "Utility Heads", "section": "Role Based",
     "contact_type_filter": False,
     "resolver": (lambda contact_type="all": resolve_utility_heads())},
    {"slug": "kmp", "name": "KMP", "section": "Role Based",
     "contact_type_filter": False,
     "resolver": (lambda contact_type="all": resolve_kmp())},
]

STANDARD_GROUPS_BY_SLUG = {g["slug"]: g for g in STANDARD_GROUPS}


# ─── CUSTOM (SAVED DYNAMIC) GROUP RESOLUTION ─────────────────────────────────

def resolve_dynamic_group(email_group):
    """Resolves a DYNAMIC EmailGroup's current membership from its
    EmailGroupFilter rows. Rows sharing a filter_type are ORed; different
    filter_types are ANDed (see models/email_group_filter.py)."""
    filters_by_type = {}
    for f in email_group.filters:
        filters_by_type.setdefault(f.filter_type, []).append(f)

    query = Employee.query.outerjoin(Organization, Employee.organization_id == Organization.id)

    org_ids = [f.organization_id for f in filters_by_type.get("ORGANIZATION", [])] or None
    category_ids = [f.category_id for f in filters_by_type.get("ORGANIZATION_CATEGORY", [])] or None

    if org_ids:
        query = query.filter(Employee.organization_id.in_(org_ids))

    if category_ids:
        query = query.filter(Organization.category_id.in_(category_ids))

    if "STATUS" in filters_by_type:
        # An admin who explicitly configured a STATUS filter (even a
        # non-active one, e.g. to build a "recently transferred" list) has
        # made an intentional choice -- respected as-is.
        status_values = [f.status_value for f in filters_by_type["STATUS"]]
        query = query.filter(Employee.status.in_(status_values))
    else:
        # No explicit STATUS filter configured -- default to ACTIVE only,
        # so a group never silently starts including archived employees
        # just because nobody thought to add a STATUS filter.
        query = query.filter(Employee.status == "ACTIVE")

    employees = query.order_by(Employee.employee_name).all()
    rows = employees

    if "ROLE" in filters_by_type:
        role_values = {f.role_value for f in filters_by_type["ROLE"]}
        allowed_ids = set()
        if "UTILITY_HEAD" in role_values:
            allowed_ids |= compute_utility_head_ids()
        if "KMP" in role_values:
            allowed_ids |= {e.id for e in employees if e.is_kmp}
        if "ADMINISTRATIVE_HEAD" in role_values:
            allowed_ids |= {
                ah.employee_id for ah in
                AdministrativeHead.query.filter_by(role_category="ADMINISTRATIVE_HEAD", status="ACTIVE").all()
                if ah.employee_id
            }
        # employees was already filtered to ACTIVE-only above (default, or
        # an explicit STATUS filter) -- narrowing against it here is enough,
        # no separate linked-employee-status check needed.
        rows = [e for e in employees if e.id in allowed_ids]

        if "ADMINISTRATIVE_HEAD" in role_values:
            # Standalone (non-Employee) heads can't appear in an Employee
            # queryset at all -- unioned in separately, honoring the
            # ORGANIZATION/ORGANIZATION_CATEGORY scoping. STATUS here is an
            # Employee-status filter (ACTIVE/TRANSFERRED/...) and doesn't
            # apply to a standalone head, so it's not applied to this half.
            standalone_query = AdministrativeHead.query.filter(
                AdministrativeHead.role_category == "ADMINISTRATIVE_HEAD",
                AdministrativeHead.employee_id.is_(None),
                AdministrativeHead.status == "ACTIVE",
            )
            if org_ids:
                standalone_query = standalone_query.filter(AdministrativeHead.organization_id.in_(org_ids))
            if category_ids:
                standalone_query = standalone_query.join(
                    Organization, AdministrativeHead.organization_id == Organization.id
                ).filter(Organization.category_id.in_(category_ids))
            rows = rows + standalone_query.all()

    # Manual overrides -- an admin's explicit include/exclude of a specific
    # employee always wins over the filter-computed result, applied last so
    # neither ordering nor filter combination can undo it.
    exclude_ids = {f.employee_id for f in filters_by_type.get("EXCLUDE_EMPLOYEE", [])}
    if exclude_ids:
        rows = [r for r in rows if getattr(r, "id", None) not in exclude_ids or not isinstance(r, Employee)]

    include_ids = {f.employee_id for f in filters_by_type.get("INCLUDE_EMPLOYEE", [])} - exclude_ids
    include_ids -= {r.id for r in rows if isinstance(r, Employee)}
    if include_ids:
        rows = rows + Employee.query.filter(Employee.id.in_(include_ids)).all()

    return rows


# ─── SHARED HELPERS ───────────────────────────────────────────────────────────

def dedupe_and_sort(rows):
    """Blank-filters, case-insensitively dedupes (keeping the first-seen
    casing), and alphabetically sorts email addresses. The one place this
    happens -- every resolver above funnels through here so no group can
    ever return a duplicate or blank address.

    Prefers resolved_email when present (AdministrativeHead/
    AdministrativeHeadAssistant: reads the linked Employee's email when
    there is one, else the row's own free-text email) -- checked before the
    plain .email attribute, since AdministrativeHead.email is itself only
    the standalone-identity field, not the resolved one."""
    seen = {}
    for row in rows:
        email = (getattr(row, "resolved_email", None) or getattr(row, "email", None) or "").strip()
        if not email:
            continue
        key = email.lower()
        if key not in seen:
            seen[key] = email
    return sorted(seen.values(), key=str.lower)


def build_mailto(emails, subject=None):
    url = "mailto:" + quote(",".join(emails))
    if subject:
        url += "?subject=" + quote(subject)
    return url


def to_contact_row(row):
    """Normalizes an Employee, DirectoryNumber, or AdministrativeHead into
    the one contact shape the browse grid, Excel export, and CSV export all
    share -- the single seam between "how we resolve contacts" and "how we
    display them" so the three can never drift apart. Control Room/
    Switchyard rows repeat their directory label as both name and
    designation (they're a place, not a person)."""
    if isinstance(row, DirectoryNumber):
        org_name = row.organization_obj.organization_name if row.organization_obj else (row.organization or "")
        return {
            "organization": org_name,
            "name": row.name or "",
            "designation": row.category or row.name or "",
            "email": row.email or "",
            "mobile": row.phone_number or "",
            "record_id": row.id, "record_type": "directory_number",
        }
    if isinstance(row, AdministrativeHead):
        return {
            "organization": row.organization.organization_name if row.organization else "",
            "name": row.resolved_name or "",
            "designation": row.resolved_designation or "",
            "email": row.resolved_email or "",
            "mobile": row.resolved_mobile_phone or "",
            "record_id": row.id, "record_type": "administrative_head",
        }
    return {
        "organization": row.organization.organization_name if row.organization else "",
        "name": row.employee_name or "",
        "designation": row.designation or "",
        "email": row.email or "",
        "mobile": row.mobile_phone or "",
        "record_id": row.id, "record_type": "employee",
    }


def summarize_contacts(rows):
    """Organizations / Employees / Valid Emails / Missing Emails /
    Duplicates Removed for the browse view's summary panel -- derived from
    the same rows and the same dedupe_and_sort used everywhere else, so the
    numbers can never disagree with what's on screen."""
    contact_rows = [to_contact_row(r) for r in rows]
    valid = sum(1 for r in contact_rows if r["email"])
    unique = len(dedupe_and_sort(rows))
    return {
        "organizations": len({r["organization"] for r in contact_rows if r["organization"]}),
        "employees": len(contact_rows),
        "valid_emails": valid,
        "missing_emails": len(contact_rows) - valid,
        "duplicates_removed": valid - unique,
    }


def filter_contact_rows(contact_rows, q):
    """Server-side mirror of the browse view's client-side search, so a
    direct API call to the data.json endpoint with ?q= is still meaningful
    on its own."""
    if not q:
        return contact_rows
    q = q.strip().lower()
    return [
        r for r in contact_rows
        if q in r["organization"].lower() or q in r["name"].lower()
        or q in r["designation"].lower() or q in r["email"].lower()
    ]
