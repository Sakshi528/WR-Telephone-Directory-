"""Data-cleanliness audit for Email Lists (EmailGroup/GroupMember/
EmailGroupFilter) and Administrative Heads (AdministrativeHead/
AdministrativeHeadAssistant/AdministrativeHeadHistory) -- separate data
models from directory_numbers, with a different actual risk shape (see
each section below). Read-only -- writes reports/email_lists_audit.csv and
reports/administrative_heads_audit.csv only, no DB writes.

Schema recap (see models/email_group.py, models/email_group_filter.py,
models/administrative_head.py, models/administrative_head_assistant.py,
models/administrative_head_history.py for the full picture):

  EmailGroup: name (unique), list_type (STATIC/DYNAMIC) -- no email field of
  its own. STATIC membership lives in GroupMember (group_id + employee_id,
  unique together); DYNAMIC groups resolve live from EmailGroupFilter rows
  instead. Neither table duplicates an employee's contact info -- they only
  reference Employee, so "clean" for these two tables is about the
  reference/filter structure, not text content.

  AdministrativeHead: has its own standalone name/office_phone/mobile_phone/
  email/office_address columns, but ONLY used when employee_id IS NULL (see
  resolved_* properties) -- when employee_id is set, contact info is read
  live from Employee instead. AdministrativeHeadAssistant follows the same
  pattern. AdministrativeHeadHistory is an archive with the same shape.

The "Select All + Copy Selected Emails" feature (static/js/email_distribution.js
dedupeSortedEmails, mirrored server-side by
services/email_distribution_service.dedupe_and_sort) already blank-filters
and case-insensitively dedupes before copying -- so a duplicate person or a
blank email can't break it, they just silently collapse/drop. What it does
NOT do is validate email syntax, so a non-blank but garbled value (typo,
placeholder text) would pass through untouched and land in the copied
output -- that's the actual residual risk this audit checks for.

Usage:
  python scripts/audit_email_lists_and_admin_heads.py
"""

import csv
import os
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.email_group import EmailGroup, GroupMember
from models.email_group_filter import EmailGroupFilter
from models.administrative_head import AdministrativeHead
from models.administrative_head_assistant import AdministrativeHeadAssistant
from models.administrative_head_history import AdministrativeHeadHistory
from models.employee import Employee

EMAIL_LISTS_CSV = "reports/email_lists_audit.csv"
ADMIN_HEADS_CSV = "reports/administrative_heads_audit.csv"

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_garbled_email(value):
    """Non-blank but not even loosely email-shaped -- exactly what
    dedupe_and_sort's blank-filter does NOT catch."""
    v = (value or "").strip()
    return bool(v) and not EMAIL_RE.match(v)


# ─── Email Lists (EmailGroup / GroupMember / EmailGroupFilter) ────────────

def audit_email_lists():
    rows = []
    with app.app_context():
        groups = EmailGroup.query.all()
        members = GroupMember.query.all()
        filters = EmailGroupFilter.query.all()

        if not groups:
            # Nothing saved yet -- report the empty state explicitly rather
            # than silently writing a 0-row file with no explanation.
            rows.append({
                "id": "", "field": "(table)", "issue_type": "no_data",
                "raw_value": "",
                "notes": "email_groups has 0 rows (and group_members / email_group_filters "
                         "have 0 rows too) -- the Email Lists feature has no saved static or "
                         "dynamic groups in this database yet, so there is nothing to audit "
                         "for duplicate/garbled/malformed content.",
            })
            return rows, {"groups": 0, "members": 0, "filters": 0}

        # (kept for when groups exist in a future run)
        seen_group_member = Counter()
        for m in members:
            seen_group_member[(m.group_id, m.employee_id)] += 1
        for (gid, eid), count in seen_group_member.items():
            if count > 1:
                rows.append({
                    "id": f"group{gid}/emp{eid}", "field": "(group_members)",
                    "issue_type": "duplicate_member", "raw_value": "",
                    "notes": f"employee {eid} appears {count}x in group {gid} -- should be "
                             "prevented by the uq_group_member unique constraint; if this shows "
                             "up, the constraint isn't doing its job.",
                })

        emp_by_id = {e.id: e for e in Employee.query.all()}
        for m in members:
            emp = emp_by_id.get(m.employee_id)
            email = emp.email if emp else None
            if is_garbled_email(email):
                rows.append({
                    "id": m.id, "field": "employee.email (via group_members)",
                    "issue_type": "garbled_email", "raw_value": email,
                    "notes": f"group {m.group_id}, employee {m.employee_id}: non-blank email "
                             "doesn't look like a valid address -- would pass through "
                             "dedupe_and_sort's blank-filter untouched and land in copied output.",
                })

        for f in filters:
            populated = [f.organization_id, f.category_id, f.role_value, f.status_value]
            non_null = [p for p in populated if p is not None]
            expected = {
                "ORGANIZATION": f.organization_id, "ORGANIZATION_CATEGORY": f.category_id,
                "ROLE": f.role_value, "STATUS": f.status_value,
            }.get(f.filter_type)
            if len(non_null) != 1 or expected is None:
                rows.append({
                    "id": f.id, "field": "filter_type/value", "issue_type": "malformed_filter",
                    "raw_value": f.filter_type,
                    "notes": f"expected exactly one of organization_id/category_id/role_value/"
                             f"status_value populated matching filter_type={f.filter_type!r}; "
                             f"found {len(non_null)} populated.",
                })

    return rows, {"groups": len(groups), "members": len(members), "filters": len(filters)}


# ─── Administrative Heads ──────────────────────────────────────────────────

def audit_administrative_heads():
    rows = []
    with app.app_context():
        heads = AdministrativeHead.query.all()
        assistants = AdministrativeHeadAssistant.query.all()
        history = AdministrativeHeadHistory.query.all()

        emp_by_id = {e.id: e for e in Employee.query.all()}

        # duplicate (organization_id, role_title) -- should be impossible for
        # ACTIVE ADMINISTRATIVE_HEAD rows (partial unique index), but KMP
        # allows multiples, and this check isn't scoped to status/category so
        # it also catches ON_LEAVE/VACANT/INACTIVE collisions the index wouldn't.
        org_role = Counter((h.organization_id, h.role_title) for h in heads)
        for (org_id, role_title), count in org_role.items():
            if count > 1:
                matching_ids = [h.id for h in heads if h.organization_id == org_id and h.role_title == role_title]
                rows.append({
                    "id": ", ".join(str(i) for i in matching_ids), "field": "(row)",
                    "issue_type": "possible_duplicate_row", "raw_value": role_title,
                    "notes": f"{count} rows share (organization_id={org_id}, role_title={role_title!r})",
                })

        for h in heads:
            if h.employee_id:
                emp = emp_by_id.get(h.employee_id)
                if emp is None:
                    rows.append({
                        "id": h.id, "field": "employee_id", "issue_type": "broken_reference",
                        "raw_value": h.employee_id,
                        "notes": "employee_id doesn't resolve to an existing Employee row -- "
                                 "resolved_name/email/etc. would all silently return blank.",
                    })
                    continue
                if is_garbled_email(emp.email):
                    rows.append({
                        "id": h.id, "field": "resolved_email (via Employee)",
                        "issue_type": "garbled_email", "raw_value": emp.email,
                        "notes": f"linked employee {h.employee_id} ({emp.employee_name}): "
                                 "non-blank email doesn't look like a valid address.",
                    })
                if not emp.email:
                    rows.append({
                        "id": h.id, "field": "resolved_email (via Employee)",
                        "issue_type": "missing_email", "raw_value": "",
                        "notes": f"linked employee {h.employee_id} ({emp.employee_name}) has no "
                                 "email at all -- silently excluded from Copy Selected Emails "
                                 "(dedupe_and_sort blank-filters it), not a crash, but a real gap.",
                    })
            else:
                # standalone-identity path -- only relevant when employee_id IS NULL
                if is_garbled_email(h.email):
                    rows.append({
                        "id": h.id, "field": "email", "issue_type": "garbled_email",
                        "raw_value": h.email, "notes": "standalone email field doesn't look valid.",
                    })
                if not h.name and not h.email:
                    rows.append({
                        "id": h.id, "field": "name/email", "issue_type": "missing_identity",
                        "raw_value": "", "notes": "employee_id is NULL and neither standalone "
                                 "name nor email is set.",
                    })

        for a in assistants:
            emp = emp_by_id.get(a.employee_id) if a.employee_id else None
            email = emp.email if emp else a.email
            if is_garbled_email(email):
                rows.append({
                    "id": a.id, "field": "resolved_email", "issue_type": "garbled_email",
                    "raw_value": email, "notes": "assistant's resolved email doesn't look valid.",
                })

        if not assistants:
            rows.append({
                "id": "", "field": "(table)", "issue_type": "no_data", "raw_value": "",
                "notes": "administrative_head_assistants has 0 rows.",
            })
        if not history:
            rows.append({
                "id": "", "field": "(table)", "issue_type": "no_data", "raw_value": "",
                "notes": "administrative_head_history has 0 rows.",
            })

    return rows, {"heads": len(heads), "assistants": len(assistants), "history": len(history)}


def write_csv(path, rows):
    fieldnames = ["id", "field", "issue_type", "raw_value", "notes"]
    os.makedirs("reports", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    email_rows, email_counts = audit_email_lists()
    admin_rows, admin_counts = audit_administrative_heads()

    write_csv(EMAIL_LISTS_CSV, email_rows)
    write_csv(ADMIN_HEADS_CSV, admin_rows)

    print("=" * 70)
    print("  Email Lists audit -- table counts:", email_counts)
    print(f"  Rows written: {len(email_rows)} -> {EMAIL_LISTS_CSV}")
    for issue, count in Counter(r["issue_type"] for r in email_rows).items():
        print(f"    - {issue}: {count}")
    print("-" * 70)
    print("  Administrative Heads audit -- table counts:", admin_counts)
    print(f"  Rows written: {len(admin_rows)} -> {ADMIN_HEADS_CSV}")
    for issue, count in Counter(r["issue_type"] for r in admin_rows).items():
        print(f"    - {issue}: {count}")
    print("=" * 70)


if __name__ == "__main__":
    main()
