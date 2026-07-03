"""
verify_data.py
==============
Run: python verify_data.py
Prints a full data quality report for employees, organizations,
and control room / directory numbers.
"""

from app import app
from models import db
from models.employee import Employee
from models.organization import Organization
from models.department import Department
from models.directory_number import DirectoryNumber
from models.emergency_contact import EmergencyContact
from sqlalchemy import text, func

SEP  = "=" * 70
SEP2 = "-" * 70

def banner(title):
    print(f"\n{SEP}")
    print(f"  {title}")
    print(SEP)

def sub(title):
    print(f"\n{SEP2}")
    print(f"  {title}")
    print(SEP2)

def ok(msg):   print(f"  [OK]  {msg}")
def warn(msg): print(f"  [!!]  {msg}")
def info(msg): print(f"  [--]  {msg}")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def run():
    with app.app_context():

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 1 – DATABASE TOTALS")
        # ══════════════════════════════════════════════════════════════════════

        n_orgs  = Organization.query.count()
        n_emps  = Employee.query.count()
        n_depts = Department.query.count()
        n_dirs  = DirectoryNumber.query.count()
        n_ec    = EmergencyContact.query.count()
        n_heads = Employee.query.filter_by(is_utility_head=True).count()

        print(f"\n  {'Organizations':<30} {n_orgs:>6}")
        print(f"  {'Departments':<30} {n_depts:>6}")
        print(f"  {'Employees':<30} {n_emps:>6}")
        print(f"  {'Utility Heads':<30} {n_heads:>6}")
        print(f"  {'Directory Numbers':<30} {n_dirs:>6}")
        print(f"  {'Emergency Contacts':<30} {n_ec:>6}")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 2 – ORGANIZATION CHECKS")
        # ══════════════════════════════════════════════════════════════════════

        # 2a – Missing address
        sub("2a  Organizations with NO address")
        no_addr = Organization.query.filter(
            (Organization.address == None) | (Organization.address == "")
        ).order_by(Organization.organization_name).all()
        if no_addr:
            warn(f"{len(no_addr)} organizations have no address:")
            for o in no_addr:
                print(f"       ID {o.id:>4}  {o.organization_name}")
        else:
            ok("All organizations have an address.")

        # 2b – Missing region
        sub("2b  Organizations with NO region")
        no_region = Organization.query.filter(
            (Organization.region == None) | (Organization.region == "")
        ).order_by(Organization.organization_name).all()
        if no_region:
            warn(f"{len(no_region)} organizations have no region:")
            for o in no_region[:20]:
                print(f"       ID {o.id:>4}  {o.organization_name}")
            if len(no_region) > 20:
                info(f"       ... and {len(no_region)-20} more")
        else:
            ok("All organizations have a region.")

        # 2c – Organizations with zero employees
        sub("2c  Organizations with ZERO employees")
        orgs_with_count = (
            db.session.query(
                Organization.id,
                Organization.organization_name,
                func.count(Employee.id).label("emp_count")
            )
            .outerjoin(Employee, Employee.organization_id == Organization.id)
            .group_by(Organization.id, Organization.organization_name)
            .having(func.count(Employee.id) == 0)
            .order_by(Organization.organization_name)
            .all()
        )
        if orgs_with_count:
            warn(f"{len(orgs_with_count)} organizations have no employees:")
            for r in orgs_with_count[:20]:
                print(f"       ID {r.id:>4}  {r.organization_name}")
            if len(orgs_with_count) > 20:
                info(f"       ... and {len(orgs_with_count)-20} more")
        else:
            ok("Every organization has at least one employee.")

        # 2d – Duplicate org names
        sub("2d  Duplicate organization names")
        dup_orgs = (
            db.session.query(
                Organization.organization_name,
                func.count(Organization.id).label("cnt")
            )
            .group_by(Organization.organization_name)
            .having(func.count(Organization.id) > 1)
            .all()
        )
        if dup_orgs:
            warn(f"{len(dup_orgs)} duplicate organization names found:")
            for d in dup_orgs:
                print(f"       '{d.organization_name}'  (x{d.cnt})")
        else:
            ok("No duplicate organization names.")

        # 2e – Utility head per org check
        sub("2e  Organizations with NO utility head marked")
        orgs_no_head = (
            db.session.query(
                Organization.id,
                Organization.organization_name,
                func.count(Employee.id).label("emp_count")
            )
            .join(Employee, Employee.organization_id == Organization.id)
            .filter(Employee.is_utility_head == False)  # noqa: E712
            .group_by(Organization.id, Organization.organization_name)
            .all()
        )
        heads_set = {
            e.organization_id
            for e in Employee.query.filter_by(is_utility_head=True).all()
        }
        orgs_missing_head = [
            r for r in orgs_no_head if r.id not in heads_set
        ]
        if orgs_missing_head:
            warn(f"{len(orgs_missing_head)} organizations have no utility head marked:")
            for r in orgs_missing_head[:15]:
                print(f"       ID {r.id:>4}  {r.organization_name}  ({r.emp_count} emps)")
            if len(orgs_missing_head) > 15:
                info(f"       ... and {len(orgs_missing_head)-15} more")
        else:
            ok(f"All organizations with employees have a utility head. ({n_heads} total)")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 3 – EMPLOYEE CHECKS")
        # ══════════════════════════════════════════════════════════════════════

        # 3a – Missing name
        sub("3a  Employees with blank / NULL name")
        no_name = Employee.query.filter(
            (Employee.employee_name == None) | (Employee.employee_name == "")
        ).all()
        if no_name:
            warn(f"{len(no_name)} employees have no name:")
            for e in no_name:
                print(f"       ID {e.id:>4}  org_id={e.organization_id}  desig=[{e.designation}]")
        else:
            ok("All employees have a name.")

        # 3b – Missing designation
        sub("3b  Employees with no designation")
        no_desig = Employee.query.filter(
            (Employee.designation == None) | (Employee.designation == "")
        ).count()
        if no_desig:
            warn(f"{no_desig} employees have no designation.")
        else:
            ok("All employees have a designation.")

        # 3c – Missing ALL contact info
        sub("3c  Employees with NO phone AND no email")
        no_contact = Employee.query.filter(
            (Employee.office_phone == None) | (Employee.office_phone == ""),
            (Employee.mobile_phone == None) | (Employee.mobile_phone == ""),
            (Employee.residence_phone == None) | (Employee.residence_phone == ""),
            (Employee.email == None) | (Employee.email == ""),
        ).all()
        if no_contact:
            warn(f"{len(no_contact)} employees have no phone AND no email:")
            for e in no_contact[:20]:
                org = e.organization.organization_name if e.organization else "—"
                print(f"       ID {e.id:>4}  {e.employee_name:<35}  {org}")
            if len(no_contact) > 20:
                info(f"       ... and {len(no_contact)-20} more")
        else:
            ok("Every employee has at least one contact field.")

        # 3d – Missing organization
        sub("3d  Employees with NO organization assigned")
        no_org = Employee.query.filter(Employee.organization_id == None).all()
        if no_org:
            warn(f"{len(no_org)} employees have no organization:")
            for e in no_org[:15]:
                print(f"       ID {e.id:>4}  {e.employee_name}  desig=[{e.designation}]")
        else:
            ok("All employees are assigned to an organization.")

        # 3e – Duplicate employee names in same org
        sub("3e  Duplicate employee names within the same organization")
        dup_emps = (
            db.session.query(
                Employee.employee_name,
                Employee.organization_id,
                func.count(Employee.id).label("cnt")
            )
            .group_by(Employee.employee_name, Employee.organization_id)
            .having(func.count(Employee.id) > 1)
            .all()
        )
        if dup_emps:
            warn(f"{len(dup_emps)} duplicate employee+org combinations:")
            for d in dup_emps[:20]:
                org = db.session.get(Organization, d.organization_id)
                org_name = org.organization_name if org else f"org_id={d.organization_id}"
                print(f"       '{d.employee_name}'  in  '{org_name}'  (x{d.cnt})")
            if len(dup_emps) > 20:
                info(f"       ... and {len(dup_emps)-20} more")
        else:
            ok("No duplicate employee names within any single organization.")

        # 3f – Coverage summary
        sub("3f  Contact field coverage summary")
        total = Employee.query.count()
        has_office  = Employee.query.filter(
            Employee.office_phone != None, Employee.office_phone != ""
        ).count()
        has_mobile  = Employee.query.filter(
            Employee.mobile_phone != None, Employee.mobile_phone != ""
        ).count()
        has_res     = Employee.query.filter(
            Employee.residence_phone != None, Employee.residence_phone != ""
        ).count()
        has_email   = Employee.query.filter(
            Employee.email != None, Employee.email != ""
        ).count()

        def pct(n): return f"{n:>4}  ({100*n//total if total else 0}%)"
        print(f"\n  {'Field':<25} {'Filled':>12}")
        print(f"  {'-'*40}")
        print(f"  {'Office Phone':<25} {pct(has_office)}")
        print(f"  {'Mobile Phone':<25} {pct(has_mobile)}")
        print(f"  {'Residence Phone':<25} {pct(has_res)}")
        print(f"  {'Email':<25} {pct(has_email)}")
        print(f"  {'Total employees':<25} {total:>4}")

        # 3g – Employees per organization (top 10)
        sub("3g  Top 10 organizations by employee count")
        top_orgs = (
            db.session.query(
                Organization.organization_name,
                func.count(Employee.id).label("cnt")
            )
            .join(Employee, Employee.organization_id == Organization.id)
            .group_by(Organization.organization_name)
            .order_by(func.count(Employee.id).desc())
            .limit(10)
            .all()
        )
        for r in top_orgs:
            bar = "#" * min(r.cnt, 40)
            print(f"  {r.organization_name[:45]:<45}  {r.cnt:>4}  {bar}")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 4 – CONTROL ROOM / DIRECTORY NUMBER CHECKS")
        # ══════════════════════════════════════════════════════════════════════

        # 4a – Total by category
        sub("4a  Directory numbers by category")
        cats = (
            db.session.query(
                func.coalesce(DirectoryNumber.category, "Uncategorized").label("cat"),
                func.count(DirectoryNumber.id).label("cnt")
            )
            .group_by(DirectoryNumber.category)
            .order_by(func.count(DirectoryNumber.id).desc())
            .all()
        )
        for c in cats:
            print(f"  {c.cat:<30} {c.cnt:>6}")

        # 4b – Control rooms with no phone
        sub("4b  Control room entries with NO phone number")
        no_phone = DirectoryNumber.query.filter(
            DirectoryNumber.category == "Control Room",
            (DirectoryNumber.phone_number == None) | (DirectoryNumber.phone_number == ""),
        ).all()
        if no_phone:
            warn(f"{len(no_phone)} control rooms have no phone number:")
            for d in no_phone:
                print(f"       ID {d.id:>4}  {d.name:<40}  org=[{d.organization}]")
        else:
            ok("All control room entries have a phone number.")

        # 4c – Duplicate control room entries
        sub("4c  Duplicate directory number entries (same name + organization)")
        dup_dirs = (
            db.session.query(
                DirectoryNumber.name,
                DirectoryNumber.organization,
                func.count(DirectoryNumber.id).label("cnt")
            )
            .group_by(DirectoryNumber.name, DirectoryNumber.organization)
            .having(func.count(DirectoryNumber.id) > 1)
            .all()
        )
        if dup_dirs:
            warn(f"{len(dup_dirs)} duplicate directory entries found:")
            for d in dup_dirs:
                print(f"       '{d.name}'  org=[{d.organization}]  (x{d.cnt})")
        else:
            ok("No duplicate directory number entries.")

        # 4d – Entries with no organization
        sub("4d  Directory numbers with no organization set")
        no_org_dir = DirectoryNumber.query.filter(
            (DirectoryNumber.organization == None) | (DirectoryNumber.organization == ""),
        ).count()
        if no_org_dir:
            warn(f"{no_org_dir} directory entries have no organization set.")
        else:
            ok("All directory entries have an organization set.")

        # 4e – List all control rooms
        sub("4e  All control room numbers (spot-check)")
        control_rooms = (
            DirectoryNumber.query
            .filter(DirectoryNumber.category == "Control Room")
            .order_by(DirectoryNumber.organization, DirectoryNumber.name)
            .all()
        )
        print(f"\n  {'Name':<40} {'Phone':<20} {'Organization'}")
        print(f"  {'-'*90}")
        for cr in control_rooms:
            phone = cr.phone_number or "—"
            org   = (cr.organization or "—")[:35]
            print(f"  {cr.name[:40]:<40} {phone:<20} {org}")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 5 – UTILITY HEAD SPOT-CHECK")
        # ══════════════════════════════════════════════════════════════════════

        sub("5a  All utility heads (name, designation, organization)")
        heads = (
            Employee.query
            .filter_by(is_utility_head=True)
            .join(Organization)
            .order_by(Organization.organization_name)
            .all()
        )
        print(f"\n  {'Name':<30} {'Designation':<35} {'Organization'}")
        print(f"  {'-'*100}")
        for h in heads:
            org = h.organization.organization_name[:35] if h.organization else "—"
            desig = (h.designation or "—")[:35]
            print(f"  {h.employee_name[:30]:<30} {desig:<35} {org}")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 6 – EMERGENCY CONTACTS CHECK")
        # ══════════════════════════════════════════════════════════════════════

        sub("6a  Emergency contacts with no phone")
        no_ec_phone = EmergencyContact.query.filter(
            (EmergencyContact.phone == None) | (EmergencyContact.phone == "")
        ).count()
        if no_ec_phone:
            warn(f"{no_ec_phone} emergency contacts have no phone number.")
        else:
            ok("All emergency contacts have a phone number.")

        sub("6b  Emergency contacts – employee link check")
        orphaned = (
            db.session.execute(
                text("""
                    SELECT ec.id, ec.contact_name
                    FROM emergency_contacts ec
                    LEFT JOIN employees e ON ec.employee_id = e.id
                    WHERE e.id IS NULL
                """)
            ).fetchall()
        )
        if orphaned:
            warn(f"{len(orphaned)} emergency contacts linked to missing employees:")
            for r in orphaned:
                print(f"       EC ID {r[0]}  contact=[{r[1]}]")
        else:
            ok("All emergency contacts have a valid employee link.")

        # ══════════════════════════════════════════════════════════════════════
        banner("SECTION 7 – DATA QUALITY CHECKS")
        # ══════════════════════════════════════════════════════════════════════

        sub("7a  Organization names with suspicious patterns")
        all_orgs = Organization.query.all()
        suspicious = []
        for o in all_orgs:
            n = o.organization_name or ""
            flags = []
            if "  " in n:
                flags.append("double-space")
            if n != n.strip():
                flags.append("leading/trailing space")
            if any(c in n for c in ["–", "—", "​", "�"]):
                flags.append("unicode dash/garbage")
            # If name looks like an address (starts with number or has "Road", "Plot", etc.)
            low = n.lower()
            if any(kw in low for kw in ["road no", "plot no", "sector ", "flat ", "survey"]):
                flags.append("looks like an address")
            if flags:
                suspicious.append((o.id, n, ", ".join(flags)))
        if suspicious:
            warn(f"{len(suspicious)} organization names with quality issues:")
            for org_id, name, flags in suspicious:
                print(f"       ID {org_id:>4}  [{name[:60]}]  -> {flags}")
        else:
            ok("No suspicious organization names found.")

        sub("7b  Employee names with suspicious patterns")
        all_emps = Employee.query.all()
        emp_issues = []
        for e in all_emps:
            n = e.employee_name or ""
            flags = []
            if "  " in n:
                flags.append("double-space")
            if n != n.strip():
                flags.append("leading/trailing space")
            if any(c in n for c in ["�", "​"]):
                flags.append("garbage char")
            if n.lower() in {"n/a", "nil", "none", "-", "–", "na"}:
                flags.append("placeholder name")
            if flags:
                emp_issues.append((e.id, n, ", ".join(flags)))
        if emp_issues:
            warn(f"{len(emp_issues)} employee names with quality issues:")
            for eid, name, flags in emp_issues[:20]:
                print(f"       ID {eid:>4}  [{name}]  -> {flags}")
            if len(emp_issues) > 20:
                info(f"       ... and {len(emp_issues)-20} more")
        else:
            ok("No suspicious employee names found.")

        # ══════════════════════════════════════════════════════════════════════
        banner("SUMMARY")
        # ══════════════════════════════════════════════════════════════════════

        print(f"""
  Total Organizations     : {n_orgs}
  Total Employees         : {n_emps}
  Total Utility Heads     : {n_heads}
  Total Directory Numbers : {n_dirs}
    └─ Control Rooms      : {control_rooms.__len__()}
  Total Emergency Contacts: {n_ec}

  Orgs missing address    : {len(no_addr)}
  Emps missing all contact: {len(no_contact)}
  Emps without org        : {len(no_org)}
  Duplicate emp+org       : {len(dup_emps)}
  Duplicate dir entries   : {len(dup_dirs)}
  Org name issues         : {len(suspicious)}
  Employee name issues    : {len(emp_issues)}
""")


if __name__ == "__main__":
    run()
