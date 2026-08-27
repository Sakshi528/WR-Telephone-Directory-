"""Applies the fixes approved on 2026-08-27 out of the full-data-quality-
audit review (full_data_quality_audit.py / build_fix_proposals.py /
build_duplicate_row_merge_plan.py), and reconstructs
reports/session_changes_audit_20260827.csv for them -- same pattern as
build_session_audit.py used for the 2026-08-26 session, since
DirectoryNumber edits still don't go through services/audit_service.py (see
routes/admin_routes.py's edit_directory_number, which commits with no
log_audit_event call).

Each apply_* function does its own read-before-write (SELECT the current
value immediately before updating it) so the audit row's old_value is the
true value at write time, not a value assumed from an earlier report.

Usage:
  python scripts/apply_session_fixes_20260827.py            # dry run, no writes
  python scripts/apply_session_fixes_20260827.py --apply    # commit + write audit CSV
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models import db
from models.directory_number import DirectoryNumber
from models.organization import Organization
from models.organization_category import OrganizationCategory
from models.employee import Employee
from services.audit_service import log_audit_event
from scripts.compare_word_excel import LEADING_NUMBER_RE

OUT_CSV = "reports/session_changes_audit_20260827.csv"
SESSION_DATE = "2026-08-27"

audit_rows = []


def record(table, row_id, field, old_value, new_value, reason):
    audit_rows.append({
        "table": table, "row_id": row_id, "field": field,
        "old_value": old_value, "new_value": new_value,
        "reason": reason, "timestamp": SESSION_DATE,
    })


# ─── Item 1: duplicated_phone import-bug fixes (2 rows) ───────────────────

PHONE_FIXES = {
    7675: "Control Room / 8408880569 / 02355235005 / 02355235311",
    7846: "02525-265510 / 265404 / NA",
}


def apply_phone_fixes(apply):
    for row_id, new_value in PHONE_FIXES.items():
        dn = db.session.get(DirectoryNumber, row_id)
        old_value = dn.phone_number
        if apply:
            dn.phone_number = new_value
        record(
            "directory_numbers", row_id, "phone_number", old_value, new_value,
            "duplicated_phone fix, origin=introduced_during_import (clean in Word, dirty only in "
            "the app -- unambiguous import bug per reports/word_vs_app_cleanliness.csv): exact-"
            "repeat segment(s) removed, order preserved, no other segment touched.",
        )


# ─── CPSU population: Power Grid + its 2 regional offices, NTPC Vidyut
# Vyapar Nigam Limited (4 orgs). Uses the same real audit_log table the
# live Manage Organizations edit route uses (routes/admin_routes.py's
# CATEGORY_CHANGE log_audit_event call), not just the session CSV, since
# organizations.category_id changes DO have proper audit infrastructure
# (unlike directory_numbers) -- this keeps script-driven changes
# indistinguishable from UI-driven ones in the real audit trail. ─────────

CPSU_MOVE_ORG_IDS = [6309, 6312, 6313, 6499]


def apply_org_category_move(apply, org_id, new_category_name, reason):
    """Shared by every single-org category reassignment this session --
    logs to both the real audit_log table (same as a Manage Organizations
    UI edit) and the session CSV."""
    new_category = OrganizationCategory.query.filter_by(category_name=new_category_name).first()
    org = db.session.get(Organization, org_id)
    old_category = db.session.get(OrganizationCategory, org.category_id) if org.category_id else None
    old_category_name = old_category.category_name if old_category else None
    if old_category_name == new_category_name:
        return  # already there (e.g. a re-run after a prior successful apply) -- not a real change, don't log
    if apply:
        log_audit_event(
            module="organizations", record_type="Organization", record_id=org.id,
            action="CATEGORY_CHANGE", field_name="category_id",
            old_value=str(org.category_id) if org.category_id else None,
            new_value=str(new_category.id),
            reason=reason,
        )
        org.category_id = new_category.id
    record(
        "organizations", org_id, "category_id (via category_name)",
        old_category_name, new_category_name,
        f"{reason} Also logged to the real audit_log table via "
        f"services.audit_service.log_audit_event, same as a Manage Organizations UI edit would.",
    )


def apply_cpsu_moves(apply):
    for org_id in CPSU_MOVE_ORG_IDS:
        apply_org_category_move(
            apply, org_id, "CPSU",
            "CPSU population per reports/cpsu_population_proposal.csv -- approved move "
            "(Power Grid + 2 regional offices + NTPC Vidyut Vyapar Nigam Limited).",
        )


# ─── Org-category reclassification proposal (reports/reclassification_proposal.csv),
# approved as-is: State SLDC -> SLDC rename (category row itself, not an
# org-level change -- no live-app route renames a category today, so
# there's no existing audit_log action precedent; using RENAME as a new
# but analogous action label to CATEGORY_CHANGE), 74x RE Generators ->
# Generation Company merge (org-level, reuses apply_org_category_move),
# and the 2 name-based DISCOM/Transmission moves (same). Read directly
# from the proposal CSV rather than hardcoding ids, so this can't drift
# from what was actually reviewed and approved. ──────────────────────────

RECLASSIFICATION_CSV = "reports/reclassification_proposal.csv"


def apply_state_sldc_rename(apply):
    cat = OrganizationCategory.query.filter_by(category_name="State SLDC").first()
    if cat is None:
        record(
            "organization_categories", "(none)", "category_name", "State SLDC", "(already renamed)",
            "State SLDC category already didn't exist at write time -- nothing to rename.",
        )
        return
    affected = Organization.query.filter_by(category_id=cat.id).count()
    old_name = cat.category_name
    if apply:
        log_audit_event(
            module="organization_categories", record_type="OrganizationCategory", record_id=cat.id,
            action="RENAME", field_name="category_name", old_value=old_name, new_value="SLDC",
            reason="Approved per reports/reclassification_proposal.csv -- category label rename, "
                   f"no org.category_id changes ({affected} orgs stay on the same category id).",
        )
        cat.category_name = "SLDC"
    record(
        "organization_categories", cat.id, "category_name", old_name, "SLDC",
        f"Approved per reports/reclassification_proposal.csv -- label rename affecting {affected} "
        "orgs' displayed category, none of which change category_id.",
    )


def apply_reclassification_org_moves(apply):
    with open(RECLASSIFICATION_CSV, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    for r in rows:
        rule = r["matched_rule"]
        if rule.endswith("(already here)"):
            continue  # no-op row, already in the proposed category
        if rule.startswith("category = State SLDC"):
            continue  # handled by apply_state_sldc_rename, not a per-org move
        # remaining rules are org-level category_id moves: the 74 RE
        # Generators -> Generation Company merge, and the 2 name-based
        # DISCOM/Transmission moves
        apply_org_category_move(
            apply, int(r["org_id"]), r["proposed_category"],
            f"Approved per reports/reclassification_proposal.csv, rule: {rule!r}.",
        )


NSPCL_ORG_ID = 6458


def apply_nspcl_move(apply):
    apply_org_category_move(
        apply, NSPCL_ORG_ID, "Generation Company",
        "NSPCL correction: initially flagged AMBIGUOUS in reports/cpsu_population_proposal.csv "
        "(generation JV, only partially NTPC-owned, didn't cleanly fit CPSU's definition) -- "
        "you clarified it should move to Generation Company instead, reflecting its actual "
        "operational function (a 'Power Company' by name), not left in Others.",
    )


# ─── IPP removal: category row is unused (0 orgs, 0 email_group_filters
# reference it -- confirmed in reports/ipp_removal_impact.csv), so this is
# a straightforward delete. No FK cascade risk. ─────────────────────────

def apply_ipp_removal(apply):
    ipp = OrganizationCategory.query.filter_by(category_name="IPP").first()
    if ipp is None:
        record(
            "organization_categories", "(none)", "category_name", "IPP", "(already absent)",
            "IPP category already didn't exist at write time -- nothing to delete.",
        )
        return
    org_count = Organization.query.filter_by(category_id=ipp.id).count()
    if apply:
        if org_count == 0:
            db.session.delete(ipp)
        else:
            raise RuntimeError(
                f"Refusing to delete IPP category -- {org_count} organizations reference it now "
                f"(was 0 when reports/ipp_removal_impact.csv was built). Re-check before retrying."
            )
    record(
        "organization_categories", ipp.id, "(entire row, DELETE)",
        f"category_name=IPP, description=NULL, 0 organizations referenced it", "DELETED",
        "Confirmed unused via reports/ipp_removal_impact.csv before deleting: 0 organizations, "
        "0 email_group_filters reference it. Durable removal also needs "
        "migrations/015_remove_ipp_category.sql (written separately) since ON CONFLICT DO NOTHING "
        "in migrations/009 would otherwise silently re-seed it on a fresh migration replay.",
    )


# ─── WRPC priority insert: Deepak Kumar (Member Secretary), 1 new Employee
# row. Employee creation has no live-app audit_log precedent either (the
# admin "add employee" route commits with no log_audit_event call, same
# gap as DirectoryNumber) -- session CSV only, same pattern. is_kmp=True is
# my own judgment call, not explicitly requested: WRPC's Chairman (Dr.
# Rohit Yadav, already in the app) is_kmp=True, and Member Secretary is the
# same tier of senior WRPC leadership -- flagged here so it's visible, not
# silently assumed. ─────────────────────────────────────────────────────

WRPC_ORG_ID = 6291


def apply_wrpc_insert(apply):
    new_emp = Employee(
        employee_name="Deepak Kumar",
        designation="Member Secretary",
        organization_id=WRPC_ORG_ID,
        office_phone="022- 28221636 Fax 022-28370193",
        email="ms-wrpc@nic.in",
        is_kmp=True,
        status="ACTIVE",
    )
    if apply:
        db.session.add(new_emp)
        db.session.flush()  # assigns new_emp.id for the audit row
    record(
        "employees", new_emp.id if apply else "(new)", "(entire row, INSERT)",
        "(did not exist)",
        f"employee_name=Deepak Kumar, designation=Member Secretary, organization_id={WRPC_ORG_ID}, "
        f"office_phone='022- 28221636 Fax 022-28370193', email=ms-wrpc@nic.in, is_kmp=True",
        "WRPC priority item: appears identically (same phone, same email) in both of Word's "
        "separate WRPC mentions -- consistent, not a stale duplicate. Confirmed absent from the "
        "app. Stephen Fernandes (the other Word-side WRPC 'Chairman' mention) was explicitly NOT "
        "added -- conflicts with the app's already-correct AdministrativeHead/Employee record for "
        "Dr. Rohit Yadav as Chairman. is_kmp=True is my own inference (Member Secretary is the "
        "same leadership tier as the Chairman, who is also is_kmp=True), not explicitly instructed.",
    )


# ─── Batch b: address fill-in (61 rows: 58 from Word paragraphs + 3 mined
# from banner title cells) ──────────────────────────────────────────────

def apply_batch_b_address(apply):
    with open("reports/fix_proposal_address.csv", encoding="utf-8") as f:
        proposals = list(csv.DictReader(f))

    org_by_name = {o.organization_name: o for o in Organization.query.all()}
    for r in proposals:
        org = org_by_name.get(r["organization"])
        if org is None:
            record(
                "organizations", "(none)", "address", "", r["proposed_address"],
                f"Could not find organization {r['organization']!r} in the live DB at write time -- skipped.",
            )
            continue
        old_value = org.address
        new_value = r["proposed_address"]
        if (old_value or "") == new_value:
            continue  # already applied (re-run) -- not a real change
        if apply:
            org.address = new_value
            log_audit_event(
                module="organizations", record_type="Organization", record_id=org.id,
                action="ADDRESS_FILL", field_name="address",
                old_value=old_value, new_value=new_value,
                reason=f"Approved per reports/fix_proposal_address.csv (source: {r['source']}).",
            )
        record(
            "organizations", org.id, "address", old_value, new_value,
            f"Batch b approved as proposed. Source: {r['source']}.",
        )


# ─── Item 4: NLDC/RLDC rename (5 rows, 0 deletions) ────────────────────────

RLDC_RENAMES = {
    7654: "NLDC (National Load Despatch Centre)",
    7655: "NRLDC (Northern Regional Load Despatch Centre)",
    7656: "ERLDC (Eastern Regional Load Despatch Centre)",
    7657: "SRLDC (Southern Regional Load Despatch Centre)",
    7658: "NERLDC (North Eastern Regional Load Despatch Centre)",
}


def apply_rldc_renames(apply):
    for row_id, new_name in RLDC_RENAMES.items():
        dn = db.session.get(DirectoryNumber, row_id)
        old_name = dn.name
        old_email = dn.email
        if apply:
            dn.name = new_name
        record(
            "directory_numbers", row_id, "name", old_name, new_name,
            f"possible_duplicate_row resolution: all 5 rows under org 6290 shared the generic name "
            f"'NLDC And RLDCs', which full_data_quality_audit.py's normalized-name grouping flagged "
            f"as a possible duplicate. Inspecting phone/email content showed each row is a distinct "
            f"real regional despatch centre (this one identified by email {old_email!r}), not a "
            f"copy-paste duplicate -- renamed to its specific centre instead of merging/deleting, "
            f"per reports/fix_proposal_duplicate_rows.csv. 0 rows deleted.",
        )


# ─── numbered_prefix strip (13 rows, confirmed inert -- see prior session's
# sort-order confirmation: no export/grouping/display logic anywhere
# depends on the leading number, and every flagged row is the sole record
# for its (organization, category)) ─────────────────────────────────────

NUMBERED_PREFIX_IDS = [
    7659, 7660, 7662, 7663, 7666, 7667, 7674, 7675, 7676, 7677, 7678, 7679, 7680,
]


def apply_numbered_prefix_strip(apply):
    for row_id in NUMBERED_PREFIX_IDS:
        dn = db.session.get(DirectoryNumber, row_id)
        old_name = dn.name
        new_name = LEADING_NUMBER_RE.sub("", old_name).strip()
        if new_name == old_name:
            continue  # shouldn't happen given the source list, but don't record a no-op
        if apply:
            dn.name = new_name
        record(
            "directory_numbers", row_id, "name", old_name, new_name,
            "numbered_prefix strip: leading Word-section-heading number removed from the name "
            "field. Confirmed inert before applying -- no export/grouping/display logic in "
            "routes/user_routes.py, services/pdf_generator.py, or services/excel_generator.py "
            "sorts by or depends on DirectoryNumber.name, and this row is the sole "
            "(organization_id, category) record for its station, so there was no sibling row the "
            "number could have been ordering against.",
        )


# ─── duplicated_phone: word_divergence batch (120 rows / 89 stations) and
# no_word_match batch (4 rows / 3 stations) -- both apply the same
# exact-repeat-segment dedup as the import_bugs batch already applied;
# read proposed_value from the already-generated proposal CSVs rather than
# recomputing, so this script and build_fix_proposals.py can't drift ──────

def _load_phone_proposals(path):
    with open(path, encoding="utf-8") as f:
        return {int(r["id"]): r for r in csv.DictReader(f)}


def apply_duplicated_phone_batch(apply, csv_path, origin_label, reason_detail):
    proposals = _load_phone_proposals(csv_path)
    for row_id, r in proposals.items():
        dn = db.session.get(DirectoryNumber, row_id)
        old_value = dn.phone_number
        new_value = r["proposed_value"]
        if new_value == old_value:
            continue  # already at the proposed value (e.g. re-run) -- nothing to record
        if apply:
            dn.phone_number = new_value
        record(
            "directory_numbers", row_id, "phone_number", old_value, new_value,
            f"duplicated_phone fix, origin={origin_label}: exact-repeat segment(s) removed, order "
            f"preserved, no other segment touched. {reason_detail}",
        )


def write_audit_csv():
    fieldnames = ["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(audit_rows)


def main():
    do_phone = "--apply-phone" in sys.argv
    do_rldc = "--apply-rldc" in sys.argv
    do_numbered_prefix = "--apply-numbered-prefix" in sys.argv
    do_word_divergence = "--apply-word-divergence" in sys.argv
    do_no_word_match = "--apply-no-word-match" in sys.argv
    do_cpsu = "--apply-cpsu" in sys.argv
    do_ipp_removal = "--apply-ipp-removal" in sys.argv
    do_nspcl = "--apply-nspcl" in sys.argv
    do_reclassification = "--apply-reclassification" in sys.argv
    do_wrpc = "--apply-wrpc" in sys.argv
    do_batch_b = "--apply-batch-b" in sys.argv
    apply = (
        do_phone or do_rldc or do_numbered_prefix or do_word_divergence or do_no_word_match
        or do_cpsu or do_ipp_removal or do_nspcl or do_reclassification or do_wrpc or do_batch_b
    )

    # test_request_context (not just app_context) -- apply_cpsu_moves calls
    # log_audit_event, which reads flask_login's current_user; that needs a
    # request context to resolve (to AnonymousUserMixin here, since this is
    # a script, not a logged-in admin) rather than raising.
    with app.test_request_context():
        if do_phone:
            apply_phone_fixes(apply=True)
        if do_rldc:
            apply_rldc_renames(apply=True)
        if do_cpsu:
            apply_cpsu_moves(apply=True)
        if do_ipp_removal:
            apply_ipp_removal(apply=True)
        if do_nspcl:
            apply_nspcl_move(apply=True)
        if do_reclassification:
            apply_state_sldc_rename(apply=True)
            apply_reclassification_org_moves(apply=True)
        if do_wrpc:
            apply_wrpc_insert(apply=True)
        if do_batch_b:
            apply_batch_b_address(apply=True)
        if do_numbered_prefix:
            apply_numbered_prefix_strip(apply=True)
        if do_word_divergence:
            apply_duplicated_phone_batch(
                apply=True,
                csv_path="reports/fix_proposal_duplicated_phone_word_divergence.csv",
                origin_label="originated_in_word/mixed",
                reason_detail=(
                    "Word itself already carries the same or a related issue for this station, but "
                    "approved decision: app data should be clean regardless of source-document "
                    "messiness."
                ),
            )
        if do_no_word_match:
            apply_duplicated_phone_batch(
                apply=True,
                csv_path="reports/fix_proposal_duplicated_phone_no_word_match.csv",
                origin_label="no_word_match",
                reason_detail=(
                    "Station never matched a Word segment, so origin vs. the source document "
                    "couldn't be determined -- applied anyway per approved decision: deduping "
                    "exact-repeat segments is safe independent of Word verification."
                ),
            )
        if apply:
            db.session.commit()

    if audit_rows:
        # merge with any prior run's rows already in the CSV so re-running
        # this script for item 4 after item 1 doesn't drop item 1's entries
        existing = []
        if os.path.exists(OUT_CSV):
            with open(OUT_CSV, encoding="utf-8") as f:
                existing = list(csv.DictReader(f))
        seen = {(r["table"], r["row_id"], r["field"]) for r in existing}
        merged = existing + [r for r in audit_rows if (r["table"], str(r["row_id"]), r["field"]) not in seen]
        fieldnames = ["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"]
        os.makedirs("reports", exist_ok=True)
        with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader()
            w.writerows(merged)

    print(f"{'APPLIED' if apply else 'DRY RUN (no writes)'} -- {len(audit_rows)} changes this run")
    for r in audit_rows:
        print(f"  [{r['table']}#{r['row_id']}] {r['field']}: {r['old_value']!r} -> {r['new_value']!r}")
    print(f"Audit trail written to {OUT_CSV}")


if __name__ == "__main__":
    main()
