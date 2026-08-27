"""Merge-plan proposal for the 12 possible_duplicate_row rows from
full_data_quality_audit.py (2 groups: 5x "NLDC And RLDCs" under org 6290,
2x "SLDC-Control room" under org 6502 / Kalwa, Thane-Belapur). Read-only --
writes reports/fix_proposal_duplicate_rows.csv only, no DB writes.

IMPORTANT DEVIATION from the requested pattern: the 5 duplicate-record
fixes from the earlier session (ACB India, Mahan Energen, etc. -- see
scripts/build_session_audit.py) were byte-for-byte near-identical rows: same
name, same org, same email, redundant phone text -- genuine copy-paste
duplicates where deleting one and keeping the other lost nothing. Inspecting
this group's actual phone/email content (not just the name match that
possible_duplicate_row grouped on) shows that is NOT the case here:

  - "NLDC And RLDCs" (5 rows): each row's email domain identifies a
    DIFFERENT real-world regional load despatch centre -- posococc@/
    systemoperationd@/nldccr@ (NLDC itself), nrldccr@/nrldcso@ (Northern
    RLDC), erldc@ (Eastern RLDC), srldccr@/srldccroom@/croom@srldc.org
    (Southern RLDC), nerldc@/nerldc1@ (North Eastern RLDC). These are 5
    distinct, real contact records that all happen to share the same
    generic placeholder name -- deleting any of them the way ACB's
    duplicate was deleted would destroy real contact data for a real RLDC.
    The actual defect is a NAMING problem (all 5 share one unhelpfully
    generic label), not a duplicate-ROW problem. Proposed fix: rename each
    row to the specific centre its email identifies; delete none.

  - "SLDC-Control room" (2 rows, Kalwa/Thane-Belapur): same email on both
    rows, but row 7682's phone text explicitly starts "Fax 022-27601769",
    distinct from row 7681's main line numbers -- consistent with these
    being a main line and a fax line sharing a name, not a redundant
    re-entry. Genuinely ambiguous either way from the data alone (could
    also be an erroneous partial re-entry), so this one is left for human
    judgment rather than resolved by inference the way the RLDC group was.

Usage:
  python scripts/build_duplicate_row_merge_plan.py
"""

import csv
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.directory_number import DirectoryNumber

OUT_CSV = "reports/fix_proposal_duplicate_rows.csv"

# email-domain/local-part signal -> the specific centre it identifies
RLDC_LABELS = [
    (("posococc", "systemoperationd", "nldccr"), "NLDC (National Load Despatch Centre)"),
    (("nrldccr", "nrldcso"), "NRLDC (Northern Regional Load Despatch Centre)"),
    (("erldc",), "ERLDC (Eastern Regional Load Despatch Centre)"),
    (("srldccr", "srldccroom", "croom"), "SRLDC (Southern Regional Load Despatch Centre)"),
    (("nerldc", "nerldc1"), "NERLDC (North Eastern Regional Load Despatch Centre)"),
]

RLDC_IDS = [7654, 7655, 7656, 7657, 7658]
KALWA_IDS = [7681, 7682]


def label_for_email(email):
    email_lower = (email or "").lower()
    for tokens, label in RLDC_LABELS:
        if any(email_lower.startswith(t + "@") or f" {t}@" in email_lower or email_lower.split()[0].startswith(t + "@") for t in tokens):
            return label
    for tokens, label in RLDC_LABELS:
        if any(t in email_lower for t in tokens):
            return label
    return None


def build():
    with app.app_context():
        rows = []

        for rid in RLDC_IDS:
            dn = DirectoryNumber.query.get(rid)
            proposed_name = label_for_email(dn.email)
            rows.append({
                "id": dn.id, "station": "NLDC And RLDCs (org 6290)", "category": "control_room",
                "current_name": dn.name, "phone": dn.phone_number, "email": dn.email,
                "proposed_action": "rename_distinct" if proposed_name else "needs_human_review",
                "proposed_new_name": proposed_name or "",
                "matched_duplicate_id": ", ".join(str(i) for i in RLDC_IDS if i != rid),
                "reasoning": (
                    f"Email identifies this as {proposed_name} -- a distinct real record, not a "
                    "copy-paste duplicate of the other 4 rows; only the generic name matched. "
                    "Do not delete."
                    if proposed_name else
                    "Could not confidently map this email to a specific RLDC from the known "
                    "signal list -- needs a human to confirm before renaming."
                ),
            })

        for rid in KALWA_IDS:
            dn = DirectoryNumber.query.get(rid)
            rows.append({
                "id": dn.id, "station": "Kalwa, Thane-Belapur (org 6502)", "category": "control_room",
                "current_name": dn.name, "phone": dn.phone_number, "email": dn.email,
                "proposed_action": "needs_human_review",
                "proposed_new_name": "",
                "matched_duplicate_id": ", ".join(str(i) for i in KALWA_IDS if i != rid),
                "reasoning": (
                    "Same email as its sibling row, but this row's phone text is prefixed 'Fax ...' "
                    "-- consistent with a genuine separate fax-line record sharing the same generic "
                    "name (rename, don't delete), but could also be a partial erroneous re-entry "
                    "(merge/delete candidate). Content alone doesn't resolve which -- flagging for "
                    "your call rather than guessing which interpretation is right."
                    if dn.phone_number and dn.phone_number.strip().lower().startswith("fax")
                    else
                    "Sibling of the Fax-prefixed row above; same ambiguity applies -- needs human "
                    "judgment before any merge/delete/rename."
                ),
            })

    fieldnames = ["id", "station", "category", "current_name", "phone", "email",
                  "proposed_action", "proposed_new_name", "matched_duplicate_id", "reasoning"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)

    print(f"Wrote {len(rows)} rows to {OUT_CSV}")
    print("Deletions proposed: 0 -- neither group is a genuine copy-paste duplicate on inspection.")
    return rows


if __name__ == "__main__":
    build()
