"""Generates: the WRPC priority proposal (item 1), full batch a
(organization_name) and batch b (address) proposals with real reasoning
per row (not just raw Word text), and splits batch d (word_only) into
three separate files by row type. Read-only -- proposal CSVs only.

Usage:
  python scripts/build_batches_ab_and_wrpc.py
"""

import csv
import os

DIFF_CSV = "reports/full_word_vs_app_field_diff.csv"


def load_rows():
    with open(DIFF_CSV, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_csv(path, rows, fieldnames):
    os.makedirs("reports", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


# ─── WRPC priority proposal ─────────────────────────────────────────────

def build_wrpc_priority():
    rows = [
        {
            "organization": "Western Regional Power Committee", "action": "insert",
            "name": "Deepak Kumar", "designation": "Member Secretary",
            "office_phone": "022- 28221636 Fax 022-28370193", "mobile": "",
            "email": "ms-wrpc@nic.in",
            "reasoning": "Genuine gap: appears identically (same phone, same email) in BOTH of "
                         "Word's two WRPC mentions -- consistent, not a stale duplicate. Confirmed "
                         "absent from the app's Employee table for org 6291. Safe to insert.",
        },
        {
            "organization": "Western Regional Power Committee", "action": "DO NOT ADD -- flagged",
            "name": "Sh. Stephen Fernandes", "designation": "Chairman, WRPC Chief Electrical Engineer, Goa",
            "office_phone": "", "mobile": "", "email": "",
            "reasoning": "CONFLICTS with data the app already has correct: AdministrativeHead for "
                         "org 6291 already links to Employee 'Dr. Rohit Yadav' (chairmancspc@gmail.com) "
                         "as Chairman, role_title 'Chairman, WRPC & Secretary (Power), DNH & DD' -- "
                         "and Word's OWN more detailed 6-person WRPC roster table (the one with a "
                         "proper Heading-style section, not a banner) independently lists Rohit Yadav "
                         "as Chairman too, matching the app exactly. Stephen Fernandes only appears in "
                         "a separate, shorter 2-person banner-style mention elsewhere in the same "
                         "document -- looks like stale/leftover text in Word itself, not a real "
                         "current gap. Adding him would create a false second Chairman. Flagging for "
                         "your judgment rather than importing possibly-wrong leadership data.",
        },
    ]
    write_csv("reports/fix_proposal_wrpc_priority.csv", rows,
              ["organization", "action", "name", "designation", "office_phone", "mobile", "email", "reasoning"])
    return rows


# ─── Batch a: organization_name, with real proposed values ─────────────

# Every row resolved to "no change" on inspection -- see reasoning per row.
# 3 rows reveal address text embedded in the Word banner itself (not a
# separate paragraph, so batch b's extractor missed it) -- ported into
# batch b below as bonus rows rather than silently dropped.
ORG_NAME_DECISIONS = {
    "Madhya Pradesh State Load Despatch Centre": (
        "no_change", "Decorative quote ('None of us is as smart as all of us... - Ken Blanchard') "
        "glued onto the real name in the same Word cell. App's name is already correct."),
    "Western Regional Power Committee": (
        "no_change", "Word adds ', MUMBAI' city suffix; app's existing naming convention across "
        "other orgs doesn't include city suffixes. Cosmetic, not an error."),
    "Eastern Regional Load Despatch Centre , Kolkata": (
        "no_change", "Decorative quote ('We = Power') glued onto the real name. App name already correct."),
    "Khopoli & Bhivpuri": (
        "no_change", "Word prefix 'MAHARASHTRA 17.' is a state label + section number, both "
        "artifacts, not part of the real name."),
    "NTPC Gadarwara": (
        "no_change_ported_to_address", "Extra text is a full descriptive address ('Super Thermal "
        "Power Project, Narsinghpur, Madhya Pradesh-487 001'), not a name variant -- belongs in the "
        "address field. Added to batch b."),
    "NTPC Khargone": (
        "no_change_ported_to_address", "Same pattern as NTPC Gadarwara -- address text embedded in "
        "the banner cell. Added to batch b."),
    "NTPC Kawas": (
        "no_change", "Word is LESS specific here (just 'Kawas') than the app's 'NTPC Kawas' -- "
        "app's fuller name is correct; no action needed."),
    "AVAADA Sunshine Energy Private Limited": (
        "no_change", "Word's '200MW' capacity prefix isn't part of the official org name (matches "
        "the app's existing convention of excluding capacity annotations elsewhere)."),
    "Tata Power Mumbai": (
        "no_change", "Stray leading '. ' in Word's heading paragraph -- looks like a formatting "
        "artifact (an auto-number field that rendered as bare punctuation), not real content."),
    "NTPC Gandhar": (
        "no_change_ported_to_address", "Same pattern as NTPC Gadarwara -- address text embedded in "
        "the banner cell. Added to batch b."),
    "Adani Power Limited": (
        "no_change", "Same stray leading '. ' artifact as Tata Power Mumbai."),
}


def build_batch_a(rows):
    out = []
    for r in rows:
        if r["field"] != "organization_name" or r["match"] != "no":
            continue
        decision, reasoning = ORG_NAME_DECISIONS.get(r["organization"], ("needs_review", "Not pre-classified -- review manually."))
        out.append({
            "organization": r["organization"], "current_name": r["app_value"],
            "word_raw_heading": r["word_value"], "decision": decision, "reasoning": reasoning,
        })
    write_csv("reports/fix_proposal_org_name.csv", out,
              ["organization", "current_name", "word_raw_heading", "decision", "reasoning"])
    return out


# ─── Batch b: address, full proposal + 3 bonus banner-embedded addresses ──

BANNER_EMBEDDED_ADDRESSES = {
    "NTPC Gadarwara": "NTPC Gadarwara Super Thermal Power Project, Narsinghpur, Madhya Pradesh-487 001.",
    "NTPC Khargone": "NTPC Khargone Super Thermal Power Project, Vishnupuri, Khandwa Road, Khargone, Madhya Pradesh-451 001",
    "NTPC Gandhar": "NTPC Gandhar Gas Power Project, P.O.: NTPC township, Dist. Bharuch, Gujarat- 392215",
}


def build_batch_b(rows):
    out = []
    for r in rows:
        if r["field"] != "address" or r["match"] != "no":
            continue
        out.append({
            "organization": r["organization"], "current_address": r["app_value"],
            "proposed_address": r["word_value"], "source": "separate Word paragraph",
        })
    for org, addr in BANNER_EMBEDDED_ADDRESSES.items():
        out.append({
            "organization": org, "current_address": "",
            "proposed_address": addr, "source": "extracted from banner title cell (not a separate paragraph)",
        })
    write_csv("reports/fix_proposal_address.csv", out,
              ["organization", "current_address", "proposed_address", "source"])
    return out


# ─── Batch d split into 3 files ─────────────────────────────────────────

def build_batch_d_split(rows):
    employees, control_rooms, orgs = [], [], []
    for r in rows:
        if r["match"] != "word_only":
            continue
        row = {"organization": r["organization"], "word_value": r["word_value"], "proposed_action": "insert new row"}
        if r["field"] == "(employee row)":
            employees.append(row)
        elif r["field"] == "(control_room row)":
            control_rooms.append(row)
        elif r["field"] == "(organization)":
            orgs.append(row)
    fieldnames = ["organization", "word_value", "proposed_action"]
    write_csv("reports/fix_proposal_word_only_employees.csv", employees, fieldnames)
    write_csv("reports/fix_proposal_word_only_control_rooms.csv", control_rooms, fieldnames)
    write_csv("reports/fix_proposal_word_only_organizations.csv", orgs, fieldnames)
    return employees, control_rooms, orgs


def main():
    rows = load_rows()

    wrpc = build_wrpc_priority()
    a = build_batch_a(rows)
    b = build_batch_b(rows)
    emp, cr, org = build_batch_d_split(rows)

    print("=" * 70)
    print(f"  WRPC priority: {len(wrpc)} rows -> reports/fix_proposal_wrpc_priority.csv")
    print(f"  Batch a (org_name): {len(a)} rows -> reports/fix_proposal_org_name.csv")
    print(f"  Batch b (address): {len(b)} rows ({len(b)-3} from Word paragraphs + 3 bonus banner-embedded) -> reports/fix_proposal_address.csv")
    print(f"  Batch d employees: {len(emp)} -> reports/fix_proposal_word_only_employees.csv")
    print(f"  Batch d control_rooms: {len(cr)} -> reports/fix_proposal_word_only_control_rooms.csv")
    print(f"  Batch d organizations: {len(org)} -> reports/fix_proposal_word_only_organizations.csv")
    print("=" * 70)


if __name__ == "__main__":
    main()
