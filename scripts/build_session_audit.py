"""One-off script that compiles reports/session_changes_audit.csv from the
values captured directly in-session (SELECT output read immediately before
each write, plus the xmin-verified list for the bulk category update).

Not a general-purpose audit tool -- there is no trigger-based change history
on this database, so this is a manual, session-specific reconstruction. See
the accompanying chat message for the caveats on timestamp precision and on
these writes not going through services/audit_service.py (see AUDIT_LOG_NOTE
in the header row written below).
"""

import csv

OUT = "reports/session_changes_audit.csv"
SESSION_DATE = "2026-08-26"

rows = []


def add(table, row_id, field, old_value, new_value, reason):
    rows.append({
        "table": table, "row_id": row_id, "field": field,
        "old_value": old_value, "new_value": new_value,
        "reason": reason, "timestamp": SESSION_DATE,
    })


# 1. Migration 014 -- schema change (not a row-level write, included for
# completeness since it was a live database write this session).
add(
    "directory_numbers", "(schema)", "switch_yard, control_room, ip_address",
    "columns did not exist",
    "columns added (switch_yard BOOLEAN DEFAULT false, control_room BOOLEAN DEFAULT false, ip_address VARCHAR(45))",
    "Applied migrations/014_directory_number_contact_fields.sql -- the model already "
    "referenced these columns, causing 500 errors on /email-lists and /telephone-directory.",
)

# 2. Nuclear category fix
for org_id, name in [
    (6359, "Tarapur Atomic Power Station Stg-1 & 2"),
    (6360, "Tarapur Atomic Power Station Stg-3 & 4"),
    (6361, "Kakrapar Atomic Power Station Stg-1 & 2"),
    (6362, "Kakrapar Atomic Power Station Stg-3 & 4"),
]:
    add(
        "organizations", org_id, "category_id",
        "17 (Others)", "20 (Nuclear)",
        f"{name}: parent org 'Nuclear Power Stations In Western Region' is correctly "
        "categorized Nuclear; this child had defaulted to Others on import and was never reclassified.",
    )

# 3. Bulk org category reassignment (91 rows, verified via xmin against the
# UPDATE's own transaction id -- see chat for how the earlier "94" estimate
# was corrected to 91).
with open("reports/_tmp_94rows.csv", newline="", encoding="utf-8") as f:
    for r in csv.DictReader(f):
        add(
            "organizations", r["id"], "category_id",
            "Others", r["new_category"],
            "Bulk fix: child organization's category defaulted to Others on import while "
            "its parent already held a specific, correct category; verified safe because "
            "the parent is a real category (not a generic grouping bucket) and the child "
            "had no independently-assigned category of its own.",
        )

# 4. ACB (India) Limited duplicate delete
add(
    "directory_numbers", 7799, "(entire row, DELETE)",
    'name="400/132 kV Switch Yard Control Room", category="Control Room", '
    'phone_number="400/132 KV Switch yard Control Room; 09302187737 / Orange no: 20221416 / '
    '09302187737 / Orange no: 20221416", email="switchyard.270@acbindia.com", org_id=6438',
    "DELETED",
    "Duplicate of id 7873 (same name, same org, same email) -- 7873 is a later, cleaner "
    "re-entry correctly categorized Switchyard; 7799 was the older Control-Room-mistagged "
    "row with duplicated/garbled phone text.",
)

# 5. 4 more duplicate deletes (same pattern as ACB)
dupe_rows = [
    (7815, "Mahan Energen Limited (Formally Essar Power (MP) Limited", "Switchyard Control Room (SCR)",
     "8966902950 / NA / 8966902950", "scr.mel@adani.com", 7874),
    (7826, "Raigarh Energy Generation Limited", "Switchyard Control Room",
     "6232005449", "reglshift.operation@adani.com", 7875),
    (7832, "SASAN Power Limited", "Switchyard Control Room",
     "Switchyard Control Room; 7583896091 / 7583896091 / 7583896091",
     "7583896091; sasan.switchyard@relianceada.com", 7876),
    (7837, "TRN Energy Private Limited", "Switchyard Control Room",
     "6265223817 / 6265223817 / 6265223817", "400kvswitchyard@trnenergy.com", 7877),
]
for row_id, org, name, phone, email, kept_id in dupe_rows:
    add(
        "directory_numbers", row_id, "(entire row, DELETE)",
        f'org="{org}", name="{name}", category="Control Room", phone_number="{phone}", email="{email}"',
        "DELETED",
        f"Same pattern as ACB (India) Limited (id 7799): duplicate of id {kept_id}, which is a "
        "later, cleaner re-entry correctly categorized Switchyard. IDs 7874-7877 are sequential "
        "with 7873 (ACB's clean row), confirming a single later correction batch never deleted "
        "the originals it was replacing.",
    )

# 6. 3 phone_number placeholder clears
phone_clears = [
    (7710, "Avikiran Solar India Private Limited", "SCADA VOIP/ PSS VOIP: / SCADA VOIP/ PSS VOIP:"),
    (7749, "Control Room", "SCADA VOIP/ PSS VOIP: / SCADA VOIP/ PSS VOIP:"),
    (7798, "ACB (India) Limited", "Control Room"),
]
for row_id, name, old_phone in phone_clears:
    add(
        "directory_numbers", row_id, "phone_number",
        old_phone, "NULL",
        f'Row name "{name}": phone_number field contained placeholder/label text with no actual '
        "phone number (no digits at all) -- cleared rather than leaving misleading text where a "
        "real number should be. No real number was recoverable from available sources.",
    )

# 7. Chhattisgarh SLDC control room name trims
name_trims = [
    (7859,
     "Control Room 1 : State Load Despatch Centre Chhattisgarh Address: Chhattisgarh State Power "
     "Transmission Co. Ltd. (A Successor Company Of CSEB), Danganiya, Raipur, C.G. - 492013, Chhattisgarh.",
     "Control Room 1 : State Load Despatch Centre Chhattisgarh",
     "Embedded address matched the organization's own stored address (Danganiya, Raipur) almost "
     "exactly -- fully redundant, safe to drop entirely."),
    (7860,
     "Control Room 2 : Backup State Load Despatch Centre, Chhattisgarh Address: Chhattisgarh State "
     "Power Transmission Co. Ltd. (A Successor Company Of CSEB), Khedamara, Bhilai, C.G. - 490024,Chhattisgarh.",
     "Control Room 2 : Backup State Load Despatch Centre, Chhattisgarh (Khedamara, Bhilai)",
     "Embedded address (Khedamara, Bhilai) is GENUINELY DIFFERENT from the organization's own "
     "stored address (Danganiya, Raipur) -- the backup control room really is at a different "
     "site, so the location detail was kept in shortened form rather than dropped."),
]
for row_id, old_name, new_name, reason in name_trims:
    add("directory_numbers", row_id, "name", old_name, new_name, reason)

with open(OUT, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["table", "row_id", "field", "old_value", "new_value", "reason", "timestamp"])
    w.writeheader()
    w.writerows(rows)

print(f"Wrote {len(rows)} rows to {OUT}")
