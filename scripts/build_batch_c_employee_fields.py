"""Full batch c proposal: all employee field mismatches from
reports/full_word_vs_app_field_diff.csv, with confirmed 10-digit
normalization applied to mobile numbers (see decision: the app's 98
existing non-blank Employee.mobile_phone values are ALL plain 10-digit,
zero exceptions -- not a majority, the only format in use). Read-only --
proposal CSV only, no DB writes.

Normalization rule for mobile:
  - digit sequence starts with '91' and is 12 digits total -> strip the
    country code, keep the last 10
  - digit sequence starts with a single leading '0' and is 11 digits total
    -> strip the leading 0, keep the last 10
  - already 10 digits -> use as-is
  - anything else (multiple numbers in one cell, e.g. "9898540640/
    9099939866", or a too-short/malformed fragment) -> NOT auto-normalized;
    flagged needs_manual_review with no proposed_value, per instruction
    not to guess which number to keep.

office_phone/designation/email are NOT mobile numbers -- proposed as the
raw Word value unchanged (no format convention established/requested for
those fields).

Usage:
  python scripts/build_batch_c_employee_fields.py
"""

import csv
import os
import re

DIFF_CSV = "reports/full_word_vs_app_field_diff.csv"
OUT_CSV = "reports/fix_proposal_batch_c_employee_fields.csv"

PLACEHOLDER = {"na", "n.a", "n.a.", "n/a", "-", "--", "nil", "none", "n. a.", ""}


def normalize_mobile(raw):
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 12 and digits.startswith("91"):
        return digits[2:], "normalized"
    if len(digits) == 11 and digits.startswith("0"):
        return digits[1:], "normalized"
    if len(digits) == 10:
        return digits, "already_10_digit"
    return None, "needs_manual_review"


def main():
    with open(DIFF_CSV, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    out = []
    for r in rows:
        if not r["field"].startswith("employee[") or r["match"] != "no":
            continue
        subfield = r["field"].split(".")[-1]
        word_value = r["word_value"]
        if word_value.strip().lower() in PLACEHOLDER:
            continue  # Word's own placeholder, not real data -- not proposed

        emp_name = r["field"][len("employee["):r["field"].rfind("].")]

        if subfield == "mobile":
            proposed, status = normalize_mobile(word_value)
            out.append({
                "organization": r["organization"], "employee": emp_name, "field": subfield,
                "current_value": r["app_value"], "raw_word_value": word_value,
                "proposed_value": proposed or "", "status": status,
            })
        else:
            out.append({
                "organization": r["organization"], "employee": emp_name, "field": subfield,
                "current_value": r["app_value"], "raw_word_value": word_value,
                "proposed_value": word_value, "status": "proposed_as_is",
            })

    fieldnames = ["organization", "employee", "field", "current_value", "raw_word_value", "proposed_value", "status"]
    os.makedirs("reports", exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(out)

    from collections import Counter
    print("=" * 70)
    print(f"  Batch c full proposal: {len(out)} rows -> {OUT_CSV}")
    print("=" * 70)
    by_field = Counter(r["field"] for r in out)
    print("  by field:", dict(by_field))
    mobile_rows = [r for r in out if r["field"] == "mobile"]
    print("  mobile status breakdown:", Counter(r["status"] for r in mobile_rows))
    print("=" * 70)


if __name__ == "__main__":
    main()
