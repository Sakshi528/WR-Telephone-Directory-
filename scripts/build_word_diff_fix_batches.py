"""Splits reports/full_word_vs_app_field_diff.csv into the review-before-
write sub-batches requested: (a) organization_name, (b) address, (c) a
20-row SAMPLE of employee.mobile only (largest sub-category, held to a
sample pending sanity-check before generating all ~638), (d) word_only
rows as new-row insertion candidates. (e) app_only is deliberately NOT
proposed for deletion -- counts only, same caveat as the original
employee-gap analysis from earlier this session (some may be legitimate
post-Word-doc additions).

Read-only -- writes proposal CSVs only, no DB writes.

Usage:
  python scripts/build_word_diff_fix_batches.py
"""

import csv
import os
from collections import Counter

DIFF_CSV = "reports/full_word_vs_app_field_diff.csv"

PLACEHOLDER = {"na", "n.a", "n.a.", "n/a", "-", "--", "nil", "none", "n. a.", ""}


def load_rows():
    with open(DIFF_CSV, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_csv(path, rows, fieldnames):
    os.makedirs("reports", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def batch_a_org_name(rows):
    out = []
    for r in rows:
        if r["field"] != "organization_name" or r["match"] != "no":
            continue
        out.append({
            "organization": r["organization"], "current_name": r["app_value"],
            "word_raw_heading": r["word_value"],
        })
    write_csv("reports/fix_proposal_org_name.csv", out,
               ["organization", "current_name", "word_raw_heading"])
    return out


def batch_b_address(rows):
    out = []
    for r in rows:
        if r["field"] != "address" or r["match"] != "no":
            continue
        out.append({
            "organization": r["organization"], "current_address": r["app_value"],
            "proposed_address": r["word_value"],
        })
    write_csv("reports/fix_proposal_address.csv", out,
               ["organization", "current_address", "proposed_address"])
    return out


def batch_c_mobile_sample(rows, n=20):
    out = []
    for r in rows:
        if not r["field"].startswith("employee[") or not r["field"].endswith(".mobile") or r["match"] != "no":
            continue
        if r["word_value"].strip().lower() in PLACEHOLDER:
            continue
        emp_name = r["field"][len("employee["):r["field"].rfind("].")]
        out.append({
            "organization": r["organization"], "employee": emp_name,
            "current_mobile": r["app_value"], "proposed_mobile": r["word_value"],
        })
    sample = out[:n]
    write_csv("reports/sample_fix_proposal_employee_mobile.csv", sample,
               ["organization", "employee", "current_mobile", "proposed_mobile"])
    return out, sample


def batch_d_word_only(rows):
    out = []
    for r in rows:
        if r["match"] != "word_only":
            continue
        out.append({
            "organization": r["organization"], "row_type": r["field"],
            "word_value": r["word_value"], "proposed_action": "insert new row",
        })
    write_csv("reports/fix_proposal_word_only_new_rows.csv", out,
               ["organization", "row_type", "word_value", "proposed_action"])
    return out


def batch_e_app_only_counts(rows):
    app_only = [r for r in rows if r["match"] == "app_only"]

    def kind(field):
        if field == "(employee row)": return "employee_row"
        if field == "(control_room row)": return "control_room_row"
        if field == "(switch_yard row)": return "switch_yard_row"
        if field == "(organization)": return "organization_row"
        return field
    return Counter(kind(r["field"]) for r in app_only)


def main():
    rows = load_rows()

    a = batch_a_org_name(rows)
    b = batch_b_address(rows)
    c_all, c_sample = batch_c_mobile_sample(rows)
    d = batch_d_word_only(rows)
    e_counts = batch_e_app_only_counts(rows)

    print("=" * 70)
    print("  Fix-proposal sub-batches (nothing applied)")
    print("=" * 70)
    print(f"  a. organization_name : {len(a)} rows -> reports/fix_proposal_org_name.csv")
    print(f"  b. address            : {len(b)} rows -> reports/fix_proposal_address.csv")
    print(f"  c. employee.mobile    : {len(c_all)} total real mismatches; "
          f"{len(c_sample)}-row SAMPLE written -> reports/sample_fix_proposal_employee_mobile.csv")
    print(f"  d. word_only new rows : {len(d)} rows -> reports/fix_proposal_word_only_new_rows.csv")
    print(f"  e. app_only (NOT proposed for deletion, counts only): {dict(e_counts)}")
    print("=" * 70)


if __name__ == "__main__":
    main()
