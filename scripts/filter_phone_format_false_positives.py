"""Re-checks every phone-type match="no" row in
reports/full_word_vs_app_field_diff_FINAL.csv for country-code / leading-zero
formatting differences before counting it as a genuine mismatch -- same
normalization used earlier this session, reproduced here since it was done
inline rather than as a saved script last time.

full_word_vs_app_field_diff.py's own phone_field_matches() only checks
"word's digit string is a literal substring of app's digit string." That
already handles Word adding surrounding text, but NOT:
  - country code: word "+91 9876543210" (12 digits incl 91) vs app plain
    "9876543210" (10 digits) -- word is not a substring of app.
  - leading zero: word "09876543210" (11 digits, STD-dialing leading 0) vs
    app "9876543210" (10 digits) -- same, word is longer.

Read-only -- writes reports/phone_format_false_positives.csv only. Does not
modify full_word_vs_app_field_diff_FINAL.csv or the database.

Usage:
  python scripts/filter_phone_format_false_positives.py
"""

import csv
import re
from collections import Counter

IN_CSV = "reports/full_word_vs_app_field_diff_FINAL.csv"
OUT_CSV = "reports/phone_format_false_positives.csv"

DIGIT_RUN_RE = re.compile(r"\d{5,}")

PHONE_FIELD_SUFFIXES = (".office_phone", ".mobile", ".residence_phone")


def normalize_number(digits):
    """Strip a leading country code (91) and/or leading zeros (STD trunk
    prefix) down to a bare 10-digit core, when the run is long enough to
    contain one. Numbers already 10 digits or shorter are returned as-is --
    nothing to strip."""
    d = digits
    while len(d) > 10 and d.startswith("0"):
        d = d[1:]
    if len(d) > 10 and d.startswith("91") and len(d) - 2 >= 10:
        d = d[2:]
    while len(d) > 10 and d.startswith("0"):
        d = d[1:]
    return d[-10:] if len(d) >= 10 else d


def number_set(value):
    return {normalize_number(m) for m in DIGIT_RUN_RE.findall(value or "")}


def is_phone_field(field):
    return any(field.endswith(suf) for suf in PHONE_FIELD_SUFFIXES)


def main():
    with open(IN_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    false_positives = []
    for r in rows:
        if r["match"] != "no" or not is_phone_field(r["field"]):
            continue
        word_nums = number_set(r["word_value"])
        app_nums = number_set(r["app_value"])
        if word_nums and app_nums and (word_nums & app_nums):
            false_positives.append(r)

    fieldnames = ["organization", "field", "word_value", "app_value", "match"]
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(false_positives)

    raw_no_phone = [r for r in rows if r["match"] == "no" and is_phone_field(r["field"])]
    raw_no_other = [r for r in rows if r["match"] == "no" and not is_phone_field(r["field"])]

    print("=" * 70)
    print("  Phone-format false-positive filter (country code / leading zero)")
    print("=" * 70)
    print(f"  Raw match=no phone-field rows      : {len(raw_no_phone)}")
    print(f"  -> reclassified as formatting-only  : {len(false_positives)}  -> {OUT_CSV}")
    print(f"  -> genuine phone mismatches remain  : {len(raw_no_phone) - len(false_positives)}")
    print(f"  Raw match=no non-phone-field rows  : {len(raw_no_other)} (unaffected by this filter)")
    print("=" * 70)


if __name__ == "__main__":
    main()
