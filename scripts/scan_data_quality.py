"""Read-only scans for two data-quality patterns in directory_numbers,
requested as a scope check BEFORE any fix is applied -- writes proposal
CSVs only, changes nothing.

1. Duplicated phone segments: directory_numbers has a single combined
   `phone_number` text field (there is no separate office_phone/mobile on
   this table -- that split exists on Employee, not DirectoryNumber). A
   value is flagged when splitting on common separators (/, ;, ,) yields the
   same normalized segment 2+ times, e.g. "X / X / X" -- not when it simply
   lists several genuinely different numbers.

2. Address-like content in the name field, table-wide (the Chhattisgarh fix
   earlier only covered the literal "Address:" pattern on 2 rows). Several
   independent signals are checked and reported per row so a human can
   judge each one; a proposed_new_value is only filled in when there's an
   unambiguous split point (a literal "Address:" marker) to cut at --
   otherwise it's left blank rather than guessing.

Usage:
  python scripts/scan_data_quality.py
"""

import csv
import os
import re
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from models.directory_number import DirectoryNumber

PHONE_OUT = "reports/duplicated_phone_fields.csv"
ADDRESS_OUT = "reports/address_in_name_field.csv"
SESSION_DATE = "2026-08-26"

SEGMENT_SPLIT_RE = re.compile(r"[/;,]")
PIN_CODE_RE = re.compile(r"-\s*\d{6}\b")
CITY_STATE_HINTS = (
    "maharashtra", "gujarat", "chhattisgarh", "madhya pradesh", "goa", "raipur",
    "bhopal", "mumbai", "ahmedabad", "vadodara", "surat", "indore", "bhilai",
    "raigad", "nagpur", "panaji", "danganiya",
)


def scan_duplicated_phone(numbers):
    rows = []
    for dn in numbers:
        for field_name, value in (("phone_number", dn.phone_number),):
            if not value:
                continue
            segments = [s.strip() for s in SEGMENT_SPLIT_RE.split(value) if s.strip()]
            if len(segments) < 2:
                continue
            counts = Counter(segments)
            repeat_count = max(counts.values())
            if repeat_count >= 2:
                rows.append({
                    "id": dn.id,
                    "station": dn.name,
                    "field": field_name,
                    "raw_value": value,
                    "detected_repeat_count": repeat_count,
                })
    return rows


def propose_address_split(name):
    if "Address:" in name:
        return name.split("Address:", 1)[0].strip()
    return ""


def scan_address_in_name(numbers):
    rows = []
    for dn in numbers:
        name = dn.name or ""
        signals = []
        if "address:" in name.lower():
            signals.append("contains 'Address:' marker")
        if PIN_CODE_RE.search(name):
            signals.append("contains a 6-digit PIN-code-like pattern")
        if name.count(",") >= 2:
            signals.append(f"{name.count(',')} commas (address-list structure)")
        if "backup" in name.lower():
            signals.append("contains 'Backup'")
        if "state load despatch centre" in name.lower() or "load despatch centre" in name.lower():
            signals.append("contains 'State Load Despatch Centre' / 'Load Despatch Centre'")
        lower = name.lower()
        hit_cities = [c for c in CITY_STATE_HINTS if c in lower]
        if hit_cities:
            signals.append(f"contains city/state name(s): {', '.join(hit_cities)}")
        if len(name) > 80:
            signals.append(f"unusually long name ({len(name)} chars)")

        if signals:
            proposed = propose_address_split(name)
            rows.append({
                "id": dn.id,
                "station": dn.name,
                "field": "name",
                "old_value": name,
                "proposed_new_value": proposed,
                "signals_detected": "; ".join(signals),
                "reason": "Flagged for review -- " + ("has an unambiguous 'Address:' split point"
                          if proposed else "no unambiguous split point found; needs manual judgment"),
                "timestamp": SESSION_DATE,
            })
    return rows


def main():
    with app.app_context():
        numbers = DirectoryNumber.query.all()

        phone_rows = scan_duplicated_phone(numbers)
        with open(PHONE_OUT, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["id", "station", "field", "raw_value", "detected_repeat_count"])
            w.writeheader()
            w.writerows(phone_rows)

        address_rows = scan_address_in_name(numbers)
        with open(ADDRESS_OUT, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=[
                "id", "station", "field", "old_value", "proposed_new_value",
                "signals_detected", "reason", "timestamp",
            ])
            w.writeheader()
            w.writerows(address_rows)

    print("=" * 70)
    print("  Data Quality Scan -- counts only, nothing changed")
    print("=" * 70)
    print(f"  Total directory_numbers rows scanned : {len(numbers)}")
    print(f"  Duplicated phone segments found      : {len(phone_rows)}  -> {PHONE_OUT}")
    print(f"  Address-like name-field rows found   : {len(address_rows)}  -> {ADDRESS_OUT}")
    print(f"    - with an unambiguous split point  : {sum(1 for r in address_rows if r['proposed_new_value'])}")
    print(f"    - needing manual judgment          : {sum(1 for r in address_rows if not r['proposed_new_value'])}")
    print("=" * 70)


if __name__ == "__main__":
    main()
