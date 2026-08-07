"""Standardizes text formatting (capitalization, spacing, separators, email
case) across WR_DB_Ready_Final_Verified_v2.xlsx. Read-only on the input.
Writes WR_DB_Ready_Final_Verified_v3.xlsx and
Formatting_Normalization_Log.xlsx.

Scope, per explicit confirmation:
- Normalized (title-case + abbreviation preservation): organization/sub-
  organization names, employee/KMP/Utility Head designations, control room
  and switchyard labels (incl. the denormalized organization_name/
  sub_organization_name columns), Hospital names, Emergency Service
  names/types.
- Lowercased only: every email column.
- Left untouched: all phone number columns, all IDs/FKs, department, and
  every person NAME field (employees.name, KMP.name, Utility_Heads.name)
  -- person names have exceptions (initials, honorifics, IAS-style
  suffixes) these rules don't address, so they're out of scope by design.
- Reference_Sections and db_schema are left untouched entirely -- not
  named in any rule category and too heterogeneous (mixes org-like and
  person-like names) to normalize safely with one rule set.

Usage:
  python scripts/normalize_formatting.py
"""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill

INPUT_XLSX = "WR_DB_Ready_Final_Verified_v2.xlsx"
OUTPUT_XLSX = "WR_DB_Ready_Final_Verified_v3.xlsx"
LOG_XLSX = "Formatting_Normalization_Log.xlsx"

LOG_HEADER = ["Sheet Name", "Primary Key", "Column Name", "Old Value", "New Value", "Reason"]

# ─── ABBREVIATION PRESERVATION ────────────────────────────────────────────────
#
# The user's explicit list covers the most important technical abbreviations,
# but real designation text turned out to use dozens more not on that list
# (AGM, ACE, AVP, ALDC, MCR, PH, SO, EEMG, O&M, ...) -- there is no way to
# enumerate every power-sector initialism in advance. So the model is
# inverted: any short token that was ALREADY all-uppercase in the source and
# isn't a recognized ordinary English word is treated as an inferred
# abbreviation and left uppercase, rather than trying to list every
# abbreviation that must be preserved. This trades a small risk of leaving
# an unlisted real word in caps against the much worse alternative already
# observed: mangling real abbreviations into nonsense ("CEO" -> "Ceo").

ABBREVIATIONS_UPPER = {
    "AIS", "GIS", "HVDC", "HVAC", "SCADA", "SLDC", "RLDC", "NLDC", "WRLDC", "NRLDC", "ERLDC",
    "SRLDC", "NERLDC", "PGCIL", "CTU", "ISTS", "STATCOM", "FACTS", "SVC", "OLTC", "DVC",
    "MPPMCL", "MSETCL", "GETCO", "CSPTCL", "OPTCL", "MPPTCL", "MSLDC", "EHV", "HV", "LV",
    "MW", "MVAR", "KVAR", "DC", "AC", "UPS", "RTU", "PLCC", "PMU", "DCB", "VCB", "SF6",
    "CT", "PT", "CVT", "ABT", "DSM", "RLNG", "GAIL", "NTPC", "NHPC", "NPCIL", "POWERGRID",
    "GMR", "JSW", "APL", "IPP", "MVA",
}
SLASH_TOKENS = {"S/S", "S/Y"}

# Ordinary English/directory words that must still be title-cased even when
# the source wrote them in ALL CAPS (checked before the "infer as
# abbreviation" fallback) -- deliberately generous, since missing a word
# here just means it stays uppercase (a visible, reviewable no-op), while a
# false entry here would wrongly title-case a real abbreviation.
ENGLISH_WORDS = {
    "control", "room", "rooms", "station", "stations", "switchyard", "switch", "yard",
    "region", "regions", "western", "eastern", "northern", "southern", "central", "other",
    "main", "new", "old", "state", "india", "phase", "project", "projects", "mega", "atomic",
    "nuclear", "thermal", "hydro", "hydel", "solar", "wind", "energy", "power", "plant",
    "plants", "load", "despatch", "dispatch", "network", "operation", "operations",
    "maintenance", "commercial", "technical", "communication", "communications",
    "scheduling", "backup", "office", "law", "department", "corporation", "company",
    "companies", "limited", "private", "transmission", "distribution", "generation",
    "generator", "generators", "generating", "connected", "buyers", "board", "committee",
    "regulatory", "development", "accounting", "port", "ports", "special", "economic",
    "zone", "holding", "service", "services", "trust", "group", "steel", "cement", "cements",
    "health", "care", "medical", "hospital", "hospitals", "research", "centre", "center",
    "institute", "facility", "campus", "floor", "address", "engineer", "engineers",
    "manager", "managers", "chief", "general", "deputy", "assistant", "senior", "executive",
    "additional", "joint", "principal", "officer", "officers", "director", "directors",
    "head", "secretary", "president", "vice", "associate", "important", "vendors", "hotline",
    "hotlines", "nos", "orange", "directory", "nodal", "provider", "providers", "charge",
    "incharge", "substation", "important", "blood", "bank", "fire", "police", "disaster",
    "management", "emergency", "contact", "numbers", "grid", "controller", "committee",
    "and", "of", "for", "the", "at", "in", "on", "to", "a", "an", "with", "by",
    # place/company names and ordinary words observed staying incorrectly
    # uppercase after a full pass over the normalized workbook
    "madhya", "pradesh", "tiroda", "adani", "tata", "thivim", "nagpur", "cyber",
    "junior", "system", "data", "red",
    "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
    "eleven", "twelve", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
    "eighty", "ninety", "hundred",
}
KV_RE = re.compile(r'(\d+)\s*-?\s*kv\b', re.IGNORECASE)
HYPHEN_RE = re.compile(r'(?<=\w)-(?=\w)')
MULTISPACE_RE = re.compile(r'\s+')
# ordinal suffix (2nd, 3rd, ...) or a plain letter-run -- matched together in
# one pass so "first word of the string" tracking is shared correctly across
# the whole string, not reset per whitespace-token
WORD_RUN_RE = re.compile(r'\d+(?:st|nd|rd|th)\b|[A-Za-z]+', re.IGNORECASE)


def _case_word(word, is_first_word):
    if re.fullmatch(r'\d+(st|nd|rd|th)', word, re.IGNORECASE):
        return word.lower()
    if not word.isalpha():
        return word
    upper = word.upper()

    if upper in SLASH_TOKENS or upper in ABBREVIATIONS_UPPER:
        return upper
    if re.fullmatch(r'kv', word, re.IGNORECASE):
        return 'kV'
    # plural of a recognized abbreviation, e.g. RLDCs, SLDCs
    if len(word) > 1 and upper.endswith('S') and upper[:-1] in ABBREVIATIONS_UPPER:
        return upper[:-1] + 's'

    lower = word.lower()
    if lower in ENGLISH_WORDS:
        return word[:1].upper() + word[1:].lower()

    # inferred abbreviation: short, was already all-uppercase in the
    # source, and isn't a recognized ordinary word
    if word == upper and len(word) <= 6:
        return upper

    return word[:1].upper() + word[1:].lower()


def _title_case_with_abbreviations(s):
    """Applies _case_word to every letter-run in s, tracking which letter-run
    is the very first one in the whole string (for stopword capitalization)
    and whether any inferred-abbreviation/explicit-abbreviation logic fired,
    for reason-tagging."""
    reasons = set()
    first_seen = [False]

    def repl(m):
        word = m.group(0)
        is_first = not first_seen[0]
        first_seen[0] = True
        new_word = _case_word(word, is_first)
        if new_word != word:
            upper = word.upper()
            if new_word == upper and word != upper:
                reasons.add("Abbreviation preserved")
            else:
                reasons.add("Capitalization")
        return new_word

    result = WORD_RUN_RE.sub(repl, s)
    return result, reasons


def _normalize_common(text, split_hyphens):
    if text is None:
        return text, set()
    original = str(text)
    reasons = set()
    s = original

    if KV_RE.search(s):
        s2 = KV_RE.sub(lambda m: f"{m.group(1)} kV", s)
        if s2 != s:
            reasons.add("Abbreviation preserved")
        s = s2

    if split_hyphens:
        s_before = s
        s = HYPHEN_RE.sub(' ', s)
        if s != s_before:
            reasons.add("Separator normalization")

    stripped = s.strip()
    collapsed = MULTISPACE_RE.sub(' ', stripped)
    if collapsed != s:
        reasons.add("Whitespace")
    s = collapsed

    if not s:
        return s, reasons

    s, word_reasons = _title_case_with_abbreviations(s)
    reasons |= word_reasons

    return s, reasons


def normalize_designation(text):
    """For designation / control-room / switchyard label fields: hyphens
    used as word separators are converted to spaces, per the Hyphen Rules
    examples (all of which are designation-shaped: "SLDC-Control room",
    "Executive-Engineer")."""
    return _normalize_common(text, split_hyphens=True)


def normalize_name(text):
    """For organization/sub-organization/hospital/emergency-service NAME
    fields: hyphens are left untouched -- the Organization Names rule says
    to only normalize capitalization and explicitly not remove meaningful
    punctuation from names, unlike the Hyphen Rules section (whose examples
    are all designations, not names)."""
    return _normalize_common(text, split_hyphens=False)


def normalize_email(text):
    if text is None:
        return text, set()
    original = str(text)
    s = MULTISPACE_RE.sub(' ', original.strip())
    lowered = s.lower()
    reasons = set()
    if s != original:
        reasons.add("Whitespace")
    if lowered != s:
        reasons.add("Email lowercase")
    return lowered, reasons


# ─── COLUMN SCOPE ─────────────────────────────────────────────────────────────
# sheet -> (pk_column_index, [(column_index, column_name, kind), ...])
# kind: "label" (title-case + abbreviations) or "email" (lowercase only)

SCOPE = {
    "organizations": (1, [(2, "name", "name")]),
    "sub_organizations": (1, [(3, "name", "name")]),
    "employees": (1, [(6, "designation", "designation")]),
    "KMP": (1, [(4, "designation", "designation"), (8, "email", "email")]),
    "Utility_Heads": (1, [(4, "designation", "designation"), (8, "email", "email")]),
    "control_rooms": (1, [
        (4, "organization_name", "name"), (5, "sub_organization_name", "name"),
        (6, "cr_label", "designation"), (7, "designation", "designation"), (12, "email_1", "email"),
    ]),
    "switchyards": (1, [
        (4, "organization_name", "name"), (5, "sub_organization_name", "name"),
        (6, "cr_label", "designation"), (9, "email_1", "email"),
    ]),
    "Hospitals": (1, [(3, "hospital_name", "name")]),
    "Emergency_Services": (1, [(2, "service_type", "name"), (3, "name", "name")]),
    "email_addresses": (1, [(4, "email", "email")]),
}


def run():
    input_mtime_before = os.path.getmtime(INPUT_XLSX)

    print(f"Loading {INPUT_XLSX} ...")
    wb = openpyxl.load_workbook(INPUT_XLSX)

    log_rows = []
    cells_reviewed = 0
    cells_modified = 0
    by_sheet = {}
    by_column = {}

    for sheet_name, (pk_col, columns) in SCOPE.items():
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2):
            pk_value = row[pk_col - 1].value
            if pk_value is None:
                continue
            for col_idx, col_name, kind in columns:
                cell = row[col_idx - 1]
                old_value = cell.value
                if old_value is None or str(old_value).strip() == "":
                    continue
                cells_reviewed += 1

                if kind == "email":
                    new_value, reasons = normalize_email(old_value)
                elif kind == "designation":
                    new_value, reasons = normalize_designation(old_value)
                else:
                    new_value, reasons = normalize_name(old_value)

                if new_value != old_value:
                    cell.value = new_value
                    cells_modified += 1
                    by_sheet[sheet_name] = by_sheet.get(sheet_name, 0) + 1
                    key = f"{sheet_name}.{col_name}"
                    by_column[key] = by_column.get(key, 0) + 1
                    log_rows.append([
                        sheet_name, pk_value, col_name, old_value, new_value,
                        "; ".join(sorted(reasons)) if reasons else "Formatting normalized",
                    ])

    print(f"Reviewed {cells_reviewed} cells, modified {cells_modified}.")
    print(f"Saving {OUTPUT_XLSX} ...")
    wb.save(OUTPUT_XLSX)

    assert os.path.getmtime(INPUT_XLSX) == input_mtime_before, "Input workbook was modified!"

    # ─── structural validation ───
    wb_orig = openpyxl.load_workbook(INPUT_XLSX, read_only=True)
    wb_new = openpyxl.load_workbook(OUTPUT_XLSX, read_only=True)
    same_sheets = wb_orig.sheetnames == wb_new.sheetnames
    row_diffs = {}
    for s in wb_orig.sheetnames:
        r0, r1 = wb_orig[s].max_row, wb_new[s].max_row
        if r0 != r1:
            row_diffs[s] = (r0, r1)
    same_rows = not row_diffs

    pk_diffs = 0
    for sheet_name, (pk_col, _cols) in SCOPE.items():
        orig_ids = [r[pk_col - 1] for r in wb_orig[sheet_name].iter_rows(min_row=2, values_only=True)]
        new_ids = [r[pk_col - 1] for r in wb_new[sheet_name].iter_rows(min_row=2, values_only=True)]
        if orig_ids != new_ids:
            pk_diffs += 1

    print("Writing log...")
    wb_log = openpyxl.Workbook()
    wb_log.remove(wb_log.active)

    ws_sum = wb_log.create_sheet("Summary", 0)
    ws_sum["A1"] = "Formatting Normalization Summary"
    ws_sum["A1"].font = Font(bold=True, size=14)
    r = 3
    for label, value in [
        ("Total Cells Reviewed", cells_reviewed),
        ("Total Cells Modified", cells_modified),
    ]:
        ws_sum.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws_sum.cell(row=r, column=2, value=value)
        r += 1
    r += 1
    ws_sum.cell(row=r, column=1, value="Changes by Worksheet").font = Font(bold=True)
    r += 1
    for s, n in sorted(by_sheet.items()):
        ws_sum.cell(row=r, column=1, value=s)
        ws_sum.cell(row=r, column=2, value=n)
        r += 1
    r += 1
    ws_sum.cell(row=r, column=1, value="Changes by Column").font = Font(bold=True)
    r += 1
    for c, n in sorted(by_column.items()):
        ws_sum.cell(row=r, column=1, value=c)
        ws_sum.cell(row=r, column=2, value=n)
        r += 1
    r += 1
    ws_sum.cell(row=r, column=1, value="Validation").font = Font(bold=True, size=12)
    r += 1
    for label, ok in [
        ("Same number of worksheets", same_sheets),
        ("Same number of rows (per sheet)", same_rows),
        ("No PK/ID values changed in normalized sheets", pk_diffs == 0),
    ]:
        ws_sum.cell(row=r, column=1, value=label)
        ws_sum.cell(row=r, column=2, value="PASS" if ok else "FAIL")
        r += 1
    if row_diffs:
        r += 1
        ws_sum.cell(row=r, column=1, value="Row count differences:")
        for s, (r0, r1) in row_diffs.items():
            r += 1
            ws_sum.cell(row=r, column=1, value=f"  {s}: {r0} -> {r1}")
    r += 2
    ws_sum.cell(row=r, column=1, value=(
        "Only the columns explicitly covered by the standardization rules were touched: "
        "organization/sub-organization names, employee/KMP/Utility Head designations, control room "
        "and switchyard labels, Hospital and Emergency Service names, and all email columns "
        "(lowercased). Person name fields, phone numbers, IDs/FKs, department, Reference_Sections, "
        "and db_schema were left untouched by design (see script docstring)."
    ))
    ws_sum.column_dimensions["A"].width = 46
    ws_sum.column_dimensions["B"].width = 60

    ws_log = wb_log.create_sheet("Change_Log")
    ws_log.append(LOG_HEADER)
    for cell in ws_log[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F6F5C")
    for row in log_rows:
        ws_log.append(row)
    widths = [20, 12, 22, 40, 40, 40]
    for col, w in zip("ABCDEF", widths):
        ws_log.column_dimensions[col].width = w
    ws_log.freeze_panes = "A2"

    wb_log.save(LOG_XLSX)

    print("=" * 70)
    print("  Formatting Normalization Summary")
    print("=" * 70)
    print(f"  Total Cells Reviewed : {cells_reviewed}")
    print(f"  Total Cells Modified : {cells_modified}")
    print("  Changes by Worksheet:")
    for s, n in sorted(by_sheet.items()):
        print(f"    {s}: {n}")
    print("  Changes by Column:")
    for c, n in sorted(by_column.items()):
        print(f"    {c}: {n}")
    print(f"  Same sheet count/names : {same_sheets}")
    print(f"  Same row counts        : {same_rows}")
    print(f"  PK/ID integrity        : {'PASS' if pk_diffs == 0 else 'FAIL'}")
    print("=" * 70)
    print(f"  Corrected workbook: {OUTPUT_XLSX}")
    print(f"  Log:               {LOG_XLSX}")


if __name__ == "__main__":
    run()
