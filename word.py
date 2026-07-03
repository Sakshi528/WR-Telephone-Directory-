"""
Western Region Phone Directory - DOCX Parser & MySQL Inserter
=============================================================
ROOT CAUSES FIXED:
  1. Merged cells repeated the same text → now de-duplicated by cell object identity
  2. Column headers like "Name (Shri/Smt)", "Name (S/Sh)" weren't matching → fuzzy prefix mapping
  3. Non-breaking spaces (\xa0), extra whitespace → cleaned on every cell
  4. Section/org headers inside tables mis-parsed as data rows → detected and tracked separately
  5. Motivational quotes embedded in cells treated as names → filtered out
  6. Phone numbers containing "Fax 022-xxx" → split into phone_office / fax_office
  7. MySQL column charset not utf8mb4 → special chars corrupted storage → enforced utf8mb4

USAGE:
  pip install python-docx mysql-connector-python
  python parse_directory.py

Update DB_CONFIG and DOCX_PATH before running.
"""

import re
import mysql.connector
from docx import Document

# ─── CONFIG ──────────────────────────────────────────────────────────────────
DB_CONFIG = {
    "host":       "localhost",
    "user":       "root",
    "password":   "your_password",   # ← change
    "database":   "telephone_dir",   # ← change
    "charset":    "utf8mb4",
    "use_unicode": True,
}

DOCX_PATH = "Western_Region_Phone_Directory_2025_Main.docx"

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def clean(text: str) -> str:
    """Normalise whitespace and remove non-breaking spaces."""
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def unique_cells(row):
    """
    Return only physically distinct cells.
    python-docx returns the SAME cell object multiple times for merged spans;
    deduplicating by id() is the only reliable fix.
    """
    seen, result = set(), []
    for cell in row.cells:
        if id(cell) not in seen:
            seen.add(id(cell))
            result.append(clean(cell.text))
    return result


def map_header(header: str) -> str | None:
    """Map raw header text → schema key (handles all variants in this document)."""
    h = header.lower().strip()
    if h.startswith("name"):                                      return "name"
    if h.startswith("designation"):                               return "designation"
    if "telephone (o)" in h or h in ("office", "telephone no",
       "phone no.", "land line 1"):                               return "phone_office"
    if "telephone (r" in h or h == "land line 2":                return "phone_residence"
    if h in ("telephone (f)", "fax"):                            return "fax_office"
    if h.startswith("mobile") or h in ("contact number",
                                        "emergency number"):      return "mobile"
    if h in ("email", "e-mail", "e-mail id", "mail id"):         return "email"
    return None


def is_header_row(cells: list[str]) -> bool:
    TRIGGERS = {"name", "designation", "telephone", "mobile", "email",
                "sl.no.", "s no.", "mail id", "department",
                "sr. no.", "sl no.", "phone no.", "contact number"}
    lowered = {c.lower().strip() for c in cells}
    return bool(lowered & TRIGGERS) or any(
        c.lower().startswith("name") and len(c) > 2 for c in cells
    )


def is_section_header(cells: list[str]) -> bool:
    """
    A section header is a single merged cell spanning the whole table width,
    containing an organisation / section name.
    """
    non_empty = [c for c in cells if c]
    if len(non_empty) != 1 or len(non_empty[0]) <= 3:
        return False
    text = non_empty[0].upper()
    ORG_KEYWORDS = ["CENTRE", "CORPORATION", "COMMITTEE", "STATION", "LIMITED",
                    "DEPARTMENT", "POWER", "GRID", "NTPC", "STATE", "NUCLEAR",
                    "ENERGY", "LOAD DESPATCH", "TRANSMISSION", "BUYER", "VENDOR",
                    "ELECTRICITY", "ATOMIC", "REGION", "SOLAR", "WIND"]
    return any(k in text for k in ORG_KEYWORDS)


def is_motivational_quote(text: str) -> bool:
    """Some cells contain inspirational quotes rather than names."""
    LOW_SIGNALS = ["wins", "together", "teamwork", "success", "intelligence",
                   "alone", "dream", "smart", "championship"]
    t = text.lower()
    return any(w in t for w in LOW_SIGNALS) or text.strip().endswith("…")


def split_phone_fax(text: str) -> tuple[str, str]:
    """'022-12345  Fax 022-67890'  →  ('022-12345', '022-67890')"""
    m = re.search(r"[Ff][Aa][Xx][\s:]*([0-9\-\s/+().]+)", text)
    if m:
        return text[:m.start()].strip().rstrip(",/").strip(), m.group(1).strip()
    return text, ""


def normalize_row(cells: list[str], col_map: list[str]) -> dict:
    record = {k: "" for k in ("name", "designation", "phone_office",
                               "phone_residence", "mobile", "email",
                               "fax_office", "organization")}
    for idx, value in enumerate(cells):
        if idx >= len(col_map):
            break
        key = map_header(col_map[idx])
        if key and value:
            record[key] = (record[key] + " / " + value) if record[key] else value

    if record["phone_office"]:
        phone, fax = split_phone_fax(record["phone_office"])
        record["phone_office"] = phone
        if fax and not record["fax_office"]:
            record["fax_office"] = fax

    # Filter out motivational quotes that slip into the name column
    if record["name"] and is_motivational_quote(record["name"]):
        record["name"] = ""

    return record


# ─── PARSE ───────────────────────────────────────────────────────────────────

def parse_docx(path: str) -> list[dict]:
    doc = Document(path)
    records = []
    current_org = "Unknown"

    for table in doc.tables:
        col_map: list[str] = []

        for row in table.rows:
            cells = unique_cells(row)

            if not any(cells):
                continue

            if is_section_header(cells):
                text = cells[0]
                # Strip any quote fragment merged with the org name
                parts = re.split(r'["""\u201c\u201d]', text)
                candidate = max(parts, key=len).strip()
                if candidate:
                    current_org = candidate
                continue

            if is_header_row(cells):
                col_map = cells
                continue

            # Pure separator rows
            if all(re.fullmatch(r"[-=*]+", c) or c == "" for c in cells):
                continue

            if not col_map:
                continue

            record = normalize_row(cells, col_map)
            record["organization"] = current_org

            if not any(record[k] for k in ("name", "phone_office", "mobile", "email")):
                continue

            records.append(record)

    return records


# ─── MYSQL ───────────────────────────────────────────────────────────────────

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS telephone_directory (
    id               INT AUTO_INCREMENT PRIMARY KEY,
    organization     VARCHAR(300),
    name             VARCHAR(255),
    designation      VARCHAR(255),
    phone_office     VARCHAR(150),
    phone_residence  VARCHAR(150),
    fax_office       VARCHAR(150),
    mobile           VARCHAR(150),
    email            VARCHAR(300),
    created_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FULLTEXT INDEX ft_search (name, designation, organization, mobile, email)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
"""

INSERT_SQL = """
INSERT INTO telephone_directory
    (organization, name, designation, phone_office, phone_residence,
     fax_office, mobile, email)
VALUES
    (%(organization)s, %(name)s, %(designation)s, %(phone_office)s,
     %(phone_residence)s, %(fax_office)s, %(mobile)s, %(email)s)
"""


def insert_to_mysql(records: list[dict]) -> None:
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()
    cursor.execute("SET NAMES utf8mb4")
    cursor.execute("SET CHARACTER SET utf8mb4")
    cursor.execute(CREATE_TABLE_SQL)
    conn.commit()
    cursor.execute("TRUNCATE TABLE telephone_directory")

    for i in range(0, len(records), 100):
        cursor.executemany(INSERT_SQL, records[i:i + 100])
        conn.commit()

    cursor.close()
    conn.close()
    print(f"✅  Inserted {len(records)} records into telephone_directory.")


# ─── SEARCH (use this to verify after import) ────────────────────────────────

def search(keyword: str) -> list[dict]:
    """Search by name, designation, organisation, mobile or email."""
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SET NAMES utf8mb4")
    like = f"%{keyword}%"
    cursor.execute("""
        SELECT organization, name, designation, phone_office, mobile, email
        FROM telephone_directory
        WHERE name LIKE %s OR designation LIKE %s
           OR organization LIKE %s OR mobile LIKE %s OR email LIKE %s
        LIMIT 20
    """, (like, like, like, like, like))
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        print(r)
    return rows


# ─── MAIN ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("📄  Parsing DOCX...")
    records = parse_docx(DOCX_PATH)
    print(f"    → {len(records)} records extracted")

    print("\nSample records:")
    shown = 0
    for r in records:
        if r["name"]:
            print(f"  [{r['organization'][:38]:<38}] {r['name']:<30} | {r['mobile']}")
            shown += 1
            if shown == 8:
                break

    print(f"\n  With name  : {sum(1 for r in records if r['name'])}")
    print(f"  With mobile: {sum(1 for r in records if r['mobile'])}")
    print(f"  With email : {sum(1 for r in records if r['email'])}")

    print("\n💾  Inserting into MySQL...")
    insert_to_mysql(records)

    print("\n🔍  Test search for 'Executive Director':")
    search("Executive Director")