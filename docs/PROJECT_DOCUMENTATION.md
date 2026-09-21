# WR Telephone Directory Management System
## Complete Project Documentation & GM Presentation Study Guide

**Prepared for:** UAT / LAN Deployment readiness review and senior management presentation
**System:** Western Region (WRLDC) Telephone Directory Management System
**Stack:** Python 3 / Flask 3.1.1 / Flask-SQLAlchemy 3.1.1 / Flask-Login 0.6.3 / PostgreSQL / psycopg[binary] (v3) / python-docx 1.1.2 / openpyxl 3.1.5 / reportlab 5.0.0 / Bootstrap / vanilla JavaScript

> **Note on accuracy:** Every claim in this document was verified directly against the current source code (`models/`, `routes/`, `services/`, `utils/`, `templates/`, `static/js/`, `migrations/`, and the import/reconciliation scripts) rather than written from generic Flask knowledge. Where a feature does **not** exist in the codebase (e.g. CSRF protection, an audit-log viewing screen, rate limiting), this document says so explicitly instead of assuming it. Exact file names, route paths, column names, and function names are quoted so you can pull up the source during Q&A if needed.

---

## Table of Contents

1. Project Overview
2. Complete System Architecture
3. Database Design
4. Backend Implementation
5. Frontend Implementation
6. User Module
7. Admin Module
8. Technical Implementation
9. Telephone Directory Module
10. Employee Directory
11. Email Distribution Lists
12. Utility Heads
13. Administrative Heads
14. Organization Classification
15. Import Pipeline
16. Export Modules
17. Security
18. Validation
19. Challenges Faced
20. Final Production Readiness
21. Interview / GM Questions (100+)
22. Code Walkthrough
23. Future Enhancements

---

# 1. Project Overview

## 1.1 Project Objective

Build a single, authoritative, searchable, web-based telephone directory for the Western Region power sector — covering WRLDC's own staff, every Discom/generator/transmission utility/SLDC/RLDC in the region, their senior leadership (Administrative Heads/KMPs), control rooms, and switchyards — that replaces a set of manually maintained Word and Excel documents with a live PostgreSQL-backed application accessible over the WRLDC LAN.

## 1.2 Business Problem

WRLDC (Western Region Load Despatch Centre) is the nodal grid-operations body for the Western Region and needs a always-current, instantly searchable way to reach the right person — a Utility Head, a Control Room, a State SLDC officer, an Administrative Head/KMP of a Discom or generator — during both routine coordination and grid emergencies. Before this application, that lookup depended on a **Word document** ("Western Region Phone Directory") that was periodically emailed around and an accompanying **Excel workbook** used to prepare/verify it.

## 1.3 The Existing Manual Process

1. A designated person periodically collected updated contact details from each organization (by phone, email, or site visit) and manually edited the master Word document.
2. Corrections were tracked informally — there was no single source of truth for "what changed since the last version" and no approval workflow.
3. When an employee left, was transferred, retired, or was promoted, the Word document was updated only when someone remembered to do it — there was no status/history mechanism.
4. Searching meant `Ctrl+F` inside a very long Word document, or scrolling a wide Excel sheet — no filtering by organization type, state, category, or role.
5. Distributing an updated directory meant re-emailing a large file to everyone, with no guarantee that everyone was using the same latest version.

## 1.4 Problems With the Excel / Word Directories

| Problem | Why it mattered |
|---|---|
| No single source of truth | Word doc, Excel workbook, and people's own contact lists all drifted independently. |
| No search/filter | Finding "the SLDC officer for Chhattisgarh" meant manual scanning. |
| No update workflow | Anyone editing the Word doc could introduce silent errors; no review/approval step. |
| No status tracking | Transferred/retired staff stayed listed as current indefinitely. |
| No audit trail | No record of who changed what, when, or why. |
| Formatting fragility | Word numbering, inconsistent Control Room naming, and copy-paste artifacts (OCR-style broken spacing, duplicated sections) accumulated over years of editing. |
| Duplicate/garbled entries | The reconciliation work uncovered 74 pairs of duplicate organizations, mis-attributed employees, and malformed control-room entries once the data was moved into a structured database (see Section 19). |
| No role/seniority awareness | There was no reliable, queryable notion of "who is the current head of this organization" — it depended on someone knowing designation titles by heart. |
| Not accessible concurrently | A Word file being edited by one person blocks concurrent editing by another. |

## 1.5 Why This Application Was Developed

To replace the document-based process with a **database-backed, role-aware, auditable, searchable** system that:
- Keeps one authoritative PostgreSQL database as the single source of truth.
- Lets the public (LAN users) search/browse without needing edit access.
- Lets a controlled Admin role make corrections, with every material change captured through an approval workflow or an audit log.
- Automatically determines "who is the current Utility Head" of an organization from designation text, rather than relying on a manually maintained flag.
- Still produces the same familiar Word/Excel/PDF outputs the organization is used to distributing — but generated on demand, directly from live data, so they can never drift from what the application shows.
- Tracks employee status (Active/Transferred/Retired/Resigned/etc.) so stale contacts are visibly flagged rather than silently wrong.

## 1.6 Expected Business Benefits

- **Always current** — one edit is visible to every user immediately; no more re-distributing files.
- **Faster lookups** — category cards, search-as-you-type, and filters replace manual scanning.
- **Self-service correction** — any user (including the contact themselves) can submit a "Request Update," reducing the burden on a single directory maintainer and closing the loop faster.
- **Accountability** — every admin-side correction to Administrative Heads, Employee status, Organization category, and Email Groups is written to an append-only `audit_logs` table with who/when/old value/new value.
- **Regional continuity** — Directory Versions (monthly PDF/Excel snapshots) preserve an official, dated record for compliance/reference even as the live data keeps changing.
- **Consistent outputs** — Excel, Word, and PDF exports are generated directly from the live database, so what is printed/distributed can never disagree with what the application shows online.

## 1.7 Scope of the Project

**In scope (implemented):**
- Employee Directory (WRLDC internal staff) with search, filters, status, emergency contacts.
- Telephone Directory (region-wide: all organizations, KMPs, Control Rooms, Switchyards) with category/organization drill-down.
- Utility Heads module (auto-detected senior-most contact per organization).
- Administrative Heads module (manually curated Chairman/MD/CEO/IAS-cadre leadership + PA/PS assistants), with role history archival.
- Email Distribution Lists / Address Book (organization-category and role-based contact lists, with static and dynamic-filter groups).
- Directory Numbers (Control Rooms, Switchyards, and miscellaneous numbers).
- Public self-service "Request Update" workflow for Employees, Directory Numbers, and Administrative Heads, reviewed/approved by an Admin.
- Directory Versions (generate/regenerate/delete official monthly PDF+Excel snapshots).
- Full Admin back office: Employee/Organization/Directory Number/Emergency Contact/Administrative Head CRUD, Update Request review queue, custom Email Group builder, Import History viewer, Export Verification Workbook.
- A multi-stage Word → Excel → validated/corrected Excel → PostgreSQL import & reconciliation pipeline (see Section 15), used to seed and periodically refresh the database from the official Word directory.

**Explicitly out of scope / not implemented (see Section 17 & 23):**
- CSRF protection (none found anywhere in the codebase).
- Rate limiting on public forms.
- A UI screen to browse the `audit_logs` table (the table and write path exist; there is no admin page to read it).
- User/account management UI (admin accounts are provisioned via a script, not through the web app).
- Automated import rollback (`ImportBatch.rollback_available` is always `False`).
- Mobile app / native client — this is a responsive web application only.

## 1.8 Technologies Used

| Layer | Technology |
|---|---|
| Web framework | Flask 3.1.1 |
| ORM | Flask-SQLAlchemy 3.1.1 |
| Authentication/session | Flask-Login 0.6.3 |
| Database | PostgreSQL, via psycopg[binary] (v3) |
| Word document parsing/export | python-docx 1.1.2 |
| Excel parsing/export | openpyxl 3.1.5 |
| PDF generation | reportlab 5.0.0 |
| Frontend templating | Jinja2 (bundled with Flask) |
| Frontend UI framework | Bootstrap (cards, accordions, modals, badges) |
| Frontend interactivity | Vanilla JavaScript (no frontend framework/bundler — `static/js/*.js`, plain `fetch()`) |
| Dev server | Werkzeug (Flask's built-in dev server) |

---

# 2. Complete System Architecture

## 2.1 Architecture Style

This is a classic **server-rendered monolith**: Flask renders Jinja2 HTML templates on the server for almost every page; a small number of endpoints return JSON for AJAX-driven widgets (autocomplete, cascading dropdowns, the Address Book grid). There is no separate frontend build step, no SPA framework, and no REST API layer beyond the handful of `jsonify` endpoints listed in Section 4.

## 2.2 High-Level Diagram

```
                                   ┌─────────────────────────────┐
                                   │        Browser (LAN)        │
                                   │  Bootstrap UI + vanilla JS   │
                                   │  (autocomplete.js,           │
                                   │   org_type_cascade.js,       │
                                   │   email_distribution*.js)    │
                                   └───────────────┬──────────────┘
                                                    │ HTTP (GET/POST, fetch/JSON)
                                                    ▼
                       ┌───────────────────────────────────────────────────┐
                       │                    app.py                         │
                       │  Flask app, LoginManager, 3 Blueprints registered │
                       └───────┬───────────────────┬───────────────────────┘
                                │                   │
              ┌─────────────────┘                   └──────────────────┐
              ▼                                                        ▼
   ┌─────────────────────┐                                 ┌───────────────────────┐
   │  routes/auth_routes  │                                 │   routes/user_routes   │
   │  (login/logout)      │                                 │   (PUBLIC, no login    │
   └──────────┬───────────┘                                 │   required anywhere)   │
              │                                              └───────────┬────────────┘
              │                                                          │
              │                                              ┌───────────▼────────────┐
              │                                              │   routes/admin_routes   │
              │                                              │  (@admin_required on    │
              │                                              │   every route)          │
              │                                              └───────────┬────────────┘
              │                                                          │
              └──────────────────────────┬───────────────────────────────┘
                                          ▼
                        ┌───────────────────────────────────┐
                        │             services/               │
                        │  audit_service, head_service,       │
                        │  email_distribution_service,        │
                        │  directory_version_service,         │
                        │  excel_generator, pdf_generator,     │
                        │  verification_workbook_service,     │
                        │  org_stats                           │
                        └───────────────────┬─────────────────┘
                                             ▼
                        ┌───────────────────────────────────┐
                        │              models/                │
                        │  SQLAlchemy models — one class per  │
                        │  table, relationships, computed     │
                        │  "resolved_*" properties            │
                        └───────────────────┬─────────────────┘
                                             ▼
                        ┌───────────────────────────────────┐
                        │            PostgreSQL                │
                        │        telephone_directory DB        │
                        └───────────────────────────────────┘

   Offline / one-time pipeline (not part of the running web app):
   Word (.docx) ──▶ scripts/*.py (repair/validate/reconcile/correct/normalize)
                 ──▶ WR_DB_Ready_Final_Verified_v3.xlsx
                 ──▶ import_from_excel_db.py / import_employees_excel.py --apply
                 ──▶ PostgreSQL (feeds the diagram above)
```

## 2.3 How the Pieces Connect

- **`app.py`** is the single composition root: it creates the Flask app, loads `config.Config`, initializes `db` (the shared `SQLAlchemy()` instance from `models/__init__.py`) and Flask-Login's `LoginManager`, registers the three blueprints (`auth_bp`, `user_bp`, `admin_bp` — none with a `url_prefix`, so each route's full path is written out in its own `@blueprint.route(...)` decorator), and defines a `@login_manager.user_loader` plus a `@login_manager.unauthorized_handler`/401 handler that both redirect to `/admin-login`.
- **Routing** is split three ways by *audience*, not by resource: `auth_routes.py` (login/logout only), `user_routes.py` (every public/anonymous page — note that **no route in this file requires login**), and `admin_routes.py` (every route decorated with a local `admin_required` decorator that layers "is this an authenticated admin" on top of Flask-Login's `@login_required`).
- **Templates** live in one flat `templates/` folder (no per-blueprint subfolders), all extending a shared `base.html` that carries the navbar, flash-message rendering, the 20-minute idle-session-timeout script, and loads the two globally-needed JS files (`autocomplete.js`, `org_type_cascade.js`).
- **Services** hold logic that is either (a) reused by more than one route, or (b) non-trivial enough that keeping it out of `routes/` matters for readability — audit logging, Administrative Head lifecycle (assign/update/end-role/assistants), Email Distribution resolution (static/dynamic groups, category browsing), Directory Version generation (PDF+Excel snapshotting), and the Verification Workbook export. Two generator modules (`excel_generator.py`, `pdf_generator.py`) are pure functions: snapshot-dict-in, bytes-out, with no DB or disk access of their own.
- **Models** are one SQLAlchemy class per table in `models/`, all sharing the single `db = SQLAlchemy()` instance defined in `models/__init__.py`. Several models expose **computed `resolved_*` properties** (see Section 3) that transparently prefer a linked `Employee`'s live data over locally stored fallback fields — this pattern is central to how Administrative Heads and their PA/PS Assistants work.
- **`utils/`** holds cross-cutting helpers with no Flask/DB dependency of their own: `designation_rank.py` (the Utility Head auto-detection algorithm), `logger.py` (shared file+console logging setup used by services and import scripts), and `user_agent.py` (a small regex-based User-Agent parser used only for audit-trail metadata, because Werkzeug 3.x removed its built-in UA parser).
- **Static files** (`static/js/*.js`) are plain, dependency-free vanilla JavaScript IIFEs — no npm, no bundler, no frontend framework. Each auto-binds itself to `data-*`-marked DOM elements on `DOMContentLoaded`, so a template opts into a script's behavior purely by including the right `data-*` attributes, not by writing page-specific JS.
- **Import scripts** (root-level `import_*.py`, `init_db.py`) and the **`scripts/` reconciliation pipeline** (`repair_workbook.py` → `validate_workbook.py` → `compare_word_excel.py`/`reconcile_word_database.py` → `generate_validation_report.py` → `apply_corrections.py` → `normalize_formatting.py` → `classify_organizations.py`/`reclassify_organization_types.py` → `generate_import_readiness_report.py`) are **offline, one-time/periodic tools**, run from the command line, never imported by the running Flask app. They read the Word directory and/or Excel workbooks and write into the same PostgreSQL database the web app reads from.
- **Export modules** (Excel/Word/PDF generation inside `routes/user_routes.py`, plus `services/excel_generator.py`, `services/pdf_generator.py`, `services/verification_workbook_service.py`) always read live PostgreSQL data at request time — none of them read from the original import workbook, so an export can never show stale data relative to the live application.

---

# 3. Database Design

## 3.1 Entity Overview

The schema is PostgreSQL, defined through SQLAlchemy models (`models/*.py`) plus a series of hand-written, numbered SQL migration files (`migrations/002_*.sql` through `migrations/013_*.sql`, applied via `psql`, **not** an ORM migration tool like Alembic). The base schema (`users`, `employees`, `organizations`, `departments`, `directory_numbers`, `update_requests`, `emergency_contacts`) predates the numbered migration system entirely — there is no committed `001` migration or schema.sql capturing it.

19 tables in total:

| # | Table | Model class | Purpose (one line) |
|---|---|---|---|
| 1 | `users` | `User` | Login accounts (admin/user), password hashes. |
| 2 | `employees` | `Employee` | Core personnel record — every named contact in the system. |
| 3 | `departments` | `Department` | Simple lookup of department names. |
| 4 | `organizations` | `Organization` | Every power-sector entity (Discom, generator, SLDC, etc.), with a parent/child hierarchy. |
| 5 | `organization_categories` | `OrganizationCategory` | Controlled "Organization Type" taxonomy. |
| 6 | `service_types` | `ServiceType` | Civil-service cadre lookup (IAS/IPS/IFS/etc.) for Administrative Heads. |
| 7 | `directory_numbers` | `DirectoryNumber` | Non-employee phone entries — Control Rooms, Switchyards, misc numbers. |
| 8 | `emergency_contacts` | `EmergencyContact` | Next-of-kin contact per (non-KMP) employee. |
| 9 | `administrative_heads` | `AdministrativeHead` | Current holder of a senior post (Administrative Head or KMP) at an organization. |
| 10 | `administrative_head_assistants` | `AdministrativeHeadAssistant` | PA/PS attached to an Administrative Head. |
| 11 | `administrative_head_history` | `AdministrativeHeadHistory` | Immutable archive of completed Administrative Head/KMP tenures. |
| 12 | `employee_status_history` | `EmployeeStatusHistory` | Change log for employee status/organization/designation (drives the Archive timeline). |
| 13 | `update_requests` | `UpdateRequest` | Public self-service correction requests (Employee/Directory/Administrative Head), admin-reviewed. |
| 14 | `audit_logs` | `AuditLog` | Append-only change log for the enterprise modules. |
| 15 | `import_batches` | `ImportBatch` | One row per import script run (Import History screen). |
| 16 | `directory_versions` | `DirectoryVersion` | One row per generated/regenerated official PDF+Excel snapshot. |
| 17 | `email_groups` / `group_members` | `EmailGroup` / `GroupMember` | Saved distribution groups — static membership or dynamic filter-based. |
| 18 | `email_group_filters` | `EmailGroupFilter` | One filter criterion (OR within type, AND across types) for a DYNAMIC email group. |
| 19 | `organization_subcategories` | `OrganizationSubcategory` | Reusable Subcategory suggestion list per Organization Category (autocomplete only). |

## 3.2 Table-by-Table Detail

### `users` (`User`)
| Column | Type | Notes |
|---|---|---|
| id | Integer PK | |
| username | String(100) | not null |
| email | String(255) | unique, not null |
| password | String(255) | Werkzeug-hashed, not plaintext |
| role | String(20) | default `"user"`; `"admin"` grants access — **plain string comparison, no DB enum/CHECK** |
| status | String(20) | default `"active"` — no CHECK constraint found |
| created_at | DateTime | server default now() |

`User.is_admin()` = `self.role == "admin"`. Relationship: `update_requests` (as reviewer/submitter, via `UpdateRequest.user_id`/`reviewed_by`). **No admin UI exists to create/edit/list users** — accounts are provisioned by the standalone `create_admin.py` script.

### `employees` (`Employee`)
| Column | Type | Notes |
|---|---|---|
| id | Integer PK | |
| employee_name | String(255) | not null |
| designation | String(255) | free text — drives Utility Head auto-detection |
| organization_id | Integer | FK → organizations.id |
| department_id | Integer | FK → departments.id |
| location, region | String(255) | |
| office_phone, residence_phone, mobile_phone | Text | |
| email | String(255) | |
| is_utility_head | Boolean | not null, default False — **not used by the Utility Head algorithm** (see §12) |
| is_kmp | Boolean | not null, default False — distinguishes region-wide KMP contacts from WRLDC-internal staff |
| status | String(20) | not null, default `"ACTIVE"`; SQL `CHECK chk_employee_status` |
| status_changed_at | DateTime | |

`EMPLOYEE_STATUSES` (Python constant in `models/employee.py`, mirrored by the SQL CHECK): `ACTIVE, INACTIVE, TRANSFERRED, RETIRED, DEPUTATION, RESIGNED, CONTRACT_ENDED, DECEASED` (the last two were added in migration 007; `email_group_filters.status_value`'s own CHECK was **never widened to match** — it still only allows the original 6).

### `departments` (`Department`)
`id`, `department_name` (String(100), unique, not null). No FKs of its own.

### `organizations` (`Organization`)
| Column | Type | Notes |
|---|---|---|
| id | Integer PK | |
| organization_name | String(500) | unique |
| region | String(255) | historically a free-text "parent import grouping label"; later reused to hold state names by `set_org_regions.py` |
| state | String(100) | added migration 010; distinct from `region`; drives Type → State → Organization cascading dropdowns |
| address | Text | |
| parent_id | Integer | self-FK, `ON DELETE SET NULL` — parent/sub-organization hierarchy |
| category_id | Integer | FK → organization_categories.id, `ON DELETE SET NULL` |

Relationship: `children` (`lazy="dynamic"`, self-referential, `backref="parent"`).

### `organization_categories` (`OrganizationCategory`)
`id`, `category_name` (unique, not null), `description`, `is_state_based` (Boolean, default False — flags types like State SLDC/STU/DISCOM that need a State selector step in the cascading dropdown).

### `organization_subcategories` (`OrganizationSubcategory`)
`id`, `category_id` (FK → organization_categories.id, `ON DELETE CASCADE`), `subcategory_name`. A pure suggestion list feeding the Subcategory autocomplete on the Add/Edit Organization form; organizations may carry any subcategory value whether or not it is listed here, so nothing references this table. Organizations store their chosen subcategory in `organizations.region`. Created by migration 018.

### `service_types` (`ServiceType`)
`id`, `service_type_name` (unique, not null — seeded: IAS, IPS, IFS, State Civil Service, Other), `description`.

### `directory_numbers` (`DirectoryNumber`)
| Column | Type | Notes |
|---|---|---|
| id | Integer PK | |
| name | String(255) | not null |
| phone_number, email | Text | |
| organization | String(255) | legacy free-text organization name |
| category | String(100) | free text — classified into Control Room/Switchyard by substring match, no enum |
| organization_id | Integer | FK → organizations.id, `ON DELETE SET NULL` |

### `emergency_contacts` (`EmergencyContact`)
`id`, `employee_id` (FK → employees.id, not null), `contact_name` (not null), `relation`, `phone`.

### `administrative_heads` (`AdministrativeHead`)
| Column | Type | Notes |
|---|---|---|
| id | Integer PK | |
| organization_id | Integer | FK → organizations.id, `ON DELETE CASCADE`, not null |
| employee_id | Integer | FK → employees.id, `ON DELETE CASCADE`, nullable (relaxed from NOT NULL in migration 006) |
| role_category | String(20) | not null; `ADMINISTRATIVE_HEAD` \| `KMP`, SQL CHECK-enforced |
| role_title | String(150) | not null |
| service_type_id | Integer | FK → service_types.id, `ON DELETE SET NULL` |
| effective_from | Date | not null |
| status | String(20) | not null, default `ACTIVE`; `ACTIVE \| ON_LEAVE \| VACANT \| INACTIVE`, CHECK `chk_administrative_head_status` |
| remarks | Text | |
| name, office_phone, mobile_phone, email, office_address | various | standalone identity/contact fields, used only when `employee_id IS NULL` |
| created_by | Integer | FK → users.id |
| created_at | DateTime | server default now() |

Row-level CHECK `chk_administrative_head_identity`: `employee_id IS NOT NULL OR name IS NOT NULL`. **Partial unique index** `uq_active_admin_head (organization_id, role_title) WHERE role_category = 'ADMINISTRATIVE_HEAD'` — only one ADMINISTRATIVE_HEAD-category row per org+title; KMP rows are exempt and can repeat freely.

**Computed properties** (the "resolved_*" pattern, central to this module):
```python
@property
def resolved_name(self):
    return self.employee.employee_name if self.employee_id and self.employee else (self.name or "")
```
Identical fallback logic for `resolved_designation` (employee's designation, else `role_title`), `resolved_office_phone`, `resolved_mobile_phone`, `resolved_email`. `resolved_office_address` differs: prefers the head's own `office_address`, falling back to the organization's address.

### `administrative_head_assistants` (`AdministrativeHeadAssistant`)
PA/PS attached to a head. `administrative_head_id` (FK, `ON DELETE CASCADE`), `employee_id` (FK, `ON DELETE SET NULL`), `designation` (not null, free text e.g. "PA / PS / Executive Assistant"), standalone `name`/`office_phone`/`mobile`/`email`, `is_primary` (Boolean). **Partial unique index** `uq_primary_assistant (administrative_head_id) WHERE is_primary = TRUE` — at most one primary assistant per head. Same `resolved_*` fallback-to-employee pattern as `AdministrativeHead`.

### `administrative_head_history` (`AdministrativeHeadHistory`)
Immutable archive row written by `head_service.end_role()` when a tenure ends: snapshot of everything on the head (organization, employee, role, dates, standalone contact fields, status) plus `effective_to`, `reason`, and `replacement_employee_id`. PA/PS assistants are **deliberately not archived here** — they belong to the live assignment, not a point-in-time snapshot.

### `employee_status_history` (`EmployeeStatusHistory`)
Despite its name, this is the **general per-employee change log** that the Archive/profile timeline is built from — a row is written whenever status, organization, or designation changes, with only the relevant old/new pair populated (`old_status`/`new_status`, `old_organization_id`/`new_organization_id`, `old_designation`/`new_designation` — the latter two pairs added in migration 007). `employee_id` FK is `ON DELETE CASCADE`.

### `update_requests` (`UpdateRequest`)
The public self-service correction queue. `user_id`, `employee_id`, `directory_number_id`, `administrative_head_id` (added migration 013) — all nullable FKs identifying the target record by request type. Newer generic fields: `requested_by`, `department`, `contact_number`, `field_name`, `new_value`, `reason`, `request_type` (default `"employee"`), `status` (default `"Pending"`), `reviewed_by`/`reviewed_at`/`admin_comment`, plus `ip_address`/`browser`/`operating_system` (added migration 007) captured from the submitter's browser for abuse tracking. Older, now largely superseded columns (`requested_mobile`, `requested_office_phone`, `requested_email`, `requested_directory_number`) remain for backward compatibility with pre-generalization submissions.

### `audit_logs` (`AuditLog`)
Append-only. `module`, `record_type`, `record_id`, `action`, `field_name`, `old_value`, `new_value`, `changed_by` (FK users.id), `changed_by_label`, `reason`, `ip_address`/`browser`/`operating_system`/`session_id` (added migration 007), `created_at`. Model docstring: *"No route should ever UPDATE or DELETE a row here."* **There is no admin screen to view this table** — it is write-only from the application's perspective today.

### `import_batches` (`ImportBatch`)
One row per run of `import_from_excel_db.py`/`import_employees_excel.py`. `workbook_name`, `imported_by`, `import_date`, `mode` (`DRY_RUN`\|`APPLIED`), `inserted_count`/`updated_count`/`skipped_count`, `status` (`SUCCESS`\|`FAILED`), `rollback_available` (always `False` — true rollback is not implemented), `summary`.

### `directory_versions` (`DirectoryVersion`)
One row per generated/regenerated official snapshot. `version_number` (unique, e.g. `"2026.08"`, or `"2026.08.2"` for a regeneration), `version_name`, `month`/`year` (CHECK-bounded), `generated_on`/`generated_by`, `employee_count`/`organization_count`, `pdf_filename`/`pdf_path`, `excel_filename`/`excel_path`, `status` (CHECK currently only allows `'Published'`), `remarks`. No employee data is duplicated into this row — only counts and file paths. Regeneration never mutates a prior row; it always inserts a new one with a suffixed version number.

### `email_groups` / `group_members` (`EmailGroup` / `GroupMember`)
`EmailGroup`: `name` (unique), `description`, `list_type` (`STATIC`\|`DYNAMIC`, CHECK-enforced), `created_by`. STATIC groups hold an explicit membership list via `GroupMember` (`group_id`+`employee_id`, both `ON DELETE CASCADE`, unique together). DYNAMIC groups instead hold `EmailGroupFilter` rows and are resolved live — their membership is never a snapshot.

### `email_group_filters` (`EmailGroupFilter`)
One filter rule for a DYNAMIC group. `filter_type` (`ORGANIZATION`\|`ORGANIZATION_CATEGORY`\|`ROLE`\|`STATUS`\|`INCLUDE_EMPLOYEE`\|`EXCLUDE_EMPLOYEE`, CHECK-enforced — the last two added by migration 017), plus exactly one of `organization_id`, `category_id`, `role_value` (`UTILITY_HEAD`\|`ADMINISTRATIVE_HEAD`\|`KMP`), `status_value` (6 of the 8 employee statuses only), `employee_id` (FK → employees.id, `ON DELETE CASCADE`, migration 017) populated, matching `filter_type`. Docstring: *"Rows sharing the same filter_type are ORed together; different filter_types are ANDed — e.g. (category=Transmission Utility OR category=SLDC) AND status=ACTIVE."* `INCLUDE_EMPLOYEE`/`EXCLUDE_EMPLOYEE` rows are the exception to the AND/OR combination rule — `resolve_dynamic_group()` applies them last, as a manual override on top of the combined result, not ANDed/ORed into the filter logic (see Section 11.8).

## 3.3 Migration History (chronological)

| # | File | What it added |
|---|---|---|
| — | `add_parent_id_and_org_fk.sql` | `organizations.parent_id` (self-FK, hierarchy) + `directory_numbers.organization_id` FK. Earliest file by timestamp; effectively "pre-002." |
| 002 | `002_enterprise_features.sql` | Employee status + history; `administrative_heads`/`administrative_head_history` (unified Admin Head/KMP table); `email_groups`/`group_members`; `audit_logs`; `update_requests` review columns. |
| 003 | `003_email_distribution_lists.sql` | `organization_categories` (+10 seed rows); `organizations.category_id`; `email_groups.list_type`; `email_group_filters`. |
| 004 | `004_contact_indexes.sql` | Performance indexes on `employees.organization_id` / `directory_numbers.organization_id`. |
| 005 | `005_administrative_heads_enhancement.sql` | `service_types` (+5 seed rows); `administrative_heads.service_type_id`; `administrative_head_assistants`. |
| 006 | `006_administrative_heads_standalone.sql` | Made `administrative_heads.employee_id` nullable; added standalone name/contact fields + `status`; dropped unused `pa_*` columns. |
| 007 | `007_archive_module.sql` | Widened employee status CHECK (+CONTRACT_ENDED, +DECEASED); extended `employee_status_history`, `audit_logs`, `update_requests` with IP/browser/OS metadata; created `import_batches`. |
| 008 | `008_directory_versions.sql` | Created `directory_versions` (explicitly replacing the earlier "Archive module" concept). |
| 009 | `009_organization_types.sql` | +7 Organization Type seed rows (CTU, STU, State SLDC, IPP, CPSU, DISCOM, Others). |
| 010 | `010_organization_state_and_types.sql` | `organizations.state`; `organization_categories.is_state_based`; +3 seed rows (Thermal, Hydel, Nuclear). |
| 011 | `011_generation_company_type.sql` | +1 seed row ("Generation Company"). |
| 012 | `012_transmission_and_re_generators.sql` | +seed rows preparing a CTU/STU merge and RE Generators split (actual reclassification handled by `scripts/reclassify_organization_types.py`, not SQL). |
| 013 | `013_update_request_administrative_head.sql` | `update_requests.administrative_head_id` — lets the public Request Update flow target Administrative Heads. |
| 014 | `014_directory_number_contact_fields.sql` | `directory_numbers.switch_yard`/`control_room` (explicit boolean flags, additive alongside the existing free-text `category`) + `ip_address`. |
| 015 | `015_remove_ipp_category.sql` | Removed the unused `IPP` organization category (0 organizations ever assigned it — see `reclassify_organization_types.py`). |
| 016 | `016_utility_head_no_head_flag.sql` | `organizations.utility_head_excluded` — lets an admin explicitly mark an organization as having no Utility Head, overriding the always-auto-resolve default. |
| 017 | `017_email_group_member_overrides.sql` | `email_group_filters.employee_id` (FK → employees.id, `ON DELETE CASCADE`); widened `filter_type` CHECK to add `INCLUDE_EMPLOYEE`/`EXCLUDE_EMPLOYEE` — manual per-employee add/remove overrides for DYNAMIC email groups (see Section 11.8). |
| 018 | `018_organization_subcategories.sql` | Creates `organization_subcategories` (per-category subcategory suggestion list, cascade-delete with its category). |
| 019 | `019_remove_re_generators_category.sql` | Removes the empty `RE Generators` category (all 74 of its organizations were merged into Generation Company on 2026-08-27); its one subcategory row cascade-deletes. |

## 3.4 Notable FK `ON DELETE` Behavior

A deliberate asymmetry exists: FKs added in the earliest enterprise migrations that point at a "parent" record whose children have no independent meaning without it were given `CASCADE` — e.g. `administrative_heads.organization_id`/`employee_id` (`CASCADE`), `group_members.group_id`/`employee_id` (`CASCADE`), `email_group_filters.group_id`/`organization_id`/`category_id` (`CASCADE`), `administrative_head_assistants.administrative_head_id` (`CASCADE`). FKs to `users.id` (`created_by`, `changed_by`, `reviewed_by`, `imported_by`, `generated_by`) were deliberately left with no `ON DELETE` clause (defaults to `NO ACTION`). Base-schema FKs (`employees.organization_id`, `employees.department_id`, etc., predating the migration system) also have no `ON DELETE` clause. This is why `admin_routes.py`'s `delete_organization()` route runs an explicit dependency check before allowing a delete — the code comment states plainly: *"Deleting an organization that still has dependents doesn't fail — the DB's ON DELETE rules silently orphan employees/directory numbers/child organizations ... and CASCADE-delete any Administrative Head tied to it."*

---

# 4. Backend Implementation

## 4.1 `app.py` — Composition Root

- Creates `Flask(__name__)`, loads `config.Config`, sets `app.config["SESSION_PERMANENT"] = True`.
- `db.init_app(app)` — `db` is the shared `SQLAlchemy()` instance from `models/__init__.py`.
- `LoginManager()`: `login_view = "auth.admin_login"`, `login_message = ""`, plus a `@login_manager.user_loader` (`db.session.get(User, int(user_id))`) and a `@login_manager.unauthorized_handler` that redirects to the hardcoded path `"/admin-login"`.
- One error handler: `@app.errorhandler(401)` also redirects to `/admin-login`. **No 404 or 500 handler is defined.**
- `@app.context_processor` injects `utility_head_ids` (from `utils.designation_rank.compute_utility_head_ids()`) into every template's context — this is how any template can check `employee.id in utility_head_ids` without each route recomputing it.
- Imports ~13 model modules purely so SQLAlchemy's metadata registry knows about every table (annotated `# noqa: F401`).
- Registers `auth_bp`, `user_bp`, `admin_bp` — no `url_prefix` on any of them.
- Entry point:
```python
debug = os.environ.get("FLASK_DEBUG", "0") == "1"
app.run(host="0.0.0.0", port=8000, debug=debug)
```
Debug is **off by default**, deliberately, with a code comment explaining that Werkzeug's interactive debugger is a remote-code-execution risk if left on and exposed on the network. (This was a production-readiness fix made during this project — see Section 19.)

## 4.2 `config.py`

Single `Config` class (no separate Dev/Prod subclasses):

| Setting | Source | Hardcoded fallback |
|---|---|---|
| `SECRET_KEY` | `os.environ.get("SECRET_KEY", ...)` | `"telephone_directory_secret_key"` |
| `SQLALCHEMY_DATABASE_URI` | `os.environ.get("DATABASE_URL", ...)` | `postgresql+psycopg://postgres:Wrldc%40123@localhost:5432/telephone_directory` |
| `SQLALCHEMY_TRACK_MODIFICATIONS` | hardcoded | `False` |
| `UPLOAD_FOLDER` | computed | `<cwd>/uploads` |
| `DIRECTORY_VERSIONS_STORAGE` | computed | `<cwd>/storage/directory_versions` |
| `PERMANENT_SESSION_LIFETIME` | hardcoded | `timedelta(minutes=20)` |
| `WRLDC_ORG_FILTER` | hardcoded | `"%WRLDC%"` |

Both `SECRET_KEY` and the database URI (including a **plaintext DB password**) fall back to hardcoded values if env vars are unset — flagged as a pre-deployment hardening item in Section 17.

## 4.3 `routes/` — Responsibility of Each File

- **`auth_routes.py`**: `/admin-login` (GET/POST, public) and `/logout` (GET, `@login_required`). Password check via Werkzeug's `check_password_hash`. Blocks login for inactive admin accounts (`admin.status != "active"`).
- **`admin_routes.py`** (1,787 lines): every admin-only page and action — Dashboard, Employee/Organization/Emergency Contact/Directory Number CRUD, Update Request review, Administrative Head lifecycle + PA/PS assistants, custom Email Group CRUD (including the members view and manual add/remove overrides, Section 11.8), Directory Version generate/regenerate/delete, Import History, Export Verification Workbook. Every route decorated with a locally defined `admin_required` (wraps `@login_required` + `current_user.role == "admin"` check).
- **`user_routes.py`** (1,704 lines): every public page — Home, Employee Directory, Telephone Directory, Utility Heads, Administrative Heads (read side), Email Distribution/Address Book, Directory Numbers (read side), Directory Versions (read side), Universal Search + autocomplete, all Exports, the three public "Request Update" submission forms, and the two cascading-dropdown JSON APIs. **No route in this file has any authentication decorator.**

## 4.4 `services/` — Responsibility of Each File

| File | Responsibility |
|---|---|
| `audit_service.py` | `log_audit_event(module, record_type, record_id, action, field_name=None, old_value=None, new_value=None, reason=None)` — the single function every enterprise-module mutation calls to append an `AuditLog` row, capturing `changed_by`/IP/browser/OS/session id automatically from the current request. |
| `head_service.py` | Full Administrative Head lifecycle: `assign_administrative_head`, `update_administrative_head`, `end_role` (archive-then-delete into `AdministrativeHeadHistory`), PA/PS assistant CRUD (`add_assistant`, `update_assistant`, `remove_assistant`, `set_primary_assistant`), and `search_administrative_heads` (the query builder behind the public Administrative Heads list). |
| `email_distribution_service.py` | Resolves both standard organization-category-based contact lists and role-based (`Role Based` section) groups; resolves DYNAMIC `EmailGroup` filters live; builds `to_contact_row()` dicts (including `record_id`/`record_type`, used by the Address Book's client-side CRUD). |
| `directory_version_service.py` | `build_directory_snapshot()` (the data structure behind PDF/Excel generation, deliberately re-implementing the Telephone Directory's own grouping logic to avoid coupling); `generate_version`/`regenerate_version`/`delete_version`/`update_remarks`/`search_versions`/`distinct_years`/`summary_stats`. |
| `excel_generator.py` | Pure function `generate_excel(version, snapshot)` → `.xlsx` bytes, 7 fixed sheets (Employees, Utility Heads, Administrative Heads, KMP, Control Rooms, Switchyards, Emergency Contacts). |
| `pdf_generator.py` | Pure function `generate_pdf(version, snapshot)` → PDF bytes via ReportLab Platypus, with a cover page, an auto-numbered two-pass Table of Contents, per-organization sections, and consolidated role sections. |
| `verification_workbook_service.py` | `generate_verification_workbook(generated_by)` — an admin-only export that independently re-queries live data (mirroring each public page's default view) into a 7-sheet + Summary `.xlsx`, used to sanity-check the application against itself before LAN deployment. |
| `org_stats.py` | `get_organization_statistics(org_id)` — live (uncached) per-organization stats: status breakdown, department count, directory-number count, emergency-contact count, pending-request count. |

## 4.5 `utils/` — Responsibility of Each File

| File | Responsibility |
|---|---|
| `designation_rank.py` | The Utility Head auto-detection algorithm — see Section 12 for the full explanation. |
| `logger.py` | `get_logger(name)` — idempotent console+file logger shared by services and import/reconciliation scripts, writing to `logs/<top-level-package>.log`. |
| `user_agent.py` | `parse_user_agent(ua_string)` — small regex-based Browser/OS parser for audit metadata, filling the gap left by Werkzeug 3.x dropping its own UA parser. |

## 4.6 `migrations/`

Plain, hand-written, numbered `.sql` files applied directly via `psql` — **not** managed by an ORM migration tool (no Alembic/Flask-Migrate). See Section 3.3 for the full chronological list.

## 4.7 Import Scripts (root level)

`init_db.py` (schema bootstrap via `db.create_all()`), `import_employees_excel.py` (WRLDC staff from `uploads/employee list.xlsx`), `import_from_excel_db.py` (the primary KMP/organization/control-room/switchyard import from the verified workbook), plus superseded/exploratory Word-parsing paths (`import_employee.py`, `import_directory_numbers.py`, `import_org_excel.py`, `word.py`, `word_document.py`) and a large set of one-off, hand-curated data-correction scripts (`merge_duplicate_orgs.py`, `set_org_regions.py`, `fix_*.py`, `check_*.py`, etc.). Full detail in Section 15.

## 4.8 Export Scripts

Export logic for Employees/Telephone Directory/Utility Heads/Administrative Heads/single-organization Word cards lives directly inside `routes/user_routes.py` (as `_export_employees_xlsx`/`_export_employees_docx`-style helper functions per feature) rather than in `services/`. The three "big picture" export/report generators — Directory Version PDF/Excel and the admin-only Verification Workbook — do live in `services/` (`pdf_generator.py`, `excel_generator.py`, `verification_workbook_service.py`) since they are reused across generate/regenerate and are non-trivial enough to warrant isolation. Full detail in Section 16.

---

# 5. Frontend Implementation

## 5.1 Template Organization

All 43 templates live flat in `templates/`, no per-blueprint subfolders, all extending `base.html`. There is no frontend build step — Jinja2 renders server-side, Bootstrap is loaded for layout/components, and four small vanilla-JS files provide interactivity. `base.html` also injects a hidden `#app-meta` element carrying server flags (like `data-is-admin`) that JS reads, and implements a 20-minute idle-session timer (mousemove/keydown/click/scroll/touchstart resets a `setTimeout`; on fire, redirects to logout or a "session expired" page depending on whether the viewer is an admin).

## 5.2 Every Page, In Detail

| Page | Template | Access | What it does |
|---|---|---|---|
| Home | `index.html` | Public | Hero search bar (universal autocomplete), 3 stat tiles (employees/organizations/last updated), 5 quick-access cards to the main modules. |
| Employee Directory | `directory.html` | Public view / admin actions | WRLDC-internal staff table (Name, Designation, Office/Mobile, Email, Emergency Contact); search box; Excel/Word export buttons; admin-only Edit/Delete/KMP-toggle column. Non-ACTIVE rows are greyed out with a status badge. |
| Employee Details | `employee_details.html` | Public view / admin status change | Single-employee profile card with Organisation/Contact panels; admin-only status-change form (posts to `/admin/employees/<id>/status`) and a link to Status History. |
| Telephone Directory (landing) | `telephone_directory.html` (`landing` mode) | Public | Category cards (org count + contact count badges) — see Section 9. |
| Telephone Directory (drill-down) | `telephone_directory.html` | Public view / admin actions | Single-expand Bootstrap accordion of organization cards, each containing a combined Employees + Control Rooms + Switchyards table — see Section 9. |
| Universal Search results | `universal_search.html` | Public | Simple two-table layout (Employees, Directory Numbers) — an older/simpler alternative to the category-driven Telephone Directory search. |
| Utility Heads | `utility_heads.html` | Public view / admin actions | Org-Type→State→Organization cascading filter, designation filter, a self-contained keyword-suggestion dropdown (not the shared autocomplete component), Copy Emails / Compose Email / Export Excel-Word action bar. |
| Administrative Heads (list) | `administrative_heads.html` | Public view / admin "Assign Head" | Same cascading filter pattern; active/inactive counts computed inline; "Standalone" badge for heads with no linked Employee record. |
| Administrative Head Details | `administrative_head_details.html` | Public view / admin actions | Official Contact, PA/PS Contact (assistants table), and Office Address cards; admin-only End Role modal with a required reason field. |
| Email Distribution Lists (home) | `email_distribution_home.html` | Public | Section-grouped cards; "Organization Based" cards route to the Address Book (`email_distribution_browse`), "Role Based" cards route to a server-rendered group page (`email_distribution_group`). |
| Distribution List (Role Based) | `email_distribution_group.html` | Public | Server-rendered contact table with per-row checkboxes, Copy/Compose/Export action bar. |
| Address Book (Organization Based) | `email_distribution_browse.html` | Public view / admin Add/Edit/Delete | Fully client-side grid — nearly everything is fetched/rendered by `email_distribution_browse.js`, not Jinja loops. Filter bar (State if applicable, Contact Type, Search). |
| Directory Numbers (admin list) | `directory_numbers.html` | Admin | Control Room/Switchyard/misc number CRUD list. |
| Directory Number Details | `directory_number_details.html` | Public | Single-entry detail page with a public "Request Update" link. |
| Directory Versions | `directory_versions_list.html` | Public view / admin Generate/Regenerate/Delete | Stat cards, filter form, results table with Preview/Download actions; admin modals per row for Regenerate/Delete plus a Generate modal with a live version-number preview. |
| Admin Login | `admin_login.html` | Public | Minimal email/password form. |
| Admin Dashboard | `admin_dashboard.html` | Admin | 4 stat cards + a grid of "Manage" shortcut buttons. |
| Manage Employees | `manage_employees.html` | Admin | Full CRUD list — KMP/Employee toggle, clickable status badge (deep-links to the details page's status form), Utility Head star toggle. |
| Manage Organizations | `manage_organizations.html` | Admin | CRUD list with category badge and employee-count badges (total/active/inactive) per row. |
| Manage Requests | `manage_requests.html` | Admin | Single queue table covering all three request types, with Approve/Reject actions for Pending rows. |
| Manage Emergency Contacts | `manage_emergency_contacts.html` | Admin | CRUD list. |
| Emergency Contacts (public) | `emergency_contacts.html` | Public view / admin Edit-Delete | Present in the codebase but not linked from the main nav — effectively a hidden/legacy page today. |
| Employee Status History | `employee_status_history.html` | Admin | Per-employee From/To/Reason/Changed By/Date table. |
| Import History | `import_history.html` | Admin | `ImportBatch` audit table with mode/status badges; explanatory note that only imports run after this feature existed are tracked. |
| Add/Edit Employee | `add_employee.html` / `edit_employee.html` | Admin | Org-cascade form; KMP checkbox (checked by default on Add); Utility-Head checkbox; Emergency Contact sub-section shown only for non-KMP employees (Edit only). |
| Add/Edit Organization | `add_organization.html` / `edit_organization.html` | Admin | Name/region/category/address form. |
| Add/Edit Administrative Head | `add_administrative_head.html` / `edit_administrative_head.html` | Admin | Dual-mode: link an existing Employee (autocomplete) **or** fill standalone name/contact fields; Edit form adds a read-only "Currently shown across the app" summary box and a warning when an Employee link overrides standalone fields. |
| Add/Edit Assistant | `administrative_head_assistant_form.html` | Admin | Same dual-mode pattern for PA/PS; "Set as Primary" checkbox only shown when adding. |
| Add/Edit Directory Number | `add_directory_number.html` / `edit_directory_number.html` | Admin | Minimal 3-field form (name, phone, category); both now carry a Back button to the Directory Numbers list. |
| Add/Edit Emergency Contact | `add_emergency_contact.html` / `edit_emergency_contact.html` | Admin | Employee dropdown + contact fields. |
| Manage Email Groups | `manage_email_groups.html` | Admin | List of custom DYNAMIC groups with filter-summary text and member/email-count badges (clickable through to the members view). |
| Add/Edit Email Group | `email_group_form.html` | Admin | Multi-select Organization Categories/Organizations + checkbox groups for Roles and Statuses. |
| Email Group Members | `email_group_members.html` | Admin | Live-resolved member table (Organization/Name/Designation/Email) for one custom group, with client-side search; the place to manually add/remove a specific employee via `INCLUDE_EMPLOYEE`/`EXCLUDE_EMPLOYEE` overrides (Section 11.8). |
| Request Update (Employee) | `update_request.html` | Public | See Section 6.7 — the most feature-rich of the three request forms (dynamic Current Information field + status-dropdown swap). |
| Request Update (Directory Number) | `directory_update_request.html` | Public | Same "Current Information" pattern, no dynamic field-type swap (no enum fields). |
| Request Update (Administrative Head) | `administrative_head_update_request.html` | Public | Same pattern as the Directory Number form. |

## 5.3 JavaScript Components (`static/js/`)

All four files are self-contained IIFEs that auto-bind to `data-*`-marked elements on `DOMContentLoaded` — a template opts into a behavior purely by including the right markup, with zero page-specific JS.

- **`autocomplete.js`** — generic, reused everywhere a text box needs suggestions (universal search, Employee Directory search, Telephone Directory search, and the Employee-picker fields in the Administrative Head/Assistant forms). Debounces 160ms, fetches `data-suggestions-url?q=...&source=...`, renders a dropdown panel. Two click behaviors, chosen per input via data attributes: (a) plain search boxes fill the input with the picked value and auto-submit the form; (b) employee-picker inputs (`data-target-id-field` set) instead write the label into the visible box and the id into a hidden field, without submitting (`data-no-autosubmit`), so the rest of the form can still be filled in.
- **`org_type_cascade.js`** — the reusable Organization Type → State → Organization cascading dropdown, auto-initializing on every `[data-org-cascade]` container (Employee forms, Administrative Head/Assistant forms, and the Utility Heads/Administrative Heads filter bars). Hides the State dropdown entirely for non-state-based types; on page load, restores a previously-selected state/organization (for Edit forms and filter bars that need to reconstruct prior selections) by re-running the fetch chain.
- **`email_distribution_browse.js`** — drives the entire Address Book page client-side: fetches filtered contact rows as JSON, renders the grid, computes summary counts, offers a "Show All Contacts instead" recovery affordance when a narrow Contact Type filter returns zero rows, builds admin Edit/Delete buttons per row (via URL templates with an id placeholder), and implements Copy Selected/All/Name+Email/Full Details plus a length-guarded `mailto:` Compose action with a fallback modal.
- **`email_distribution.js`** — the simpler sibling for the server-rendered "Role Based" Distribution List page: same Copy/Compose feature set, but working off data already rendered into the page (`data-rows` JSON) rather than a live fetch.

## 5.4 Bootstrap Usage

Bootstrap supplies the navbar, cards, the single-expand accordion (Telephone Directory), modals (Generate/Regenerate/Delete Directory Version, End Administrative Head Role), badges (status colors, category tags, KMP/Employee type), and dismissible flash-message alerts. No custom CSS framework or design system beyond Bootstrap's defaults plus the app's own navy/white brand palette (also reused in the generated Excel/PDF exports for visual consistency between the web UI and the documents it produces).

---

# 6. User Module

Every feature below lives in `routes/user_routes.py`, and — as established in Section 4 — **none of it requires login**. This is intentional: the directory itself is meant to be freely browsable across the WRLDC LAN, with editing gated behind the Admin role instead.

## 6.1 Searching

Three distinct search surfaces exist, each backed by different logic:
- **Universal Search** (`/search` + `/suggestions`) — searches across Employees and Directory Numbers via `employee_search_filter`/`directory_number_search_filter` helper predicates; the `/suggestions` JSON endpoint additionally covers Organizations and Departments, ranks results with a "startswith-match boosted, then alphabetical" heuristic, and caps at the top 10.
- **Telephone Directory search** (`/telephone-directory?keyword=...`) — keyword-driven "subtree" search: matches organizations and pulls in **all** their child organizations (even if a child itself has zero matching data), matches employees within that subtree, and separately buckets matching Directory Numbers into control_rooms/switchyards/other by substring match on category. Employee results are capped to the first 200.
- **Utility Heads suggestion box** (`/utility-heads/suggest`) — a bespoke, debounced, fetch-based dropdown independent of the shared `autocomplete.js`, returning distinct organization names.

## 6.2 Viewing Employees

`/directory` lists WRLDC-internal staff only (`is_kmp == False`, `status == "ACTIVE"`), with each row's first `EmergencyContact` pre-fetched in one batched query (not N+1). `/employee/<id>` shows the full profile.

## 6.3 Viewing the Telephone Directory

See Section 9 for the full walk-through of the category-card → organization-accordion → contact-table hierarchy.

## 6.4 Utility Heads

`/utility-heads` — see Section 12 for the auto-detection algorithm behind who appears here.

## 6.5 Administrative Heads

`/administrative-heads` and `/administrative-heads/<id>` — see Section 13.

## 6.6 Email Lists / Address Book

`/email-lists`, `/email-lists/<slug>`, `/email-lists/browse` (+ their JSON/export siblings) — see Section 11. Admin-created custom (DYNAMIC) groups also appear on `/email-lists` as a "Custom Groups" section of cards, each linking to `/email-lists/custom/<id>` (a public members-and-emails view, `email_distribution_custom_group.html`; 404s for non-DYNAMIC groups).

## 6.7 Request Information Update

Three public, unauthenticated POST endpoints let anyone submit a correction request against an Employee, a Directory Number, or an Administrative Head:
- `/employee/request-update/<id>`
- `/directory-number/request-update/<id>`
- `/administrative-head/request-update/<id>`

Each renders a form with: requester name/department/contact (all but name optional), a "Field to Update" dropdown (options vary by target — e.g. Employee offers mobile/office/residence phone, email, designation, address, and (for non-KMP employees only) three Emergency Contact sub-fields, plus **Status**), a **read-only, dynamically-populated "Current Information" field** that shows the record's present value for whichever field is selected (looked up client-side from a `current_values` JSON payload the route already computed server-side — no extra request), and a "Correct Information" field for the proposed new value. For the Employee form specifically, selecting **Status** swaps that free-text input for a `<select>` populated from the same `EMPLOYEE_STATUSES` enum used everywhere else in the app, so a status correction can only ever be a valid value — this was added specifically so that anyone (not just an admin) can report that a person has resigned, retired, transferred, or otherwise changed status, with the change taking effect only after admin approval. Submissions capture the requester's IP, browser, and OS (via `utils.user_agent.parse_user_agent`) for the admin's review context.

## 6.8 Directory Versions

`/directory-versions` — browse published monthly snapshots (filter by month/year/keyword, paginated); Preview/Download PDF/Download Excel just `send_file` the pre-generated files written to disk at generation time (Section 16.1).

## 6.9 Export Features

Employees, Telephone Directory, per-organization Word cards, Utility Heads, and Administrative Heads can each be exported (`xlsx`/`docx`/`csv` depending on the feature) directly from their respective list pages — see Section 16.

---

# 7. Admin Module

Every route below lives in `routes/admin_routes.py` and is protected by the locally defined `admin_required` decorator (`@login_required` + `current_user.role == "admin"`).

## 7.1 Login

`/admin-login` (in `auth_routes.py`) — email/password form, `check_password_hash` verification, blocks inactive admin accounts, redirects to the Dashboard on success.

## 7.2 Dashboard

`/admin/dashboard` — 4 live-computed stat cards (Employees, Organizations, Departments, Pending Requests) and a grid of shortcut buttons into every management screen.

## 7.3 Employee Management

Search/list (`/admin/employees`, keyword search across name/designation/email/mobile/org/department via `ilike`), Add/Edit (with organization-type-mismatch and duplicate-name-in-same-org guards), Delete (audit-logged), KMP toggle and Utility-Head toggle (both single-click, audit-logged for KMP), Status change (delegates to the shared `_apply_employee_status_change` helper — see Section 10.4), and Status History (`/admin/employees/<id>/status-history`).

## 7.4 Organization Management

Search/list (with a per-organization employee status breakdown computed via a single grouped query), Add/Edit (optional Subcategory field with autocomplete from the category's suggestion list; uniqueness check, falls back to the seeded "Others" category if none chosen, audit-logs category changes), Delete — **guarded**: blocks deletion if the organization still has any dependent Employees, Directory Numbers, child Organizations, or Administrative Heads, specifically because the database's own `ON DELETE` rules would otherwise silently orphan or cascade-delete related data (Section 3.4).

### Organization Categories (`/admin/categories`)
Reached from the Admin Dashboard's "Add Category" card. Admins can add/edit/delete Organization Categories (name, description, state-based flag) and add/remove Subcategory suggestions per category, with no code change or migration. A category cannot be deleted while any organization uses it; a suggestion can always be removed. "Sync from live data" (`POST /admin/categories/sync`) adds suggestions for combinations in live use and removes ones nothing matches. Routes: `/admin/categories/add`, `/edit/<id>`, `/delete/<id>`, `/<category_id>/subcategories/add`, `/subcategories/delete/<id>`; suggestions JSON at `GET /api/organization-categories/<id>/subcategories`.

## 7.5 User Management

**Not implemented as an admin screen.** No route exists anywhere to create, edit, disable, or list `User` (login account) records through the web UI — the only code path touching `users` besides login/session loading is the standalone `create_admin.py` script. This should be flagged explicitly to the GM as a known gap, not assumed to exist.

## 7.6 Administrative Heads

Assign (`/admin/administrative-heads/add`), Edit, End Role (requires a reason; archives into `AdministrativeHeadHistory` and cascades-deletes the live row's assistants), and full PA/PS Assistant CRUD including "Set as Primary" — all delegating to `services/head_service.py` (Section 13).

## 7.7 Directory Versions

Generate (blocked if a version already exists for that month/year — must Regenerate instead), Regenerate (always allowed, always creates a new suffixed version, never touches the prior one), Delete (removes both the DB row and its PDF/Excel files from disk), Update Remarks — all delegating to `services/directory_version_service.py`.

## 7.8 Update Request Management

`/admin/requests` — single queue listing all Pending/Approved/Rejected requests across all three types. Approve (`/admin/requests/approve/<id>`) branches on `request_type`: for Administrative Head requests, contact fields (`office_phone`, `mobile_phone`, `email`) are routed to the linked `Employee` if the head has one (else the head's own fields) — mirroring the `resolved_*` read-side logic on the write side; for Employee requests, a `field_name == "status"` special case delegates to `_apply_employee_status_change` instead of a plain `setattr`, so a status-change request gets the exact same history/audit/auto-archive side effects as an admin using the direct status dropdown. Reject just marks the row Rejected.

## 7.9 Employee Status

Covered in 7.3/10.4 — status changes (whether from the admin dropdown or an approved public request) always go through the same helper, guaranteeing consistent side effects regardless of entry point.

## 7.10 Import History

`/admin/import-history` — paginated (50/page) read-only view of `ImportBatch` rows. Explicitly an internal tool, not linked from the public UI, and only covers imports run **after** this feature was added (older import runs predate the table and aren't retroactively visible).

## 7.11 Audit Logs

**Write path exists; read/browse UI does not.** `log_audit_event(...)` is called from Employee delete/KMP-toggle/status-change, Organization add/edit/delete, and all Email Group CRUD — every one of those actions is captured in `audit_logs`. But there is no admin route or template anywhere that lists or filters that table. Today, inspecting the audit trail requires a direct database query. This is called out explicitly in Sections 17 and 23 as a near-term enhancement, not something to imply already exists.

## 7.12 Exports

`/admin/export-verification-workbook` — an admin-only, live-data cross-check export (Section 16.3). All other export routes (Employees, Telephone Directory, Utility Heads, Administrative Heads, Email Lists) are public, not admin-gated, since exporting is a read-only convenience available to any user.

---

# 8. Technical Implementation

This section explains *how* specific mechanisms work at the code level — useful for answering "how did you build X" questions.

## 8.1 Search Implementation

Search is implemented with plain SQLAlchemy `ilike` (case-insensitive `LIKE`) predicates combined with `or_()`, not a search engine (no Elasticsearch/Postgres full-text search/trigram indexes). For example, `manage_employees()`'s keyword search:
```python
query = Employee.query.outerjoin(Organization).outerjoin(Department)
query = query.filter(or_(
    Employee.employee_name.ilike(f"%{kw}%"),
    Employee.designation.ilike(f"%{kw}%"),
    Employee.email.ilike(f"%{kw}%"),
    Employee.mobile_phone.ilike(f"%{kw}%"),
    Organization.organization_name.ilike(f"%{kw}%"),
    Department.department_name.ilike(f"%{kw}%"),
))
```
`outerjoin` is used (not `join`) so employees with no organization/department still appear. The Telephone Directory's keyword search is more involved — it first resolves a matching-organization "subtree" (parent + all children) before filtering employees and directory numbers, so a search for a parent utility's name surfaces all of its sub-stations even if the substation name itself doesn't match.

## 8.2 SQLAlchemy Query Patterns Used

- **Outer joins** for optional relationships (Employee↔Organization, Employee↔Department, AdministrativeHead↔Employee) so rows with a NULL FK are never silently dropped.
- **Aggregate/group-by queries** for stats, e.g. `manage_organizations()`'s per-org status breakdown:
```python
db.session.query(Employee.organization_id, Employee.status, func.count(Employee.id))
    .group_by(Employee.organization_id, Employee.status).all()
```
computed once and looked up in Python per row, rather than one query per organization.
- **`db.or_()` / `db.and_()`** for combined filter logic, including the "defensive" employee-status check in `search_administrative_heads`: `db.or_(AdministrativeHead.employee_id.is_(None), Employee.status == "ACTIVE")`.
- **Flush-before-audit-log pattern** — several services call `db.session.flush()` immediately after adding a new row specifically to obtain its auto-generated `id` before writing the paired `AuditLog`/history entry that references it, without a full commit.

## 8.3 Pagination

Only two features paginate: **Import History** (`/admin/import-history`, fixed 50/page, manual `offset()/limit()`) and **Directory Versions** (`/directory-versions`, via `directory_version_service._paginate`, default 25/page). Every other list page (Employees, Organizations, Directory Numbers, Utility Heads, Administrative Heads, Telephone Directory results) renders its full result set unpaginated — acceptable at the current data scale, but worth knowing as a scaling limit if the data volume grows substantially (Section 23).

## 8.4 Filtering

Structured filters (organization type/state/organization/category/service type/status) are applied at the SQL level via query `.filter()` chaining; free-text keyword filters are layered on top, sometimes in Python rather than SQL when the match target spans a computed property (e.g. `search_administrative_heads`'s keyword pass matches against `resolved_name`, which doesn't exist as a database column — it can only be evaluated after the ORM objects are loaded).

## 8.5 Cascading Dropdown

Organization Type → State → Organization, implemented as two small JSON APIs plus `static/js/org_type_cascade.js`:
- `GET /api/organization-types/<type_id>/states` — distinct, non-null `Organization.state` values for that category.
- `GET /api/organizations?type_id=&state=` — organizations matching both filters.
The State step is entirely skipped (hidden) for non-state-based categories, driven by an `is_state_based` flag baked into each `<option>`'s `data-state-based` attribute so the client never needs to re-ask the server whether a type needs a state step.

## 8.6 Fetch API / AJAX / JSON APIs

No AJAX library — every dynamic call uses the browser's native `fetch()`. Full inventory of JSON endpoints:

| Endpoint | Auth | Returns |
|---|---|---|
| `GET /admin/employees/suggest` | Admin | Employee id-picker suggestions |
| `GET /api/organization-types/<id>/states` | Public | Cascading dropdown states |
| `GET /api/organizations` | Public | Cascading dropdown organizations |
| `GET /utility-heads/suggest` | Public | Plain list of organization names |
| `GET /suggestions` | Public | Universal search-box suggestions |
| `GET /api/organization-categories/<id>/subcategories` | Admin | Subcategory suggestions for a category |
| `GET /email-lists/<slug>/emails.json` | Public | Group email list + precomputed mailto string |
| `GET /email-lists/browse/data.json` | Public | Address Book grid rows + summary |

## 8.7 Dynamic UI Patterns

- **Debounced autocomplete** (160ms in `autocomplete.js`, 200ms in the bespoke Utility Heads suggestion box).
- **Single-expand accordion** (Telephone Directory — `data-bs-parent` makes Bootstrap auto-collapse siblings).
- **Auto-expand-if-single-result** (`{% set auto_expand = grouped|length == 1 %}` in `telephone_directory.html`).
- **Input-type swap** — the Employee Request Update form swaps a text `<input>` for a `<select>` when "Status" is chosen as the field to update, so only a valid enum value can be submitted.
- **Live current-value lookup** — all three Request Update forms populate a read-only "Current Information" field client-side from a small JSON payload the server already rendered, with no extra network round-trip.
- **mailto: length guard** — Compose Email actions check the built `mailto:` URL's length client-side (Address Book) or server-side (`emails.json`'s `mailto_safe` flag) and fall back to a "copy emails" modal if it would exceed practical URL length limits.

## 8.8 Session Management

Flask-Login manages the session cookie (`login_user`/`logout_user`/`current_user`). `PERMANENT_SESSION_LIFETIME = timedelta(minutes=20)` combined with `app.config["SESSION_PERMANENT"] = True` in `app.py` means sessions expire after 20 minutes of *inactivity at the server level*; `base.html`'s inline idle-timer additionally proactively redirects the browser to logout/session-expired after 20 minutes of no user interaction, so an unattended session doesn't sit on an active-looking page indefinitely.

## 8.9 Authentication

Email+password only, via `routes/auth_routes.py`. No self-registration flow exists for admin accounts (`create_admin.py` is the only provisioning path) and no "forgot password" flow exists anywhere in the codebase.

## 8.10 Authorization

Single binary role check — `current_user.role == "admin"` — applied per-route via the local `admin_required` decorator. There is no finer-grained permission system (no "editor" vs "super-admin" distinction, no per-module permissions).

## 8.11 Password Hashing

Werkzeug's `generate_password_hash`/`check_password_hash` (used in `create_admin.py` and `auth_routes.py` respectively) — industry-standard salted hashing, not plaintext, not a custom scheme.

## 8.12 Audit Logging

Covered in depth in Section 4.4/7.11 — a single `log_audit_event()` call site pattern, write-only today (no viewer UI).

## 8.13 Import Pipeline (implementation)

Covered fully in Section 15.

## 8.14 Export Generation (implementation)

Covered fully in Section 16.

---

# 9. Telephone Directory Module

This is the flagship public module and the one most likely to draw detailed functional questions — it is generated **entirely from live PostgreSQL data** on every request, with no caching layer.

## 9.1 Category Cards (Landing State)

When `/telephone-directory` is requested with neither a `keyword` nor a `category_id`, the route returns `landing=True` and a list built by `_compute_telephone_directory_categories()`: for every `OrganizationCategory`, it counts member organizations and active-KMP contacts, and the template renders one card per category with those two counts as badges. Clicking a card navigates to `/telephone-directory?category_id=N`.

## 9.2 Organization Cards / Accordion Behavior

With a `category_id` (no keyword), the route groups by organization *within that category* and pre-seeds the `grouped` dict with **every** organization in the category up front — not only ones that happen to have an employee or directory-number row. (This fixed a real bug found during production-readiness review: 55 organizations across 6 categories — mostly private TBCB transmission SPVs listed by name/address only — were previously invisible because the grouping dict was only populated reactively as employee/number rows were encountered.) The template then renders one Bootstrap card per organization, `data-bs-parent="#groupedOrgAccordion"` making it a **true single-expand accordion** — opening one organization automatically collapses whichever was previously open. If the current filter/search happens to yield exactly one organization, that card starts already expanded (`{% set auto_expand = grouped|length == 1 %}`) so the user isn't forced to click through to a lone result.

## 9.3 Search

With a `keyword`, the route builds a "subtree" match: any organization whose name matches the keyword contributes itself **and all of its child organizations** (via `parent_id`) to the result set — explicitly including children that have zero employees/numbers of their own, since a parent match implies the user is interested in the whole group. Employees are then queried via `Employee.organization_id.in_(subtree_ids)` OR a direct keyword match on the employee's own fields, capped to the first 200 results. Directory Numbers matching the keyword are separately bucketed into control_rooms/switchyards/other by a case-insensitive substring check on their free-text `category` field.

## 9.4 Employee Details Within the Directory

Each organization's card shows a combined table: employees (name, designation, office/mobile phone, email — heads-first via `utility_head_ids`), then Control Rooms (highlighted `table-warning`), then Switchyards (`table-info`). Only **KMP, ACTIVE** employees are shown in the Telephone Directory (WRLDC-internal, non-KMP staff belong to the separate Employee Directory module, Section 10).

## 9.5 Control Rooms and Switchyards

Both are `DirectoryNumber` rows classified purely by case-insensitive substring matching on the free-text `category` column (`"control room" in category.lower()`, `"switchyard" in category.lower()`) — there is no dedicated enum/type column for this distinction anywhere in the schema. Anything that doesn't match either pattern falls into a residual "other numbers" bucket, shown separately (or, in a keyword search, under an "(Other / Unassigned Numbers)" label in the Verification Workbook's equivalent export).

## 9.6 Utility Heads Within the Directory

Utility Head status is purely a **display affordance** here (a star badge/heads-first sort), computed via the same `compute_utility_head_ids()` algorithm used everywhere else — see Section 12. It is not a separate section of the Telephone Directory; it's layered onto the same employee rows.

## 9.7 Organization Hierarchy

`Organization.parent_id` (self-referential FK, `ON DELETE SET NULL`) is the only hierarchy mechanism — a flat one level of parent/child in practice (e.g. a Discom parent with named sub-station "child" organizations), used both for the subtree search behavior above and for the Organization Type → State → Organization cascading dropdown (which resolves at the Organization level, not per-hierarchy-level).

## 9.8 How the Page Is Generated, End to End

1. Route receives `keyword`/`category_id` query args.
2. Branches: landing (no args) → category counts; category only → full-category grouping (pre-seeded so empty organizations still show); keyword → subtree search.
3. Employees/DirectoryNumbers are queried per the branch's logic above.
4. A `grouped` dict (`{organization_name: {"employees": [...], "control_rooms": [...], "switchyards": [...], "address": ...}}`) is assembled and, for the category branch, sorted alphabetically.
5. `telephone_directory.html` renders category cards (landing) or the organization accordion (category/keyword), with Excel/Word export buttons that re-run the *same* keyword/category logic server-side at export time (Section 16.2) so the exported file always matches what's on screen.

---

# 10. Employee Directory

## 10.1 Search

`/directory?keyword=...` filters via `employee_search_filter(keyword)` (name/designation/email/mobile, `ilike`-based) scoped to `is_kmp == False`.

## 10.2 Filters

The list is implicitly filtered to `is_kmp == False` and `status == "ACTIVE"` at all times — there is no UI toggle to include inactive or KMP staff on this page (KMPs belong to the Telephone Directory; inactive staff are only visible via Manage Employees in the admin back office or via an employee's own Status History).

## 10.3 Employee Details

`/employee/<id>` — profile card with a colored header (status badge, Utility Head star if applicable), Organisation and Contact info panels.

## 10.4 Status

`Employee.status` is a Python/SQL-enum-like field (`ACTIVE, INACTIVE, TRANSFERRED, RETIRED, DEPUTATION, RESIGNED, CONTRACT_ENDED, DECEASED`). Every status change — whether made directly by an admin (`/admin/employees/<id>/status`) or approved from a public "Status" Request Update — funnels through one shared function, `_apply_employee_status_change(employee, new_status, reason)`, which:
1. Sets `employee.status` and `status_changed_at`.
2. Writes an `EmployeeStatusHistory` row (old/new status, reason, who changed it).
3. Calls `log_audit_event(module="employee_status", action="STATUS_CHANGE", ...)`.
4. If the new status is anything other than `ACTIVE`, automatically ends (via `head_service.end_role`) any `AdministrativeHead` role currently held by that employee — so a retired/transferred person doesn't linger as a listed Administrative Head after their employee status changes.

## 10.5 Emergency Contacts

One-to-many via `EmergencyContact.employee_id`. Shown only for non-KMP employees in both the public directory and the admin Edit form (a KMP is a region-wide contact, not WRLDC's own staff, so an emergency contact doesn't apply the same way). The directory list page pre-fetches each visible employee's first emergency contact in a single batched query rather than one query per row.

## 10.6 Exports

`/export/employees` — Excel or Word, restricted to WRLDC organization IDs (matched via `ilike` against "%WRLDC%" and two spelling variants of "Western Region Load Dis/spatch"), `status == "ACTIVE"`, with optional keyword/organization/department filters reapplied at export time.

## 10.7 Backend Logic Summary

`Employee.outerjoin(Organization).outerjoin(Department)` → filter `is_kmp=False, status=ACTIVE` → optional keyword filter → order by name → batched emergency-contact lookup → render.

---

# 11. Email Distribution Lists

## 11.1 Organization Categories

The "Organization Based" section of `/email-lists` iterates `OrganizationCategory` rows (skipping a section literally named `"Role Based"`, which is handled separately) and computes per-category contact counts; empty categories are skipped from the card grid.

## 11.2 Address Book

`/email-lists/browse?category=...` — the fully client-side grid page. The server only renders the shell and validates the `category` query param against `BROWSE_CATEGORY_NAMES`; all row data comes from `GET /email-lists/browse/data.json`, which calls `resolve_category_contacts(category, contact_type, state=...)` and applies a `q` (search) filter server-side, sorted by organization then name. This split (server validates + serves shell, client fetches + renders + filters) is why the Address Book supports instant client-side search without a full page reload.

## 11.3 Distribution Lists

The "Role Based" section's groups render via `email_distribution_group.html` — a server-rendered (not client-fetched) contact table, with an optional contact-type filter that auto-submits on change.

## 11.4 Contact Types

A `contact_type` concept (e.g. filtering to only Utility Heads, or only Administrative Heads, within a category/group) is threaded through both the Address Book and Distribution List query paths as a query-string filter.

## 11.5 Copy Features

Both the Address Book and Distribution List pages offer: Copy Selected / Copy All / Copy Name+Email / Copy Full Details (via `navigator.clipboard.writeText`, with de-duplicated, alphabetically-sorted email lists computed by a shared `dedupeSortedEmails()` helper in each JS file), and a Compose Email action that builds a `mailto:` link — guarded by a length check (client-side on the Address Book, server-precomputed `mailto_safe` flag on the Distribution List) that falls back to a "copy emails instead" modal when the recipient list is too long to fit in a `mailto:` URL.

## 11.6 Export

Both surfaces offer `.xlsx` and `.csv` exports (`/email-lists/<slug>/export.xlsx`/`.csv`, `/email-lists/browse/export.xlsx`/`.csv`), each re-running the same resolution/filter logic as the on-screen view at export time.

## 11.7 Backend Queries

`services/email_distribution_service.py` is the shared resolution layer: `to_contact_row()` builds a uniform contact dict (including `record_id`/`record_type` — `employee`/`directory_number`/`administrative_head` — which is what lets the Address Book's client-side JS build correct per-row Edit/Delete URLs regardless of the underlying record type), DYNAMIC `EmailGroup`s are resolved **live** every time (never cached/snapshotted — their membership always reflects the current database), and a `dedupe_and_sort()` step removes duplicate contacts (e.g. someone who is both a Utility Head and an Administrative Head of the same org) before counting/rendering.

## 11.8 Custom (Saved Dynamic) Groups — Membership View & Manual Overrides

Admin-only, at `/admin/email-groups`. A custom group (`EmailGroup.list_type == "DYNAMIC"`) is built from `EmailGroupFilter` rows (Organization Categories/Organizations, Roles, Status — OR within a type, AND across types) and always resolves **live** from `resolve_dynamic_group()`, never as a stored snapshot.

- **Viewing members**: the Members column on `manage_email_groups.html` shows two badges — total matched count and unique-email count — both linking to `/admin/email-groups/<id>/members` (`email_group_members.html`), a server-rendered table (Organization, Name, Designation, Email) with a client-side search box, resolved fresh on every visit so it can never go stale relative to the underlying Employee/Administrative Head data.
- **Manual add/remove**: since membership is filter-computed, there's no membership list to directly edit. `EmailGroupFilter.filter_type` has two override values added for this — `INCLUDE_EMPLOYEE` and `EXCLUDE_EMPLOYEE` (each pairs with the row's `employee_id` FK, migration 017) — which `resolve_dynamic_group()` applies **last**, after the category/role/status filters, so a manual override always wins regardless of whether the person matches the group's filters. `POST /admin/email-groups/<id>/members/add` (with an `employee_id`) adds an `INCLUDE_EMPLOYEE` override (first clearing any prior override for that same employee, so re-adding someone who was excluded cancels the exclusion); `POST /admin/email-groups/<id>/members/remove` adds an `EXCLUDE_EMPLOYEE` override the same way. `_build_filter_summary()` appends a `+N manually added` / `−N manually removed` suffix to the group's filter-summary text whenever overrides exist, so the list page always shows that a group's membership isn't purely filter-derived.
- **Changing who's in a group otherwise**: the two non-override ways are (a) editing the group's filters themselves (`email_group_form.html`, Category/Organization/Role/Status), or (b) editing the underlying Employee's own status/role/organization, since those drive filter matching directly.

---

# 12. Utility Heads

## 12.1 Auto Detection

Utility Heads are **never manually flagged** in normal operation — `Employee.is_utility_head` exists as a column but is not what the Utility Heads module actually uses (that boolean is a leftover/independent toggle exposed as an admin convenience star-toggle, not the source of truth). The real mechanism is `utils/designation_rank.py`, which infers seniority purely from the free-text `designation` field.

## 12.2 Designation Ranking — the Algorithm

`RANKING_TIERS` is an ordered list of 30 seniority tiers (most senior first), each listing full titles and abbreviations that resolve to that tier — e.g. tier 1 `["chairman"]`, tier 2 `["chairman & managing director", "cmd"]`, tier 3 `["managing director", "md"]`, tier 9 `["general manager", "gm"]`, tier 12 `["chief engineer", "ce"]`, down through Manager/Engineer/Officer/Supervisor tiers to the least senior, `["operator"]`. At import time, every phrase in every tier is precompiled into a case-insensitive, whole-word regex (`\bphrase\b`).

`rank_designation(designation)` scans **every** tier/phrase (not stopping at the first hit) and keeps the best match by this rule: *the longest matching phrase wins*, regardless of which tier it came from. This deliberately prevents a more specific/senior title like "Additional Chief Engineer" from being shadowed by the shorter substring "Chief Engineer" also matching — since more senior titles tend to be worded more specifically/longer, this naturally favors the correct (more senior) tier in the common case, though the code comment notes it's theoretically possible (not otherwise guarded against) for a longer phrase in a *less* senior tier to win over a shorter phrase in a more senior tier if both match the same text.

`resolve_utility_head(employees)` picks the single employee with the numerically lowest rank from a given organization's employee list (ties broken by first-occurrence order, since the comparison uses strict `<`). `compute_utility_head_ids()` is the batch version: loads all ACTIVE employees, buckets by organization, resolves each organization's head, and returns the full set of employee ids — recomputed fresh on **every call**, never cached or read from a stored flag. This is what `app.py`'s context processor injects into every template as `utility_head_ids`.

## 12.3 Filters

`/utility-heads` supports the Organization Type → State → Organization cascade, a free-text designation filter, and a keyword search with its own dedicated (non-shared) autocomplete dropdown.

## 12.4 Search

Keyword search resolves via `/utility-heads/suggest`, returning distinct matching organization names (not employee names).

## 12.5 Edit Workflow

There is no dedicated "edit utility head" form — because a Utility Head is *computed*, not stored, the only way to change who appears is to edit the underlying employee's `designation` (via the normal Employee edit form), which will cause `resolve_utility_head` to select a different person on the next page load automatically.

---

# 13. Administrative Heads

## 13.1 How Administrative Heads Are Determined

Unlike Utility Heads, Administrative Heads are **not auto-detected**. They are explicit, manually curated rows in the `administrative_heads` table (`role_category = "ADMINISTRATIVE_HEAD"`), assigned and maintained by an admin through `services/head_service.py`'s `assign_administrative_head`/`update_administrative_head`/`end_role` functions. A partial unique index enforces at most one *active-category* Administrative Head per organization+role_title combination, while `KMP`-category rows (Company Secretary, etc.) are allowed to repeat freely.

## 13.2 IAS/CMD/MD Logic

There is no automatic classification of *which* titles count as "Administrative Head" vs "KMP" — that distinction is made once, by the admin, at assignment time via the `role_category` field. `ServiceType` (IAS/IPS/IFS/State Civil Service/Other) is a separate, optional classification attached to a head purely for filtering/reporting (e.g. "show me all IAS-cadre heads") — it plays no role in determining seniority or whether someone is shown as a head at all.

## 13.3 Manual Management

A head can be either **linked to an existing Employee** (identity/contact fields are then always read live via the `resolved_*` properties, and any standalone fields on the head row itself are forced to `None` and ignored) or **standalone** (a person with no Employee record at all — e.g. a government-appointed official external to the utility's own staff, such as a Principal Secretary) with their own `name`/`office_phone`/`mobile_phone`/`email` stored directly on the `administrative_heads` row. Ending a role (`head_service.end_role`) archives a full snapshot into `AdministrativeHeadHistory` (including the reason and an optional replacement employee) and then deletes the live row — which cascades to delete that head's PA/PS assistants too, since the model explicitly treats assistants as belonging to the *live* assignment, not a point-in-time historical record.

## 13.4 Search

`head_service.search_administrative_heads()` combines SQL-level structured filters (organization, service type, status, category, state, designation `ilike`) with a Python-level keyword pass (since keyword matching needs each head's *resolved* name/assistant names, which don't exist as plain database columns) — and always applies a defensive filter excluding employee-linked heads whose linked employee is not currently ACTIVE (a safety net in addition to the automatic role-ending triggered by `_apply_employee_status_change`).

## 13.5 Filters

Category, State, Organization, Service Type, Status, and free-text Designation/keyword — identical cascading-dropdown pattern to Utility Heads.

---

# 14. Organization Classification

## 14.1 Organization Type

Backed by `OrganizationCategory` (`organization_categories` table) — a controlled taxonomy that grew across five migrations (003, 009, 010, 011, 012) from an initial 10 seed values (Transmission Utility, Generator, RE Generator, Distribution Company, SLDC, RLDC, Central Utility, Government, Private Utility, Other) to include CTU, STU, State SLDC, IPP, CPSU, DISCOM, Others, Thermal, Hydel, Nuclear, Generation Company, and a later-added, more precisely named "RE Generators" — with an explicitly noted intent (in migration 012's own comment) to eventually merge CTU+STU back into "Transmission Utility" and split RE Generators cleanly out of IPP, a reclassification carried out by `scripts/reclassify_organization_types.py` rather than in SQL.

## 14.2 State

`Organization.state` (migration 010) is distinct from the older, looser `region` field — `region` started as a free-text "parent import grouping label" and was later reused by `set_org_regions.py` to hold Indian state names for search purposes on organizations that predate the dedicated `state` column, so both fields can carry state-like information depending on when a given organization row was created/touched.

## 14.3 Cascading Dropdowns

Type → (State, only if `OrganizationCategory.is_state_based`) → Organization — implemented identically everywhere it's needed (Employee forms, Administrative Head/Assistant forms, Utility Heads/Administrative Heads filter bars) via the shared `org_type_cascade.js` and the two cascading JSON APIs (Section 8.5).

## 14.4 Category Mapping

New organizations created by the import pipeline default to the catch-all "Others" category if not explicitly classified — `import_from_excel_db.py` states this directly in its own comments, noting that `scripts/reclassify_organization_types.py` must be re-run afterward to restore accurate classifications, since the import step alone has no way to know them.

## 14.5 Reclassification Script

`scripts/reclassify_organization_types.py` (and its sibling `scripts/classify_organizations.py`) exist specifically to backfill/correct `category_id` (and related state/type taxonomy) after a bulk import, decoupling "get the data into PostgreSQL" from "correctly classify every organization" as two separate, independently re-runnable steps.

## 14.6 Why This Approach Was Chosen

A single `organizations` table with a nullable `category_id` FK (rather than a separate table per organization type) keeps the schema simple and lets an organization's classification be corrected after the fact without any data migration between tables — critical given how much of this project's real-world effort went into iteratively discovering the *correct* classification for hundreds of organizations sourced from a Word document that itself grouped things inconsistently. Treating "import" and "classify" as separate, repeatable passes (rather than trying to get classification perfectly right inside the import script itself) is what made it practical to re-run reclassification multiple times as the category taxonomy itself evolved (migrations 009–012 each added new categories that earlier-imported rows needed to be moved into).

---

# 15. Import Pipeline

The project went through **two generations** of import pipeline, both discoverable directly from the scripts' own docstrings.

## 15.1 Generation 1 — Direct Word → Database (superseded)

`import_employee.py`, `import_directory_numbers.py`, `word.py`, `word_document.py` parsed the master Word directory directly with `python-docx` and wrote straight into `Organization`/`Employee`/`Department`/`DirectoryNumber`. `import_from_excel_db.py`'s own header comment confirms this path was abandoned in favor of Generation 2: *"[the old source workbook] no longer exists (superseded by the validation/correction/normalization pipeline that produced v3)."* `word.py` in particular is a standalone MySQL-based prototype, not even wired to this app's PostgreSQL/SQLAlchemy stack — a clear artifact of an earlier exploratory phase.

## 15.2 Generation 2 — Word → Validated/Corrected Excel → PostgreSQL (current pipeline)

```
Word (.docx, official source of truth)
   │
   ▼
scripts/repair_workbook.py ─────► WR_DB_Ready_Final_Verified.xlsx  (+ reports/repair_change_log.csv)
   │  edits a COPY only; source workbook is never opened in write mode
   ▼
scripts/validate_workbook.py    (read-only gate: duplicate IDs, blank required fields, orphan org/suborg refs)
   │
   ▼
scripts/compare_word_excel.py ──► reports/comparison_report.csv   (Word vs Excel: missing employees/control rooms per station)
scripts/reconcile_word_database.py ──► reports/directory_reconciliation_report.csv, reports/perfect_stations.csv
   │        (Word vs live DB, record-level: name + normalized phone digits + email must all match for a station to be "Perfect")
   ▼
scripts/generate_validation_report.py ──► Validation_Report.xlsx
   │
   ▼
scripts/apply_corrections.py ──► WR_DB_Ready_Final_Verified_v2.xlsx  (+ Correction_Log.xlsx)
   │  only evidence-backed, unambiguous corrections auto-applied
   ▼
scripts/normalize_formatting.py ──► WR_DB_Ready_Final_Verified_v3.xlsx  (+ Formatting_Normalization_Log.xlsx)
   │  standardizes capitalization/spacing/case (org names, designations, CR/switchyard labels, emails)
   ▼
import_from_excel_db.py --apply ──► PostgreSQL  (organizations, KMP employees, control rooms, switchyards)
import_employees_excel.py --apply ──► PostgreSQL (WRLDC-internal employees, from uploads/employee list.xlsx, run separately)
   │
   ▼
scripts/classify_organizations.py / scripts/reclassify_organization_types.py ──► backfills category_id / state / type
   │
   ▼
Manual hand-curated cleanup pass: set_org_regions.py, merge_duplicate_orgs.py (guided by find_duplicate_orgs.py),
fix_data_quality.py, fix_directory_orgs.py, fix_black_start_orgs.py / fix_black_start_all.py, fix_missing_data.py,
fix_remaining_issues.py, fix_remaining_orgs.py, update_addresses_and_heads.py
   │
   ▼
scripts/generate_import_readiness_report.py ──► Import_Readiness_Report.xlsx  (Critical/Warning/Informational/False-Positive)
scripts/validate_database.py / verify_data.py ──► ongoing read-only data-quality audits against the live DB
   │
   ▼
Application (app.py / routes / templates) serves the resulting PostgreSQL data
   │
   ▼
Exports (Excel/Word/PDF) — generated on demand, always from live data, closing the loop back to distributable documents
```

## 15.3 Script-by-Script Explanation

| Script | Role |
|---|---|
| `scripts/repair_workbook.py` | Repairs a raw workbook into a verified copy; every change logged to `reports/repair_change_log.csv`; source file never opened in write mode. |
| `scripts/validate_workbook.py` | Read-only pre-import gate: duplicate IDs, blank required fields, orphan org/suborg references across `organizations`/`sub_organizations`/`employees`/`control_rooms`/`switchyards` sheets. |
| `scripts/compare_word_excel.py` | Read-only: compares Word directory rosters against the Excel workbook, reports missing employees/control rooms per station (`reports/comparison_report.csv`). Home of shared Word-table-parsing primitives (`clean`, `strip_leading_number`, `iter_block_items`, header/segment detection) reused by several downstream scripts. |
| `scripts/reconcile_word_database.py` | Read-only: reconciles the **live PostgreSQL database** against the Word directory per station. A station is only "Perfect" when every Control Room/Switchyard record matches by name, normalized phone digits, and email on both sides, with no missing/extra/duplicate records — a stricter, record-level check than a simple presence/absence flag. Writes `reports/directory_reconciliation_report.csv` and `reports/perfect_stations.csv`. |
| `scripts/generate_validation_report.py` | Validates the near-final workbook against the Word source of truth, writes `Validation_Report.xlsx`. |
| `scripts/apply_corrections.py` | Applies only evidence-backed, unambiguous corrections (Word-present/Excel-blank fill-ins; single-purpose field overwrites like designation or a lone email). Explicitly **never** auto-overwrites multi-slot phone fields (Office/Residence/Mobile/Fax, Contact/OPD/Emergency) on a mismatch, because Word's own phone extraction doesn't preserve which slot a number belongs to — those go to manual verification instead. Writes `WR_DB_Ready_Final_Verified_v2.xlsx` + `Correction_Log.xlsx`. |
| `scripts/normalize_formatting.py` | Standardizes capitalization/spacing/separators/email case. By explicit design, it normalizes organization/designation/control-room/switchyard/hospital/emergency-service names, lowercases emails only, and **leaves phone numbers, IDs/FKs, department, and every person-name field untouched** (person names have exceptions — initials, honorifics, IAS-style suffixes — that a generic rule set can't safely handle). Writes `WR_DB_Ready_Final_Verified_v3.xlsx` + `Formatting_Normalization_Log.xlsx` — this is the version actually consumed by the import step. |
| `import_from_excel_db.py --apply` | The primary production import: organizations, KMP employees (`is_kmp=True`), control rooms, switchyards, from `WR_DB_Ready_Final_Verified_v3.xlsx`. Deliberately never touches the WRLDC organization or its own employees. Uses a documented priority order for resolving an ambiguous employee's organization (valid sub-org id from Excel → hand-maintained manual override map → designation text naming a known station → skip if genuinely ambiguous across multiple stations → parent-org fallback). New/re-inserted organizations default to the "Others" category. Writes an `ImportBatch` row. |
| `import_employees_excel.py --apply` | Imports WRLDC-internal staff from `uploads/employee list.xlsx`, run as a separate step; explicitly sets `is_kmp=False` for every row. Writes an `ImportBatch` row. |
| `scripts/classify_organizations.py` | One-off, **hand-reviewed** (not keyword-heuristic) mapping of every organization to one of the ten original `organization_categories`, explicitly avoiding a naive `%NTPC%`-style pattern match because it would misfile entities like "NTPC Vidyut Vyapar Nigam Limited" (a trading arm, not a generator). |
| `scripts/reclassify_organization_types.py` | Two independent, source-grounded passes: **State** — pulled from the verified workbook's `state` column where present (~44% coverage, never invented for the rest); **Type** — remapped to the 12-type taxonomy using structured signals (org name, the `region` grouping label, and sub-organization address text) via a documented priority-ordered rule set, not a blind keyword classifier. |
| `scripts/generate_import_readiness_report.py` | Read-only: classifies every row remaining in `Correction_Log.xlsx`'s manual-verification sheet into Critical/Warning/Informational/False-Positive, using rules re-derived from how `apply_corrections.py` originally produced each row (so the classification stays correct even if the log is regenerated). Writes `Import_Readiness_Report.xlsx`. |
| `scripts/validate_database.py` | Read-only, ongoing: validates the current live database state — counts, duplicate employee IDs, duplicate phone numbers (warning only), orphan foreign keys, invalid org/suborg references. |
| `scripts/review_reconciliation.py` | **Interactive, human-in-the-loop** reconciliation tool — reads the reconciliation report, cross-references Word/DB/workbook, and presents each suggested repair with an explicit `[Y/n]` prompt; decisions are permanently recorded in `config/manual_reconciliation_decisions.csv` so already-approved/rejected items are never re-presented. This is the tool that enforced the project's "never guess, only apply what's evidence-backed or explicitly human-approved" rule during the final reconciliation phase. |
| `verify_data.py` | Read-only, repeatable data-quality report — organization/employee/directory-number checks, suspicious-pattern detection (garbled Unicode, address text stored as an org name, placeholder names). |
| Hand-curated one-off scripts (`merge_duplicate_orgs.py`, `find_duplicate_orgs.py`, `set_org_regions.py`, `fix_data_quality.py`, `fix_directory_orgs.py`, `fix_black_start_orgs.py`/`fix_black_start_all.py`, `fix_missing_data.py`, `fix_remaining_issues.py`, `fix_remaining_orgs.py`, `update_addresses_and_heads.py`) | Each targets a specific, hand-identified batch of records by exact ID/name (never a blanket heuristic), always with a dry-run default and an explicit `--apply` flag — the same conservative pattern used throughout the whole pipeline. |

## 15.4 Design Principle Running Through the Whole Pipeline

Every script defaults to a **dry run** and requires an explicit `--apply` flag to write anything; every correction step distinguishes between *evidence-backed, unambiguous* changes (auto-applied) and *ambiguous* ones (routed to manual verification, ultimately via the interactive `review_reconciliation.py` tool or the human-reviewed `Import_Readiness_Report.xlsx`); and nothing overwrites the original Word source or the original raw Excel workbook — every stage writes to a new, version-suffixed copy, leaving a full paper trail (`_v2`, `_v3`, `*_Log.xlsx`, `*_Report.xlsx`) of exactly what changed and why.

---

# 16. Export Modules

## 16.1 Excel Generation

Two independent Excel-generation code paths exist:
- **Directory Versions** (`services/excel_generator.py`) — pure function, 7 fixed sheets (Employees, Utility Heads, Administrative Heads, KMP, Control Rooms, Switchyards, Emergency Contacts), styled with a shared navy/white header convention, saved to disk as part of a version's generation (Section 3.2 `directory_versions`).
- **Ad-hoc exports** (`routes/user_routes.py`) — Employees, Telephone Directory, Utility Heads, Administrative Heads, and Email Lists/Address Book each have their own `openpyxl`-based export function/route, generated on request (not persisted to disk), always re-running that page's current filter/keyword logic so the download matches exactly what's on screen.

## 16.2 Word Generation

`python-docx`-based exports exist for: Employees (`/export/employees?format=docx`), Telephone Directory (`/export/telephone?format=docx`), Utility Heads (`/export/utility-heads?format=docx`), and single-organization "cards" (`/export/org-word/<org_name>`, including that organization's employees plus its Control Room directory numbers).

## 16.3 Verification Workbook

`/admin/export-verification-workbook` (admin-only) — `services/verification_workbook_service.py`'s `generate_verification_workbook()`. Purpose-built for the pre-LAN-deployment sanity check: it independently re-implements each public page's own default-view query (Organizations, Employees, Telephone Directory, Utility Heads, Administrative Heads, Control Rooms, Switchyards) rather than importing route code, specifically so this export can never be silently coupled to (and broken by) unrelated route changes. A "Summary" sheet with record counts is inserted as the **first** tab even though it's generated last in code. Filename: `WRLDC_Telephone_Directory_Verification.xlsx`. Every sheet has frozen header row + Excel auto-filter enabled.

## 16.4 Formatting

All generated Excel output (ad-hoc exports, Directory Version snapshots, and the Verification Workbook) shares the same visual convention: bold white header text on a navy (`#1a3a6b`) fill. The Directory Version PDF (`services/pdf_generator.py`, ReportLab Platypus) additionally zebra-stripes every data table row and uses the same navy brand color for headers/cover page — giving the web UI, the Excel exports, and the PDF exports a single consistent visual identity.

## 16.5 Data Source

**Every** export in the system — ad-hoc, Directory Version, or Verification Workbook — reads live PostgreSQL data at generation time. None of them reads from the original import workbook. This is a deliberate design choice, stated explicitly in the Verification Workbook's own code comment: exports must reflect exactly what the application currently displays, not what was originally imported, so the two can never drift apart.

---

# 17. Security

## 17.1 Authentication

Email + Werkzeug-hashed password, via Flask-Login (`routes/auth_routes.py`). No self-registration for admin accounts, no password-reset flow.

## 17.2 Authorization

Single binary role check (`current_user.role == "admin"`), applied per-route by a local `admin_required` decorator layered on Flask-Login's `@login_required`. No finer-grained permission tiers.

## 17.3 Password Hashing

Werkzeug's `generate_password_hash`/`check_password_hash` — salted, industry-standard, never plaintext at rest.

## 17.4 Admin Protection

All state-changing admin routes require `admin_required`; the public `user_routes.py` blueprint is entirely read/submit (no destructive actions) except for the three Request Update POST endpoints, which only *create* a pending row for later admin review, never mutate live data directly.

## 17.5 Audit Logging

`log_audit_event()` covers Employee delete/KMP-toggle/status-change, Organization add/edit/delete, and Email Group CRUD — capturing who (`changed_by`/`changed_by_label`), what changed (`field_name`/`old_value`/`new_value`), why (`reason`), and request metadata (IP, browser, OS, session id). As noted in Section 7.11, there is currently no UI to browse this log.

## 17.6 Delete Protection

`delete_organization()` explicitly blocks deletion if the organization still has any dependent Employees, Directory Numbers, child Organizations, or Administrative Heads — a deliberate safeguard against the database's own inconsistent `ON DELETE` behavior (some FKs `SET NULL`, some `CASCADE`, some have no rule at all — Section 3.4).

## 17.7 Environment Variables

`SECRET_KEY` and `DATABASE_URL` are both read from environment variables **with hardcoded fallbacks** (`config.py`) — including a **plaintext database password** in the fallback connection string. This is acceptable for local development but must be addressed before LAN/production deployment (Section 20).

## 17.8 Debug Mode Removal

Fixed during the production-readiness pass: `app.py`'s `if __name__ == "__main__"` block now reads `FLASK_DEBUG` from the environment and defaults to **off**, with an explicit code comment explaining that Werkzeug's interactive debugger is a remote-code-execution risk if a debug-enabled server is reachable over the network. Previously the app always ran with `debug=True`.

## 17.9 Security Improvements Made During Development

- Removed an unauthenticated `/count` debug route entirely.
- Made debug mode opt-in via environment variable instead of hardcoded on.
- Added the `delete_organization` dependency-check guard described above.
- Added audit logging to previously-silent mutating actions (employee delete, KMP toggle, status changes, organization category changes, email group CRUD).

## 17.10 Known Security Limitations (explicitly not implemented — state plainly, do not imply otherwise)

| Gap | Detail |
|---|---|
| **No CSRF protection** | Confirmed absent everywhere — no `Flask-WTF`/`CSRFProtect` in `requirements.txt`, no token rendered or validated in any form across `auth_routes.py`, `admin_routes.py`, or `user_routes.py`. All state-changing admin actions rely solely on the session cookie. |
| **No rate limiting** | No `Flask-Limiter` or equivalent. The three public, unauthenticated "Request Update" POST endpoints have no throttling against spam/abuse. |
| **Spoofable IP logging** | `_request_metadata()` trusts the `X-Forwarded-For` header directly from the client with no reverse-proxy validation — a client can set this header to an arbitrary value, which would then be recorded as the "true" IP in audit logs and update requests. |
| **Hardcoded credential fallbacks** | `SECRET_KEY` and the database URI (with plaintext password) both fall back to hardcoded values if env vars are unset. |
| **Documented default admin credentials** | `create_admin.py` provisions `admin@gmail.com` / `admin123`; this exact credential pair is also written out in plaintext inside the project's own auto-generated `generate_summary.py` Word output. Must be rotated before go-live. |
| **No audit log viewer** | The data is captured but not reviewable through the UI today. |
| **No account lockout / login throttling** | No limit on repeated failed login attempts was found in `auth_routes.py`. |

---

# 18. Validation

## 18.1 Duplicate Detection

`find_duplicate_orgs.py` (substring-based heuristic across all organization names, writing `duplicate_orgs_report.txt`) feeds the hand-curated merge list in `merge_duplicate_orgs.py`. `scripts/validate_database.py` and `verify_data.py` both separately check for duplicate employee records (by name+organization) and duplicate directory-number entries.

## 18.2 Data Validation

`scripts/validate_workbook.py` gates the source Excel workbook before it's ever imported (duplicate IDs, blank required fields, orphan org/suborg references). `verify_data.py` runs the same class of checks against the *live database* on an ongoing basis, not just at import time — organization address/region gaps, zero-employee organizations, employee contact-field coverage percentages, and suspicious-pattern detection (garbled Unicode, address text mistakenly stored as an organization name, placeholder employee names like "n/a"/"nil"/"-").

## 18.3 Control Room Validation

Handled through the record-level comparison in `scripts/reconcile_word_database.py` — a station's Control Room/Switchyard set is only marked "Perfect" when every record matches the Word source by name, normalized phone digits, and email, with zero missing/extra/duplicate entries on either side (a materially stricter bar than "does a control room exist for this station").

## 18.4 Organization Validation

`scripts/validate_workbook.py` (pre-import), `scripts/reclassify_organization_types.py` (state/type correctness, source-grounded rather than guessed), and `verify_data.py` (ongoing, live-DB) together cover organization data quality from three different angles: before import, during classification, and continuously afterward.

## 18.5 Import Validation

Every import script (`import_from_excel_db.py`, `import_employees_excel.py`) defaults to a dry run, requires `--apply` to write, wraps its delete+insert in a single transaction (rolled back whole on any failure), and writes an `ImportBatch` row recording mode/status/counts either way.

## 18.6 Verification Workbook

Covered in Section 16.3 — the admin-only, live-data cross-check export used as a final pre-deployment sanity check against the running application itself.

## 18.7 Final Reconciliation

The final verification stage (immediately preceding this documentation effort) was a full end-to-end reconciliation between the official Word directory, the live PostgreSQL database, and the generated Telephone Directory/Verification exports, governed by an explicit rule: only correct issues fully supported by the official Word directory; anything the Word document itself left inconsistent, or that couldn't be resolved with certainty, was logged as a manual-review item rather than guessed at. `scripts/review_reconciliation.py`'s interactive `[Y/n]`-approval workflow and its permanent `config/manual_reconciliation_decisions.csv` record are the concrete mechanism that enforced this rule in practice.

---

# 19. Challenges Faced

| Challenge | How it was solved |
|---|---|
| **Duplicate imports** | A pre-existing WRLDC double-import (65 pairs) plus 9 self-inflicted duplicates (created during reconciliation, when employees flagged "missing" by an automated script were re-inserted without first checking they already existed under a slightly different designation string) — together 74 duplicate pairs — were identified by exact-ID-range inspection and resolved by deleting the newer/incomplete duplicate and keeping the pre-existing, more complete record. |
| **Phantom sub-organizations** | Private TBCB transmission SPVs and similar entities that exist in the Word directory by name/address only (no employees, no control room) were initially invisible in the Telephone Directory's category view, because the grouping logic only created an entry when it encountered an employee or directory-number row for that organization. Fixed by pre-seeding the grouping dict with every organization in the category up front — restoring visibility for 55 organizations across 6 categories. |
| **Organization classification** | No file in the entire project ever contained an explicit "Organization Type" column. `scripts/reclassify_organization_types.py` had to derive Type from indirect, structured signals (organization name, the legacy `region` grouping label, and sub-organization address text) via a documented priority-ordered rule set — explicitly avoiding both a blind keyword classifier and any invented/memorized company facts. |
| **Word document inconsistencies** | The Word directory itself contains inconsistent Control Room naming, Word-numbering artifacts, OCR-style broken spacing in organization names (e.g. "N TPC" → "NTPC"), and stray non-data text (motivational quotes, headers) mixed into table cells. `import_employee.py`'s Word parser has dedicated heuristics for each: `clean_org_name()`, `is_junk_paragraph()`, `is_address_line()`, and explicit "check control room before switchyard" ordering because some cell text (e.g. "400KV GIS Control room") would otherwise match both patterns. |
| **Data reconciliation** | Reconciling three independently-evolving sources (Word, Excel workbook, live database) required building a dedicated multi-stage pipeline (Section 15) rather than a single script, specifically so that ambiguous cases could be routed to human review instead of being silently guessed. |
| **Import bugs** | Two concrete, code-comment-documented bugs from earlier import runs: (1) the workbook's own `organizations` sheet included a row for the WRLDC organization itself, which the import loop would otherwise try to re-insert, violating the unique-name constraint; (2) `control_rooms` and `switchyards` sheets have different column counts (12 vs 9), which an earlier version of the insert function didn't account for, causing a `ValueError`. Both are now explicitly guarded against in `import_from_excel_db.py`. |
| **UI consistency** | The Telephone Directory's organization list needed to behave as a true single-expand accordion (opening one card collapses any other), and needed to auto-expand automatically when a filter/search narrowed the result to exactly one organization — both implemented purely with Bootstrap's native `data-bs-parent` mechanism and a Jinja conditional, no custom JS. |
| **Control Room formatting** | Standardized via `scripts/normalize_formatting.py`'s explicit, scoped rule set (title-case + abbreviation preservation for control room/switchyard labels) — deliberately excluding phone numbers and person names from the same normalization pass, since a blanket rule would have corrupted both. |
| **Performance** | Addressed at the query level rather than via caching: outer joins to avoid N+1-style per-row lookups, batched emergency-contact pre-fetching on the Employee Directory list, and dedicated indexes added specifically for the Email Distribution "browse" view (migration 004) once it was identified that PostgreSQL doesn't automatically index foreign-key columns. |
| **Security** | Debug mode was found running with Werkzeug's interactive debugger reachable over the network (RCE risk) and was made opt-in/off-by-default; an unauthenticated `/count` debug route was removed entirely; a silent-data-loss risk in Organization deletion (inconsistent `ON DELETE` rules across FKs) was closed with an explicit dependency check. |

---

# 20. Final Production Readiness

## 20.1 Final Testing

A dedicated production-readiness review pass covered: admin pages, employee/organization workflows, search/filtering/pagination, exports, the Active/Inactive employee-status workflow, the import pipeline, and audit logging — with an explicit constraint that only fixes improving correctness, stability, maintainability, or production readiness would be made (no speculative new features, no cosmetic-only UI churn) during that phase. New self-service features (Address Book CRUD, KMP/status visibility and editing, Administrative Head request-update, the Employee status-report dropdown) were added afterward, at explicit user request, superseding that constraint for those specific items.

## 20.2 UAT

The application has reached the stage of a final, Word-directory-anchored end-to-end reconciliation (Section 18.7) intended as the last gate before UAT — the recommended UAT scope is: every organization category's Telephone Directory view, the Employee/Administrative Head Request Update flows (submission → admin queue → approval side effects), the Directory Version generate/regenerate/delete lifecycle, and every export format (Excel/Word/PDF/CSV).

## 20.3 LAN Deployment

The app currently runs via Werkzeug's built-in dev server (`app.run(host="0.0.0.0", port=8000, debug=False)`), which is explicitly **not** a production-grade WSGI server. For LAN deployment, running behind a production WSGI server (e.g. Waitress on Windows, or Gunicorn behind a reverse proxy on Linux) is recommended rather than relying on the Flask dev server directly, even for internal-LAN-only exposure.

## 20.4 Backup Strategy

**Not implemented in-application.** No automated PostgreSQL backup/restore mechanism, scheduled dump job, or `pg_dump` wrapper script exists anywhere in this codebase — backup strategy would need to be handled at the infrastructure/DBA level (e.g. scheduled `pg_dump`, PostgreSQL's own WAL archiving), outside this application's own scope.

## 20.5 Verification

`/admin/export-verification-workbook` (Section 16.3) is the built-in mechanism for a final, point-in-time cross-check of exactly what the live application shows, ahead of go-live sign-off.

## 20.6 Remaining Manual-Review Items

Per the final reconciliation rule (Section 18.7), any Word-directory inconsistency that could not be resolved with certainty was deliberately left as a documented manual-review item rather than corrected by guesswork — these are tracked in the reconciliation pipeline's own output files (`Correction_Log.xlsx`'s Needs_Manual_Verification sheet, classified by severity in `Import_Readiness_Report.xlsx`) rather than silently resolved. Combined with the security gaps listed in Section 17.10 (no CSRF, no rate limiting, hardcoded credential fallbacks, no audit-log viewer, default admin credentials needing rotation), this is the concrete punch list to close out — or explicitly accept as known limitations — before go-live.

---

# 21. Interview / GM Questions

118 likely questions, grouped by theme, each with a detailed, code-grounded answer.

## 21.1 Business Questions

**Q1. What problem does this system actually solve?**
It replaces a Word document + Excel workbook as the region's telephone directory with a single live, searchable, PostgreSQL-backed application, so contact information for WRLDC staff, every Western-Region utility's leadership, and all control rooms/switchyards is always current and instantly findable, instead of depending on someone remembering to re-edit and re-circulate a file.

**Q2. Who are the users of this system?**
Two groups: (1) anyone on the WRLDC LAN who needs to look someone up — no login required for any browsing, search, or export feature; (2) a small set of Admin users who maintain the data — organizations, employees, administrative heads, directory numbers — and review public correction requests.

**Q3. Why was a database chosen over continuing to maintain Word/Excel?**
A database gives one source of truth, structured search/filtering, referential integrity between employees/organizations/roles, a status/history mechanism, and the ability to regenerate the same familiar Word/Excel/PDF outputs on demand from live data — none of which a manually edited document can offer.

**Q4. What is the single biggest operational improvement this delivers?**
Self-service correction: any user, including the person whose details are wrong, can submit a "Request Update" that lands in an admin review queue — replacing an informal, single-point-of-failure update process with a lightweight, tracked workflow.

**Q5. How does this system stay accurate over time without becoming stale like the old Word document?**
Through the Request Update workflow (public corrections reviewed by an admin), the Employee status mechanism (so transferred/retired people are flagged rather than silently left listed), and the periodic re-run of the Word-to-database reconciliation pipeline against the official Word directory.

**Q6. Does this replace the official Word directory entirely?**
No — the Word document remains the official source of truth that the database is periodically reconciled against (Section 15/18). The application is the operational, always-current interface; Word/Excel/PDF exports generated from live data are still produced for distribution where a document is needed.

**Q7. What happens to an employee's history when they leave or transfer?**
Their `status` changes (directly by an admin, or via an approved public "Status" request), which writes a permanent `EmployeeStatusHistory` row and — if they held an Administrative Head role — automatically archives that role into `AdministrativeHeadHistory` rather than leaving a stale listing.

**Q8. How is leadership information (Chairman/MD/CEO) different from regular staff information?**
It's a separate module, "Administrative Heads," manually curated rather than auto-detected, because these are often senior/statutory appointments (including, in some cases, government officials with no WRLDC employee record at all) that need deliberate assignment, an effective-from date, and a formal "End Role" action with a documented reason — not something that should change automatically just because a designation string changes.

**Q9. What is a "Utility Head" and how is it different from an "Administrative Head"?**
A Utility Head is automatically computed per organization from the most senior-sounding designation text among that organization's active employees (Section 12) — no one manually flags it. An Administrative Head is a manually assigned, formally tracked role (Chairman/MD/CEO/etc.), independent of that computation, often reserved for statutory or the very top leadership positions.

**Q10. What is the Email Distribution Lists / Address Book feature for, in business terms?**
It lets a user quickly assemble and copy/compose an email to every contact in a category (e.g. all State SLDCs) or role (e.g. all Utility Heads) without manually collecting individual addresses — directly addressing the "who do I need to email for X" coordination need.

**Q11. Who can delete data in this system?**
Only Admin users, and even then Organization deletion is blocked outright if the organization still has dependent employees, directory numbers, child organizations, or administrative heads — the system will not let an admin accidentally orphan or cascade-delete related records without first clearing them out explicitly.

**Q12. What is the "Directory Versions" feature for?**
It produces an official, dated, unchangeable PDF+Excel snapshot of the whole directory at a point in time (e.g. monthly), for record-keeping/compliance/reference — distinct from the live, always-changing application, which keeps evolving after a version is published.

## 21.2 Technical Questions

**Q13. What is the overall architecture pattern?**
A server-rendered Flask monolith: Jinja2 templates rendered on the server for nearly every page, a thin SQLAlchemy ORM layer over PostgreSQL, a handful of `services/` modules for reusable/non-trivial business logic, and a small number of JSON endpoints for AJAX-driven widgets (autocomplete, cascading dropdowns, the Address Book grid). No SPA framework, no separate frontend build.

**Q14. Why was Flask chosen over a heavier framework like Django?**
(Not documented in the codebase — this is a judgment call, not something the source proves.) Flask's minimalism suits an application of this size well: a handful of blueprints, a modest schema, and no need for Django's built-in admin/ORM-migration/permissions machinery, which this project instead hand-rolled where actually needed (its own `admin_required` decorator, its own audit log, plain numbered SQL migrations).

**Q15. How is search implemented — is it a search engine like Elasticsearch?**
No — it's plain SQLAlchemy `ilike` (case-insensitive SQL `LIKE`) filtering combined with `or_()` across the relevant text columns. There is no full-text search engine, no trigram index, no external search service.

**Q16. How does the cascading Organization Type → State → Organization dropdown work technically?**
Two small JSON endpoints (`/api/organization-types/<id>/states`, `/api/organizations`) plus a shared JavaScript file (`org_type_cascade.js`) that fetches and repopulates the dependent dropdowns on change, hiding the State step entirely for organization types not flagged `is_state_based`.

**Q17. Is there a REST API for this application?**
Not a general-purpose one. There are seven specific `jsonify` endpoints (autocomplete suggestions, cascading-dropdown data, Address Book grid data, email-list data) that exist purely to power in-page JavaScript widgets — they are not a documented, versioned, general API surface for external consumption.

**Q18. How is the frontend built — React, Vue, something else?**
None of those. It's server-rendered Jinja2 HTML with Bootstrap for layout/components and four small, dependency-free vanilla JavaScript files (no bundler, no npm build step) that progressively enhance specific widgets (autocomplete, cascading dropdowns, the Address Book grid, email copy/compose actions).

**Q19. How does the single-expand accordion on the Telephone Directory work?**
Purely via Bootstrap's native collapse component — every organization card shares one `data-bs-parent` container, so Bootstrap automatically collapses whichever card was previously open when a new one is expanded. No custom JavaScript was needed for this behavior.

**Q20. How are Excel/Word/PDF files actually generated — are they templates or generated programmatically?**
Fully programmatic: `openpyxl` builds Excel workbooks cell-by-cell (with shared styling helpers for headers/fonts/fills), `python-docx` builds Word documents the same way, and `reportlab`'s Platypus layout engine builds PDFs (including a genuine two-pass, auto-numbered Table of Contents). None of these read from a static template file.

**Q21. How does the system decide who the "Utility Head" of an organization is?**
Purely from designation text, via a 30-tier seniority ranking table with whole-word regex matching and a "longest matching phrase wins" tie-break rule — fully explained in Section 12. It is recomputed fresh on every page load, never cached.

**Q22. What happens if an organization has two employees with identical, highest-ranked designations (e.g. two "General Manager"s)?**
The algorithm picks the first one encountered in the query's ordering (by employee `id`) — ties are broken by first-occurrence order, since the comparison in `resolve_utility_head` is a strict `<`, not `<=`.

**Q23. How is the public "Request Update" workflow implemented end to end?**
A public, unauthenticated form posts to one of three routes (Employee/Directory Number/Administrative Head), creating a `Pending` `UpdateRequest` row with the requester's stated info plus captured IP/browser/OS. An admin reviews it in a single unified queue (`/admin/requests`) and either Approves (which applies the change via type-specific routing logic — including, for a Status change, the exact same helper function the admin's own status dropdown uses) or Rejects it.

**Q24. Why does the Employee "Status" field use a dropdown in the public request form but other fields use free text?**
Because Status is a closed set of valid values (`ACTIVE`, `TRANSFERRED`, `RETIRED`, etc.) enforced by a database CHECK constraint — allowing free text there could produce a request an admin can't legally approve as-is. Every other public-request field (phone, email, address, designation) has no such constraint, so free text is appropriate.

**Q25. How does an Administrative Head's contact information stay in sync with their Employee record?**
Through computed `resolved_*` properties on the `AdministrativeHead` model — when a head is linked to an `Employee` (`employee_id` set), the standalone contact fields on the head row itself are always forced to `None` and every read (`resolved_name`, `resolved_office_phone`, etc.) transparently falls through to the linked employee's live data instead, so there is only ever one place a linked head's contact info can be edited.

**Q26. What happens to a head's PA/PS assistants when their role ends?**
They are deleted along with the head row (a `CASCADE` foreign key), and are deliberately **not** copied into the historical archive — the data model treats assistants as belonging to the live assignment, not a point-in-time record of who held the role.

**Q27. How are duplicate contacts avoided in Email Distribution Lists (e.g. someone who is both a Utility Head and an Administrative Head)?**
`services/email_distribution_service.py` runs a `dedupe_and_sort()` pass over resolved contact rows before counting or rendering, specifically to collapse exactly that kind of overlap.

## 21.3 Database Questions

**Q28. What database engine is used, and why?**
PostgreSQL, accessed via `psycopg[binary]` (v3) and Flask-SQLAlchemy. (Note: the project's own `README.md` is stale and still describes a MySQL setup — the actual, current configuration in `config.py` is PostgreSQL; this should be corrected in the README as a housekeeping item.)

**Q29. How many tables are in the schema, and what are the core ones?**
18 tables. The core operational ones are `employees`, `organizations`, `directory_numbers`, `administrative_heads`, and `update_requests`; the rest are lookups (`departments`, `organization_categories`, `service_types`), audit/history (`audit_logs`, `employee_status_history`, `administrative_head_history`, `import_batches`), and feature-specific (`emergency_contacts`, `administrative_head_assistants`, `directory_versions`, `email_groups`/`group_members`/`email_group_filters`).

**Q30. Is the schema managed by an ORM migration tool like Alembic?**
No — migrations are plain, hand-written, numbered `.sql` files (`migrations/002_*.sql` through `013_*.sql`) applied directly via `psql`. There is no Alembic/Flask-Migrate version tracking table.

**Q31. Where did the base schema (before migration 002) come from?**
It predates the numbered migration system entirely — `users`, `employees`, `organizations`, `departments`, `directory_numbers`, `update_requests`, and `emergency_contacts` have no captured SQL file in this repository; they were the application's original schema before the enterprise-feature migrations began.

**Q32. How does an Organization's hierarchy (parent/sub-organization) work at the database level?**
A single self-referential foreign key, `organizations.parent_id` (`ON DELETE SET NULL`), giving each organization at most one parent. There's no separate "levels" table — hierarchy depth is whatever `parent_id` chains happen to form in the data.

**Q33. How is "who is the current Administrative Head" enforced at the database level, not just in application code?**
A partial unique index: `uq_active_admin_head` on `(organization_id, role_title)` where `role_category = 'ADMINISTRATIVE_HEAD'` — the database itself will reject a second concurrently-active Administrative Head row for the same organization+title, while KMP-category rows are exempt from this constraint entirely.

**Q34. Are employee statuses enforced at the database level or only in Python?**
Both — `Employee.status` has a SQL `CHECK` constraint (`chk_employee_status`) mirroring the Python-side `EMPLOYEE_STATUSES` list, so an invalid status value is rejected by PostgreSQL even if application code had a bug that tried to set one.

**Q35. Are all the enum-like fields in this schema backed by real database constraints?**
No, not uniformly — several are enforced only by convention in Python with no matching SQL CHECK, including `User.role`/`User.status`, `UpdateRequest.status`/`request_type`, and `AuditLog.action`. This is a documented inconsistency worth being aware of, not a defect that's been hidden.

**Q36. Is there a mismatch anywhere between two enum-like constraints that are supposed to represent the same thing?**
Yes, one concrete example: `employees.status`'s CHECK was widened in migration 007 to add `CONTRACT_ENDED` and `DECEASED` (8 values total), but `email_group_filters.status_value`'s own CHECK constraint was never updated to match — it still only allows the original 6 values, so a dynamic email group cannot currently be filtered on those two newer statuses.

**Q37. How are foreign keys handled when a parent record is deleted — is it consistent across the schema?**
No, deliberately not — FKs added in the earliest enterprise migrations that point at a true "owning" parent use `CASCADE` (e.g. `administrative_heads` → `organizations`/`employees`, `group_members` → `email_groups`/`employees`), while FKs to `users.id` (audit/attribution columns like `created_by`, `changed_by`) have no `ON DELETE` clause at all (defaulting to `NO ACTION`), and several base-schema FKs predating the migration system also have no rule specified. This asymmetry is exactly why `delete_organization()` needed an explicit application-level dependency check (Section 17.6).

**Q38. How does the "resolved_name"/"resolved_office_phone" pattern work, technically?**
They're plain Python `@property` computed attributes on the SQLAlchemy model (`AdministrativeHead`, `AdministrativeHeadAssistant`, `AdministrativeHeadHistory`) — not database columns, not database views. Each one checks whether `employee_id` is set and the related `Employee` is loaded; if so it reads that employee's live field, otherwise it falls back to a same-named column stored directly on the head/assistant row itself.

**Q39. How is the Administrative Head's tenure history preserved when a role ends?**
`head_service.end_role()` copies essentially every field off the live `AdministrativeHead` row (organization, employee, role, dates, standalone contact fields, status) into a new `AdministrativeHeadHistory` row, adds an `effective_to` date and reason, writes an audit log entry, and only then deletes the live row.

**Q40. What is `EmployeeStatusHistory` actually used for, despite its name suggesting only status changes?**
It's the general per-employee change log driving the Archive/profile timeline — a row is written whenever status, organization, **or** designation changes, with only the relevant old/new column pair populated each time (the other pair stays NULL). One reused table, three different kinds of change.

**Q41. How does a DYNAMIC Email Group's membership stay current without manual maintenance?**
It's never materialized or snapshotted — `email_group_filters` rows (OR'd within the same `filter_type`, AND'd across different `filter_type`s) are re-evaluated live, every time the group is viewed or exported, directly against the current `employees`/`organizations` tables.

**Q42. Is there any data in this schema that gets duplicated/denormalized on purpose?**
Yes, deliberately, in a few narrow, documented cases: `AdministrativeHeadHistory` snapshots contact fields at the moment a role ends (so history isn't silently rewritten if the underlying employee record changes later), and `DirectoryVersion` stores only counts and file paths (not employee data) so a published snapshot's numbers don't shift after the fact even though it deliberately does **not** duplicate the underlying employee/organization rows themselves.

## 21.4 Flask Questions

**Q43. How many Flask blueprints does the application have, and how are they organized?**
Three: `auth_bp` (login/logout), `user_bp` (every public page/action), `admin_bp` (every admin-only page/action). None of them is registered with a `url_prefix` — each route's full path is written directly in its own decorator.

**Q44. How is "is this user an admin" enforced across dozens of routes without repeating the check everywhere?**
A single locally-defined `admin_required` decorator in `admin_routes.py`, layering a role check (`current_user.role == "admin"`) on top of Flask-Login's `@login_required`, applied to every route in that blueprint.

**Q45. Does every route in the application require login?**
No — the entire `user_routes.py` blueprint (Employee Directory, Telephone Directory, Utility/Administrative Heads, Email Lists, Directory Versions, Universal Search, all exports, and the three public Request Update submission forms) has **no** authentication decorator anywhere. Only `admin_routes.py`'s routes require an authenticated admin.

**Q46. How does Flask-Login know how to load a user from the session?**
Via `app.py`'s `@login_manager.user_loader` function, `load_user(user_id)`, which does `db.session.get(User, int(user_id))`.

**Q47. What happens when an unauthenticated user hits an admin-only page?**
`login_manager.unauthorized_handler` and the app's single `@app.errorhandler(401)` handler both redirect to `/admin-login`, hardcoded as a literal path rather than via `url_for`.

**Q48. Are there any Flask error handlers for 404 or 500 errors?**
No — only a 401 handler is defined in `app.py`. 404s and 500s fall through to Flask/Werkzeug's default behavior.

**Q49. How does session expiry work?**
`app.config["SESSION_PERMANENT"] = True` plus `PERMANENT_SESSION_LIFETIME = timedelta(minutes=20)` in `config.py` expires the server-side session after 20 minutes of inactivity; `base.html` additionally runs a client-side idle timer that proactively redirects to logout/session-expired after the same 20-minute window of no user interaction, rather than waiting for the user to notice a stale page.

**Q50. How does every template get access to `utility_head_ids` without every route computing it manually?**
`app.py` registers an `@app.context_processor` that injects `utility_head_ids` (from `utils.designation_rank.compute_utility_head_ids()`) into every template's rendering context automatically.

**Q51. Is Flask running in debug mode in production?**
No, and this was a deliberate production-readiness fix — `app.py` now reads `FLASK_DEBUG` from the environment and defaults to off, specifically because Werkzeug's debug-mode interactive debugger is a remote-code-execution risk if reachable over the network.

**Q52. What WSGI server does the application use in production?**
None dedicated — it currently runs on Werkzeug's built-in development server (`app.run(...)`), which is explicitly not recommended for production use even with debug off. See Section 20.3 for the recommended fix.

**Q53. Does the application use Flask's built-in `flash()` messaging?**
Yes, extensively — every create/update/delete admin action and most public-form submissions flash a success/error message, rendered as dismissible Bootstrap alerts in `base.html`.

**Q54. Are Flask blueprints tested with any automated test framework (pytest, etc.)?**
No automated test suite exists in this codebase — `test_import.py` is a two-line manual sanity check (imports `db`, prints "SUCCESS"), and `test_import_employee.py` is actually an early draft/variant of the Word-parsing import logic, not a pytest-style test. This is a documented gap, not an oversight to gloss over.

## 21.5 PostgreSQL Questions

**Q55. Why PostgreSQL specifically, and not SQLite or MySQL?**
Not explicitly documented in the codebase as a deliberate choice narrative — but PostgreSQL's support for partial unique indexes (used for the "one active Administrative Head per org+title" and "one primary assistant" constraints) and robust `CHECK` constraint support are both actively relied upon by this schema in ways a simpler engine like SQLite would handle less cleanly at scale.

**Q56. How does the schema guarantee "at most one active Administrative Head per organization+role"?**
A PostgreSQL **partial unique index** — `uq_active_admin_head ON administrative_heads(organization_id, role_title) WHERE role_category = 'ADMINISTRATIVE_HEAD'` — enforced by the database itself, not just application logic.

**Q57. Are foreign key columns automatically indexed in PostgreSQL?**
No — this is explicitly why migration 004 exists, adding indexes on `employees.organization_id` and `directory_numbers.organization_id` specifically to support the Email Distribution "browse" view once it became clear PostgreSQL doesn't auto-index FK columns the way some other engines do.

**Q58. How is connection configuration handled?**
Via `SQLALCHEMY_DATABASE_URI` in `config.py`, read from the `DATABASE_URL` environment variable with a hardcoded PostgreSQL connection-string fallback (including a plaintext password) — see Section 17.7 for the security implication.

**Q59. Does the application use raw SQL anywhere, or is it 100% ORM?**
Primarily ORM (SQLAlchemy query builder), but at least one script (`verify_data.py`) uses a raw SQL `LEFT JOIN` specifically for an orphaned-foreign-key integrity check that's more naturally expressed that way than through the ORM.

**Q60. How are database transactions handled during import?**
Each import script wraps its delete+insert sequence in a single transaction and calls `db.session.rollback()` on any exception, restoring the database to its pre-import state and logging a `FAILED` `ImportBatch` row rather than leaving a half-applied import in place.

**Q61. Is there a rollback mechanism for a completed (applied) import?**
Not really — `ImportBatch.rollback_available` exists as a column but is always `False` in current code; true rollback of an applied import is not implemented.

**Q62. How is the connection pool / session lifecycle managed?**
Standard Flask-SQLAlchemy behavior — `db.session` is scoped to the request context automatically; no custom pooling configuration was found beyond SQLAlchemy's defaults.

## 21.6 Security Questions

**Q63. Is there CSRF protection on the forms?**
No — confirmed absent everywhere in the codebase. No CSRF library is installed, and no token is rendered or validated on any form, including admin state-changing actions.

**Q64. Is there rate limiting on the public Request Update forms?**
No — nothing prevents automated or repeated submission of the three public, unauthenticated Request Update endpoints.

**Q65. How are passwords stored?**
Hashed with Werkzeug's `generate_password_hash`, verified with `check_password_hash` — never stored or compared in plaintext.

**Q66. What is the default admin login, and is it safe to leave as-is?**
`admin@gmail.com` / `admin123`, provisioned by `create_admin.py`. It is not safe to leave as-is for a production/LAN deployment — this exact credential pair is even documented in plaintext in the project's own auto-generated summary document, and must be rotated before go-live.

**Q67. Can the IP address recorded on an audit log or update request be trusted?**
Not fully — `_request_metadata()` reads the `X-Forwarded-For` header directly from the incoming request with no reverse-proxy validation, so a client could set an arbitrary value there; it should be treated as a best-effort field, not a forensically reliable one, unless a trusted reverse proxy is guaranteed to be the only thing setting that header.

**Q68. Is there any way to see a history of admin actions?**
The data exists (`audit_logs`, populated for Employee delete/KMP-toggle/status-change, Organization CRUD, and Email Group CRUD) but there is currently **no admin UI page** to browse or filter it — today that requires a direct database query.

**Q69. How is authorization enforced — is there role-based access control beyond admin/non-admin?**
No — it's a single binary check (`current_user.role == "admin"`). There's no "editor," "viewer-with-edit," or per-module permission tier.

**Q70. Are there any known vulnerabilities in file/data handling (e.g. path traversal in exports)?**
Not identified in the areas researched for this document (export routes take internal database ids/slugs, not raw user-supplied file paths); a dedicated penetration test was outside the scope of this documentation effort and is recommended before go-live regardless.

**Q71. Is user input sanitized against SQL injection?**
Yes, structurally — the application uses SQLAlchemy's parameterized query builder (`.filter()`, `.ilike()`, etc.) throughout rather than string-concatenated SQL, which is the standard, effective defense against SQL injection in this pattern.

**Q72. What happens if the SECRET_KEY environment variable is never set?**
The app silently falls back to a hardcoded value (`"telephone_directory_secret_key"`) baked into `config.py` — this weakens session-cookie signing security and must be set to a real secret before any non-local deployment.

**Q73. Is there account lockout after repeated failed login attempts?**
No such mechanism was found in `auth_routes.py` — login attempts are not throttled or counted.

**Q74. Does the application log security-relevant events like failed logins?**
Not specifically — `audit_logs` covers data mutations, not authentication events; failed login attempts are not written anywhere beyond a flashed message to the user attempting to log in.

## 21.7 Architecture Questions

**Q75. Is this a microservices architecture?**
No — it's a single Flask monolith with one PostgreSQL database. There is no service decomposition, no message queue, no separate API gateway.

**Q76. How is business logic separated from route-handling code?**
Partially: reusable/non-trivial logic lives in `services/` (audit logging, Administrative Head lifecycle, Email Distribution resolution, Directory Version generation, export generation, org stats), while simpler, route-specific logic (validation, straightforward CRUD) stays inline in `routes/admin_routes.py`/`user_routes.py`. This is a pragmatic split, not a strict layered-architecture rule applied everywhere.

**Q77. Why do `directory_version_service.py` and `verification_workbook_service.py` each re-implement the Telephone Directory's grouping logic instead of importing it from `routes/user_routes.py`?**
Deliberately, per their own code comments — to avoid coupling an export/reporting feature to a live, frequently-edited route, which would put the export at risk of silent regression any time that route's logic changes for unrelated reasons.

**Q78. How does the application avoid tightly coupling templates to database column names?**
It doesn't fully avoid this — many templates read model attributes and `resolved_*` properties directly (e.g. `employee.status`, `head.resolved_name`). This is a typical, acceptable pattern for a server-rendered Flask app of this size, though it does mean a model rename requires a template audit.

**Q79. Is there a caching layer (Redis, etc.)?**
No caching layer exists anywhere in the codebase. Every page, including computed values like Utility Head resolution and Organization statistics, is recomputed from the database on every request — explicitly by design for `org_stats.py` ("would need a refresh job and can go stale at this data scale") though this could become a performance consideration at much larger data volumes (Section 23).

**Q80. How extensible is the Organization Category taxonomy if new types are needed later?**
Straightforward — `organization_categories` is just a table with an `is_state_based` flag; adding a new category is an `INSERT` (as migrations 009–012 repeatedly did), and existing organizations can be reclassified afterward via `scripts/reclassify_organization_types.py` or the admin Edit Organization form, without any schema change.

**Q81. What is the role of `models/__init__.py`?**
It defines the single shared `SQLAlchemy()` instance (`db`) that every other model file imports — the one true source of the ORM's metadata registry, initialized against the Flask app in `app.py` via `db.init_app(app)`.

**Q82. How does the "resolved_*" property pattern relate to the overall architecture philosophy of this project?**
It reflects a broader pattern seen throughout: prefer a single source of truth with a read-time fallback, rather than duplicating/syncing data across two places — the same philosophy shows up in how DYNAMIC email groups are resolved live instead of snapshotted, and in why exports always read live data instead of the original import file.

**Q83. Is the codebase organized by feature (e.g. one folder per module) or by technical layer?**
By technical layer — `models/`, `routes/`, `services/`, `templates/`, `static/js/`, `utils/`, `migrations/` — not by feature folders. A given feature like "Administrative Heads" therefore has its pieces spread across `models/administrative_head*.py`, parts of `routes/admin_routes.py` and `routes/user_routes.py`, `services/head_service.py`, and several `templates/*administrative_head*.html` files.

**Q84. How many lines of route code exist, roughly, and what does that suggest?**
`admin_routes.py` is roughly 1,427 lines and `user_routes.py` roughly 1,616 lines — both quite large single files. This is a reasonable structure at the project's current size, but if the application grows further, splitting each into feature-specific route modules (employees, organizations, administrative_heads, etc.) would improve maintainability.

## 21.8 Design Questions

**Q85. Why is Utility Head auto-detected but Administrative Head manually assigned?**
Because a Utility Head is meant to represent "whoever currently holds the most senior operational designation in this organization" — something that should track staff changes automatically — while an Administrative Head is often a specific, sometimes statutory or externally-appointed role (Chairman/MD/CEO/IAS officer) that an admin needs to deliberately confirm and formally end, not something that should silently change because someone's designation text happened to shift.

**Q86. Why does an Administrative Head support being "standalone" (not linked to an Employee)?**
Because some Administrative Heads are government-appointed officials (e.g. a Principal Secretary or District Collector) with no corresponding record in this application's own Employee Directory at all — the standalone name/contact fields exist specifically to represent that case without forcing a fake Employee record to be created just to hang a role off of.

**Q87. Why is KMP (`is_kmp`) a separate flag rather than a separate table from Employee?**
Because a KMP is still fundamentally an employee-shaped record (name, designation, organization, contact fields) — the flag exists to control *where* they're shown (Telephone Directory, not the WRLDC-internal Employee Directory) and *what* related data applies to them (no Emergency Contact section, since that's a WRLDC-staff-specific concept), without duplicating the whole employee schema into a second table.

**Q88. Why does the public Request Update form show a read-only "Current Information" field instead of just letting people type the new value blind?**
So the person submitting a correction (and the admin reviewing it later) can see exactly what's currently on file before proposing a change — reducing accidental "corrections" that are actually already correct, and giving the admin immediate old-vs-new context without needing to look the record up separately.

**Q89. Why does the Employee Request Update form use a dropdown for Status but the other two Request Update forms (Directory Number, Administrative Head) don't have any equivalent dynamic-input behavior?**
Because Status is the only field, across all three request types, backed by a closed enum with a database CHECK constraint — Directory Number and Administrative Head fields (phone, email, name, address, role title) are all free text with no such constraint, so there's no equivalent value set to constrain the input to.

**Q90. Why was Directory Versions built as "insert a new row, never mutate an old one" instead of updating in place?**
So a previously published/distributed snapshot can never silently change out from under anyone who downloaded it — Regenerate always produces a new, distinctly-numbered version (e.g. `2026.08.2`) with its own new files, leaving the original version and its files completely untouched.

**Q91. Why does `AdministrativeHeadHistory` snapshot contact fields instead of just keeping a foreign key back to the Employee?**
Because contact details can keep changing on the live Employee record after a role has ended — a pure foreign-key reference would make the historical record silently drift to show the person's *current* details rather than what was actually true while they held that role. Snapshotting fields at the moment of archival preserves the historically accurate picture.

**Q92. Why is there a separate `EmployeeStatusHistory` table instead of just relying on `AuditLog` for status changes?**
`EmployeeStatusHistory` is purpose-built with typed old/new columns for status, organization, and designation specifically to drive a per-employee timeline view efficiently; `AuditLog` is a generic, cross-module catch-all not optimized for "show me this one employee's full change history" queries.

**Q93. Why does the import pipeline default every script to a dry run, requiring an explicit `--apply` flag?**
A deliberate safety convention applied consistently across every import and one-off correction script in the project — so that running a script to "see what it would do" is always the default, zero-risk behavior, and an actual write to the shared production database is always an explicit, conscious choice.

**Q94. Why wasn't a keyword/heuristic classifier used to assign Organization Type automatically?**
Because it produces confident-looking but wrong answers in exactly the cases that matter most — `scripts/classify_organizations.py`'s own comment gives the concrete example of "NTPC Vidyut Vyapar Nigam Limited" (a trading subsidiary), which a naive `%NTPC%` pattern would misfile as a generation Type alongside NTPC's actual power stations. The team chose a slower, hand-reviewed mapping instead, accepting more manual effort in exchange for correctness.

## 21.9 Performance Questions

**Q95. Does the application use caching to keep pages fast?**
No — there is no caching layer anywhere; every page, including computed values like Utility Head detection, recomputes from the database on every request.

**Q96. Is pagination used consistently across list pages?**
No — only Import History and Directory Versions paginate. Employees, Organizations, Directory Numbers, Utility Heads, Administrative Heads, and Telephone Directory results all render their full result set on one page, which is fine at current data volumes but is a scaling consideration if the dataset grows substantially.

**Q97. How does the Employee Directory avoid an N+1 query problem when showing each employee's emergency contact?**
It pre-fetches all relevant `EmergencyContact` rows in a single `IN`-filtered query up front and builds an in-memory lookup dict, rather than issuing one query per visible employee row.

**Q98. Are there database indexes tuned for the application's actual query patterns?**
Yes, in the specific places that needed it — e.g. migration 004 added indexes on `employees.organization_id`/`directory_numbers.organization_id` once the Email Distribution browse view's performance need was identified, and `idx_employees_status` was added in migration 007. This is targeted, need-driven indexing rather than a blanket index-everything approach.

**Q99. What is the single biggest performance risk if this application's dataset grows 10x?**
The unpaginated list pages (Section 21.9-Q96) and the Utility Head computation running fresh on every request (`compute_utility_head_ids()` scans every ACTIVE employee every time it's called) are the two most likely places response time would start to degrade first.

**Q100. How does the Telephone Directory keyword search avoid returning an unbounded, page-crashing result set?**
Employee matches are explicitly capped to the first 200 results (`employees[:200]`) inside the route.

**Q101. Is there any background job processing (Celery, task queues) for slow operations like PDF generation?**
No — PDF/Excel generation for Directory Versions happens synchronously within the admin's Generate/Regenerate request. For the current data volume this is acceptable, but it would become a UX concern (a long-blocking request) if the dataset grew large enough to make generation noticeably slow.

**Q102. How is the two-pass PDF Table of Contents generation handled without hurting performance too much?**
ReportLab's own `multiBuild()` mechanism does exactly two passes over the document (not more), which is the standard, bounded-cost way to resolve real page numbers for a TOC in a Platypus-based PDF — not an open-ended or iterative process.

## 21.10 Deployment Questions

**Q103. What is required to run this application from scratch?**
A PostgreSQL database reachable via `DATABASE_URL` (or the hardcoded local fallback), Python dependencies from `requirements.txt`, `init_db.py` run once to create tables, `create_admin.py` run once to provision the initial admin account, and then `python app.py` (or a production WSGI server pointed at the same Flask app).

**Q104. Is the application containerized (Docker)?**
No Dockerfile or container configuration was found in this codebase.

**Q105. What web server does this run on today, and is that suitable for LAN deployment?**
Werkzeug's built-in Flask development server. It is explicitly not a production-grade WSGI server; for LAN deployment, fronting it with a proper WSGI server (Waitress on Windows, or Gunicorn behind a reverse proxy on Linux) is recommended even for internal-only exposure.

**Q106. Does the application require HTTPS?**
Nothing in the codebase enforces or configures HTTPS/TLS — that would need to be handled at the deployment/reverse-proxy layer if required, and is worth deciding explicitly for a LAN deployment (even internal traffic carrying login credentials benefits from TLS).

**Q107. How are environment-specific settings (dev vs. LAN/production) managed?**
Through environment variables (`DATABASE_URL`, `SECRET_KEY`, `FLASK_DEBUG`) read by `config.py`/`app.py`, with hardcoded fallbacks for local development — there is no separate `Config` subclass per environment (no `DevConfig`/`ProdConfig` split).

**Q108. Where are generated files (Directory Version PDFs/Excels) stored?**
On local disk, under `storage/directory_versions/<year>/<month>/`, as configured by `DIRECTORY_VERSIONS_STORAGE` in `config.py` — not in the database and not in cloud object storage.

**Q109. What is the backup story before go-live?**
Not implemented in-application (Section 20.4) — this needs to be handled at the infrastructure level (scheduled `pg_dump`, filesystem backup of the `storage/`/`uploads/` folders) before the system is relied upon for production use.

**Q110. What's the recommended pre-go-live checklist based on everything documented here?**
Rotate the default admin credentials; set real `SECRET_KEY`/`DATABASE_URL` environment variables (don't rely on the hardcoded fallbacks); put a production WSGI server in front of the app; decide on a PostgreSQL backup strategy; and explicitly accept or address the security gaps in Section 17.10 (CSRF, rate limiting, IP-header trust, no audit-log viewer) rather than leaving them undecided.

## 21.11 Future Enhancement Questions

**Q111. What is the most valuable near-term addition to this system?**
An admin UI to browse/filter the `audit_logs` table — the data is already fully captured, so this is a comparatively small effort for a real accountability/visibility gain (see Section 23 for the full list).

**Q112. Could this system support a mobile app in the future?**
The existing JSON endpoints (autocomplete, cascading dropdowns, Address Book data) demonstrate the pattern needed, but there is no general-purpose, authenticated REST API today — building one would be the necessary first step before a native mobile client could be built.

**Q113. Could role-based permissions (beyond just admin/non-admin) be added later?**
Yes — `User.role` is already a plain string field, so introducing an intermediate role (e.g. "editor" who can update but not delete, or a module-scoped admin) is a schema-and-route-decorator change, not an architectural rewrite.

**Q114. Could automated email notifications be added (e.g. notify an admin when a new Request Update arrives)?**
Yes, and it would be a natural extension — the `UpdateRequest` creation code path already exists as a single, well-defined point where a notification (email/webhook) could be triggered; no such mechanism exists today.

**Q115. Could the audit log be extended to cover more modules than it does today?**
Yes — today it only covers Employee delete/KMP-toggle/status-change, Organization CRUD, and Email Group CRUD. Extending the same `log_audit_event()` call pattern to Administrative Head CRUD (which already has its own audit calls via `head_service.py`, so this is more complete than it first appears) and Directory Number CRUD would close the remaining gaps.

**Q116. Is there a plan for CSRF protection?**
Not implemented yet, but adding `Flask-WTF`'s CSRF protection is a well-understood, moderate-effort addition given the form-heavy nature of this application — recommended as a pre-go-live or early-post-go-live item given its absence today.

**Q117. Could the reconciliation pipeline (Word → Excel → PostgreSQL) be made self-service for non-technical staff?**
Currently it's a sequence of command-line Python scripts requiring a developer to run — wrapping the pipeline's later, safer stages (validation, report generation) behind an admin-UI "Run Reconciliation Check" button is a plausible future enhancement, though the correction-application steps deliberately involve human judgment calls that shouldn't be fully automated away.

**Q118. What would need to change to support multiple regions beyond Western Region, if WRLDC's mandate expanded?**
The schema itself doesn't hardcode "Western Region" anywhere structurally (organizations simply have a `state`/`region` field) — the main region-specific logic lives in a handful of filters (e.g. the WRLDC organization-name matching used by exports) and in the source Word document itself, both of which would need generalizing rather than the database schema needing a redesign.

---

# 22. Code Walkthrough

Concrete, step-by-step execution traces for the scenarios most likely to come up in a live walkthrough or demo.

## 22.1 User Opens the Application

1. Browser requests `/`. No auth decorator on this route (it's in `user_bp`).
2. `app.py`'s registered `@app.context_processor` runs first, computing `utility_head_ids` via `compute_utility_head_ids()` (fresh query, every request — Section 12.2).
3. The `index()` view (in `user_routes.py`) queries summary counts (employees, organizations) for the stats bar and renders `index.html`, which extends `base.html`.
4. `base.html` renders the navbar (conditionally showing Admin/Logout vs. Admin Login based on `current_user.is_authenticated and current_user.role == "admin"`), loads `autocomplete.js` and `org_type_cascade.js`, and starts the 20-minute idle-session timer.
5. `index.html`'s search box has `data-suggestions-url`/`data-suggestion-source="all"`, which `autocomplete.js` auto-binds to on `DOMContentLoaded`.

## 22.2 User Searches for an Employee

1. User types into a search box wired to `autocomplete.js`. After a 160ms debounce and ≥2 characters, it calls `fetch("/suggestions?q=...&source=...")`.
2. `search_suggestions()` (in `user_routes.py`) queries Organization/Department/Employee/DirectoryNumber (scoped by the `source` param), each capped at 12–24 rows, ranks results (startswith-match boosted, then alphabetical), returns the top 10 as JSON.
3. `autocomplete.js` renders the dropdown; clicking a plain-search suggestion fills the input with the picked value and calls `input.form.submit()`.
4. The form submits to `/search?keyword=...` (or the page's own search route, e.g. `/directory?keyword=...`), which runs `employee_search_filter(keyword)` (an `or_()` of `ilike` predicates) and re-renders the results page.

## 22.3 User Opens the Telephone Directory

1. `/telephone-directory` with no query args → `landing=True`, `_compute_telephone_directory_categories()` counts organizations/contacts per `OrganizationCategory`, `telephone_directory.html` renders category cards.
2. User clicks a category card → `/telephone-directory?category_id=N`. The route loads every `Organization` in that category, pre-seeds a `grouped` dict for **all** of them (not just ones with data — the fix described in Section 9.2), then queries KMP+ACTIVE employees and `DirectoryNumber` rows scoped to that category's organization ids, bucketing numbers into control_rooms/switchyards/other by substring match on `category`.
3. `all_employees.sort(key=lambda e: e.id not in utility_head_ids)` moves each organization's Utility Head to the front of its group (stable sort).
4. `telephone_directory.html` renders one Bootstrap accordion card per organization (`data-bs-parent` making it single-expand), auto-expanding if the category yielded exactly one organization.

## 22.4 User Opens Email Lists

1. `/email-lists` → `email_distribution_home()` iterates `STANDARD_GROUPS` (organization-category-based and role-based sections), computing counts per group and skipping empty ones.
2. User clicks an "Organization Based" card → `/email-lists/browse?category=X`. The server validates `category` against `BROWSE_CATEGORY_NAMES` and renders the shell (`email_distribution_browse.html`) — **no row data yet**.
3. `email_distribution_browse.js` runs on load, reading its configuration off `#edbApp`'s `data-*` attributes, and calls `fetch("/email-lists/browse/data.json?category=X&contact_type=&state=")`.
4. `email_distribution_browse_data_json()` calls `resolve_category_contacts(category, contact_type, state)` (in `services/email_distribution_service.py`), which builds contact rows via `to_contact_row()` and dedupes/sorts them, returning JSON.
5. The JS renders the grid client-side; State/Contact Type/search changes re-fetch or re-filter without a full page reload.

## 22.5 Admin Edits an Employee

1. Admin (already authenticated, `role == "admin"`) navigates to `/admin/employees/edit/<id>` — `admin_required` verifies both `@login_required` and the role check before the view runs.
2. GET renders `edit_employee.html` pre-populated, with the org-cascade widget restoring the employee's current organization selection via `data-selected-org-id` (Section 8.5).
3. On POST: server-side validation (required fields, org-type mismatch guard, email format, mobile character whitelist, duplicate-name-in-same-org check excluding the record being edited).
4. If organization or designation changed, an `EmployeeStatusHistory` row is written (feeding the Archive timeline) even though this isn't strictly a "status" change — the model comment explains this table is the general change log, not status-only.
5. `db.session.commit()`; flash success; redirect to `manage_employees`.

## 22.6 User Submits an Update Request

1. Anonymous user opens `/employee/request-update/<id>` (no login). The route computes `current_values` (a dict of every editable field's present value, including `status`) and renders `update_request.html` with that dict serialized into a `<script type="application/json">` tag.
2. Inline JS reads that JSON; on selecting a field from the dropdown, it fills the read-only "Current Information" box from `current_values[field]`, and — only for `field == "status"` — swaps the free-text input for a `<select>` populated from a second JSON payload (`employee_statuses`).
3. On submit, `request_update()` builds an `UpdateRequest(request_type="employee", field_name=..., new_value=..., status="Pending", ...)`, capturing `ip_address`/`browser`/`operating_system` via `_request_metadata()`, and commits.
4. Later, an admin opens `/admin/requests`, sees the new Pending row (with `FIELD_LABELS` translating `field_name` to a friendly label), and clicks Approve.
5. `approve_request()` branches on `request_type == "employee"`. If `field_name == "status"`, it calls `_apply_employee_status_change(employee, new_value, reason=...)` — the exact same function the admin's own status dropdown uses — which updates `Employee.status`, writes `EmployeeStatusHistory`, calls `log_audit_event`, and auto-ends any Administrative Head role that employee currently holds if the new status isn't `ACTIVE`. For any other field, a `field_map` resolves the target attribute and applies it via `setattr`.
6. `update_request.status = "Approved"`, `reviewed_by`/`reviewed_at` set, commit.

## 22.7 Data Is Imported

1. Operator runs `import_from_excel_db.py --apply` (or without `--apply` first, for a dry-run preview) against `WR_DB_Ready_Final_Verified_v3.xlsx` — itself the end product of the multi-stage reconciliation pipeline in Section 15.
2. The script opens a single transaction, seeds `name_to_org` with the existing WRLDC organization (to avoid a duplicate-name insert), then walks the `organizations`, `sub_organizations`, `employees`, `control_rooms`, and `switchyards` sheets, resolving each employee's organization via the documented priority order (valid sub-org id → manual override map → designation-text station match → skip if ambiguous → parent-org fallback).
3. New/re-inserted organizations default to the "Others" category.
4. On success: commit, write an `ImportBatch(mode="APPLIED", status="SUCCESS", inserted_count=..., ...)` row. On any exception: `db.session.rollback()`, write an `ImportBatch(status="FAILED", ...)` row, re-raise.
5. Operator separately runs `scripts/reclassify_organization_types.py` to backfill accurate `category_id`/`state` values (Section 15), since the import step alone can't know them.

## 22.8 Excel Is Exported

1. User clicks "Export Excel" on, e.g., the Telephone Directory page (with whatever `keyword`/`category_id` filter is currently active).
2. `/export/telephone?format=xlsx&keyword=...` **re-runs the exact same grouping/filter logic** as the on-screen `/telephone-directory` route (not a cached copy of what was displayed), guaranteeing the download matches the current live data.
3. An `openpyxl.Workbook()` is built in memory, sheet(s) styled with the shared navy-header convention, and returned as a `Response`/`send_file` with the appropriate `Content-Disposition` for download — no temporary file is written to disk for this ad-hoc export path (contrast with Directory Versions, Section 16.1, which does persist its generated files).

---

# 23. Future Enhancements

Concrete, code-grounded suggestions for Version 2.0 — each tied to a specific gap identified elsewhere in this document, not a generic wishlist.

| Priority | Enhancement | Why (tied to this document) |
|---|---|---|
| High | **Audit Log viewer** | The full write path already exists (`log_audit_event`, `audit_logs` table) — only a read-only, filterable admin page is missing (Section 7.11/21.11-Q111). Comparatively low effort for a real accountability gain. |
| High | **CSRF protection** (`Flask-WTF`) | Confirmed absent across every form in the application (Section 17.10). |
| High | **Rate limiting** on public Request Update endpoints | No throttling exists today on the only unauthenticated write paths in the app (Section 17.10). |
| High | **Rotate default admin credentials + externalize secrets** | `admin@gmail.com`/`admin123` and hardcoded `SECRET_KEY`/DB-password fallbacks are real, documented risks (Section 17.7/17.9). |
| High | **Production WSGI server** (Waitress/Gunicorn + reverse proxy) | The app currently runs on Werkzeug's dev server, unsuitable for real deployment even on a LAN (Section 20.3). |
| Medium | **User/account management UI** | No admin screen exists to create/disable/list login accounts today — everything goes through `create_admin.py` (Section 7.5). |
| Medium | **Pagination on remaining list pages** | Employees, Organizations, Directory Numbers, Utility/Administrative Heads, and Telephone Directory results are all unpaginated (Section 21.9-Q96) — fine today, a scaling risk later. |
| Medium | **Notifications on new Update Requests** | Admins currently have to check `/admin/requests` manually; an email/webhook trigger at request-creation time would close the loop faster (Section 21.11-Q114). |
| Medium | **Automated database backup strategy** | Nothing in-application handles this today (Section 20.4) — needs an infrastructure-level scheduled `pg_dump` or equivalent regardless of any application change. |
| Medium | **Extend `email_group_filters.status_value` CHECK** | Currently stuck at 6 of the 8 employee statuses (Section 21.3-Q36) — a small migration to add `CONTRACT_ENDED`/`DECEASED` would let dynamic email groups filter on the full status set. |
| Low | **General-purpose authenticated REST API** | Would be the necessary foundation for any future mobile client (Section 21.11-Q112); today's JSON endpoints are narrow, page-specific widgets, not a general API. |
| Low | **Role tiers beyond admin/non-admin** | `User.role` is already a plain string, so an intermediate "editor" role is a small, additive change if the organization later wants finer-grained permissions (Section 21.11-Q113). |
| Low | **Self-service reconciliation pipeline UI** | Wrapping the later, safer stages of the Word/Excel/DB reconciliation pipeline (validation, report generation) behind an admin button, while deliberately keeping human-judgment correction steps manual (Section 21.11-Q117). |
| Low | **README correction** | `README.md` still documents a MySQL setup; the application is actually PostgreSQL-based (Section 15.2/21.5-Q28) — a small but real housekeeping fix for onboarding accuracy. |

---

*End of document.*
