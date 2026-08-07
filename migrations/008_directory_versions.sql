-- Migration: Directory Versions module (replaces Archive).
-- Metadata-only table -- no employee data is duplicated. Each row points at
-- a PDF/Excel snapshot on disk (see services/directory_version_service.py).
-- Previously generated versions must never change: regeneration for a
-- month that already has a version inserts a NEW row with a revision-
-- suffixed version_number rather than overwriting the original.
--
-- Safe to run multiple times: every statement is guarded.

CREATE TABLE IF NOT EXISTS directory_versions (
    id SERIAL PRIMARY KEY,
    version_number  VARCHAR(20) NOT NULL UNIQUE,
    version_name    VARCHAR(255) NOT NULL,
    month           INTEGER NOT NULL CHECK (month BETWEEN 1 AND 12),
    year            INTEGER NOT NULL CHECK (year BETWEEN 2000 AND 2100),
    generated_on    TIMESTAMP DEFAULT now(),
    generated_by    INTEGER REFERENCES users(id),
    employee_count      INTEGER NOT NULL DEFAULT 0,
    organization_count  INTEGER NOT NULL DEFAULT 0,
    pdf_filename    VARCHAR(255),
    pdf_path        VARCHAR(500),
    excel_filename  VARCHAR(255),
    excel_path      VARCHAR(500),
    status          VARCHAR(20) NOT NULL DEFAULT 'Published' CHECK (status IN ('Published')),
    remarks         TEXT,
    created_at      TIMESTAMP DEFAULT now(),
    updated_at      TIMESTAMP DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_directory_versions_year_month ON directory_versions(year, month);
