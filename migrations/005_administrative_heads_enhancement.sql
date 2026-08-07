-- Migration: Administrative Heads module -- Service Type classification and
-- normalized, multi-row PA/PS assistant management. Extends the
-- administrative_heads table added in 002_enterprise_features.sql; the old
-- pa_name/pa_designation/pa_phone/pa_email columns on that table are left
-- in place (unused going forward, table has zero rows so no data to
-- migrate) rather than dropped.
-- Safe to run multiple times: every statement is guarded with IF NOT EXISTS.

CREATE TABLE IF NOT EXISTS service_types (
    id SERIAL PRIMARY KEY,
    service_type_name VARCHAR(100) NOT NULL UNIQUE,
    description TEXT
);

INSERT INTO service_types (service_type_name) VALUES
    ('IAS'), ('IPS'), ('IFS'), ('State Civil Service'), ('Other')
ON CONFLICT (service_type_name) DO NOTHING;

ALTER TABLE administrative_heads
    ADD COLUMN IF NOT EXISTS service_type_id INTEGER
        REFERENCES service_types(id) ON DELETE SET NULL;

CREATE TABLE IF NOT EXISTS administrative_head_assistants (
    id SERIAL PRIMARY KEY,
    administrative_head_id INTEGER NOT NULL
        REFERENCES administrative_heads(id) ON DELETE CASCADE,
    employee_id INTEGER REFERENCES employees(id) ON DELETE SET NULL,
    designation VARCHAR(100) NOT NULL,   -- PA / PS / Executive Assistant / ...
    name VARCHAR(255),                   -- used only when employee_id IS NULL
    office_phone VARCHAR(50),
    mobile VARCHAR(50),
    email VARCHAR(255),
    is_primary BOOLEAN NOT NULL DEFAULT FALSE,
    created_by INTEGER REFERENCES users(id),
    created_at TIMESTAMP DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_admin_head_assistants_head
    ON administrative_head_assistants(administrative_head_id);

CREATE UNIQUE INDEX IF NOT EXISTS uq_primary_assistant
    ON administrative_head_assistants(administrative_head_id)
    WHERE is_primary = TRUE;
