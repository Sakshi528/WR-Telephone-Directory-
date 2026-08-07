-- Migration: Archive module.
-- No new "archive" tables -- employees.status is the sole membership test
-- (status != 'ACTIVE' = archived). This migration only: (1) widens the
-- employee status vocabulary, (2) extends employee_status_history to also
-- capture organization/designation changes so it can drive the Archive
-- profile's full timeline, (3) extends audit_logs and update_requests with
-- IP/browser/OS/session fields, (4) adds the one genuinely new table,
-- import_batches, since no import tracking has ever existed.
--
-- Safe to run multiple times: every statement is guarded.

-- 1. Widen the employee status vocabulary (CONTRACT_ENDED, DECEASED)
DO $$
BEGIN
    IF EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'chk_employee_status'
    ) THEN
        ALTER TABLE employees DROP CONSTRAINT chk_employee_status;
    END IF;
    ALTER TABLE employees
        ADD CONSTRAINT chk_employee_status CHECK (status IN
            ('ACTIVE', 'INACTIVE', 'TRANSFERRED', 'RETIRED', 'DEPUTATION',
             'RESIGNED', 'CONTRACT_ENDED', 'DECEASED'));
END $$;

CREATE INDEX IF NOT EXISTS idx_employees_status ON employees(status);

-- 1b. administrative_head_history never captured service_type_id (migration
--     005 only added it to the live administrative_heads table) -- without
--     it, the Archive's "Former Administrative Heads" view can't show
--     Service Type at all. Added here since this table is already being
--     extended for the same reason as everything else in this migration.
ALTER TABLE administrative_head_history
    ADD COLUMN IF NOT EXISTS service_type_id INTEGER REFERENCES service_types(id);

-- 2. employee_status_history: also capture organization/designation deltas,
--    so this one reused table drives the Archive profile's Status Timeline,
--    Organization History, and Designation History sections together.
ALTER TABLE employee_status_history
    ADD COLUMN IF NOT EXISTS old_organization_id INTEGER REFERENCES organizations(id),
    ADD COLUMN IF NOT EXISTS new_organization_id INTEGER REFERENCES organizations(id),
    ADD COLUMN IF NOT EXISTS old_designation VARCHAR(255),
    ADD COLUMN IF NOT EXISTS new_designation VARCHAR(255);

-- 3. audit_logs: who/where/how, for every Archive action (and everything
--    else that already calls log_audit_event)
ALTER TABLE audit_logs
    ADD COLUMN IF NOT EXISTS ip_address VARCHAR(64),
    ADD COLUMN IF NOT EXISTS browser VARCHAR(100),
    ADD COLUMN IF NOT EXISTS operating_system VARCHAR(100),
    ADD COLUMN IF NOT EXISTS session_id VARCHAR(64);

-- 4. update_requests: same request-metadata capture, at submission time
ALTER TABLE update_requests
    ADD COLUMN IF NOT EXISTS ip_address VARCHAR(64),
    ADD COLUMN IF NOT EXISTS browser VARCHAR(100),
    ADD COLUMN IF NOT EXISTS operating_system VARCHAR(100);

-- 5. import_batches -- the first-ever persisted record of an import run.
--    Written going forward by import_from_excel_db.py / import_employees_excel.py;
--    historical past imports can't be retroactively populated (no data exists
--    for them). rollback_available is always FALSE for now -- true rollback
--    (undoing an already-committed import) isn't implemented.
CREATE TABLE IF NOT EXISTS import_batches (
    id SERIAL PRIMARY KEY,
    workbook_name VARCHAR(255) NOT NULL,
    imported_by INTEGER REFERENCES users(id),
    import_date TIMESTAMP DEFAULT now(),
    mode VARCHAR(20) NOT NULL CHECK (mode IN ('DRY_RUN', 'APPLIED')),
    inserted_count INTEGER NOT NULL DEFAULT 0,
    updated_count INTEGER NOT NULL DEFAULT 0,
    skipped_count INTEGER NOT NULL DEFAULT 0,
    status VARCHAR(20) NOT NULL CHECK (status IN ('SUCCESS', 'FAILED')),
    rollback_available BOOLEAN NOT NULL DEFAULT FALSE,
    summary TEXT,
    created_at TIMESTAMP DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_import_batches_import_date ON import_batches(import_date);
