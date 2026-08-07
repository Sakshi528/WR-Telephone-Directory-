-- Migration: enterprise feature expansion (Administrative Heads/KMP, Email
-- Distribution Groups, Employee status, Audit Trail, Update Request review).
-- Safe to run multiple times: every statement is guarded with IF NOT EXISTS,
-- and the CHECK constraint add is wrapped so it is skipped if already present.
-- No existing columns, data, or constraints are modified.

-- 1. Employee status (Feature 5)
ALTER TABLE employees
    ADD COLUMN IF NOT EXISTS status VARCHAR(20) NOT NULL DEFAULT 'ACTIVE',
    ADD COLUMN IF NOT EXISTS status_changed_at TIMESTAMP;

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'chk_employee_status'
    ) THEN
        ALTER TABLE employees
            ADD CONSTRAINT chk_employee_status CHECK (status IN
                ('ACTIVE', 'INACTIVE', 'TRANSFERRED', 'RETIRED', 'DEPUTATION', 'RESIGNED'));
    END IF;
END $$;

CREATE TABLE IF NOT EXISTS employee_status_history (
    id SERIAL PRIMARY KEY,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    old_status VARCHAR(20),
    new_status VARCHAR(20) NOT NULL,
    reason TEXT,
    changed_by INTEGER REFERENCES users(id),
    changed_at TIMESTAMP DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_employee_status_history_employee
    ON employee_status_history(employee_id);

-- 2. Administrative Heads + KMP, unified (Features 1, 3, 7)
--    role_category = 'ADMINISTRATIVE_HEAD' -> at most one ACTIVE row per org+role.
--    role_category = 'KMP'                 -> multiple concurrent holders allowed.
CREATE TABLE IF NOT EXISTS administrative_heads (
    id SERIAL PRIMARY KEY,
    organization_id INTEGER NOT NULL REFERENCES organizations(id) ON DELETE CASCADE,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    role_category VARCHAR(20) NOT NULL CHECK (role_category IN ('ADMINISTRATIVE_HEAD', 'KMP')),
    role_title VARCHAR(150) NOT NULL,
    effective_from DATE NOT NULL,
    remarks TEXT,
    -- Personal Assistant / office contact for this head -- free text like
    -- emergency_contacts, since a PA is often not in the employees table.
    pa_name VARCHAR(255),
    pa_designation VARCHAR(150),
    pa_phone VARCHAR(50),
    pa_email VARCHAR(255),
    created_by INTEGER REFERENCES users(id),
    created_at TIMESTAMP DEFAULT now()
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_active_admin_head
    ON administrative_heads(organization_id, role_title)
    WHERE role_category = 'ADMINISTRATIVE_HEAD';
CREATE INDEX IF NOT EXISTS idx_admin_heads_org ON administrative_heads(organization_id);

CREATE TABLE IF NOT EXISTS administrative_head_history (
    id SERIAL PRIMARY KEY,
    organization_id INTEGER REFERENCES organizations(id),
    employee_id INTEGER REFERENCES employees(id),
    role_category VARCHAR(20) NOT NULL,
    role_title VARCHAR(150) NOT NULL,
    effective_from DATE NOT NULL,
    effective_to DATE NOT NULL,
    reason TEXT,
    replacement_employee_id INTEGER REFERENCES employees(id),
    remarks TEXT,
    pa_name VARCHAR(255),
    pa_designation VARCHAR(150),
    pa_phone VARCHAR(50),
    pa_email VARCHAR(255),
    moved_at TIMESTAMP DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_admin_head_history_org
    ON administrative_head_history(organization_id);

-- 3. Email Distribution Groups (Feature 2) -- saved/named groups only;
--    ad-hoc "copy all GMs" style groups are resolved live, no table needed.
CREATE TABLE IF NOT EXISTS email_groups (
    id SERIAL PRIMARY KEY,
    name VARCHAR(150) NOT NULL UNIQUE,
    description TEXT,
    created_by INTEGER REFERENCES users(id),
    created_at TIMESTAMP DEFAULT now()
);
CREATE TABLE IF NOT EXISTS group_members (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES email_groups(id) ON DELETE CASCADE,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    added_at TIMESTAMP DEFAULT now(),
    UNIQUE(group_id, employee_id)
);

-- 4. Audit trail (Feature 8) -- append-only; no route ever issues UPDATE/DELETE
--    against this table.
CREATE TABLE IF NOT EXISTS audit_logs (
    id SERIAL PRIMARY KEY,
    module VARCHAR(50) NOT NULL,
    record_type VARCHAR(50) NOT NULL,
    record_id INTEGER,
    action VARCHAR(20) NOT NULL,
    field_name VARCHAR(100),
    old_value TEXT,
    new_value TEXT,
    changed_by INTEGER REFERENCES users(id),
    changed_by_label VARCHAR(150),
    reason TEXT,
    created_at TIMESTAMP DEFAULT now()
);
CREATE INDEX IF NOT EXISTS idx_audit_logs_module ON audit_logs(module, record_id);

-- 5. Update Request review upgrade (Feature 4)
ALTER TABLE update_requests
    ADD COLUMN IF NOT EXISTS reviewed_by INTEGER REFERENCES users(id),
    ADD COLUMN IF NOT EXISTS reviewed_at TIMESTAMP,
    ADD COLUMN IF NOT EXISTS admin_comment TEXT;
