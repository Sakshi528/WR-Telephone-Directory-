-- Migration: Administrative Heads no longer require an existing Employee row.
-- Mirrors the standalone-identity pattern already proven by
-- administrative_head_assistants (optional employee_id + free-text
-- name/office_phone/mobile/email fallback fields). Adds a `status` field
-- (ACTIVE/ON_LEAVE/VACANT/INACTIVE) so a post can be marked vacant or a
-- sitting head on leave without ending their whole tenure record. Drops the
-- pa_name/pa_designation/pa_phone/pa_email columns, superseded by the
-- administrative_head_assistants table per migration 005's own comment and
-- confirmed to have zero rows using them.
--
-- Safe to run multiple times: every statement is guarded.
-- Both administrative_heads and administrative_head_history are confirmed
-- empty (0 rows) as of this migration, so no data backfill is needed.

-- 1. administrative_heads: employee_id becomes optional
ALTER TABLE administrative_heads
    ALTER COLUMN employee_id DROP NOT NULL;

ALTER TABLE administrative_heads
    ADD COLUMN IF NOT EXISTS name VARCHAR(255),
    ADD COLUMN IF NOT EXISTS office_phone VARCHAR(50),
    ADD COLUMN IF NOT EXISTS mobile_phone VARCHAR(50),
    ADD COLUMN IF NOT EXISTS email VARCHAR(255),
    ADD COLUMN IF NOT EXISTS office_address TEXT,
    ADD COLUMN IF NOT EXISTS status VARCHAR(20) NOT NULL DEFAULT 'ACTIVE';

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'chk_administrative_head_status'
    ) THEN
        ALTER TABLE administrative_heads
            ADD CONSTRAINT chk_administrative_head_status CHECK (status IN
                ('ACTIVE', 'ON_LEAVE', 'VACANT', 'INACTIVE'));
    END IF;
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'chk_administrative_head_identity'
    ) THEN
        ALTER TABLE administrative_heads
            ADD CONSTRAINT chk_administrative_head_identity CHECK (
                employee_id IS NOT NULL OR name IS NOT NULL
            );
    END IF;
END $$;

ALTER TABLE administrative_heads
    DROP COLUMN IF EXISTS pa_name,
    DROP COLUMN IF EXISTS pa_designation,
    DROP COLUMN IF EXISTS pa_phone,
    DROP COLUMN IF EXISTS pa_email;

-- 2. administrative_head_history: same snapshot fields, no identity CHECK
--    (a history row is a point-in-time record, not a live constraint target)
ALTER TABLE administrative_head_history
    ADD COLUMN IF NOT EXISTS name VARCHAR(255),
    ADD COLUMN IF NOT EXISTS office_phone VARCHAR(50),
    ADD COLUMN IF NOT EXISTS mobile_phone VARCHAR(50),
    ADD COLUMN IF NOT EXISTS email VARCHAR(255),
    ADD COLUMN IF NOT EXISTS office_address TEXT,
    ADD COLUMN IF NOT EXISTS status VARCHAR(20);

ALTER TABLE administrative_head_history
    DROP COLUMN IF EXISTS pa_name,
    DROP COLUMN IF EXISTS pa_designation,
    DROP COLUMN IF EXISTS pa_phone,
    DROP COLUMN IF EXISTS pa_email;
