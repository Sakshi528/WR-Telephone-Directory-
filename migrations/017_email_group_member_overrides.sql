-- Migration: Manual member overrides for DYNAMIC email groups.
-- Lets an admin always-include or always-exclude a specific Employee on top
-- of a group's filter-computed membership (e.g. someone who doesn't match
-- the category/role/status filters but should still get the mail, or a
-- filter-match who shouldn't). Reuses email_group_filters with two new
-- filter_type values instead of a separate table, since these rows are the
-- same "one criterion of a group" shape as the existing filters.
-- Safe to run multiple times.

ALTER TABLE email_group_filters
    ADD COLUMN IF NOT EXISTS employee_id INTEGER
        REFERENCES employees(id) ON DELETE CASCADE;

DO $$
BEGIN
    IF EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'email_group_filters_filter_type_check'
    ) THEN
        ALTER TABLE email_group_filters DROP CONSTRAINT email_group_filters_filter_type_check;
    END IF;
    ALTER TABLE email_group_filters
        ADD CONSTRAINT email_group_filters_filter_type_check CHECK (filter_type IN
            ('ORGANIZATION', 'ORGANIZATION_CATEGORY', 'ROLE', 'STATUS',
             'INCLUDE_EMPLOYEE', 'EXCLUDE_EMPLOYEE'));
END $$;

CREATE INDEX IF NOT EXISTS idx_email_group_filters_employee ON email_group_filters(employee_id);
