-- Migration: supporting indexes for the Email Distribution Lists "browse"
-- view (Organization Category + Contact Type address book), which filters
-- Employee/DirectoryNumber rows by organization_id more heavily than
-- before. Postgres does not auto-index FK columns (only the referenced PK
-- side), so these are added explicitly. Safe to run multiple times.

CREATE INDEX IF NOT EXISTS idx_employees_organization_id ON employees(organization_id);
CREATE INDEX IF NOT EXISTS idx_directory_numbers_organization_id ON directory_numbers(organization_id);
