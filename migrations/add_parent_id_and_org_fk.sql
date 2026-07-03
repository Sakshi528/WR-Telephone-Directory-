-- Migration: add parent_id to organizations, organization_id to directory_numbers
-- Safe to run multiple times: every statement is guarded with IF NOT EXISTS.
-- No existing columns, data, or constraints are modified.

-- 1. Self-referencing FK on organizations
--    parent_id IS NULL  → top-level (parent) organization
--    parent_id = N      → sub-organization whose parent has id = N
ALTER TABLE organizations
    ADD COLUMN IF NOT EXISTS parent_id INTEGER
        REFERENCES organizations(id) ON DELETE SET NULL;

-- 2. FK from directory_numbers to organizations
--    The existing "organization" TEXT column is kept completely unchanged.
--    organization_id is only populated after the next importer run.
ALTER TABLE directory_numbers
    ADD COLUMN IF NOT EXISTS organization_id INTEGER
        REFERENCES organizations(id) ON DELETE SET NULL;

-- 3. Index for fast child-org lookup  (filter_by(parent_id=...))
CREATE INDEX IF NOT EXISTS idx_organizations_parent_id
    ON organizations(parent_id);

-- 4. Index for fast CR / switchyard lookup by org FK
CREATE INDEX IF NOT EXISTS idx_directory_numbers_org_id
    ON directory_numbers(organization_id);
