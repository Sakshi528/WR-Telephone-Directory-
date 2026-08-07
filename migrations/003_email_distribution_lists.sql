-- Migration: Email Distribution Lists module.
-- Adds organization classification (organization_categories -> organizations)
-- and saved dynamic/custom group filters (email_group_filters, extends the
-- existing email_groups table with list_type).
-- Safe to run multiple times: every statement is guarded with IF NOT EXISTS,
-- and constraint adds are wrapped so they are skipped if already present.
-- No existing columns, data, or constraints are modified.

-- 1. Organization classification
CREATE TABLE IF NOT EXISTS organization_categories (
    id SERIAL PRIMARY KEY,
    category_name VARCHAR(100) NOT NULL UNIQUE,
    description TEXT
);

INSERT INTO organization_categories (category_name, description) VALUES
    ('Transmission Utility', NULL),
    ('Generator', NULL),
    ('RE Generator', NULL),
    ('Distribution Company', NULL),
    ('SLDC', NULL),
    ('RLDC', NULL),
    ('Central Utility', NULL),
    ('Government', NULL),
    ('Private Utility', NULL),
    ('Other', NULL)
ON CONFLICT (category_name) DO NOTHING;

ALTER TABLE organizations
    ADD COLUMN IF NOT EXISTS category_id INTEGER
        REFERENCES organization_categories(id) ON DELETE SET NULL;

CREATE INDEX IF NOT EXISTS idx_organizations_category ON organizations(category_id);

-- Backfill: run scripts/classify_organizations.py after this migration to
-- assign real categories, then apply the remainder to 'Other' below and
-- optionally tighten to NOT NULL once every row has a category:
-- UPDATE organizations SET category_id =
--     (SELECT id FROM organization_categories WHERE category_name = 'Other')
--     WHERE category_id IS NULL;
-- ALTER TABLE organizations ALTER COLUMN category_id SET NOT NULL;

-- 2. Saved dynamic/custom email groups (extends email_groups from 002)
ALTER TABLE email_groups
    ADD COLUMN IF NOT EXISTS list_type VARCHAR(10) NOT NULL DEFAULT 'STATIC';

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'chk_email_group_list_type'
    ) THEN
        ALTER TABLE email_groups
            ADD CONSTRAINT chk_email_group_list_type CHECK (list_type IN ('STATIC', 'DYNAMIC'));
    END IF;
END $$;

CREATE TABLE IF NOT EXISTS email_group_filters (
    id SERIAL PRIMARY KEY,
    group_id INTEGER NOT NULL REFERENCES email_groups(id) ON DELETE CASCADE,
    filter_type VARCHAR(30) NOT NULL CHECK (filter_type IN
        ('ORGANIZATION', 'ORGANIZATION_CATEGORY', 'ROLE', 'STATUS')),
    organization_id INTEGER REFERENCES organizations(id) ON DELETE CASCADE,
    category_id INTEGER REFERENCES organization_categories(id) ON DELETE CASCADE,
    role_value VARCHAR(30) CHECK (role_value IN
        ('UTILITY_HEAD', 'ADMINISTRATIVE_HEAD', 'KMP')),
    status_value VARCHAR(20) CHECK (status_value IN
        ('ACTIVE', 'INACTIVE', 'TRANSFERRED', 'RETIRED', 'DEPUTATION', 'RESIGNED'))
);
CREATE INDEX IF NOT EXISTS idx_email_group_filters_group ON email_group_filters(group_id);
