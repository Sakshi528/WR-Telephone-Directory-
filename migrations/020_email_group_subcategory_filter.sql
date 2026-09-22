-- Migration: adds an ORGANIZATION_SUBCATEGORY filter type for DYNAMIC email
-- groups, so a group can be built from a subcategory (e.g. Generation
-- Company -> Thermal) the same way it can already be built from a category
-- or a specific organization. Reuses email_group_filters with a new
-- subcategory_value column (free text, since Organization.region/subcategory
-- is a free-text field, not an FK) rather than a separate table.
-- Safe to run multiple times.

ALTER TABLE email_group_filters
    ADD COLUMN IF NOT EXISTS subcategory_value VARCHAR(255);

DO $$
BEGIN
    IF EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'email_group_filters_filter_type_check'
    ) THEN
        ALTER TABLE email_group_filters DROP CONSTRAINT email_group_filters_filter_type_check;
    END IF;
    ALTER TABLE email_group_filters
        ADD CONSTRAINT email_group_filters_filter_type_check
        CHECK (filter_type IN (
            'ORGANIZATION', 'ORGANIZATION_CATEGORY', 'ORGANIZATION_SUBCATEGORY',
            'ROLE', 'STATUS', 'INCLUDE_EMPLOYEE', 'EXCLUDE_EMPLOYEE'
        ));
END $$;
