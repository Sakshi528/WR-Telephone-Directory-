-- Migration: adds "Generation Company" as a 12th Organization Type -- a
-- catch-all for confirmed power-generation companies/stations whose
-- specific fuel type (Thermal/Hydel/Nuclear) or ownership (CPSU/IPP) isn't
-- confidently determinable from the uploaded source files. Not
-- state-based, same as Thermal/Hydel/Nuclear/CPSU/IPP/Others.
--
-- Safe to run multiple times.

INSERT INTO organization_categories (category_name, description) VALUES
    ('Generation Company', NULL)
ON CONFLICT (category_name) DO NOTHING;
