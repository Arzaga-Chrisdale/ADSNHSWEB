-- Add Track (SHS only) support to classes for SF9.
-- Safe to run more than once.

ALTER TABLE public.classes
ADD COLUMN IF NOT EXISTS track_shs TEXT;

COMMENT ON COLUMN public.classes.track_shs IS
'Senior High School track used by SF9, e.g. STEM, ABM, HUMSS, GAS, TVL.';