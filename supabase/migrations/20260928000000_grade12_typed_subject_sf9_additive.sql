-- SIGLA Grade 12: typed subject -> Grade 12 SF9 fixed subject and term.
-- Safe to run multiple times AFTER your existing core classes table migration.
-- This is an ADD-ONLY migration, not a replacement for your full Supabase
-- schema, Grade Requests, teacher role policies or other existing functions.
--
-- The E-Class Record and SF9 Excel file still generate client-side from the
-- existing classes, grades, grade_components/activities, activity_scores tables.
-- No grade data is rewritten and no old or historical classes are deleted.
-- Grade 12 subject names must match one of the 16 subjects in the original
-- supplied SF9 workbook. The frontend keeps a PLAIN-TEXT Subject input.

begin;

-- Preserve the existing optional-Units functionality for Grade 12.
-- This is a no-op if the optional-Units migration is already installed.
-- Existing units CHECK constraints, school-year locks and Grade 11 rules
-- are deliberately NOT dropped/overwritten by this targeted migration.
alter table public.classes
  alter column units drop not null;

-- Read-only lookup of the ORIGINAL Grade 12 SF9 template's 16 row positions.
-- Uses the same case/space/ampersand normalization as the frontend.
-- The client standardizes common PE punctuation to the canonical label
-- BEFORE calling this RPC, which is why the SQL also accepts P.E. forms.
create or replace function public.grade12_sf9_subject_match(p_subject text)
returns table (
  canonical_subject text,
  assigned_term text,
  excel_row integer,
  subject_group text
)
language sql
stable
security invoker
set search_path = ''
as $$
  with sf9_subjects(canonical_subject, assigned_term, excel_row, subject_group) as (
    values
      ('Media and Information Literacy', '1', 33, 'core'),
      ('PE and Health 3', '1', 34, 'core'),
      ('Introduction to Human Philosophy', '2', 35, 'core'),
      ('Disaster Readiness and Risk Reduction', '2', 36, 'core'),
      ('Contemporary Philippine Arts', '3', 37, 'core'),
      ('PE and Health 4', '3', 38, 'core'),
      ('Filipino sa Piling Larang', '1', 40, 'applied'),
      ('Practical Research 2', '1', 41, 'applied'),
      ('General Biology 1', '1', 42, 'applied'),
      ('General Physics 1', '1', 43, 'applied'),
      ('English for Academic & Professional Purposes', '2', 44, 'applied'),
      ('Entrepreneurship', '2', 45, 'applied'),
      ('General Physics 2', '2', 46, 'applied'),
      ('General Biology 2', '3', 47, 'applied'),
      ('Inquiries, Investigation and Immersion', '3', 49, 'applied'),
      ('Capstone Project', '3', 51, 'applied')
  ),
  requested as (
    select btrim(regexp_replace(
      replace(replace(lower(btrim(coalesce(p_subject, ''))), 'p.e.', 'pe'), '&', 'and'),
      '[^a-z0-9]+', ' ', 'g'
    )) as normalized_name
  )
  select
    subject_row.canonical_subject::text,
    subject_row.assigned_term::text,
    subject_row.excel_row::integer,
    subject_row.subject_group::text
  from sf9_subjects as subject_row
  cross join requested
  where btrim(regexp_replace(
    replace(replace(lower(subject_row.canonical_subject), 'p.e.', 'pe'), '&', 'and'),
    '[^a-z0-9]+', ' ', 'g'
  )) = requested.normalized_name
  limit 1;
$$;

revoke all on function public.grade12_sf9_subject_match(text)
  from public, anon, authenticated;
grant execute on function public.grade12_sf9_subject_match(text)
  to authenticated;

commit;
notify pgrst, 'reload schema';

-- Read-only verification with THREE RESULT ROWS (no data is deleted).
-- All three test_passed values should be true in the SQL Editor output.
select
  verification.typed_subject,
  coalesce(matched.canonical_subject, 'Not matched (expected)') as mapped_subject,
  matched.assigned_term,
  matched.excel_row,
  (verification.should_match = (matched.canonical_subject is not null)) as test_passed
from (
  values
    ('PE and Health 3'::text, true),
    ('English for Academic & Professional Purposes'::text, true),
    ('Nonexistent subject'::text, false)
) as verification(typed_subject, should_match)
left join lateral public.grade12_sf9_subject_match(verification.typed_subject) as matched
  on true;
