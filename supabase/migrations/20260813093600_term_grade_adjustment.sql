-- Term Grade storage for both Class Adviser and Subject Teacher E-Class Records.
-- Keeps the old adjustment column for backward compatibility and adds the
-- editable Base column used by the current E-Class Record UI.
-- Safe to run more than once.

begin;

-- Legacy adjustment column
alter table public.grades
  add column if not exists term_grade_adjustment numeric not null default 0;

alter table public.grades
  drop constraint if exists grades_term_grade_adjustment_range;

alter table public.grades
  add constraint grades_term_grade_adjustment_range
  check (
    term_grade_adjustment >= 0
    and term_grade_adjustment <= 100
  );

comment on column public.grades.term_grade_adjustment is
  'Legacy optional positive adjustment. The current E-Class Record UI uses term_grade_base instead.';

-- Current editable Term Grade Base
alter table public.grades
  add column if not exists term_grade_base numeric;

alter table public.grades
  drop constraint if exists grades_term_grade_base_range;

alter table public.grades
  add constraint grades_term_grade_base_range
  check (
    term_grade_base is null
    or (term_grade_base >= 0 and term_grade_base <= 100)
  );

comment on column public.grades.term_grade_base is
  'Optional teacher-edited Term Grade Base. NULL uses the automatically rounded Initial Grade.';

-- Existing grade RLS/policies continue to control who can update a grade row.
-- No new table is required.

commit;

notify pgrst, 'reload schema';
