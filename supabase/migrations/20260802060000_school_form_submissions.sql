-- ============================================================
-- SCHOOL FORM SUBMISSIONS
-- Individual form submission / review workflow
-- ============================================================


-- ============================================================
-- FORM READINESS
-- ============================================================

create table if not exists public.class_form_readiness (
  class_id uuid not null references public.classes(id) on delete cascade,
  adviser_id uuid not null references auth.users(id) on delete cascade,
  form_code text not null,
  is_ready boolean not null default false,
  updated_at timestamptz not null default now(),
  primary key (class_id, form_code)
);


-- ============================================================
-- INDIVIDUAL FORM SUBMISSIONS
-- ============================================================

create table if not exists public.school_form_submissions (
  id uuid primary key default gen_random_uuid(),

  class_id uuid not null
    references public.classes(id)
    on delete cascade,

  adviser_id uuid not null
    references auth.users(id)
    on delete cascade,

  adviser_name text,

  grade_level text not null,
  section text,
  school_year text,

  status text not null default 'submitted'
    check (
      status in (
        'submitted',
        'approved',
        'returned'
      )
    ),

  admin_remarks text,

  submitted_at timestamptz not null default now(),
  reviewed_at timestamptz,

  reviewed_by uuid
    references auth.users(id),

  snapshot jsonb not null default '{}'::jsonb
);

-- IMPORTANT FOR EXISTING DATABASES:
-- CREATE TABLE IF NOT EXISTS does not add new columns to a table that already exists.
-- Add form_code explicitly so this SQL works with your current Supabase table.
alter table public.school_form_submissions
add column if not exists form_code text;

-- Keep old package rows valid if they already exist (their form_code can stay NULL).
-- All NEW individual submissions created by submit_school_form_to_admin()
-- will always write a real form_code.
alter table public.school_form_submissions
drop constraint if exists school_form_submissions_form_code_check;

alter table public.school_form_submissions
add constraint school_form_submissions_form_code_check
check (
  form_code is null
  or form_code in (
    'GSA',
    'SF5',
    'SOG Report',
    'SF1',
    'SF9 (New)',
    'SF8',
    'SF10',
    'Anecdotal'
  )
);


-- ============================================================
-- INDEXES
-- ============================================================

-- Remove the OLD package-wide uniqueness rule.
-- It allowed only one active submission for the whole class and would block
-- SF1/SF5/SF9/etc. from being submitted independently.
drop index if exists public.one_open_school_form_submission_per_class;

-- Each form is independent.
--
-- Example:
-- Grade 9 Polaris can have:
-- SF1 submitted
-- SF5 approved
-- SF9 returned
-- SF10 not submitted
--
-- at the same time.

create unique index if not exists
one_active_school_form_submission_per_form
on public.school_form_submissions (
  class_id,
  form_code
)
where status in ('submitted', 'approved');


create index if not exists school_form_submissions_admin_sort
on public.school_form_submissions (
  grade_level,
  adviser_name,
  section,
  submitted_at desc
);


create index if not exists school_form_submissions_form_code_idx
on public.school_form_submissions (
  form_code
);


create index if not exists school_form_submissions_status_idx
on public.school_form_submissions (
  status
);


-- ============================================================
-- ROW LEVEL SECURITY
-- ============================================================

alter table public.class_form_readiness
enable row level security;


alter table public.school_form_submissions
enable row level security;


-- ============================================================
-- READINESS POLICY
-- ============================================================

drop policy if exists
"advisers read own form readiness"
on public.class_form_readiness;


create policy
"advisers read own form readiness"
on public.class_form_readiness
for select
to authenticated
using (
  adviser_id = auth.uid()
  or public.is_admin(auth.uid())
);


-- ============================================================
-- SUBMISSIONS POLICY
-- ============================================================

drop policy if exists
"advisers read own form submissions"
on public.school_form_submissions;


create policy
"advisers read own form submissions"
on public.school_form_submissions
for select
to authenticated
using (
  adviser_id = auth.uid()
  or public.is_admin(auth.uid())
);


-- ============================================================
-- ADMIN DELETE SUBMITTED FORM
-- Allows only Admin users to delete an individual submitted form
-- from public.school_form_submissions.
-- ============================================================

drop policy if exists
"admins can delete school form submissions"
on public.school_form_submissions;


create policy
"admins can delete school form submissions"
on public.school_form_submissions
for delete
to authenticated
using (
  public.is_admin(auth.uid())
);


-- ============================================================
-- MARK ONE FORM READY / NOT READY
-- ============================================================

create or replace function public.set_class_form_readiness(
  p_class_id uuid,
  p_form_code text,
  p_is_ready boolean
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare

  allowed_forms constant text[] :=
    array[
      'GSA',
      'SF5',
      'SOG Report',
      'SF1',
      'SF9 (New)',
      'SF8',
      'SF10',
      'Anecdotal'
    ];

begin

  -- Must be the Class Adviser who owns the class.
  if not exists (
    select 1
    from public.classes c

    join public.profiles p
      on p.id = auth.uid()

    where
      c.id = p_class_id
      and c.teacher_id = auth.uid()
      and p.teacher_type = 'class_adviser'
  )
  then
    raise exception
      'Only the owning Class Adviser can update form readiness.';
  end if;


  -- Validate form.
  if not (p_form_code = any(allowed_forms))
  then
    raise exception 'Unknown required form.';
  end if;


  -- IMPORTANT:
  -- Only lock THIS specific form.
  --
  -- Other forms remain editable.

  if exists (
    select 1
    from public.school_form_submissions s
    where
      s.class_id = p_class_id
      and s.form_code = p_form_code
      and s.status in ('submitted', 'approved')
  )
  then

    raise exception
      'This form is locked while under Admin review or after approval.';

  end if;


  insert into public.class_form_readiness (
    class_id,
    adviser_id,
    form_code,
    is_ready,
    updated_at
  )
  values (
    p_class_id,
    auth.uid(),
    p_form_code,
    p_is_ready,
    now()
  )

  on conflict (class_id, form_code)
  do update set

    adviser_id = excluded.adviser_id,
    is_ready = excluded.is_ready,
    updated_at = now();

end;
$$;


-- ============================================================
-- SUBMIT ONE FORM TO ADMIN
-- ============================================================

create or replace function public.submit_school_form_to_admin(
  p_class_id uuid,
  p_form_code text
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare

  v_class public.classes%rowtype;

  v_submission_id uuid;

  v_adviser_name text;

  allowed_forms constant text[] :=
    array[
      'GSA',
      'SF5',
      'SOG Report',
      'SF1',
      'SF9 (New)',
      'SF8',
      'SF10',
      'Anecdotal'
    ];

begin

  -- ==========================================================
  -- CHECK CLASS ADVISER
  -- ==========================================================

  select c.*
  into v_class

  from public.classes c

  join public.profiles p
    on p.id = auth.uid()

  where
    c.id = p_class_id
    and c.teacher_id = auth.uid()
    and p.teacher_type = 'class_adviser';


  if not found
  then
    raise exception
      'Only the owning Class Adviser can submit this form.';
  end if;


  -- ==========================================================
  -- CHECK FORM CODE
  -- ==========================================================

  if not (p_form_code = any(allowed_forms))
  then
    raise exception 'Unknown school form.';
  end if;


  -- ==========================================================
  -- FORM MUST BE READY
  -- ==========================================================

  if not exists (
    select 1

    from public.class_form_readiness r

    where
      r.class_id = p_class_id
      and r.adviser_id = auth.uid()
      and r.form_code = p_form_code
      and r.is_ready = true
  )
  then

    raise exception
      'Mark this form Ready before submitting it to Admin.';

  end if;


  -- ==========================================================
  -- PREVENT DUPLICATE ACTIVE SUBMISSION
  -- ==========================================================

  if exists (
    select 1

    from public.school_form_submissions s

    where
      s.class_id = p_class_id
      and s.form_code = p_form_code
      and s.status in ('submitted', 'approved')
  )
  then

    raise exception
      'This form is already submitted or approved.';

  end if;


  -- ==========================================================
  -- ADVISER NAME
  -- ==========================================================

  select
    coalesce(
      nullif(trim(p.full_name), ''),
      nullif(trim(p.email), ''),
      'Class Adviser'
    )

  into v_adviser_name

  from public.profiles p

  where p.id = auth.uid();


  -- ==========================================================
  -- RESUBMIT A RETURNED FORM
  -- ==========================================================

  update public.school_form_submissions

  set
    adviser_id = auth.uid(),

    adviser_name = v_adviser_name,

    grade_level = v_class.grade_level,

    section = v_class.section,

    school_year = v_class.school_year,

    status = 'submitted',

    admin_remarks = null,

    submitted_at = now(),

    reviewed_at = null,

    reviewed_by = null,

    snapshot = jsonb_build_object(
      'class',
      to_jsonb(v_class),

      'form_code',
      p_form_code,

      'captured_at',
      now()
    )

  where id = (

    select s.id

    from public.school_form_submissions s

    where
      s.class_id = p_class_id
      and s.form_code = p_form_code
      and s.status = 'returned'

    order by s.submitted_at desc

    limit 1

  )

  returning id
  into v_submission_id;


  -- ==========================================================
  -- FIRST SUBMISSION
  -- ==========================================================

  if v_submission_id is null
  then

    insert into public.school_form_submissions (

      class_id,
      form_code,

      adviser_id,
      adviser_name,

      grade_level,
      section,
      school_year,

      status,

      snapshot

    )
    values (

      p_class_id,
      p_form_code,

      auth.uid(),
      v_adviser_name,

      v_class.grade_level,
      v_class.section,
      v_class.school_year,

      'submitted',

      jsonb_build_object(
        'class',
        to_jsonb(v_class),

        'form_code',
        p_form_code,

        'captured_at',
        now()
      )

    )

    returning id
    into v_submission_id;

  end if;


  return v_submission_id;

end;
$$;


-- ============================================================
-- ADMIN REVIEW ONE FORM
-- ============================================================

create or replace function public.review_school_form_submission(
  p_submission_id uuid,
  p_status text,
  p_admin_remarks text default null
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin

  -- Admin only.
  if not public.is_admin(auth.uid())
  then
    raise exception 'Admin access required.';
  end if;


  -- Only these two Admin actions are valid.
  if p_status not in ('approved', 'returned')
  then
    raise exception
      'Status must be approved or returned.';
  end if;


  -- Returned forms should explain what needs correction.
  if
    p_status = 'returned'
    and coalesce(trim(p_admin_remarks), '') = ''
  then

    raise exception
      'Please enter Admin remarks before returning the form.';

  end if;


  update public.school_form_submissions

  set

    status = p_status,

    admin_remarks =
      nullif(trim(p_admin_remarks), ''),

    reviewed_at = now(),

    reviewed_by = auth.uid()

  where
    id = p_submission_id
    and status = 'submitted';


  if not found
  then

    raise exception
      'Pending submission not found or it has already been reviewed.';

  end if;

end;
$$;


-- ============================================================
-- REMOVE OLD PACKAGE SUBMISSION FUNCTION
-- ============================================================

drop function if exists
public.submit_school_forms_to_admin(uuid);


-- ============================================================
-- PERMISSIONS
-- ============================================================

revoke all
on function public.set_class_form_readiness(
  uuid,
  text,
  boolean
)
from public;


revoke all
on function public.submit_school_form_to_admin(
  uuid,
  text
)
from public;


revoke all
on function public.review_school_form_submission(
  uuid,
  text,
  text
)
from public;


grant execute
on function public.set_class_form_readiness(
  uuid,
  text,
  boolean
)
to authenticated;


grant execute
on function public.submit_school_form_to_admin(
  uuid,
  text
)
to authenticated;


grant execute
on function public.review_school_form_submission(
  uuid,
  text,
  text
)
to authenticated;
