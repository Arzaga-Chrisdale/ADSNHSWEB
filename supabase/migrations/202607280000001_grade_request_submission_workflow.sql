-- ============================================================
-- 202607280000001_grade_request_submission_workflow.sql
-- MERGED VERSION
--
-- Original workflow + Complete / Reject request extension.
-- ============================================================

-- Two independent grade workflows:
-- 1. Class Adviser -> Subject Teacher grade request
-- 2. Subject Teacher -> Class Adviser direct grade submission
--
-- This migration is additive. It does not delete existing grade_requests rows.

create extension if not exists pgcrypto;

create table if not exists public.grade_request_batches (
  id uuid primary key default gen_random_uuid(),
  adviser_id uuid not null references auth.users(id) on delete cascade,
  advisory_class_id uuid not null references public.classes(id) on delete cascade,
  grading_period text not null check (grading_period in ('1', '2', '3', 'final')),
  recipient_role text not null default 'subject_teacher'
    check (recipient_role in ('class_adviser', 'subject_teacher')),
  recipient_id uuid references auth.users(id) on delete set null,
  request_message text,
  overall_status text not null default 'pending'
    check (overall_status in ('pending', 'completed')),
  is_finalized boolean not null default false,
  requested_at timestamptz not null default now(),
  completed_at timestamptz,
  finalized_at timestamptz
);

create table if not exists public.grade_request_subjects (
  id uuid primary key default gen_random_uuid(),
  batch_id uuid not null references public.grade_request_batches(id) on delete cascade,
  subject_class_id uuid references public.classes(id) on delete set null,
  subject_teacher_id uuid not null references auth.users(id) on delete cascade,
  subject text not null,
  status text not null default 'pending'
    check (status in ('pending', 'submitted')),
  requested_at timestamptz not null default now(),
  submitted_at timestamptz,
  teacher_note text,
  unique (batch_id, subject_teacher_id, subject)
);

create table if not exists public.grade_request_scores (
  subject_request_id uuid not null
    references public.grade_request_subjects(id) on delete cascade,
  student_id uuid not null references public.students(id) on delete cascade,
  score numeric(5,2) not null check (score between 0 and 100),
  primary key (subject_request_id, student_id)
);

create table if not exists public.grade_request_history (
  id uuid primary key default gen_random_uuid(),
  batch_id uuid not null references public.grade_request_batches(id) on delete cascade,
  subject_request_id uuid references public.grade_request_subjects(id) on delete cascade,
  actor_id uuid references auth.users(id) on delete set null,
  event_type text not null,
  old_status text,
  new_status text,
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

-- Upgrade compatibility for projects where grade_request_batches was created
-- by an older migration. CREATE TABLE IF NOT EXISTS does not add new columns
-- to a table that already exists, so add the two-workflow fields explicitly
-- before creating indexes, policies, the view, or RPC functions that use them.
alter table public.grade_request_batches
  add column if not exists recipient_role text;

alter table public.grade_request_batches
  add column if not exists recipient_id uuid;

alter table public.grade_request_batches
  add column if not exists request_message text;

alter table public.grade_request_batches
  add column if not exists overall_status text;

alter table public.grade_request_batches
  add column if not exists is_finalized boolean;

alter table public.grade_request_batches
  add column if not exists requested_at timestamptz;

alter table public.grade_request_batches
  add column if not exists completed_at timestamptz;

alter table public.grade_request_batches
  add column if not exists finalized_at timestamptz;

update public.grade_request_batches
set recipient_role = 'subject_teacher'
where recipient_role is null;

update public.grade_request_batches
set overall_status = 'pending'
where overall_status is null;

update public.grade_request_batches
set is_finalized = false
where is_finalized is null;

update public.grade_request_batches
set requested_at = now()
where requested_at is null;

alter table public.grade_request_batches
  alter column recipient_role set default 'subject_teacher';

alter table public.grade_request_batches
  alter column recipient_role set not null;

alter table public.grade_request_batches
  alter column overall_status set default 'pending';

alter table public.grade_request_batches
  alter column overall_status set not null;

alter table public.grade_request_batches
  alter column is_finalized set default false;

alter table public.grade_request_batches
  alter column is_finalized set not null;

alter table public.grade_request_batches
  alter column requested_at set default now();

alter table public.grade_request_batches
  alter column requested_at set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.grade_request_batches'::regclass
      and conname = 'grade_request_batches_recipient_role_check'
  ) then
    alter table public.grade_request_batches
      add constraint grade_request_batches_recipient_role_check
      check (recipient_role in ('class_adviser', 'subject_teacher'));
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.grade_request_batches'::regclass
      and conname = 'grade_request_batches_overall_status_check'
  ) then
    alter table public.grade_request_batches
      add constraint grade_request_batches_overall_status_check
      check (overall_status in ('pending', 'completed'));
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.grade_request_batches'::regclass
      and conname = 'grade_request_batches_recipient_id_fkey'
  ) then
    alter table public.grade_request_batches
      add constraint grade_request_batches_recipient_id_fkey
      foreign key (recipient_id)
      references auth.users(id)
      on delete set null
      not valid;
  end if;
end
$$;

-- Replace the older open-cycle rule, which allowed only one open request for
-- an entire class and grading period. The recipient-aware rule below allows
-- separate open requests for different Subject Teachers or Class Advisers,
-- while still preventing a duplicate open request to the same recipient.
alter table public.grade_request_batches
  drop constraint if exists grade_request_batches_one_open_cycle;

drop index if exists public.grade_request_batches_one_open_cycle;

create unique index grade_request_batches_one_open_cycle
  on public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period,
    recipient_role,
    recipient_id
  )
  where overall_status = 'pending'
    and is_finalized = false
    and recipient_id is not null;

create index if not exists grade_request_batches_adviser_idx
  on public.grade_request_batches (adviser_id, requested_at desc);
create index if not exists grade_request_batches_recipient_idx
  on public.grade_request_batches (recipient_id, requested_at desc);
create index if not exists grade_request_subjects_teacher_idx
  on public.grade_request_subjects (subject_teacher_id, requested_at desc);

alter table public.grade_request_batches enable row level security;
alter table public.grade_request_subjects enable row level security;
alter table public.grade_request_scores enable row level security;
alter table public.grade_request_history enable row level security;

drop policy if exists "participants read grade request batches"
  on public.grade_request_batches;
create policy "participants read grade request batches"
on public.grade_request_batches for select to authenticated
using (
  adviser_id = auth.uid()
  or recipient_id = auth.uid()
  or public.is_admin(auth.uid())
);

drop policy if exists "advisers create grade request batches"
  on public.grade_request_batches;
create policy "advisers create grade request batches"
on public.grade_request_batches for insert to authenticated
with check (adviser_id = auth.uid() or public.is_admin(auth.uid()));

drop policy if exists "participants read grade request subjects"
  on public.grade_request_subjects;
create policy "participants read grade request subjects"
on public.grade_request_subjects for select to authenticated
using (
  subject_teacher_id = auth.uid()
  or exists (
    select 1 from public.grade_request_batches batch
    where batch.id = batch_id
      and (batch.adviser_id = auth.uid() or batch.recipient_id = auth.uid())
  )
  or public.is_admin(auth.uid())
);

drop policy if exists "participants read grade request scores"
  on public.grade_request_scores;
create policy "participants read grade request scores"
on public.grade_request_scores for select to authenticated
using (
  exists (
    select 1
    from public.grade_request_subjects request
    join public.grade_request_batches batch on batch.id = request.batch_id
    where request.id = subject_request_id
      and (
        request.subject_teacher_id = auth.uid()
        or batch.adviser_id = auth.uid()
        or batch.recipient_id = auth.uid()
      )
  )
  or public.is_admin(auth.uid())
);

drop policy if exists "participants read grade request history"
  on public.grade_request_history;
create policy "participants read grade request history"
on public.grade_request_history for select to authenticated
using (
  exists (
    select 1 from public.grade_request_batches batch
    where batch.id = batch_id
      and (
        batch.adviser_id = auth.uid()
        or batch.recipient_id = auth.uid()
      )
  )
  or public.is_admin(auth.uid())
);

-- The workflow view may already exist with an older column order.
-- CREATE OR REPLACE VIEW cannot insert or rename columns in the middle of an
-- existing view, so remove only the view definition and recreate it below.
-- This does not delete any table rows or workflow records.
drop view if exists public.grade_request_workflow_view;

create view public.grade_request_workflow_view
with (security_invoker = true)
as
select
  batch.id as batch_id,
  batch.adviser_id,
  adviser.full_name as adviser_name,
  adviser.email as adviser_email,
  batch.advisory_class_id,
  advisory_class.grade_level,
  advisory_class.section,
  advisory_class.school_year,
  batch.grading_period,
  batch.recipient_role,
  batch.recipient_id,
  batch.overall_status,
  batch.is_finalized,
  batch.requested_at,
  batch.completed_at,
  batch.finalized_at,
  request.id as subject_request_id,
  request.subject_class_id,
  request.subject_teacher_id,
  subject_teacher.full_name as subject_teacher_name,
  subject_teacher.email as subject_teacher_email,
  request.subject,
  request.status as subject_status,
  request.requested_at as subject_requested_at,
  request.submitted_at,
  request.teacher_note,
  batch.request_message,
  count(score.student_id)::integer as submitted_grade_count
from public.grade_request_batches batch
join public.classes advisory_class on advisory_class.id = batch.advisory_class_id
join public.profiles adviser on adviser.id = batch.adviser_id
join public.grade_request_subjects request on request.batch_id = batch.id
left join public.profiles subject_teacher
  on subject_teacher.id = request.subject_teacher_id
left join public.grade_request_scores score
  on score.subject_request_id = request.id
group by
  batch.id,
  adviser.full_name,
  adviser.email,
  advisory_class.grade_level,
  advisory_class.section,
  advisory_class.school_year,
  request.id,
  subject_teacher.full_name,
  subject_teacher.email;

grant select on public.grade_request_workflow_view to authenticated;
grant select on public.grade_request_batches,
  public.grade_request_subjects,
  public.grade_request_scores,
  public.grade_request_history to authenticated;

create or replace function public.list_grade_request_recipient_directory(
  p_advisory_class_id uuid
)
returns table (
  teacher_id uuid,
  full_name text,
  email text,
  teacher_type text,
  assigned_to_class boolean,
  subjects text[]
)
language sql
security definer
set search_path = public
as $$
  with selected_class as (
    select grade_level, section, school_year
    from public.classes
    where id = p_advisory_class_id
      and teacher_id = auth.uid()
  )
  select
    profile.id,
    coalesce(profile.full_name, 'Teacher'),
    coalesce(profile.email, ''),
    coalesce(profile.teacher_type, 'subject_teacher'),
    exists (
      select 1
      from public.classes class_row, selected_class selected
      where class_row.teacher_id = profile.id
        and class_row.grade_level is not distinct from selected.grade_level
        and class_row.section is not distinct from selected.section
        and class_row.school_year is not distinct from selected.school_year
    ),
    coalesce(
      array(
        select distinct class_row.subject
        from public.classes class_row
        where class_row.teacher_id = profile.id
          and class_row.subject is not null
        order by class_row.subject
      ),
      array[]::text[]
    )
  from public.profiles profile
  where profile.teacher_type in ('class_adviser', 'subject_teacher')
  order by profile.teacher_type, profile.email;
$$;

create or replace function public.list_grade_request_teacher_classes(
  p_teacher_id uuid
)
returns table (
  class_id uuid,
  teacher_id uuid,
  subject text,
  grade_level text,
  section text,
  school_year text
)
language sql
security definer
set search_path = public
as $$
  select
    class_row.id,
    class_row.teacher_id,
    class_row.subject,
    class_row.grade_level,
    class_row.section,
    class_row.school_year
  from public.classes class_row
  where class_row.teacher_id = p_teacher_id
  order by class_row.created_at desc;
$$;

-- Remove obsolete overloads before creating the canonical RPC below.
-- PostgreSQL/PostgREST treats functions with the same name but different
-- parameter lists as overloads. Older migrations may have left one or more
-- of these signatures behind, which makes an RPC call ambiguous.
--
-- These statements remove functions only. They do not delete workflow
-- tables, requests, submitted scores, or history records.
drop function if exists public.create_grade_request_batch(
  uuid,
  text
);

drop function if exists public.create_grade_request_batch(
  uuid,
  text,
  uuid
);

drop function if exists public.create_grade_request_batch(
  uuid,
  text,
  uuid,
  text
);

drop function if exists public.create_grade_request_batch(
  uuid,
  text,
  uuid,
  text,
  text
);

drop function if exists public.create_grade_request_batch(
  uuid,
  text,
  uuid,
  text,
  text,
  uuid
);

create function public.create_grade_request_batch(
  p_advisory_class_id uuid,
  p_grading_period text,
  p_recipient_id uuid default null,
  p_recipient_role text default 'subject_teacher',
  p_message text default null,
  p_subject_teacher_id uuid default null
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  new_batch_id uuid;
  selected_recipient uuid := coalesce(p_recipient_id, p_subject_teacher_id);
  selected_subject text;
begin
  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  if p_recipient_role not in ('class_adviser', 'subject_teacher') then
    raise exception 'Invalid recipient role.';
  end if;

  if selected_recipient is null then
    raise exception 'Select a recipient.';
  end if;

  select coalesce(subject, 'Grades')
  into selected_subject
  from public.classes
  where id = p_advisory_class_id;

  insert into public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period,
    recipient_role,
    recipient_id,
    request_message
  )
  values (
    auth.uid(),
    p_advisory_class_id,
    p_grading_period,
    p_recipient_role,
    selected_recipient,
    p_message
  )
  returning id into new_batch_id;

  insert into public.grade_request_subjects (
    batch_id,
    subject_class_id,
    subject_teacher_id,
    subject
  )
  values (
    new_batch_id,
    p_advisory_class_id,
    selected_recipient,
    coalesce(selected_subject, 'Grades')
  );

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    new_status,
    details
  )
  values (
    new_batch_id,
    auth.uid(),
    case
      when p_recipient_role = 'subject_teacher' then 'grade_request_sent'
      else 'grade_submission_sent'
    end,
    'pending',
    jsonb_build_object('recipient_role', p_recipient_role)
  );

  return new_batch_id;
end;
$$;

create or replace function public.submit_grade_request_subject(
  p_subject_request_id uuid,
  p_scores jsonb,
  p_teacher_note text default null
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  target_request public.grade_request_subjects%rowtype;
  score_row jsonb;
begin
  select * into target_request
  from public.grade_request_subjects
  where id = p_subject_request_id
    and subject_teacher_id = auth.uid();

  if not found then
    raise exception 'Grade request not found or not assigned to you.';
  end if;

  delete from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  for score_row in select * from jsonb_array_elements(p_scores)
  loop
    insert into public.grade_request_scores (
      subject_request_id,
      student_id,
      score
    )
    values (
      p_subject_request_id,
      (score_row->>'student_id')::uuid,
      (score_row->>'score')::numeric
    );
  end loop;

  update public.grade_request_subjects
  set status = 'submitted',
      submitted_at = now(),
      teacher_note = nullif(trim(p_teacher_note), '')
  where id = p_subject_request_id;

  update public.grade_request_batches batch
  set overall_status = 'completed',
      completed_at = now()
  where batch.id = target_request.batch_id
    and not exists (
      select 1 from public.grade_request_subjects request
      where request.batch_id = batch.id
        and request.status <> 'submitted'
    );

  insert into public.grade_request_history (
    batch_id,
    subject_request_id,
    actor_id,
    event_type,
    old_status,
    new_status
  )
  values (
    target_request.batch_id,
    p_subject_request_id,
    auth.uid(),
    'grades_submitted',
    target_request.status,
    'submitted'
  );
end;
$$;

create or replace function public.finalize_grade_request_batch(
  p_batch_id uuid
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  update public.grade_request_batches
  set is_finalized = true,
      overall_status = 'completed',
      completed_at = coalesce(completed_at, now()),
      finalized_at = now()
  where id = p_batch_id
    and adviser_id = auth.uid()
    and not exists (
      select 1 from public.grade_request_subjects request
      where request.batch_id = p_batch_id
        and request.status <> 'submitted'
    );

  if not found then
    raise exception 'All Subject Teachers must submit before finalizing.';
  end if;
end;
$$;

grant execute on function public.list_grade_request_recipient_directory(uuid)
  to authenticated;
grant execute on function public.list_grade_request_teacher_classes(uuid)
  to authenticated;
grant execute on function public.create_grade_request_batch(
  uuid, text, uuid, text, text, uuid
) to authenticated;
grant execute on function public.submit_grade_request_subject(uuid, jsonb, text)
  to authenticated;
grant execute on function public.finalize_grade_request_batch(uuid)
  to authenticated;

-- ============================================================
-- COMPLETE / REJECT GRADE REQUEST EXTENSION
-- Merged into 202607280000001_grade_request_submission_workflow.sql
--
-- UI status:
--   pending   -> Pending
--   submitted -> Completed
--   rejected  -> Rejected
--
-- "submitted" remains the internal completed state so existing
-- finalization/import/notification logic remains compatible.
-- ============================================================


-- ------------------------------------------------------------
-- Compatibility columns used by the newer request-form workflow.
-- These statements are safe when the columns already exist.
-- ------------------------------------------------------------

alter table public.grade_request_batches
  add column if not exists status text;

update public.grade_request_batches
set status = coalesce(overall_status, 'pending')
where status is null;

alter table public.grade_request_batches
  alter column status set default 'pending';

alter table public.grade_request_batches
  alter column status set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.grade_request_batches'::regclass
      and conname = 'grade_request_batches_status_check'
  ) then
    alter table public.grade_request_batches
      add constraint grade_request_batches_status_check
      check (status in ('pending', 'completed'));
  end if;
end
$$;


alter table public.grade_request_subjects
  add column if not exists message text;

alter table public.grade_request_subjects
  add column if not exists recipient_role text;

update public.grade_request_subjects
set recipient_role = 'subject_teacher'
where recipient_role is null;

alter table public.grade_request_subjects
  alter column recipient_role set default 'subject_teacher';

alter table public.grade_request_subjects
  alter column recipient_role set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.grade_request_subjects'::regclass
      and conname = 'grade_request_subjects_recipient_role_check'
  ) then
    alter table public.grade_request_subjects
      add constraint grade_request_subjects_recipient_role_check
      check (recipient_role in ('class_adviser', 'subject_teacher'));
  end if;
end
$$;


alter table public.grade_request_subjects
  add column if not exists updated_at timestamptz;

update public.grade_request_subjects
set updated_at = coalesce(requested_at, now())
where updated_at is null;

alter table public.grade_request_subjects
  alter column updated_at set default now();

alter table public.grade_request_subjects
  alter column updated_at set not null;


-- ------------------------------------------------------------
-- Add Rejected to the existing subject request status.
-- ------------------------------------------------------------

alter table public.grade_request_subjects
  drop constraint if exists grade_request_subjects_status_check;

alter table public.grade_request_subjects
  add constraint grade_request_subjects_status_check
  check (status in ('pending', 'submitted', 'rejected'));


-- ------------------------------------------------------------
-- Complete a request.
--
-- Internally this calls the existing submit routine, so the
-- database status stays "submitted". The UI view below exposes
-- it as "completed".
-- ------------------------------------------------------------

create or replace function public.complete_grade_request_subject(
  p_subject_request_id uuid,
  p_scores jsonb,
  p_teacher_note text default null
)
returns void
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_status text;
begin
  select status
  into v_status
  from public.grade_request_subjects
  where id = p_subject_request_id
    and subject_teacher_id = auth.uid();

  if not found then
    raise exception 'You are not assigned to this Grade Request.';
  end if;

  if v_status = 'rejected' then
    raise exception 'This Grade Request was rejected. Wait for the Class Adviser to send another request.';
  end if;

  perform public.submit_grade_request_subject(
    p_subject_request_id,
    p_scores,
    p_teacher_note
  );
end;
$$;


-- ------------------------------------------------------------
-- Reject a pending request.
--
-- A rejected request becomes read-only for the recipient.
-- The Class Adviser can resend the same teacher/class/term later;
-- create_selected_teacher_grade_request() below will reopen the
-- rejected subject row as Pending.
-- ------------------------------------------------------------

create or replace function public.reject_grade_request_subject(
  p_subject_request_id uuid,
  p_teacher_note text default null
)
returns void
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_item public.grade_request_subjects%rowtype;
  v_batch public.grade_request_batches%rowtype;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  select *
  into v_item
  from public.grade_request_subjects
  where id = p_subject_request_id;

  if not found then
    raise exception 'The Grade Request does not exist.';
  end if;

  if v_item.subject_teacher_id <> v_user_id then
    raise exception 'You are not assigned to this Grade Request.';
  end if;

  select *
  into v_batch
  from public.grade_request_batches
  where id = v_item.batch_id;

  if not found then
    raise exception 'The Grade Request batch does not exist.';
  end if;

  if v_batch.is_finalized then
    raise exception 'This grading period has already been finalized.';
  end if;

  if v_item.status = 'submitted' then
    raise exception 'A completed Grade Request cannot be rejected.';
  end if;

  if v_item.status = 'rejected' then
    return;
  end if;

  -- Remove any previously saved scores for this request.
  delete from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  update public.grade_request_subjects
  set
    status = 'rejected',
    submitted_at = null,
    teacher_note = nullif(trim(p_teacher_note), ''),
    updated_at = now()
  where id = p_subject_request_id;

  insert into public.grade_request_history (
    batch_id,
    subject_request_id,
    actor_id,
    event_type,
    old_status,
    new_status,
    details
  ) values (
    v_item.batch_id,
    p_subject_request_id,
    v_user_id,
    'subject_request_rejected',
    v_item.status,
    'rejected',
    jsonb_build_object(
      'subject', v_item.subject,
      'note', nullif(trim(p_teacher_note), ''),
      'rejected_at', now()
    )
  );
end;
$$;


-- ------------------------------------------------------------
-- Create / resend one selected teacher Grade Request.
--
-- If the same teacher/class/subject request was rejected earlier,
-- the existing subject request row is reopened as Pending.
-- ------------------------------------------------------------

create or replace function public.create_selected_teacher_grade_request(
  p_advisory_class_id uuid,
  p_grading_period text,
  p_recipient_id uuid,
  p_recipient_role text,
  p_message text default null
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
  v_batch_created boolean := false;
  v_inserted integer := 0;
  v_class public.classes%rowtype;
  v_actual_role text;
  v_subject text;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  if p_recipient_role not in ('class_adviser', 'subject_teacher') then
    raise exception 'The selected teacher type is invalid.';
  end if;

  if p_recipient_id = v_user_id then
    raise exception 'You cannot send a Grade Request to your own account.';
  end if;

  if not exists (
    select 1
    from public.profiles requester
    where requester.id = v_user_id
      and requester.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can create Grade Requests.';
  end if;

  select recipient.teacher_type
  into v_actual_role
  from public.profiles recipient
  where recipient.id = p_recipient_id;

  if not found then
    raise exception 'The selected teacher account does not exist.';
  end if;

  if v_actual_role is distinct from p_recipient_role then
    raise exception 'The selected email does not match its teacher type.';
  end if;

  select selected_class.*
  into v_class
  from public.classes selected_class
  where selected_class.id = p_advisory_class_id;

  if not found then
    raise exception 'The selected class does not exist.';
  end if;

  if v_class.teacher_id is distinct from p_recipient_id then
    raise exception 'The selected class does not belong to the selected teacher.';
  end if;

  v_subject := coalesce(
    nullif(trim(v_class.subject), ''),
    case
      when p_recipient_role = 'class_adviser' then 'Advisory Grades'
      else 'Subject Grades'
    end
  );

  -- Reuse the latest unfinished batch for this adviser/class/term.
  select batch.id
  into v_batch_id
  from public.grade_request_batches batch
  where batch.adviser_id = v_user_id
    and batch.advisory_class_id = p_advisory_class_id
    and batch.grading_period = p_grading_period
    and batch.is_finalized = false
  order by batch.requested_at desc
  limit 1;

  if v_batch_id is null then
    insert into public.grade_request_batches (
      adviser_id,
      advisory_class_id,
      grading_period,
      status,
      overall_status,
      recipient_role,
      recipient_id,
      request_message
    ) values (
      v_user_id,
      p_advisory_class_id,
      p_grading_period,
      'pending',
      'pending',
      p_recipient_role,
      p_recipient_id,
      nullif(trim(p_message), '')
    )
    returning id into v_batch_id;

    v_batch_created := true;
  else
    -- Keep the batch open when a rejected request is resent.
    update public.grade_request_batches
    set
      status = 'pending',
      overall_status = 'pending',
      completed_at = null,
      recipient_role = p_recipient_role,
      recipient_id = p_recipient_id,
      request_message = nullif(trim(p_message), '')
    where id = v_batch_id;
  end if;

  insert into public.grade_request_subjects (
    batch_id,
    subject_class_id,
    subject_teacher_id,
    recipient_role,
    subject,
    status,
    message
  ) values (
    v_batch_id,
    p_advisory_class_id,
    p_recipient_id,
    p_recipient_role,
    v_subject,
    'pending',
    nullif(trim(p_message), '')
  )
  on conflict (batch_id, subject_teacher_id, subject)
  do update set
    subject_class_id = excluded.subject_class_id,
    recipient_role = excluded.recipient_role,
    status = 'pending',
    requested_at = now(),
    submitted_at = null,
    teacher_note = null,
    message = excluded.message,
    updated_at = now()
  where public.grade_request_subjects.status = 'rejected';

  get diagnostics v_inserted = row_count;

  if v_inserted = 0 then
    if v_batch_created then
      delete from public.grade_request_batches
      where id = v_batch_id;
    end if;

    raise exception 'A Grade Request was already sent to this teacher for the selected class and grading period.';
  end if;

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    new_status,
    details
  ) values (
    v_batch_id,
    v_user_id,
    case
      when v_batch_created then 'request_created'
      else 'request_recipients_added'
    end,
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'advisory_class_id', p_advisory_class_id,
      'recipient_id', p_recipient_id,
      'recipient_role', p_recipient_role,
      'message', nullif(trim(p_message), '')
    )
  );

  return v_batch_id;
end;
$$;


-- ------------------------------------------------------------
-- Recreate the workflow view.
--
-- We drop only the VIEW definition, not any workflow tables or
-- records. This lets us safely change the output column order and
-- expose "submitted" to React as "completed".
-- ------------------------------------------------------------

drop view if exists public.grade_request_workflow_view;

create view public.grade_request_workflow_view
with (security_invoker = true) as
select
  b.id as batch_id,
  b.adviser_id,
  adviser.full_name as adviser_name,
  adviser.email as adviser_email,
  b.advisory_class_id,
  ac.grade_level,
  ac.section,
  ac.school_year,
  b.grading_period,
  b.status as overall_status,
  b.is_finalized,
  b.requested_at,
  b.completed_at,
  b.finalized_at,
  s.id as subject_request_id,
  s.subject_class_id,
  s.subject_teacher_id,
  recipient.full_name as subject_teacher_name,
  recipient.email as subject_teacher_email,
  s.subject,
  case
    when s.status = 'submitted' then 'completed'
    else s.status
  end as subject_status,
  s.requested_at as subject_requested_at,
  s.submitted_at,
  s.teacher_note,
  count(sc.student_id)::integer as submitted_grade_count,
  s.message as request_message,
  s.recipient_role,
  s.subject_teacher_id as recipient_id
from public.grade_request_batches b
join public.classes ac
  on ac.id = b.advisory_class_id
join public.profiles adviser
  on adviser.id = b.adviser_id
join public.grade_request_subjects s
  on s.batch_id = b.id
join public.profiles recipient
  on recipient.id = s.subject_teacher_id
left join public.grade_request_scores sc
  on sc.subject_request_id = s.id
group by
  b.id,
  adviser.full_name,
  adviser.email,
  ac.grade_level,
  ac.section,
  ac.school_year,
  s.id,
  recipient.full_name,
  recipient.email,
  s.message,
  s.recipient_role;


-- ------------------------------------------------------------
-- Permissions
-- ------------------------------------------------------------

grant select on public.grade_request_workflow_view to authenticated;

grant execute on function public.complete_grade_request_subject(
  uuid,
  jsonb,
  text
) to authenticated;

grant execute on function public.reject_grade_request_subject(
  uuid,
  text
) to authenticated;

grant execute on function public.create_selected_teacher_grade_request(
  uuid,
  text,
  uuid,
  text,
  text
) to authenticated;

notify pgrst, 'reload schema';



-- ============================================================
-- FINAL OVERRIDE:
-- PRESERVE REJECTED REQUESTS + ALLOW MULTIPLE SENDS
--
-- This final section intentionally overrides only the send behavior.
-- It does NOT delete the earlier SQL source in this file.
-- ============================================================

-- ============================================================
-- Preserve Rejected Requests + Allow Multiple Grade Requests
--
-- FIX:
-- Previously, resending after a rejection reopened the SAME
-- grade_request_subjects row:
--
--   Rejected -> Pending
--
-- That caused the old Rejected record to disappear from the
-- Class Adviser's request history.
--
-- NEW BEHAVIOR:
-- Every click of "Send Grade Request" creates a NEW batch/request.
--
-- Example:
--   Request #1 -> Rejected   (kept permanently)
--   Request #2 -> Pending    (new row)
--   Request #3 -> Completed  (another independent row)
--
-- No old rejected request is changed back to Pending.
-- ============================================================


-- ------------------------------------------------------------
-- 1. Remove the old "one open request" restriction.
--
-- This allows the Class Adviser to send multiple independent
-- requests to the same recipient/class/term.
-- ------------------------------------------------------------

alter table public.grade_request_batches
  drop constraint if exists grade_request_batches_one_open_cycle;

drop index if exists public.grade_request_batches_one_open_cycle;


-- ------------------------------------------------------------
-- 2. Replace the send RPC.
--
-- IMPORTANT:
-- This function NEVER searches for/reuses an older batch.
-- It NEVER changes a rejected subject row back to pending.
-- Every send creates:
--   - one new grade_request_batches row
--   - one new grade_request_subjects row
--   - one new history event
-- ------------------------------------------------------------

create or replace function public.create_selected_teacher_grade_request(
  p_advisory_class_id uuid,
  p_grading_period text,
  p_recipient_id uuid,
  p_recipient_role text,
  p_message text default null
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
  v_subject_request_id uuid;
  v_class public.classes%rowtype;
  v_actual_role text;
  v_subject text;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  if p_recipient_role not in ('class_adviser', 'subject_teacher') then
    raise exception 'The selected teacher type is invalid.';
  end if;

  if p_recipient_id is null then
    raise exception 'Select a recipient.';
  end if;

  if p_recipient_id = v_user_id then
    raise exception 'You cannot send a Grade Request to your own account.';
  end if;

  -- Only a Class Adviser can send this Grade Request workflow.
  if not exists (
    select 1
    from public.profiles requester
    where requester.id = v_user_id
      and requester.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can create Grade Requests.';
  end if;

  -- Verify the selected recipient exists and has the chosen role.
  select recipient.teacher_type
  into v_actual_role
  from public.profiles recipient
  where recipient.id = p_recipient_id;

  if not found then
    raise exception 'The selected teacher account does not exist.';
  end if;

  if v_actual_role is distinct from p_recipient_role then
    raise exception 'The selected email does not match its teacher type.';
  end if;

  -- Verify the selected class.
  select selected_class.*
  into v_class
  from public.classes selected_class
  where selected_class.id = p_advisory_class_id;

  if not found then
    raise exception 'The selected class does not exist.';
  end if;

  if v_class.teacher_id is distinct from p_recipient_id then
    raise exception 'The selected class does not belong to the selected teacher.';
  end if;

  v_subject := coalesce(
    nullif(trim(v_class.subject), ''),
    case
      when p_recipient_role = 'class_adviser' then 'Advisory Grades'
      else 'Subject Grades'
    end
  );

  -- ALWAYS create a brand-new batch.
  insert into public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period,
    status,
    overall_status,
    recipient_role,
    recipient_id,
    request_message,
    is_finalized,
    requested_at,
    completed_at,
    finalized_at
  )
  values (
    v_user_id,
    p_advisory_class_id,
    p_grading_period,
    'pending',
    'pending',
    p_recipient_role,
    p_recipient_id,
    nullif(trim(p_message), ''),
    false,
    now(),
    null,
    null
  )
  returning id into v_batch_id;

  -- ALWAYS create a brand-new subject request under the new batch.
  insert into public.grade_request_subjects (
    batch_id,
    subject_class_id,
    subject_teacher_id,
    recipient_role,
    subject,
    status,
    message,
    requested_at,
    submitted_at,
    teacher_note,
    updated_at
  )
  values (
    v_batch_id,
    p_advisory_class_id,
    p_recipient_id,
    p_recipient_role,
    v_subject,
    'pending',
    nullif(trim(p_message), ''),
    now(),
    null,
    null,
    now()
  )
  returning id into v_subject_request_id;

  insert into public.grade_request_history (
    batch_id,
    subject_request_id,
    actor_id,
    event_type,
    old_status,
    new_status,
    details
  )
  values (
    v_batch_id,
    v_subject_request_id,
    v_user_id,
    'request_created',
    null,
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'advisory_class_id', p_advisory_class_id,
      'recipient_id', p_recipient_id,
      'recipient_role', p_recipient_role,
      'subject', v_subject,
      'message', nullif(trim(p_message), ''),
      'created_at', now()
    )
  );

  return v_batch_id;
end;
$$;


grant execute on function public.create_selected_teacher_grade_request(
  uuid,
  text,
  uuid,
  text,
  text
) to authenticated;


-- Ask PostgREST/Supabase to refresh the RPC schema cache.
notify pgrst, 'reload schema';
