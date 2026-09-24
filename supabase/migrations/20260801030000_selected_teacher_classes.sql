-- 20260801030000_selected_teacher_classes.sql
--
-- Combined migration:
-- 1. Grade 12 custom subject + optional units support.
-- 2. Load all classes owned by a selected Class Adviser or Subject Teacher.
-- 3. Create a Grade Request for the selected teacher's chosen class.
--
-- Safe to run more than once.
-- This migration does not delete existing class, learner, grade, or request data.

begin;

-- ============================================================================
-- 1. Grade 12 custom subject + optional units support
-- ============================================================================
--
-- Frontend rule:
--   Grade 12 subject is free-text.
--   Grade 12 units may be left blank (NULL).
--
-- The existing classes.subject column remains text and continues to store
-- both predefined subjects and custom Grade 12 subject names.

alter table public.classes
  alter column units drop not null;

comment on column public.classes.units is
  'Optional class units. Grade 12 classes may store NULL when no unit value is selected.';


-- ============================================================================
-- 2. Load all classes owned by the selected teacher
-- ============================================================================
--
-- Both Class Adviser and Subject Teacher email selections load all classes
-- owned by the selected account. The request is attached to the selected
-- teacher's chosen class.

create or replace function public.list_selected_teacher_classes(
  p_teacher_id uuid
)
returns table (
  id uuid,
  teacher_id uuid,
  subject text,
  grade_level text,
  section text,
  school_year text
)
language plpgsql
stable
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if not exists (
    select 1
    from public.profiles requester
    where requester.id = v_user_id
      and requester.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can load another teacher''s classes.';
  end if;

  if not exists (
    select 1
    from public.profiles recipient
    where recipient.id = p_teacher_id
      and recipient.teacher_type in ('class_adviser', 'subject_teacher')
  ) then
    raise exception 'The selected email is not a Class Adviser or Subject Teacher account.';
  end if;

  return query
  select
    c.id,
    c.teacher_id,
    c.subject,
    c.grade_level,
    c.section,
    c.school_year
  from public.classes c
  where c.teacher_id = p_teacher_id
  order by
    c.school_year desc,
    c.grade_level,
    c.section,
    c.subject,
    c.created_at desc;
end;
$$;

revoke all on function public.list_selected_teacher_classes(uuid)
from public;

grant execute on function public.list_selected_teacher_classes(uuid)
to authenticated;


-- ============================================================================
-- 3. Create a Grade Request for the selected teacher's chosen class
-- ============================================================================

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
      status
    ) values (
      v_user_id,
      p_advisory_class_id,
      p_grading_period,
      'pending'
    )
    returning id into v_batch_id;

    v_batch_created := true;
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
  on conflict (batch_id, subject_teacher_id, subject) do nothing;

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

revoke all on function public.create_selected_teacher_grade_request(
  uuid,
  text,
  uuid,
  text,
  text
)
from public;

grant execute on function public.create_selected_teacher_grade_request(
  uuid,
  text,
  uuid,
  text,
  text
)
to authenticated;

commit;

-- Tell PostgREST/Supabase to refresh its schema cache after the migration.
notify pgrst, 'reload schema';
