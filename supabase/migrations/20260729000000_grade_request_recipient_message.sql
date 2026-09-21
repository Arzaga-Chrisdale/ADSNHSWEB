-- Add recipient selection and an optional Class Adviser message to the
-- existing Grade Request workflow without deleting current request data.

alter table public.grade_request_subjects
  add column if not exists message text;

-- Replace the two-parameter function with a backward-compatible function.
-- The last two parameters have defaults, so existing calls that only send
-- class ID and grading period continue to request all assigned teachers.
drop function if exists public.create_grade_request_batch(uuid, text);

create or replace function public.create_grade_request_batch(
  p_advisory_class_id uuid,
  p_grading_period text,
  p_subject_teacher_id uuid default null,
  p_message text default null
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
  v_class public.classes%rowtype;
  v_inserted_subject_count integer := 0;
  v_batch_created boolean := false;
  v_event_type text;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  select * into v_class
  from public.classes
  where id = p_advisory_class_id;

  if not found then
    raise exception 'The selected advisory class does not exist.';
  end if;

  if v_class.teacher_id <> v_user_id then
    raise exception 'Only the assigned Class Adviser can request grades for this class.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can create grade requests.';
  end if;

  -- A specific recipient must be an assigned Subject Teacher for the same
  -- grade level, section, and school year as the advisory class.
  if p_subject_teacher_id is not null and not exists (
    select 1
    from public.classes c
    join public.profiles p on p.id = c.teacher_id
    where c.teacher_id = p_subject_teacher_id
      and p.teacher_type = 'subject_teacher'
      and c.subject is not null
      and trim(c.subject) <> ''
      and lower(trim(coalesce(c.grade_level, ''))) =
          lower(trim(coalesce(v_class.grade_level, '')))
      and lower(trim(coalesce(c.section, ''))) =
          lower(trim(coalesce(v_class.section, '')))
      and lower(trim(coalesce(c.school_year, ''))) =
          lower(trim(coalesce(v_class.school_year, '')))
  ) then
    raise exception 'The selected Subject Teacher is not assigned to this grade, section, and school year.';
  end if;

  -- Reuse an existing open cycle. This allows the adviser to request one
  -- teacher first and add other assigned teachers later without creating a
  -- duplicate grading-period cycle.
  select b.id into v_batch_id
  from public.grade_request_batches b
  where b.advisory_class_id = p_advisory_class_id
    and b.grading_period = p_grading_period
    and b.is_finalized = false
  order by b.requested_at desc
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
    subject,
    status,
    message
  )
  select distinct on (lower(trim(c.subject)), c.teacher_id)
    v_batch_id,
    c.id,
    c.teacher_id,
    trim(c.subject),
    'pending',
    nullif(trim(p_message), '')
  from public.classes c
  join public.profiles p on p.id = c.teacher_id
  where p.teacher_type = 'subject_teacher'
    and c.subject is not null
    and trim(c.subject) <> ''
    and lower(trim(coalesce(c.grade_level, ''))) =
        lower(trim(coalesce(v_class.grade_level, '')))
    and lower(trim(coalesce(c.section, ''))) =
        lower(trim(coalesce(v_class.section, '')))
    and lower(trim(coalesce(c.school_year, ''))) =
        lower(trim(coalesce(v_class.school_year, '')))
    and (
      p_subject_teacher_id is null
      or c.teacher_id = p_subject_teacher_id
    )
  order by lower(trim(c.subject)), c.teacher_id, c.id
  on conflict (batch_id, subject_teacher_id, subject) do nothing;

  get diagnostics v_inserted_subject_count = row_count;

  if v_inserted_subject_count = 0 then
    if v_batch_created then
      delete from public.grade_request_batches
      where id = v_batch_id;
    end if;

    if p_subject_teacher_id is null then
      raise exception 'Requests have already been sent to all currently assigned Subject Teachers for this grading period.';
    end if;

    raise exception 'A request has already been sent to the selected Subject Teacher for this grading period.';
  end if;

  v_event_type := case
    when v_batch_created then 'request_created'
    else 'request_recipients_added'
  end;

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    new_status,
    details
  ) values (
    v_batch_id,
    v_user_id,
    v_event_type,
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'subject_count', v_inserted_subject_count,
      'advisory_class_id', p_advisory_class_id,
      'recipient_teacher_id', p_subject_teacher_id,
      'recipient_scope', case
        when p_subject_teacher_id is null then 'all_assigned_subject_teachers'
        else 'selected_subject_teacher'
      end,
      'message', nullif(trim(p_message), '')
    )
  );

  return v_batch_id;
end;
$$;

-- Keep every existing column in the same order and append request_message at
-- the end so current TypeScript queries remain compatible.
create or replace view public.grade_request_workflow_view
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
  teacher.full_name as subject_teacher_name,
  teacher.email as subject_teacher_email,
  s.subject,
  s.status as subject_status,
  s.requested_at as subject_requested_at,
  s.submitted_at,
  s.teacher_note,
  count(sc.id) as submitted_grade_count,
  s.message as request_message
from public.grade_request_batches b
join public.classes ac on ac.id = b.advisory_class_id
join public.profiles adviser on adviser.id = b.adviser_id
join public.grade_request_subjects s on s.batch_id = b.id
join public.profiles teacher on teacher.id = s.subject_teacher_id
left join public.grade_request_scores sc on sc.subject_request_id = s.id
group by
  b.id,
  adviser.full_name,
  adviser.email,
  ac.grade_level,
  ac.section,
  ac.school_year,
  s.id,
  teacher.full_name,
  teacher.email,
  s.message;

grant execute on function public.create_grade_request_batch(uuid, text, uuid, text)
to authenticated;

grant select on public.grade_request_workflow_view to authenticated;

notify pgrst, 'reload schema';
