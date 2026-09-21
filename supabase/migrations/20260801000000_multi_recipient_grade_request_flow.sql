-- Unified multi-recipient Grade Request flow.
-- One Class Adviser may select multiple Subject Teachers and other Class
-- Advisers. Every recipient receives an independent subject request for the
-- same advisory class, enters grades, and submits them to the requester.

alter table public.grade_request_subjects
  add column if not exists recipient_role text;

alter table public.grade_request_subjects
  add column if not exists message text;

update public.grade_request_subjects
set recipient_role = 'subject_teacher'
where recipient_role is null;

alter table public.grade_request_subjects
  alter column recipient_role set default 'subject_teacher';

alter table public.grade_request_subjects
  alter column recipient_role set not null;

drop index if exists public.grade_request_batches_one_open_cycle;

create unique index grade_request_batches_one_open_cycle
  on public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period
  )
  where is_finalized = false;

create or replace function public.create_multi_recipient_grade_request(
  p_advisory_class_id uuid,
  p_grading_period text,
  p_recipients jsonb,
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
  v_class public.classes%rowtype;
  v_recipient jsonb;
  v_recipient_id uuid;
  v_recipient_role text;
  v_actual_role text;
  v_inserted integer;
  v_total_inserted integer := 0;
  v_batch_created boolean := false;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  if jsonb_typeof(p_recipients) <> 'array'
    or jsonb_array_length(p_recipients) = 0 then
    raise exception 'Select at least one recipient.';
  end if;

  select *
  into v_class
  from public.classes c
  where c.id = p_advisory_class_id;

  if not found then
    raise exception 'The selected advisory class does not exist.';
  end if;

  if v_class.teacher_id is distinct from v_user_id then
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

  select b.id
  into v_batch_id
  from public.grade_request_batches b
  where b.adviser_id = v_user_id
    and b.advisory_class_id = p_advisory_class_id
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

  for v_recipient in
    select value from jsonb_array_elements(p_recipients)
  loop
    begin
      v_recipient_id := (v_recipient ->> 'recipient_id')::uuid;
    exception when others then
      raise exception 'A selected recipient is invalid.';
    end;

    v_recipient_role := v_recipient ->> 'recipient_role';

    if v_recipient_id = v_user_id then
      raise exception 'You cannot send a Grade Request to your own account.';
    end if;

    if v_recipient_role not in ('class_adviser', 'subject_teacher') then
      raise exception 'A selected teacher type is invalid.';
    end if;

    select p.teacher_type
    into v_actual_role
    from public.profiles p
    where p.id = v_recipient_id;

    if not found then
      raise exception 'A selected teacher account does not exist.';
    end if;

    if v_actual_role is distinct from v_recipient_role then
      raise exception 'A selected email does not match its teacher type.';
    end if;

    if v_recipient_role = 'subject_teacher' then
      insert into public.grade_request_subjects (
        batch_id,
        subject_class_id,
        subject_teacher_id,
        recipient_role,
        subject,
        status,
        message
      )
      select distinct on (lower(trim(c.subject)))
        v_batch_id,
        c.id,
        v_recipient_id,
        'subject_teacher',
        trim(c.subject),
        'pending',
        nullif(trim(p_message), '')
      from public.classes c
      where c.teacher_id = v_recipient_id
        and c.subject is not null
        and trim(c.subject) <> ''
        and lower(trim(coalesce(c.grade_level, ''))) =
            lower(trim(coalesce(v_class.grade_level, '')))
        and lower(trim(coalesce(c.section, ''))) =
            lower(trim(coalesce(v_class.section, '')))
        and lower(trim(coalesce(c.school_year, ''))) =
            lower(trim(coalesce(v_class.school_year, '')))
      order by lower(trim(c.subject)), c.id
      on conflict (batch_id, subject_teacher_id, subject) do nothing;

      get diagnostics v_inserted = row_count;

      if v_inserted = 0 and not exists (
        select 1
        from public.grade_request_subjects s
        where s.batch_id = v_batch_id
          and s.subject_teacher_id = v_recipient_id
      ) then
        raise exception 'A selected Subject Teacher is not assigned to this grade, section, and school year.';
      end if;
    else
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
        v_recipient_id,
        'class_adviser',
        'Advisory Grades',
        'pending',
        nullif(trim(p_message), '')
      )
      on conflict (batch_id, subject_teacher_id, subject) do nothing;

      get diagnostics v_inserted = row_count;
    end if;

    v_total_inserted := v_total_inserted + v_inserted;
  end loop;

  if v_total_inserted = 0 then
    if v_batch_created then
      delete from public.grade_request_batches where id = v_batch_id;
    end if;

    raise exception 'Grade Requests have already been sent to all selected recipients for this class and grading period.';
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
    case when v_batch_created then 'request_created' else 'request_recipients_added' end,
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'advisory_class_id', p_advisory_class_id,
      'selected_recipient_count', jsonb_array_length(p_recipients),
      'created_subject_request_count', v_total_inserted,
      'message', nullif(trim(p_message), '')
    )
  );

  return v_batch_id;
end;
$$;

create or replace function public.finalize_grade_request_batch(
  p_batch_id uuid
)
returns void
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch public.grade_request_batches%rowtype;
begin
  select *
  into v_batch
  from public.grade_request_batches b
  where b.id = p_batch_id;

  if not found then
    raise exception 'The Grade Request does not exist.';
  end if;

  if v_batch.adviser_id is distinct from v_user_id then
    raise exception 'Only the requesting Class Adviser can finalize these grades.';
  end if;

  if v_batch.status <> 'completed' then
    raise exception 'All selected recipients must submit their grades before finalization.';
  end if;

  if v_batch.is_finalized then
    return;
  end if;

  update public.grade_request_batches
  set
    is_finalized = true,
    finalized_at = now()
  where id = p_batch_id;

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    old_status,
    new_status,
    details
  ) values (
    p_batch_id,
    v_user_id,
    'grades_finalized_after_review',
    'completed',
    'completed',
    jsonb_build_object('finalized_at', now())
  );
end;
$$;

-- Recipients need read access to the requester's class, learners, and current
-- subject grades while the request is assigned to them.
drop policy if exists grade_request_classes_select_participants
on public.classes;

create policy grade_request_classes_select_participants
on public.classes
for select
to authenticated
using (
  teacher_id = auth.uid()
  or exists (
    select 1
    from public.grade_request_batches b
    join public.grade_request_subjects s on s.batch_id = b.id
    where b.advisory_class_id = classes.id
      and s.subject_teacher_id = auth.uid()
  )
  or public.is_admin(auth.uid())
);

drop policy if exists grade_request_students_select_participants
on public.students;

create policy grade_request_students_select_participants
on public.students
for select
to authenticated
using (
  teacher_id = auth.uid()
  or exists (
    select 1
    from public.grade_request_batches b
    join public.grade_request_subjects s on s.batch_id = b.id
    where b.advisory_class_id = students.class_id
      and s.subject_teacher_id = auth.uid()
  )
  or public.is_admin(auth.uid())
);

drop policy if exists grade_request_grades_select_participants
on public.grades;

create policy grade_request_grades_select_participants
on public.grades
for select
to authenticated
using (
  teacher_id = auth.uid()
  or exists (
    select 1
    from public.grade_request_batches b
    join public.grade_request_subjects s on s.batch_id = b.id
    where b.advisory_class_id = grades.class_id
      and s.subject_teacher_id = auth.uid()
      and s.subject = grades.subject
      and b.grading_period = grades.term
  )
  or public.is_admin(auth.uid())
);

drop policy if exists grade_request_profiles_select_participants
on public.profiles;

create policy grade_request_profiles_select_participants
on public.profiles
for select
to authenticated
using (
  id = auth.uid()
  or exists (
    select 1
    from public.grade_request_subjects s
    join public.grade_request_batches b on b.id = s.batch_id
    where (
      b.adviser_id = auth.uid()
      and s.subject_teacher_id = profiles.id
    ) or (
      s.subject_teacher_id = auth.uid()
      and b.adviser_id = profiles.id
    )
  )
  or public.is_admin(auth.uid())
);

-- Include a direct recipient id in the view and keep recipient role on each
-- request item, allowing both teacher types to use the same Receive Form.
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
  s.status as subject_status,
  s.requested_at as subject_requested_at,
  s.submitted_at,
  s.teacher_note,
  count(sc.id)::integer as submitted_grade_count,
  s.message as request_message,
  s.recipient_role,
  s.subject_teacher_id as recipient_id
from public.grade_request_batches b
join public.classes ac on ac.id = b.advisory_class_id
join public.profiles adviser on adviser.id = b.adviser_id
join public.grade_request_subjects s on s.batch_id = b.id
join public.profiles recipient on recipient.id = s.subject_teacher_id
left join public.grade_request_scores sc on sc.subject_request_id = s.id
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

create or replace function public.delete_assigned_grade_request(
  p_subject_request_id uuid
)
returns boolean
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
begin
  select s.batch_id
  into v_batch_id
  from public.grade_request_subjects s
  where s.id = p_subject_request_id
    and s.subject_teacher_id = v_user_id;

  if not found then
    return false;
  end if;

  delete from public.grade_request_subjects
  where id = p_subject_request_id
    and subject_teacher_id = v_user_id;

  if not exists (
    select 1 from public.grade_request_subjects s where s.batch_id = v_batch_id
  ) then
    delete from public.grade_request_batches where id = v_batch_id;
  end if;

  return true;
end;
$$;

create or replace function public.delete_received_class_adviser_submission(
  p_subject_request_id uuid
)
returns boolean
language sql
security definer
set search_path = public, pg_temp
as $$
  select public.delete_assigned_grade_request(p_subject_request_id);
$$;

create or replace function public.delete_sent_grade_request(
  p_subject_request_id uuid
)
returns boolean
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
begin
  select s.batch_id
  into v_batch_id
  from public.grade_request_subjects s
  join public.grade_request_batches b on b.id = s.batch_id
  where s.id = p_subject_request_id
    and b.adviser_id = v_user_id;

  if not found then
    return false;
  end if;

  delete from public.grade_request_subjects where id = p_subject_request_id;

  if not exists (
    select 1 from public.grade_request_subjects s where s.batch_id = v_batch_id
  ) then
    delete from public.grade_request_batches where id = v_batch_id;
  end if;

  return true;
end;
$$;

grant execute on function public.create_multi_recipient_grade_request(
  uuid,
  text,
  jsonb,
  text
) to authenticated;

grant execute on function public.finalize_grade_request_batch(uuid)
to authenticated;

grant execute on function public.delete_assigned_grade_request(uuid)
to authenticated;

grant execute on function public.delete_received_class_adviser_submission(uuid)
to authenticated;

grant execute on function public.delete_sent_grade_request(uuid)
to authenticated;

grant select on public.grade_request_workflow_view to authenticated;

notify pgrst, 'reload schema';
