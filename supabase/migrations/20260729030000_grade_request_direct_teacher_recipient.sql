-- Direct grade-request recipient update.
-- A logged-in Class Adviser can send one request to one selected
-- Class Adviser or Subject Teacher email without requiring the selected
-- class to match the requester's own advisory class.

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
    where conname = 'grade_request_subjects_recipient_role_check'
      and conrelid = 'public.grade_request_subjects'::regclass
  ) then
    alter table public.grade_request_subjects
      add constraint grade_request_subjects_recipient_role_check
      check (recipient_role in ('class_adviser', 'subject_teacher'));
  end if;
end;
$$;

-- Allow separate requesting advisers to have an open request for the same
-- selected class and grading period.
drop index if exists public.grade_request_batches_one_open_cycle;
create unique index if not exists grade_request_batches_one_open_cycle
  on public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period
  )
  where is_finalized = false;

-- Remove older overloads so PostgREST exposes one unambiguous RPC contract.
drop function if exists public.create_grade_request_batch(uuid, text);
drop function if exists public.create_grade_request_batch(uuid, text, uuid, text);
drop function if exists public.create_grade_request_batch(uuid, text, uuid, text, text);

create or replace function public.create_grade_request_batch(
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
  v_class public.classes%rowtype;
  v_recipient_type text;
  v_subject text;
  v_batch_created boolean := false;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can create grade requests.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  if p_recipient_role not in ('class_adviser', 'subject_teacher') then
    raise exception 'Select a valid teacher type.';
  end if;

  select p.teacher_type
  into v_recipient_type
  from public.profiles p
  where p.id = p_recipient_id;

  if not found then
    raise exception 'The selected teacher account does not exist.';
  end if;

  if v_recipient_type is distinct from p_recipient_role then
    raise exception 'The selected email does not match the selected teacher type.';
  end if;

  select *
  into v_class
  from public.classes c
  where c.id = p_advisory_class_id;

  if not found then
    raise exception 'The selected class does not exist.';
  end if;

  -- The class dropdown is populated from the selected email account. Verify
  -- the selected class still belongs to that recipient at submission time.
  if v_class.teacher_id is distinct from p_recipient_id then
    raise exception 'The selected class is not assigned to the selected email account.';
  end if;

  v_subject := nullif(trim(coalesce(v_class.subject, '')), '');
  if v_subject is null then
    v_subject := case
      when p_recipient_role = 'class_adviser' then 'Advisory Grades'
      else 'Assigned Subject'
    end;
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

  if not found then
    if v_batch_created then
      delete from public.grade_request_batches
      where id = v_batch_id;
    end if;

    raise exception 'A request has already been sent to the selected email for this class and grading period.';
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
    case when v_batch_created then 'request_created' else 'request_recipient_added' end,
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'class_id', p_advisory_class_id,
      'recipient_id', p_recipient_id,
      'recipient_role', p_recipient_role,
      'subject', v_subject,
      'message', nullif(trim(p_message), '')
    )
  );

  return v_batch_id;
end;
$$;

-- Allow the selected recipient to submit the requested grades whether the
-- account is a Class Adviser or Subject Teacher.
create or replace function public.submit_grade_request_subject(
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
  v_user_id uuid := auth.uid();
  v_item public.grade_request_subjects%rowtype;
  v_batch public.grade_request_batches%rowtype;
  v_profile_type text;
  v_score jsonb;
  v_student_id uuid;
  v_numeric_score numeric;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  select *
  into v_item
  from public.grade_request_subjects
  where id = p_subject_request_id;

  if not found then
    raise exception 'The grade request does not exist.';
  end if;

  if v_item.subject_teacher_id <> v_user_id then
    raise exception 'This grade request was not sent to your account.';
  end if;

  select p.teacher_type
  into v_profile_type
  from public.profiles p
  where p.id = v_user_id;

  if v_profile_type is distinct from v_item.recipient_role then
    raise exception 'Your account role does not match this grade-request recipient.';
  end if;

  select *
  into v_batch
  from public.grade_request_batches
  where id = v_item.batch_id;

  if not found then
    raise exception 'The grade-request cycle does not exist.';
  end if;

  if v_batch.is_finalized then
    raise exception 'This grading period has already been finalized.';
  end if;

  if jsonb_typeof(p_scores) <> 'array' or jsonb_array_length(p_scores) = 0 then
    raise exception 'Submit at least one learner grade.';
  end if;

  delete from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  for v_score in
    select * from jsonb_array_elements(p_scores)
  loop
    v_student_id := (v_score ->> 'student_id')::uuid;
    v_numeric_score := (v_score ->> 'score')::numeric;

    if v_numeric_score < 0 or v_numeric_score > 100 then
      raise exception 'Every grade must be between 0 and 100.';
    end if;

    if not exists (
      select 1
      from public.students s
      where s.id = v_student_id
        and s.class_id = v_batch.advisory_class_id
    ) then
      raise exception 'A submitted learner does not belong to the selected class.';
    end if;

    insert into public.grade_request_scores (
      subject_request_id,
      student_id,
      score,
      submitted_at
    ) values (
      p_subject_request_id,
      v_student_id,
      v_numeric_score,
      now()
    );

    update public.grades
    set
      score = v_numeric_score,
      teacher_id = v_user_id,
      class_id = v_batch.advisory_class_id
    where student_id = v_student_id
      and subject = v_item.subject
      and term = v_batch.grading_period;

    if not found then
      insert into public.grades (
        student_id,
        class_id,
        teacher_id,
        subject,
        term,
        score
      ) values (
        v_student_id,
        v_batch.advisory_class_id,
        v_user_id,
        v_item.subject,
        v_batch.grading_period,
        v_numeric_score
      );
    end if;
  end loop;

  update public.grade_request_subjects
  set
    status = 'submitted',
    submitted_at = now(),
    teacher_note = nullif(trim(p_teacher_note), '')
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
    'recipient_grades_submitted',
    v_item.status,
    'submitted',
    jsonb_build_object(
      'recipient_role', v_item.recipient_role,
      'subject', v_item.subject,
      'score_count', jsonb_array_length(p_scores),
      'submitted_at', now()
    )
  );
end;
$$;

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
  recipient.full_name as subject_teacher_name,
  recipient.email as subject_teacher_email,
  s.subject,
  s.status as subject_status,
  s.requested_at as subject_requested_at,
  s.submitted_at,
  s.teacher_note,
  count(sc.id) as submitted_grade_count,
  s.message as request_message,
  s.recipient_role
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

grant execute on function public.create_grade_request_batch(
  uuid,
  text,
  uuid,
  text,
  text
) to authenticated;

grant execute on function public.submit_grade_request_subject(
  uuid,
  jsonb,
  text
) to authenticated;

grant select on public.grade_request_workflow_view to authenticated;

notify pgrst, 'reload schema';
