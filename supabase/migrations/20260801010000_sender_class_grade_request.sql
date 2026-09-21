-- Grade Requests are based on the sending Class Adviser's selected Grade
-- Level and Section. A selected recipient does not need to own a matching
-- class. This migration replaces only the request-creation function and does
-- not delete existing requests, grades, students, classes, or history.

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
      -- Use the Subject Teacher's distinct subject assignment(s), but do not
      -- require the teacher to own the sender's Grade Level or Section.
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
      order by lower(trim(c.subject)), c.created_at desc, c.id
      on conflict (batch_id, subject_teacher_id, subject) do nothing;

      get diagnostics v_inserted = row_count;

      -- A Subject Teacher without a class assignment can still receive one
      -- generic request for the sender's selected Grade Level and Section.
      if v_inserted = 0 and not exists (
        select 1
        from public.grade_request_subjects s
        where s.batch_id = v_batch_id
          and s.subject_teacher_id = v_recipient_id
      ) then
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
          'subject_teacher',
          'Subject Grades',
          'pending',
          nullif(trim(p_message), '')
        )
        on conflict (batch_id, subject_teacher_id, subject) do nothing;

        get diagnostics v_inserted = row_count;
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

revoke all on function public.create_multi_recipient_grade_request(uuid, text, jsonb, text)
from public;

grant execute on function public.create_multi_recipient_grade_request(uuid, text, jsonb, text)
to authenticated;

notify pgrst, 'reload schema';
