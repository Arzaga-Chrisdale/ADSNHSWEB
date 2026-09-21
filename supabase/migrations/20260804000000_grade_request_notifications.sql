-- Grade-request notifications for Class Advisers and Subject Teachers.
--
-- Run after:
--   1. 20260728000000_grade_request_submission_workflow.sql
--   2. 20260803000000_notifications.sql
--
-- Safe to run more than once.
--
-- Notification flow:
--   1. New Grade Request -> recipient is notified
--   2. Completed Request -> Class Adviser is notified
--   3. Rejected Request -> Class Adviser is notified
--   4. Finalized Grades -> recipient is notified
--
-- Compatibility:
--   database status "submitted" = UI status "Completed"
--   database status "rejected"  = UI status "Rejected"


create or replace function public.notify_grade_request_subject_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_batch public.grade_request_batches%rowtype;
  v_class public.classes%rowtype;
  v_adviser_name text;
  v_recipient_name text;
  v_class_label text;
  v_period_label text;
  v_note text;
begin
  select *
  into v_batch
  from public.grade_request_batches
  where id = new.batch_id;

  if not found then
    return new;
  end if;

  select *
  into v_class
  from public.classes
  where id = v_batch.advisory_class_id;

  select coalesce(nullif(trim(full_name), ''), 'A Class Adviser')
  into v_adviser_name
  from public.profiles
  where id = v_batch.adviser_id;

  select coalesce(nullif(trim(full_name), ''), 'The recipient')
  into v_recipient_name
  from public.profiles
  where id = new.subject_teacher_id;

  v_adviser_name := coalesce(v_adviser_name, 'A Class Adviser');
  v_recipient_name := coalesce(v_recipient_name, 'The recipient');

  v_class_label := concat_ws(
    ' - ',
    nullif(trim(v_class.grade_level), ''),
    nullif(trim(v_class.section), '')
  );

  v_class_label := coalesce(
    nullif(v_class_label, ''),
    'the advisory class'
  );

  v_period_label := case v_batch.grading_period
    when '1' then 'Term 1'
    when '2' then 'Term 2'
    when '3' then 'Term 3'
    when 'final' then 'Final Grade'
    else coalesce(v_batch.grading_period, 'the selected grading period')
  end;

  v_note := nullif(trim(new.teacher_note), '');

  -- ==========================================================
  -- 1. NEW REQUEST
  -- Notify the selected Subject Teacher or Class Adviser.
  -- ==========================================================
  if tg_op = 'INSERT' then
    perform public.create_notification(
      new.subject_teacher_id,
      'New Grade Request',
      concat(
        v_adviser_name,
        ' requested ',
        new.subject,
        ' grades for ',
        v_class_label,
        ', ',
        v_period_label,
        '.'
      ),
      'grade_submission',
      '/received-requests',
      v_batch.adviser_id,
      concat(
        'grade-request-subject:',
        new.id,
        ':created'
      ),
      jsonb_build_object(
        'batch_id', new.batch_id,
        'subject_request_id', new.id,
        'advisory_class_id', v_batch.advisory_class_id,
        'subject_class_id', new.subject_class_id,
        'subject', new.subject,
        'grading_period', v_batch.grading_period,
        'status', new.status
      )
    );

    return new;
  end if;

  -- For UPDATE events, do nothing unless the status changed.
  if old.status is not distinct from new.status then
    return new;
  end if;

  -- ==========================================================
  -- 2. COMPLETED REQUEST
  --
  -- The existing database stores a completed request as
  -- "submitted". The React UI displays this as "Completed".
  --
  -- Notify the Class Adviser and open Request Form.
  -- ==========================================================
  if new.status = 'submitted' then
    perform public.create_notification(
      v_batch.adviser_id,
      'Grade Request Completed',
      concat(
        v_recipient_name,
        ' completed the ',
        new.subject,
        ' Grade Request for ',
        v_class_label,
        ', ',
        v_period_label,
        '.',
        case
          when v_note is not null
            then concat(' Message: ', v_note)
          else ''
        end
      ),
      'grade_review',
      '/request-form',
      new.subject_teacher_id,
      concat(
        'grade-request-subject:',
        new.id,
        ':completed:',
        extract(
          epoch from coalesce(new.submitted_at, now())
        )
      ),
      jsonb_build_object(
        'batch_id', new.batch_id,
        'subject_request_id', new.id,
        'advisory_class_id', v_batch.advisory_class_id,
        'subject_class_id', new.subject_class_id,
        'subject', new.subject,
        'grading_period', v_batch.grading_period,
        'status', 'completed',
        'database_status', new.status,
        'teacher_note', v_note
      )
    );

    return new;
  end if;

  -- ==========================================================
  -- 3. REJECTED REQUEST
  --
  -- Notify the Class Adviser and keep the rejected request in
  -- Previous Requests.
  -- ==========================================================
  if new.status = 'rejected' then
    perform public.create_notification(
      v_batch.adviser_id,
      'Grade Request Rejected',
      concat(
        v_recipient_name,
        ' rejected the ',
        new.subject,
        ' Grade Request for ',
        v_class_label,
        ', ',
        v_period_label,
        '.',
        case
          when v_note is not null
            then concat(' Message: ', v_note)
          else ''
        end
      ),
      'grade_review',
      '/request-form',
      new.subject_teacher_id,
      concat(
        'grade-request-subject:',
        new.id,
        ':rejected:',
        extract(epoch from now())
      ),
      jsonb_build_object(
        'batch_id', new.batch_id,
        'subject_request_id', new.id,
        'advisory_class_id', v_batch.advisory_class_id,
        'subject_class_id', new.subject_class_id,
        'subject', new.subject,
        'grading_period', v_batch.grading_period,
        'status', 'rejected',
        'teacher_note', v_note
      )
    );

    return new;
  end if;

  return new;
end;
$$;


revoke all on function public.notify_grade_request_subject_change()
  from public, anon, authenticated;


drop trigger if exists grade_request_subject_notifications
  on public.grade_request_subjects;

create trigger grade_request_subject_notifications
after insert or update of status on public.grade_request_subjects
for each row
execute function public.notify_grade_request_subject_change();


-- ============================================================
-- 4. FINALIZED GRADES
--
-- Keep the previous finalization notification behavior.
-- When the Class Adviser finalizes a batch, notify every
-- recipient included in the batch.
-- ============================================================

create or replace function public.notify_grade_request_batch_finalized()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_subject_request record;
  v_class public.classes%rowtype;
  v_adviser_name text;
  v_class_label text;
  v_period_label text;
begin
  if new.is_finalized is not true
    or old.is_finalized is not distinct from new.is_finalized
  then
    return new;
  end if;

  select *
  into v_class
  from public.classes
  where id = new.advisory_class_id;

  select coalesce(nullif(trim(full_name), ''), 'The Class Adviser')
  into v_adviser_name
  from public.profiles
  where id = new.adviser_id;

  v_adviser_name := coalesce(
    v_adviser_name,
    'The Class Adviser'
  );

  v_class_label := concat_ws(
    ' - ',
    nullif(trim(v_class.grade_level), ''),
    nullif(trim(v_class.section), '')
  );

  v_class_label := coalesce(
    nullif(v_class_label, ''),
    'the advisory class'
  );

  v_period_label := case new.grading_period
    when '1' then 'Term 1'
    when '2' then 'Term 2'
    when '3' then 'Term 3'
    when 'final' then 'Final Grade'
    else coalesce(
      new.grading_period,
      'the selected grading period'
    )
  end;

  for v_subject_request in
    select
      s.id,
      s.subject_teacher_id,
      s.subject
    from public.grade_request_subjects s
    where s.batch_id = new.id
  loop
    perform public.create_notification(
      v_subject_request.subject_teacher_id,
      'Grades Finalized',
      concat(
        v_adviser_name,
        ' finalized your ',
        v_subject_request.subject,
        ' grades for ',
        v_class_label,
        ', ',
        v_period_label,
        '.'
      ),
      'grade_review',
      '/received-requests',
      new.adviser_id,
      concat(
        'grade-request-batch:',
        new.id,
        ':subject:',
        v_subject_request.id,
        ':finalized'
      ),
      jsonb_build_object(
        'batch_id', new.id,
        'subject_request_id', v_subject_request.id,
        'advisory_class_id', new.advisory_class_id,
        'subject', v_subject_request.subject,
        'grading_period', new.grading_period,
        'is_finalized', true
      )
    );
  end loop;

  return new;
end;
$$;


revoke all on function public.notify_grade_request_batch_finalized()
  from public, anon, authenticated;


drop trigger if exists grade_request_batch_finalized_notifications
  on public.grade_request_batches;

create trigger grade_request_batch_finalized_notifications
after update of is_finalized on public.grade_request_batches
for each row
execute function public.notify_grade_request_batch_finalized();


notify pgrst, 'reload schema';
