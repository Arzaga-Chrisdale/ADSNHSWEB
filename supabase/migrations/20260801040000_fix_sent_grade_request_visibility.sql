-- Return the signed-in Class Adviser's sent Grade Requests without relying on
-- view joins that can be hidden by RLS when the selected class belongs to the
-- recipient. The function is strictly limited to auth.uid().

create or replace function public.list_my_sent_grade_request_workflow()
returns table (
  batch_id uuid,
  adviser_id uuid,
  adviser_name text,
  adviser_email text,
  advisory_class_id uuid,
  grade_level text,
  section text,
  school_year text,
  grading_period text,
  overall_status text,
  is_finalized boolean,
  requested_at timestamptz,
  completed_at timestamptz,
  finalized_at timestamptz,
  subject_request_id uuid,
  subject_class_id uuid,
  subject_teacher_id uuid,
  subject_teacher_name text,
  subject_teacher_email text,
  subject text,
  subject_status text,
  subject_requested_at timestamptz,
  submitted_at timestamptz,
  teacher_note text,
  submitted_grade_count integer,
  request_message text,
  recipient_role text,
  recipient_id uuid
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
    raise exception 'Only a Class Adviser can view sent Grade Requests.';
  end if;

  return query
  select
    batch.id,
    batch.adviser_id,
    adviser.full_name,
    adviser.email,
    batch.advisory_class_id,
    selected_class.grade_level,
    selected_class.section,
    selected_class.school_year,
    batch.grading_period::text,
    batch.status::text,
    batch.is_finalized,
    batch.requested_at,
    batch.completed_at,
    batch.finalized_at,
    request.id,
    request.subject_class_id,
    request.subject_teacher_id,
    recipient.full_name,
    recipient.email,
    request.subject,
    request.status::text,
    request.requested_at,
    request.submitted_at,
    request.teacher_note,
    count(score.id)::integer,
    request.message,
    request.recipient_role,
    request.subject_teacher_id
  from public.grade_request_batches batch
  join public.classes selected_class on selected_class.id = batch.advisory_class_id
  join public.profiles adviser on adviser.id = batch.adviser_id
  join public.grade_request_subjects request on request.batch_id = batch.id
  join public.profiles recipient on recipient.id = request.subject_teacher_id
  left join public.grade_request_scores score on score.subject_request_id = request.id
  where batch.adviser_id = v_user_id
  group by
    batch.id,
    adviser.full_name,
    adviser.email,
    selected_class.grade_level,
    selected_class.section,
    selected_class.school_year,
    request.id,
    recipient.full_name,
    recipient.email
  order by batch.requested_at desc, request.requested_at desc;
end;
$$;

revoke all on function public.list_my_sent_grade_request_workflow()
from public;

grant execute on function public.list_my_sent_grade_request_workflow()
to authenticated;

notify pgrst, 'reload schema';
