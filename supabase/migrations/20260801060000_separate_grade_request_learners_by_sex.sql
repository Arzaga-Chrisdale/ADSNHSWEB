-- Return the learner sex with submitted grades so the requesting Class Adviser
-- can display separate MALE and FEMALE sections in the eye-icon review dialog.
create or replace function public.list_my_sent_grade_request_review_by_sex(
  p_subject_request_id uuid
)
returns table (
  id uuid,
  first_name text,
  middle_name text,
  last_name text,
  lrn text,
  sex text,
  score numeric
)
language sql
security definer
set search_path = public
stable
as $$
  select
    student.id,
    student.first_name,
    student.middle_name,
    student.last_name,
    student.lrn,
    student.sex,
    submitted_score.score
  from public.grade_request_subjects subject_request
  join public.grade_request_batches batch
    on batch.id = subject_request.batch_id
  join public.grade_request_scores submitted_score
    on submitted_score.subject_request_id = subject_request.id
  join public.students student
    on student.id = submitted_score.student_id
  where subject_request.id = p_subject_request_id
    and batch.adviser_id = auth.uid()
  order by
    case student.sex when 'male' then 1 when 'female' then 2 else 3 end,
    student.last_name,
    student.first_name,
    student.middle_name;
$$;

revoke all on function public.list_my_sent_grade_request_review_by_sex(uuid) from public;
grant execute on function public.list_my_sent_grade_request_review_by_sex(uuid) to authenticated;
