-- Permanently delete one assigned grade request as its assigned Subject Teacher.
-- Related submitted scores and history are removed before the request itself.

create or replace function public.delete_assigned_grade_request(
  p_subject_request_id uuid
)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  v_batch_id uuid;
  v_deleted_count integer;
begin
  select subject.batch_id
    into v_batch_id
  from public.grade_request_subjects as subject
  where subject.id = p_subject_request_id
    and subject.subject_teacher_id = auth.uid();

  if v_batch_id is null then
    raise exception
      'Assigned grade request not found or you do not have permission to delete it.';
  end if;

  delete from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  delete from public.grade_request_history
  where subject_request_id = p_subject_request_id;

  delete from public.grade_request_subjects
  where id = p_subject_request_id
    and subject_teacher_id = auth.uid();

  get diagnostics v_deleted_count = row_count;

  -- Remove the parent batch only when it no longer has subject requests.
  delete from public.grade_request_batches as batch
  where batch.id = v_batch_id
    and not exists (
      select 1
      from public.grade_request_subjects as remaining
      where remaining.batch_id = batch.id
    );

  return v_deleted_count = 1;
end;
$$;

revoke all on function public.delete_assigned_grade_request(uuid) from public;
grant execute on function public.delete_assigned_grade_request(uuid)
to authenticated;
