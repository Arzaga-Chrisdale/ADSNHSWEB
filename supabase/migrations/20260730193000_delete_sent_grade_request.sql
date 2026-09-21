-- Permanently delete one mistakenly sent Grade Request.
-- Only the Class Adviser who sent the request can delete it.
-- Submitted requests are protected and cannot be deleted here.

create or replace function public.delete_sent_grade_request(
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
  select request.batch_id
    into v_batch_id
  from public.grade_request_subjects as request
  join public.grade_request_batches as batch
    on batch.id = request.batch_id
  where request.id = p_subject_request_id
    and batch.adviser_id = auth.uid()
    and batch.recipient_role = 'subject_teacher'
    and request.status = 'pending';

  if v_batch_id is null then
    raise exception
      'Pending Grade Request not found, already submitted, or you do not have permission to delete it.';
  end if;

  delete from public.grade_request_subjects
  where id = p_subject_request_id
    and status = 'pending';

  get diagnostics v_deleted_count = row_count;

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

revoke all on function public.delete_sent_grade_request(uuid) from public;
grant execute on function public.delete_sent_grade_request(uuid)
  to authenticated;
