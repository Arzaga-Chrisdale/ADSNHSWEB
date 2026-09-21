-- Permanently delete one adviser-to-adviser submission from the recipient's
-- Receive Form. Only the Class Adviser named as the recipient (or an admin)
-- can delete it. Related scores and history are removed by cascade rules.

create or replace function public.delete_received_class_adviser_submission(
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
    and batch.recipient_role = 'class_adviser'
    and (
      batch.recipient_id = auth.uid()
      or public.is_admin(auth.uid())
    );

  if v_batch_id is null then
    raise exception
      'Class Adviser submission not found or you do not have permission to delete it.';
  end if;

  delete from public.grade_request_subjects
  where id = p_subject_request_id;

  get diagnostics v_deleted_count = row_count;

  -- Remove the parent only when no other submission remains in the batch.
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

revoke all
  on function public.delete_received_class_adviser_submission(uuid)
  from public;

grant execute
  on function public.delete_received_class_adviser_submission(uuid)
  to authenticated;
