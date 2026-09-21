-- Consolidate the credential workflow into Pending Review, Approved, and Rejected.
-- After Admin approval, the requesting Class Adviser securely assigns the
-- complete SF1 learner record to one of their own classes.
-- Safe to run more than once.

begin;

alter table public.student_credential_requests
  drop constraint if exists student_credential_requests_status_check;

update public.student_credential_requests
set status = case lower(trim(coalesce(status, '')))
  when 'pending' then 'pending_review'
  when 'pending review' then 'pending_review'
  when 'pending_review' then 'pending_review'
  when 'completed' then 'approved'
  when 'released' then 'approved'
  when 'approved' then 'approved'
  when 'rejected' then 'rejected'
  else 'pending_review'
end;

alter table public.student_credential_requests
  alter column status set default 'pending_review';

alter table public.student_credential_requests
  add constraint student_credential_requests_status_check
  check (status in ('pending_review', 'approved', 'rejected'));

-- The requesting adviser may read the Admin-verified learner before class
-- assignment so the correct SF1 name is shown instead of the typed request.
drop policy if exists "requesters read approved credential students"
  on public.students;

create policy "requesters read approved credential students"
on public.students
for select
to authenticated
using (
  exists (
    select 1
    from public.student_credential_requests request_record
    where request_record.student_id = students.id
      and request_record.requester_id = auth.uid()
      and lower(trim(request_record.status)) in ('approved', 'released', 'completed')
  )
);

create or replace function public.assign_approved_credential_student(
  p_request_id uuid,
  p_class_id uuid
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  credential_request public.student_credential_requests%rowtype;
  target_class public.classes%rowtype;
  v_previous_class_id uuid;
begin
  if auth.uid() is null then
    raise exception 'You must be signed in.';
  end if;

  select request_record.*
  into credential_request
  from public.student_credential_requests request_record
  where request_record.id = p_request_id
  for update;

  if not found then
    raise exception 'Credential request not found.';
  end if;

  if credential_request.requester_id <> auth.uid() then
    raise exception 'Only the requesting Class Adviser can assign this student.';
  end if;

  if lower(trim(credential_request.status)) not in ('approved', 'released', 'completed') then
    raise exception 'Admin must approve this credential request first.';
  end if;

  if credential_request.student_id is null then
    raise exception 'Admin has not connected this request to an SF1 learner record.';
  end if;

  select class_record.*
  into target_class
  from public.classes class_record
  where class_record.id = p_class_id;

  if not found then
    raise exception 'Selected class not found.';
  end if;

  if target_class.teacher_id <> auth.uid() then
    raise exception 'You can only add the student to one of your own classes.';
  end if;

  select student.class_id
  into v_previous_class_id
  from public.students student
  where student.id = credential_request.student_id
  for update;

  if not found then
    raise exception 'The approved SF1 learner record no longer exists.';
  end if;

  update public.students
  set class_id = target_class.id,
      teacher_id = auth.uid()
  where id = credential_request.student_id;

  update public.student_credential_requests
  set status = 'approved',
      requesting_class_id = target_class.id,
      previous_class_id = coalesce(
        credential_request.previous_class_id,
        v_previous_class_id
      ),
      updated_at = now()
  where id = credential_request.id;

  return credential_request.student_id;
end;
$$;

revoke all on function public.assign_approved_credential_student(uuid, uuid)
  from public, anon;

grant execute on function public.assign_approved_credential_student(uuid, uuid)
  to authenticated, service_role;

commit;

notify pgrst, 'reload schema';
