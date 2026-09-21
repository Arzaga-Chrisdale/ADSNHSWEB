-- Allow Class Advisers to delete only their own credential requests.
-- Allow Admin users to delete any credential request.

alter table public.student_credential_requests enable row level security;

drop policy if exists "requesters delete own credential requests"
  on public.student_credential_requests;

create policy "requesters delete own credential requests"
on public.student_credential_requests
for delete
to authenticated
using (requester_id = auth.uid());

drop policy if exists "admins delete credential requests"
  on public.student_credential_requests;

create policy "admins delete credential requests"
on public.student_credential_requests
for delete
to authenticated
using (public.is_admin(auth.uid()));