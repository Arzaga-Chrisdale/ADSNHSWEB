-- Allow an assigned Subject Teacher to read the profile
-- of the Class Adviser who sent the request.

alter table public.profiles enable row level security;

drop policy if exists "subject teachers read request requester profiles"
on public.profiles;

create policy "subject teachers read request requester profiles"
on public.profiles
for select
to authenticated
using (
  exists (
    select 1
    from public.grade_requests as request
    where request.subject_teacher_id = auth.uid()
      and request.requester_id = profiles.id
  )
);