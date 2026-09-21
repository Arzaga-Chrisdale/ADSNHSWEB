-- Allow authenticated users to load Subject Teacher accounts
-- for the Request Form teacher selector.

alter table public.profiles enable row level security;

drop policy if exists "authenticated read subject teacher profiles"
on public.profiles;

create policy "authenticated read subject teacher profiles"
on public.profiles
for select
to authenticated
using (
  teacher_type = 'subject_teacher'
  or id = auth.uid()
  or public.is_admin(auth.uid())
);

grant select
on public.profiles
to authenticated;