alter table public.profiles enable row level security;

drop policy if exists "teachers read request recipient directory"
on public.profiles;

create policy "teachers read request recipient directory"
on public.profiles
for select
to authenticated
using (
  teacher_type in ('class_adviser', 'subject_teacher')
);