-- =========================================================
-- Grade Request Delete Policies
-- Subject Teacher: delete requests they sent
-- Class Adviser: delete requests assigned to them
-- =========================================================

alter table public.grade_requests enable row level security;

-- Remove older delete policies before recreating them
drop policy if exists "requesters delete own grade requests"
on public.grade_requests;

drop policy if exists "class advisers delete assigned grade requests"
on public.grade_requests;

-- Subject Teacher can delete only requests they created
create policy "requesters delete own grade requests"
on public.grade_requests
for delete
to authenticated
using (
  requester_id = auth.uid()
);

-- Class Adviser can delete only requests assigned to their account
create policy "class advisers delete assigned grade requests"
on public.grade_requests
for delete
to authenticated
using (
  subject_teacher_id = auth.uid()
  and exists (
    select 1
    from public.profiles as profile
    where profile.id = auth.uid()
      and profile.teacher_type = 'class_adviser'
  )
);

grant delete
on public.grade_requests
to authenticated;