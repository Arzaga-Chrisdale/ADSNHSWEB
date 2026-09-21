-- ============================================================
-- Admin Access + SF1 Learner Management + Admin Profile Update
-- Safe to run more than once.
--
-- Includes:
-- 1) Admin read access policies
-- 2) Admin SF1 learner CRUD policies
-- 3) Nullable class_id / teacher_id for Admin-created learners
-- 4) Admin RPC for editing another teacher/user profile
-- 5) Sync classes.teacher_name after profile name changes
-- ============================================================

begin;

-- ------------------------------------------------------------
-- 1. Allow Admin-created SF1 learner records before assignment
-- ------------------------------------------------------------

alter table public.students
  alter column class_id drop not null,
  alter column teacher_id drop not null;

-- ------------------------------------------------------------
-- 2. Profiles / Users
-- ------------------------------------------------------------

drop policy if exists "admins read all profiles" on public.profiles;

create policy "admins read all profiles"
on public.profiles
for select
to authenticated
using (public.is_admin(auth.uid()));

-- ------------------------------------------------------------
-- 3. User Roles
-- ------------------------------------------------------------

drop policy if exists "admins read all user roles" on public.user_roles;

create policy "admins read all user roles"
on public.user_roles
for select
to authenticated
using (public.is_admin(auth.uid()));

drop policy if exists "users can read own role" on public.user_roles;

create policy "users can read own role"
on public.user_roles
for select
to authenticated
using (auth.uid() = user_id);

-- ------------------------------------------------------------
-- 4. Classes
-- ------------------------------------------------------------

drop policy if exists "admins read all classes" on public.classes;

create policy "admins read all classes"
on public.classes
for select
to authenticated
using (public.is_admin(auth.uid()));

-- ------------------------------------------------------------
-- 5. Students / SF1 Learner Records
-- ------------------------------------------------------------

drop policy if exists "admins read all students" on public.students;

create policy "admins read all students"
on public.students
for select
to authenticated
using (public.is_admin(auth.uid()));

drop policy if exists "admins insert students" on public.students;

create policy "admins insert students"
on public.students
for insert
to authenticated
with check (public.is_admin(auth.uid()));

drop policy if exists "admins update students" on public.students;

create policy "admins update students"
on public.students
for update
to authenticated
using (public.is_admin(auth.uid()))
with check (public.is_admin(auth.uid()));

drop policy if exists "admins delete students" on public.students;

create policy "admins delete students"
on public.students
for delete
to authenticated
using (public.is_admin(auth.uid()));

-- ------------------------------------------------------------
-- 6. Grades
-- ------------------------------------------------------------

drop policy if exists "admins read all grades" on public.grades;

create policy "admins read all grades"
on public.grades
for select
to authenticated
using (public.is_admin(auth.uid()));

-- ------------------------------------------------------------
-- 7. Table Grants
-- ------------------------------------------------------------

grant select on public.profiles to authenticated;
grant select on public.user_roles to authenticated;
grant select on public.classes to authenticated;
grant select, insert, update, delete on public.students to authenticated;
grant select on public.grades to authenticated;

-- ------------------------------------------------------------
-- 8. Admin Profile Update RPC
--
-- Admins can read other profiles, but direct client UPDATE can
-- still be blocked by profiles RLS. This SECURITY DEFINER RPC
-- performs the update after verifying the caller is an Admin.
-- ------------------------------------------------------------

create or replace function public.admin_update_profile(
  target_user_id uuid,
  new_full_name text,
  new_school_name text default null,
  new_school_id text default null,
  new_region text default null,
  new_division text default null,
  new_district text default null
)
returns public.profiles
language plpgsql
security definer
set search_path = public
as $$
declare
  updated_profile public.profiles;
begin
  if auth.uid() is null then
    raise exception 'You must be signed in.';
  end if;

  if not public.is_admin(auth.uid()) then
    raise exception 'Admin access is required.';
  end if;

  if target_user_id is null then
    raise exception 'Target user is required.';
  end if;

  if nullif(btrim(coalesce(new_full_name, '')), '') is null then
    raise exception 'Full Name is required.';
  end if;

  update public.profiles
  set
    full_name = btrim(new_full_name),
    school_name = nullif(btrim(coalesce(new_school_name, '')), ''),
    school_id = nullif(btrim(coalesce(new_school_id, '')), ''),
    region = nullif(btrim(coalesce(new_region, '')), ''),
    division = nullif(btrim(coalesce(new_division, '')), ''),
    district = nullif(btrim(coalesce(new_district, '')), '')
  where id = target_user_id
  returning * into updated_profile;

  if updated_profile.id is null then
    raise exception 'Profile not found.';
  end if;

  -- Keep the copied teacher name in classes synchronized.
  update public.classes
  set teacher_name = updated_profile.full_name
  where teacher_id = target_user_id;

  return updated_profile;
end;
$$;

-- Only signed-in users can execute it, and the function itself
-- verifies that the caller is an Admin.
revoke all on function public.admin_update_profile(
  uuid, text, text, text, text, text, text
) from public;

grant execute on function public.admin_update_profile(
  uuid, text, text, text, text, text, text
) to authenticated;

commit;

-- Reload PostgREST schema cache so the new RPC is immediately available.
notify pgrst, 'reload schema';
