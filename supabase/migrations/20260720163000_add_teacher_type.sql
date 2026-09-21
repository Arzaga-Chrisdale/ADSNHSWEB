alter table public.profiles
add column if not exists teacher_type text;

alter table public.profiles
drop constraint if exists profiles_teacher_type_check;

alter table public.profiles
add constraint profiles_teacher_type_check
check (
  teacher_type is null
  or teacher_type in ('class_adviser', 'subject_teacher')
);

-- Keep existing teacher accounts on the Class Adviser dashboard.
update public.profiles as profile
set teacher_type = 'class_adviser'
where profile.teacher_type is null
  and exists (
    select 1
    from public.user_roles as user_role
    where user_role.user_id = profile.id
      and user_role.role = 'teacher'
  );