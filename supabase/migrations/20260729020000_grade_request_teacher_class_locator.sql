-- Allow a logged-in Class Adviser to locate the grade level, section,
-- school year, and subject assignments of a selected Class Adviser or
-- Subject Teacher from the Request Form email directory.

create or replace function public.list_grade_request_teacher_classes(
  p_teacher_id uuid
)
returns table (
  class_id uuid,
  teacher_id uuid,
  subject text,
  grade_level text,
  section text,
  school_year text
)
language plpgsql
stable
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can view the teacher class directory.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    where p.id = p_teacher_id
      and p.teacher_type in ('class_adviser', 'subject_teacher')
  ) then
    raise exception 'The selected teacher account does not exist.';
  end if;

  return query
  select
    c.id as class_id,
    c.teacher_id,
    c.subject,
    c.grade_level,
    c.section,
    c.school_year
  from public.classes c
  where c.teacher_id = p_teacher_id
  order by
    lower(coalesce(c.school_year, '')) desc,
    lower(coalesce(c.grade_level, '')),
    lower(coalesce(c.section, '')),
    lower(coalesce(c.subject, ''));
end;
$$;

revoke all on function public.list_grade_request_teacher_classes(uuid)
from public;

grant execute on function public.list_grade_request_teacher_classes(uuid)
to authenticated;

notify pgrst, 'reload schema';
