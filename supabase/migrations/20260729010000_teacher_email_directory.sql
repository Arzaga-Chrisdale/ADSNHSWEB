-- Allow a Class Adviser to view the email directory for all Class Advisers
-- and Subject Teachers while keeping grade-request recipients restricted to
-- Subject Teachers assigned to the selected advisory class.

create or replace function public.list_grade_request_recipient_directory(
  p_advisory_class_id uuid
)
returns table (
  teacher_id uuid,
  full_name text,
  email text,
  teacher_type text,
  assigned_to_class boolean,
  subjects text[]
)
language plpgsql
stable
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_advisory_class public.classes%rowtype;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  select *
  into v_advisory_class
  from public.classes
  where id = p_advisory_class_id;

  if not found then
    raise exception 'The selected advisory class does not exist.';
  end if;

  if v_advisory_class.teacher_id <> v_user_id then
    raise exception 'Only the assigned Class Adviser can view this teacher directory.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can view this teacher directory.';
  end if;

  return query
  select
    p.id as teacher_id,
    p.full_name,
    p.email,
    p.teacher_type,
    case
      when p.teacher_type = 'subject_teacher' then exists (
        select 1
        from public.classes c
        where c.teacher_id = p.id
          and c.subject is not null
          and trim(c.subject) <> ''
          and lower(trim(coalesce(c.grade_level, ''))) =
              lower(trim(coalesce(v_advisory_class.grade_level, '')))
          and lower(trim(coalesce(c.section, ''))) =
              lower(trim(coalesce(v_advisory_class.section, '')))
          and lower(trim(coalesce(c.school_year, ''))) =
              lower(trim(coalesce(v_advisory_class.school_year, '')))
      )
      else false
    end as assigned_to_class,
    case
      when p.teacher_type = 'subject_teacher' then coalesce(
        array(
          select distinct trim(c.subject)
          from public.classes c
          where c.teacher_id = p.id
            and c.subject is not null
            and trim(c.subject) <> ''
            and lower(trim(coalesce(c.grade_level, ''))) =
                lower(trim(coalesce(v_advisory_class.grade_level, '')))
            and lower(trim(coalesce(c.section, ''))) =
                lower(trim(coalesce(v_advisory_class.section, '')))
            and lower(trim(coalesce(c.school_year, ''))) =
                lower(trim(coalesce(v_advisory_class.school_year, '')))
          order by trim(c.subject)
        ),
        array[]::text[]
      )
      else array[]::text[]
    end as subjects
  from public.profiles p
  where p.teacher_type in ('class_adviser', 'subject_teacher')
    and p.email is not null
    and trim(p.email) <> ''
  order by
    case when p.teacher_type = 'class_adviser' then 0 else 1 end,
    lower(p.email);
end;
$$;

revoke all on function public.list_grade_request_recipient_directory(uuid)
from public;

grant execute on function public.list_grade_request_recipient_directory(uuid)
to authenticated;

notify pgrst, 'reload schema';
