-- Let an assigned Subject Teacher edit the copied class information after
-- accepting a student list. The Class Adviser's original class is untouched.

create or replace function public.update_accepted_subject_class_details(
  p_class_id uuid,
  p_school_name text,
  p_school_id text,
  p_school_year text,
  p_region text,
  p_division text,
  p_district text,
  p_grade_level text,
  p_section text,
  p_start_date date,
  p_end_date date,
  p_subject text
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_assignment_id uuid;
  v_old_subject text;
  v_teacher_name text;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to update this class.';
  end if;

  if nullif(btrim(p_grade_level), '') is null then
    raise exception 'Grade Level is required.';
  end if;
  if nullif(btrim(p_section), '') is null then
    raise exception 'Section is required.';
  end if;
  if nullif(btrim(p_subject), '') is null then
    raise exception 'Subject is required.';
  end if;
  if p_start_date is not null and p_end_date is not null
     and p_end_date < p_start_date then
    raise exception 'End Date cannot be earlier than Start Date.';
  end if;

  select c.student_list_assignment_id, c.subject,
         coalesce(nullif(btrim(p.full_name), ''), p.email, c.teacher_name)
  into v_assignment_id, v_old_subject, v_teacher_name
  from public.classes c
  join public.student_list_assignments a
    on a.id = c.student_list_assignment_id
  left join public.profiles p on p.id = v_user_id
  where c.id = p_class_id
    and c.teacher_id = v_user_id
    and a.subject_teacher_id = v_user_id
    and a.status = 'accepted'
  for update of c, a;

  if not found then
    raise exception 'Only the assigned Subject Teacher can update this accepted class.';
  end if;

  update public.classes
  set school_name = nullif(btrim(p_school_name), ''),
      school_id = nullif(btrim(p_school_id), ''),
      school_year = nullif(btrim(p_school_year), ''),
      region = nullif(btrim(p_region), ''),
      division = nullif(btrim(p_division), ''),
      district = nullif(btrim(p_district), ''),
      grade_level = btrim(p_grade_level),
      section = btrim(p_section),
      start_date = p_start_date,
      end_date = p_end_date,
      subject = btrim(p_subject),
      teacher_name = v_teacher_name
  where id = p_class_id;

  update public.student_list_assignments
  set subject = btrim(p_subject), updated_at = now()
  where id = v_assignment_id;

  if v_old_subject is distinct from btrim(p_subject) then
    update public.grades
    set subject = btrim(p_subject)
    where class_id = p_class_id
      and teacher_id = v_user_id
      and subject = v_old_subject;
  end if;

  return jsonb_build_object(
    'class_id', p_class_id,
    'grade_level', btrim(p_grade_level),
    'section', btrim(p_section),
    'subject', btrim(p_subject)
  );
end;
$$;

revoke all on function public.update_accepted_subject_class_details(
  uuid, text, text, text, text, text, text, text, text, date, date, text
) from public;
grant execute on function public.update_accepted_subject_class_details(
  uuid, text, text, text, text, text, text, text, text, date, date, text
) to authenticated;

notify pgrst, 'reload schema';
