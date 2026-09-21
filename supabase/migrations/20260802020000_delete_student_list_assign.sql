-- Permanently remove a sent student-list assignment and its accepted copy.
-- Either participant may remove it. Deleting the assignment releases the
-- unique (class, teacher, subject) combination so the adviser can send again.

create or replace function public.delete_student_list_assignment(
  p_assignment_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_assignment public.student_list_assignments%rowtype;
  v_subject_class_id uuid;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to delete this assignment.';
  end if;

  select * into v_assignment
  from public.student_list_assignments
  where id = p_assignment_id
  for update;

  if not found then
    raise exception 'Student-list assignment not found.';
  end if;

  if v_user_id <> v_assignment.subject_teacher_id
     and v_user_id <> v_assignment.adviser_id then
    raise exception 'Only the Class Adviser or assigned Subject Teacher can delete this assignment.';
  end if;

  select id into v_subject_class_id
  from public.classes
  where student_list_assignment_id = v_assignment.id;

  -- Existing class foreign keys remove the copied learners, grades, and class
  -- records. The assignment must be deleted second because this link is
  -- intentionally ON DELETE RESTRICT.
  if v_subject_class_id is not null then
    delete from public.classes
    where id = v_subject_class_id
      and teacher_id = v_assignment.subject_teacher_id;

    if not found then
      raise exception 'The copied Subject Teacher class could not be deleted.';
    end if;
  end if;

  delete from public.student_list_assignments
  where id = v_assignment.id;

  return jsonb_build_object(
    'assignment_id', v_assignment.id,
    'deleted_class_id', v_subject_class_id,
    'deleted', true,
    'can_send_again', true
  );
end;
$$;

revoke all on function public.delete_student_list_assignment(uuid) from public;
grant execute on function public.delete_student_list_assignment(uuid)
  to authenticated;

notify pgrst, 'reload schema';
