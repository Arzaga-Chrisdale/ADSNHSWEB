-- Only a Class Adviser may send an official advisory-class student list.
-- This supplements the UI role check with a database-level authorization check.

create or replace function public.enforce_class_adviser_student_list_sender()
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if auth.uid() is null then
    raise exception 'You must be signed in to send students.';
  end if;

  if not exists (
    select 1
    from public.profiles profile
    where profile.id = auth.uid()
      and profile.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can send an official student list.';
  end if;
end;
$$;

-- Add the role check to the existing send RPC without changing the TSX call.
-- Rename the current implementation once, then expose a guarded wrapper using
-- the same public function name and parameters expected by the application.
do $$
begin
  if to_regprocedure(
    'public.send_student_list_assignment_implementation(uuid,text,uuid)'
  ) is null then
    alter function public.send_student_list_assignment(uuid, text, uuid)
      rename to send_student_list_assignment_implementation;
  end if;
end;
$$;

create or replace function public.send_student_list_assignment(
  p_class_id uuid,
  p_subject text,
  p_subject_teacher_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
begin
  perform public.enforce_class_adviser_student_list_sender();

  return public.send_student_list_assignment_implementation(
    p_class_id,
    p_subject,
    p_subject_teacher_id
  );
end;
$$;

revoke all on function public.enforce_class_adviser_student_list_sender()
  from public;
revoke all on function public.send_student_list_assignment_implementation(uuid, text, uuid)
  from public;
revoke all on function public.send_student_list_assignment(uuid, text, uuid)
  from public;

grant execute on function public.enforce_class_adviser_student_list_sender()
  to authenticated;
grant execute on function public.send_student_list_assignment(uuid, text, uuid)
  to authenticated;

notify pgrst, 'reload schema';
