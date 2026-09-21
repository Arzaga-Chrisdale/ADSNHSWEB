-- Send an advisory class student list to an assigned Subject Teacher.
-- The Subject Teacher must accept before an assigned-subject class is created.

create table if not exists public.student_list_assignments (
  id uuid primary key default gen_random_uuid(),
  class_id uuid not null references public.classes(id) on delete cascade,
  adviser_id uuid not null references auth.users(id) on delete cascade,
  subject_teacher_id uuid not null references auth.users(id) on delete cascade,
  subject text not null check (btrim(subject) <> ''),
  status text not null default 'pending'
    check (status in ('pending', 'accepted', 'declined')),
  decline_reason text,
  sent_at timestamptz not null default now(),
  responded_at timestamptz,
  accepted_at timestamptz,
  declined_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (class_id, subject_teacher_id, subject)
);

create index if not exists student_list_assignments_adviser_idx
  on public.student_list_assignments (adviser_id, sent_at desc);
create index if not exists student_list_assignments_teacher_idx
  on public.student_list_assignments (subject_teacher_id, status, sent_at desc);

create table if not exists public.student_list_assignment_learners (
  id uuid primary key default gen_random_uuid(),
  assignment_id uuid not null
    references public.student_list_assignments(id) on delete cascade,
  source_student_id uuid references public.students(id) on delete set null,
  lrn text,
  first_name text not null,
  middle_name text,
  last_name text not null,
  suffix text,
  sex text check (sex in ('male', 'female')),
  birthdate date,
  address text,
  mother_name text,
  father_name text,
  guardian text,
  contact_number text,
  list_position integer not null default 0,
  created_at timestamptz not null default now(),
  unique (assignment_id, source_student_id)
);

create index if not exists student_list_assignment_learners_assignment_idx
  on public.student_list_assignment_learners (assignment_id, sex, list_position);

-- The accepted class is still a normal row in public.classes, so the existing
-- Subject Teacher dashboard can display it in My Classes without UI changes.
alter table public.classes
  add column if not exists student_list_assignment_id uuid
  references public.student_list_assignments(id) on delete restrict;

create unique index if not exists classes_student_list_assignment_uidx
  on public.classes (student_list_assignment_id)
  where student_list_assignment_id is not null;

alter table public.student_list_assignments enable row level security;
alter table public.student_list_assignment_learners enable row level security;

drop policy if exists "assignment participants read assignments"
  on public.student_list_assignments;
create policy "assignment participants read assignments"
on public.student_list_assignments for select to authenticated
using (adviser_id = auth.uid() or subject_teacher_id = auth.uid());

drop policy if exists "assignment participants read learner snapshots"
  on public.student_list_assignment_learners;
create policy "assignment participants read learner snapshots"
on public.student_list_assignment_learners for select to authenticated
using (
  exists (
    select 1
    from public.student_list_assignments assignment
    where assignment.id = assignment_id
      and (
        assignment.adviser_id = auth.uid()
        or assignment.subject_teacher_id = auth.uid()
      )
  )
);

grant select on public.student_list_assignments to authenticated;
grant select on public.student_list_assignment_learners to authenticated;

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
declare
  v_user_id uuid := auth.uid();
  v_assignment_id uuid;
  v_status text;
  v_learner_count integer := 0;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to send students.';
  end if;

  if nullif(btrim(p_subject), '') is null then
    raise exception 'Select a subject before sending the student list.';
  end if;

  if p_subject_teacher_id is null or p_subject_teacher_id = v_user_id then
    raise exception 'Select a valid Subject Teacher.';
  end if;

  perform 1
  from public.classes class_row
  where class_row.id = p_class_id
    and class_row.teacher_id = v_user_id
    and class_row.student_list_assignment_id is null;

  if not found then
    raise exception 'Only the Class Adviser who owns this class can send its student list.';
  end if;

  perform 1
  from public.profiles profile
  where profile.id = p_subject_teacher_id
    and coalesce(profile.teacher_type, '') = 'subject_teacher';

  if not found then
    raise exception 'The selected recipient is not a Subject Teacher.';
  end if;

  select assignment.id, assignment.status
  into v_assignment_id, v_status
  from public.student_list_assignments assignment
  where assignment.class_id = p_class_id
    and assignment.subject_teacher_id = p_subject_teacher_id
    and assignment.subject = btrim(p_subject)
  for update;

  if found and v_status = 'accepted' then
    raise exception 'This Subject Teacher has already accepted these students for this subject.';
  end if;

  if v_assignment_id is null then
    insert into public.student_list_assignments (
      class_id, adviser_id, subject_teacher_id, subject
    ) values (
      p_class_id, v_user_id, p_subject_teacher_id, btrim(p_subject)
    )
    returning id into v_assignment_id;
  else
    update public.student_list_assignments
    set status = 'pending',
        decline_reason = null,
        sent_at = now(),
        responded_at = null,
        accepted_at = null,
        declined_at = null,
        updated_at = now()
    where id = v_assignment_id;

    delete from public.student_list_assignment_learners
    where assignment_id = v_assignment_id;
  end if;

  insert into public.student_list_assignment_learners (
    assignment_id,
    source_student_id,
    lrn,
    first_name,
    middle_name,
    last_name,
    suffix,
    sex,
    birthdate,
    address,
    mother_name,
    father_name,
    guardian,
    contact_number,
    list_position
  )
  select
    v_assignment_id,
    student.id,
    student.lrn,
    student.first_name,
    student.middle_name,
    student.last_name,
    null,
    student.sex,
    student.birthdate,
    student.address,
    student.mother_name,
    student.father_name,
    student.guardian,
    student.contact_number,
    row_number() over (
      order by
        case student.sex when 'male' then 0 when 'female' then 1 else 2 end,
        lower(student.last_name), lower(student.first_name), student.id
    )::integer
  from public.students student
  where student.class_id = p_class_id
    and student.teacher_id = v_user_id;

  get diagnostics v_learner_count = row_count;

  if v_learner_count = 0 then
    raise exception 'Add at least one student before sending the list.';
  end if;

  return jsonb_build_object(
    'assignment_id', v_assignment_id,
    'status', 'pending',
    'learner_count', v_learner_count
  );
end;
$$;

create or replace function public.list_my_student_list_assignments()
returns table (
  id uuid,
  class_id uuid,
  subject text,
  status text,
  sent_at timestamptz,
  responded_at timestamptz,
  decline_reason text,
  grade_level text,
  section text,
  school_year text,
  adviser_name text,
  learner_count bigint
)
language sql
security definer
set search_path = public
stable
as $$
  select
    assignment.id,
    assignment.class_id,
    assignment.subject,
    assignment.status,
    assignment.sent_at,
    assignment.responded_at,
    assignment.decline_reason,
    class_row.grade_level,
    class_row.section,
    class_row.school_year,
    coalesce(adviser.full_name, adviser.email, 'Class Adviser'),
    count(learner.id)
  from public.student_list_assignments assignment
  join public.classes class_row on class_row.id = assignment.class_id
  left join public.profiles adviser on adviser.id = assignment.adviser_id
  left join public.student_list_assignment_learners learner
    on learner.assignment_id = assignment.id
  where assignment.subject_teacher_id = auth.uid()
  group by assignment.id, class_row.id, adviser.id
  order by
    case assignment.status when 'pending' then 0 when 'accepted' then 1 else 2 end,
    assignment.sent_at desc;
$$;

create or replace function public.list_student_list_assignment_learners(
  p_assignment_id uuid
)
returns table (
  id uuid,
  lrn text,
  first_name text,
  middle_name text,
  last_name text,
  suffix text,
  sex text
)
language sql
security definer
set search_path = public
stable
as $$
  select
    learner.id,
    learner.lrn,
    learner.first_name,
    learner.middle_name,
    learner.last_name,
    learner.suffix,
    learner.sex
  from public.student_list_assignment_learners learner
  join public.student_list_assignments assignment
    on assignment.id = learner.assignment_id
  where learner.assignment_id = p_assignment_id
    and (
      assignment.adviser_id = auth.uid()
      or assignment.subject_teacher_id = auth.uid()
    )
  order by
    case learner.sex when 'male' then 0 when 'female' then 1 else 2 end,
    learner.list_position,
    lower(learner.last_name),
    lower(learner.first_name);
$$;

create or replace function public.accept_student_list_assignment(
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
  v_source_class public.classes%rowtype;
  v_subject_class_id uuid;
  v_learner_count integer := 0;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to accept students.';
  end if;

  select * into v_assignment
  from public.student_list_assignments
  where id = p_assignment_id
    and subject_teacher_id = v_user_id
  for update;

  if not found then
    raise exception 'This student assignment was not sent to you.';
  end if;

  if v_assignment.status = 'declined' then
    raise exception 'A declined assignment must be sent again by the Class Adviser.';
  end if;

  select id into v_subject_class_id
  from public.classes
  where student_list_assignment_id = v_assignment.id;

  if v_assignment.status = 'accepted' and v_subject_class_id is not null then
    return jsonb_build_object(
      'assignment_id', v_assignment.id,
      'status', 'accepted',
      'class_id', v_subject_class_id,
      'already_accepted', true
    );
  end if;

  select * into v_source_class
  from public.classes
  where id = v_assignment.class_id;

  if not found then
    raise exception 'The Class Adviser class no longer exists.';
  end if;

  insert into public.classes (
    teacher_id,
    subject,
    grade_level,
    section,
    school_year,
    start_date,
    end_date,
    student_list_assignment_id
  ) values (
    v_user_id,
    v_assignment.subject,
    v_source_class.grade_level,
    v_source_class.section,
    v_source_class.school_year,
    v_source_class.start_date,
    v_source_class.end_date,
    v_assignment.id
  )
  returning id into v_subject_class_id;

  insert into public.students (
    class_id,
    teacher_id,
    last_name,
    first_name,
    middle_name,
    sex,
    lrn,
    birthdate,
    address,
    mother_name,
    father_name,
    guardian,
    contact_number
  )
  select
    v_subject_class_id,
    v_user_id,
    learner.last_name,
    learner.first_name,
    learner.middle_name,
    learner.sex,
    learner.lrn,
    learner.birthdate,
    learner.address,
    learner.mother_name,
    learner.father_name,
    learner.guardian,
    learner.contact_number
  from public.student_list_assignment_learners learner
  where learner.assignment_id = v_assignment.id
  order by learner.list_position;

  get diagnostics v_learner_count = row_count;

  update public.student_list_assignments
  set status = 'accepted',
      decline_reason = null,
      responded_at = now(),
      accepted_at = now(),
      declined_at = null,
      updated_at = now()
  where id = v_assignment.id;

  return jsonb_build_object(
    'assignment_id', v_assignment.id,
    'status', 'accepted',
    'class_id', v_subject_class_id,
    'learner_count', v_learner_count
  );
end;
$$;

create or replace function public.decline_student_list_assignment(
  p_assignment_id uuid,
  p_reason text default null
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_assignment public.student_list_assignments%rowtype;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to decline students.';
  end if;

  select * into v_assignment
  from public.student_list_assignments
  where id = p_assignment_id
    and subject_teacher_id = v_user_id
  for update;

  if not found then
    raise exception 'This student assignment was not sent to you.';
  end if;

  if v_assignment.status = 'accepted' then
    raise exception 'An accepted student assignment cannot be declined.';
  end if;

  update public.student_list_assignments
  set status = 'declined',
      decline_reason = nullif(btrim(p_reason), ''),
      responded_at = now(),
      declined_at = now(),
      accepted_at = null,
      updated_at = now()
  where id = v_assignment.id;

  return jsonb_build_object(
    'assignment_id', v_assignment.id,
    'status', 'declined'
  );
end;
$$;

-- Replace the broad owner policies with operation-specific policies. Normal
-- classes retain their existing behavior. Assignment-created class metadata
-- and learner details are read-only to the Subject Teacher.
drop policy if exists "own classes" on public.classes;
drop policy if exists "teachers read own classes" on public.classes;
drop policy if exists "teachers insert own classes" on public.classes;
drop policy if exists "teachers update own regular classes" on public.classes;
drop policy if exists "teachers delete own regular classes" on public.classes;

create policy "teachers read own classes"
on public.classes for select to authenticated
using (teacher_id = auth.uid());

create policy "teachers insert own classes"
on public.classes for insert to authenticated
with check (teacher_id = auth.uid() and student_list_assignment_id is null);

create policy "teachers update own regular classes"
on public.classes for update to authenticated
using (teacher_id = auth.uid() and student_list_assignment_id is null)
with check (teacher_id = auth.uid() and student_list_assignment_id is null);

create policy "teachers delete own regular classes"
on public.classes for delete to authenticated
using (teacher_id = auth.uid() and student_list_assignment_id is null);

drop policy if exists "own students" on public.students;
drop policy if exists "teachers read own students" on public.students;
drop policy if exists "teachers insert own regular class students" on public.students;
drop policy if exists "teachers update own regular class students" on public.students;
drop policy if exists "teachers delete own regular class students" on public.students;

create policy "teachers read own students"
on public.students for select to authenticated
using (teacher_id = auth.uid());

create policy "teachers insert own regular class students"
on public.students for insert to authenticated
with check (
  teacher_id = auth.uid()
  and exists (
    select 1 from public.classes class_row
    where class_row.id = students.class_id
      and class_row.teacher_id = auth.uid()
      and class_row.student_list_assignment_id is null
  )
);

create policy "teachers update own regular class students"
on public.students for update to authenticated
using (
  teacher_id = auth.uid()
  and exists (
    select 1 from public.classes class_row
    where class_row.id = students.class_id
      and class_row.teacher_id = auth.uid()
      and class_row.student_list_assignment_id is null
  )
)
with check (
  teacher_id = auth.uid()
  and exists (
    select 1 from public.classes class_row
    where class_row.id = students.class_id
      and class_row.teacher_id = auth.uid()
      and class_row.student_list_assignment_id is null
  )
);

create policy "teachers delete own regular class students"
on public.students for delete to authenticated
using (
  teacher_id = auth.uid()
  and exists (
    select 1 from public.classes class_row
    where class_row.id = students.class_id
      and class_row.teacher_id = auth.uid()
      and class_row.student_list_assignment_id is null
  )
);

drop policy if exists "own grades" on public.grades;
drop policy if exists "teachers read own grades" on public.grades;
drop policy if exists "teachers insert grades for allowed subject" on public.grades;
drop policy if exists "teachers update grades for allowed subject" on public.grades;
drop policy if exists "teachers delete grades for allowed subject" on public.grades;

create policy "teachers read own grades"
on public.grades for select to authenticated
using (teacher_id = auth.uid());

create policy "teachers insert grades for allowed subject"
on public.grades for insert to authenticated
with check (
  teacher_id = auth.uid()
  and exists (
    select 1
    from public.classes class_row
    left join public.student_list_assignments assignment
      on assignment.id = class_row.student_list_assignment_id
    where class_row.id = grades.class_id
      and class_row.teacher_id = auth.uid()
      and (
        class_row.student_list_assignment_id is null
        or (
          assignment.status = 'accepted'
          and assignment.subject_teacher_id = auth.uid()
          and lower(btrim(grades.subject)) = lower(btrim(assignment.subject))
        )
      )
  )
);

create policy "teachers update grades for allowed subject"
on public.grades for update to authenticated
using (
  teacher_id = auth.uid()
  and exists (
    select 1
    from public.classes class_row
    left join public.student_list_assignments assignment
      on assignment.id = class_row.student_list_assignment_id
    where class_row.id = grades.class_id
      and class_row.teacher_id = auth.uid()
      and (
        class_row.student_list_assignment_id is null
        or (
          assignment.status = 'accepted'
          and assignment.subject_teacher_id = auth.uid()
          and lower(btrim(grades.subject)) = lower(btrim(assignment.subject))
        )
      )
  )
)
with check (
  teacher_id = auth.uid()
  and exists (
    select 1
    from public.classes class_row
    left join public.student_list_assignments assignment
      on assignment.id = class_row.student_list_assignment_id
    where class_row.id = grades.class_id
      and class_row.teacher_id = auth.uid()
      and (
        class_row.student_list_assignment_id is null
        or (
          assignment.status = 'accepted'
          and assignment.subject_teacher_id = auth.uid()
          and lower(btrim(grades.subject)) = lower(btrim(assignment.subject))
        )
      )
  )
);

create policy "teachers delete grades for allowed subject"
on public.grades for delete to authenticated
using (
  teacher_id = auth.uid()
  and exists (
    select 1
    from public.classes class_row
    left join public.student_list_assignments assignment
      on assignment.id = class_row.student_list_assignment_id
    where class_row.id = grades.class_id
      and class_row.teacher_id = auth.uid()
      and (
        class_row.student_list_assignment_id is null
        or (
          assignment.status = 'accepted'
          and assignment.subject_teacher_id = auth.uid()
          and lower(btrim(grades.subject)) = lower(btrim(assignment.subject))
        )
      )
  )
);

revoke all on function public.send_student_list_assignment(uuid, text, uuid) from public;
revoke all on function public.list_my_student_list_assignments() from public;
revoke all on function public.list_student_list_assignment_learners(uuid) from public;
revoke all on function public.accept_student_list_assignment(uuid) from public;
revoke all on function public.decline_student_list_assignment(uuid, text) from public;

grant execute on function public.send_student_list_assignment(uuid, text, uuid)
  to authenticated;
grant execute on function public.list_my_student_list_assignments()
  to authenticated;
grant execute on function public.list_student_list_assignment_learners(uuid)
  to authenticated;
grant execute on function public.accept_student_list_assignment(uuid)
  to authenticated;
grant execute on function public.decline_student_list_assignment(uuid, text)
  to authenticated;

notify pgrst, 'reload schema';
