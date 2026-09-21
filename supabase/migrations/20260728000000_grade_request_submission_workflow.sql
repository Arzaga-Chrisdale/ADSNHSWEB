create extension if not exists "pgcrypto";

create table if not exists public.grade_request_batches (
  id uuid primary key default gen_random_uuid(),
  adviser_id uuid not null references auth.users(id) on delete cascade,
  advisory_class_id uuid not null references public.classes(id) on delete cascade,
  grading_period text not null check (grading_period in ('1', '2', '3', 'final')),
  status text not null default 'pending' check (status in ('pending', 'completed')),
  is_finalized boolean not null default false,
  requested_at timestamptz not null default now(),
  completed_at timestamptz,
  finalized_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists grade_request_batches_one_open_cycle
  on public.grade_request_batches (advisory_class_id, grading_period)
  where is_finalized = false;

create table if not exists public.grade_request_subjects (
  id uuid primary key default gen_random_uuid(),
  batch_id uuid not null references public.grade_request_batches(id) on delete cascade,
  subject_class_id uuid references public.classes(id) on delete set null,
  subject_teacher_id uuid not null references auth.users(id) on delete cascade,
  subject text not null,
  status text not null default 'pending' check (status in ('pending', 'submitted')),
  requested_at timestamptz not null default now(),
  submitted_at timestamptz,
  teacher_note text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (batch_id, subject_teacher_id, subject)
);

create table if not exists public.grade_request_scores (
  id uuid primary key default gen_random_uuid(),
  subject_request_id uuid not null references public.grade_request_subjects(id) on delete cascade,
  student_id uuid not null references public.students(id) on delete cascade,
  score numeric(5,2) not null check (score >= 0 and score <= 100),
  submitted_at timestamptz not null default now(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (subject_request_id, student_id)
);

create table if not exists public.grade_request_history (
  id uuid primary key default gen_random_uuid(),
  batch_id uuid not null references public.grade_request_batches(id) on delete cascade,
  subject_request_id uuid references public.grade_request_subjects(id) on delete cascade,
  actor_id uuid references auth.users(id) on delete set null,
  event_type text not null,
  old_status text,
  new_status text,
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create or replace function public.set_grade_request_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists set_grade_request_batches_updated_at on public.grade_request_batches;
create trigger set_grade_request_batches_updated_at
before update on public.grade_request_batches
for each row execute function public.set_grade_request_updated_at();

drop trigger if exists set_grade_request_subjects_updated_at on public.grade_request_subjects;
create trigger set_grade_request_subjects_updated_at
before update on public.grade_request_subjects
for each row execute function public.set_grade_request_updated_at();

drop trigger if exists set_grade_request_scores_updated_at on public.grade_request_scores;
create trigger set_grade_request_scores_updated_at
before update on public.grade_request_scores
for each row execute function public.set_grade_request_updated_at();

create or replace function public.refresh_grade_request_batch_status()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_batch_id uuid;
  v_total integer;
  v_submitted integer;
  v_old_status text;
  v_new_status text;
begin
  if tg_op = 'DELETE' then
    v_batch_id := old.batch_id;
  else
    v_batch_id := new.batch_id;
  end if;

  select status into v_old_status
  from public.grade_request_batches
  where id = v_batch_id;

  select
    count(*),
    count(*) filter (where status = 'submitted')
  into v_total, v_submitted
  from public.grade_request_subjects
  where batch_id = v_batch_id;

  v_new_status := case
    when v_total > 0 and v_total = v_submitted then 'completed'
    else 'pending'
  end;

  update public.grade_request_batches
  set
    status = v_new_status,
    completed_at = case
      when v_new_status = 'completed' then coalesce(completed_at, now())
      else null
    end
  where id = v_batch_id
    and is_finalized = false;

  if v_old_status is distinct from v_new_status then
    insert into public.grade_request_history (
      batch_id,
      actor_id,
      event_type,
      old_status,
      new_status,
      details
    ) values (
      v_batch_id,
      auth.uid(),
      'overall_status_changed',
      v_old_status,
      v_new_status,
      jsonb_build_object('submitted_subjects', v_submitted, 'total_subjects', v_total)
    );
  end if;

  if tg_op = 'DELETE' then
    return old;
  end if;

  return new;
end;
$$;

drop trigger if exists refresh_grade_request_batch_after_subject_change on public.grade_request_subjects;
create trigger refresh_grade_request_batch_after_subject_change
after insert or update of status or delete on public.grade_request_subjects
for each row execute function public.refresh_grade_request_batch_status();

create or replace function public.create_grade_request_batch(
  p_advisory_class_id uuid,
  p_grading_period text
)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch_id uuid;
  v_class public.classes%rowtype;
  v_subject_count integer;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if p_grading_period not in ('1', '2', '3', 'final') then
    raise exception 'Invalid grading period.';
  end if;

  select * into v_class
  from public.classes
  where id = p_advisory_class_id;

  if not found then
    raise exception 'The selected advisory class does not exist.';
  end if;

  if v_class.teacher_id <> v_user_id then
    raise exception 'Only the assigned Class Adviser can request grades for this class.';
  end if;

  if not exists (
    select 1 from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can create grade requests.';
  end if;

  if exists (
    select 1
    from public.grade_request_batches
    where advisory_class_id = p_advisory_class_id
      and grading_period = p_grading_period
      and is_finalized = false
  ) then
    raise exception 'An open grade request already exists for this class and grading period.';
  end if;

  insert into public.grade_request_batches (
    adviser_id,
    advisory_class_id,
    grading_period,
    status
  ) values (
    v_user_id,
    p_advisory_class_id,
    p_grading_period,
    'pending'
  )
  returning id into v_batch_id;

  insert into public.grade_request_subjects (
    batch_id,
    subject_class_id,
    subject_teacher_id,
    subject,
    status
  )
  select distinct on (lower(trim(c.subject)), c.teacher_id)
    v_batch_id,
    c.id,
    c.teacher_id,
    trim(c.subject),
    'pending'
  from public.classes c
  join public.profiles p on p.id = c.teacher_id
  where p.teacher_type = 'subject_teacher'
    and c.subject is not null
    and trim(c.subject) <> ''
    and lower(trim(coalesce(c.grade_level, ''))) = lower(trim(coalesce(v_class.grade_level, '')))
    and lower(trim(coalesce(c.section, ''))) = lower(trim(coalesce(v_class.section, '')))
    and lower(trim(coalesce(c.school_year, ''))) = lower(trim(coalesce(v_class.school_year, '')))
  order by lower(trim(c.subject)), c.teacher_id, c.id;

  get diagnostics v_subject_count = row_count;

  if v_subject_count = 0 then
    delete from public.grade_request_batches where id = v_batch_id;
    raise exception 'No assigned Subject Teachers were found for the same grade level, section, and school year.';
  end if;

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    new_status,
    details
  ) values (
    v_batch_id,
    v_user_id,
    'request_created',
    'pending',
    jsonb_build_object(
      'grading_period', p_grading_period,
      'subject_count', v_subject_count,
      'advisory_class_id', p_advisory_class_id
    )
  );

  return v_batch_id;
end;
$$;

create or replace function public.submit_grade_request_subject(
  p_subject_request_id uuid,
  p_scores jsonb,
  p_teacher_note text default null
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_item public.grade_request_subjects%rowtype;
  v_batch public.grade_request_batches%rowtype;
  v_score jsonb;
  v_student_id uuid;
  v_numeric_score numeric;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  select * into v_item
  from public.grade_request_subjects
  where id = p_subject_request_id;

  if not found then
    raise exception 'The grade request does not exist.';
  end if;

  if v_item.subject_teacher_id <> v_user_id then
    raise exception 'You are not assigned to this subject request.';
  end if;

  if not exists (
    select 1 from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'subject_teacher'
  ) then
    raise exception 'Only a Subject Teacher can submit requested grades.';
  end if;

  select * into v_batch
  from public.grade_request_batches
  where id = v_item.batch_id;

  if v_batch.is_finalized then
    raise exception 'This grading period has already been finalized.';
  end if;

  if jsonb_typeof(p_scores) <> 'array' or jsonb_array_length(p_scores) = 0 then
    raise exception 'Submit at least one learner grade.';
  end if;

  delete from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  for v_score in select * from jsonb_array_elements(p_scores)
  loop
    v_student_id := (v_score ->> 'student_id')::uuid;
    v_numeric_score := (v_score ->> 'score')::numeric;

    if v_numeric_score < 0 or v_numeric_score > 100 then
      raise exception 'Every grade must be between 0 and 100.';
    end if;

    if not exists (
      select 1
      from public.students s
      where s.id = v_student_id
        and s.class_id = v_batch.advisory_class_id
    ) then
      raise exception 'A submitted learner does not belong to the advisory class.';
    end if;

    insert into public.grade_request_scores (
      subject_request_id,
      student_id,
      score,
      submitted_at
    ) values (
      p_subject_request_id,
      v_student_id,
      v_numeric_score,
      now()
    );

    update public.grades
    set
      score = v_numeric_score,
      teacher_id = v_user_id,
      class_id = v_batch.advisory_class_id
    where student_id = v_student_id
      and subject = v_item.subject
      and term = v_batch.grading_period;

    if not found then
      insert into public.grades (
        student_id,
        class_id,
        teacher_id,
        subject,
        term,
        score
      ) values (
        v_student_id,
        v_batch.advisory_class_id,
        v_user_id,
        v_item.subject,
        v_batch.grading_period,
        v_numeric_score
      );
    end if;
  end loop;

  update public.grade_request_subjects
  set
    status = 'submitted',
    submitted_at = now(),
    teacher_note = nullif(trim(p_teacher_note), '')
  where id = p_subject_request_id;

  insert into public.grade_request_history (
    batch_id,
    subject_request_id,
    actor_id,
    event_type,
    old_status,
    new_status,
    details
  ) values (
    v_item.batch_id,
    p_subject_request_id,
    v_user_id,
    'subject_grades_submitted',
    v_item.status,
    'submitted',
    jsonb_build_object(
      'subject', v_item.subject,
      'score_count', jsonb_array_length(p_scores),
      'submitted_at', now()
    )
  );
end;
$$;

create or replace function public.finalize_grade_request_batch(p_batch_id uuid)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_batch public.grade_request_batches%rowtype;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  select * into v_batch
  from public.grade_request_batches
  where id = p_batch_id;

  if not found then
    raise exception 'The grade request does not exist.';
  end if;

  if v_batch.adviser_id <> v_user_id then
    raise exception 'Only the requesting Class Adviser can finalize this grading period.';
  end if;

  if not exists (
    select 1 from public.profiles p
    where p.id = v_user_id
      and p.teacher_type = 'class_adviser'
  ) then
    raise exception 'Only a Class Adviser can finalize grades.';
  end if;

  if v_batch.status <> 'completed' then
    raise exception 'All Subject Teachers must submit their grades before finalization.';
  end if;

  if v_batch.is_finalized then
    return;
  end if;

  update public.grade_request_batches
  set
    is_finalized = true,
    finalized_at = now()
  where id = p_batch_id;

  insert into public.grade_request_history (
    batch_id,
    actor_id,
    event_type,
    old_status,
    new_status,
    details
  ) values (
    p_batch_id,
    v_user_id,
    'grades_finalized',
    'completed',
    'completed',
    jsonb_build_object('finalized_at', now())
  );
end;
$$;

create or replace view public.grade_request_workflow_view
with (security_invoker = true) as
select
  b.id as batch_id,
  b.adviser_id,
  adviser.full_name as adviser_name,
  adviser.email as adviser_email,
  b.advisory_class_id,
  ac.grade_level,
  ac.section,
  ac.school_year,
  b.grading_period,
  b.status as overall_status,
  b.is_finalized,
  b.requested_at,
  b.completed_at,
  b.finalized_at,
  s.id as subject_request_id,
  s.subject_class_id,
  s.subject_teacher_id,
  teacher.full_name as subject_teacher_name,
  teacher.email as subject_teacher_email,
  s.subject,
  s.status as subject_status,
  s.requested_at as subject_requested_at,
  s.submitted_at,
  s.teacher_note,
  count(sc.id) as submitted_grade_count
from public.grade_request_batches b
join public.classes ac on ac.id = b.advisory_class_id
join public.profiles adviser on adviser.id = b.adviser_id
join public.grade_request_subjects s on s.batch_id = b.id
join public.profiles teacher on teacher.id = s.subject_teacher_id
left join public.grade_request_scores sc on sc.subject_request_id = s.id
group by
  b.id,
  adviser.full_name,
  adviser.email,
  ac.grade_level,
  ac.section,
  ac.school_year,
  s.id,
  teacher.full_name,
  teacher.email;

alter table public.grade_request_batches enable row level security;
alter table public.grade_request_subjects enable row level security;
alter table public.grade_request_scores enable row level security;
alter table public.grade_request_history enable row level security;

-- ============================================================
-- Non-recursive access helpers
--
-- These SECURITY DEFINER functions read the workflow tables without
-- re-entering their RLS policies. They always use auth.uid() internally,
-- so callers cannot impersonate another user by passing a user ID.
-- ============================================================

create or replace function public.can_access_grade_request_batch(
  p_batch_id uuid
)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.grade_request_batches b
    where b.id = p_batch_id
      and (
        b.adviser_id = auth.uid()
        or exists (
          select 1
          from public.grade_request_subjects s
          where s.batch_id = b.id
            and s.subject_teacher_id = auth.uid()
        )
      )
  );
$$;

create or replace function public.can_access_grade_request_subject(
  p_subject_request_id uuid
)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.grade_request_subjects s
    join public.grade_request_batches b
      on b.id = s.batch_id
    where s.id = p_subject_request_id
      and (
        s.subject_teacher_id = auth.uid()
        or b.adviser_id = auth.uid()
      )
  );
$$;

revoke all on function public.can_access_grade_request_batch(uuid) from public;
revoke all on function public.can_access_grade_request_subject(uuid) from public;
grant execute on function public.can_access_grade_request_batch(uuid) to authenticated;
grant execute on function public.can_access_grade_request_subject(uuid) to authenticated;

-- ============================================================
-- Recreate workflow policies without circular table-policy checks
-- ============================================================

drop policy if exists grade_request_batches_select_participants
on public.grade_request_batches;

create policy grade_request_batches_select_participants
on public.grade_request_batches
for select
to authenticated
using (
  public.can_access_grade_request_batch(id)
);

drop policy if exists grade_request_subjects_select_participants
on public.grade_request_subjects;

create policy grade_request_subjects_select_participants
on public.grade_request_subjects
for select
to authenticated
using (
  public.can_access_grade_request_subject(id)
);

drop policy if exists grade_request_subjects_update_assigned_teacher
on public.grade_request_subjects;

create policy grade_request_subjects_update_assigned_teacher
on public.grade_request_subjects
for update
to authenticated
using (
  subject_teacher_id = auth.uid()
)
with check (
  subject_teacher_id = auth.uid()
);

drop policy if exists grade_request_scores_select_participants
on public.grade_request_scores;

create policy grade_request_scores_select_participants
on public.grade_request_scores
for select
to authenticated
using (
  public.can_access_grade_request_subject(subject_request_id)
);

drop policy if exists grade_request_history_select_participants
on public.grade_request_history;

create policy grade_request_history_select_participants
on public.grade_request_history
for select
to authenticated
using (
  public.can_access_grade_request_batch(batch_id)
);

grant select on public.grade_request_batches to authenticated;
grant select on public.grade_request_subjects to authenticated;
grant select on public.grade_request_scores to authenticated;
grant select on public.grade_request_history to authenticated;
grant select on public.grade_request_workflow_view to authenticated;
grant execute on function public.create_grade_request_batch(uuid, text) to authenticated;
grant execute on function public.submit_grade_request_subject(uuid, jsonb, text) to authenticated;
grant execute on function public.finalize_grade_request_batch(uuid) to authenticated;

notify pgrst, 'reload schema';
