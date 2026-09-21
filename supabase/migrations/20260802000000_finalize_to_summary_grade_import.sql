-- Connect the reviewed Grade Request workflow to Summary of Grades.
-- Only finalized requests owned by the logged-in Class Adviser can be imported.

create table if not exists public.grade_request_subject_imports (
  id uuid primary key default gen_random_uuid(),
  subject_request_id uuid not null references public.grade_request_subjects(id) on delete cascade,
  target_class_id uuid not null references public.classes(id) on delete cascade,
  imported_by uuid not null references auth.users(id) on delete cascade,
  imported_at timestamptz not null default now(),
  imported_count integer not null default 0,
  unmatched_count integer not null default 0,
  unique (subject_request_id, target_class_id)
);

alter table public.grade_request_subject_imports enable row level security;

drop policy if exists "advisers read own finalized grade imports"
  on public.grade_request_subject_imports;
create policy "advisers read own finalized grade imports"
on public.grade_request_subject_imports for select to authenticated
using (imported_by = auth.uid());

grant select on public.grade_request_subject_imports to authenticated;

create or replace function public.list_my_finalized_grade_request_subjects(
  p_advisory_class_id uuid
)
returns table (
  subject_request_id uuid,
  batch_id uuid,
  advisory_class_id uuid,
  subject text,
  grading_period text,
  grade_level text,
  section text,
  school_year text,
  recipient_name text,
  submitted_at timestamptz,
  finalized_at timestamptz,
  learner_count bigint,
  is_imported boolean
)
language sql
security definer
set search_path = public
stable
as $$
  select
    s.id,
    b.id,
    b.advisory_class_id,
    s.subject,
    b.grading_period,
    c.grade_level,
    c.section,
    c.school_year,
    coalesce(p.full_name, p.email),
    s.submitted_at,
    b.finalized_at,
    count(sc.student_id),
    exists (
      select 1
      from public.grade_request_subject_imports i
      where i.subject_request_id = s.id
        and i.target_class_id = b.advisory_class_id
    )
  from public.grade_request_batches b
  join public.grade_request_subjects s on s.batch_id = b.id
  join public.classes c on c.id = b.advisory_class_id
  left join public.profiles p on p.id = s.subject_teacher_id
  left join public.grade_request_scores sc on sc.subject_request_id = s.id
  where b.adviser_id = auth.uid()
    and b.advisory_class_id = p_advisory_class_id
    and b.is_finalized = true
    and s.status = 'submitted'
  group by s.id, b.id, c.id, p.id
  order by b.finalized_at desc, s.subject;
$$;

create or replace function public.import_finalized_grade_request_subject(
  p_subject_request_id uuid,
  p_target_class_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid := auth.uid();
  v_subject text;
  v_term text;
  v_batch_id uuid;
  v_imported_count integer := 0;
  v_score_count integer := 0;
  v_unmatched_count integer := 0;
begin
  if v_user_id is null then
    raise exception 'You must be signed in to import grades.';
  end if;

  select s.subject, b.grading_period, b.id
  into v_subject, v_term, v_batch_id
  from public.grade_request_subjects s
  join public.grade_request_batches b on b.id = s.batch_id
  join public.classes c on c.id = b.advisory_class_id
  where s.id = p_subject_request_id
    and b.adviser_id = v_user_id
    and b.advisory_class_id = p_target_class_id
    and c.teacher_id = v_user_id
    and b.is_finalized = true
    and s.status = 'submitted';

  if not found then
    raise exception 'Only a finalized Grade Request for this advisory class can be imported.';
  end if;

  if exists (
    select 1
    from public.grade_request_subject_imports i
    where i.subject_request_id = p_subject_request_id
      and i.target_class_id = p_target_class_id
  ) then
    raise exception 'This finalized form has already been added to Summary of Grades.';
  end if;

  select count(*) into v_score_count
  from public.grade_request_scores
  where subject_request_id = p_subject_request_id;

  with matched_scores as (
    select distinct on (sc.student_id)
      target.id as target_student_id,
      sc.score
    from public.grade_request_scores sc
    join public.students source on source.id = sc.student_id
    join lateral (
      select candidate.id
      from public.students candidate
      where candidate.class_id = p_target_class_id
        -- Add a finalized grade only to the learner with the same complete name.
        -- Lower-casing and collapsing spaces avoids false mismatches caused only
        -- by capitalization or accidental extra spaces.
        and regexp_replace(
          lower(concat_ws(' ', candidate.first_name, candidate.middle_name, candidate.last_name)),
          '\s+', ' ', 'g'
        ) = regexp_replace(
          lower(concat_ws(' ', source.first_name, source.middle_name, source.last_name)),
          '\s+', ' ', 'g'
        )
      order by candidate.id
      limit 1
    ) target on true
    where sc.subject_request_id = p_subject_request_id
    order by sc.student_id
  ), inserted as (
    insert into public.grades (
      student_id, class_id, teacher_id, subject, term, score
    )
    select
      target_student_id, p_target_class_id, v_user_id, v_subject, v_term, score
    from matched_scores
    on conflict (student_id, subject, term)
    do update set
      class_id = excluded.class_id,
      teacher_id = excluded.teacher_id,
      score = excluded.score
    returning 1
  )
  select count(*) into v_imported_count from inserted;

  v_unmatched_count := greatest(v_score_count - v_imported_count, 0);

  insert into public.grade_request_subject_imports (
    subject_request_id, target_class_id, imported_by, imported_count, unmatched_count
  ) values (
    p_subject_request_id, p_target_class_id, v_user_id,
    v_imported_count, v_unmatched_count
  );

  insert into public.grade_request_history (
    batch_id, subject_request_id, actor_id, event_type, new_status, details
  ) values (
    v_batch_id, p_subject_request_id, v_user_id,
    'grades_imported_to_summary', 'finalized',
    jsonb_build_object(
      'target_class_id', p_target_class_id,
      'subject', v_subject,
      'grading_period', v_term,
      'imported_count', v_imported_count,
      'unmatched_count', v_unmatched_count
    )
  );

  return jsonb_build_object(
    'subject', v_subject,
    'grading_period', v_term,
    'imported_count', v_imported_count,
    'unmatched_count', v_unmatched_count
  );
end;
$$;

revoke all on function public.list_my_finalized_grade_request_subjects(uuid) from public;
revoke all on function public.import_finalized_grade_request_subject(uuid, uuid) from public;
grant execute on function public.list_my_finalized_grade_request_subjects(uuid) to authenticated;
grant execute on function public.import_finalized_grade_request_subject(uuid, uuid) to authenticated;
