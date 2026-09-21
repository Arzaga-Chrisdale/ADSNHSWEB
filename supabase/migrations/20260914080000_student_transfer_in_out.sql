-- ============================================================
-- 20260914080000_student_transfer_in_out.sql
-- Transfer In / Transfer Out for existing public.students/classes
-- Keeps learner and grade history. Does NOT delete grades.
-- ============================================================

begin;

-- 1) Add transfer state to students
alter table public.students
  add column if not exists enrollment_status text not null default 'active',
  add column if not exists is_active boolean not null default true,
  add column if not exists transfer_date date,
  add column if not exists transfer_effective_term text,
  add column if not exists previous_school text,
  add column if not exists destination_school text,
  add column if not exists transfer_reason text,
  add column if not exists transferred_at timestamptz;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'students_enrollment_status_check'
      and conrelid = 'public.students'::regclass
  ) then
    alter table public.students
      add constraint students_enrollment_status_check
      check (enrollment_status in ('active', 'transferred_in', 'transferred_out'));
  end if;
end
$$;

create index if not exists students_class_active_idx
  on public.students (class_id, is_active);

create index if not exists students_enrollment_status_idx
  on public.students (enrollment_status);

-- 2) Transfer audit/history table
create table if not exists public.student_transfers (
  id uuid primary key default gen_random_uuid(),
  student_id uuid not null references public.students(id) on delete restrict,
  transfer_type text not null check (transfer_type in ('in', 'out')),
  from_class_id uuid references public.classes(id) on delete set null,
  to_class_id uuid references public.classes(id) on delete set null,
  school_year text,
  effective_term text,
  transfer_date date not null default current_date,
  previous_school text,
  destination_school text,
  reason text,
  remarks text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create index if not exists student_transfers_student_idx
  on public.student_transfers (student_id, created_at desc);

create index if not exists student_transfers_from_class_idx
  on public.student_transfers (from_class_id);

create index if not exists student_transfers_to_class_idx
  on public.student_transfers (to_class_id);

-- 3) Permission helper
-- Allows Admin or the user attached to the class through common owner/adviser fields.
create or replace function public.can_manage_student_transfer(p_class_id uuid)
returns boolean
language sql
stable
security definer
set search_path = public, pg_temp
as $$
  select
    auth.uid() is not null
    and (
      exists (
        select 1
        from public.user_roles ur
        where ur.user_id = auth.uid()
          and ur.role = 'admin'
      )
      or exists (
        select 1
        from public.classes c
        where c.id = p_class_id
          and auth.uid() in (
            nullif(to_jsonb(c)->>'teacher_id', '')::uuid,
            nullif(to_jsonb(c)->>'adviser_id', '')::uuid,
            nullif(to_jsonb(c)->>'class_adviser_id', '')::uuid,
            nullif(to_jsonb(c)->>'owner_id', '')::uuid,
            nullif(to_jsonb(c)->>'user_id', '')::uuid
          )
      )
    );
$$;

grant execute on function public.can_manage_student_transfer(uuid)
to authenticated;

-- 4) Transfer OUT
-- Keeps student.class_id pointing at the historical class.
-- Makes the learner inactive for the current roster.
create or replace function public.transfer_out_student(
  p_student_id uuid,
  p_transfer_date date default current_date,
  p_destination_school text default null,
  p_reason text default null,
  p_effective_term text default null,
  p_remarks text default null
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_student public.students%rowtype;
  v_school_year text;
  v_transfer_id uuid;
begin
  if auth.uid() is null then
    raise exception 'You must be signed in.';
  end if;

  select *
  into v_student
  from public.students
  where id = p_student_id
  for update;

  if not found then
    raise exception 'Learner not found.';
  end if;

  if v_student.class_id is null then
    raise exception 'Learner is not assigned to a class.';
  end if;

  if not public.can_manage_student_transfer(v_student.class_id) then
    raise exception 'You are not allowed to transfer this learner.';
  end if;

  if coalesce(v_student.is_active, true) = false
     and v_student.enrollment_status = 'transferred_out' then
    raise exception 'Learner is already transferred out.';
  end if;

  select nullif(to_jsonb(c)->>'school_year', '')
  into v_school_year
  from public.classes c
  where c.id = v_student.class_id;

  insert into public.student_transfers (
    student_id,
    transfer_type,
    from_class_id,
    school_year,
    effective_term,
    transfer_date,
    destination_school,
    reason,
    remarks,
    created_by
  )
  values (
    v_student.id,
    'out',
    v_student.class_id,
    v_school_year,
    nullif(trim(p_effective_term), ''),
    coalesce(p_transfer_date, current_date),
    nullif(trim(p_destination_school), ''),
    nullif(trim(p_reason), ''),
    nullif(trim(p_remarks), ''),
    auth.uid()
  )
  returning id into v_transfer_id;

  update public.students
  set
    enrollment_status = 'transferred_out',
    is_active = false,
    transfer_date = coalesce(p_transfer_date, current_date),
    transfer_effective_term = nullif(trim(p_effective_term), ''),
    destination_school = nullif(trim(p_destination_school), ''),
    transfer_reason = nullif(trim(p_reason), ''),
    transferred_at = now()
  where id = v_student.id;

  return v_transfer_id;
end;
$$;

grant execute on function public.transfer_out_student(
  uuid, date, text, text, text, text
) to authenticated;

-- 5) Transfer IN
-- For an incoming learner that already exists in public.students.
-- For a completely new learner, first create the learner using your existing
-- Add Student flow, then call this function.
create or replace function public.transfer_in_student(
  p_student_id uuid,
  p_to_class_id uuid,
  p_transfer_date date default current_date,
  p_previous_school text default null,
  p_effective_term text default null,
  p_reason text default null,
  p_remarks text default null
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_student public.students%rowtype;
  v_from_class_id uuid;
  v_school_year text;
  v_transfer_id uuid;
begin
  if auth.uid() is null then
    raise exception 'You must be signed in.';
  end if;

  if not exists (
    select 1
    from public.classes
    where id = p_to_class_id
  ) then
    raise exception 'Target class not found.';
  end if;

  if not public.can_manage_student_transfer(p_to_class_id) then
    raise exception 'You are not allowed to transfer a learner into this class.';
  end if;

  select *
  into v_student
  from public.students
  where id = p_student_id
  for update;

  if not found then
    raise exception 'Learner not found.';
  end if;

  v_from_class_id := v_student.class_id;

  if v_student.class_id = p_to_class_id
     and coalesce(v_student.is_active, true) = true
     and v_student.enrollment_status in ('active', 'transferred_in') then
    raise exception 'Learner is already active in the selected class.';
  end if;

  select nullif(to_jsonb(c)->>'school_year', '')
  into v_school_year
  from public.classes c
  where c.id = p_to_class_id;

  insert into public.student_transfers (
    student_id,
    transfer_type,
    from_class_id,
    to_class_id,
    school_year,
    effective_term,
    transfer_date,
    previous_school,
    reason,
    remarks,
    created_by
  )
  values (
    v_student.id,
    'in',
    v_from_class_id,
    p_to_class_id,
    v_school_year,
    nullif(trim(p_effective_term), ''),
    coalesce(p_transfer_date, current_date),
    nullif(trim(p_previous_school), ''),
    nullif(trim(p_reason), ''),
    nullif(trim(p_remarks), ''),
    auth.uid()
  )
  returning id into v_transfer_id;

  update public.students
  set
    class_id = p_to_class_id,
    enrollment_status = 'transferred_in',
    is_active = true,
    transfer_date = coalesce(p_transfer_date, current_date),
    transfer_effective_term = nullif(trim(p_effective_term), ''),
    previous_school = nullif(trim(p_previous_school), ''),
    destination_school = null,
    transfer_reason = nullif(trim(p_reason), ''),
    transferred_at = now()
  where id = v_student.id;

  return v_transfer_id;
end;
$$;

grant execute on function public.transfer_in_student(
  uuid, uuid, date, text, text, text, text
) to authenticated;

-- 6) Transfer history helper
create or replace function public.get_student_transfer_history(p_student_id uuid)
returns setof public.student_transfers
language sql
stable
security definer
set search_path = public, pg_temp
as $$
  select st.*
  from public.student_transfers st
  where st.student_id = p_student_id
    and (
      exists (
        select 1
        from public.user_roles ur
        where ur.user_id = auth.uid()
          and ur.role = 'admin'
      )
      or public.can_manage_student_transfer(st.from_class_id)
      or public.can_manage_student_transfer(st.to_class_id)
    )
  order by st.created_at desc;
$$;

grant execute on function public.get_student_transfer_history(uuid)
to authenticated;

-- 7) RLS
alter table public.student_transfers enable row level security;

drop policy if exists "read manageable student transfers"
on public.student_transfers;

create policy "read manageable student transfers"
on public.student_transfers
for select
to authenticated
using (
  exists (
    select 1
    from public.user_roles ur
    where ur.user_id = auth.uid()
      and ur.role = 'admin'
  )
  or public.can_manage_student_transfer(from_class_id)
  or public.can_manage_student_transfer(to_class_id)
);

revoke insert, update, delete
on public.student_transfers
from authenticated;

grant select
on public.student_transfers
to authenticated;

commit;

-- ============================================================
-- FRONTEND EXAMPLES
-- ============================================================
--
-- Transfer Out:
--
-- await supabase.rpc("transfer_out_student", {
--   p_student_id: studentId,
--   p_transfer_date: "2026-09-14",
--   p_destination_school: "Destination School",
--   p_reason: "Family relocation",
--   p_effective_term: "Term 2",
--   p_remarks: null,
-- });
--
-- Transfer In:
--
-- await supabase.rpc("transfer_in_student", {
--   p_student_id: studentId,
--   p_to_class_id: classId,
--   p_transfer_date: "2026-09-14",
--   p_previous_school: "Previous School",
--   p_effective_term: "Term 2",
--   p_reason: null,
--   p_remarks: null,
-- });
--
-- Current active roster:
--   .eq("is_active", true)
--
-- IMPORTANT:
-- * Do NOT delete learner grade/history rows during Transfer Out.
-- * Missing grades after a transfer must stay NULL/blank, not 0.
-- ============================================================
