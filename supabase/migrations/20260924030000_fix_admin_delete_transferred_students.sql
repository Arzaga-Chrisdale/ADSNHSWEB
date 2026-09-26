-- SIGLA: Admin -> All Learners / Students -> Delete (Transfer In / Out fix)
--
-- The reported error is:
--   update or delete on table "students" violates foreign key constraint
--   "student_transfers_student_id_fkey" on table "student_transfers"
--
-- BEFORE RUNNING: back up your Supabase database and test with an
-- expendable learner. Deleting a learner is PERMANENT. This migration
-- intentionally removes that learner's Transfer In / Transfer Out history.
-- Other learner-related rows are handled by their existing FK rules.
--
-- This is a NEW migration. Keep all prior migration files unchanged.
-- Safe to execute again if already installed.

begin;

-- 1) Ensure deleting a learner also deletes their transfer history.
-- This fixes cascading deletes initiated by other existing workflows too.
-- The existing constraint name is taken from the reported database error.
do $migration$
declare
  v_delete_rule "char";
begin
  if to_regclass('public.student_transfers') is null then
    raise exception 'public.student_transfers is missing; apply the Transfer In / Out schema first.';
  end if;

  select fk.confdeltype
    into v_delete_rule
  from pg_catalog.pg_constraint as fk
  where fk.conrelid = 'public.student_transfers'::regclass
    and fk.conname = 'student_transfers_student_id_fkey'
    and fk.contype = 'f';

  -- Only rebuild when necessary, to avoid an unnecessary lock on reruns.
  if v_delete_rule is distinct from 'c' then
    alter table public.student_transfers
      drop constraint if exists student_transfers_student_id_fkey;

    alter table public.student_transfers
      add constraint student_transfers_student_id_fkey
      foreign key (student_id)
      references public.students(id)
      on delete cascade;
  end if;
end;
$migration$;

-- 2) Give the Admin Delete icon one secure, atomic database operation.
-- No client-side deletion of transfer rows is needed; normal teachers
-- cannot call this function because it verifies the administrator role.
create or replace function public.admin_delete_student(p_student_id uuid)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_caller_id uuid := auth.uid();
begin
  if v_caller_id is null then
    raise exception 'You must be signed in to delete a learner.';
  end if;

  if not public.is_admin(v_caller_id) then
    raise exception 'Only an administrator can delete learners.';
  end if;

  if p_student_id is null then
    raise exception 'Select a learner to delete.';
  end if;

  -- Prevent deletion of the same learner by two concurrent requests.
  perform 1
  from public.students
  where id = p_student_id
  for update;

  if not found then
    raise exception 'This learner no longer exists.';
  end if;

  -- Explicitly remove transfer history before deleting its parent learner.
  -- With the FK fix above, transfers also cascade for other deletion paths.
  delete from public.student_transfers
  where student_id = p_student_id;

  -- Existing ON DELETE CASCADE / SET NULL foreign keys handle the other
  -- learner records in accordance with the database's existing rules.
  delete from public.students
  where id = p_student_id;

  if not found then
    raise exception 'The learner could not be deleted.';
  end if;
end;
$$;

revoke all on function public.admin_delete_student(uuid) from public;
revoke all on function public.admin_delete_student(uuid) from anon;
grant execute on function public.admin_delete_student(uuid) to authenticated;

commit;

notify pgrst, 'reload schema';

-- 3) Non-destructive verification: both columns should be true.
select
  exists (
    select 1
    from pg_catalog.pg_constraint as fk
    where fk.conrelid = 'public.student_transfers'::regclass
      and fk.conname = 'student_transfers_student_id_fkey'
      and fk.confdeltype = 'c'  -- c = ON DELETE CASCADE
  ) as transfer_history_cascade_enabled,
  to_regprocedure('public.admin_delete_student(uuid)') is not null
    as admin_delete_student_installed;
