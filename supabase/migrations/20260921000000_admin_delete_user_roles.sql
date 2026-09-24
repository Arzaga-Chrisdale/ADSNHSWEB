-- SIGLA: Combined Admin -> All Teachers / Users -> Delete Teacher / User fix
--
-- This migration replaces the original admin_delete_user RPC and repairs
-- refresh_grade_request_batch_status so cascading deletion does not create
-- orphan grade_request_history records.
--
-- The original admin_delete_user protections are preserved:
-- * only a signed-in administrator may delete accounts;
-- * the current administrator cannot delete their own account;
-- * school-form submissions reviewed by the target user are retained,
--   with reviewed_by cleared before that user is deleted.
--
-- Apply this migration AFTER 20260921000000_admin_delete_user_roles.sql and
-- AFTER any earlier 20260924000000_admin_delete_teacher_user.sql migration.
-- If an earlier migration was already applied, do not edit its migration
-- history: apply this as the NEW migration named above instead.
--
-- CAUTION: Successful deletion permanently removes the user and any data
-- configured with ON DELETE CASCADE. Back up your database and test with
-- an expendable account before using in production.

begin;

-- 1) Guard the Grade Request trigger during cascade deletes.
-- When a batch is deleted, its subjects may be deleted by cascade. Their
-- AFTER DELETE trigger can still execute after the parent batch is gone.
-- Never insert history or update the status of a missing parent batch.
create or replace function public.refresh_grade_request_batch_status()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_batch_id uuid;
  v_total integer;
  v_submitted integer;
  v_old_status text;
  v_new_status text;
begin
  v_batch_id := case
    when tg_op = 'DELETE' then old.batch_id
    else new.batch_id
  end;

  select batch.status
    into v_old_status
  from public.grade_request_batches as batch
  where batch.id = v_batch_id;

  if not found then
    if tg_op = 'DELETE' then
      return old;
    end if;
    return new;
  end if;

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

  -- Do not change finalized batches.
  update public.grade_request_batches as batch
  set status = v_new_status,
      completed_at = case
        when v_new_status = 'completed'
          then coalesce(batch.completed_at, now())
        else null
      end
  where batch.id = v_batch_id
    and batch.is_finalized = false;

  -- The parent may have disappeared during cascade deletion. Recheck
  -- before writing a history row referencing the parent's foreign key.
  if v_old_status is distinct from v_new_status
     and exists (
       select 1
       from public.grade_request_batches as batch
       where batch.id = v_batch_id
         and batch.is_finalized = false
         and batch.status = v_new_status
     ) then
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
      pg_catalog.jsonb_build_object(
        'submitted_subjects', v_submitted,
        'total_subjects', v_total
      )
    );
  end if;

  if tg_op = 'DELETE' then
    return old;
  end if;
  return new;
end;
$$;

-- 2) Merge the original admin_delete_user RPC with child-first Grade Request
-- cleanup. Delete the affected subjects and history BEFORE removing their
-- parent batches; this prevents the reported batch_id foreign-key error.
create or replace function public.admin_delete_user(target_user_id uuid)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  caller_id uuid := auth.uid();
begin
  -- Preserve the original account-deletion authorization checks.
  if caller_id is null then
    raise exception 'You must be signed in to delete a user.';
  end if;

  if not public.is_admin(caller_id) then
    raise exception 'Only an administrator can delete teacher/user accounts.';
  end if;

  if target_user_id is null then
    raise exception 'A target user ID is required.';
  end if;

  if target_user_id = caller_id then
    raise exception 'You cannot delete the administrator account that is currently signed in.';
  end if;

  if not exists (
    select 1 from auth.users where id = target_user_id
  ) then
    raise exception 'The selected user account no longer exists.';
  end if;

  -- Delete subjects while their Grade Request batches still exist. Include
  -- batches owned by the user and batches for their owned advisory classes.
  delete from public.grade_request_subjects as request
  where request.batch_id in (
    select batch.id
    from public.grade_request_batches as batch
    where batch.adviser_id = target_user_id
       or exists (
         select 1
         from public.classes as owned_class
         where owned_class.id = batch.advisory_class_id
           and owned_class.teacher_id = target_user_id
       )
  );

  -- The subjects' DELETE triggers may have added fresh status history.
  -- Remove all history for these batches after their subjects are deleted.
  delete from public.grade_request_history as history
  where history.batch_id in (
    select batch.id
    from public.grade_request_batches as batch
    where batch.adviser_id = target_user_id
       or exists (
         select 1
         from public.classes as owned_class
         where owned_class.id = batch.advisory_class_id
           and owned_class.teacher_id = target_user_id
       )
  );

  -- The request subjects and history are now gone; remove the parent batches.
  delete from public.grade_request_batches as batch
  where batch.adviser_id = target_user_id
     or exists (
       select 1
       from public.classes as owned_class
       where owned_class.id = batch.advisory_class_id
         and owned_class.teacher_id = target_user_id
     );

  -- Preserve OTHER users' School Form submissions previously reviewed by
  -- this user. Clear only the reviewer's reference, as in the original RPC.
  if to_regclass('public.school_form_submissions') is not null then
    execute '
      update public.school_form_submissions
         set reviewed_by = null
       where reviewed_by = $1
    ' using target_user_id;
  end if;

  -- Delete the Supabase Auth account last. Existing foreign-key CASCADE /
  -- SET NULL rules manage remaining related records as configured in the DB.
  delete from auth.users where id = target_user_id;

  if not found then
    raise exception 'The selected user account could not be deleted.';
  end if;
end;
$$;

-- 3) Only signed-in users can call the RPC; the function itself checks
-- that the caller has the administrator role.
revoke all on function public.admin_delete_user(uuid) from public;
revoke all on function public.admin_delete_user(uuid) from anon;
grant execute on function public.admin_delete_user(uuid) to authenticated;

commit;

-- Ask the Supabase API to reload its database function definitions.
notify pgrst, 'reload schema';

-- Verification only: these checks do not delete accounts.
-- Both columns should return TRUE after successful installation.
select
  coalesce(
    position(
      'if not found' in lower(
        pg_get_functiondef(
          to_regprocedure('public.refresh_grade_request_batch_status()')
        )
      )
    ) > 0,
    false
  ) as grade_request_history_guard_installed,
  coalesce(
    position(
      'delete from public.grade_request_subjects' in lower(
        pg_get_functiondef(to_regprocedure('public.admin_delete_user(uuid)'))
      )
    ) > 0,
    false
  ) as child_first_user_deletion_installed;
