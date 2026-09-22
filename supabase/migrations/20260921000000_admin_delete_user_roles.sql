-- SIGLA Admin: delete a teacher/user account
-- Adds the RPC used by src/routes/_authenticated/admin.tsx:
--   supabase.rpc('admin_delete_user', { target_user_id: profile.id })
--
-- Safe to run more than once.

begin;

create or replace function public.admin_delete_user(
  target_user_id uuid
)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  caller_id uuid := auth.uid();
begin
  -- Must be signed in.
  if caller_id is null then
    raise exception 'You must be signed in to delete a user.';
  end if;

  -- Only an administrator can use this RPC.
  if not public.is_admin(caller_id) then
    raise exception 'Only an administrator can delete teacher/user accounts.';
  end if;

  if target_user_id is null then
    raise exception 'A target user ID is required.';
  end if;

  -- Protect the currently signed-in administrator.
  if target_user_id = caller_id then
    raise exception 'You cannot delete the administrator account that is currently signed in.';
  end if;

  if not exists (
    select 1
    from auth.users
    where id = target_user_id
  ) then
    raise exception 'The selected user account no longer exists.';
  end if;

  -- school_form_submissions.reviewed_by was created without ON DELETE SET NULL.
  -- Clear it first so a user who reviewed another class form does not block deletion.
  if to_regclass('public.school_form_submissions') is not null then
    execute
      'update public.school_form_submissions
         set reviewed_by = null
       where reviewed_by = $1'
      using target_user_id;
  end if;

  -- Delete the Supabase Auth account.
  -- Existing ON DELETE CASCADE / SET NULL foreign keys remove or detach
  -- the user's profile, roles, classes, learners, grades, requests,
  -- notifications, sessions, and other related records as configured
  -- by the project migrations.
  delete from auth.users
  where id = target_user_id;

  if not found then
    raise exception 'The selected user account could not be deleted.';
  end if;
end;
$$;

revoke all on function public.admin_delete_user(uuid) from public;
revoke all on function public.admin_delete_user(uuid) from anon;
grant execute on function public.admin_delete_user(uuid) to authenticated;

commit;

-- Tell PostgREST/Supabase API to refresh its RPC schema cache immediately.
notify pgrst, 'reload schema';
