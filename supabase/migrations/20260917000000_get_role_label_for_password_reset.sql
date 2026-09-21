-- =============================================================================
-- Migration: get_role_label_for_password_reset
-- Purpose  : Expose a SECURITY DEFINER RPC that unauthenticated (anon) users
--            can call from the Forgot Password page to look up the display-role
--            label (Admin / Class Adviser / Subject Teacher) for a given email.
--            Only the label string is returned — no passwords, IDs, or PII.
-- =============================================================================

begin;

-- ---------------------------------------------------------------------------
-- Function: public.get_role_label_for_reset(p_email text)
-- Returns : text  →  'Admin' | 'Class Adviser' | 'Subject Teacher' | 'User'
--                     NULL if the email is not found in profiles
-- Security: SECURITY DEFINER so it runs as the DB owner and bypasses RLS.
--           The anon role is granted EXECUTE only on this specific function.
-- ---------------------------------------------------------------------------
create or replace function public.get_role_label_for_reset(p_email text)
returns text
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id  uuid;
  v_tt       text;
  v_role     text;
  v_label    text;
begin
  -- 1. Resolve the profile by email
  select id, teacher_type
  into   v_user_id, v_tt
  from   public.profiles
  where  email = lower(trim(p_email))
  limit  1;

  -- Email not registered → return NULL so the caller can show a proper error
  if v_user_id is null then
    return null;
  end if;

  -- 2. Check user_roles table for the 'admin' role first
  select role
  into   v_role
  from   public.user_roles
  where  user_id = v_user_id
  limit  1;

  if v_role = 'admin' then
    return 'Admin';
  end if;

  -- 3. Fall back to teacher_type from profiles
  v_label := case lower(coalesce(v_tt, ''))
    when 'class_adviser'   then 'Class Adviser'
    when 'subject_teacher' then 'Subject Teacher'
    else                        coalesce(nullif(trim(v_tt), ''), 'User')
  end;

  return v_label;
end;
$$;

-- Grant execute to the anonymous role so unauthenticated callers can use it
grant execute on function public.get_role_label_for_reset(text) to anon;
-- Authenticated users can also call it (e.g. after sign-in edge cases)
grant execute on function public.get_role_label_for_reset(text) to authenticated;

commit;

notify pgrst, 'reload schema';

