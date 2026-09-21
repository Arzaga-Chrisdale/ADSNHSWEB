-- SIGLA password-reset database support
-- Safe to run more than once.
--
-- IMPORTANT:
-- This SQL supports the role lookup used by auth.tsx.
-- It CANNOT create or replace the HTTP Edge Function route:
--   /functions/v1/reset-user-password
-- Your current auth.tsx still calls that Edge Function when saving
-- the new password, so that function must exist if auth.tsx is unchanged.

begin;

-- -------------------------------------------------------------------
-- 1. Server-side role lookup used by the forgot-password email step.
--    auth.tsx calls:
--      supabase.rpc("get_role_label_for_reset", { p_email: ... })
-- -------------------------------------------------------------------

create or replace function public.get_role_label_for_reset(
  p_email text
)
returns text
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  v_email text;
  v_user_id uuid;
  v_teacher_type text;
begin
  v_email := lower(trim(coalesce(p_email, '')));

  if v_email = '' then
    return null;
  end if;

  select u.id
  into v_user_id
  from auth.users as u
  where lower(coalesce(u.email, '')) = v_email
  limit 1;

  if v_user_id is null then
    return null;
  end if;

  -- Admin role has priority.
  if exists (
    select 1
    from public.user_roles ur
    where ur.user_id = v_user_id
      and ur.role = 'admin'
  ) then
    return 'Admin';
  end if;

  select lower(trim(coalesce(p.teacher_type, '')))
  into v_teacher_type
  from public.profiles p
  where p.id = v_user_id
  limit 1;

  if v_teacher_type = 'class_adviser' then
    return 'Class Adviser';
  end if;

  if v_teacher_type = 'subject_teacher' then
    return 'Subject Teacher';
  end if;

  return null;
end;
$$;

revoke all on function public.get_role_label_for_reset(text) from public;
grant execute on function public.get_role_label_for_reset(text) to anon;
grant execute on function public.get_role_label_for_reset(text) to authenticated;


-- -------------------------------------------------------------------
-- 2. Optional reset-challenge table.
--    Your current auth.tsx generates/verifies the OTP in the browser,
--    so this table is not required by that exact frontend code.
--    It is kept ready for a future server-verified OTP flow.
-- -------------------------------------------------------------------

create table if not exists public.password_reset_challenges (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  email text not null,
  role_label text not null,
  otp_hash text not null,
  expires_at timestamptz not null,
  attempts integer not null default 0 check (attempts >= 0),
  verified_at timestamptz,
  reset_token_hash text,
  reset_token_expires_at timestamptz,
  used_at timestamptz,
  created_at timestamptz not null default now()
);

alter table public.password_reset_challenges
  enable row level security;

-- Do not expose challenge data directly to browser clients.
revoke all on table public.password_reset_challenges
  from anon, authenticated;

create index if not exists password_reset_challenges_email_created_idx
  on public.password_reset_challenges (lower(email), created_at desc);

create index if not exists password_reset_challenges_user_created_idx
  on public.password_reset_challenges (user_id, created_at desc);

create index if not exists password_reset_challenges_active_idx
  on public.password_reset_challenges (user_id, used_at, created_at desc);

commit;

notify pgrst, 'reload schema';
