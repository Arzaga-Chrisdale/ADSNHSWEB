-- ============================================================
-- SIGLA: MERGED Admin sessions + All Teachers/Users online status
--
-- Keeps ALL existing Admin presence features:
--   * Class Adviser / Subject Teacher duplicate-device rules remain unchanged.
--   * Up to 3 DISTINCT Admin accounts can be online simultaneously.
--   * The same Admin can open additional devices without consuming a slot.
--   * The 4th distinct Admin is refused Admin-panel access.
--   * Admin presence expires after 90 seconds without a heartbeat.
--
-- Adds:
--   * Admin-only RPC returning online Class Adviser and Subject Teacher IDs.
--   * Uses the existing teacher active_account_sessions heartbeat, not
--     the Admin presence table, and never changes teacher device rules.
--
-- Apply as a single SQL migration IF the original Admin presence migration
-- has NOT already been applied. If it HAS already been applied, use the
-- smaller add-only migration provided separately instead.
-- Safe to rerun on a compatible existing SIGLA schema.
-- ============================================================

begin;

-- ============================================================
-- 1) ORIGINAL: Admin presence table, indexes, and SELECT RLS policy
-- ============================================================

create table if not exists public.admin_active_sessions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  device_id text not null,
  device_name text,
  browser_name text,
  last_seen_at timestamptz not null default now(),
  created_at timestamptz not null default now()
);

create unique index if not exists
  admin_active_sessions_user_device_uidx
on public.admin_active_sessions (user_id, device_id);

create index if not exists
  admin_active_sessions_last_seen_idx
on public.admin_active_sessions (last_seen_at desc);

create index if not exists
  admin_active_sessions_user_last_seen_idx
on public.admin_active_sessions (user_id, last_seen_at desc);

alter table public.admin_active_sessions enable row level security;

drop policy if exists "Admins can view Admin active sessions"
  on public.admin_active_sessions;

create policy "Admins can view Admin active sessions"
on public.admin_active_sessions
for select
to authenticated
using (
  exists (
    select 1
    from public.user_roles ur
    where ur.user_id = auth.uid()
      and ur.role::text = 'admin'
  )
);

-- Direct INSERT / UPDATE / DELETE access is intentionally not granted.
-- Admin presence changes go through the SECURITY DEFINER RPCs below.

-- ============================================================
-- 2) ORIGINAL: Claim Admin session, enforce 3 DISTINCT Admin limit
-- ============================================================

create or replace function public.claim_admin_session(
  p_device_id text,
  p_device_name text default null,
  p_browser_name text default null
)
returns table (
  allowed boolean,
  online_count integer,
  max_admins integer
)
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_max_admins integer := 3;
  v_online_count integer := 0;
  v_user_is_already_online boolean := false;
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if nullif(trim(p_device_id), '') is null then
    raise exception 'A device id is required.';
  end if;

  if not exists (
    select 1
    from public.user_roles ur
    where ur.user_id = v_user_id
      and ur.role::text = 'admin'
  ) then
    raise exception 'Only an Administrator can claim an Admin session.';
  end if;

  -- Serialize slot claims so two Admins cannot both claim the last slot.
  perform pg_advisory_xact_lock(
    hashtext('sigla-admin-online-limit')::bigint
  );

  -- A browser that has not sent a heartbeat for 90 seconds is no longer online.
  delete from public.admin_active_sessions s
  where s.last_seen_at < now() - interval '90 seconds';

  select exists (
    select 1
    from public.admin_active_sessions s
    where s.user_id = v_user_id
      and s.last_seen_at >= now() - interval '90 seconds'
  )
  into v_user_is_already_online;

  select count(distinct s.user_id)::integer
  into v_online_count
  from public.admin_active_sessions s
  where s.last_seen_at >= now() - interval '90 seconds';

  -- Existing online Admins do not consume another slot when refreshing
  -- or opening another browser/device.
  if not v_user_is_already_online and v_online_count >= v_max_admins then
    return query
    select false, v_online_count, v_max_admins;
    return;
  end if;

  insert into public.admin_active_sessions (
    user_id,
    device_id,
    device_name,
    browser_name,
    last_seen_at
  )
  values (
    v_user_id,
    trim(p_device_id),
    nullif(trim(p_device_name), ''),
    nullif(trim(p_browser_name), ''),
    now()
  )
  on conflict (user_id, device_id)
  do update set
    device_name = excluded.device_name,
    browser_name = excluded.browser_name,
    last_seen_at = now();

  select count(distinct s.user_id)::integer
  into v_online_count
  from public.admin_active_sessions s
  where s.last_seen_at >= now() - interval '90 seconds';

  return query
  select true, v_online_count, v_max_admins;
end;
$$;

revoke all on function public.claim_admin_session(text, text, text)
  from public, anon, authenticated;

grant execute on function public.claim_admin_session(text, text, text)
  to authenticated;

-- ============================================================
-- 3) ORIGINAL: Release an Admin device session
-- ============================================================

create or replace function public.release_admin_session(
  p_device_id text
)
returns void
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then
    return;
  end if;

  delete from public.admin_active_sessions s
  where s.user_id = v_user_id
    and s.device_id = p_device_id;
end;
$$;

revoke all on function public.release_admin_session(text)
  from public, anon, authenticated;

grant execute on function public.release_admin_session(text)
  to authenticated;

-- ============================================================
-- 4) ORIGINAL: Retrieve online Admins (including public profile fields)
-- ============================================================

create or replace function public.get_online_admins()
returns table (
  user_id uuid,
  full_name text,
  email text,
  avatar_url text,
  last_seen_at timestamptz
)
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then
    raise exception 'You must be signed in.';
  end if;

  if not exists (
    select 1
    from public.user_roles ur
    where ur.user_id = v_user_id
      and ur.role::text = 'admin'
  ) then
    raise exception 'Only an Administrator can view online Admins.';
  end if;

  delete from public.admin_active_sessions s
  where s.last_seen_at < now() - interval '90 seconds';

  return query
  select
    s.user_id,
    p.full_name,
    p.email,
    p.avatar_url,
    max(s.last_seen_at) as last_seen_at
  from public.admin_active_sessions s
  join public.user_roles ur
    on ur.user_id = s.user_id
   and ur.role::text = 'admin'
  left join public.profiles p
    on p.id = s.user_id
  where s.last_seen_at >= now() - interval '90 seconds'
  group by
    s.user_id,
    p.full_name,
    p.email,
    p.avatar_url
  order by max(s.last_seen_at) desc;
end;
$$;

revoke all on function public.get_online_admins()
  from public, anon, authenticated;

grant execute on function public.get_online_admins()
  to authenticated;

-- ============================================================
-- 5) NEW: Online Class Adviser / Subject Teacher IDs for Admin UI
-- ============================================================
-- Uses the EXISTING teacher heartbeat table. It does not modify or grant
-- client access to teacher sessions. Return IDs only to verified Admins.
-- The role check matches the existing Admin presence RPCs above.

create or replace function public.get_online_teacher_ids()
returns table (user_id uuid)
language plpgsql
security definer
set search_path = ''
as $$
begin
  if auth.uid() is null or not exists (
    select 1
    from public.user_roles ur
    where ur.user_id = auth.uid()
      and ur.role::text = 'admin'
  ) then
    raise exception 'Only an administrator can view teacher online status.'
      using errcode = '42501';
  end if;

  return query
  select distinct session_row.user_id
  from public.active_account_sessions as session_row
  join public.profiles as profile
    on profile.id = session_row.user_id
  where profile.teacher_type in ('class_adviser', 'subject_teacher')
    and session_row.last_seen_at >= now() - interval '90 seconds';
end;
$$;

revoke all on function public.get_online_teacher_ids()
  from public, anon, authenticated;

grant execute on function public.get_online_teacher_ids()
  to authenticated;

commit;

notify pgrst, 'reload schema';
