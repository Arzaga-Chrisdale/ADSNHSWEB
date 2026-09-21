-- ============================================================
-- Admin online presence + maximum 3 active Admin accounts.
--
-- Behavior:
--   * Class Adviser / Subject Teacher duplicate-device notice is unchanged.
--   * Admin accounts use this separate presence table.
--   * Up to 3 DISTINCT Admin accounts may be online at the same time.
--   * A 4th Admin account is refused Admin-panel access.
--   * The same already-online Admin may refresh/open another browser/device
--     without consuming another Admin slot.
--   * Presence expires automatically after 90 seconds without a heartbeat.
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
-- Presence changes go through the SECURITY DEFINER functions below.

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

notify pgrst, 'reload schema';
