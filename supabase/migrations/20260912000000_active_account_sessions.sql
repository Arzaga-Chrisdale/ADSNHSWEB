-- ============================================================
-- Active Account Notice (notification only)
-- Applies in the frontend to:
--   - Class Adviser
--   - Subject Teacher
--
-- IMPORTANT:
-- This migration DOES NOT force logout and DOES NOT invalidate sessions.
-- It only stores lightweight presence rows so the UI can notify a user
-- when the same account is active on another device/browser.
-- ============================================================

create table if not exists public.active_account_sessions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  device_id text not null,
  device_name text,
  browser_name text,
  last_seen_at timestamptz not null default now(),
  created_at timestamptz not null default now(),

  constraint active_account_sessions_user_device_unique
    unique (user_id, device_id)
);

create index if not exists active_account_sessions_user_last_seen_idx
  on public.active_account_sessions (user_id, last_seen_at desc);

alter table public.active_account_sessions enable row level security;

drop policy if exists "Users can view own active sessions"
  on public.active_account_sessions;

create policy "Users can view own active sessions"
  on public.active_account_sessions
  for select
  to authenticated
  using (auth.uid() = user_id);


drop policy if exists "Users can create own active sessions"
  on public.active_account_sessions;

create policy "Users can create own active sessions"
  on public.active_account_sessions
  for insert
  to authenticated
  with check (auth.uid() = user_id);


drop policy if exists "Users can update own active sessions"
  on public.active_account_sessions;

create policy "Users can update own active sessions"
  on public.active_account_sessions
  for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);


drop policy if exists "Users can delete own active sessions"
  on public.active_account_sessions;

create policy "Users can delete own active sessions"
  on public.active_account_sessions
  for delete
  to authenticated
  using (auth.uid() = user_id);


grant select, insert, update, delete
  on public.active_account_sessions
  to authenticated;

revoke all
  on public.active_account_sessions
  from anon;


-- Realtime lets Device A notice Device B without reloading the page.
-- Add the table only if it has not already been added.
do $$
begin
  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'active_account_sessions'
  ) then
    alter publication supabase_realtime
      add table public.active_account_sessions;
  end if;
end
$$;


-- Helpful for realtime update/delete payloads.
alter table public.active_account_sessions replica identity full;

notify pgrst, 'reload schema';
