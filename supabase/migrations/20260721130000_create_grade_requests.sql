-- =========================================================
-- Grade Requests (Fixed Setup)
-- IMPORTANT:
-- This script deletes the existing public.grade_requests table
-- and recreates it with the correct columns.
-- Run the entire script in the Supabase SQL Editor.
-- =========================================================

create extension if not exists pgcrypto;

-- Remove the old/incomplete table that caused:
-- column "requester_id" does not exist
drop table if exists public.grade_requests cascade;

-- 1. Main table
create table public.grade_requests (
  id uuid primary key default gen_random_uuid(),

  requester_id uuid not null default auth.uid()
    references auth.users(id) on delete cascade,

  subject_teacher_id uuid
    references auth.users(id) on delete set null,

  subject text not null,
  term text not null,
  message text not null,

  priority text not null default 'low'
    check (priority in ('low', 'medium', 'high')),

  status text not null default 'pending'
    check (status in ('pending', 'in_progress', 'completed', 'rejected')),

  attachment_url text,
  attachment_name text,
  teacher_response text,

  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- 2. Helpful indexes
create index grade_requests_requester_id_idx
  on public.grade_requests (requester_id);

create index grade_requests_subject_teacher_id_idx
  on public.grade_requests (subject_teacher_id);

create index grade_requests_status_idx
  on public.grade_requests (status);

create index grade_requests_created_at_idx
  on public.grade_requests (created_at desc);

-- 3. Automatically update updated_at
create or replace function public.set_grade_requests_updated_at()
returns trigger
language plpgsql
security invoker
set search_path = public
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create trigger set_grade_requests_updated_at
before update on public.grade_requests
for each row
execute function public.set_grade_requests_updated_at();

-- 4. Enable Row Level Security
alter table public.grade_requests enable row level security;

-- Class advisers/requesters can read their own requests
create policy "requesters read own grade requests"
on public.grade_requests
for select
to authenticated
using (requester_id = auth.uid());

-- Class advisers/requesters can create requests only for themselves
create policy "requesters create own grade requests"
on public.grade_requests
for insert
to authenticated
with check (
  requester_id = auth.uid()
  and status = 'pending'
);

-- Requesters can edit only their own pending requests
create policy "requesters update own pending grade requests"
on public.grade_requests
for update
to authenticated
using (
  requester_id = auth.uid()
  and status = 'pending'
)
with check (
  requester_id = auth.uid()
  and status = 'pending'
);

-- Assigned subject teachers can read requests sent to them
create policy "subject teachers read assigned grade requests"
on public.grade_requests
for select
to authenticated
using (subject_teacher_id = auth.uid());

-- Assigned subject teachers can update assigned requests
create policy "subject teachers update assigned grade requests"
on public.grade_requests
for update
to authenticated
using (subject_teacher_id = auth.uid())
with check (subject_teacher_id = auth.uid());

-- Admins can manage every request
create policy "admins manage all grade requests"
on public.grade_requests
for all
to authenticated
using (public.is_admin(auth.uid()))
with check (public.is_admin(auth.uid()));

-- 5. Grant table access
grant select, insert, update, delete
on public.grade_requests
to authenticated;

-- =========================================================
-- OPTIONAL ATTACHMENT STORAGE
-- Private bucket: grade-request-attachments
-- Expected file path:
-- <requester-user-id>/<request-id>/<filename>
-- =========================================================

insert into storage.buckets (id, name, public)
values (
  'grade-request-attachments',
  'grade-request-attachments',
  false
)
on conflict (id) do nothing;

drop policy if exists "users upload own grade request attachments"
  on storage.objects;

drop policy if exists "users read own grade request attachments"
  on storage.objects;

drop policy if exists "assigned teachers read grade request attachments"
  on storage.objects;

drop policy if exists "users delete own grade request attachments"
  on storage.objects;

create policy "users upload own grade request attachments"
on storage.objects
for insert
to authenticated
with check (
  bucket_id = 'grade-request-attachments'
  and (storage.foldername(name))[1] = auth.uid()::text
);

create policy "users read own grade request attachments"
on storage.objects
for select
to authenticated
using (
  bucket_id = 'grade-request-attachments'
  and (storage.foldername(name))[1] = auth.uid()::text
);

create policy "assigned teachers read grade request attachments"
on storage.objects
for select
to authenticated
using (
  bucket_id = 'grade-request-attachments'
  and exists (
    select 1
    from public.grade_requests as gr
    where gr.subject_teacher_id = auth.uid()
      and gr.attachment_url = storage.objects.name
  )
);

create policy "users delete own grade request attachments"
on storage.objects
for delete
to authenticated
using (
  bucket_id = 'grade-request-attachments'
  and (storage.foldername(name))[1] = auth.uid()::text
);
