-- Student Credential Request System
-- Updated version:
-- Removed "Submit Student Credentials to Admin"
-- Kept only "Request Student Credentials"

create extension if not exists "pgcrypto";

create table if not exists public.student_credential_requests (
  id uuid primary key default gen_random_uuid(),

  student_id uuid not null references public.students(id) on delete cascade,
  requester_id uuid not null default auth.uid() references auth.users(id) on delete cascade,

  requesting_class_id uuid references public.classes(id) on delete set null,
  previous_class_id uuid references public.classes(id) on delete set null,

  reason text,

  -- Main status column used by the credential request UI.
  status text not null default 'pending'
    check (status in ('pending', 'completed', 'Pending', 'Completed')),

  request_date timestamptz not null default now(),

  processed_by uuid references auth.users(id) on delete set null,
  processed_at timestamptz,

  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_student_credential_requests_student_id
on public.student_credential_requests(student_id);

create index if not exists idx_student_credential_requests_requester_id
on public.student_credential_requests(requester_id);

create index if not exists idx_student_credential_requests_status
on public.student_credential_requests(status);

alter table public.student_credential_requests enable row level security;

drop policy if exists "Admins can manage all credential requests"
on public.student_credential_requests;

drop policy if exists "Teachers can insert own credential requests"
on public.student_credential_requests;

drop policy if exists "Teachers can view own credential requests"
on public.student_credential_requests;

create policy "Admins can manage all credential requests"
on public.student_credential_requests
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

create policy "Teachers can insert own credential requests"
on public.student_credential_requests
for insert
to authenticated
with check (requester_id = auth.uid());

create policy "Teachers can view own credential requests"
on public.student_credential_requests
for select
to authenticated
using (requester_id = auth.uid());

-- This lets PostgREST/Supabase API detect the table immediately.
notify pgrst, 'reload schema';
