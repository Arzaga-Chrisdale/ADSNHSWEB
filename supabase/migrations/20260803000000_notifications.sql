-- Role-targeted notifications for Admin, Class Adviser, and Subject Teacher.
-- Safe to run more than once.

create table if not exists public.notifications (
  id uuid primary key default gen_random_uuid(),
  recipient_id uuid not null references auth.users(id) on delete cascade,
  actor_id uuid references auth.users(id) on delete set null,
  title text not null check (char_length(trim(title)) between 1 and 160),
  message text not null check (char_length(trim(message)) between 1 and 2000),
  category text not null default 'system' check (
    category in (
      'account', 'assignment', 'learner', 'grade_submission',
      'grade_review', 'form_submission', 'form_review',
      'announcement', 'security', 'system'
    )
  ),
  related_path text,
  event_key text,
  metadata jsonb not null default '{}'::jsonb,
  read_at timestamptz,
  created_at timestamptz not null default now(),
  constraint notifications_related_path_check check (
    related_path is null or related_path like '/%'
  )
);

create index if not exists notifications_recipient_created_idx
  on public.notifications (recipient_id, created_at desc);

create index if not exists notifications_recipient_unread_idx
  on public.notifications (recipient_id, created_at desc)
  where read_at is null;

create unique index if not exists notifications_recipient_event_key_idx
  on public.notifications (recipient_id, event_key)
  where event_key is not null;

alter table public.notifications enable row level security;

drop policy if exists "users read own notifications" on public.notifications;
create policy "users read own notifications"
on public.notifications for select
to authenticated
using (recipient_id = auth.uid());

drop policy if exists "users update own notifications" on public.notifications;
create policy "users update own notifications"
on public.notifications for update
to authenticated
using (recipient_id = auth.uid())
with check (recipient_id = auth.uid());

drop policy if exists "users delete own notifications" on public.notifications;
create policy "users delete own notifications"
on public.notifications for delete
to authenticated
using (recipient_id = auth.uid());

-- There is intentionally no client INSERT policy. Notifications are created by
-- trusted database workflows so users cannot send arbitrary alerts to others.

revoke all on table public.notifications from anon;
grant select, update, delete on table public.notifications to authenticated;
grant all on table public.notifications to service_role;

-- Internal helper used by trusted SECURITY DEFINER workflows and triggers.
create or replace function public.create_notification(
  p_recipient_id uuid,
  p_title text,
  p_message text,
  p_category text default 'system',
  p_related_path text default null,
  p_actor_id uuid default null,
  p_event_key text default null,
  p_metadata jsonb default '{}'::jsonb
) returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  v_id uuid;
begin
  if p_recipient_id is null then
    raise exception 'A notification recipient is required.';
  end if;

  insert into public.notifications (
    recipient_id, actor_id, title, message, category,
    related_path, event_key, metadata
  ) values (
    p_recipient_id,
    p_actor_id,
    trim(p_title),
    trim(p_message),
    coalesce(nullif(trim(p_category), ''), 'system'),
    nullif(trim(p_related_path), ''),
    nullif(trim(p_event_key), ''),
    coalesce(p_metadata, '{}'::jsonb)
  )
  on conflict (recipient_id, event_key) where event_key is not null
  do nothing
  returning id into v_id;

  return v_id;
end;
$$;

revoke all on function public.create_notification(
  uuid, text, text, text, text, uuid, text, jsonb
) from public, anon, authenticated;

create or replace function public.mark_notification_read(p_notification_id uuid)
returns void
language sql
security invoker
set search_path = public
as $$
  update public.notifications
  set read_at = coalesce(read_at, now())
  where id = p_notification_id
    and recipient_id = auth.uid();
$$;

create or replace function public.mark_all_notifications_read()
returns integer
language plpgsql
security invoker
set search_path = public
as $$
declare
  v_count integer;
begin
  update public.notifications
  set read_at = now()
  where recipient_id = auth.uid()
    and read_at is null;
  get diagnostics v_count = row_count;
  return v_count;
end;
$$;

revoke all on function public.mark_notification_read(uuid) from public;
revoke all on function public.mark_all_notifications_read() from public;
grant execute on function public.mark_notification_read(uuid) to authenticated;
grant execute on function public.mark_all_notifications_read() to authenticated;

-- Admin-only announcement helper. A null recipient sends to every active user.
create or replace function public.send_system_announcement(
  p_title text,
  p_message text,
  p_recipient_id uuid default null,
  p_related_path text default null
) returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  v_recipient record;
  v_count integer := 0;
begin
  if not public.is_admin(auth.uid()) then
    raise exception 'Admin access required.';
  end if;

  for v_recipient in
    select user_id
    from public.user_roles
    where p_recipient_id is null or user_id = p_recipient_id
  loop
    perform public.create_notification(
      v_recipient.user_id,
      p_title,
      p_message,
      'announcement',
      p_related_path,
      auth.uid(),
      null,
      '{}'::jsonb
    );
    v_count := v_count + 1;
  end loop;

  return v_count;
end;
$$;

revoke all on function public.send_system_announcement(text, text, uuid, text)
  from public;
grant execute on function public.send_system_announcement(text, text, uuid, text)
  to authenticated;

-- Automatically notify Admins when an Adviser submits/resubmits forms and
-- notify that Adviser when an Admin approves or returns the submission.
create or replace function public.notify_school_form_submission_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin record;
  v_class_label text;
begin
  v_class_label := concat_ws(' - ', new.grade_level, nullif(new.section, ''));

  if tg_op = 'INSERT' or (
    tg_op = 'UPDATE'
    and new.status = 'submitted'
    and old.status is distinct from new.status
  ) then
    for v_admin in
      select ur.user_id
      from public.user_roles ur
      where ur.role = 'admin'
    loop
      perform public.create_notification(
        v_admin.user_id,
        case when tg_op = 'INSERT'
          then 'School Forms Submitted'
          else 'School Forms Resubmitted'
        end,
        concat(
          coalesce(nullif(new.adviser_name, ''), 'A Class Adviser'),
          ' submitted the school forms for ', v_class_label, '.'
        ),
        'form_submission',
        '/admin',
        new.adviser_id,
        concat('school-form:', new.id, ':submitted:', extract(epoch from new.submitted_at)),
        jsonb_build_object(
          'submission_id', new.id,
          'class_id', new.class_id,
          'status', new.status
        )
      );
    end loop;
  elsif tg_op = 'UPDATE'
    and new.status in ('approved', 'returned')
    and old.status is distinct from new.status
  then
    perform public.create_notification(
      new.adviser_id,
      case new.status
        when 'approved' then 'School Forms Approved'
        else 'School Forms Returned'
      end,
      case new.status
        when 'approved' then concat('Your school forms for ', v_class_label, ' were approved.')
        else concat(
          'Your school forms for ', v_class_label,
          ' were returned for correction.',
          case when nullif(trim(new.admin_remarks), '') is not null
            then concat(' Remarks: ', trim(new.admin_remarks))
            else ''
          end
        )
      end,
      'form_review',
      '/school-forms',
      new.reviewed_by,
      concat('school-form:', new.id, ':', new.status, ':', extract(epoch from new.reviewed_at)),
      jsonb_build_object(
        'submission_id', new.id,
        'class_id', new.class_id,
        'status', new.status,
        'remarks', new.admin_remarks
      )
    );
  end if;

  return new;
end;
$$;

revoke all on function public.notify_school_form_submission_change()
  from public, anon, authenticated;

drop trigger if exists school_form_submission_notifications
  on public.school_form_submissions;
create trigger school_form_submission_notifications
after insert or update of status on public.school_form_submissions
for each row execute function public.notify_school_form_submission_change();

-- Add the table to Supabase Realtime once. RLS still limits delivered rows.
do $$
begin
  if not exists (
    select 1
    from pg_publication_tables
    where pubname = 'supabase_realtime'
      and schemaname = 'public'
      and tablename = 'notifications'
  ) then
    alter publication supabase_realtime add table public.notifications;
  end if;
end;
$$;

