-- Student credential-request notifications for Admins and Class Advisers.
--
-- Run after:
--   1. 20260725000000_student_credential_requests.sql
--   2. 20260803000000_notifications.sql
--   3. 20260823130000_sf1_credential_review_flow.sql
--
-- Safe to run more than once.

create or replace function public.notify_student_credential_request_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_admin record;
  v_student_name text;
  v_student_data jsonb;
  v_request_data jsonb;
  v_requester_name text;
  v_new_status text;
  v_old_status text;
  v_notification_title text;
  v_notification_message text;
  v_actor_id uuid;
begin
  v_request_data := to_jsonb(new);

  -- student_id may be NULL when the adviser initially submits only a name.
  -- Admin later connects the request to the correct SF1 learner record.
  if new.student_id is not null then
    select to_jsonb(student_record)
    into v_student_data
    from public.students student_record
    where student_record.id = new.student_id;
  end if;

  v_student_name := coalesce(
    nullif(trim(v_student_data ->> 'full_name'), ''),
    nullif(trim(v_student_data ->> 'name'), ''),
    nullif(
      trim(concat_ws(
        ' ',
        nullif(trim(v_student_data ->> 'first_name'), ''),
        nullif(trim(v_student_data ->> 'middle_name'), ''),
        nullif(trim(v_student_data ->> 'last_name'), '')
      )),
      ''
    ),
    nullif(trim(v_request_data ->> 'requested_student_name'), ''),
    'the requested learner'
  );

  select coalesce(
    nullif(trim(profile_record.full_name), ''),
    nullif(trim(profile_record.email), ''),
    'A Class Adviser'
  )
  into v_requester_name
  from public.profiles profile_record
  where profile_record.id = new.requester_id;

  v_requester_name := coalesce(v_requester_name, 'A Class Adviser');

  v_new_status := case lower(trim(coalesce(new.status, '')))
    when 'pending' then 'pending_review'
    when 'pending review' then 'pending_review'
    when 'completed' then 'approved'
    when 'released' then 'approved'
    else lower(trim(coalesce(new.status, 'pending_review')))
  end;

  if tg_op = 'UPDATE' then
    v_old_status := case lower(trim(coalesce(old.status, '')))
      when 'pending' then 'pending_review'
      when 'pending review' then 'pending_review'
      when 'completed' then 'approved'
      when 'released' then 'approved'
      else lower(trim(coalesce(old.status, 'pending_review')))
    end;
  end if;

  if tg_op = 'INSERT' then
    for v_admin in
      select distinct user_role.user_id
      from public.user_roles user_role
      where lower(user_role.role::text) = 'admin'
    loop
      perform public.create_notification(
        v_admin.user_id,
        'New SF1 Credential Request',
        concat(
          v_requester_name,
          ' requested the SF1 record for ',
          v_student_name,
          case
            when nullif(trim(new.reason), '') is not null
              then concat('. Reason: ', trim(new.reason))
            else '.'
          end
        ),
        'form_submission',
        '/admin?section=credential_requests',
        new.requester_id,
        concat('credential-request:', new.id, ':created'),
        jsonb_build_object(
          'credential_request_id', new.id,
          'student_id', new.student_id,
          'requested_student_name', v_request_data ->> 'requested_student_name',
          'requester_id', new.requester_id,
          'requesting_class_id', new.requesting_class_id,
          'previous_class_id', new.previous_class_id,
          'status', v_new_status
        )
      );
    end loop;

  elsif tg_op = 'UPDATE'
    and v_old_status is distinct from v_new_status
    and v_new_status in ('approved', 'rejected')
  then
    v_notification_title := case v_new_status
      when 'approved' then 'SF1 Request Approved'
      when 'rejected' then 'SF1 Request Rejected'
    end;

    v_notification_message := case v_new_status
      when 'approved' then concat(
        'The SF1 request for ',
        v_student_name,
        ' was approved. Select a class to add the learner and complete the enrollment.'
      )
      when 'rejected' then concat(
        'The SF1 request for ',
        v_student_name,
        ' was rejected by Admin.'
      )
    end;

    v_actor_id := coalesce(
      nullif(v_request_data ->> 'reviewed_by', '')::uuid,
      new.processed_by
    );

    perform public.create_notification(
      new.requester_id,
      v_notification_title,
      v_notification_message,
      'form_review',
      '/credential-requests',
      v_actor_id,
      concat('credential-request:', new.id, ':', v_new_status),
      jsonb_build_object(
        'credential_request_id', new.id,
        'student_id', new.student_id,
        'requested_student_name', v_request_data ->> 'requested_student_name',
        'requester_id', new.requester_id,
        'status', v_new_status,
        'admin_remarks', v_request_data ->> 'admin_remarks',
        'reviewed_at', v_request_data ->> 'reviewed_at',
        'released_at', v_request_data ->> 'released_at',
        'processed_by', new.processed_by,
        'processed_at', new.processed_at
      )
    );
  end if;

  return new;
end;
$$;

revoke all on function public.notify_student_credential_request_change()
  from public, anon, authenticated;

drop trigger if exists student_credential_request_notifications
  on public.student_credential_requests;

create trigger student_credential_request_notifications
after insert or update of status on public.student_credential_requests
for each row execute function public.notify_student_credential_request_change();

notify pgrst, 'reload schema';
