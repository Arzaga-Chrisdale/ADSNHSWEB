-- Add editable profile image support for Class Adviser and Subject Teacher profiles.

alter table public.profiles
add column if not exists avatar_url text;

insert into storage.buckets (id, name, public)
values ('profile-images', 'profile-images', true)
on conflict (id) do update set public = true;

drop policy if exists "Profile images are publicly viewable" on storage.objects;
drop policy if exists "Users can upload own profile images" on storage.objects;
drop policy if exists "Users can update own profile images" on storage.objects;
drop policy if exists "Users can delete own profile images" on storage.objects;

create policy "Profile images are publicly viewable"
on storage.objects
for select
to public
using (bucket_id = 'profile-images');

create policy "Users can upload own profile images"
on storage.objects
for insert
to authenticated
with check (
  bucket_id = 'profile-images'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "Users can update own profile images"
on storage.objects
for update
to authenticated
using (
  bucket_id = 'profile-images'
  and auth.uid()::text = (storage.foldername(name))[1]
)
with check (
  bucket_id = 'profile-images'
  and auth.uid()::text = (storage.foldername(name))[1]
);

create policy "Users can delete own profile images"
on storage.objects
for delete
to authenticated
using (
  bucket_id = 'profile-images'
  and auth.uid()::text = (storage.foldername(name))[1]
);

notify pgrst, 'reload schema';
