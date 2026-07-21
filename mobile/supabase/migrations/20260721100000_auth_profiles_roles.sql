-- Auth profiles with roles for BÜW-Toolbox (web + mobile)
create extension if not exists "pgcrypto";

do $$
begin
  if not exists (select 1 from pg_type where typname = 'user_role') then
    create type public.user_role as enum ('admin', 'benutzer');
  end if;
end $$;

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  email text,
  first_name text,
  last_name text,
  display_name text,
  role public.user_role not null default 'benutzer',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.profiles
  add column if not exists first_name text,
  add column if not exists last_name text,
  add column if not exists display_name text;

alter table public.profiles
  add column if not exists role public.user_role;

update public.profiles
set role = 'benutzer'
where role is null;

alter table public.profiles
  alter column role set default 'benutzer';

alter table public.profiles
  alter column role set not null;

alter table public.profiles enable row level security;

drop policy if exists "Users can read own profile" on public.profiles;
create policy "Users can read own profile"
  on public.profiles
  for select
  using (auth.uid() = id);

drop policy if exists "Users can update own profile" on public.profiles;
create policy "Users can update own profile"
  on public.profiles
  for update
  using (auth.uid() = id)
  with check (auth.uid() = id and role = (select p.role from public.profiles p where p.id = auth.uid()));

drop policy if exists "Users can insert own profile" on public.profiles;
create policy "Users can insert own profile"
  on public.profiles
  for insert
  with check (auth.uid() = id and role = 'benutzer');

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  first_name_value text := coalesce(new.raw_user_meta_data ->> 'first_name', '');
  last_name_value text := coalesce(new.raw_user_meta_data ->> 'last_name', '');
  display_name_value text := coalesce(
    new.raw_user_meta_data ->> 'display_name',
    nullif(trim(both from first_name_value || ' ' || last_name_value), ''),
    split_part(new.email, '@', 1)
  );
begin
  insert into public.profiles (id, email, first_name, last_name, display_name, role)
  values (
    new.id,
    new.email,
    nullif(first_name_value, ''),
    nullif(last_name_value, ''),
    display_name_value,
    'benutzer'
  )
  on conflict (id) do update
    set email = excluded.email,
        first_name = coalesce(excluded.first_name, public.profiles.first_name),
        last_name = coalesce(excluded.last_name, public.profiles.last_name),
        display_name = coalesce(excluded.display_name, public.profiles.display_name),
        updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();
