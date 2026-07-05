-- Financeiro Motoboys - Supabase schema
-- Rode este arquivo no SQL Editor do Supabase.
-- Crie os usuarios no Supabase Auth:
-- admin@financeiro.local, gil@financeiro.local, sales@financeiro.local, guilherme@financeiro.local

create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  username text not null unique,
  role text not null check (role in ('admin', 'GIL', 'SALES', 'GUILHERME', 'BASE')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.profile_role()
returns text
language sql
security definer
set search_path = public
stable
as $$
  select coalesce((select role from public.profiles where id = auth.uid()), 'BASE')
$$;

create or replace function public.is_admin()
returns boolean
language sql
security definer
set search_path = public
stable
as $$
  select public.profile_role() = 'admin'
$$;

create or replace function public.can_write_owner(target_owner text)
returns boolean
language sql
security definer
set search_path = public
stable
as $$
  select public.is_admin() or coalesce(target_owner, 'BASE') = public.profile_role()
$$;

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  uname text := lower(split_part(new.email, '@', 1));
  urole text := 'BASE';
begin
  if uname = 'admin' then urole := 'admin';
  elsif uname = 'gil' then urole := 'GIL';
  elsif uname = 'sales' then urole := 'SALES';
  elsif uname in ('guilherme', 'gm') then urole := 'GUILHERME';
  end if;

  insert into public.profiles(id, username, role)
  values (new.id, uname, urole)
  on conflict (id) do update set username = excluded.username, role = excluded.role, updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created_financeiro on auth.users;
create trigger on_auth_user_created_financeiro
after insert on auth.users
for each row execute function public.handle_new_user();

create table if not exists public.motoboys (
  id text primary key,
  owner text not null default 'BASE',
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.daily_launches (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.package_entries (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.discounts (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.expenses (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.payments (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.receipts (
  id text primary key,
  owner text not null,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.settings (
  key text primary key,
  value jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.backups (
  id uuid primary key default gen_random_uuid(),
  data jsonb not null,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now()
);

create table if not exists public.change_logs (
  id uuid primary key default gen_random_uuid(),
  action text not null,
  detail text,
  metadata jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id) default auth.uid(),
  created_at timestamptz not null default now()
);

create index if not exists idx_motoboys_owner on public.motoboys(owner);
create index if not exists idx_daily_launches_owner on public.daily_launches(owner);
create index if not exists idx_package_entries_owner on public.package_entries(owner);
create index if not exists idx_discounts_owner on public.discounts(owner);
create index if not exists idx_expenses_owner on public.expenses(owner);
create index if not exists idx_payments_owner on public.payments(owner);
create index if not exists idx_receipts_owner on public.receipts(owner);
create index if not exists idx_change_logs_created_at on public.change_logs(created_at desc);

drop trigger if exists trg_motoboys_touch on public.motoboys;
create trigger trg_motoboys_touch before update on public.motoboys for each row execute function public.touch_updated_at();
drop trigger if exists trg_daily_launches_touch on public.daily_launches;
create trigger trg_daily_launches_touch before update on public.daily_launches for each row execute function public.touch_updated_at();
drop trigger if exists trg_package_entries_touch on public.package_entries;
create trigger trg_package_entries_touch before update on public.package_entries for each row execute function public.touch_updated_at();
drop trigger if exists trg_discounts_touch on public.discounts;
create trigger trg_discounts_touch before update on public.discounts for each row execute function public.touch_updated_at();
drop trigger if exists trg_expenses_touch on public.expenses;
create trigger trg_expenses_touch before update on public.expenses for each row execute function public.touch_updated_at();
drop trigger if exists trg_payments_touch on public.payments;
create trigger trg_payments_touch before update on public.payments for each row execute function public.touch_updated_at();
drop trigger if exists trg_receipts_touch on public.receipts;
create trigger trg_receipts_touch before update on public.receipts for each row execute function public.touch_updated_at();
drop trigger if exists trg_settings_touch on public.settings;
create trigger trg_settings_touch before update on public.settings for each row execute function public.touch_updated_at();

alter table public.profiles enable row level security;
alter table public.motoboys enable row level security;
alter table public.daily_launches enable row level security;
alter table public.package_entries enable row level security;
alter table public.discounts enable row level security;
alter table public.expenses enable row level security;
alter table public.payments enable row level security;
alter table public.receipts enable row level security;
alter table public.settings enable row level security;
alter table public.backups enable row level security;
alter table public.change_logs enable row level security;

create policy "profiles_select_authenticated" on public.profiles for select to authenticated using (true);
create policy "profiles_admin_update" on public.profiles for update to authenticated using (public.is_admin()) with check (public.is_admin());

create policy "motoboys_select_all" on public.motoboys for select to authenticated using (true);
create policy "motoboys_admin_write" on public.motoboys for all to authenticated using (public.is_admin()) with check (public.is_admin());

create policy "daily_select_all" on public.daily_launches for select to authenticated using (true);
create policy "daily_insert_owner" on public.daily_launches for insert to authenticated with check (public.can_write_owner(owner));
create policy "daily_update_owner" on public.daily_launches for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "daily_delete_owner" on public.daily_launches for delete to authenticated using (public.can_write_owner(owner));

create policy "entries_select_all" on public.package_entries for select to authenticated using (true);
create policy "entries_insert_owner" on public.package_entries for insert to authenticated with check (public.can_write_owner(owner));
create policy "entries_update_owner" on public.package_entries for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "entries_delete_owner" on public.package_entries for delete to authenticated using (public.can_write_owner(owner));

create policy "discounts_select_all" on public.discounts for select to authenticated using (true);
create policy "discounts_insert_owner" on public.discounts for insert to authenticated with check (public.can_write_owner(owner));
create policy "discounts_update_owner" on public.discounts for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "discounts_delete_owner" on public.discounts for delete to authenticated using (public.can_write_owner(owner));

create policy "expenses_select_all" on public.expenses for select to authenticated using (true);
create policy "expenses_insert_owner" on public.expenses for insert to authenticated with check (public.can_write_owner(owner));
create policy "expenses_update_owner" on public.expenses for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "expenses_delete_owner" on public.expenses for delete to authenticated using (public.can_write_owner(owner));

create policy "payments_select_all" on public.payments for select to authenticated using (true);
create policy "payments_insert_owner" on public.payments for insert to authenticated with check (public.can_write_owner(owner));
create policy "payments_update_owner" on public.payments for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "payments_delete_owner" on public.payments for delete to authenticated using (public.can_write_owner(owner));

create policy "receipts_select_all" on public.receipts for select to authenticated using (true);
create policy "receipts_insert_owner" on public.receipts for insert to authenticated with check (public.can_write_owner(owner));
create policy "receipts_update_owner" on public.receipts for update to authenticated using (public.can_write_owner(owner)) with check (public.can_write_owner(owner));
create policy "receipts_delete_owner" on public.receipts for delete to authenticated using (public.can_write_owner(owner));

create policy "settings_select_all" on public.settings for select to authenticated using (true);
create policy "settings_admin_write" on public.settings for all to authenticated using (public.is_admin()) with check (public.is_admin());

create policy "backups_admin_all" on public.backups for all to authenticated using (public.is_admin()) with check (public.is_admin());
create policy "change_logs_select_all" on public.change_logs for select to authenticated using (true);
create policy "change_logs_insert_authenticated" on public.change_logs for insert to authenticated with check (true);
