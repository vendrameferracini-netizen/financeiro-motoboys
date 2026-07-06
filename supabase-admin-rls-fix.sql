-- Financeiro Motoboys - correcao RLS para migracao do admin
-- Rode este arquivo no Supabase SQL Editor.
--
-- Causa comum do erro:
-- As tabelas motoboys/settings estavam protegidas por RLS e a policy de escrita
-- dependia de public.is_admin(). Se o perfil do usuario admin ainda nao estava
-- cadastrado em public.profiles, o Supabase bloqueava a migracao.

begin;

create or replace function public.current_profile_role()
returns text
language sql
security definer
set search_path = public
stable
as $$
  select coalesce(
    (select role from public.profiles where id = auth.uid() and active is true),
    nullif(auth.jwt() -> 'app_metadata' ->> 'role', ''),
    nullif(auth.jwt() -> 'user_metadata' ->> 'role', ''),
    case
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'admin@financeiro.local' then 'admin'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'gil@financeiro.local' then 'GIL'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'sales@financeiro.local' then 'SALES'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'guilherme@financeiro.local' then 'GUILHERME'
      else null
    end,
    'anon'
  );
$$;

create or replace function public.is_admin()
returns boolean
language sql
security definer
set search_path = public
stable
as $$
  select auth.role() = 'authenticated'
    and public.current_profile_role() = 'admin';
$$;

alter table public.motoboys enable row level security;
alter table public.settings enable row level security;

drop policy if exists motoboys_select_all on public.motoboys;
drop policy if exists motoboys_admin_all on public.motoboys;
drop policy if exists motoboys_admin_write on public.motoboys;
drop policy if exists motoboys_select_authenticated on public.motoboys;
drop policy if exists motoboys_insert_admin on public.motoboys;
drop policy if exists motoboys_update_admin on public.motoboys;
drop policy if exists motoboys_delete_admin on public.motoboys;

create policy motoboys_select_authenticated
on public.motoboys
for select
to authenticated
using (true);

create policy motoboys_insert_admin
on public.motoboys
for insert
to authenticated
with check (public.is_admin());

create policy motoboys_update_admin
on public.motoboys
for update
to authenticated
using (public.is_admin())
with check (public.is_admin());

create policy motoboys_delete_admin
on public.motoboys
for delete
to authenticated
using (public.is_admin());

drop policy if exists settings_select_all on public.settings;
drop policy if exists settings_admin_all on public.settings;
drop policy if exists settings_admin_write on public.settings;
drop policy if exists settings_select_authenticated on public.settings;
drop policy if exists settings_insert_authenticated on public.settings;
drop policy if exists settings_update_authenticated on public.settings;
drop policy if exists settings_delete_admin on public.settings;

create policy settings_select_authenticated
on public.settings
for select
to authenticated
using (true);

create policy settings_insert_admin
on public.settings
for insert
to authenticated
with check (public.is_admin());

create policy settings_update_admin
on public.settings
for update
to authenticated
using (public.is_admin())
with check (public.is_admin());

create policy settings_delete_admin
on public.settings
for delete
to authenticated
using (public.is_admin());

commit;
