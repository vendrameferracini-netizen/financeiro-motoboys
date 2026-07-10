-- Financeiro Motoboys - persistencia da Entrada de Pacotes da Base
-- Tabela correta usada pelo app: public.package_entries
-- Rode no Supabase SQL Editor.

begin;

create or replace function public.normalize_responsible(input text)
returns text
language plpgsql
immutable
as $$
declare
  v text := lower(trim(coalesce(input, 'BASE')));
begin
  v := replace(replace(replace(v, '.', ''), '-', ''), '_', '');

  if v in ('admin', 'administrador') then
    return 'admin';
  elsif v = 'gil' then
    return 'GIL';
  elsif v in ('sales', 'salles') then
    return 'SALES';
  elsif v in ('gm', 'g m', 'guilherme', 'guilhermem', 'guilhermemendes') then
    return 'GUILHERME';
  elsif v in ('operador', 'operator') then
    return 'OPERADOR';
  elsif v in ('base', 'empresa', 'operacao') then
    return 'BASE';
  end if;

  return upper(coalesce(input, 'BASE'));
end;
$$;

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
      when lower(coalesce(auth.jwt() ->> 'email', '')) in ('guilherme@financeiro.local', 'gm@financeiro.local') then 'GUILHERME'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'operador@financeiro.local' then 'OPERADOR'
      else null
    end,
    'anon'
  );
$$;

create or replace function public.can_write_responsible(target text)
returns boolean
language sql
security definer
set search_path = public
stable
as $$
  select case
    when public.current_profile_role() = 'admin' then true
    when public.current_profile_role() = 'OPERADOR' then false
    when public.normalize_responsible(target) = 'BASE' then false
    else public.normalize_responsible(target) = public.current_profile_role()
  end;
$$;

drop policy if exists package_select_all on public.package_entries;
drop policy if exists package_insert_owner on public.package_entries;
drop policy if exists package_update_owner on public.package_entries;
drop policy if exists package_delete_owner on public.package_entries;
drop policy if exists entries_select_all on public.package_entries;
drop policy if exists entries_insert_owner on public.package_entries;
drop policy if exists entries_update_owner on public.package_entries;
drop policy if exists entries_delete_owner on public.package_entries;

create policy package_select_all
on public.package_entries
for select
to authenticated
using (true);

create policy package_insert_owner
on public.package_entries
for insert
to authenticated
with check (public.can_write_responsible(owner));

create policy package_update_owner
on public.package_entries
for update
to authenticated
using (public.can_write_responsible(owner))
with check (public.can_write_responsible(owner));

create policy package_delete_owner
on public.package_entries
for delete
to authenticated
using (public.can_write_responsible(owner));

create index if not exists idx_package_entries_owner_entry_date
on public.package_entries(owner, entry_date);

commit;

select
  policyname,
  cmd,
  qual,
  with_check
from pg_policies
where schemaname = 'public'
  and tablename = 'package_entries'
order by policyname;
