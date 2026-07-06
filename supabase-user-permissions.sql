-- Financeiro Motoboys - permissoes por perfil e usuario OPERADOR
-- Rode no Supabase SQL Editor depois de criar operador@financeiro.local no Auth.

begin;

alter table public.profiles drop constraint if exists profiles_role_check;
alter table public.profiles
  add constraint profiles_role_check
  check (role in ('admin', 'GIL', 'SALES', 'GUILHERME', 'OPERADOR', 'BASE'));

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
  elsif v in ('gil') then
    return 'GIL';
  elsif v in ('sales', 'salles') then
    return 'SALES';
  elsif v in ('gm', 'g m', 'guilherme', 'guilhermem', 'guilhermemendes', 'guilhermem.') then
    return 'GUILHERME';
  elsif v in ('operador', 'operator') then
    return 'OPERADOR';
  elsif v in ('base', 'empresa', 'operacao', 'operacao') then
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
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'guilherme@financeiro.local' then 'GUILHERME'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'operador@financeiro.local' then 'OPERADOR'
      else null
    end,
    'anon'
  );
$$;

drop policy if exists daily_insert_operator on public.daily_launches;
drop policy if exists daily_update_operator_today on public.daily_launches;
drop policy if exists daily_delete_operator_today on public.daily_launches;

create policy daily_insert_operator
on public.daily_launches
for insert
to authenticated
with check (public.current_profile_role() = 'OPERADOR' and launch_date = current_date);

create policy daily_update_operator_today
on public.daily_launches
for update
to authenticated
using (public.current_profile_role() = 'OPERADOR' and launch_date = current_date)
with check (public.current_profile_role() = 'OPERADOR' and launch_date = current_date);

create policy daily_delete_operator_today
on public.daily_launches
for delete
to authenticated
using (public.current_profile_role() = 'OPERADOR' and launch_date = current_date);

insert into public.profiles(id, username, role, full_name, active)
select id, 'operador', 'OPERADOR', 'Operador', true
from auth.users
where lower(email) = 'operador@financeiro.local'
on conflict (id) do update
  set username = excluded.username,
      role = excluded.role,
      full_name = excluded.full_name,
      active = true,
      updated_at = now();

commit;
