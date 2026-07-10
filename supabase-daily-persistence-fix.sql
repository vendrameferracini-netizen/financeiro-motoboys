-- Financeiro Motoboys - persistencia dos lancamentos diarios
-- Rode no Supabase SQL Editor para liberar ADMIN e OPERADOR na tabela daily_launches.

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
      when lower(coalesce(auth.jwt() ->> 'email', '')) in ('guilherme@financeiro.local', 'gm@financeiro.local') then 'GUILHERME'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'operador@financeiro.local' then 'OPERADOR'
      else null
    end,
    'anon'
  );
$$;

insert into public.profiles(id, username, role, full_name, active)
select id, 'operador', 'OPERADOR', coalesce(raw_user_meta_data->>'full_name', 'Operador'), true
from auth.users
where lower(email) = 'operador@financeiro.local'
on conflict (id) do update
  set username = 'operador',
      role = 'OPERADOR',
      full_name = coalesce(excluded.full_name, public.profiles.full_name, 'Operador'),
      active = true,
      updated_at = now();

drop policy if exists daily_insert_operator on public.daily_launches;
drop policy if exists daily_update_operator_today on public.daily_launches;
drop policy if exists daily_delete_operator_today on public.daily_launches;

create policy daily_insert_operator
on public.daily_launches
for insert
to authenticated
with check (public.current_profile_role() = 'OPERADOR');

create policy daily_update_operator_today
on public.daily_launches
for update
to authenticated
using (public.current_profile_role() = 'OPERADOR')
with check (public.current_profile_role() = 'OPERADOR');

create policy daily_delete_operator_today
on public.daily_launches
for delete
to authenticated
using (public.current_profile_role() = 'OPERADOR');

commit;

select
  u.email,
  p.username,
  p.role,
  p.active
from auth.users u
left join public.profiles p on p.id = u.id
where lower(u.email) in ('admin@financeiro.local', 'operador@financeiro.local');
