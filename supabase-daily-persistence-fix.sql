-- Financeiro Motoboys - persistencia dos lancamentos diarios
-- Rode no Supabase SQL Editor para liberar ADMIN e OPERADOR na tabela daily_launches.
-- Esta versao tambem corrige o caso em que operador@financeiro.local volta
-- para role BASE ao fazer login por causa de trigger/function antiga.

begin;

create or replace function public.current_profile_role()
returns text
language sql
security definer
set search_path = public
stable
as $$
  select coalesce(
    case
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'admin@financeiro.local' then 'admin'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'gil@financeiro.local' then 'GIL'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'sales@financeiro.local' then 'SALES'
      when lower(coalesce(auth.jwt() ->> 'email', '')) in ('guilherme@financeiro.local', 'gm@financeiro.local') then 'GUILHERME'
      when lower(coalesce(auth.jwt() ->> 'email', '')) = 'operador@financeiro.local' then 'OPERADOR'
      else null
    end,
    (select role from public.profiles where id = auth.uid() and active is true),
    nullif(auth.jwt() -> 'app_metadata' ->> 'role', ''),
    nullif(auth.jwt() -> 'user_metadata' ->> 'role', ''),
    'anon'
  );
$$;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  user_name text := lower(split_part(new.email, '@', 1));
  user_role text;
begin
  user_role := case
    when user_name = 'admin' then 'admin'
    when user_name = 'gil' then 'GIL'
    when user_name = 'sales' then 'SALES'
    when user_name in ('guilherme', 'gm') then 'GUILHERME'
    when user_name = 'operador' then 'OPERADOR'
    else coalesce(nullif(new.raw_user_meta_data->>'role', ''), 'BASE')
  end;

  insert into public.profiles(id, username, role, full_name, active)
  values (new.id, user_name, user_role, coalesce(new.raw_user_meta_data->>'full_name', user_name), true)
  on conflict (id) do update
    set username = excluded.username,
        role = excluded.role,
        full_name = excluded.full_name,
        active = true,
        updated_at = now();

  return new;
end;
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

update auth.users
set raw_user_meta_data = coalesce(raw_user_meta_data, '{}'::jsonb)
  || '{"username":"operador","role":"OPERADOR","full_name":"Operador"}'::jsonb
where lower(email) = 'operador@financeiro.local';

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
