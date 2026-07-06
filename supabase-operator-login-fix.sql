-- Financeiro Motoboys - correcao do login/perfil OPERADOR
-- Rode este arquivo no Supabase SQL Editor depois de criar o usuario no Auth.
-- Login esperado no app:
--   operador -> operador@financeiro.local
-- Senha esperada pelo app:
--   a senha cadastrada no Supabase Auth para operador@financeiro.local
--   se voce usou a senha temporaria sugerida: Operador@2026#Temp

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
set
  raw_app_meta_data = coalesce(raw_app_meta_data, '{}'::jsonb) || '{"role":"OPERADOR"}'::jsonb,
  raw_user_meta_data = coalesce(raw_user_meta_data, '{}'::jsonb) || '{"username":"operador","role":"OPERADOR","full_name":"Operador"}'::jsonb
where lower(email) = 'operador@financeiro.local';

do $$
begin
  if not exists (select 1 from auth.users where lower(email) = 'operador@financeiro.local') then
    raise exception 'Usuario operador@financeiro.local nao existe em auth.users. Crie o usuario no Supabase Auth e rode este SQL novamente.';
  end if;

  if not exists (
    select 1
    from public.profiles p
    join auth.users u on u.id = p.id
    where lower(u.email) = 'operador@financeiro.local'
      and p.username = 'operador'
      and p.role = 'OPERADOR'
      and p.active is true
  ) then
    raise exception 'Perfil OPERADOR nao foi sincronizado corretamente em public.profiles.';
  end if;
end;
$$;

commit;

select
  u.id as auth_user_id,
  u.email,
  u.email_confirmed_at is not null as email_confirmado,
  u.last_sign_in_at,
  p.username,
  p.role,
  p.active,
  p.updated_at
from auth.users u
left join public.profiles p on p.id = u.id
where lower(u.email) = 'operador@financeiro.local';
