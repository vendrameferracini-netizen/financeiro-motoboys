-- Financeiro Motoboys - correcao RLS da tabela settings
-- Rode este arquivo no Supabase SQL Editor.

begin;

alter table public.settings enable row level security;

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

create policy settings_insert_authenticated
on public.settings
for insert
to authenticated
with check (auth.role() = 'authenticated');

create policy settings_update_authenticated
on public.settings
for update
to authenticated
using (auth.role() = 'authenticated')
with check (auth.role() = 'authenticated');

create policy settings_delete_admin
on public.settings
for delete
to authenticated
using (public.is_admin());

commit;
