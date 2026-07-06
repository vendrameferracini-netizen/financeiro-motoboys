-- Limpeza cirurgica da transportadora duplicada "Frellin".
--
-- Alvo incorreto: Frellin
-- Alvo correto:   Freelin
--
-- Seguranca:
-- - Nao recria banco.
-- - Nao roda seed.
-- - Nao altera "Freelin".
-- - Atua apenas em registros com nome exatamente igual a "Frellin"
--   ou vinculados ao id da transportadora "Frellin".
-- - Dashboards, rankings, relatorios, PDFs e Excel sao recalculados pelo app
--   a partir das tabelas/settings apos a remocao.
--
-- Execute em Supabase SQL Editor como usuario/admin com permissao de escrita.
-- Leia o resultado final "RESUMO FINAL" antes de confirmar a limpeza.

begin;

create temp table _target_carrier as
select id, name, data
from public.motoboys
where name = 'Frellin';

create temp table _correct_carrier_before as
select id, name, data
from public.motoboys
where name = 'Freelin';

create temp table _freelin_counts_before as
select 'motoboys' as table_name, count(*)::bigint as total from public.motoboys where name = 'Freelin'
union all
select 'daily_launches', count(*) from public.daily_launches
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'discounts', count(*) from public.discounts
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'payments', count(*) from public.payments
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'receipts', count(*) from public.receipts
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin';

create temp table _candidate_daily as
select id
from public.daily_launches
where motoboy_name = 'Frellin'
   or motoboy_id in (select id from _target_carrier)
   or data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin';

create temp table _candidate_discounts as
select id
from public.discounts
where motoboy_name = 'Frellin'
   or motoboy_id in (select id from _target_carrier)
   or data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin';

create temp table _candidate_payments as
select id
from public.payments
where motoboy_name = 'Frellin'
   or motoboy_id in (select id from _target_carrier)
   or data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin';

create temp table _candidate_receipts as
select id
from public.receipts
where payment_id in (select id from _candidate_payments)
   or motoboy_name = 'Frellin'
   or motoboy_id in (select id from _target_carrier)
   or data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin';

-- Tabelas operacionais sem FK direta para motoboys; so entram se o JSON
-- possuir campos de transportadora/motoboy exatamente iguais a "Frellin".
create temp table _candidate_package_entries as
select id
from public.package_entries
where data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin'
   or data->>'name' = 'Frellin';

create temp table _candidate_expenses as
select id
from public.expenses
where data->>'rider' = 'Frellin'
   or data->>'motoboyName' = 'Frellin'
   or data->>'closingRider' = 'Frellin'
   or data->>'name' = 'Frellin';

create temp table _precheck as
select 'motoboys' as table_name, count(*)::bigint as records_to_delete from _target_carrier
union all select 'daily_launches', count(*) from _candidate_daily
union all select 'discounts', count(*) from _candidate_discounts
union all select 'payments', count(*) from _candidate_payments
union all select 'receipts', count(*) from _candidate_receipts
union all select 'package_entries', count(*) from _candidate_package_entries
union all select 'expenses', count(*) from _candidate_expenses;

select
  'CONFERENCIA ANTES DA EXCLUSAO' as etapa,
  table_name,
  records_to_delete
from _precheck
order by table_name;

select
  'TRANSPORTADORAS ENCONTRADAS' as etapa,
  'Frellin' as nome,
  count(*)::bigint as quantidade,
  coalesce(jsonb_agg(jsonb_build_object('id', id, 'name', name)), '[]'::jsonb) as registros
from _target_carrier
union all
select
  'TRANSPORTADORAS ENCONTRADAS',
  'Freelin',
  count(*)::bigint,
  coalesce(jsonb_agg(jsonb_build_object('id', id, 'name', name)), '[]'::jsonb)
from _correct_carrier_before;

do $$
begin
  if (select count(*) from _target_carrier) = 0 then
    raise exception 'Abortado: nao existe transportadora com nome exatamente igual a Frellin.';
  end if;

  if exists (
    select 1
    from public.motoboys
    where name = 'Freelin'
      and id in (select id from _target_carrier)
  ) then
    raise exception 'Abortado: ids de Frellin e Freelin se sobrepoem, revise manualmente.';
  end if;
end $$;

create temp table _deleted_receipts as
with deleted as (
  delete from public.receipts
  where id in (select id from _candidate_receipts)
  returning id
)
select id from deleted;

create temp table _deleted_payments as
with deleted as (
  delete from public.payments
  where id in (select id from _candidate_payments)
  returning id
)
select id from deleted;

create temp table _deleted_discounts as
with deleted as (
  delete from public.discounts
  where id in (select id from _candidate_discounts)
  returning id
)
select id from deleted;

create temp table _deleted_daily as
with deleted as (
  delete from public.daily_launches
  where id in (select id from _candidate_daily)
  returning id
)
select id from deleted;

create temp table _deleted_package_entries as
with deleted as (
  delete from public.package_entries
  where id in (select id from _candidate_package_entries)
  returning id
)
select id from deleted;

create temp table _deleted_expenses as
with deleted as (
  delete from public.expenses
  where id in (select id from _candidate_expenses)
  returning id
)
select id from deleted;

create temp table _deleted_motoboys as
with deleted as (
  delete from public.motoboys
  where id in (select id from _target_carrier)
    and name = 'Frellin'
  returning id
)
select id from deleted;

create temp table _settings_before as
select
  key,
  value,
  coalesce(jsonb_array_length(value->'riders'), 0) as riders_before,
  coalesce(jsonb_array_length(value->'daily'), 0) as daily_before,
  coalesce(jsonb_array_length(value->'discounts'), 0) as discounts_before,
  coalesce(jsonb_array_length(value->'payments'), 0) as payments_before,
  coalesce(jsonb_array_length(value->'receipts'), 0) as receipts_before,
  coalesce((select count(*) from jsonb_each(coalesce(value->'paid', '{}'::jsonb))), 0) as paid_before
from public.settings
where key = 'app_state';

update public.settings s
set value =
  jsonb_set(
    jsonb_set(
      jsonb_set(
        jsonb_set(
          jsonb_set(
            jsonb_set(
              s.value,
              '{riders}',
              coalesce((
                select jsonb_agg(item)
                from jsonb_array_elements(coalesce(s.value->'riders', '[]'::jsonb)) as item
                where coalesce(item->>'name', '') <> 'Frellin'
                  and coalesce(item->>'rider', '') <> 'Frellin'
                  and coalesce(item->>'motoboyName', '') <> 'Frellin'
                  and coalesce(item->>'closingRider', '') <> 'Frellin'
              ), '[]'::jsonb)
            ),
            '{daily}',
            coalesce((
              select jsonb_agg(item)
              from jsonb_array_elements(coalesce(s.value->'daily', '[]'::jsonb)) as item
              where coalesce(item->>'name', '') <> 'Frellin'
                and coalesce(item->>'rider', '') <> 'Frellin'
                and coalesce(item->>'motoboyName', '') <> 'Frellin'
                and coalesce(item->>'closingRider', '') <> 'Frellin'
            ), '[]'::jsonb)
          ),
          '{discounts}',
          coalesce((
            select jsonb_agg(item)
            from jsonb_array_elements(coalesce(s.value->'discounts', '[]'::jsonb)) as item
            where coalesce(item->>'name', '') <> 'Frellin'
              and coalesce(item->>'rider', '') <> 'Frellin'
              and coalesce(item->>'motoboyName', '') <> 'Frellin'
              and coalesce(item->>'closingRider', '') <> 'Frellin'
          ), '[]'::jsonb)
        ),
        '{payments}',
        coalesce((
          select jsonb_agg(item)
          from jsonb_array_elements(coalesce(s.value->'payments', '[]'::jsonb)) as item
          where coalesce(item->>'name', '') <> 'Frellin'
            and coalesce(item->>'rider', '') <> 'Frellin'
            and coalesce(item->>'motoboyName', '') <> 'Frellin'
            and coalesce(item->>'closingRider', '') <> 'Frellin'
        ), '[]'::jsonb)
      ),
      '{receipts}',
      coalesce((
        select jsonb_agg(item)
        from jsonb_array_elements(coalesce(s.value->'receipts', '[]'::jsonb)) as item
        where coalesce(item->>'name', '') <> 'Frellin'
          and coalesce(item->>'rider', '') <> 'Frellin'
          and coalesce(item->>'motoboyName', '') <> 'Frellin'
          and coalesce(item->>'closingRider', '') <> 'Frellin'
      ), '[]'::jsonb)
    ),
    '{paid}',
    coalesce((
      select jsonb_object_agg(key, value)
      from jsonb_each(coalesce(s.value->'paid', '{}'::jsonb))
      where lower(key) not like 'frellin|%'
    ), '{}'::jsonb)
  ),
  updated_at = now()
where s.key = 'app_state';

create temp table _settings_after as
select
  key,
  value,
  coalesce(jsonb_array_length(value->'riders'), 0) as riders_after,
  coalesce(jsonb_array_length(value->'daily'), 0) as daily_after,
  coalesce(jsonb_array_length(value->'discounts'), 0) as discounts_after,
  coalesce(jsonb_array_length(value->'payments'), 0) as payments_after,
  coalesce(jsonb_array_length(value->'receipts'), 0) as receipts_after,
  coalesce((select count(*) from jsonb_each(coalesce(value->'paid', '{}'::jsonb))), 0) as paid_after
from public.settings
where key = 'app_state';

create temp table _freelin_counts_after as
select 'motoboys' as table_name, count(*)::bigint as total from public.motoboys where name = 'Freelin'
union all
select 'daily_launches', count(*) from public.daily_launches
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'discounts', count(*) from public.discounts
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'payments', count(*) from public.payments
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin'
union all
select 'receipts', count(*) from public.receipts
where motoboy_name = 'Freelin'
   or motoboy_id in (select id from _correct_carrier_before)
   or data->>'rider' = 'Freelin'
   or data->>'motoboyName' = 'Freelin'
   or data->>'closingRider' = 'Freelin';

do $$
begin
  if exists (select 1 from public.motoboys where name = 'Frellin') then
    raise exception 'Abortado: ainda existe transportadora com nome exatamente igual a Frellin.';
  end if;

  if exists (
    select 1
    from _freelin_counts_before b
    join _freelin_counts_after a using (table_name)
    where b.total <> a.total
  ) then
    raise exception 'Abortado: contagens vinculadas a Freelin mudaram.';
  end if;
end $$;

select 'RESUMO FINAL' as etapa, 'motoboys' as table_name, count(*)::bigint as registros_excluidos from _deleted_motoboys
union all select 'RESUMO FINAL', 'daily_launches', count(*) from _deleted_daily
union all select 'RESUMO FINAL', 'discounts', count(*) from _deleted_discounts
union all select 'RESUMO FINAL', 'payments', count(*) from _deleted_payments
union all select 'RESUMO FINAL', 'receipts', count(*) from _deleted_receipts
union all select 'RESUMO FINAL', 'package_entries', count(*) from _deleted_package_entries
union all select 'RESUMO FINAL', 'expenses', count(*) from _deleted_expenses
union all
select
  'RESUMO FINAL',
  'settings.app_state.riders',
  coalesce(sum(b.riders_before - a.riders_after), 0)::bigint
from _settings_before b
join _settings_after a using (key)
union all
select 'RESUMO FINAL', 'settings.app_state.daily', coalesce(sum(b.daily_before - a.daily_after), 0)::bigint
from _settings_before b join _settings_after a using (key)
union all
select 'RESUMO FINAL', 'settings.app_state.discounts', coalesce(sum(b.discounts_before - a.discounts_after), 0)::bigint
from _settings_before b join _settings_after a using (key)
union all
select 'RESUMO FINAL', 'settings.app_state.payments', coalesce(sum(b.payments_before - a.payments_after), 0)::bigint
from _settings_before b join _settings_after a using (key)
union all
select 'RESUMO FINAL', 'settings.app_state.receipts', coalesce(sum(b.receipts_before - a.receipts_after), 0)::bigint
from _settings_before b join _settings_after a using (key)
union all
select 'RESUMO FINAL', 'settings.app_state.paid', coalesce(sum(b.paid_before - a.paid_after), 0)::bigint
from _settings_before b join _settings_after a using (key)
order by table_name;

select
  'CONFIRMACAO' as etapa,
  (select count(*) from public.motoboys where name = 'Frellin') as frellin_restante,
  (select count(*) from public.motoboys where name = 'Freelin') as freelin_restante,
  'Apenas Frellin foi removida; Freelin manteve as mesmas contagens pre-limpeza.' as observacao;

commit;
