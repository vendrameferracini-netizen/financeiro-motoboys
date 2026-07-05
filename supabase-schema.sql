-- Financeiro Motoboys - Supabase schema completo
-- Execute em ordem: Parte 1, Parte 2, Parte 3 e Parte 4.
-- Primeiro crie os usuarios no Supabase Auth, se ainda nao existirem:
-- admin@financeiro.local, gil@financeiro.local, sales@financeiro.local,
-- guilherme@financeiro.local.

-- =========================================================
-- PARTE 1/4 - Extensao e tabelas
-- =========================================================

create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  username text unique not null,
  role text not null default 'BASE'
    check (role in ('admin', 'GIL', 'SALES', 'GUILHERME', 'BASE')),
  full_name text,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.motoboys (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  name text not null,
  region text,
  work_type text not null default 'com_coleta'
    check (work_type in ('com_coleta', 'sem_coleta', 'freelancer')),
  has_collection boolean not null default true,
  rate_ml numeric(12,2) not null default 8,
  rate_shopee numeric(12,2) not null default 5,
  rate_avulso numeric(12,2) not null default 8,
  status text not null default 'ativo' check (status in ('ativo', 'inativo')),
  notes text,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.daily_launches (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  launch_date date not null default current_date,
  motoboy_id text references public.motoboys(id) on delete set null,
  motoboy_name text,
  launch_type text not null default 'com_coleta'
    check (launch_type in ('com_coleta', 'sem_coleta')),
  ml_qty integer not null default 0 check (ml_qty >= 0),
  shopee_qty integer not null default 0 check (shopee_qty >= 0),
  avulso_qty integer not null default 0 check (avulso_qty >= 0),
  rate_ml numeric(12,2) not null default 8,
  rate_shopee numeric(12,2) not null default 5,
  rate_avulso numeric(12,2) not null default 8,
  gross_total numeric(12,2) not null default 0,
  advances_total numeric(12,2) not null default 0,
  discounts_total numeric(12,2) not null default 0,
  bonuses_total numeric(12,2) not null default 0,
  net_total numeric(12,2) not null default 0,
  responsible_name text,
  status text not null default 'pendente' check (status in ('pendente', 'conferido', 'pago')),
  note text,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.package_entries (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  entry_date date not null default current_date,
  responsible text not null default 'BASE'
    check (responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  ml_qty integer not null default 0 check (ml_qty >= 0),
  shopee_qty integer not null default 0 check (shopee_qty >= 0),
  total_packages integer not null default 0 check (total_packages >= 0),
  rate_ml numeric(12,2) not null default 8,
  rate_shopee numeric(12,2) not null default 5,
  value_ml numeric(12,2) not null default 0,
  value_shopee numeric(12,2) not null default 0,
  total_value numeric(12,2) not null default 0,
  status text not null default 'pendente' check (status in ('pendente', 'conferido', 'pago')),
  note text,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.discounts (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  discount_date date default current_date,
  responsible text not null default 'BASE'
    check (responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  motoboy_id text references public.motoboys(id) on delete set null,
  motoboy_name text,
  discount_type text not null default 'OUTROS'
    check (discount_type in ('VALE', 'EXTRAVIO', 'OCORRENCIA', 'EXTRAVIO/OCORRENCIA', 'OUTROS')),
  value numeric(12,2) not null default 0 check (value >= 0),
  reason text,
  occurrence text,
  package_code text,
  observation text,
  original_sheet text,
  original_line integer,
  original_column text,
  unique_import_key text unique,
  status text not null default 'pendente' check (status in ('pendente', 'conferido', 'pago')),
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.expenses (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  expense_date date not null default current_date,
  responsible text not null default 'BASE'
    check (responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  expense_type text not null default 'variavel' check (expense_type in ('fixa', 'variavel')),
  category text not null,
  description text,
  value numeric(12,2) not null default 0 check (value >= 0),
  observation text,
  origin_period_start date,
  origin_period_end date,
  origin_period_label text,
  discount_period_start date,
  discount_period_end date,
  discount_period_label text,
  status text not null default 'pendente'
    check (status in ('pendente', 'lancado_no_fechamento', 'pago')),
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint expenses_variable_responsible_required
    check (expense_type <> 'variavel' or responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE'))
);

create table if not exists public.payments (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  payment_date date not null default current_date,
  responsible text not null default 'BASE'
    check (responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  motoboy_id text references public.motoboys(id) on delete set null,
  motoboy_name text,
  period_start date,
  period_end date,
  period_label text,
  gross_total numeric(12,2) not null default 0,
  discounts_total numeric(12,2) not null default 0,
  advances_total numeric(12,2) not null default 0,
  bonuses_total numeric(12,2) not null default 0,
  net_paid numeric(12,2) not null default 0,
  status text not null default 'pendente' check (status in ('pendente', 'conferido', 'pago')),
  note text,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.receipts (
  id text primary key default gen_random_uuid()::text,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  receipt_number text unique,
  payment_id text references public.payments(id) on delete set null,
  responsible text not null default 'BASE'
    check (responsible in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  motoboy_id text references public.motoboys(id) on delete set null,
  motoboy_name text,
  period_start date,
  period_end date,
  period_label text,
  payment_date date,
  ml_qty integer not null default 0,
  shopee_qty integer not null default 0,
  avulso_qty integer not null default 0,
  gross_total numeric(12,2) not null default 0,
  discounts_total numeric(12,2) not null default 0,
  advances_total numeric(12,2) not null default 0,
  net_paid numeric(12,2) not null default 0,
  observations text,
  responsible_signature text,
  motoboy_signature text,
  html text,
  data jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.settings (
  key text primary key,
  value jsonb not null default '{}'::jsonb,
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.backups (
  id uuid primary key default gen_random_uuid(),
  owner text not null default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  backup_name text,
  data jsonb not null,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now()
);

create table if not exists public.change_logs (
  id uuid primary key default gen_random_uuid(),
  table_name text,
  record_id text,
  action text not null,
  detail text,
  owner text default 'BASE'
    check (owner in ('GIL', 'SALES', 'GUILHERME', 'BASE')),
  old_data jsonb,
  new_data jsonb,
  metadata jsonb not null default '{}'::jsonb,
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now()
);

-- =========================================================
-- PARTE 2/4 - Funcoes
-- =========================================================

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
  elsif v in ('base', 'empresa', 'operacao', 'operação') then
    return 'BASE';
  end if;

  return upper(coalesce(input, 'BASE'));
end;
$$;

create or replace function public.safe_date(input text)
returns date
language plpgsql
immutable
as $$
begin
  if input is null or trim(input) = '' then
    return null;
  end if;
  return input::date;
exception when others then
  return null;
end;
$$;

create or replace function public.safe_numeric(input text)
returns numeric
language plpgsql
immutable
as $$
declare
  cleaned text;
begin
  if input is null or trim(input) = '' then
    return null;
  end if;
  cleaned := replace(replace(replace(input, 'R$', ''), '.', ''), ',', '.');
  return trim(cleaned)::numeric;
exception when others then
  return null;
end;
$$;

create or replace function public.safe_int(input text)
returns integer
language plpgsql
immutable
as $$
begin
  if input is null or trim(input) = '' then
    return null;
  end if;
  return regexp_replace(input, '[^0-9-]', '', 'g')::integer;
exception when others then
  return null;
end;
$$;

create or replace function public.fortnight_start(input_date date)
returns date
language sql
immutable
as $$
  select case
    when input_date is null then null
    when extract(day from input_date)::int <= 15 then date_trunc('month', input_date)::date
    else (date_trunc('month', input_date)::date + interval '15 day')::date
  end;
$$;

create or replace function public.fortnight_end(input_date date)
returns date
language sql
immutable
as $$
  select case
    when input_date is null then null
    when extract(day from input_date)::int <= 15 then (date_trunc('month', input_date)::date + interval '14 day')::date
    else (date_trunc('month', input_date)::date + interval '1 month - 1 day')::date
  end;
$$;

create or replace function public.fortnight_label(start_date date, end_date date)
returns text
language sql
immutable
as $$
  select case
    when start_date is null or end_date is null then null
    when extract(day from start_date)::int = 1 then '1ª quinzena ' || to_char(start_date, 'MM/YYYY')
    else '2ª quinzena ' || to_char(start_date, 'MM/YYYY')
  end;
$$;

create or replace function public.variable_expense_discount_start(expense_date date)
returns date
language sql
immutable
as $$
  select case
    when expense_date is null then null
    when extract(day from expense_date)::int <= 15
      then (date_trunc('month', expense_date)::date + interval '15 day')::date
    else (date_trunc('month', expense_date)::date + interval '1 month')::date
  end;
$$;

create or replace function public.variable_expense_discount_end(expense_date date)
returns date
language sql
immutable
as $$
  select case
    when expense_date is null then null
    when extract(day from expense_date)::int <= 15
      then (date_trunc('month', expense_date)::date + interval '1 month - 1 day')::date
    else (date_trunc('month', expense_date)::date + interval '1 month 14 day')::date
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
    (select role from public.profiles where id = auth.uid()),
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
  select public.current_profile_role() = 'admin';
$$;

create or replace function public.can_write_responsible(target text)
returns boolean
language sql
security definer
set search_path = public
stable
as $$
  select case
    when public.is_admin() then true
    when public.normalize_responsible(target) = 'BASE' then false
    else public.normalize_responsible(target) = public.current_profile_role()
  end;
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

create or replace function public.apply_default_actor()
returns trigger
language plpgsql
as $$
begin
  if tg_op = 'INSERT' and new.created_by is null then
    new.created_by = auth.uid();
  end if;
  return new;
end;
$$;

create or replace function public.prepare_motoboy()
returns trigger
language plpgsql
as $$
begin
  new.owner := public.normalize_responsible(coalesce(new.data->>'owner', new.data->>'responsible', new.owner, 'BASE'));
  new.name := coalesce(nullif(new.name, ''), new.data->>'name', new.data->>'nome');
  new.region := coalesce(new.region, new.data->>'region', new.data->>'regiao');
  new.work_type := coalesce(new.data->>'workType', new.data->>'type', nullif(new.work_type, ''), 'com_coleta');
  if new.work_type in ('Com coleta', 'com coleta') then new.work_type := 'com_coleta'; end if;
  if new.work_type in ('Sem coleta', 'sem coleta') then new.work_type := 'sem_coleta'; end if;
  if new.work_type in ('Freelancer', 'freelancer') then new.work_type := 'freelancer'; end if;
  new.has_collection := new.work_type = 'com_coleta';
  new.rate_ml := coalesce(public.safe_numeric(new.data->>'mlValue'), public.safe_numeric(new.data->>'valorMl'), new.rate_ml);
  new.rate_shopee := coalesce(public.safe_numeric(new.data->>'shopeeValue'), public.safe_numeric(new.data->>'valorShopee'), new.rate_shopee);
  new.rate_avulso := coalesce(public.safe_numeric(new.data->>'avulsoValue'), public.safe_numeric(new.data->>'valorAvulso'), new.rate_avulso);
  new.status := lower(coalesce(new.data->>'status', nullif(new.status, ''), 'ativo'));
  return new;
end;
$$;

create or replace function public.prepare_daily_launch()
returns trigger
language plpgsql
as $$
begin
  new.owner := public.normalize_responsible(coalesce(new.data->>'owner', new.data->>'responsible', new.data->>'partner', new.owner, 'BASE'));
  new.launch_date := coalesce(public.safe_date(new.data->>'date'), public.safe_date(new.data->>'data'), new.launch_date, current_date);
  new.motoboy_name := coalesce(new.motoboy_name, new.data->>'riderName', new.data->>'motoboyName', new.data->>'motoboy');
  new.launch_type := coalesce(new.data->>'launchType', new.data->>'collectionType', new.launch_type, 'com_coleta');
  if lower(new.launch_type) in ('sem coleta', 'sem_coleta', 'no_collection') then
    new.launch_type := 'sem_coleta';
  else
    new.launch_type := 'com_coleta';
  end if;
  new.ml_qty := coalesce(public.safe_int(new.data->>'mlQty'), public.safe_int(new.data->>'ml'), new.ml_qty, 0);
  new.shopee_qty := coalesce(public.safe_int(new.data->>'shopeeQty'), public.safe_int(new.data->>'shopee'), new.shopee_qty, 0);
  new.avulso_qty := coalesce(public.safe_int(new.data->>'avulsoQty'), public.safe_int(new.data->>'avulso'), new.avulso_qty, 0);
  new.rate_ml := coalesce(public.safe_numeric(new.data->>'mlValue'), public.safe_numeric(new.data->>'rateMl'), new.rate_ml);
  new.rate_shopee := coalesce(public.safe_numeric(new.data->>'shopeeValue'), public.safe_numeric(new.data->>'rateShopee'), new.rate_shopee);
  new.rate_avulso := coalesce(public.safe_numeric(new.data->>'avulsoValue'), public.safe_numeric(new.data->>'rateAvulso'), new.rate_avulso);
  new.gross_total := (new.ml_qty * new.rate_ml) + (new.shopee_qty * new.rate_shopee) + (new.avulso_qty * new.rate_avulso);
  new.net_total := new.gross_total - coalesce(new.discounts_total, 0) - coalesce(new.advances_total, 0) + coalesce(new.bonuses_total, 0);
  return new;
end;
$$;

create or replace function public.prepare_package_entry()
returns trigger
language plpgsql
as $$
begin
  new.responsible := public.normalize_responsible(coalesce(new.data->>'responsible', new.data->>'partner', new.responsible, new.owner, 'BASE'));
  new.owner := new.responsible;
  new.entry_date := coalesce(public.safe_date(new.data->>'date'), public.safe_date(new.data->>'data'), new.entry_date, current_date);
  new.ml_qty := coalesce(public.safe_int(new.data->>'mlQty'), public.safe_int(new.data->>'ml'), new.ml_qty, 0);
  new.shopee_qty := coalesce(public.safe_int(new.data->>'shopeeQty'), public.safe_int(new.data->>'shopee'), new.shopee_qty, 0);
  new.total_packages := coalesce(public.safe_int(new.data->>'totalPackages'), new.ml_qty + new.shopee_qty, 0);
  new.value_ml := new.ml_qty * new.rate_ml;
  new.value_shopee := new.shopee_qty * new.rate_shopee;
  new.total_value := new.value_ml + new.value_shopee;
  return new;
end;
$$;

create or replace function public.prepare_discount()
returns trigger
language plpgsql
as $$
begin
  new.responsible := public.normalize_responsible(coalesce(new.data->>'responsible', new.data->>'partner', new.responsible, new.owner, 'BASE'));
  new.owner := new.responsible;
  new.discount_date := coalesce(public.safe_date(new.data->>'date'), public.safe_date(new.data->>'data'), new.discount_date, current_date);
  new.motoboy_name := coalesce(new.motoboy_name, new.data->>'riderName', new.data->>'motoboyName', new.data->>'motoboy');
  new.discount_type := upper(coalesce(new.data->>'type', new.data->>'discountType', new.discount_type, 'OUTROS'));
  if new.discount_type in ('OCORRÊNCIA', 'OCORRENCIA') then new.discount_type := 'OCORRENCIA'; end if;
  if new.discount_type in ('EXTRAVIO/OCORRÊNCIA') then new.discount_type := 'EXTRAVIO/OCORRENCIA'; end if;
  new.value := coalesce(public.safe_numeric(new.data->>'value'), public.safe_numeric(new.data->>'amount'), new.value, 0);
  new.reason := coalesce(new.reason, new.data->>'reason', new.data->>'motivo');
  new.occurrence := coalesce(new.occurrence, new.data->>'occurrence', new.data->>'ocorrencia');
  new.observation := coalesce(new.observation, new.data->>'observation', new.data->>'observacao');
  new.original_sheet := coalesce(new.original_sheet, new.data->>'sheet', new.data->>'abaOriginal');
  new.original_line := coalesce(new.original_line, public.safe_int(new.data->>'line'), public.safe_int(new.data->>'linhaOriginal'));
  new.original_column := coalesce(new.original_column, new.data->>'column', new.data->>'colunaOriginal');
  new.unique_import_key := coalesce(
    new.unique_import_key,
    md5(concat_ws('|', new.responsible, coalesce(new.motoboy_name, ''), new.discount_type, new.value::text, coalesce(new.original_sheet, ''), coalesce(new.original_line::text, ''), coalesce(new.original_column, '')))
  );
  return new;
end;
$$;

create or replace function public.prepare_expense()
returns trigger
language plpgsql
as $$
declare
  manual_key text;
  discount_start date;
  discount_end date;
begin
  new.responsible := public.normalize_responsible(coalesce(new.data->>'responsible', new.data->>'partner', new.responsible, new.owner, 'BASE'));
  new.owner := new.responsible;
  new.expense_date := coalesce(public.safe_date(new.data->>'date'), public.safe_date(new.data->>'data'), new.expense_date, current_date);
  new.expense_type := lower(coalesce(new.data->>'type', new.data->>'expenseType', new.expense_type, 'variavel'));
  if new.expense_type in ('fixa', 'fixed') then new.expense_type := 'fixa'; else new.expense_type := 'variavel'; end if;
  new.category := coalesce(nullif(new.category, ''), new.data->>'category', new.data->>'categoria', 'Sem categoria');
  new.description := coalesce(new.description, new.data->>'description', new.data->>'descricao');
  new.value := coalesce(public.safe_numeric(new.data->>'value'), public.safe_numeric(new.data->>'amount'), new.value, 0);
  new.observation := coalesce(new.observation, new.data->>'observation', new.data->>'observacao');
  new.origin_period_start := public.fortnight_start(new.expense_date);
  new.origin_period_end := public.fortnight_end(new.expense_date);
  new.origin_period_label := public.fortnight_label(new.origin_period_start, new.origin_period_end);

  if new.expense_type = 'variavel' then
    manual_key := coalesce(new.data->>'discountPeriodKey', '');
    if manual_key ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}\|[0-9]{4}-[0-9]{2}-[0-9]{2}$' then
      discount_start := split_part(manual_key, '|', 1)::date;
      discount_end := split_part(manual_key, '|', 2)::date;
    else
      discount_start := public.variable_expense_discount_start(new.expense_date);
      discount_end := public.variable_expense_discount_end(new.expense_date);
    end if;
    new.discount_period_start := coalesce(new.discount_period_start, discount_start);
    new.discount_period_end := coalesce(new.discount_period_end, discount_end);
    new.discount_period_label := coalesce(new.discount_period_label, new.data->>'discountPeriodLabel', public.fortnight_label(new.discount_period_start, new.discount_period_end));
  end if;

  return new;
end;
$$;

create or replace function public.prepare_responsible_record()
returns trigger
language plpgsql
as $$
begin
  new.responsible := public.normalize_responsible(coalesce(new.data->>'responsible', new.data->>'partner', new.responsible, new.owner, 'BASE'));
  new.owner := new.responsible;
  return new;
end;
$$;

create or replace function public.audit_row_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  rec_id text;
  rec_owner text;
begin
  if tg_op = 'DELETE' then
    rec_id := old.id::text;
    rec_owner := public.normalize_responsible(coalesce(old.owner, 'BASE'));
    insert into public.change_logs(table_name, record_id, action, detail, owner, old_data, new_data, created_by)
    values (tg_table_name, rec_id, 'delete', 'Registro removido', rec_owner, to_jsonb(old), null, auth.uid());
    return old;
  elsif tg_op = 'UPDATE' then
    rec_id := new.id::text;
    rec_owner := public.normalize_responsible(coalesce(new.owner, 'BASE'));
    insert into public.change_logs(table_name, record_id, action, detail, owner, old_data, new_data, created_by)
    values (tg_table_name, rec_id, 'update', 'Registro atualizado', rec_owner, to_jsonb(old), to_jsonb(new), auth.uid());
    return new;
  else
    rec_id := new.id::text;
    rec_owner := public.normalize_responsible(coalesce(new.owner, 'BASE'));
    insert into public.change_logs(table_name, record_id, action, detail, owner, old_data, new_data, created_by)
    values (tg_table_name, rec_id, 'insert', 'Registro criado', rec_owner, null, to_jsonb(new), auth.uid());
    return new;
  end if;
end;
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
    else 'BASE'
  end;

  insert into public.profiles(id, username, role, full_name)
  values (new.id, user_name, user_role, coalesce(new.raw_user_meta_data->>'full_name', user_name))
  on conflict (id) do update
    set username = excluded.username,
        role = excluded.role,
        full_name = excluded.full_name,
        updated_at = now();

  return new;
end;
$$;

-- =========================================================
-- PARTE 3/4 - Triggers, perfis iniciais e indices
-- =========================================================

drop trigger if exists trg_profiles_touch on public.profiles;
create trigger trg_profiles_touch before update on public.profiles
for each row execute function public.touch_updated_at();

drop trigger if exists on_auth_user_created_financeiro on auth.users;
drop trigger if exists trg_auth_user_profile on auth.users;
create trigger trg_auth_user_profile after insert or update on auth.users
for each row execute function public.handle_new_user();

drop trigger if exists trg_motoboys_touch on public.motoboys;
drop trigger if exists trg_motoboys_actor on public.motoboys;
drop trigger if exists trg_motoboys_prepare on public.motoboys;
drop trigger if exists trg_motoboys_audit on public.motoboys;
create trigger trg_motoboys_prepare before insert or update on public.motoboys for each row execute function public.prepare_motoboy();
create trigger trg_motoboys_actor before insert on public.motoboys for each row execute function public.apply_default_actor();
create trigger trg_motoboys_touch before update on public.motoboys for each row execute function public.touch_updated_at();
create trigger trg_motoboys_audit after insert or update or delete on public.motoboys for each row execute function public.audit_row_change();

drop trigger if exists trg_daily_prepare on public.daily_launches;
drop trigger if exists trg_daily_actor on public.daily_launches;
drop trigger if exists trg_daily_touch on public.daily_launches;
drop trigger if exists trg_daily_audit on public.daily_launches;
create trigger trg_daily_prepare before insert or update on public.daily_launches for each row execute function public.prepare_daily_launch();
create trigger trg_daily_actor before insert on public.daily_launches for each row execute function public.apply_default_actor();
create trigger trg_daily_touch before update on public.daily_launches for each row execute function public.touch_updated_at();
create trigger trg_daily_audit after insert or update or delete on public.daily_launches for each row execute function public.audit_row_change();

drop trigger if exists trg_package_prepare on public.package_entries;
drop trigger if exists trg_package_actor on public.package_entries;
drop trigger if exists trg_package_touch on public.package_entries;
drop trigger if exists trg_package_audit on public.package_entries;
create trigger trg_package_prepare before insert or update on public.package_entries for each row execute function public.prepare_package_entry();
create trigger trg_package_actor before insert on public.package_entries for each row execute function public.apply_default_actor();
create trigger trg_package_touch before update on public.package_entries for each row execute function public.touch_updated_at();
create trigger trg_package_audit after insert or update or delete on public.package_entries for each row execute function public.audit_row_change();

drop trigger if exists trg_discounts_prepare on public.discounts;
drop trigger if exists trg_discounts_actor on public.discounts;
drop trigger if exists trg_discounts_touch on public.discounts;
drop trigger if exists trg_discounts_audit on public.discounts;
create trigger trg_discounts_prepare before insert or update on public.discounts for each row execute function public.prepare_discount();
create trigger trg_discounts_actor before insert on public.discounts for each row execute function public.apply_default_actor();
create trigger trg_discounts_touch before update on public.discounts for each row execute function public.touch_updated_at();
create trigger trg_discounts_audit after insert or update or delete on public.discounts for each row execute function public.audit_row_change();

drop trigger if exists trg_expenses_prepare on public.expenses;
drop trigger if exists trg_expenses_actor on public.expenses;
drop trigger if exists trg_expenses_touch on public.expenses;
drop trigger if exists trg_expenses_audit on public.expenses;
create trigger trg_expenses_prepare before insert or update on public.expenses for each row execute function public.prepare_expense();
create trigger trg_expenses_actor before insert on public.expenses for each row execute function public.apply_default_actor();
create trigger trg_expenses_touch before update on public.expenses for each row execute function public.touch_updated_at();
create trigger trg_expenses_audit after insert or update or delete on public.expenses for each row execute function public.audit_row_change();

drop trigger if exists trg_payments_prepare on public.payments;
drop trigger if exists trg_payments_actor on public.payments;
drop trigger if exists trg_payments_touch on public.payments;
drop trigger if exists trg_payments_audit on public.payments;
create trigger trg_payments_prepare before insert or update on public.payments for each row execute function public.prepare_responsible_record();
create trigger trg_payments_actor before insert on public.payments for each row execute function public.apply_default_actor();
create trigger trg_payments_touch before update on public.payments for each row execute function public.touch_updated_at();
create trigger trg_payments_audit after insert or update or delete on public.payments for each row execute function public.audit_row_change();

drop trigger if exists trg_receipts_prepare on public.receipts;
drop trigger if exists trg_receipts_actor on public.receipts;
drop trigger if exists trg_receipts_touch on public.receipts;
drop trigger if exists trg_receipts_audit on public.receipts;
create trigger trg_receipts_prepare before insert or update on public.receipts for each row execute function public.prepare_responsible_record();
create trigger trg_receipts_actor before insert on public.receipts for each row execute function public.apply_default_actor();
create trigger trg_receipts_touch before update on public.receipts for each row execute function public.touch_updated_at();
create trigger trg_receipts_audit after insert or update or delete on public.receipts for each row execute function public.audit_row_change();

drop trigger if exists trg_settings_touch on public.settings;
drop trigger if exists trg_settings_actor on public.settings;
create trigger trg_settings_actor before insert on public.settings for each row execute function public.apply_default_actor();
create trigger trg_settings_touch before update on public.settings for each row execute function public.touch_updated_at();

drop trigger if exists trg_backups_actor on public.backups;
drop trigger if exists trg_backups_audit on public.backups;
create trigger trg_backups_actor before insert on public.backups for each row execute function public.apply_default_actor();
create trigger trg_backups_audit after insert or update or delete on public.backups for each row execute function public.audit_row_change();

insert into public.profiles(id, username, role, full_name)
select
  u.id,
  lower(split_part(u.email, '@', 1)) as username,
  case
    when lower(split_part(u.email, '@', 1)) = 'admin' then 'admin'
    when lower(split_part(u.email, '@', 1)) = 'gil' then 'GIL'
    when lower(split_part(u.email, '@', 1)) = 'sales' then 'SALES'
    when lower(split_part(u.email, '@', 1)) in ('guilherme', 'gm') then 'GUILHERME'
    else 'BASE'
  end as role,
  lower(split_part(u.email, '@', 1)) as full_name
from auth.users u
where lower(split_part(u.email, '@', 1)) in ('admin', 'gil', 'sales', 'guilherme', 'gm')
on conflict (id) do update
  set username = excluded.username,
      role = excluded.role,
      full_name = excluded.full_name,
      updated_at = now();

create index if not exists idx_profiles_role on public.profiles(role);
create index if not exists idx_motoboys_name on public.motoboys(name);
create index if not exists idx_motoboys_status on public.motoboys(status);
create index if not exists idx_motoboys_owner on public.motoboys(owner);
create index if not exists idx_daily_launches_owner_date on public.daily_launches(owner, launch_date);
create index if not exists idx_daily_launches_motoboy on public.daily_launches(motoboy_id, motoboy_name);
create index if not exists idx_package_entries_owner_date on public.package_entries(owner, entry_date);
create index if not exists idx_discounts_owner_date on public.discounts(owner, discount_date);
create index if not exists idx_discounts_motoboy on public.discounts(motoboy_id, motoboy_name);
create index if not exists idx_expenses_owner_date on public.expenses(owner, expense_date);
create index if not exists idx_expenses_discount_period on public.expenses(discount_period_start, discount_period_end);
create index if not exists idx_expenses_type on public.expenses(expense_type);
create index if not exists idx_payments_owner_period on public.payments(owner, period_start, period_end);
create index if not exists idx_receipts_owner_period on public.receipts(owner, period_start, period_end);
create index if not exists idx_change_logs_table_record on public.change_logs(table_name, record_id);
create index if not exists idx_change_logs_created_at on public.change_logs(created_at desc);
create index if not exists idx_motoboys_data_gin on public.motoboys using gin(data);
create index if not exists idx_daily_data_gin on public.daily_launches using gin(data);
create index if not exists idx_package_data_gin on public.package_entries using gin(data);
create index if not exists idx_discounts_data_gin on public.discounts using gin(data);
create index if not exists idx_expenses_data_gin on public.expenses using gin(data);

-- =========================================================
-- PARTE 4/4 - RLS, policies, views e grants
-- =========================================================

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

drop policy if exists profiles_select_all on public.profiles;
drop policy if exists profiles_admin_all on public.profiles;
drop policy if exists profiles_select_authenticated on public.profiles;
drop policy if exists profiles_admin_update on public.profiles;
create policy profiles_select_all on public.profiles for select to authenticated using (true);
create policy profiles_admin_all on public.profiles for all to authenticated using (public.is_admin()) with check (public.is_admin());

drop policy if exists motoboys_select_all on public.motoboys;
drop policy if exists motoboys_admin_all on public.motoboys;
drop policy if exists motoboys_admin_write on public.motoboys;
create policy motoboys_select_all on public.motoboys for select to authenticated using (true);
create policy motoboys_admin_all on public.motoboys for all to authenticated using (public.is_admin()) with check (public.is_admin());

drop policy if exists daily_select_all on public.daily_launches;
drop policy if exists daily_admin_all on public.daily_launches;
drop policy if exists daily_insert_owner on public.daily_launches;
drop policy if exists daily_update_owner on public.daily_launches;
drop policy if exists daily_delete_owner on public.daily_launches;
create policy daily_select_all on public.daily_launches for select to authenticated using (true);
create policy daily_admin_all on public.daily_launches for all to authenticated using (public.is_admin()) with check (public.is_admin());

drop policy if exists package_select_all on public.package_entries;
drop policy if exists package_insert_owner on public.package_entries;
drop policy if exists package_update_owner on public.package_entries;
drop policy if exists package_delete_owner on public.package_entries;
drop policy if exists entries_select_all on public.package_entries;
drop policy if exists entries_insert_owner on public.package_entries;
drop policy if exists entries_update_owner on public.package_entries;
drop policy if exists entries_delete_owner on public.package_entries;
create policy package_select_all on public.package_entries for select to authenticated using (true);
create policy package_insert_owner on public.package_entries for insert to authenticated with check (public.can_write_responsible(owner));
create policy package_update_owner on public.package_entries for update to authenticated using (public.can_write_responsible(owner)) with check (public.can_write_responsible(owner));
create policy package_delete_owner on public.package_entries for delete to authenticated using (public.can_write_responsible(owner));

drop policy if exists discounts_select_all on public.discounts;
drop policy if exists discounts_insert_owner on public.discounts;
drop policy if exists discounts_update_owner on public.discounts;
drop policy if exists discounts_delete_owner on public.discounts;
create policy discounts_select_all on public.discounts for select to authenticated using (true);
create policy discounts_insert_owner on public.discounts for insert to authenticated with check (public.can_write_responsible(owner));
create policy discounts_update_owner on public.discounts for update to authenticated using (public.can_write_responsible(owner)) with check (public.can_write_responsible(owner));
create policy discounts_delete_owner on public.discounts for delete to authenticated using (public.can_write_responsible(owner));

drop policy if exists expenses_select_all on public.expenses;
drop policy if exists expenses_insert_owner on public.expenses;
drop policy if exists expenses_update_owner on public.expenses;
drop policy if exists expenses_delete_owner on public.expenses;
create policy expenses_select_all on public.expenses for select to authenticated using (true);
create policy expenses_insert_owner on public.expenses for insert to authenticated with check (public.can_write_responsible(owner));
create policy expenses_update_owner on public.expenses for update to authenticated using (public.can_write_responsible(owner)) with check (public.can_write_responsible(owner));
create policy expenses_delete_owner on public.expenses for delete to authenticated using (public.can_write_responsible(owner));

drop policy if exists payments_select_all on public.payments;
drop policy if exists payments_insert_owner on public.payments;
drop policy if exists payments_update_owner on public.payments;
drop policy if exists payments_delete_owner on public.payments;
create policy payments_select_all on public.payments for select to authenticated using (true);
create policy payments_insert_owner on public.payments for insert to authenticated with check (public.can_write_responsible(owner));
create policy payments_update_owner on public.payments for update to authenticated using (public.can_write_responsible(owner)) with check (public.can_write_responsible(owner));
create policy payments_delete_owner on public.payments for delete to authenticated using (public.can_write_responsible(owner));

drop policy if exists receipts_select_all on public.receipts;
drop policy if exists receipts_insert_owner on public.receipts;
drop policy if exists receipts_update_owner on public.receipts;
drop policy if exists receipts_delete_owner on public.receipts;
create policy receipts_select_all on public.receipts for select to authenticated using (true);
create policy receipts_insert_owner on public.receipts for insert to authenticated with check (public.can_write_responsible(owner));
create policy receipts_update_owner on public.receipts for update to authenticated using (public.can_write_responsible(owner)) with check (public.can_write_responsible(owner));
create policy receipts_delete_owner on public.receipts for delete to authenticated using (public.can_write_responsible(owner));

drop policy if exists settings_select_all on public.settings;
drop policy if exists settings_admin_all on public.settings;
drop policy if exists settings_admin_write on public.settings;
drop policy if exists settings_select_authenticated on public.settings;
drop policy if exists settings_insert_authenticated on public.settings;
drop policy if exists settings_update_authenticated on public.settings;
drop policy if exists settings_delete_admin on public.settings;
create policy settings_select_authenticated on public.settings for select to authenticated using (true);
create policy settings_insert_authenticated on public.settings for insert to authenticated with check (auth.role() = 'authenticated');
create policy settings_update_authenticated on public.settings for update to authenticated using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy settings_delete_admin on public.settings for delete to authenticated using (public.is_admin());

drop policy if exists backups_select_admin on public.backups;
drop policy if exists backups_admin_all on public.backups;
create policy backups_select_admin on public.backups for select to authenticated using (public.is_admin());
create policy backups_admin_all on public.backups for all to authenticated using (public.is_admin()) with check (public.is_admin());

drop policy if exists change_logs_select_all on public.change_logs;
drop policy if exists change_logs_insert_authenticated on public.change_logs;
drop policy if exists change_logs_admin_delete on public.change_logs;
create policy change_logs_select_all on public.change_logs for select to authenticated using (true);
create policy change_logs_insert_authenticated on public.change_logs for insert to authenticated with check (true);
create policy change_logs_admin_delete on public.change_logs for delete to authenticated using (public.is_admin());

create or replace view public.v_variable_expenses_by_discount_period as
select
  owner,
  responsible,
  discount_period_start,
  discount_period_end,
  discount_period_label,
  sum(value) as total_value,
  count(*) as total_records
from public.expenses
where expense_type = 'variavel'
group by owner, responsible, discount_period_start, discount_period_end, discount_period_label;

create or replace view public.v_dashboard_summary as
select
  coalesce((select sum(gross_total) from public.daily_launches), 0) as total_bruto,
  coalesce((select sum(value) from public.discounts), 0) as total_descontos,
  coalesce((select sum(value) from public.discounts where discount_type = 'VALE'), 0) as total_vales,
  coalesce((select sum(net_total) from public.daily_launches), 0) as total_liquido,
  coalesce((select sum(net_paid) from public.payments where status = 'pago'), 0) as total_pago,
  coalesce((select sum(net_paid) from public.payments where status <> 'pago'), 0) as total_pendente,
  coalesce((select count(*) from public.motoboys where status = 'ativo'), 0) as quantidade_motoboys,
  coalesce((select sum(ml_qty + shopee_qty + avulso_qty) from public.daily_launches), 0) as quantidade_pacotes,
  coalesce((select sum(value) from public.expenses where expense_type = 'variavel'), 0) as despesas_variaveis,
  coalesce((select sum(value) from public.expenses where expense_type = 'fixa'), 0) as despesas_fixas;

grant usage on schema public to anon, authenticated;
grant select on public.v_variable_expenses_by_discount_period to authenticated;
grant select on public.v_dashboard_summary to authenticated;

-- Fim do schema.
