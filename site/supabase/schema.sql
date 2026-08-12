-- ValeFluxo: execute no SQL Editor de um projeto Supabase novo.
create extension if not exists "pgcrypto";

create type public.user_role as enum ('admin', 'operator');
create type public.card_type as enum ('BHBus', 'Ótimo', 'Outro');
create type public.recharge_status as enum ('Pendente', 'Recarregado', 'Atrasado', 'Próximo');

create table public.stores (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  active boolean not null default true,
  created_at timestamptz not null default now()
);

create table public.users (
  id uuid primary key references auth.users(id) on delete cascade,
  name text not null,
  role public.user_role not null default 'operator',
  created_at timestamptz not null default now()
);

create table public.employees (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  cpf text not null unique,
  role_title text not null,
  store_id uuid references public.stores(id),
  phone text,
  card_type public.card_type not null,
  card_number text not null,
  hired_at date not null,
  active boolean not null default true,
  credit_day smallint check (credit_day between 1 and 31),
  specific_credit_date date,
  advance_days smallint not null default 3 check (advance_days between 1 and 10),
  created_at timestamptz not null default now(),
  constraint credit_schedule check (credit_day is not null or specific_credit_date is not null)
);

create table public.recharges (
  id uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.employees(id) on delete cascade,
  credit_date date not null,
  advance_days smallint not null default 3,
  recharge_date date generated always as (credit_date - advance_days) stored,
  status public.recharge_status not null default 'Próximo',
  completed_at timestamptz,
  completed_by uuid references public.users(id),
  amount numeric(10,2),
  created_at timestamptz not null default now(),
  unique(employee_id, credit_date)
);

create table public.settings (
  id uuid primary key default gen_random_uuid(),
  key text not null unique,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

-- Mantém os status coerentes diariamente; pode ser chamado pelo Supabase Cron.
create or replace function public.refresh_recharge_statuses()
returns void language plpgsql security definer set search_path = public as $$
begin
  update recharges set status = case
    when completed_at is not null then 'Recarregado'::recharge_status
    when recharge_date < current_date then 'Atrasado'::recharge_status
    when recharge_date = current_date then 'Pendente'::recharge_status
    else 'Próximo'::recharge_status end;
end $$;

-- Cria ocorrências dos próximos 40 dias para funcionários com agenda mensal.
create or replace function public.generate_upcoming_recharges()
returns void language plpgsql security definer set search_path = public as $$
declare e employees%rowtype; target_date date;
begin
  for e in select * from employees where active loop
    if e.specific_credit_date is not null and e.specific_credit_date between current_date and current_date + 40 then
      insert into recharges(employee_id, credit_date, advance_days) values(e.id,e.specific_credit_date,e.advance_days) on conflict do nothing;
    elsif e.credit_day is not null then
      for target_date in select (d + (least(e.credit_day,extract(day from (d + interval '1 month - 1 day')))::int - 1))::date from generate_series(date_trunc('month',current_date),date_trunc('month',current_date + interval '40 day'),interval '1 month') d loop
        insert into recharges(employee_id, credit_date, advance_days) values(e.id,target_date,e.advance_days) on conflict do nothing;
      end loop;
    end if;
  end loop;
  perform refresh_recharge_statuses();
end $$;

alter table stores enable row level security;
alter table users enable row level security;
alter table employees enable row level security;
alter table recharges enable row level security;
alter table settings enable row level security;

create or replace function public.current_role() returns user_role language sql stable security definer set search_path=public as $$ select role from users where id=auth.uid() $$;
create policy "authenticated read stores" on stores for select to authenticated using (true);
create policy "authenticated read users" on users for select to authenticated using (true);
create policy "authenticated read employees" on employees for select to authenticated using (true);
create policy "admins manage employees" on employees for all to authenticated using (current_role()='admin') with check (current_role()='admin');
create policy "authenticated read recharges" on recharges for select to authenticated using (true);
create policy "staff update recharges" on recharges for update to authenticated using (true) with check (true);
create policy "admins manage recharges" on recharges for insert to authenticated with check (current_role()='admin');
create policy "admins delete recharges" on recharges for delete to authenticated using (current_role()='admin');
create policy "admins manage stores" on stores for all to authenticated using (current_role()='admin') with check (current_role()='admin');
create policy "authenticated read settings" on settings for select to authenticated using (true);
create policy "admins manage settings" on settings for all to authenticated using (current_role()='admin') with check (current_role()='admin');

insert into public.settings(key,value) values ('default_advance_days','3'::jsonb) on conflict do nothing;
