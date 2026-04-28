create extension if not exists pgcrypto;

create table if not exists public.facilities (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  area text,
  manager_note text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.vehicles (
  id uuid primary key default gen_random_uuid(),
  facility_id uuid not null references public.facilities(id) on delete cascade,
  plate_number text not null,
  insurance_end_date date not null,
  insurance_file_path text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.huv_usage_records (
  id uuid primary key default gen_random_uuid(),
  vehicle_id uuid not null references public.vehicles(id) on delete cascade,
  record_type text not null default 'reservation'
    check (record_type in ('reservation', 'history')),
  usage_date date not null,
  start_time time not null,
  end_time time not null,
  sales_amount numeric(12, 0) not null default 0,
  maintenance_date date,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists vehicles_plate_number_key
  on public.vehicles (plate_number);

create index if not exists vehicles_facility_id_idx
  on public.vehicles (facility_id);

create index if not exists huv_usage_records_vehicle_id_idx
  on public.huv_usage_records (vehicle_id);

create index if not exists huv_usage_records_usage_date_idx
  on public.huv_usage_records (usage_date desc);

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists facilities_set_updated_at on public.facilities;
create trigger facilities_set_updated_at
before update on public.facilities
for each row
execute function public.set_updated_at();

drop trigger if exists vehicles_set_updated_at on public.vehicles;
create trigger vehicles_set_updated_at
before update on public.vehicles
for each row
execute function public.set_updated_at();

drop trigger if exists huv_usage_records_set_updated_at on public.huv_usage_records;
create trigger huv_usage_records_set_updated_at
before update on public.huv_usage_records
for each row
execute function public.set_updated_at();

alter table public.facilities enable row level security;
alter table public.vehicles enable row level security;
alter table public.huv_usage_records enable row level security;

drop policy if exists "public can read facilities" on public.facilities;
create policy "public can read facilities"
on public.facilities
for select
to anon, authenticated
using (true);

drop policy if exists "public can read vehicles" on public.vehicles;
create policy "public can read vehicles"
on public.vehicles
for select
to anon, authenticated
using (true);

drop policy if exists "public can read huv usage records" on public.huv_usage_records;
drop policy if exists "authenticated can read huv usage records" on public.huv_usage_records;
create policy "authenticated can read huv usage records"
on public.huv_usage_records
for select
to authenticated
using (true);

drop policy if exists "authenticated can insert facilities" on public.facilities;
create policy "authenticated can insert facilities"
on public.facilities
for insert
to authenticated
with check (true);

drop policy if exists "authenticated can update facilities" on public.facilities;
create policy "authenticated can update facilities"
on public.facilities
for update
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated can delete facilities" on public.facilities;
create policy "authenticated can delete facilities"
on public.facilities
for delete
to authenticated
using (true);

drop policy if exists "authenticated can insert vehicles" on public.vehicles;
create policy "authenticated can insert vehicles"
on public.vehicles
for insert
to authenticated
with check (true);

drop policy if exists "authenticated can update vehicles" on public.vehicles;
create policy "authenticated can update vehicles"
on public.vehicles
for update
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated can delete vehicles" on public.vehicles;
create policy "authenticated can delete vehicles"
on public.vehicles
for delete
to authenticated
using (true);

drop policy if exists "authenticated can insert huv usage records" on public.huv_usage_records;
create policy "authenticated can insert huv usage records"
on public.huv_usage_records
for insert
to authenticated
with check (true);

drop policy if exists "authenticated can update huv usage records" on public.huv_usage_records;
create policy "authenticated can update huv usage records"
on public.huv_usage_records
for update
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated can delete huv usage records" on public.huv_usage_records;
create policy "authenticated can delete huv usage records"
on public.huv_usage_records
for delete
to authenticated
using (true);
