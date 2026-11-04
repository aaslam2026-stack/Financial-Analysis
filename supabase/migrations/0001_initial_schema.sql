-- Finance Studio: initial schema for user accounts and saved models (Supabase / PostgreSQL).
-- Not wired into the app yet: models are saved in the browser today. Apply with `supabase db push`
-- once accounts are switched on. Every table uses row-level security so users only ever see their own data.

create extension if not exists "pgcrypto";

create table public.users (
  id uuid primary key references auth.users (id) on delete cascade,
  email text not null,
  display_name text,
  plan text not null default 'free',
  created_at timestamptz not null default now()
);

create table public.templates (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  industry text,
  description text,
  payload jsonb not null,                       -- default assumptions and settings
  is_public boolean not null default true,
  created_at timestamptz not null default now()
);

create table public.models (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.users (id) on delete cascade,
  template_id uuid references public.templates (id) on delete set null,
  name text not null check (char_length(name) between 1 and 80),
  company text,
  currency char(3) not null default 'USD',
  units text not null default 'thousands' check (units in ('actual', 'thousands', 'millions')),
  period_type text not null default 'annual' check (period_type in ('annual', 'quarterly', 'monthly')),
  years smallint not null default 5 check (years between 3 and 10),
  first_year smallint not null,
  settings jsonb not null default '{}'::jsonb,   -- interest method, revolver
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  deleted_at timestamptz                         -- soft delete: never remove without confirmation
);

create table public.scenarios (
  id uuid primary key default gen_random_uuid(),
  model_id uuid not null references public.models (id) on delete cascade,
  name text not null default 'Base case',
  is_base boolean not null default false,
  created_at timestamptz not null default now()
);

-- Money is stored as numeric (exact decimal), never floating point
create table public.assumptions (
  id uuid primary key default gen_random_uuid(),
  scenario_id uuid not null references public.scenarios (id) on delete cascade,
  key text not null,                             -- e.g. growth, grossMargin, dso, debtRepayment
  year_index smallint not null check (year_index between 0 and 9),
  value numeric(24, 8) not null,
  unique (scenario_id, key, year_index)
);

create table public.model_versions (
  id uuid primary key default gen_random_uuid(),
  model_id uuid not null references public.models (id) on delete cascade,
  version integer not null,
  snapshot jsonb not null,                       -- full model as exported to JSON
  created_by uuid references public.users (id),
  created_at timestamptz not null default now(),
  unique (model_id, version)
);

create index on public.models (user_id) where deleted_at is null;
create index on public.model_versions (model_id, version desc);

alter table public.users enable row level security;
alter table public.templates enable row level security;
alter table public.models enable row level security;
alter table public.scenarios enable row level security;
alter table public.assumptions enable row level security;
alter table public.model_versions enable row level security;

create policy "own profile" on public.users for all using (id = auth.uid()) with check (id = auth.uid());
create policy "read public templates" on public.templates for select using (is_public);
create policy "own models" on public.models for all using (user_id = auth.uid()) with check (user_id = auth.uid());
create policy "own scenarios" on public.scenarios for all
  using (exists (select 1 from public.models m where m.id = model_id and m.user_id = auth.uid()))
  with check (exists (select 1 from public.models m where m.id = model_id and m.user_id = auth.uid()));
create policy "own assumptions" on public.assumptions for all
  using (exists (select 1 from public.scenarios s join public.models m on m.id = s.model_id where s.id = scenario_id and m.user_id = auth.uid()))
  with check (exists (select 1 from public.scenarios s join public.models m on m.id = s.model_id where s.id = scenario_id and m.user_id = auth.uid()));
create policy "own versions" on public.model_versions for all
  using (exists (select 1 from public.models m where m.id = model_id and m.user_id = auth.uid()))
  with check (exists (select 1 from public.models m where m.id = model_id and m.user_id = auth.uid()));
