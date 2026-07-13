create table if not exists public.korah_crm_state (
  id text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.korah_crm_state enable row level security;

drop policy if exists "korah crm public read" on public.korah_crm_state;
drop policy if exists "korah crm public insert" on public.korah_crm_state;
drop policy if exists "korah crm public update" on public.korah_crm_state;
drop policy if exists "korah crm public delete" on public.korah_crm_state;

create policy "korah crm public read"
on public.korah_crm_state
for select
to anon
using (true);

create policy "korah crm public insert"
on public.korah_crm_state
for insert
to anon
with check (true);

create policy "korah crm public update"
on public.korah_crm_state
for update
to anon
using (true)
with check (true);

create policy "korah crm public delete"
on public.korah_crm_state
for delete
to anon
using (true);
