-- NHT Operations Dashboard V8 - role-based upload access
-- Run once AFTER the existing nht_dashboard_data table exists.
-- This keeps all authenticated users able to VIEW data, while only
-- ADMIN and UPLOADER roles can change the shared dataset.

create table if not exists public.nht_user_roles (
  user_id uuid primary key references auth.users(id) on delete cascade,
  role text not null check (role in ('admin','uploader','viewer')),
  created_at timestamptz not null default now()
);

alter table public.nht_user_roles enable row level security;

drop policy if exists "Users can read their own NHT role" on public.nht_user_roles;
create policy "Users can read their own NHT role"
on public.nht_user_roles
for select
to authenticated
using (user_id = auth.uid());

create or replace function public.is_nht_uploader()
returns boolean
language sql
security definer
stable
set search_path = public
as $$
  select exists (
    select 1
    from public.nht_user_roles
    where user_id = auth.uid()
      and role in ('admin','uploader')
  );
$$;

grant execute on function public.is_nht_uploader() to authenticated;

-- Replace the old "all authenticated users can write" policies.
drop policy if exists "Authenticated users can write NHT data" on public.nht_dashboard_data;
drop policy if exists "NHT uploaders can insert NHT data" on public.nht_dashboard_data;
drop policy if exists "NHT uploaders can update NHT data" on public.nht_dashboard_data;
drop policy if exists "NHT uploaders can delete NHT data" on public.nht_dashboard_data;

create policy "NHT uploaders can insert NHT data"
on public.nht_dashboard_data
for insert
to authenticated
with check (public.is_nht_uploader());

create policy "NHT uploaders can update NHT data"
on public.nht_dashboard_data
for update
to authenticated
using (public.is_nht_uploader())
with check (public.is_nht_uploader());

create policy "NHT uploaders can delete NHT data"
on public.nht_dashboard_data
for delete
to authenticated
using (public.is_nht_uploader());

grant select, insert, update, delete on public.nht_dashboard_data to authenticated;


-- Atomic month-scoped replacement used by the V8 upload button.
-- If a workbook contains only June, only June is replaced.
-- If it contains June-August, those three months are replaced.
create or replace function public.replace_nht_months(
  p_rows jsonb,
  p_months text[]
)
returns void
language plpgsql
security definer
set search_path = public
as $$
begin
  if auth.uid() is null then
    raise exception 'Authentication required';
  end if;

  if not public.is_nht_uploader() then
    raise exception 'Upload permission required';
  end if;

  if p_rows is null or jsonb_typeof(p_rows) <> 'array' or jsonb_array_length(p_rows) = 0 then
    raise exception 'No valid rows supplied';
  end if;

  if p_months is null or cardinality(p_months) = 0 then
    raise exception 'No months supplied';
  end if;

  delete from public.nht_dashboard_data
  where month = any(p_months);

  insert into public.nht_dashboard_data
  (month, location, batch, inflow, outflow, hr, training, throughput, joined, joining_throughput, updated_at, updated_by)
  select
    x.month, x.location, coalesce(x.batch,0), coalesce(x.inflow,0), coalesce(x.outflow,0),
    coalesce(x.hr,0), coalesce(x.training,0), coalesce(x.throughput,0), coalesce(x.joined,0),
    coalesce(x.joining_throughput,0), now(), auth.uid()
  from jsonb_to_recordset(p_rows) as x(
    month text, location text, batch numeric, inflow numeric, outflow numeric,
    hr numeric, training numeric, throughput numeric, joined numeric, joining_throughput numeric
  );
end;
$$;

grant execute on function public.replace_nht_months(jsonb, text[]) to authenticated;
notify pgrst, 'reload schema';

-- IMPORTANT:
-- After running this migration, assign roles to users.
-- Find each user's UUID in Supabase > Authentication > Users.
-- Then run one of these examples in SQL Editor:
--
-- INSERT INTO public.nht_user_roles (user_id, role)
-- VALUES ('USER-UUID-HERE', 'admin')
-- ON CONFLICT (user_id) DO UPDATE SET role = excluded.role;
--
-- INSERT INTO public.nht_user_roles (user_id, role)
-- VALUES ('USER-UUID-HERE', 'uploader')
-- ON CONFLICT (user_id) DO UPDATE SET role = excluded.role;
--
-- INSERT INTO public.nht_user_roles (user_id, role)
-- VALUES ('USER-UUID-HERE', 'viewer')
-- ON CONFLICT (user_id) DO UPDATE SET role = excluded.role;
