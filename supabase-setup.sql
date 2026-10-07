create table if not exists public.herd_members (
  user_id uuid primary key references auth.users(id) on delete cascade,
  role text not null check (role in ('owner', 'viewer')),
  created_at timestamptz not null default now()
);

create table if not exists public.herd_app_state (
  id smallint primary key check (id = 1),
  data jsonb not null default '{"lots": [], "weighings": [], "feedings": []}'::jsonb,
  updated_at timestamptz not null default now()
);

alter table public.herd_members enable row level security;
alter table public.herd_app_state enable row level security;

create or replace function public.is_herd_member()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1
    from public.herd_members
    where user_id = (select auth.uid())
  );
$$;

create or replace function public.is_herd_owner()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1
    from public.herd_members
    where user_id = (select auth.uid())
      and role = 'owner'
  );
$$;

revoke all on function public.is_herd_member() from public;
revoke all on function public.is_herd_owner() from public;
grant execute on function public.is_herd_member() to authenticated;
grant execute on function public.is_herd_owner() to authenticated;

grant select on public.herd_members to authenticated;
grant select, insert, update on public.herd_app_state to authenticated;

drop policy if exists "Members can read their own role" on public.herd_members;
create policy "Members can read their own role"
on public.herd_members for select to authenticated
using ((select auth.uid()) = user_id);

drop policy if exists "Members can read shared herd data" on public.herd_app_state;
create policy "Members can read shared herd data"
on public.herd_app_state for select to authenticated
using ((select public.is_herd_member()));

drop policy if exists "Owners can create shared herd data" on public.herd_app_state;
create policy "Owners can create shared herd data"
on public.herd_app_state for insert to authenticated
with check ((select public.is_herd_owner()));

drop policy if exists "Owners can update shared herd data" on public.herd_app_state;
create policy "Owners can update shared herd data"
on public.herd_app_state for update to authenticated
using ((select public.is_herd_owner()))
with check ((select public.is_herd_owner()));