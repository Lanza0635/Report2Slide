-- Excel Araçları SaaS — Supabase profiles şeması
-- Supabase SQL Editor'de bir kez çalıştırın.

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  email text unique,
  plan text not null default 'free' check (plan in ('free', 'pro')),
  subscription_status text,
  lemon_customer_id text,
  lemon_subscription_id text,
  updated_at timestamptz default now(),
  created_at timestamptz default now()
);

create index if not exists profiles_email_idx on public.profiles (email);
create index if not exists profiles_lemon_customer_id_idx on public.profiles (lemon_customer_id);

alter table public.profiles enable row level security;

-- Kullanıcı kendi profilini okuyabilir
drop policy if exists "Users can read own profile" on public.profiles;
create policy "Users can read own profile"
  on public.profiles for select
  using (auth.uid() = id);

-- Yazma: service role (Flask webhook) RLS'yi bypass eder.
-- İsteğe bağlı: kullanıcı kendi satırını güncellemesin (plan client'tan değiştirilemez).

-- Yeni auth kullanıcısında otomatik free profil
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email, plan)
  values (new.id, lower(new.email), 'free')
  on conflict (id) do update set email = excluded.email;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();
