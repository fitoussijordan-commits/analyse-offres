-- Paramètres partagés par toute l'équipe (ex. clients exclus de l'analyse des campagnes).
-- Accès uniquement via la passerelle serveur /api/db (RLS sans policy, droits publics retirés).
create table if not exists public.parametres_app (
  cle text primary key,
  valeur jsonb not null,
  updated_at timestamptz not null default now()
);
alter table public.parametres_app enable row level security;
revoke all on table public.parametres_app from anon, authenticated;
insert into public.parametres_app (cle, valeur) values ('clients_exclus', '["Réservation"]'::jsonb) on conflict (cle) do nothing;
