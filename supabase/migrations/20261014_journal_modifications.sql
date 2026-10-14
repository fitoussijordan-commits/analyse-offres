-- Journal des modifications : qui a créé / modifié / supprimé quoi, avec la version d'avant
-- et d'après. Écrit par la passerelle serveur /api/db, lu par /api/journal (admins seuls).
-- PAS dans la liste blanche de /api/db : aucun utilisateur ne peut le lire ni le modifier.
create table if not exists public.journal_modifications (
  id bigint generated always as identity primary key,
  cree_le timestamptz not null default now(),
  uid integer, utilisateur text, login text,
  table_cible text not null, action text not null,
  objet_id text, objet_nom text,
  avant jsonb, apres jsonb
);
create index if not exists journal_modifications_cree_le on public.journal_modifications (cree_le desc);
create index if not exists journal_modifications_login on public.journal_modifications (login);
alter table public.journal_modifications enable row level security;
revoke all on table public.journal_modifications from anon, authenticated;
