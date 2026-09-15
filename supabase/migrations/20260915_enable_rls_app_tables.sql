-- Active le RLS sur les tables de l'app Analyse Offres et retire tout accès aux rôles publics.
-- À appliquer UNIQUEMENT après le déploiement de la passerelle /api/db (clé secrète côté
-- serveur) : sans elle, l'app ne peut plus lire ni écrire ces tables.
-- Aucune policy n'est créée volontairement : anon/authenticated n'ont plus aucun accès,
-- seule la clé secrète (qui contourne le RLS) passe, depuis le serveur.

do $$
declare t text;
begin
  foreach t in array array[
    'campagnes', 'campagnes_creees', 'analyse_offres', 'projets_kits',
    'planning_produits', 'planning_quantites', 'planning_historique'
  ] loop
    execute format('alter table public.%I enable row level security', t);
    execute format('revoke all on table public.%I from anon, authenticated', t);
  end loop;
end $$;
