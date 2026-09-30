-- Période d'analyse d'une campagne : borne les ventes des produits autonomes (réfs vendues
-- hors offre MEA). Sans elle, toutes les ventes de ces réfs, tous temps confondus, étaient
-- comptées. Colonnes nullables : les campagnes existantes restent inchangées.
alter table public.campagnes add column if not exists date_debut date;
alter table public.campagnes add column if not exists date_fin date;
