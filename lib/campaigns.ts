// lib/campaigns.ts — Gestion des campagnes (Supabase) + chargement des offres
import { supabase } from "@/lib/supabase";

export interface Offre {
  id: string;
  code: string;
  label: string;
  produits: string[];
  codeInterne?: string;
}

export interface Campagne {
  id: string;
  nom: string;
  offres: string[];   // codes d'offres (référencent analyse_offres.code)
  produits: string[]; // références produits autonomes
  notes: string[];    // notes internes (x_note_interne)
  createdAt?: string;
}

export function genId() {
  return Math.random().toString(36).slice(2, 10);
}

// ─── Offres (réutilise la table analyse_offres) ───────────────────────────────
function rowToOffre(row: any): Offre {
  return {
    id: row.id,
    code: row.code,
    label: row.label || "",
    produits: row.produits || [],
    codeInterne: row.code_interne || undefined,
  };
}

function offreToRow(o: Offre) {
  return { id: o.id, code: o.code, label: o.label || "", produits: o.produits, code_interne: o.codeInterne || null };
}

export async function loadOffres(): Promise<Offre[]> {
  const { data, error } = await supabase.from("analyse_offres").select("*").order("created_at");
  if (error) throw new Error(error.message);
  return (data || []).map(rowToOffre);
}

export async function upsertOffre(o: Offre): Promise<void> {
  const { error } = await supabase.from("analyse_offres").upsert(offreToRow(o));
  if (error) throw new Error(error.message);
}

export async function deleteOffre(id: string): Promise<void> {
  const { error } = await supabase.from("analyse_offres").delete().eq("id", id);
  if (error) throw new Error(error.message);
}

// ─── Campagnes ────────────────────────────────────────────────────────────────
function rowToCampagne(row: any): Campagne {
  return {
    id: row.id,
    nom: row.nom,
    offres: row.offres || [],
    produits: row.produits || [],
    notes: row.notes || [],
    createdAt: row.created_at,
  };
}
function campagneToRow(c: Campagne) {
  return {
    id: c.id,
    nom: c.nom,
    offres: c.offres,
    produits: c.produits,
    notes: c.notes,
  };
}

export async function loadCampagnes(): Promise<Campagne[]> {
  const { data, error } = await supabase.from("campagnes").select("*").order("created_at");
  if (error) throw new Error(error.message);
  return (data || []).map(rowToCampagne);
}

export async function upsertCampagne(c: Campagne): Promise<void> {
  const { error } = await supabase.from("campagnes").upsert(campagneToRow(c));
  if (error) throw new Error(error.message);
}

export async function deleteCampagne(id: string): Promise<void> {
  const { error } = await supabase.from("campagnes").delete().eq("id", id);
  if (error) throw new Error(error.message);
}

// ─── Campagnes créées "à blanc" (table campagnes_creees) ──────────────────────
// La structure (paliers + articles) est riche → on la sérialise en JSON dans une colonne
// `data` pour éviter un schéma rigide. Colonnes : id, nom, date_debut, date_fin,
// periode_debut, periode_fin, data (jsonb), created_at.
import type { CampagneCreee } from "@/lib/create-campaign";

function rowToCampagneCreee(row: any): CampagneCreee {
  const d = (typeof row.data === "string" ? JSON.parse(row.data) : row.data) || {};
  return {
    id: row.id,
    nom: row.nom,
    dateDebut: row.date_debut || "",
    dateFin: row.date_fin || "",
    periodeDebut: row.periode_debut || "",
    periodeFin: row.periode_fin || "",
    paliers: d.paliers || [],
    createdAt: row.created_at,
  };
}

function campagneCreeeToRow(c: CampagneCreee) {
  return {
    id: c.id,
    nom: c.nom,
    date_debut: c.dateDebut || null,
    date_fin: c.dateFin || null,
    periode_debut: c.periodeDebut || null,
    periode_fin: c.periodeFin || null,
    data: { paliers: c.paliers },
  };
}

export async function loadCampagnesCreees(): Promise<CampagneCreee[]> {
  const { data, error } = await supabase.from("campagnes_creees").select("*").order("created_at");
  if (error) throw new Error(error.message);
  return (data || []).map(rowToCampagneCreee);
}

export async function upsertCampagneCreee(c: CampagneCreee): Promise<void> {
  const { error } = await supabase.from("campagnes_creees").upsert(campagneCreeeToRow(c));
  if (error) throw new Error(error.message);
}

export async function deleteCampagneCreee(id: string): Promise<void> {
  const { error } = await supabase.from("campagnes_creees").delete().eq("id", id);
  if (error) throw new Error(error.message);
}
