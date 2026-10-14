// lib/server/journal.ts — Journal des modifications (écrit par la passerelle /api/db).
// Pour chaque écriture réussie : utilisateur Odoo, table, action, objet, version d'avant
// (relue juste avant l'écriture) et d'après. Les photos (base64) ne sont pas journalisées.
// Best effort : un échec de journalisation n'empêche jamais l'écriture métier.

import { SUPABASE_REST, adminHeaders } from "@/lib/server/supabase-admin";
import type { UtilisateurOdoo } from "@/lib/server/odoo-auth";

// Clé primaire par table (pour relire l'« avant » et identifier l'objet).
const CLE: Record<string, string> = {
  campagnes_creees: "id", campagnes: "id", analyse_offres: "id", projets_kits: "id",
  planning_produits: "ref", planning_quantites: "id", planning_historique: "id", parametres_app: "cle",
};
// Historique du planning = déjà un journal métier : inutile de le dupliquer.
const NON_JOURNALISEES = new Set(["planning_historique"]);
const MAX_LIGNES = 200;

const allege = (row: any) => {
  if (!row || typeof row !== "object") return row;
  const { photo, ...reste } = row;
  return photo ? { ...reste, photo: "(image non journalisée)" } : reste;
};
const nomDe = (row: any) => row?.nom || row?.label || row?.material_text || row?.ref || row?.cle || row?.code || null;

/** Relit les lignes visées AVANT l'écriture (filtres de la requête, ou ids du corps). */
export async function lireAvant(table: string, method: string, qs: URLSearchParams, corps: string): Promise<any[] | null> {
  if (NON_JOURNALISEES.has(table)) return null;
  const cle = CLE[table];
  if (!cle) return null;
  try {
    let filtre = "";
    if (method === "PATCH" || method === "DELETE") {
      const f = new URLSearchParams();
      for (const [k, v] of qs) if (!["select", "order", "limit", "offset", "on_conflict", "columns"].includes(k)) f.append(k, v);
      filtre = f.toString();
    } else if (method === "POST") {
      const lignes = [].concat(JSON.parse(corps || "null") || []) as any[];
      const ids = lignes.map(l => l?.[cle]).filter(v => v != null && v !== "");
      if (!ids.length) return [];
      filtre = `${cle}=in.(${ids.map(v => `"${String(v).replace(/"/g, "")}"`).join(",")})`;
    }
    if (!filtre) return null;
    const res = await fetch(`${SUPABASE_REST}/${table}?select=*&${filtre}&limit=${MAX_LIGNES}`, { headers: adminHeaders(), cache: "no-store" });
    return res.ok ? await res.json() : null;
  } catch { return null; }
}

/** Enregistre l'écriture dans le journal. */
export async function journaliser(user: UtilisateurOdoo, table: string, method: string, corps: string, avant: any[] | null) {
  if (NON_JOURNALISEES.has(table)) return;
  try {
    const cle = CLE[table] || "id";
    let apres: any[] = [];
    if (method === "POST" || method === "PATCH") { try { apres = [].concat(JSON.parse(corps || "null") || []); } catch { /* corps illisible */ } }
    const avantParId = new Map((avant || []).map(r => [String(r?.[cle]), r]));
    const lignes: any[] = [];
    const base = { uid: user.uid, utilisateur: user.nom, login: user.login, table_cible: table };
    if (method === "DELETE") {
      for (const r of (avant || []).slice(0, MAX_LIGNES)) lignes.push({ ...base, action: "suppression", objet_id: String(r?.[cle] ?? ""), objet_nom: nomDe(r), avant: allege(r), apres: null });
    } else if (method === "PATCH") {
      const patch = apres[0] || {};
      const action = "deleted_at" in patch ? (patch.deleted_at ? "corbeille" : "restauration") : "modification";
      for (const r of (avant || []).slice(0, MAX_LIGNES)) lignes.push({ ...base, action, objet_id: String(r?.[cle] ?? ""), objet_nom: nomDe(r), avant: allege(r), apres: allege({ ...r, ...patch }) });
    } else {
      for (const r of apres.slice(0, MAX_LIGNES)) {
        const prec = avantParId.get(String(r?.[cle]));
        lignes.push({ ...base, action: prec ? "modification" : "création", objet_id: String(r?.[cle] ?? ""), objet_nom: nomDe(r) || nomDe(prec), avant: prec ? allege(prec) : null, apres: allege(prec ? { ...prec, ...r } : r) });
      }
    }
    if (!lignes.length) return;
    await fetch(`${SUPABASE_REST}/journal_modifications`, {
      method: "POST", headers: { ...adminHeaders(), "Content-Type": "application/json", Prefer: "return=minimal" },
      body: JSON.stringify(lignes), cache: "no-store",
    });
  } catch (e: any) { console.error("[journal]", e?.message); }
}
