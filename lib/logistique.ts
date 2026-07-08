// lib/logistique.ts — Synthèse logistique annuelle : besoins par référence et par mois.
//
// Pour chaque campagne et chaque référence, le besoin TOTAL = somme sur les paliers de
// (qté/pack de la réf × nb packs du palier). Ce besoin est ensuite réparti sur les mois
// de livraison selon le profil :
//   - Mois -1 (le mois précédant le début de l'offre) : 40% du total.
//   - Reste (60%) lissé à parts égales du mois de début jusqu'à 1 mois avant la fin.
// Les besoins de toutes les campagnes sont agrégés par référence et par mois calendaire
// (janvier → décembre).

import type ExcelJS from "exceljs";
import type { CampagneCreee } from "@/lib/create-campaign";
import { qtyParPack, totalPacks, ventilationPalier } from "@/lib/create-campaign";

export const MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

/** Écrit un onglet "Synthèse logistique" (réf × mois) dans le classeur. Partagé entre les
 *  exports simple et multi. nameByRef : libellés de secours par réf. */
export function writeSyntheseLogistiqueSheet(wb: ExcelJS.Workbook, log: SyntheseLogistique, nameByRef: Record<string, string> = {}) {
  const TEAL = "0D9488", DARK = "1A1A2E", WHITE = "FFFFFF";
  const mois = log.moisLabels && log.moisLabels.length ? log.moisLabels : MOIS_FR;
  const nbCols = 2 + mois.length + 1;               // Réf + Produit + N mois + Total
  const totalCol = nbCols;                          // index colonne "Total"
  // Le template peut déjà contenir un onglet "Synthèse logistique" → on le retire d'abord
  // pour éviter l'erreur "Worksheet name already exists", puis on le (re)crée proprement.
  const existant = wb.getWorksheet("Synthèse logistique");
  if (existant) wb.removeWorksheet(existant.id);
  const sw = wb.addWorksheet("Synthèse logistique", { views: [{ showGridLines: false }] });
  sw.columns = [{ width: 16 }, { width: 42 }, ...mois.map(() => ({ width: 12 })), { width: 12 }];

  const titleRow = sw.addRow(["Synthèse besoins logistiques — par référence et par mois"]);
  sw.mergeCells(titleRow.number, 1, titleRow.number, nbCols);
  const tc = sw.getCell(titleRow.number, 1);
  tc.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
  tc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
  tc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  titleRow.height = 24;
  sw.addRow([]);

  const head = sw.addRow(["Réf", "Produit", ...mois, "Total"]);
  head.height = 20;
  head.eachCell(c => {
    c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
    c.alignment = { horizontal: "center", vertical: "middle" };
  });

  const numFmt = "#,##0";
  for (const l of log.lignes) {
    const libelle = l.name || nameByRef[l.ref] || "";
    const row = sw.addRow([l.ref, libelle, ...l.parMois, l.total]);
    row.eachCell((c, col) => {
      c.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK }, bold: col === totalCol };
      c.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
      if (col >= 3) { c.numFmt = numFmt; c.alignment = { horizontal: "right" }; }
      if (col === 1) c.font = { ...c.font, name: "Consolas" };
    });
  }
  const totRow = sw.addRow(["", "TOTAL", ...log.totalParMois, log.totalGeneral]);
  totRow.eachCell((c, col) => {
    c.font = { bold: true, size: 10, name: "Calibri", color: { argb: "FF" + WHITE } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
    if (col >= 3) { c.numFmt = numFmt; c.alignment = { horizontal: "right" }; }
  });
  sw.addRow([]);
  const note = sw.addRow(["Profil de livraison : 40 % le mois précédant le début de l'offre, puis 60 % lissé à parts égales jusqu'à 1 mois avant la fin. Les mois s'étendent sur l'année suivante si une offre déborde."]);
  sw.mergeCells(note.number, 1, note.number, nbCols);
  sw.getCell(note.number, 1).font = { italic: true, size: 9, color: { argb: "FF6B7280" }, name: "Calibri" };
}

export interface LigneLogistique {
  ref: string;
  name: string;
  parMois: number[];   // 12 valeurs (index 0 = janvier … 11 = décembre)
  total: number;
}

export interface SyntheseLogistique {
  lignes: LigneLogistique[];
  totalParMois: number[];   // aligné sur moisLabels
  totalGeneral: number;
  moisLabels: string[];     // libellés des mois (peut déborder sur N+1 : "Janvier 2027"…)
}

const PART_VAGUE = 0.4; // 40% le mois -1

// Index de mois ABSOLU d'une date "YYYY-MM-DD" = année*12 + (mois-1). Permet de gérer un axe
// temporel qui déborde d'une année sur l'autre. Renvoie null si invalide.
function absMonth(date: string): number | null {
  if (!date) return null;
  const m = date.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (!m) return null;
  return parseInt(m[1], 10) * 12 + (parseInt(m[2], 10) - 1);
}

// Libellé d'un mois absolu : "Janvier 2027".
export function libelleMoisAbsolu(abs: number): string {
  const y = Math.floor(abs / 12), mi = abs % 12;
  return `${MOIS_FR[mi]} ${y}`;
}

/**
 * Répartit un besoin total selon le profil 40% (mois -1) + 60% lissé, sur des mois ABSOLUS.
 * @returns Map<moisAbsolu, quantité>
 * @param absDebut mois absolu du début de l'offre
 * @param absFin   mois absolu de la fin de l'offre
 * Profil : mois -1 (avant début) = 40% ; reste lissé du mois de début jusqu'à 1 mois avant la fin.
 */
export function repartirAbsolu(total: number, absDebut: number, absFin: number): Map<number, number> {
  const out = new Map<number, number>();
  if (total <= 0) return out;
  const add = (m: number, q: number) => out.set(m, (out.get(m) || 0) + q);

  const moisVague = absDebut - 1;                 // mois -1 (peut être l'année précédente)
  const vague = Math.round(total * PART_VAGUE);
  add(moisVague, vague);

  const reste = total - vague;
  const lissDebut = absDebut, lissFin = absFin - 1; // jusqu'à 1 mois avant la fin
  const mois: number[] = [];
  for (let m = lissDebut; m <= lissFin; m++) mois.push(m);
  if (mois.length > 0) {
    const part = Math.floor(reste / mois.length);
    let residu = reste - part * mois.length;
    for (const m of mois) { add(m, part); if (residu > 0) { add(m, 1); residu -= 1; } }
  } else {
    add(absDebut, reste); // offre très courte → tout au mois de début
  }
  return out;
}

/** Besoin total par référence d'une campagne = somme(qté/pack × nb packs) sur les paliers. */
function besoinsCampagne(camp: CampagneCreee): Record<string, number> {
  const out: Record<string, number> = {};
  for (const art of camp.articles) {
    const ref = art.ref.trim();
    if (!ref) continue;
    let total = 0;
    const totalP = totalPacks(camp.paliers);
    for (const pal of camp.paliers) {
      const vent = ventilationPalier(camp.articles, pal);
      total += qtyParPack(art, pal, totalP, vent, camp.articles) * (pal.nbPacks || 0);
    }
    out[ref] = (out[ref] || 0) + total;
  }
  return out;
}

/** Agrège les besoins logistiques de plusieurs campagnes par réf et par mois ABSOLU.
 *  L'axe des mois s'étend automatiquement (déborde sur N+1 si une offre finit l'année suivante). */
export function buildSyntheseLogistique(campagnes: CampagneCreee[]): SyntheseLogistique {
  // accum[ref] = Map<moisAbsolu, qté>
  const accum: Record<string, { name: string; parMoisAbs: Map<number, number> }> = {};
  let minAbs = Infinity, maxAbs = -Infinity;

  for (const camp of campagnes) {
    const ad = absMonth(camp.dateDebut);
    const af = absMonth(camp.dateFin);
    if (ad == null || af == null) continue;
    const besoins = besoinsCampagne(camp);
    const nameByRef: Record<string, string> = {};
    for (const a of camp.articles) if (a.ref.trim()) nameByRef[a.ref.trim()] = a.name || "";

    for (const [ref, total] of Object.entries(besoins)) {
      if (!accum[ref]) accum[ref] = { name: nameByRef[ref] || "", parMoisAbs: new Map() };
      if (!accum[ref].name && nameByRef[ref]) accum[ref].name = nameByRef[ref];
      const rep = repartirAbsolu(total, ad, af);
      for (const [m, q] of rep) {
        accum[ref].parMoisAbs.set(m, (accum[ref].parMoisAbs.get(m) || 0) + q);
        if (m < minAbs) minAbs = m;
        if (m > maxAbs) maxAbs = m;
      }
    }
  }

  // Aucune donnée → synthèse vide.
  if (!isFinite(minAbs)) return { lignes: [], totalParMois: [], totalGeneral: 0, moisLabels: [] };

  // Axe des mois : du 1er au dernier mois de livraison (continu).
  const nbMois = maxAbs - minAbs + 1;
  const moisLabels: string[] = [];
  for (let i = 0; i < nbMois; i++) moisLabels.push(libelleMoisAbsolu(minAbs + i));

  const lignes: LigneLogistique[] = Object.entries(accum).map(([ref, v]) => {
    const parMois = new Array(nbMois).fill(0);
    for (const [m, q] of v.parMoisAbs) parMois[m - minAbs] += q;
    return { ref, name: v.name, parMois, total: parMois.reduce((s, x) => s + x, 0) };
  }).sort((a, b) => b.total - a.total);

  const totalParMois = new Array(nbMois).fill(0);
  for (const l of lignes) for (let i = 0; i < nbMois; i++) totalParMois[i] += l.parMois[i];
  const totalGeneral = totalParMois.reduce((s, x) => s + x, 0);

  return { lignes, totalParMois, totalGeneral, moisLabels };
}
