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
import { qtyParPack } from "@/lib/create-campaign";

export const MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

/** Écrit un onglet "Synthèse logistique" (réf × mois) dans le classeur. Partagé entre les
 *  exports simple et multi. nameByRef : libellés de secours par réf. */
export function writeSyntheseLogistiqueSheet(wb: ExcelJS.Workbook, log: SyntheseLogistique, nameByRef: Record<string, string> = {}) {
  const TEAL = "0D9488", DARK = "1A1A2E", WHITE = "FFFFFF";
  const sw = wb.addWorksheet("Synthèse logistique", { views: [{ showGridLines: false }] });
  sw.columns = [{ width: 16 }, { width: 42 }, ...MOIS_FR.map(() => ({ width: 10 })), { width: 12 }];

  const titleRow = sw.addRow(["Synthèse besoins logistiques — par référence et par mois"]);
  sw.mergeCells(titleRow.number, 1, titleRow.number, 15);
  const tc = sw.getCell(titleRow.number, 1);
  tc.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
  tc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
  tc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  titleRow.height = 24;
  sw.addRow([]);

  const head = sw.addRow(["Réf", "Produit", ...MOIS_FR, "Total"]);
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
      c.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK }, bold: col === 15 };
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
  const note = sw.addRow(["Profil de livraison : 40 % le mois précédant le début de l'offre, puis 60 % lissé à parts égales jusqu'à 1 mois avant la fin."]);
  sw.mergeCells(note.number, 1, note.number, 15);
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
  totalParMois: number[]; // 12 valeurs
  totalGeneral: number;
}

const PART_VAGUE = 0.4; // 40% le mois -1

// Index de mois calendaire (0-11) d'une date "YYYY-MM-DD". Renvoie null si invalide.
function moisIndex(date: string): number | null {
  if (!date) return null;
  const m = date.match(/^\d{4}-(\d{2})-\d{2}/);
  if (!m) return null;
  const mi = parseInt(m[1], 10) - 1;
  return mi >= 0 && mi <= 11 ? mi : null;
}

/**
 * Répartit un besoin total sur 12 mois selon le profil 40% (mois -1) + 60% lissé.
 * @param total       besoin total de la réf
 * @param moisDebut   index 0-11 du mois de début de l'offre
 * @param moisFin     index 0-11 du mois de fin de l'offre
 * Note : on raisonne en mois calendaires sur une seule année (jan-déc). Si l'offre déborde
 * sur l'année suivante, on borne au 31 déc.
 */
export function repartir(total: number, moisDebut: number, moisFin: number): number[] {
  const out = new Array(12).fill(0);
  if (total <= 0) return out;

  // Mois -1 = mois avant le début. S'il sort de l'année (offre démarrant en janvier →
  // mois -1 = décembre N-1), on le place en colonne Décembre (index 11) : la prod est
  // anticipée le décembre précédent.
  const moisVague = moisDebut - 1 < 0 ? 11 : moisDebut - 1;
  // Fenêtre de lissage : du mois de début jusqu'à 1 mois avant la fin (inclus).
  const lissDebut = moisDebut;
  const lissFin = moisFin - 1;

  // 40% en 1re vague.
  const vague = Math.round(total * PART_VAGUE);
  out[moisVague] += vague;

  // 60% lissé à parts égales sur la fenêtre [lissDebut .. lissFin].
  const reste = total - vague;
  const moisLissage: number[] = [];
  for (let m = lissDebut; m <= lissFin; m++) if (m >= 0 && m <= 11) moisLissage.push(m);
  if (moisLissage.length > 0) {
    const part = Math.floor(reste / moisLissage.length);
    let residu = reste - part * moisLissage.length; // arrondi → on met le reliquat sur le 1er mois
    for (const m of moisLissage) { out[m] += part; if (residu > 0) { out[m] += 1; residu -= 1; } }
  } else if (moisDebut >= 0 && moisDebut <= 11) {
    // Pas de fenêtre de lissage (offre très courte) → tout le reste au mois de début.
    out[moisDebut] += reste;
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
    for (const pal of camp.paliers) total += qtyParPack(art, pal) * (pal.nbPacks || 0);
    out[ref] = (out[ref] || 0) + total;
  }
  return out;
}

/** Agrège les besoins logistiques de plusieurs campagnes par réf et par mois. */
export function buildSyntheseLogistique(campagnes: CampagneCreee[]): SyntheseLogistique {
  const accum: Record<string, { name: string; parMois: number[] }> = {};

  for (const camp of campagnes) {
    const md = moisIndex(camp.dateDebut);
    const mf = moisIndex(camp.dateFin);
    if (md == null || mf == null) continue; // sans dates, pas de planning possible
    const besoins = besoinsCampagne(camp);
    const nameByRef: Record<string, string> = {};
    for (const a of camp.articles) if (a.ref.trim()) nameByRef[a.ref.trim()] = a.name || "";

    for (const [ref, total] of Object.entries(besoins)) {
      if (!accum[ref]) accum[ref] = { name: nameByRef[ref] || "", parMois: new Array(12).fill(0) };
      if (!accum[ref].name && nameByRef[ref]) accum[ref].name = nameByRef[ref];
      const rep = repartir(total, md, mf);
      for (let m = 0; m < 12; m++) accum[ref].parMois[m] += rep[m];
    }
  }

  const lignes: LigneLogistique[] = Object.entries(accum)
    .map(([ref, v]) => ({ ref, name: v.name, parMois: v.parMois, total: v.parMois.reduce((s, x) => s + x, 0) }))
    .sort((a, b) => b.total - a.total);

  const totalParMois = new Array(12).fill(0);
  for (const l of lignes) for (let m = 0; m < 12; m++) totalParMois[m] += l.parMois[m];
  const totalGeneral = totalParMois.reduce((s, x) => s + x, 0);

  return { lignes, totalParMois, totalGeneral };
}
