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
import { qtyParPack, totalPacks, ventilationPalier, articlesPalier, qtyPalierOpca } from "@/lib/create-campaign";
import { estOpca } from "@/lib/type-produit";

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
  if (log.ignorees && log.ignorees.length) {
    const warn = sw.addRow([`⚠ Campagne(s) NON comptée(s), dates de début/fin manquantes ou invalides : ${log.ignorees.join(", ")}`]);
    sw.mergeCells(warn.number, 1, warn.number, nbCols);
    const wc = sw.getCell(warn.number, 1);
    wc.font = { bold: true, size: 10, color: { argb: "FFB91C1C" }, name: "Calibri" };
    wc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } };
    wc.alignment = { wrapText: true, vertical: "middle", indent: 1 };
    warn.height = 30;
  }
  sw.addRow([]);

  const head = sw.addRow(["Réf", "Produit", ...mois, "Total"]);
  head.height = 20;
  head.eachCell(c => {
    c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
    c.alignment = { horizontal: "center", vertical: "middle" };
  });

  const numFmt = "#,##0";
  // Détail regroupé par réf, pour les notes de survol.
  const detailParRef = new Map<string, LigneDetailLogistique[]>();
  for (const d of log.detail || []) {
    if (!detailParRef.has(d.ref)) detailParRef.set(d.ref, []);
    detailParRef.get(d.ref)!.push(d);
  }
  const fmt = (n: number) => Math.round(n).toLocaleString("fr-FR");
  // Texte de note : une ligne par offre contributrice (campagne / offre : qté), triée.
  const noteOffres = (parts: { d: LigneDetailLogistique; q: number }[]) => parts
    .filter(p => p.q > 0).sort((a, b) => b.q - a.q)
    .map(p => `${p.d.campagne} / ${p.d.palier} : ${fmt(p.q)}`).join("\n");

  for (const l of log.lignes) {
    const libelle = l.name || nameByRef[l.ref] || "";
    const row = sw.addRow([l.ref, libelle, ...l.parMois, l.total]);
    row.eachCell((c, col) => {
      c.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK }, bold: col === totalCol };
      c.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
      if (col >= 3) { c.numFmt = numFmt; c.alignment = { horizontal: "right" }; }
      if (col === 1) c.font = { ...c.font, name: "Consolas" };
    });
    const parts = detailParRef.get(l.ref) || [];
    if (parts.length) {
      l.parMois.forEach((q, i) => {
        if (q > 0) { const t = noteOffres(parts.map(d => ({ d, q: d.parMois[i] || 0 }))); if (t) row.getCell(3 + i).note = t; }
      });
      const tTot = noteOffres(parts.map(d => ({ d, q: d.total })));
      if (tTot) row.getCell(totalCol).note = tTot;
    }
  }
  // Aucun besoin calculable (dates de campagne manquantes) : on l'indique clairement
  // plutôt que de laisser un tableau vide qu'on pourrait confondre avec un bug.
  if (!log.lignes.length) {
    const empty = sw.addRow(["", "Aucun besoin calculé — renseigne les dates de début et de fin de campagne."]);
    sw.mergeCells(empty.number, 2, empty.number, nbCols);
    sw.getCell(empty.number, 2).font = { italic: true, size: 10, color: { argb: "FF6B7280" }, name: "Calibri" };
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
  if ((log.detail || []).length) {
    const aide = sw.addRow(["Survole une case pour voir les offres qui la composent. Le détail complet (filtrable) est dans l'onglet « Détail logistique »."]);
    sw.mergeCells(aide.number, 1, aide.number, nbCols);
    sw.getCell(aide.number, 1).font = { italic: true, size: 9, color: { argb: "FF6B7280" }, name: "Calibri" };
  }

  writeDetailLogistiqueSheet(wb, log, nameByRef);
}

/** Onglet « Détail logistique » : une ligne par référence × campagne × offre, mêmes colonnes
 *  de mois que la synthèse. Trié par référence (plus gros besoin d'abord) puis par quantité,
 *  avec filtres automatiques pour isoler une campagne ou une offre. */
function writeDetailLogistiqueSheet(wb: ExcelJS.Workbook, log: SyntheseLogistique, nameByRef: Record<string, string>) {
  const NOM = "Détail logistique";
  const existant = wb.getWorksheet(NOM);
  if (existant) wb.removeWorksheet(existant.id);
  const detail = log.detail || [];
  if (!detail.length) return;

  const TEAL = "0D9488", DARK = "1A1A2E", WHITE = "FFFFFF";
  const mois = log.moisLabels && log.moisLabels.length ? log.moisLabels : MOIS_FR;
  const nbCols = 4 + mois.length + 1;
  const ws = wb.addWorksheet(NOM, { views: [{ state: "frozen", xSplit: 4, ySplit: 3, showGridLines: false }] });
  ws.columns = [{ width: 14 }, { width: 38 }, { width: 30 }, { width: 28 }, ...mois.map(() => ({ width: 11 })), { width: 11 }];

  const titre = ws.addRow(["Détail des besoins logistiques — par référence, campagne et offre"]);
  ws.mergeCells(titre.number, 1, titre.number, nbCols);
  const tc = ws.getCell(titre.number, 1);
  tc.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
  tc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
  tc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  titre.height = 24;
  ws.addRow([]);

  const head = ws.addRow(["Réf", "Produit", "Campagne", "Offre", ...mois, "Total"]);
  head.height = 20;
  head.eachCell(c => {
    c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
    c.alignment = { horizontal: "center", vertical: "middle" };
  });

  // Ordre : références dans l'ordre de la synthèse (plus gros besoin d'abord), puis offres.
  const rang = new Map(log.lignes.map((l, i) => [l.ref, i]));
  const tri = [...detail].sort((a, b) => ((rang.get(a.ref) ?? 1e9) - (rang.get(b.ref) ?? 1e9)) || (b.total - a.total));
  let refPrec = "", bande = false;
  for (const d of tri) {
    if (d.ref !== refPrec) { bande = !bande; refPrec = d.ref; }
    const row = ws.addRow([d.ref, d.name || nameByRef[d.ref] || "", d.campagne, d.palier, ...d.parMois, d.total]);
    row.eachCell({ includeEmpty: true }, (c, col) => {
      c.font = { size: 10, name: col === 1 ? "Consolas" : "Calibri", color: { argb: "FF" + DARK }, bold: col === nbCols };
      if (bande) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F7FA" } };
      c.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
      if (col >= 5) { c.numFmt = "#,##0;-#,##0;\"\""; c.alignment = { horizontal: "right" }; }
    });
    ws.getCell(row.number, 1).numFmt = "@";
  }
  ws.autoFilter = { from: { row: head.number, column: 1 }, to: { row: head.number, column: nbCols } };
}

export interface LigneLogistique {
  ref: string;
  name: string;
  parMois: number[];   // 12 valeurs (index 0 = janvier … 11 = décembre)
  total: number;
}

/** Détail d'un besoin : une ligne par référence × campagne × offre (palier). */
export interface LigneDetailLogistique {
  ref: string;
  name: string;
  campagne: string;
  palier: string;
  parMois: number[];
  total: number;
}

export interface SyntheseLogistique {
  lignes: LigneLogistique[];
  // Même chose éclatée par campagne et par offre : explique chaque case de la synthèse.
  detail?: LigneDetailLogistique[];
  totalParMois: number[];   // aligné sur moisLabels
  totalGeneral: number;
  moisLabels: string[];     // libellés des mois (peut déborder sur N+1 : "Janvier 2027"…)
  // Campagnes exclues faute de dates exploitables (manquantes, invalides ou fin < début).
  // Remontées à l'utilisateur : sans ça, leurs besoins disparaissaient en silence.
  ignorees?: string[];
}

const PART_VAGUE = 0.4; // 40% le mois -1

// Index de mois ABSOLU d'une date "YYYY-MM-DD" = année*12 + (mois-1). Permet de gérer un axe
// temporel qui déborde d'une année sur l'autre. Renvoie null si invalide.
function absMonth(date: string): number | null {
  if (!date) return null;
  const m = date.match(/^(\d{4})-(\d{2})-\d{2}$/);
  if (!m) return null;
  const y = parseInt(m[1], 10), mo = parseInt(m[2], 10);
  if (y < 2000 || y > 2100 || mo < 1 || mo > 12) return null;
  return y * 12 + (mo - 1);
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

/** Besoin d'une campagne détaillé par OFFRE : { ref, libellé du palier, quantité totale }.
 *  Le total par référence (somme sur les paliers) reste identique à avant. */
function besoinsParOffre(camp: CampagneCreee): { ref: string; palier: string; qte: number }[] {
  const out: { ref: string; palier: string; qte: number }[] = [];
  const totalP = totalPacks(camp.paliers);
  for (const pal of camp.paliers) {
    const vent = ventilationPalier(camp.articles, pal);
    const libelle = [pal.label, pal.code].map(x => (x || "").trim()).filter(Boolean).join(" — ") || "(palier sans nom)";
    // Palier OPCA : seuls les gratuits offerts sont à approvisionner (le panachage de
    // l'institut n'est pas une composition figée, la ligne OPCA n'est pas un produit).
    for (const art of articlesPalier(camp.articles, pal)) {
      const ref = art.ref.trim();
      if (!ref || estOpca(art.typProd)) continue;
      const qte = (pal.opca
        ? qtyPalierOpca(art, pal, camp.articles)
        : qtyParPack(art, pal, totalP, vent, camp.articles)) * (pal.nbPacks || 0);
      if (qte > 0) out.push({ ref, palier: libelle, qte });
    }
  }
  return out;
}

/** Agrège les besoins logistiques de plusieurs campagnes par réf et par mois ABSOLU.
 *  L'axe des mois s'étend automatiquement (déborde sur N+1 si une offre finit l'année suivante). */
export function buildSyntheseLogistique(campagnes: CampagneCreee[]): SyntheseLogistique {
  // accum[ref] = Map<moisAbsolu, qté> ; detailAbs = même chose par campagne × offre.
  const accum: Record<string, { name: string; parMoisAbs: Map<number, number> }> = {};
  const detailAbs: { ref: string; name: string; campagne: string; palier: string; parMoisAbs: Map<number, number> }[] = [];
  let minAbs = Infinity, maxAbs = -Infinity;
  const ignorees: string[] = [];

  for (const camp of campagnes) {
    const ad = absMonth(camp.dateDebut);
    const af = absMonth(camp.dateFin);
    if (ad == null || af == null || af < ad) { ignorees.push(camp.nom || "(sans nom)"); continue; }
    const nameByRef: Record<string, string> = {};
    for (const a of camp.articles) if (a.ref.trim()) nameByRef[a.ref.trim()] = a.name || "";

    // On répartit OFFRE PAR OFFRE, puis on agrège par référence : le total par réf est
    // inchangé, et chaque case de la synthèse devient explicable (onglet Détail + survol).
    for (const b of besoinsParOffre(camp)) {
      const rep = repartirAbsolu(b.qte, ad, af);
      if (!accum[b.ref]) accum[b.ref] = { name: nameByRef[b.ref] || "", parMoisAbs: new Map() };
      if (!accum[b.ref].name && nameByRef[b.ref]) accum[b.ref].name = nameByRef[b.ref];
      detailAbs.push({ ref: b.ref, name: nameByRef[b.ref] || "", campagne: camp.nom || "(sans nom)", palier: b.palier, parMoisAbs: rep });
      for (const [m, q] of rep) {
        accum[b.ref].parMoisAbs.set(m, (accum[b.ref].parMoisAbs.get(m) || 0) + q);
        if (m < minAbs) minAbs = m;
        if (m > maxAbs) maxAbs = m;
      }
    }
  }

  // Aucune donnée → synthèse vide.
  if (!isFinite(minAbs)) return { lignes: [], totalParMois: [], totalGeneral: 0, moisLabels: [], ignorees, detail: [] };

  // Axe des mois : du 1er au dernier mois de livraison (continu).
  const nbMois = maxAbs - minAbs + 1;
  const moisLabels: string[] = [];
  for (let i = 0; i < nbMois; i++) moisLabels.push(libelleMoisAbsolu(minAbs + i));

  const lignes: LigneLogistique[] = Object.entries(accum).map(([ref, v]) => {
    const parMois = new Array(nbMois).fill(0);
    for (const [m, q] of v.parMoisAbs) parMois[m - minAbs] += q;
    return { ref, name: v.name, parMois, total: parMois.reduce((s, x) => s + x, 0) };
  }).sort((a, b) => b.total - a.total);

  const detail: LigneDetailLogistique[] = detailAbs.map(d => {
    const parMois = new Array(nbMois).fill(0);
    for (const [m, q] of d.parMoisAbs) parMois[m - minAbs] += q;
    return { ref: d.ref, name: d.name, campagne: d.campagne, palier: d.palier, parMois, total: parMois.reduce((s, x) => s + x, 0) };
  }).sort((a, b) => b.total - a.total);

  const totalParMois = new Array(nbMois).fill(0);
  for (const l of lignes) for (let i = 0; i < nbMois; i++) totalParMois[i] += l.parMois[i];
  const totalGeneral = totalParMois.reduce((s, x) => s + x, 0);

  return { lignes, totalParMois, totalGeneral, moisLabels, ignorees, detail };
}
