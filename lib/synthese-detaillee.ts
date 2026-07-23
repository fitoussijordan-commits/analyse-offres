// lib/synthese-detaillee.ts — Onglet "Synthèse détaillée" du fichier exporté.
// Reproduit l'onglet Synthèse détaillée de l'Aperçu : CA/marge par palier (offre) +
// ventilation CA/marge/nb offres par statut client (typologie), agrégée sur tous les paliers.

import type ExcelJS from "exceljs";
import { calcPalier, CalcPalier, TYPOLOGIES, DEFAULT_PCTS, DEFAULT_REMISES, REMISE_ADD_DEFAUT, calcGrandsComptes, GcEnseigneCalc, GcProduitInfo } from "@/lib/calc-offre";
import type { PropPalier, GcEnseignePayload, PropProduit } from "@/lib/fill-proposition";

/** Pricing par article pour le calcul GC : la clé suit la même logique que fill-proposition
 *  (réf seule si unique dans le palier de référence, sinon réf#type). Le palier 1 sert de
 *  référence de pricing (mêmes articles/prix dans tous les paliers). */
function gcProduitsInfo(paliers: PropPalier[]): GcProduitInfo[] {
  const pal1 = paliers[0];
  if (!pal1?.produits?.length) return [];
  const count: Record<string, number> = {};
  for (const p of pal1.produits) { const r = (p.ref || "").trim(); if (r) count[r] = (count[r] || 0) + 1; }
  const keyOf = (p: PropProduit) => { const r = (p.ref || "").trim(); return (r && count[r] > 1) ? `${r}#${p.typProd || "Produit Vente"}` : r; };
  const remiseAdd = pal1.remiseAddTaux != null ? pal1.remiseAddTaux : REMISE_ADD_DEFAUT;
  return pal1.produits.map(p => ({
    key: keyOf(p),
    listPrice: (p.typProd || "Produit Vente") === "Produit Vente" ? (p.listPrice || 0) : 0,
    standardPrice: p.standardPrice || 0,
    remiseAdd,
  }));
}

const FMT_EUR = '#,##0 "€";[Red]-#,##0 "€"';
const FMT_PCT = "0.0%";
const TEAL = "0D9488", BLUE = "1F3A5F", DARK = "1A1A2E", WHITE = "FFFFFF", SOFT = "F5F7FA";

/** Convertit un PropPalier (payload export) en CalcPalier (couche de calcul). */
function toCalc(p: PropPalier): CalcPalier {
  return {
    code: p.code, label: p.label, nbPacks: p.qtyPacks || 0,
    pcts: p.pctOffres && p.pctOffres.length === 7 ? p.pctOffres : DEFAULT_PCTS,
    remises: p.remises && p.remises.length === 7 ? p.remises : DEFAULT_REMISES,
    remiseAdd: p.remiseAddTaux != null ? p.remiseAddTaux : REMISE_ADD_DEFAUT,
    produits: (p.produits || []).map(pr => ({
      ref: pr.ref, name: pr.name || "",
      qtyParPack: pr.qtyParPack || 0,
      standardPrice: pr.standardPrice || 0,
      listPrice: pr.listPrice || 0,
      ppc: pr.ppc || 0,
    })),
  };
}

/** Écrit l'onglet "Synthèse détaillée" dans le classeur. */
export function writeSyntheseDetailleeSheet(wb: ExcelJS.Workbook, paliers: PropPalier[], gcEnseignes?: GcEnseignePayload[]) {
  // Supprime un onglet existant du même nom (idempotent).
  const existant = wb.getWorksheet("Synthèse détaillée");
  if (existant) wb.removeWorksheet(existant.id);
  const ws = wb.addWorksheet("Synthèse détaillée", { views: [{ showGridLines: false }] });

  const calcs = paliers.map(toCalc);
  ws.columns = [{ width: 26 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }];

  const bold = (cell: ExcelJS.Cell, color = DARK, size = 11) => { cell.font = { bold: true, color: { argb: color }, size }; };

  // ── Tableau 1 : CA / Marge par offre ──
  let row = 1;
  const t1 = ws.getCell(row, 1); t1.value = "CA / Marge par offre"; bold(t1, BLUE, 13); row += 1;
  const head1 = ["Offre", "Nb offres", "CA", "Marge €", "Marge %"];
  head1.forEach((h, i) => {
    const c = ws.getCell(row, i + 1); c.value = h;
    c.font = { bold: true, color: { argb: WHITE } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE } };
    c.alignment = { horizontal: i === 0 ? "left" : "right" };
  });
  row += 1;
  for (let i = 0; i < paliers.length; i++) {
    const r = calcPalier(calcs[i]);
    const p = paliers[i];
    ws.getCell(row, 1).value = `${p.code} — ${p.label}`.trim().replace(/^—\s*/, "");
    ws.getCell(row, 2).value = p.qtyPacks || 0;
    const ca = ws.getCell(row, 3); ca.value = r.caTotal; ca.numFmt = FMT_EUR;
    const mg = ws.getCell(row, 4); mg.value = r.margeTotal; mg.numFmt = FMT_EUR;
    const mp = ws.getCell(row, 5); mp.value = r.margePct; mp.numFmt = FMT_PCT;
    if (i % 2 === 1) for (let c = 1; c <= 5; c++) ws.getCell(row, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SOFT } };
    row += 1;
  }
  // Totaux : Retail + Institut (offres), Grands Comptes, puis total général.
  {
    const caRI = calcs.reduce((s, c) => s + calcPalier(c).caTotal, 0);
    const mgRI = calcs.reduce((s, c) => s + calcPalier(c).margeTotal, 0);
    const nbTot = paliers.reduce((s, p) => s + (p.qtyPacks || 0), 0);
    // CA / marge GC (0 si aucune enseigne saisie).
    const gc = calcGrandsComptes(
      (gcEnseignes || []).map(e => ({ nom: e.nom, remise: e.remise, qties: e.qties } as GcEnseigneCalc)),
      gcProduitsInfo(paliers),
    );

    const writeTot = (label: string, nb: number | "", ca: number, mg: number, color: string) => {
      ws.getCell(row, 1).value = label; bold(ws.getCell(row, 1), color);
      if (nb !== "") ws.getCell(row, 2).value = nb;
      const c3 = ws.getCell(row, 3); c3.value = ca; c3.numFmt = FMT_EUR; bold(c3, color);
      const c4 = ws.getCell(row, 4); c4.value = mg; c4.numFmt = FMT_EUR; bold(c4, color);
      const c5 = ws.getCell(row, 5); c5.value = ca > 0 ? mg / ca : 0; c5.numFmt = FMT_PCT; bold(c5, color);
      row += 1;
    };
    writeTot("TOTAL Retail + Institut", nbTot, caRI, mgRI, TEAL);
    if (gc.caTotal > 0) {
      writeTot("TOTAL Grands Comptes", "", gc.caTotal, gc.margeTotal, BLUE);
      writeTot("TOTAL GÉNÉRAL (RI + GC)", "", caRI + gc.caTotal, mgRI + gc.margeTotal, DARK);
    }
    row += 1;

    // Détail par enseigne GC (si des quantités ont été saisies).
    if (gc.parEnseigne.some(e => e.ca > 0)) {
      const tg = ws.getCell(row, 1); tg.value = "Détail Grands Comptes par enseigne"; bold(tg, BLUE, 13); row += 1;
      ["Enseigne", "Qté", "CA", "Marge €", "Marge %"].forEach((h, i) => {
        const c = ws.getCell(row, i + 1); c.value = h;
        c.font = { bold: true, color: { argb: WHITE } };
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE } };
        c.alignment = { horizontal: i === 0 ? "left" : "right" };
      });
      row += 1;
      gc.parEnseigne.filter(e => e.ca > 0).forEach((e, i) => {
        ws.getCell(row, 1).value = e.nom;
        ws.getCell(row, 2).value = Math.round(e.qty);
        const c3 = ws.getCell(row, 3); c3.value = e.ca; c3.numFmt = FMT_EUR;
        const c4 = ws.getCell(row, 4); c4.value = e.marge; c4.numFmt = FMT_EUR;
        const c5 = ws.getCell(row, 5); c5.value = e.ca > 0 ? e.marge / e.ca : 0; c5.numFmt = FMT_PCT;
        if (i % 2 === 1) for (let c = 1; c <= 5; c++) ws.getCell(row, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SOFT } };
        row += 1;
      });
      row += 1;
    }
  }

  // ── Tableau 2 : Ventilation par statut client (tous paliers) ──
  const t2 = ws.getCell(row, 1); t2.value = "Ventilation par statut client (tous paliers)"; bold(t2, TEAL, 13); row += 1;
  // En-tête : vide + 7 typologies + Total
  ws.getCell(row, 1).value = "";
  TYPOLOGIES.forEach((t, i) => {
    const c = ws.getCell(row, i + 2); c.value = t;
    c.font = { bold: true, color: { argb: WHITE } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TEAL } };
    c.alignment = { horizontal: "right" };
  });
  const cTot = ws.getCell(row, 9); cTot.value = "Total"; cTot.font = { bold: true, color: { argb: WHITE } };
  cTot.fill = { type: "pattern", pattern: "solid", fgColor: { argb: TEAL } }; cTot.alignment = { horizontal: "right" };
  row += 1;

  // Agrège par typologie sur tous les paliers.
  const nbOff = new Array(7).fill(0), ca = new Array(7).fill(0), marge = new Array(7).fill(0);
  for (const cp of calcs) {
    const r = calcPalier(cp);
    for (let t = 0; t < 7; t++) { nbOff[t] += r.parTypo[t].nbOffres; ca[t] += r.parTypo[t].ca; marge[t] += r.parTypo[t].marge; }
  }
  const writeRow = (label: string, arr: number[], numFmt?: string) => {
    bold(ws.getCell(row, 1));
    ws.getCell(row, 1).value = label;
    arr.forEach((v, i) => { const c = ws.getCell(row, i + 2); c.value = v; if (numFmt) c.numFmt = numFmt; });
    const tot = ws.getCell(row, 9); tot.value = arr.reduce((s, x) => s + x, 0); if (numFmt) tot.numFmt = numFmt; bold(tot);
    row += 1;
  };
  writeRow("Nb offres", nbOff.map(v => Math.round(v)));
  writeRow("CA", ca, FMT_EUR);
  writeRow("Marge", marge, FMT_EUR);

  return ws;
}

/** Onglet "Synthèse CA annuelle" : CA/marge/marge% par campagne + total sur l'année. */
export function writeSyntheseAnnuelleSheet(wb: ExcelJS.Workbook, campagnes: { nom: string; paliers: PropPalier[] }[]) {
  const existant = wb.getWorksheet("Synthèse CA annuelle");
  if (existant) wb.removeWorksheet(existant.id);
  const ws = wb.addWorksheet("Synthèse CA annuelle", { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 34 }, { width: 12 }, { width: 16 }, { width: 16 }, { width: 12 }];

  const titre = ws.getCell(1, 1); titre.value = "Synthèse CA annuelle — toutes campagnes"; titre.font = { bold: true, color: { argb: BLUE }, size: 14 };
  const head = ["Campagne", "Nb offres", "CA", "Marge €", "Marge %"];
  head.forEach((h, i) => {
    const c = ws.getCell(3, i + 1); c.value = h;
    c.font = { bold: true, color: { argb: WHITE } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BLUE } };
    c.alignment = { horizontal: i === 0 ? "left" : "right" };
  });

  let row = 4, caGrand = 0, mgGrand = 0, nbGrand = 0;
  for (let ci = 0; ci < campagnes.length; ci++) {
    const camp = campagnes[ci];
    const calcs = (camp.paliers || []).map(toCalc);
    let ca = 0, mg = 0, nb = 0;
    for (const cp of calcs) { const r = calcPalier(cp); ca += r.caTotal; mg += r.margeTotal; nb += cp.nbPacks; }
    caGrand += ca; mgGrand += mg; nbGrand += nb;
    ws.getCell(row, 1).value = camp.nom || `Campagne ${ci + 1}`;
    ws.getCell(row, 2).value = nb;
    const cca = ws.getCell(row, 3); cca.value = ca; cca.numFmt = FMT_EUR;
    const cmg = ws.getCell(row, 4); cmg.value = mg; cmg.numFmt = FMT_EUR;
    const cmp = ws.getCell(row, 5); cmp.value = ca > 0 ? mg / ca : 0; cmp.numFmt = FMT_PCT;
    if (ci % 2 === 1) for (let c = 1; c <= 5; c++) ws.getCell(row, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: SOFT } };
    row += 1;
  }
  // Total annuel
  ws.getCell(row, 1).value = "TOTAL ANNÉE"; ws.getCell(row, 1).font = { bold: true, color: { argb: TEAL } };
  ws.getCell(row, 2).value = nbGrand;
  const tca = ws.getCell(row, 3); tca.value = caGrand; tca.numFmt = FMT_EUR; tca.font = { bold: true, color: { argb: TEAL } };
  const tmg = ws.getCell(row, 4); tmg.value = mgGrand; tmg.numFmt = FMT_EUR; tmg.font = { bold: true, color: { argb: TEAL } };
  const tmp = ws.getCell(row, 5); tmp.value = caGrand > 0 ? mgGrand / caGrand : 0; tmp.numFmt = FMT_PCT; tmp.font = { bold: true, color: { argb: TEAL } };
  return ws;
}
