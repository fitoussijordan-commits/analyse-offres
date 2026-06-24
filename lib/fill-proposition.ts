// lib/fill-proposition.ts — Remplissage d'une feuille "Proposition" à partir d'une campagne.
// Logique partagée entre l'export d'une campagne (/api/export-template) et l'export
// multi-campagnes (/api/export-multi). NE restructure PAS la feuille : écrit seulement les
// valeurs dans les lignes produits des 3 blocs + GRANDS COMPTES + Besoins logistiques.

import type ExcelJS from "exceljs";

export interface PropProduit {
  ref: string; name: string; productId: number;
  qtyParPack: number;
  barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
}
export interface PropPalier { code: string; label: string; qtyPacks: number; produits: PropProduit[]; }
export interface PropPayload { nom: string; paliers: PropPalier[]; }

export const PROP_SHEET = "Proposition template";

const BLOCKS = [
  { title: 3,  nbOffresRow: 4,  nbProduitsRow: 5,  pvFirst: 17, pvCount: 13, remiseRow: 10, nbOffRow: 11, synRow: 15, dataFirst: 17, dataLast: 36 },
  { title: 40, nbOffresRow: 42, nbProduitsRow: 43, pvFirst: 55, pvCount: 13, remiseRow: 48, nbOffRow: 49, synRow: 53, dataFirst: 55, dataLast: 74 },
  { title: 76, nbOffresRow: 78, nbProduitsRow: 79, pvFirst: 91, pvCount: 13, remiseRow: 84, nbOffRow: 85, synRow: 89, dataFirst: 91, dataLast: 110 },
];
const TYPO_COLS = [
  { ca: "M", marge: "N", param: "N" }, { ca: "O", marge: "P", param: "P" },
  { ca: "Q", marge: "R", param: "R" }, { ca: "S", marge: "T", param: "T" },
  { ca: "U", marge: "V", param: "V" }, { ca: "W", marge: "X", param: "X" },
  { ca: "Y", marge: "Z", param: "Z" },
];
const FMT_EUR = '#,##0.0 "€";(#,##0.0) "€";" - "';

function round2(n: number): number { return Math.round((n + Number.EPSILON) * 100) / 100; }

function writeTypoFormulas(ws: ExcelJS.Worksheet, row: number, remiseRow: number, nbOffRow: number) {
  for (const t of TYPO_COLS) {
    const nbOff = `$${t.param}$${nbOffRow}`, rem = `$${t.param}$${remiseRow}`;
    const ca = ws.getCell(`${t.ca}${row}`), mg = ws.getCell(`${t.marge}${row}`);
    ca.value = { formula: `E${row}*H${row}*(1-I${row})*${nbOff}*(1-${rem})` };
    mg.value = { formula: `${t.ca}${row}-E${row}*F${row}*${nbOff}` };
    ca.numFmt = FMT_EUR; mg.numFmt = FMT_EUR;
  }
}

/** Remplit la feuille `ws` (un gabarit Proposition vierge) avec les données d'une campagne. */
export function fillPropositionSheet(ws: ExcelJS.Worksheet, payload: PropPayload) {
  const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
  const used = Math.min(paliers.length, BLOCKS.length);

  for (let b = 0; b < BLOCKS.length; b++) {
    const blk = BLOCKS[b];
    const pal = b < used ? paliers[b] : null;

    if (pal) {
      const titre = [payload.nom, pal.code, pal.label].map(s => (s || "").trim()).filter(Boolean).join(" — ") || "Offre";
      ws.getCell(blk.title, 1).value = titre;
      ws.getCell(blk.nbOffresRow, 2).value = pal.qtyPacks || 0;
      ws.getCell(blk.nbProduitsRow, 2).value = pal.produits.reduce((s, p) => s + (p.qtyParPack || 0), 0);
    }

    // Recâbler les "Nb Offres" (bug gabarit : pointent tous sur $B$4).
    const bOffres = `$B$${blk.nbOffresRow}`;
    for (const t of TYPO_COLS) {
      const cell = ws.getCell(`${t.param}${blk.nbOffRow}`);
      const v: any = cell.value;
      if (v && typeof v === "object" && typeof v.formula === "string") cell.value = { formula: v.formula.replace(/\$B\$4/g, bOffres) };
    }

    for (let i = 0; i < blk.pvCount; i++) {
      const row = blk.pvFirst + i;
      const p = pal ? pal.produits[i] : undefined;
      if (p) {
        ws.getCell(row, 1).value = p.ref || "";
        ws.getCell(row, 2).value = p.name || "";
        ws.getCell(row, 3).value = p.barcode || "";
        ws.getCell(row, 5).value = p.qtyParPack || 0;
        ws.getCell(row, 6).value = round2(p.standardPrice || 0);
        ws.getCell(row, 8).value = round2(p.listPrice || 0);
        ws.getCell(row, 10).value = round2(p.ppc || 0);
        writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);
      } else {
        for (const c of [1, 2, 3, 5, 6, 8, 10]) ws.getCell(row, c).value = null;
        for (const t of TYPO_COLS) { ws.getCell(`${t.ca}${row}`).value = null; ws.getCell(`${t.marge}${row}`).value = null; }
      }
    }

    // Synthèse : SUM par colonne CA/Marges.
    for (const t of TYPO_COLS) {
      const ca = ws.getCell(`${t.ca}${blk.synRow}`), mg = ws.getCell(`${t.marge}${blk.synRow}`);
      ca.value = { formula: `SUM(${t.ca}${blk.dataFirst}:${t.ca}${blk.dataLast})` };
      mg.value = { formula: `SUM(${t.marge}${blk.dataFirst}:${t.marge}${blk.dataLast})` };
      ca.numFmt = FMT_EUR; mg.numFmt = FMT_EUR;
    }
    const ab = ws.getCell(blk.synRow, 28);
    ab.style = JSON.parse(JSON.stringify(ab.style || {}));
    ws.getCell(blk.synRow, 29).numFmt = FMT_EUR;
    ws.getCell(blk.synRow, 30).numFmt = FMT_EUR;
    ab.numFmt = '#,##0;(#,##0);" - "';
  }

  // GRANDS COMPTES (palier 1) : code, libellé, prix.
  const pal1 = paliers[0];
  const GC_PV_FIRST = 123, GC_PV_COUNT = 13, LOG_PV_FIRST = 149, LOG_PV_COUNT = 13;
  for (let i = 0; i < GC_PV_COUNT; i++) {
    const row = GC_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    if (p) {
      ws.getCell(row, 1).value = p.ref || ""; ws.getCell(row, 2).value = p.name || "";
      ws.getCell(row, 6).value = round2(p.standardPrice || 0); ws.getCell(row, 8).value = round2(p.listPrice || 0);
      ws.getCell(row, 10).value = round2(p.ppc || 0);
    } else for (const c of [1, 2, 6, 8, 10]) ws.getCell(row, c).value = null;
  }
  // Besoins logistiques (palier 1) : code, libellé.
  for (let i = 0; i < LOG_PV_COUNT; i++) {
    const row = LOG_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    ws.getCell(row, 1).value = p ? (p.ref || "") : null;
    ws.getCell(row, 2).value = p ? (p.name || "") : null;
  }
  // Vider PLV/Testeurs fixes du gabarit.
  for (const row of [136, 137, 138, 139, 140, 141, 142]) for (const c of [1, 2, 4, 6, 8, 10]) ws.getCell(row, c).value = null;
  for (const row of [162, 163, 164, 165, 166, 167, 168]) for (const c of [1, 2, 4]) ws.getCell(row, c).value = null;
}
