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
export interface PropPalier {
  code: string; label: string; qtyPacks: number; produits: PropProduit[];
  // Remise statut : si standard, on applique le même taux à toutes les typologies.
  remiseStandard?: boolean;   // true = standard (même % pour tous)
  remiseStandardTaux?: number; // ex. 0.17 pour 17%
}
export interface PropPayload { nom: string; paliers: PropPalier[]; }

export const PROP_SHEET = "Proposition template";
export const MAPPING_SHEET = "Mapping";
// Colonnes de l'onglet Mapping : A=réf, B=désignation, C=EAN, D=coût, E=tarif, F=PPC.
const MAP_HEADER = ["Code article", "Libellé article", "Code à barres (EAN)", "Coût achat unitaire", "Tarif revendeur unitaire", "PPC"];

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

// Remplit l'onglet Mapping avec tous les articles uniques (Odoo + réfs libres) et renvoie
// l'ensemble des réfs présentes (pour savoir si un VLOOKUP est possible).
function fillMapping(wb: ExcelJS.Workbook, payload: PropPayload): Set<string> {
  const map = wb.getWorksheet(MAPPING_SHEET);
  const refs = new Set<string>();
  if (!map) return refs;

  // En-tête.
  MAP_HEADER.forEach((h, i) => {
    const c = map.getCell(1, i + 1);
    c.value = h; c.font = { bold: true };
  });

  // Articles uniques par réf (1re occurrence non vide gagne).
  const byRef = new Map<string, PropProduit>();
  for (const pal of payload.paliers) for (const p of pal.produits) {
    const ref = (p.ref || "").trim();
    if (!ref) continue;
    if (!byRef.has(ref) || (!byRef.get(ref)!.name && p.name)) byRef.set(ref, p);
  }
  let row = 2;
  for (const [ref, p] of byRef) {
    map.getCell(row, 1).value = ref;
    map.getCell(row, 2).value = p.name || "";
    map.getCell(row, 3).value = p.barcode || "";
    map.getCell(row, 4).value = round2(p.standardPrice || 0);
    map.getCell(row, 5).value = round2(p.listPrice || 0);
    map.getCell(row, 6).value = round2(p.ppc || 0);
    refs.add(ref);
    row++;
  }
  return refs;
}

// VLOOKUP vers l'onglet Mapping : colonne `mapCol` (2=libellé,3=EAN,4=coût,5=tarif,6=PPC).
function vlookup(refCell: string, mapCol: number): string {
  return `IFERROR(VLOOKUP(${refCell},${MAPPING_SHEET}!$A:$F,${mapCol},FALSE),"")`;
}

/**
 * Remplit le classeur (3 onglets) à partir d'une campagne :
 *  - onglet Mapping : base articles (réf, désignation, EAN, coût, tarif, PPC),
 *  - onglet Proposition : code article (valeur) + B/C/F/H/J en VLOOKUP vers Mapping,
 *    remise standard/spécifique par palier, CA/Marges, GC, logistique.
 */
export function fillPropositionWorkbook(wb: ExcelJS.Workbook, payload: PropPayload) {
  const ws = wb.getWorksheet(PROP_SHEET);
  if (!ws) return;
  const mapRefs = fillMapping(wb, payload);
  fillProposition(ws, payload, mapRefs);
}

function fillProposition(ws: ExcelJS.Worksheet, payload: PropPayload, mapRefs: Set<string>) {
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

      // Remise statut STANDARD : même taux pour toutes les typologies du palier.
      if (pal.remiseStandard && typeof pal.remiseStandardTaux === "number") {
        for (const t of TYPO_COLS) ws.getCell(`${t.param}${blk.remiseRow}`).value = pal.remiseStandardTaux;
      }
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
        const ref = (p.ref || "").trim();
        ws.getCell(row, 1).value = ref;                                  // A : code (valeur)
        // B/C/F/H/J = VLOOKUP vers Mapping si la réf y est, sinon valeur en dur (sécurité).
        if (mapRefs.has(ref)) {
          ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) }; // libellé
          ws.getCell(row, 3).value = { formula: vlookup(`A${row}`, 3) }; // EAN
          ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) }; // coût
          ws.getCell(row, 8).value = { formula: vlookup(`A${row}`, 5) }; // tarif
          ws.getCell(row, 10).value = { formula: vlookup(`A${row}`, 6) }; // PPC
        } else {
          ws.getCell(row, 2).value = p.name || "";
          ws.getCell(row, 3).value = p.barcode || "";
          ws.getCell(row, 6).value = round2(p.standardPrice || 0);
          ws.getCell(row, 8).value = round2(p.listPrice || 0);
          ws.getCell(row, 10).value = round2(p.ppc || 0);
        }
        ws.getCell(row, 5).value = p.qtyParPack || 0;                    // E : qté/pack
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

  // GRANDS COMPTES (palier 1) : code (valeur) + libellé/prix en VLOOKUP vers Mapping.
  const pal1 = paliers[0];
  const GC_PV_FIRST = 123, GC_PV_COUNT = 13, LOG_PV_FIRST = 149, LOG_PV_COUNT = 13;
  for (let i = 0; i < GC_PV_COUNT; i++) {
    const row = GC_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    if (p) {
      const ref = (p.ref || "").trim();
      ws.getCell(row, 1).value = ref;
      if (mapRefs.has(ref)) {
        ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
        ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
        ws.getCell(row, 8).value = { formula: vlookup(`A${row}`, 5) };
        ws.getCell(row, 10).value = { formula: vlookup(`A${row}`, 6) };
      } else {
        ws.getCell(row, 2).value = p.name || "";
        ws.getCell(row, 6).value = round2(p.standardPrice || 0); ws.getCell(row, 8).value = round2(p.listPrice || 0);
        ws.getCell(row, 10).value = round2(p.ppc || 0);
      }
    } else for (const c of [1, 2, 6, 8, 10]) ws.getCell(row, c).value = null;
  }
  // Besoins logistiques (palier 1) : code + libellé en VLOOKUP.
  for (let i = 0; i < LOG_PV_COUNT; i++) {
    const row = LOG_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    const ref = p ? (p.ref || "").trim() : "";
    ws.getCell(row, 1).value = ref || null;
    ws.getCell(row, 2).value = ref && mapRefs.has(ref) ? { formula: vlookup(`A${row}`, 2) } : (p ? (p.name || "") : null);
  }
  // Vider PLV/Testeurs fixes du gabarit.
  for (const row of [136, 137, 138, 139, 140, 141, 142]) for (const c of [1, 2, 4, 6, 8, 10]) ws.getCell(row, c).value = null;
  for (const row of [162, 163, 164, 165, 166, 167, 168]) for (const c of [1, 2, 4]) ws.getCell(row, c).value = null;
}

// Compat : ancienne signature (feuille seule, sans Mapping). Garde le remplissage en valeurs.
export function fillPropositionSheet(ws: ExcelJS.Worksheet, payload: PropPayload) {
  fillProposition(ws, payload, new Set());
}
