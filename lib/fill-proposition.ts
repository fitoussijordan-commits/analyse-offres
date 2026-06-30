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
// Ligne du Mapping (catalogue complet ou articles campagne).
export interface MapRow { ref: string; name?: string; barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number; }
export interface PropPayload { nom: string; paliers: PropPalier[]; mapping?: MapRow[]; }

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
// Valeurs par défaut du gabarit (% offres et remises par typologie : Ambassadeur..Calendula).
const DEFAULT_PCTS = [0.5, 0.1, 0.1, 0.1, 0.1, 0.05, 0.05];
const DEFAULT_REMISES = [0.17, 0.13, 0.08, 0.325, 0.3, 0.28, 0.25];

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

  // Source : catalogue complet fourni (payload.mapping) si présent, sinon articles de la campagne.
  const byRef = new Map<string, MapRow>();
  if (payload.mapping && payload.mapping.length) {
    for (const m of payload.mapping) {
      const ref = (m.ref || "").trim();
      if (ref && !byRef.has(ref)) byRef.set(ref, m);
    }
  }
  // Toujours s'assurer que les articles de la campagne (dont réfs libres avec prix manuels)
  // sont présents — ils complètent/écrasent le catalogue pour les réfs manuelles.
  for (const pal of payload.paliers) for (const p of pal.produits) {
    const ref = (p.ref || "").trim();
    if (!ref) continue;
    const isManual = p.productId === 0; // réf libre / hors Odoo
    if (isManual || !byRef.has(ref)) byRef.set(ref, p);
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
  fillSynthese(wb, payload);
}

const SYNTHESE_SHEET = "Synthese";
// Mots-clés pour mapper un palier à une ligne d'offre de l'onglet Synthese.
const OFFRE_ROWS: Array<{ row: number; keys: string[] }> = [
  { row: 2, keys: ["tg vip", "vip"] },
  { row: 3, keys: ["premium"] },
  { row: 4, keys: ["standard"] },
  { row: 5, keys: ["essentiel"] },
  { row: 6, keys: ["gc", "grand compte"] },
];

function norm(s: string): string {
  return (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

/**
 * Remplit l'onglet Synthese :
 *  - HAUT : par offre (ligne mappée via mot-clé du label du palier) → Nb offres (B),
 *    CA (C), Marge € (D), Marge % (E).
 *  - BAS : pour chaque réf des tableaux (Produits ventes / Testeur / PLV…), colonne
 *    "Offres retail+Institut" (E) = besoin total = somme(qté/pack × nb packs) sur les paliers ;
 *    TOTAL (K) = somme E:J. Les colonnes GC/MDRH/ESHOP/DC/Marketing restent vides (saisie main).
 */
function fillSynthese(wb: ExcelJS.Workbook, payload: PropPayload) {
  const sw = wb.getWorksheet(SYNTHESE_SHEET);
  if (!sw) return;
  const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);

  // ── HAUT : CA/Marge par offre ──────────────────────────────────────────────
  const usedRows = new Set<number>();
  for (const pal of paliers) {
    const label = norm(`${pal.label} ${pal.code}`);
    const match = OFFRE_ROWS.find(o => !usedRows.has(o.row) && o.keys.some(k => label.includes(k)));
    if (!match) continue;
    usedRows.add(match.row);
    const nbOffres = pal.qtyPacks || 0;
    // CA et marge = somme sur les produits : CA = E×H×(1-I)×NbOffres×(1-remise) cumulé toutes typologies.
    // On recalcule simplement ici en € (mêmes formules que le bloc), sommé sur produits × typologies.
    let ca = 0, marge = 0;
    const remises = pal.remiseStandard && pal.remiseStandardTaux != null
      ? new Array(7).fill(pal.remiseStandardTaux)
      : DEFAULT_REMISES;
    const pcts = DEFAULT_PCTS;
    for (const p of pal.produits) {
      const E = p.qtyParPack || 0, H = p.listPrice || 0, F = p.standardPrice || 0, I = 0.15;
      for (let t = 0; t < 7; t++) {
        const nbOff = pcts[t] * nbOffres;
        const caT = E * H * (1 - I) * nbOff * (1 - remises[t]);
        ca += caT;
        marge += caT - E * F * nbOff;
      }
    }
    sw.getCell(match.row, 2).value = nbOffres;
    sw.getCell(match.row, 3).value = round2(ca);
    sw.getCell(match.row, 4).value = round2(marge);
    sw.getCell(match.row, 5).value = ca > 0 ? round2(marge / ca) : 0;
    sw.getCell(match.row, 3).numFmt = FMT_EUR; sw.getCell(match.row, 4).numFmt = FMT_EUR;
    sw.getCell(match.row, 5).numFmt = "0.0%";
  }

  // ── BAS : besoin total par réf (col E) + TOTAL (col K) ─────────────────────
  // Besoin total d'une réf = somme sur paliers de (qté/pack × nb packs).
  const besoin: Record<string, number> = {};
  for (const pal of paliers) for (const p of pal.produits) {
    const ref = (p.ref || "").trim();
    if (!ref) continue;
    besoin[ref] = (besoin[ref] || 0) + (p.qtyParPack || 0) * (pal.qtyPacks || 0);
  }
  // Parcourir toutes les lignes de l'onglet : si col A = une réf connue, remplir E et K.
  const maxRow = sw.rowCount;
  for (let r = 1; r <= maxRow; r++) {
    const a = sw.getCell(r, 1).value;
    const ref = a == null ? "" : String(a).trim();
    if (ref && besoin[ref] != null) {
      sw.getCell(r, 5).value = besoin[ref];                       // E : Offres retail+Institut
      sw.getCell(r, 11).value = { formula: `SUM(E${r}:J${r})` };  // K : TOTAL
      sw.getCell(r, 5).numFmt = "#,##0"; sw.getCell(r, 11).numFmt = "#,##0";
    }
  }
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

      // % Offres par typologie : on écrit la reco (valeurs par défaut DEFAULT_PCTS) sur chaque
      // palier, pour que la répartition soit toujours renseignée. Ligne %Offres = remiseRow-1.
      const pctRow = blk.remiseRow - 1;
      TYPO_COLS.forEach((t, idx) => { ws.getCell(`${t.param}${pctRow}`).value = DEFAULT_PCTS[idx]; });

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

    // VLOOKUP sur TOUTES les lignes produits (même vides) : ainsi, si l'utilisateur tape une
    // réf dans une ligne vide, libellé/EAN/prix/PPC se remplissent automatiquement depuis le
    // Mapping. B/C/F/H/J sont donc toujours des formules VLOOKUP basées sur la colonne A.
    for (let i = 0; i < blk.pvCount; i++) {
      const row = blk.pvFirst + i;
      const p = pal ? pal.produits[i] : undefined;
      const ref = p ? (p.ref || "").trim() : "";
      // Réf manuelle hors catalogue (productId 0 et absente du Mapping) → valeurs en dur.
      const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);

      ws.getCell(row, 1).value = ref || null;                            // A : code (valeur)
      if (horsMapping) {
        ws.getCell(row, 2).value = p!.name || "";
        ws.getCell(row, 3).value = p!.barcode || "";
        ws.getCell(row, 6).value = round2(p!.standardPrice || 0);
        ws.getCell(row, 8).value = round2(p!.listPrice || 0);
        ws.getCell(row, 10).value = round2(p!.ppc || 0);
      } else {
        // VLOOKUP (renvoie "" si A vide → ligne vide propre, mais remplissable).
        ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
        ws.getCell(row, 3).value = { formula: vlookup(`A${row}`, 3) };
        ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
        ws.getCell(row, 8).value = { formula: vlookup(`A${row}`, 5) };
        ws.getCell(row, 10).value = { formula: vlookup(`A${row}`, 6) };
      }
      ws.getCell(row, 5).value = p ? (p.qtyParPack || 0) : null;          // E : qté/pack
      // CA/Marges : toujours en formules (donnent 0 si E vide), pour rester remplissables.
      writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);
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
    const ref = p ? (p.ref || "").trim() : "";
    const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);
    ws.getCell(row, 1).value = ref || null;
    if (horsMapping) {
      ws.getCell(row, 2).value = p!.name || "";
      ws.getCell(row, 6).value = round2(p!.standardPrice || 0); ws.getCell(row, 8).value = round2(p!.listPrice || 0);
      ws.getCell(row, 10).value = round2(p!.ppc || 0);
    } else {
      ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
      ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
      ws.getCell(row, 8).value = { formula: vlookup(`A${row}`, 5) };
      ws.getCell(row, 10).value = { formula: vlookup(`A${row}`, 6) };
    }
  }
  // Besoins logistiques (palier 1) : code + libellé en VLOOKUP sur toutes les lignes.
  for (let i = 0; i < LOG_PV_COUNT; i++) {
    const row = LOG_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    const ref = p ? (p.ref || "").trim() : "";
    const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);
    ws.getCell(row, 1).value = ref || null;
    ws.getCell(row, 2).value = horsMapping ? (p!.name || "") : { formula: vlookup(`A${row}`, 2) };
  }
  // Vider PLV/Testeurs fixes du gabarit.
  for (const row of [136, 137, 138, 139, 140, 141, 142]) for (const c of [1, 2, 4, 6, 8, 10]) ws.getCell(row, c).value = null;
  for (const row of [162, 163, 164, 165, 166, 167, 168]) for (const c of [1, 2, 4]) ws.getCell(row, c).value = null;
}

// Compat : ancienne signature (feuille seule, sans Mapping). Garde le remplissage en valeurs.
export function fillPropositionSheet(ws: ExcelJS.Worksheet, payload: PropPayload) {
  fillProposition(ws, payload, new Set());
}
