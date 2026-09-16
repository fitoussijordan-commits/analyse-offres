// lib/fill-proposition.ts — Remplissage d'une feuille "Proposition" à partir d'une campagne.
// Logique partagée entre l'export d'une campagne (/api/export-template) et l'export
// multi-campagnes (/api/export-multi). NE restructure PAS la feuille : écrit seulement les
// valeurs dans les lignes produits des 3 blocs + GRANDS COMPTES + Besoins logistiques.

import type ExcelJS from "exceljs";
import { genereCA, estOpca, OPCA_TAUX_COUT } from "@/lib/type-produit";

export interface PropProduit {
  ref: string; name: string; productId: number;
  qtyParPack: number;
  barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
  typProd?: string;
}
export interface PropPalier {
  code: string; label: string; qtyPacks: number; produits: PropProduit[]; descriptif?: string;
  // Remise statut : si standard, on applique le même taux à toutes les typologies.
  remiseStandard?: boolean;   // true = standard (même % pour tous)
  remiseStandardTaux?: number; // ex. 0.17 pour 17%
  // Remise additionnelle (colonne I) appliquée à tous les produits du palier (ex. 0.15).
  remiseAddTaux?: number;
  // % offres reco par typologie (7 valeurs, ordre Ambassadeur..Calendula) propre à ce palier.
  pctOffres?: number[];
  // Remises par typologie (7 valeurs) — surchargent les remises du gabarit si fournies.
  remises?: number[];
}
// Ligne du Mapping (catalogue complet ou articles campagne).
export interface MapRow { ref: string; name?: string; barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number; }
export interface GcEnseignePayload { nom: string; remise: number; qties: Record<string, number>; }
export interface PropPayload { nom: string; paliers: PropPalier[]; mapping?: MapRow[]; gcEnseignes?: GcEnseignePayload[]; canauxNonB2B?: GcEnseignePayload[]; }

export const PROP_SHEET = "Proposition template";

/** Un produit génère du CA seulement s'il est "Produit Vente". UG / Testeur / PLV /
 *  Échantillon sont gratuits → prix de vente (col 8) et PPC (col 10) forcés à 0. */
function estVente(p?: { typProd?: string }): boolean {
  return !p || genereCA(p.typProd);
}
export const MAPPING_SHEET = "Mapping";
// Colonnes de l'onglet Mapping : A=réf, B=désignation, C=EAN, D=coût, E=tarif, F=PPC.
const MAP_HEADER = ["Code article", "Libellé article", "Code à barres (EAN)", "Coût achat unitaire", "Tarif revendeur unitaire", "PPC"];

// Positions calées sur "TEMPLATE VIERGE_Campagne marketing.xlsx" (4 paliers + GC en L149).
// Chaque bloc : 20 lignes de données (13 Produit Vente + 7 PLV/Testeurs).
const BLOCKS = [
  { title: 3,   nbOffresRow: 4,   nbProduitsRow: 5,   pvFirst: 17,  pvCount: 13, remiseRow: 10,  nbOffRow: 11,  synRow: 15,  dataFirst: 17,  dataLast: 36 },
  { title: 39,  nbOffresRow: 40,  nbProduitsRow: 41,  pvFirst: 53,  pvCount: 13, remiseRow: 46,  nbOffRow: 47,  synRow: 51,  dataFirst: 53,  dataLast: 72 },
  { title: 75,  nbOffresRow: 77,  nbProduitsRow: 78,  pvFirst: 90,  pvCount: 13, remiseRow: 83,  nbOffRow: 84,  synRow: 88,  dataFirst: 90,  dataLast: 109 },
  { title: 112, nbOffresRow: 114, nbProduitsRow: 115, pvFirst: 127, pvCount: 13, remiseRow: 120, nbOffRow: 121, synRow: 125, dataFirst: 127, dataLast: 146 },
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
export function fillMapping(wb: ExcelJS.Workbook, payload: PropPayload): Set<string> {
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
    // Réf du Mapping aussi en TEXTE pour matcher la colonne A du template (VLOOKUP homogène).
    map.getCell(row, 1).numFmt = "@";
    map.getCell(row, 1).value = String(ref);
    refs.add(ref);
    row++;
  }
  return refs;
}

// Écrit une référence en TEXTE (format "@") pour que le VLOOKUP compare des types homogènes.
// Sans ça, une réf nombre (3020202) ne matche pas une réf texte dans le Mapping.
function setRefText(ws: ExcelJS.Worksheet, row: number, ref: string) {
  const c = ws.getCell(row, 1);
  c.numFmt = "@";
  c.value = ref ? String(ref) : null;
}

// VLOOKUP vers l'onglet Mapping : colonne `mapCol` (2=libellé,3=EAN,4=coût,5=tarif,6=PPC).
// Le repli (réf vide ou introuvable) dépend du type de colonne :
//   - texte (libellé, EAN) → "" ;
//   - numérique (coût, tarif, PPC) → 0, sinon les formules CA/Marges plantent en #VALEUR!.
function vlookup(refCell: string, mapCol: number): string {
  const numericCols = [4, 5, 6]; // coût, tarif, PPC
  const fallback = numericCols.includes(mapCol) ? "0" : '""';
  return `IFERROR(VLOOKUP(${refCell},${MAPPING_SHEET}!$A:$F,${mapCol},FALSE),${fallback})`;
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
  const blocks = fillProposition(ws, payload, mapRefs);
  fillSynthese(wb, payload, blocks);
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
function fillSynthese(wb: ExcelJS.Workbook, payload: PropPayload, blocks: BlockPos[] = BLOCKS) {
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
      const E = p.qtyParPack || 0, H = p.listPrice || 0, F = p.standardPrice || 0, I = estOpca(p.typProd) ? 0 : 0.15;
      for (let t = 0; t < 7; t++) {
        const nbOff = pcts[t] * nbOffres;
        const caT = E * H * (1 - I) * nbOff * (1 - remises[t]);
        ca += caT;
        marge += caT - E * F * nbOff;
      }
    }
    // Formats : B = Nombre d'offres (NOMBRE), C = CA (€), D = Marge (€), E = Marge % (%).
    // On clone le style de chaque cellule avant de fixer son format pour casser tout partage
    // d'objet style hérité du gabarit (qui mettait du % partout).
    const setFmt = (col: number, val: any, fmt: string) => {
      const c = sw.getCell(match.row, col);
      c.style = JSON.parse(JSON.stringify(c.style || {}));
      c.value = val; c.numFmt = fmt;
    };
    setFmt(2, nbOffres, '#,##0;(#,##0);" - "');                 // B : nombre
    setFmt(3, round2(ca), FMT_EUR);                              // C : CA €
    setFmt(4, round2(marge), FMT_EUR);                           // D : Marge €
    setFmt(5, ca > 0 ? round2(marge / ca) : 0, "0.0%");          // E : Marge %
  }

  // ── BAS : par réf des tableaux (les réfs sont déjà écrites en colonne A du gabarit) ──
  // Pour que la Synthese SUIVE les changements de réf faits dans le template :
  //   - B (libellé) et C (EAN) → VLOOKUP sur le Mapping (basé sur la réf en colonne A).
  //   - E (Offres retail+Institut) → besoin total de la réf, agrégé par formule depuis les
  //     blocs Proposition : pour chaque bloc, SUMIF(A produits = réf ; E produits) × Nb offres
  //     du bloc. Ainsi, changer une réf en colonne A met tout à jour automatiquement.
  //   - K (TOTAL) → somme E:J.
  // Plages "Produit Vente" + Testeurs/PLV de chaque bloc dans Proposition.
  // Positions RÉELLES des blocs (décalées si lignes d'articles ou blocs de paliers ajoutés).
  const propRanges = blocks.map(b => ({ first: b.pvFirst, last: b.dataLast, nbOffresCell: `'${PROP_SHEET}'!$B$${b.nbOffresRow}` }));

  const maxRow = sw.rowCount;
  // IMPORTANT : ne traiter QUE les tableaux du bas (à partir de la ligne 10). La zone du HAUT
  // (lignes 1-6 : SYNTHESE + offres) ne doit PAS être touchée, sinon ses CA/Marges sont écrasés.
  const BAS_FIRST = 10;
  for (let r = BAS_FIRST; r <= maxRow; r++) {
    const a = sw.getCell(r, 1).value;
    const ref = a == null ? "" : String(a).trim();
    if (!ref) continue;
    // Ne traiter que les lignes DATA (réf article), pas les en-têtes de tableau.
    const typ = sw.getCell(r, 4).value; // colonne D = Typ. Prod sur les lignes data
    const isDataRow = typ != null && String(typ).trim() !== "";
    if (!isDataRow) continue;

    // Réf en TEXTE + libellé/EAN en VLOOKUP (suivent si on change la réf).
    sw.getCell(r, 1).numFmt = "@";
    sw.getCell(r, 1).value = String(ref);
    sw.getCell(r, 2).value = { formula: `IFERROR(VLOOKUP(A${r},${MAPPING_SHEET}!$A:$F,2,FALSE),"")` };
    sw.getCell(r, 3).value = { formula: `IFERROR(VLOOKUP(A${r},${MAPPING_SHEET}!$A:$F,3,FALSE),"")` };

    // E : besoin = Σ_bloc SUMIF(A bloc = A<r> ; E bloc) × Nb offres bloc.
    const terms = propRanges.map(rg =>
      `SUMIF('${PROP_SHEET}'!$A$${rg.first}:$A$${rg.last},$A${r},'${PROP_SHEET}'!$E$${rg.first}:$E$${rg.last})*${rg.nbOffresCell}`
    );
    sw.getCell(r, 5).value = { formula: terms.join("+") };
    sw.getCell(r, 11).value = { formula: `SUM(E${r}:J${r})` };
    sw.getCell(r, 5).numFmt = "#,##0"; sw.getCell(r, 11).numFmt = "#,##0";
  }
}

/** Convertit toutes les formules PARTAGÉES du gabarit en formules normales. Obligatoire
 *  avant d'insérer des lignes : sinon les "clones" perdent leur cellule maître et Excel
 *  refuse d'ouvrir le fichier ("Shared Formula master must exist"). */
function unshareFormulas(ws: ExcelJS.Worksheet) {
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      const v: any = cell.value;
      if (v && typeof v === "object" && v.sharedFormula) {
        try { cell.value = { formula: cell.formula }; } catch { /* laisse tel quel */ }
      }
    });
  });
}

/** Après une insertion de `n` lignes à la position `at`, décale les références de ligne des
 *  formules existantes (ExcelJS, contrairement à Excel, ne le fait PAS : sans ça les formules
 *  du gabarit continuent de pointer sur les anciennes lignes → #VALUE! en cascade).
 *  Les références absolues ($10) se décalent aussi : une insertion déplace la cellule visée. */
function shiftFormulaRefs(ws: ExcelJS.Worksheet, at: number, n: number) {
  const re = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      const v: any = cell.value;
      if (!v || typeof v !== "object" || typeof v.formula !== "string") return;
      const shifted = v.formula.replace(re, (m: string, d1: string, col: string, d2: string, rowStr: string) => {
        const r = parseInt(rowStr, 10);
        return r >= at ? `${d1}${col}${d2}${r + n}` : m;
      });
      if (shifted !== v.formula) cell.value = { formula: shifted };
    });
  });
}

type BlockPos = (typeof BLOCKS)[number];

function fillProposition(ws: ExcelJS.Worksheet, payload: PropPayload, mapRefs: Set<string>): BlockPos[] {
  const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
  let used = Math.min(paliers.length, BLOCKS.length);

  // ── Capacité des blocs : le gabarit réserve 20 lignes d'articles par palier. Si une
  //    campagne en a davantage, on AGRANDIT chaque bloc en insérant les lignes manquantes
  //    (du dernier bloc vers le premier) et on décale toutes les positions en aval.
  //    Sans ça, les articles au-delà de la capacité étaient purement et simplement perdus.
  const blocks = BLOCKS.map(b => ({ ...b }));
  const gcPos = { top: 149, band: 150, nom: 152, remise: 153, hdr: 154, unit: 155, syn: 157, pvFirst: 159, last: 178 };
  let logFirst = 185;   // première ligne de la zone « Besoins logistiques » du gabarit

  // ── Plus de 4 paliers : le gabarit n'a que 4 blocs, les suivants étaient perdus. On
  //    recopie le bloc 1 (titre → lignes vides de séparation) autant de fois que nécessaire,
  //    juste avant GRANDS COMPTES, et on décale le GC et la zone logistique d'autant.
  const blocsManquants = Math.max(0, paliers.length - blocks.length);
  if (blocsManquants > 0) {
    unshareFormulas(ws);
    const SRC_FIRST = BLOCKS[0].title, SRC_LAST = BLOCKS[1].title - 1;   // lignes 3 → 38
    const HAUT = SRC_LAST - SRC_FIRST + 1;
    const NB_COLS_COPIE = 40;
    // Instantané du bloc source AVANT l'insertion (valeurs, formules, styles, hauteurs, fusions).
    const snap = [] as { height?: number; cells: { value: any; style: any }[] }[];
    for (let r = SRC_FIRST; r <= SRC_LAST; r++) {
      const row = ws.getRow(r);
      const cells = [];
      for (let c = 1; c <= NB_COLS_COPIE; c++) {
        const cell = row.getCell(c);
        const v: any = cell.value;
        cells.push({ value: v && typeof v === "object" && v.formula ? { formula: v.formula } : (v && typeof v === "object" ? JSON.parse(JSON.stringify(v)) : v), style: JSON.parse(JSON.stringify(cell.style || {})) });
      }
      snap.push({ height: row.height, cells });
    }
    const fusions = ((ws.model as any).merges as string[] || [])
      .map(rg => rg.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/))
      .filter((m): m is RegExpMatchArray => !!m && +m[2] >= SRC_FIRST && +m[4] <= SRC_LAST);

    const at = gcPos.top;
    const total = HAUT * blocsManquants;
    ws.spliceRows(at, 0, ...Array.from({ length: total }, () => [] as any[]));
    shiftFormulaRefs(ws, at, total);

    const refLigne = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
    for (let k = 0; k < blocsManquants; k++) {
      const delta = at + k * HAUT - SRC_FIRST;
      snap.forEach((sr, i) => {
        const dst = ws.getRow(SRC_FIRST + i + delta);
        if (sr.height) dst.height = sr.height;
        sr.cells.forEach((sc, ci) => {
          const cell = dst.getCell(ci + 1);
          cell.style = JSON.parse(JSON.stringify(sc.style));
          if (sc.value && typeof sc.value === "object" && typeof sc.value.formula === "string") {
            // Références internes au bloc décalées vers la copie ; le reste (Mapping!$A:$F…) intact.
            cell.value = { formula: sc.value.formula.replace(refLigne, (m: string, d1: string, col: string, d2: string, rs: string) => {
              const r = parseInt(rs, 10);
              return r >= SRC_FIRST && r <= SRC_LAST ? `${d1}${col}${d2}${r + delta}` : m;
            }) };
          } else {
            cell.value = sc.value ?? null;
          }
        });
      });
      for (const m of fusions) ws.mergeCells(`${m[1]}${+m[2] + delta}:${m[3]}${+m[4] + delta}`);
      const b0 = BLOCKS[0];
      blocks.push({
        title: b0.title + delta, nbOffresRow: b0.nbOffresRow + delta, nbProduitsRow: b0.nbProduitsRow + delta,
        pvFirst: b0.pvFirst + delta, pvCount: b0.pvCount, remiseRow: b0.remiseRow + delta, nbOffRow: b0.nbOffRow + delta,
        synRow: b0.synRow + delta, dataFirst: b0.dataFirst + delta, dataLast: b0.dataLast + delta,
      });
    }
    for (const key of Object.keys(gcPos) as (keyof typeof gcPos)[]) gcPos[key] += total;
    logFirst += total;
  }
  used = Math.min(paliers.length, blocks.length);
  const capacite = blocks[0].dataLast - blocks[0].pvFirst + 1;             // 20 dans le gabarit
  const maxProduits = paliers.length ? Math.max(...paliers.slice(0, blocks.length).map(p => p.produits.length)) : 0;
  const extra = Math.max(0, maxProduits - capacite);

  if (extra > 0) {
    unshareFormulas(ws);
    const styleCols = 34;
    for (let bi = blocks.length - 1; bi >= 0; bi--) {
      const at = blocks[bi].dataLast + 1;                                  // insertion en fin de zone data
      ws.spliceRows(at, 0, ...Array.from({ length: extra }, () => [] as any[]));
      shiftFormulaRefs(ws, at, extra);   // recale les formules du gabarit sur les nouvelles lignes
      // Style des nouvelles lignes recopié d'une ligne d'article existante du même bloc.
      const src = ws.getRow(blocks[bi].pvFirst + 1);
      for (let k = 0; k < extra; k++) {
        const dst = ws.getRow(at + k);
        dst.height = src.height;
        for (let c = 1; c <= styleCols; c++) dst.getCell(c).style = JSON.parse(JSON.stringify(src.getCell(c).style || {}));
      }
      // Décaler tout ce qui se trouve APRÈS le point d'insertion.
      blocks[bi].dataLast += extra;
      for (let bj = bi + 1; bj < blocks.length; bj++) {
        const b = blocks[bj];
        b.title += extra; b.nbOffresRow += extra; b.nbProduitsRow += extra; b.pvFirst += extra;
        b.remiseRow += extra; b.nbOffRow += extra; b.synRow += extra; b.dataFirst += extra; b.dataLast += extra;
      }
      for (const k of Object.keys(gcPos) as (keyof typeof gcPos)[]) gcPos[k] += extra;
      logFirst += extra;
    }
  }

  for (let b = 0; b < blocks.length; b++) {
    const blk = blocks[b];
    const pal = b < used ? paliers[b] : null;

    if (pal) {
      const titre = [payload.nom, pal.code, pal.label].map(s => (s || "").trim()).filter(Boolean).join(" — ") || "Offre";
      ws.getCell(blk.title, 1).value = titre;
      // Descriptif libre du palier → colonne D de la ligne titre (à côté du nom Premium/Standard).
      if (pal.descriptif != null) ws.getCell(blk.title, 4).value = pal.descriptif.trim();
      ws.getCell(blk.nbOffresRow, 2).value = pal.qtyPacks || 0;
      // Nombre de produits = SUM des "Pdt dans offre" (colonne E) des lignes Produit Vente.
      // Formule (pas une valeur) pour rester cohérent si on édite une qté à la main.
      ws.getCell(blk.nbProduitsRow, 2).value = { formula: `SUM(E${blk.pvFirst}:E${blk.dataLast})` };

      // % Offres par typologie : si une reco PROPRE AU PALIER (commandes N-1 du code offre par
      // statut) est fournie, on l'écrit sur la ligne %Offres (= remiseRow-1). Sinon on laisse
      // les valeurs en dur du gabarit. Chaque palier a donc sa propre répartition.
      if (pal.pctOffres && pal.pctOffres.length === TYPO_COLS.length) {
        const pctRow = blk.remiseRow - 1;
        TYPO_COLS.forEach((t, idx) => { ws.getCell(`${t.param}${pctRow}`).value = pal.pctOffres![idx]; });
      }

      // Remises par typologie : priorité (1) remises éditées (aperçu), (2) remise standard,
      // (3) valeurs du gabarit (on ne touche pas).
      if (pal.remises && pal.remises.length === TYPO_COLS.length) {
        TYPO_COLS.forEach((t, idx) => { ws.getCell(`${t.param}${blk.remiseRow}`).value = pal.remises![idx]; });
      } else if (pal.remiseStandard && typeof pal.remiseStandardTaux === "number") {
        for (const t of TYPO_COLS) ws.getCell(`${t.param}${blk.remiseRow}`).value = pal.remiseStandardTaux;
      }
    }

    // Recâbler les "Nb Offres" (bug gabarit : pointent tous sur $B$4).
    // ATTENTION : remplacer "$B$4" uniquement quand il n'est PAS suivi d'un chiffre, sinon
    // "$B$42" (bloc 2) deviendrait "$B$422". On utilise (?!\d).
    const bOffres = `$B$${blk.nbOffresRow}`;
    for (const t of TYPO_COLS) {
      const cell = ws.getCell(`${t.param}${blk.nbOffRow}`);
      const v: any = cell.value;
      if (v && typeof v === "object" && typeof v.formula === "string") cell.value = { formula: v.formula.replace(/\$B\$4(?!\d)/g, bOffres) };
    }

    // VLOOKUP sur TOUTES les lignes produits (même vides) : ainsi, si l'utilisateur tape une
    // réf dans une ligne vide, libellé/EAN/prix/PPC se remplissent automatiquement depuis le
    // Mapping. B/C/F/H/J sont donc toujours des formules VLOOKUP basées sur la colonne A.
    // On parcourt TOUTE la zone d'articles du bloc (pvFirst..dataLast), pas seulement les
    // 13 premières lignes : sinon les articles au-delà (UG, testeurs, PLV…) étaient perdus.
    const nbLignes = blk.dataLast - blk.pvFirst + 1;
    for (let i = 0; i < nbLignes; i++) {
      const row = blk.pvFirst + i;
      const p = pal ? pal.produits[i] : undefined;
      const ref = p ? (p.ref || "").trim() : "";
      // Réf manuelle hors catalogue (productId 0 et absente du Mapping) → valeurs en dur.
      const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);

      setRefText(ws, row, ref);                                          // A : code en TEXTE
      const vente = estVente(p);
      if (horsMapping) {
        ws.getCell(row, 2).value = p!.name || "";
        ws.getCell(row, 3).value = p!.barcode || "";
        ws.getCell(row, 6).value = round2(p!.standardPrice || 0);
        ws.getCell(row, 8).value = vente ? round2(p!.listPrice || 0) : 0;   // prix vente = 0 si gratuit
        ws.getCell(row, 10).value = vente ? round2(p!.ppc || 0) : 0;        // PPC = 0 si gratuit
      } else {
        // VLOOKUP (renvoie "" si A vide → ligne vide propre, mais remplissable).
        ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
        ws.getCell(row, 3).value = { formula: vlookup(`A${row}`, 3) };
        ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
        // Prix vente / PPC : VLOOKUP si Produit Vente, sinon 0 en dur (UG/Testeur/PLV gratuits).
        ws.getCell(row, 8).value = vente ? { formula: vlookup(`A${row}`, 5) } : 0;
        ws.getCell(row, 10).value = vente ? { formula: vlookup(`A${row}`, 6) } : 0;
      }
      // D : Typ. Prod — vidé si la ligne n'a pas d'article (évite les résidus du gabarit).
      ws.getCell(row, 4).value = p ? (p.typProd || "Produit Vente") : null;
      ws.getCell(row, 5).value = p ? (p.qtyParPack || 0) : null;          // E : qté/pack
      // I : Remise additionnelle du palier (si renseignée) appliquée à tous les produits.
      if (p && estOpca(p.typProd)) {
        // OPCA : panier virtuel hors catalogue → valeurs en dur. Coût = 55 % du seuil en formule
        // (suit le seuil si on le modifie dans Excel), pas de PPC, jamais de remise additionnelle.
        ws.getCell(row, 2).value = p.name || "";
        ws.getCell(row, 3).value = "";
        ws.getCell(row, 6).value = { formula: `ROUND(H${row}*${OPCA_TAUX_COUT},2)` };
        ws.getCell(row, 8).value = round2(p.listPrice || 0);
        ws.getCell(row, 10).value = 0;
        ws.getCell(row, 9).value = 0;
      } else if (pal && typeof pal.remiseAddTaux === "number") ws.getCell(row, 9).value = pal.remiseAddTaux;
      ws.getCell(row, 7).value = { formula: `E${row}*F${row}` };               // G : montant achat
      ws.getCell(row, 11).value = { formula: `J${row}*(1-I${row})` };           // K : PPC remisé
      ws.getCell(row, 12).value = { formula: `IFERROR(J${row}-K${row},"")` };   // L : Montant BRI
      // CA/Marges : toujours en formules (donnent 0 si E vide), pour rester remplissables.
      writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);
      // Masquer les lignes VIDES (sans réf) pour alléger le fichier ; les lignes remplies
      // restent visibles. La formule/format reste en place (démasquable dans Excel).
      ws.getRow(row).hidden = !ref;
    }

    // « Poids Gratuités achats » (UG/PLV/Échantillon/Testeur, colonnes AC/AD) : le gabarit
    // fait pointer le bloc 2 sur les lignes du bloc 1. On réécrit la formule sur les lignes
    // réelles de CHAQUE bloc (y compris blocs ajoutés et lignes d'articles insérées).
    for (let r = blk.title; r <= blk.dataLast; r++) {
      const lbl = ws.getCell(r, 29).value;
      if (typeof lbl === "string" && ["UG", "PLV", "Echantillon", "Échantillon", "Testeur"].includes(lbl.trim())) {
        ws.getCell(r, 30).value = { formula: `IFERROR(SUMIF($D$${blk.pvFirst}:$D$${blk.dataLast},AC${r},$G$${blk.pvFirst}:$G$${blk.dataLast})/$AC$${blk.synRow},0)` };
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

  // GRANDS COMPTES : bloc en L149 (template vierge). Code (valeur) + libellé/prix VLOOKUP.
  const pal1 = paliers[0];

  // Le bloc GC du gabarit a 20 lignes d'articles (159-178). Au-delà, les références en trop
  // étaient perdues dans GC, Non B2B et la liste logistique : on insère les lignes manquantes
  // (formules recopiées de la dernière ligne) et on étend les SUM/SUMIF qui couvraient 159:178.
  const nbArticlesGc = pal1?.produits?.length || 0;
  const extraGc = Math.max(0, nbArticlesGc - (gcPos.last - gcPos.pvFirst + 1));
  if (extraGc > 0) {
    unshareFormulas(ws);
    const srcRow = gcPos.last, oldLast = gcPos.last, at = gcPos.last + 1;
    const src = ws.getRow(srcRow);
    const srcCells: { value: any; style: any }[] = [];
    for (let c = 1; c <= 40; c++) {
      const cell = src.getCell(c); const v: any = cell.value;
      srcCells.push({ value: v && typeof v === "object" && v.formula ? { formula: v.formula } : v, style: JSON.parse(JSON.stringify(cell.style || {})) });
    }
    ws.spliceRows(at, 0, ...Array.from({ length: extraGc }, () => [] as any[]));
    shiftFormulaRefs(ws, at, extraGc);
    for (let k = 1; k <= extraGc; k++) {
      const dst = ws.getRow(srcRow + k);
      dst.height = src.height;
      srcCells.forEach((sc, ci) => {
        const cell = dst.getCell(ci + 1);
        cell.style = JSON.parse(JSON.stringify(sc.style));
        cell.value = sc.value && typeof sc.value === "object" && typeof sc.value.formula === "string"
          ? { formula: sc.value.formula.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (m: string, col: string, rs: string) => +rs === srcRow ? `${col}${srcRow + k}` : m) }
          : (sc.value ?? null);
      });
    }
    const newLast = oldLast + extraGc;
    const plage = new RegExp(`(\\$?[A-Z]{1,3}\\$?)${gcPos.pvFirst}:(\\$?[A-Z]{1,3}\\$?)${oldLast}(?!\\d)`, "g");
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
      const v: any = cell.value;
      if (v && typeof v === "object" && typeof v.formula === "string" && plage.test(v.formula)) {
        plage.lastIndex = 0;
        cell.value = { formula: v.formula.replace(plage, `$1${gcPos.pvFirst}:$2${newLast}`) };
      }
      plage.lastIndex = 0;
    }));
    gcPos.last = newLast;
    logFirst += extraGc;
  }

  // GC_PV_COUNT = toutes les lignes d'articles du bloc (plus seulement les 13 "Produit Vente").
  const GC_PV_FIRST = gcPos.pvFirst, GC_PV_COUNT = gcPos.last - gcPos.pvFirst + 1, LOG_PV_FIRST = logFirst, LOG_PV_COUNT = 13;
  const GC_NOM_ROW = gcPos.nom, GC_REMISE_ROW = gcPos.remise;
  // 6 colonnes GC fixes dans le template (M/P/S/V/Y/AB = cols 13/16/19/22/25/28).
  const GC_ENSEIGNE_COLS = [
    { nom: 13, qte: 13, remise: 15 },
    { nom: 16, qte: 16, remise: 18 },
    { nom: 19, qte: 19, remise: 21 },
    { nom: 22, qte: 22, remise: 24 },
    { nom: 25, qte: 25, remise: 27 },
    { nom: 28, qte: 28, remise: 30 },
  ];
  // GC 7+ : colonnes insérées à la suite de NewPharma (col 31+), 3 colonnes chacune
  // (Qtités/CA/Marges) avec les styles copiés de la colonne NewPharma. Le bloc
  // "Grands Comptes Total" + "Poids Gratuités" (cols 31-34 du gabarit) est réécrit
  // décalé d'autant vers la droite, formules ajustées.
  const allGcEnseignes = payload.gcEnseignes || [];
  const gcExtras = allGcEnseignes.slice(6);
  const nbExtras = gcExtras.length;
  const GC_EXTRA_COLS = gcExtras.map((_, i) => ({ nom: 31 + i * 3, qte: 31 + i * 3, remise: 33 + i * 3 }));
  const ALL_GC_COLS = [...GC_ENSEIGNE_COLS, ...GC_EXTRA_COLS];
  const GC_ROW_TOP = gcPos.top, GC_ROW_LAST = gcPos.last, GC_SYN_ROW = gcPos.syn;
  const colL = (n: number): string => { let s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; };
  const cloneStyle = (st: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> => JSON.parse(JSON.stringify(st || {}));

  // Snapshot des styles du bloc "Total" du gabarit (cols 31-34 = spacer/titre/CA/Marges),
  // pris AVANT toute modification. Réutilisé pour le Total GC décalé ET le Total non B2B.
  const totSnap: Record<string, Partial<ExcelJS.Style>> = {};
  for (let r = gcPos.top - 1; r <= gcPos.top + 21; r++) for (let c = 31; c <= 34; c++) {
    totSnap[`${r}:${c}`] = cloneStyle(ws.getCell(r, c).style);
  }

  if (nbExtras > 0) {
    const shift = nbExtras * 3;
    // 1) Effacer l'ancienne position du bloc Total (cols 31-34) — sera réécrit décalé.
    for (let r = gcPos.top - 1; r <= gcPos.top + 21; r++) for (let c = 31; c <= 34; c++) {
      const cell = ws.getCell(r, c); cell.value = null; cell.style = {} as ExcelJS.Style;
    }
    // 2) Styles des colonnes extra copiés de NewPharma (cols 28/29/30) + largeurs.
    for (let i = 0; i < nbExtras; i++) {
      const base = 31 + i * 3;
      for (let k = 0; k < 3; k++) {
        for (let r = GC_ROW_TOP; r <= GC_ROW_LAST; r++) {
          ws.getCell(r, base + k).style = cloneStyle(ws.getCell(r, 28 + k).style) as ExcelJS.Style;
        }
        const w = ws.getColumn(28 + k).width;
        if (w) ws.getColumn(base + k).width = w;
      }
    }
    // Bandeau catégorie (ligne 150) « Autre GC » fusionné sur toutes les colonnes extra.
    ws.mergeCells(gcPos.band, 31, gcPos.band, 30 + shift);
    const cat = ws.getCell(gcPos.band, 31);
    cat.value = "Autre GC";
    cat.style = cloneStyle(ws.getCell(gcPos.band, 25).style) as ExcelJS.Style;
    // 3) En-têtes + formules de chaque colonne extra (mêmes formules que les 6 fixes).
    for (const c of GC_EXTRA_COLS) {
      const qL = colL(c.qte), caL = colL(c.qte + 1), mgL = colL(c.qte + 2);
      ws.getCell(GC_REMISE_ROW, c.qte).value = "Remise";
      ws.getCell(gcPos.hdr, c.qte).value = "Qtités"; ws.getCell(gcPos.hdr, c.qte + 1).value = "CA"; ws.getCell(gcPos.hdr, c.qte + 2).value = "Marges";
      ws.getCell(gcPos.unit, c.qte).value = "'#"; ws.getCell(gcPos.unit, c.qte + 1).value = "'€"; ws.getCell(gcPos.unit, c.qte + 2).value = "'€";
      ws.getCell(GC_SYN_ROW, c.qte).value = { formula: `SUM(${qL}${GC_PV_FIRST}:${qL}${GC_ROW_LAST})` };
      ws.getCell(GC_SYN_ROW, c.qte + 1).value = { formula: `SUM(${caL}${GC_PV_FIRST}:${caL}${GC_ROW_LAST})` };
      ws.getCell(GC_SYN_ROW, c.qte + 2).value = { formula: `SUM(${mgL}${GC_PV_FIRST}:${mgL}${GC_ROW_LAST})` };
      for (let r = GC_PV_FIRST; r <= GC_ROW_LAST; r++) {
        ws.getCell(r, c.qte + 1).value = { formula: `${qL}${r}*H${r}*(1-I${r})*(1-$${colL(c.remise)}$${GC_REMISE_ROW})` };
        ws.getCell(r, c.qte + 2).value = { formula: `${caL}${r}-(F${r}*${qL}${r})` };
      }
    }
    // 4) Col E (Qtités totale) : somme de toutes les colonnes Qtités. On remplace la formule
    //    partagée du gabarit sur TOUTE la plage pour ne laisser aucun clone orphelin.
    const qteLetters = ALL_GC_COLS.map(c => colL(c.qte));
    for (let r = GC_PV_FIRST; r <= GC_ROW_LAST; r++) {
      ws.getCell(r, 5).value = { formula: qteLetters.map(L => `${L}${r}`).join("+") };
    }
    // 5) Bloc "Grands Comptes Total" + "Poids Gratuités" réécrit à sa nouvelle position.
    for (let r = gcPos.top - 1; r <= gcPos.top + 21; r++) for (let c = 31; c <= 34; c++) {
      const st = totSnap[`${r}:${c}`];
      if (st) ws.getCell(r, c + shift).style = st as ExcelJS.Style;
    }
    const T1 = 32 + shift, T2 = 33 + shift, T3 = 34 + shift;
    const caTL = colL(T2), mgTL = colL(T3);
    ws.getCell(gcPos.top, T1).value = "Grands Comptes";
    ws.getCell(gcPos.top + 2, T1).value = "Total";
    ws.getCell(gcPos.hdr, T2).value = "CA"; ws.getCell(gcPos.hdr, T3).value = "Marges";
    ws.getCell(gcPos.unit, T2).value = "'€"; ws.getCell(gcPos.unit, T3).value = "'€";
    ws.getCell(GC_SYN_ROW, T2).value = { formula: ALL_GC_COLS.map(c => `${colL(c.qte + 1)}${GC_SYN_ROW}`).join("+") };
    ws.getCell(GC_SYN_ROW, T3).value = { formula: ALL_GC_COLS.map(c => `${colL(c.qte + 2)}${GC_SYN_ROW}`).join("+") };
    ws.getCell(gcPos.top + 12, T2).value = "Poids Gratuités achats :";
    ws.getCell(gcPos.top + 12, T3).value = { formula: `SUM(${mgTL}${gcPos.top + 14}:${mgTL}${gcPos.top + 17})` };
    ["UG", "PLV", "Echantillon", "Testeur"].forEach((lbl, gi) => {
      const r = gcPos.top + 14 + gi;
      ws.getCell(r, T2).value = lbl;
      ws.getCell(r, T3).value = { formula: `SUMIF($D$${GC_PV_FIRST}:$D$${GC_ROW_LAST},${caTL}${r},$G$${GC_PV_FIRST}:$G$${GC_ROW_LAST})/$${caTL}$${GC_SYN_ROW}` };
    });
    // Largeurs des colonnes CA/Marges du Total (évite "###").
    ws.getColumn(T2).width = 13;
    ws.getColumn(T3).width = 13;
  }

  // En-têtes enseignes : nom (ligne 152) + taux de remise (ligne 153) pour toutes.
  ALL_GC_COLS.forEach((cols, idx) => {
    const ens = allGcEnseignes[idx];
    if (!ens) return;
    if (ens.nom) ws.getCell(GC_NOM_ROW, cols.nom).value = ens.nom;
    if (typeof ens.remise === "number") ws.getCell(GC_REMISE_ROW, cols.remise).value = ens.remise;
  });
  // Clé GC : réf seule si unique dans le palier, sinon "ref#type" (doublons vendu/UG).
  const gcRefCount: Record<string, number> = {};
  for (const pr of pal1?.produits || []) { const r = (pr.ref || "").trim(); if (r) gcRefCount[r] = (gcRefCount[r] || 0) + 1; }
  const gcKey = (pr?: PropProduit): string => {
    const r = (pr?.ref || "").trim();
    return (r && gcRefCount[r] > 1) ? `${r}#${pr!.typProd || "Produit Vente"}` : r;
  };
  for (let i = 0; i < GC_PV_COUNT; i++) {
    const row = GC_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    const ref = p ? (p.ref || "").trim() : "";
    const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);
    setRefText(ws, row, ref);
    const venteGC = estVente(p);
    if (horsMapping) {
      ws.getCell(row, 2).value = p!.name || "";
      ws.getCell(row, 6).value = round2(p!.standardPrice || 0);
      ws.getCell(row, 8).value = venteGC ? round2(p!.listPrice || 0) : 0;
      ws.getCell(row, 10).value = venteGC ? round2(p!.ppc || 0) : 0;
    } else {
      ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
      ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
      ws.getCell(row, 8).value = venteGC ? { formula: vlookup(`A${row}`, 5) } : 0;
      ws.getCell(row, 10).value = venteGC ? { formula: vlookup(`A${row}`, 6) } : 0;
    }
    // Grands Comptes : qté pour toutes les enseignes (formules CA/Marges des extras déjà posées).
    ws.getCell(row, 4).value = p ? (p.typProd || "Produit Vente") : null;   // D : sert au « Poids gratuités »
    if (!p) for (const c of [2, 6, 8, 10]) ws.getCell(row, c).value = null;
    else ws.getCell(row, 11).value = { formula: `J${row}*(1-I${row})` };      // K : absent du gabarit après la ligne 13
    const gk = gcKey(p);
    if (gk && allGcEnseignes.length) {
      ALL_GC_COLS.forEach((cols, idx) => {
        const ens = allGcEnseignes[idx];
        if (!ens) return;
        const q = ens.qties[gk];
        if (q != null) {
          const cell = ws.getCell(row, cols.qte);
          cell.value = q;
          cell.numFmt = "0";
        }
      });
    }
  }
  // Élargir les colonnes CA/Marges de toutes les enseignes GC pour éviter "########".
  for (const cols of ALL_GC_COLS) {
    ws.getColumn(cols.qte + 1).width = 13;
    ws.getColumn(cols.qte + 2).width = 13;
  }


  // Zone « Besoins logistiques » du gabarit (lignes 185-197) : entièrement vidée ici.
  // Elle n'a ni titre ni en-tête et fait doublon avec l'onglet « Synthèse logistique » ;
  // sa liste réf/libellé est réécrite À LA FIN, après le bloc non B2B (voir plus bas).
  for (let i = 0; i < LOG_PV_COUNT; i++) {
    const row = LOG_PV_FIRST + i;
    ws.getCell(row, 1).value = null; ws.getCell(row, 2).value = null;
  }
  // Vider PLV/Testeurs fixes du gabarit (positions du template vierge : GC 172-178, log 192-197).
  for (let row = LOG_PV_FIRST + LOG_PV_COUNT - 6; row < LOG_PV_FIRST + LOG_PV_COUNT; row++) for (const c of [1, 2, 4]) ws.getCell(row, c).value = null;

  // ── BESOINS NON B2B : NOUVEAU TABLEAU juste sous le bloc GRANDS COMPTES.
  //    Réplique de la structure GC : titre, bandeau, noms/remises des canaux, en-têtes,
  //    synthèse, lignes articles avec VLOOKUP et formules CA/Marges.
  const nonB2B = payload.canauxNonB2B || [];
  let logistiqueRow = LOG_PV_FIRST; // position de la liste logistique (recalculée si non B2B)
  if (nonB2B.length && pal1?.produits?.length) {
    const OFF = (GC_ROW_LAST + 3) - gcPos.top; // titre du bloc = 3 lignes sous la fin du bloc GC (178)
    const NB_TITLE = gcPos.top + OFF, NB_BAND = gcPos.band + OFF, NB_NOM = GC_NOM_ROW + OFF, NB_REM = GC_REMISE_ROW + OFF;
    const NB_HDR = gcPos.hdr + OFF, NB_UNIT = gcPos.unit + OFF, NB_SYN = GC_SYN_ROW + OFF;
    const NB_FIRST = GC_PV_FIRST + OFF, NB_LAST = GC_ROW_LAST + OFF;
    const NB_COLS = nonB2B.map((_, i) => ({ qte: 13 + i * 3, remise: 15 + i * 3 }));
    const lastCanalCol = 12 + nonB2B.length * 3;

    // 1) Styles : recopie ligne à ligne du bloc GC (cols A..L + groupe BIOCOOP par canal).
    for (let r = GC_ROW_TOP; r <= GC_ROW_LAST; r++) {
      for (let c = 1; c <= 12; c++) ws.getCell(r + OFF, c).style = cloneStyle(ws.getCell(r, c).style) as ExcelJS.Style;
      for (let i = 0; i < nonB2B.length; i++) for (let k = 0; k < 3; k++) {
        ws.getCell(r + OFF, 13 + i * 3 + k).style = cloneStyle(ws.getCell(r, 13 + k).style) as ExcelJS.Style;
      }
    }

    // 2) Titre + bandeau.
    ws.getCell(NB_TITLE, 1).value = "BESOINS NON B2B";
    ws.mergeCells(NB_BAND, 13, NB_BAND, lastCanalCol);
    const band = ws.getCell(NB_BAND, 13);
    band.value = "NON B2B";
    band.style = cloneStyle(ws.getCell(gcPos.band, 25).style) as ExcelJS.Style;

    // 3) En-têtes fixes (libellés colonnes A..L identiques au bloc GC).
    for (let c = 1; c <= 12; c++) {
      const v154 = ws.getCell(gcPos.hdr, c).value, v155 = ws.getCell(gcPos.unit, c).value;
      if (v154 != null && typeof v154 !== "object") ws.getCell(NB_HDR, c).value = v154;
      if (v155 != null && typeof v155 !== "object") ws.getCell(NB_UNIT, c).value = v155;
    }
    ws.getCell(NB_SYN, 1).value = "Synthèse";

    // 4) Canaux : nom, remise, en-têtes, synthèse, formules par ligne.
    NB_COLS.forEach((c, idx) => {
      const ens = nonB2B[idx];
      const qL = colL(c.qte), caL = colL(c.qte + 1), mgL = colL(c.qte + 2);
      ws.getCell(NB_NOM, c.qte).value = ens.nom;
      ws.getCell(NB_REM, c.qte).value = "Remise";
      ws.getCell(NB_REM, c.remise).value = typeof ens.remise === "number" ? ens.remise : 0;
      ws.getCell(NB_HDR, c.qte).value = "Qtités"; ws.getCell(NB_HDR, c.qte + 1).value = "CA"; ws.getCell(NB_HDR, c.qte + 2).value = "Marges";
      ws.getCell(NB_UNIT, c.qte).value = "'#"; ws.getCell(NB_UNIT, c.qte + 1).value = "'€"; ws.getCell(NB_UNIT, c.qte + 2).value = "'€";
      ws.getCell(NB_SYN, c.qte).value = { formula: `SUM(${qL}${NB_FIRST}:${qL}${NB_LAST})` };
      ws.getCell(NB_SYN, c.qte + 1).value = { formula: `SUM(${caL}${NB_FIRST}:${caL}${NB_LAST})` };
      ws.getCell(NB_SYN, c.qte + 2).value = { formula: `SUM(${mgL}${NB_FIRST}:${mgL}${NB_LAST})` };
      for (let r = NB_FIRST; r <= NB_LAST; r++) {
        ws.getCell(r, c.qte + 1).value = { formula: `${qL}${r}*H${r}*(1-I${r})*(1-$${colL(c.remise)}$${NB_REM})` };
        ws.getCell(r, c.qte + 2).value = { formula: `${caL}${r}-(F${r}*${qL}${r})` };
      }
      ws.getColumn(c.qte + 1).width = 13; ws.getColumn(c.qte + 2).width = 13;
    });

    // 5) Lignes articles : réf + libellé/prix (VLOOKUP ou valeurs) + qtés par canal +
    //    Qtités totale (col E) + PPC remisé / BRI (cols K/L) comme le gabarit.
    const qteLettersNb = NB_COLS.map(c => colL(c.qte));
    for (let i = 0; i < GC_PV_COUNT; i++) {
      const row = NB_FIRST + i, p = pal1.produits[i];
      const ref = p ? (p.ref || "").trim() : "";
      if (p) {
        const horsMapping = p.productId === 0 && !mapRefs.has(ref);
        setRefText(ws, row, ref);
        const vente = estVente(p);
        if (horsMapping) {
          ws.getCell(row, 2).value = p.name || "";
          ws.getCell(row, 6).value = round2(p.standardPrice || 0);
          ws.getCell(row, 8).value = vente ? round2(p.listPrice || 0) : 0;
          ws.getCell(row, 10).value = vente ? round2(p.ppc || 0) : 0;
        } else {
          ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
          ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
          ws.getCell(row, 8).value = vente ? { formula: vlookup(`A${row}`, 5) } : 0;
          ws.getCell(row, 10).value = vente ? { formula: vlookup(`A${row}`, 6) } : 0;
        }
        ws.getCell(row, 11).value = { formula: `J${row}*(1-I${row})` };
        ws.getCell(row, 12).value = { formula: `IFERROR(J${row}-K${row},"")` };
      }
      ws.getCell(row, 5).value = { formula: qteLettersNb.map(L => `${L}${row}`).join("+") };
      // Quantités saisies par canal (même clé composite que les GC).
      const gk = gcKey(p);
      if (gk) NB_COLS.forEach((c, idx) => {
        const q = nonB2B[idx].qties[gk];
        if (q != null) { const cell = ws.getCell(row, c.qte); cell.value = q; cell.numFmt = "0"; }
      });
    }
    ws.getCell(NB_SYN, 5).value = { formula: `SUM(E${NB_FIRST}:E${NB_LAST})` };

    // 6) Bloc « Total non B2B » à droite du tableau (styles du bloc Total du gabarit).
    const TB = lastCanalCol + 2;
    for (let r = gcPos.top - 1; r <= gcPos.top + 9; r++) for (let c = 31; c <= 34; c++) {
      const st = totSnap[`${r}:${c}`];
      if (st) ws.getCell(r + OFF, c - 31 + TB).style = cloneStyle(st) as ExcelJS.Style;
    }
    ws.getCell(gcPos.top + OFF, TB + 1).value = "Non B2B";
    ws.getCell(gcPos.top + 2 + OFF, TB + 1).value = "Total";
    ws.getCell(gcPos.hdr + OFF, TB + 2).value = "CA"; ws.getCell(gcPos.hdr + OFF, TB + 3).value = "Marges";
    ws.getCell(gcPos.unit + OFF, TB + 2).value = "'€"; ws.getCell(gcPos.unit + OFF, TB + 3).value = "'€";
    ws.getCell(NB_SYN, TB + 2).value = { formula: NB_COLS.map(c => `${colL(c.qte + 1)}${NB_SYN}`).join("+") };
    ws.getCell(NB_SYN, TB + 3).value = { formula: NB_COLS.map(c => `${colL(c.qte + 2)}${NB_SYN}`).join("+") };
    ws.getColumn(TB + 2).width = 13; ws.getColumn(TB + 3).width = 13;
    // La liste logistique passe APRÈS ce bloc (3 lignes sous sa dernière ligne d'article).
    logistiqueRow = NB_LAST + 3;
  }

  // ── Liste « Besoins logistiques » (réf + libellé), écrite en DERNIER, tout en bas.
  //    Le détail par mois est dans l'onglet dédié « Synthèse logistique » ; ici on ne
  //    garde que le rappel des références, avec un titre pour qu'il ne soit plus orphelin.
  // OPCA exclues : ce ne sont pas des produits à approvisionner.
  const produitsLog = (pal1?.produits || []).filter(p => !estOpca(p.typProd));
  if (produitsLog.length) {
    ws.getCell(logistiqueRow - 1, 1).value = "Besoins logistiques — références";
    ws.getCell(logistiqueRow - 1, 1).font = { bold: true, size: 11 };
    for (let i = 0; i < Math.max(LOG_PV_COUNT, produitsLog.length); i++) {
      const row = logistiqueRow + i, p = produitsLog[i];
      if (!p) { ws.getCell(row, 1).value = null; ws.getCell(row, 2).value = null; continue; }
      const ref = (p.ref || "").trim();
      const horsMapping = p.productId === 0 && !mapRefs.has(ref);
      setRefText(ws, row, ref);
      ws.getCell(row, 2).value = horsMapping ? (p.name || "") : { formula: vlookup(`A${row}`, 2) };
    }
  }
  return blocks;
}

// Remplit une feuille Proposition. mapRefs = réfs présentes dans le Mapping partagé (pour que
// les VLOOKUP fonctionnent). Si omis, aucun Mapping (valeurs en dur uniquement).
export function fillPropositionSheet(ws: ExcelJS.Worksheet, payload: PropPayload, mapRefs: Set<string> = new Set()) {
  fillProposition(ws, payload, mapRefs);
}
