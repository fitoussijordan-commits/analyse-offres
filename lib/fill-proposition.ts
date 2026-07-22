// lib/fill-proposition.ts — Remplissage d'une feuille "Proposition" à partir d'une campagne.
// Logique partagée entre l'export d'une campagne (/api/export-template) et l'export
// multi-campagnes (/api/export-multi). NE restructure PAS la feuille : écrit seulement les
// valeurs dans les lignes produits des 3 blocs + GRANDS COMPTES + Besoins logistiques.

import type ExcelJS from "exceljs";

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
export interface PropPayload { nom: string; paliers: PropPalier[]; mapping?: MapRow[]; gcEnseignes?: GcEnseignePayload[]; }

export const PROP_SHEET = "Proposition template";

/** Un produit génère du CA seulement s'il est "Produit Vente". UG / Testeur / PLV /
 *  Échantillon sont gratuits → prix de vente (col 8) et PPC (col 10) forcés à 0. */
function estVente(p?: { typProd?: string }): boolean {
  return !p || (p.typProd || "Produit Vente") === "Produit Vente";
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
  const propRanges = BLOCKS.map(b => ({ first: b.pvFirst, last: b.dataLast, nbOffresCell: `'${PROP_SHEET}'!$B$${b.nbOffresRow}` }));

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

function fillProposition(ws: ExcelJS.Worksheet, payload: PropPayload, mapRefs: Set<string>) {
  const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
  const used = Math.min(paliers.length, BLOCKS.length);

  for (let b = 0; b < BLOCKS.length; b++) {
    const blk = BLOCKS[b];
    const pal = b < used ? paliers[b] : null;

    if (pal) {
      const titre = [payload.nom, pal.code, pal.label].map(s => (s || "").trim()).filter(Boolean).join(" — ") || "Offre";
      ws.getCell(blk.title, 1).value = titre;
      // Descriptif libre du palier → colonne D de la ligne titre (à côté du nom Premium/Standard).
      if (pal.descriptif != null) ws.getCell(blk.title, 4).value = pal.descriptif.trim();
      ws.getCell(blk.nbOffresRow, 2).value = pal.qtyPacks || 0;
      // Nombre de produits = SUM des "Pdt dans offre" (colonne E) des lignes Produit Vente.
      // Formule (pas une valeur) pour rester cohérent si on édite une qté à la main.
      const pvLast = blk.pvFirst + blk.pvCount - 1;
      ws.getCell(blk.nbProduitsRow, 2).value = { formula: `SUM(E${blk.pvFirst}:E${pvLast})` };

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
    for (let i = 0; i < blk.pvCount; i++) {
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
      if (p && p.typProd) ws.getCell(row, 4).value = p.typProd;            // D : Typ. Prod
      ws.getCell(row, 5).value = p ? (p.qtyParPack || 0) : null;          // E : qté/pack
      // I : Remise additionnelle du palier (si renseignée) appliquée à tous les produits.
      if (pal && typeof pal.remiseAddTaux === "number") ws.getCell(row, 9).value = pal.remiseAddTaux;
      // CA/Marges : toujours en formules (donnent 0 si E vide), pour rester remplissables.
      writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);
      // Masquer les lignes Produit Vente VIDES (sans réf) pour alléger le fichier ; les lignes
      // remplies restent visibles. La formule/format reste en place (démasquable dans Excel).
      ws.getRow(row).hidden = !ref;
    }

    // Lignes PLV / Testeurs / SR (au-delà des Produit Vente, jusqu'à dataLast).
    // RÈGLE : ne garder QUE les réfs présentes dans la campagne (pal.produits). Toute réf
    // pré-remplie du gabarit (PANNEAU REGE, PRESENTOIR REGE, SR REGE RETAIL 1…) absente de la
    // campagne est VIDÉE (réf, libellé, prix, quantités) et la ligne masquée — sinon on
    // exportait des références d'une ancienne campagne, ce qui n'a aucun sens.
    const refsCampagne = new Set((pal?.produits || []).map(p => (p.ref || "").trim()).filter(Boolean));
    for (let row = blk.pvFirst + blk.pvCount; row <= blk.dataLast; row++) {
      const cur = ws.getCell(row, 1).value;
      const ref = cur == null ? "" : String(cur).trim();
      if (!ref) continue;
      // Résidu du gabarit (réf absente de la campagne) → on vide entièrement la ligne.
      if (!refsCampagne.has(ref)) {
        for (let col = 1; col <= 30; col++) ws.getCell(row, col).value = null;
        ws.getRow(row).hidden = true;
        continue;
      }
      setRefText(ws, row, ref);
      ws.getCell(row, 2).value = { formula: vlookup(`A${row}`, 2) };
      ws.getCell(row, 3).value = { formula: vlookup(`A${row}`, 3) };
      ws.getCell(row, 6).value = { formula: vlookup(`A${row}`, 4) };
      // Ces lignes sont des PLV / Testeurs / SR (non "Produit Vente") → prix vente + PPC = 0.
      ws.getCell(row, 8).value = 0;
      ws.getCell(row, 10).value = 0;
      // Remise additionnelle du palier (col I) aussi sur ces lignes.
      if (pal && typeof pal.remiseAddTaux === "number") ws.getCell(row, 9).value = pal.remiseAddTaux;
      ws.getCell(row, 11).value = { formula: `J${row}*(1-I${row})` };           // K : PPC remisé
      ws.getCell(row, 12).value = { formula: `IFERROR(J${row}-K${row},"")` };   // L : Montant BRI
      writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);                  // M..Z : CA/Marges
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
  const GC_PV_FIRST = 159, GC_PV_COUNT = 13, LOG_PV_FIRST = 185, LOG_PV_COUNT = 13;
  const GC_NOM_ROW = 152, GC_REMISE_ROW = 153;
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
  const GC_ROW_TOP = 149, GC_ROW_LAST = 178, GC_SYN_ROW = 157;
  const colL = (n: number): string => { let s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - 1 - r) / 26; } return s; };
  const cloneStyle = (st: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> => JSON.parse(JSON.stringify(st || {}));

  if (nbExtras > 0) {
    const shift = nbExtras * 3;
    // 1) Snapshot des styles du bloc Total du gabarit (cols 31-34 = spacer/titre/CA/Marges).
    const totSnap: Record<string, Partial<ExcelJS.Style>> = {};
    for (let r = 148; r <= 170; r++) for (let c = 31; c <= 34; c++) {
      const cell = ws.getCell(r, c);
      totSnap[`${r}:${c}`] = cloneStyle(cell.style);
      cell.value = null;
      cell.style = {} as ExcelJS.Style;
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
    ws.mergeCells(150, 31, 150, 30 + shift);
    const cat = ws.getCell(150, 31);
    cat.value = "Autre GC";
    cat.style = cloneStyle(ws.getCell(150, 25).style) as ExcelJS.Style;
    // 3) En-têtes + formules de chaque colonne extra (mêmes formules que les 6 fixes).
    for (const c of GC_EXTRA_COLS) {
      const qL = colL(c.qte), caL = colL(c.qte + 1), mgL = colL(c.qte + 2);
      ws.getCell(GC_REMISE_ROW, c.qte).value = "Remise";
      ws.getCell(154, c.qte).value = "Qtités"; ws.getCell(154, c.qte + 1).value = "CA"; ws.getCell(154, c.qte + 2).value = "Marges";
      ws.getCell(155, c.qte).value = "'#"; ws.getCell(155, c.qte + 1).value = "'€"; ws.getCell(155, c.qte + 2).value = "'€";
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
    for (let r = 148; r <= 170; r++) for (let c = 31; c <= 34; c++) {
      const st = totSnap[`${r}:${c}`];
      if (st) ws.getCell(r, c + shift).style = st as ExcelJS.Style;
    }
    const T1 = 32 + shift, T2 = 33 + shift, T3 = 34 + shift;
    const caTL = colL(T2), mgTL = colL(T3);
    ws.getCell(149, T1).value = "Grands Comptes";
    ws.getCell(151, T1).value = "Total";
    ws.getCell(154, T2).value = "CA"; ws.getCell(154, T3).value = "Marges";
    ws.getCell(155, T2).value = "'€"; ws.getCell(155, T3).value = "'€";
    ws.getCell(GC_SYN_ROW, T2).value = { formula: ALL_GC_COLS.map(c => `${colL(c.qte + 1)}${GC_SYN_ROW}`).join("+") };
    ws.getCell(GC_SYN_ROW, T3).value = { formula: ALL_GC_COLS.map(c => `${colL(c.qte + 2)}${GC_SYN_ROW}`).join("+") };
    ws.getCell(161, T2).value = "Poids Gratuités achats :";
    ws.getCell(161, T3).value = { formula: `SUM(${mgTL}163:${mgTL}166)` };
    ["UG", "PLV", "Echantillon", "Testeur"].forEach((lbl, gi) => {
      const r = 163 + gi;
      ws.getCell(r, T2).value = lbl;
      ws.getCell(r, T3).value = { formula: `SUMIF($D$159:$D$178,${caTL}${r},$G$159:$G$178)/$${caTL}$${GC_SYN_ROW}` };
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
  // Besoins logistiques (palier 1) : code + libellé en VLOOKUP sur toutes les lignes.
  for (let i = 0; i < LOG_PV_COUNT; i++) {
    const row = LOG_PV_FIRST + i, p = pal1 ? pal1.produits[i] : undefined;
    const ref = p ? (p.ref || "").trim() : "";
    const horsMapping = !!p && p.productId === 0 && !mapRefs.has(ref);
    setRefText(ws, row, ref);
    ws.getCell(row, 2).value = horsMapping ? (p!.name || "") : { formula: vlookup(`A${row}`, 2) };
  }
  // Vider PLV/Testeurs fixes du gabarit (positions du template vierge : GC 172-178, log 192-197).
  for (const row of [172, 173, 174, 175, 176, 177, 178]) for (const c of [1, 2, 4, 6, 8, 10]) ws.getCell(row, c).value = null;
  for (const row of [192, 193, 194, 195, 196, 197]) for (const c of [1, 2, 4]) ws.getCell(row, c).value = null;
}


// Remplit une feuille Proposition. mapRefs = réfs présentes dans le Mapping partagé (pour que
// les VLOOKUP fonctionnent). Si omis, aucun Mapping (valeurs en dur uniquement).
export function fillPropositionSheet(ws: ExcelJS.Worksheet, payload: PropPayload, mapRefs: Set<string> = new Set()) {
  fillProposition(ws, payload, mapRefs);
}
