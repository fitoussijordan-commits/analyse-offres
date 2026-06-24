// app/api/export-template/route.ts
// Remplit le template "Proposition template" (Fichier trade 2027.xlsx) à partir de la
// préco N+1 enrichie des données tarifaires Odoo (EAN, coût achat, tarif revendeur, PPC).
//
// Principe :
//  - On part du gabarit lib/templates/proposition-template.xlsx (mise en forme conservée).
//  - L'onglet "Proposition template" contient 1 bloc-modèle (REGENERANTS 1, lignes 3→36).
//    On le réplique autant de fois qu'il y a de paliers dans la préco.
//  - Pour chaque produit on écrit : code article (A), libellé (B), EAN (C), typ. prod (D),
//    qté/pack (E), coût achat (F), tarif revendeur (H), PPC (J) — en VALEURS (issues d'Odoo).
//  - Les colonnes calculées (G, K, L, M…Z, et la Synthèse) restent en FORMULES, mais on
//    répare les références cassées du gabarit ([1]…, #REF!) pour qu'elles tournent en
//    autonomie dans le classeur.
//
// La grille de répartition par typologie (M..Z : %Offres / Remises) reprend les valeurs
// par défaut du gabarit (bloc 1), conformément au choix utilisateur.

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";

export const maxDuration = 60;

const SHEET = "Proposition template";
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template.xlsx");

// ── Géométrie du bloc-modèle (bloc 1 du gabarit) ──────────────────────────────
// Toutes les lignes sont exprimées en offset relatif au titre du bloc (ligne "REGENERANTS X").
const TITLE = 0;          // L3  : titre
const NB_OFFRES = 1;      // L4  : "Nombre d'offres" | B = valeur
const NB_PRODUITS = 2;    // L5  : "Nombre de produits" | B = valeur
const RETAIL_HDR = 3;     // L6  : RETAIL / INSTITUT / Retail+Institut
const TYPO_HDR = 5;       // L8  : Ambassadeur..Calendula | Total
const PCT_OFFRES = 6;     // L9  : "% Offres" + valeurs N9.. + Z9
const REMISE = 7;         // L10 : "Remise" + valeurs
const NB_OFFRES_ROW = 8;  // L11 : "Nb Offres" + formules =N9*$B$<nbOffres>
const COL_HDR1 = 9;       // L12 : Pdt dans offre, Coût achat…  (+ AB12 Nbrs Offres planifiées)
const COL_HDR2 = 10;      // L13 : Code article, Libellé… (+ unités '# '€)
const SYNTHESE = 12;      // L15 : Synthèse + formules de totaux
const DATA_START = 14;    // L17 : 1re ligne produit
const DATA_LEN = 20;      // 20 lignes produits max par bloc (L17→L36)
const BLOCK_LEN = 38;     // hauteur totale d'un bloc (titre L3 → fin marge avant L40)

interface PrecoLigneIn {
  ref: string; name: string; productId: number;
  qtyParPack: number; conserve: boolean;
  // pricing Odoo
  barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
  typProd?: string; // "Produit Vente" | "PLV" | "Testeur" | "UG" | "Echantillon"...
}
interface PrecoPalierIn {
  code: string; label: string; qtyPacks: number;
  produits: PrecoLigneIn[];
}
interface PayloadIn {
  nom: string;
  paliers: PrecoPalierIn[];
}

// Remplace dans une chaîne de formule les références cassées du gabarit.
// - Liens externes [1]'Paramètres Articles' / [1]'New Base PA' : on ne devrait plus les
//   rencontrer (colonnes B/C/J réécrites en valeurs), mais on neutralise par sécurité.
// - #REF! résiduels : on neutralise pour éviter une erreur Excel à l'ouverture.
function repairFormula(f: string): string | null {
  if (f.includes("[1]") || f.includes("#REF!")) return null; // → sera remplacé par valeur/blank
  return f;
}

export async function POST(req: NextRequest) {
  try {
    const payload: PayloadIn = await req.json();
    const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
    if (!paliers.length) {
      return NextResponse.json({ error: "Aucun palier à exporter" }, { status: 400 });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_PATH);
    const ws = wb.getWorksheet(SHEET);
    if (!ws) return NextResponse.json({ error: `Onglet "${SHEET}" introuvable dans le gabarit` }, { status: 500 });

    // 1) Capturer le bloc-modèle (bloc 1, lignes 3..3+BLOCK_LEN) : styles + valeurs + formules.
    const modelTitleRow = 3;
    const model: Array<Array<{ value: any; style: any; numFmt?: string }>> = [];
    for (let off = 0; off <= BLOCK_LEN; off++) {
      const srcRow = ws.getRow(modelTitleRow + off);
      const rowCells: Array<{ value: any; style: any; numFmt?: string }> = [];
      for (let c = 1; c <= 35; c++) {
        const cell = srcRow.getCell(c);
        rowCells.push({
          value: cell.value,
          style: JSON.parse(JSON.stringify(cell.style || {})),
          numFmt: cell.numFmt,
        });
      }
      model.push(rowCells);
    }
    // Capturer les merges du bloc-modèle (relatifs).
    const modelMerges: Array<[number, number, number, number]> = [];
    const mergeRanges: string[] = (ws as any).model?.merges || [];
    for (const m of mergeRanges) {
      const dec = decodeRange(m);
      if (!dec) continue;
      const { top, left, bottom, right } = dec;
      if (top >= modelTitleRow && bottom <= modelTitleRow + BLOCK_LEN) {
        modelMerges.push([top - modelTitleRow, left, bottom - modelTitleRow, right]);
      }
    }

    // 2) Effacer tout le contenu existant de l'onglet (on régénère intégralement),
    //    en conservant largeurs de colonnes.
    const lastRow = ws.rowCount;
    // unmerge tout
    for (const m of [...mergeRanges]) { try { ws.unMergeCells(m); } catch { /* ignore */ } }
    for (let r = 1; r <= lastRow; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 35; c++) { const cell = row.getCell(c); cell.value = null; }
    }

    // 3) Réécrire un bloc par palier.
    let cursor = 1; // 1re ligne du 1er bloc (= ancien L1 → on commence à 1 comme l'original commençait à 3 ; on garde 2 lignes de marge)
    cursor = 3;
    paliers.forEach((pal, idx) => {
      const titleRow = cursor;
      const nbOffres = pal.qtyPacks || 0;
      const bOffresCell = `$B$${titleRow + NB_OFFRES}`;       // réf "Nombre d'offres" de CE bloc

      for (let off = 0; off <= BLOCK_LEN; off++) {
        const destRow = ws.getRow(titleRow + off);
        for (let c = 1; c <= 35; c++) {
          const m = model[off][c - 1];
          const cell = destRow.getCell(c);
          // style + format toujours recopiés
          cell.style = JSON.parse(JSON.stringify(m.style || {}));
          if (m.numFmt) cell.numFmt = m.numFmt;

          let v = m.value;
          // Réécriture des valeurs dynamiques selon la ligne logique (off) :
          if (off === TITLE && c === 1) v = `${pal.code} — ${pal.label || "Offre"}`;
          else if (off === NB_OFFRES && c === 2) v = nbOffres;
          else if (off === NB_PRODUITS && c === 2) v = pal.produits.reduce((s, p) => s + (p.qtyParPack || 0), 0);
          else if (off === NB_OFFRES_ROW) {
            // "Nb Offres" = %Offres * Nombre d'offres du bloc → recâbler sur bOffresCell
            if (typeof v === "object" && v?.formula) {
              const fixed = String(v.formula).replace(/\$B\$4/g, bOffresCell);
              v = { formula: fixed };
            }
          } else if (off === SYNTHESE) {
            v = remapSynthese(v, titleRow, c);
          } else if (off >= DATA_START && off < DATA_START + DATA_LEN) {
            v = buildDataCell(off - DATA_START, c, pal.produits, titleRow, m.value);
          } else if (typeof v === "object" && v?.formula) {
            const fixed = remapRelativeFormula(String(v.formula), titleRow, modelTitleRow);
            v = fixed === null ? null : { formula: fixed };
          }
          cell.value = v === undefined ? null : v;
        }
      }
      // ré-appliquer les merges du bloc
      for (const [t, l, b, r] of modelMerges) {
        try { ws.mergeCells(titleRow + t, l, titleRow + b, r); } catch { /* ignore */ }
      }
      cursor = titleRow + BLOCK_LEN + 2; // 2 lignes de marge entre blocs (comme le gabarit : L36→L40)
    });

    const buf = await wb.xlsx.writeBuffer();
    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="proposition_${(payload.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx"`,
      },
    });
  } catch (e: any) {
    console.error("Export template error:", e);
    return NextResponse.json({ error: e?.message || "Erreur export template" }, { status: 500 });
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

// Mapping des 7 typologies clients → colonnes du template.
// Chaque typologie occupe une paire (CA, Marges). Les paramètres %Offres / Remise / NbOffres
// sont dans la colonne de DROITE de la paire (ex. Ambassadeur : CA=M, Marges=N, params en N).
//   caCol    : colonne où écrire la formule de CA
//   margeCol : colonne où écrire la formule de Marge
//   paramCol : colonne des paramètres ($<paramCol>$<ligne Remise>, $<paramCol>$<ligne NbOffres>)
const TYPO_COLS: Array<{ caCol: number; margeCol: number; paramCol: string }> = [
  { caCol: 13, margeCol: 14, paramCol: "N" }, // Ambassadeur : M / N
  { caCol: 15, margeCol: 16, paramCol: "P" }, // Compagnon   : O / P
  { caCol: 17, margeCol: 18, paramCol: "R" }, // Challenger  : Q / R
  { caCol: 19, margeCol: 20, paramCol: "T" }, // Rose        : S / T
  { caCol: 21, margeCol: 22, paramCol: "V" }, // Prunelier   : U / V
  { caCol: 23, margeCol: 24, paramCol: "X" }, // Anthylide   : W / X
  { caCol: 25, margeCol: 26, paramCol: "Z" }, // Calendula   : Y / Z
];

// Construit une cellule de ligne produit (off-relatif 0..DATA_LEN-1).
function buildDataCell(
  prodIdx: number, col: number, produits: PrecoLigneIn[],
  titleRow: number, modelValue: any
): any {
  const absRow = titleRow + DATA_START + prodIdx;
  const p = produits[prodIdx];

  // Au-delà de la liste produits → ligne vide (mais on garde le style).
  if (!p) return null;

  // Lignes de paramètres de CE bloc (offsets relatifs au titre).
  const remiseRow = titleRow + REMISE;          // ligne "Remise"
  const nbOffresRow = titleRow + NB_OFFRES_ROW; // ligne "Nb Offres"

  switch (col) {
    case 1:  return p.ref || "";                                  // A : Code article
    case 2:  return p.name || "";                                 // B : Libellé (valeur Odoo)
    case 3:  return p.barcode || "";                              // C : EAN (valeur Odoo)
    case 4:  return p.typProd || "Produit Vente";                 // D : Typ. Prod
    case 5:  return p.qtyParPack || 0;                            // E : Pdt dans offre (qté/pack)
    case 6:  return round2(p.standardPrice || 0);                 // F : Coût achat unitaire (valeur)
    case 7:  return { formula: `E${absRow}*F${absRow}` };         // G : Coûts achats total
    case 8:  return round2(p.listPrice || 0);                     // H : Tarif revendeur unitaire (valeur)
    case 9:  return typeof modelValue === "number" ? modelValue : 0.15; // I : Remise additionnelle
    case 10: return round2(p.ppc || 0);                           // J : PPC (valeur Odoo)
    case 11: return { formula: `J${absRow}*(1-I${absRow})` };     // K : PPC remisé
    case 12: return { formula: `IFERROR(J${absRow}-K${absRow},"")` }; // L : Montant BRI
    default: {
      // Colonnes M..Z : pour chaque typologie, CA puis Marges.
      const typo = TYPO_COLS.find(t => t.caCol === col || t.margeCol === col);
      if (!typo) return null;
      const nbOff = `$${typo.paramCol}$${nbOffresRow}`;  // Nb Offres de la typologie
      const rem = `$${typo.paramCol}$${remiseRow}`;      // Remise de la typologie
      if (col === typo.caCol) {
        // CA = E × Tarif revendeur × (1 − remise add.) × Nb Offres × (1 − Remise typologie)
        //   E = quantité du produit dans le pack (colonne "Pdt dans offre").
        return { formula: `E${absRow}*H${absRow}*(1-I${absRow})*${nbOff}*(1-${rem})` };
      }
      // Marges € = CA − coût d'achat écoulé = CA − (E × Coût achat unitaire × Nb Offres typologie).
      const caRef = colNumToLetter(typo.caCol) + absRow;
      return { formula: `${caRef}-E${absRow}*F${absRow}*${nbOff}` };
    }
  }
}

function colNumToLetter(n: number): string {
  let s = ""; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s;
}

// Remappe une formule du bloc-modèle (titre L3) vers le bloc courant (titre titleRow).
// Décale toutes les références de lignes de (titleRow - 3) et neutralise [1]/#REF!.
function remapRelativeFormula(formula: string, titleRow: number, modelTitleRow: number): string | null {
  if (formula.includes("[1]") || formula.includes("#REF!")) return null;
  const delta = titleRow - modelTitleRow;
  if (delta === 0) return formula;
  // Décale les références $?LET$?NUM (en évitant de toucher aux noms de feuille)
  return shiftRowRefs(formula, delta);
}

// Synthèse (ligne SYNTHESE) : agrège chaque colonne sur les lignes data du bloc.
// - Colonnes CA et Marges de chaque typologie (M..Z) → SUM sur la plage data, même
//   si le gabarit ne contenait la formule que pour Ambassadeur.
// - Autres cellules (E, G, J, K, L, AB, AC, AD) : on décale les formules du gabarit
//   et on neutralise les #REF!/[1].
function remapSynthese(v: any, titleRow: number, col: number): any {
  const dataFirst = titleRow + DATA_START;
  const dataLast = titleRow + DATA_START + DATA_LEN - 1;
  const L = (n: number) => colNumToLetter(n);

  // Toute colonne CA/Marges d'une typologie → SUM de sa plage data.
  const isTypoCol = TYPO_COLS.some(t => t.caCol === col || t.margeCol === col);
  if (isTypoCol) {
    const c = L(col);
    return { formula: `SUM(${c}${dataFirst}:${c}${dataLast})` };
  }

  if (typeof v !== "object" || !v?.formula) return v;
  const f = String(v.formula);
  if (f.includes("[1]") || f.includes("#REF!")) return null;
  const fixed = shiftRowRefs(f, titleRow - 3);
  return { formula: fixed };
}

// Décale les numéros de ligne dans les références de cellules d'une formule.
// Gère les formes A1, $A$1, A1:B2, et préserve les références à d'autres feuilles
// (REGENERANTS!$C..) en ne décalant pas après un "!".
function shiftRowRefs(formula: string, delta: number): string {
  // On ne décale pas les réfs qui suivent un "!" (autre feuille).
  // Approche : tokeniser en repérant les segments "Feuille!..." pour les laisser intacts.
  return formula.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (match, colPart, numPart, offset, full) => {
    // si juste avant le match il y a un "!", c'est une réf vers une autre feuille → ne pas toucher
    const prevChar = offset > 0 ? full[offset - 1] : "";
    if (prevChar === "!") return match;
    const n = parseInt(numPart, 10) + delta;
    return `${colPart}${n}`;
  });
}

function decodeRange(a1: string): { top: number; left: number; bottom: number; right: number } | null {
  const m = a1.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return null;
  return { left: colToNum(m[1]), top: parseInt(m[2], 10), right: colToNum(m[3]), bottom: parseInt(m[4], 10) };
}
function colToNum(s: string): number {
  let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n;
}
function round2(n: number): number { return Math.round((n + Number.EPSILON) * 100) / 100; }
