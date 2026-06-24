// app/api/export-template/route.ts
// Remplit le template "Proposition" SANS le restructurer.
//
// Principe (non destructif) :
//  - On part du gabarit nettoyé proposition-template-clean.xlsx, identique à l'original
//    (3 blocs REGENERANTS 1/2/3, section GRANDS COMPTES, mise en page, formats, formules),
//    mais sans les onglets parasites ni les liens externes morts qui corrompaient le fichier.
//  - On NE touche NI à la structure, NI à GRANDS COMPTES, NI aux PLV/Testeurs.
//  - Pour chaque palier de la préco (max 3, mappés sur les 3 blocs existants), on écrit
//    uniquement dans les lignes "Produit Vente" du bloc :
//       A = code article (préco)   B = libellé (Odoo)   C = EAN (Odoo)
//       E = qté/pack (préco)        F = coût achat (Odoo) H = tarif revendeur (Odoo)
//       J = PPC (Odoo)
//    Les colonnes calculées (G, K, L, CA/Marges M..Z) restent les formules du gabarit.
//  - Les lignes "Produit Vente" en trop (produit non fourni) sont vidées de leur code/valeurs
//    pour ne pas laisser les anciens articles du template.

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";

export const maxDuration = 60;

const SHEET = "Proposition template";
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template-clean.xlsx");

// Géométrie FIXE des 3 blocs du gabarit (lignes 1-based). Les lignes "Produit Vente"
// sont les 13 premières lignes data de chaque bloc ; au-delà : PLV / Testeurs (intouchés).
// pctRow / remiseRow / nbOffRow : lignes des paramètres typologie de CE bloc (elles ne sont
// PAS à offset constant — le bloc 1 a une ligne de moins que les blocs 2 et 3).
// nbOffresCell : cellule "Nombre d'offres" du bloc (B<nbOffresRow>), pour recâbler les
// formules "Nb Offres" qui, dans le gabarit d'origine, pointaient toutes (à tort) sur B4.
// synRow : ligne "Synthèse" du bloc. dataFirst/dataLast : plage des lignes data (pour les SUM).
const BLOCKS = [
  { title: 3,  nbOffresRow: 4,  nbProduitsRow: 5,  pvFirst: 17, pvCount: 13, pctRow: 9,  remiseRow: 10, nbOffRow: 11, synRow: 15, dataFirst: 17, dataLast: 36 },  // REGENERANTS 1
  { title: 40, nbOffresRow: 42, nbProduitsRow: 43, pvFirst: 55, pvCount: 13, pctRow: 47, remiseRow: 48, nbOffRow: 49, synRow: 53, dataFirst: 55, dataLast: 74 },  // REGENERANTS 2
  { title: 76, nbOffresRow: 78, nbProduitsRow: 79, pvFirst: 91, pvCount: 13, pctRow: 83, remiseRow: 84, nbOffRow: 85, synRow: 89, dataFirst: 91, dataLast: 110 }, // REGENERANTS 3
];

interface PrecoLigneIn {
  ref: string; name: string; productId: number;
  qtyParPack: number;
  barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
}
interface PrecoPalierIn {
  code: string; label: string; qtyPacks: number;
  produits: PrecoLigneIn[];
}
interface PayloadIn { nom: string; paliers: PrecoPalierIn[]; }

function round2(n: number): number { return Math.round((n + Number.EPSILON) * 100) / 100; }

// Mapping des 7 typologies clients → colonnes (lettres) du template.
// Chaque typologie = une paire (CA, Marges). Les paramètres %Offres/Remise/NbOffres sont
// dans la colonne de DROITE de la paire, aux lignes (titre+9) %Offres, (titre+10) Remise,
// (titre+11) Nb Offres. Ex. bloc1 (titre L3) : Ambassadeur CA=M, Marges=N, params en N9/N10/N11.
const TYPO_COLS: Array<{ ca: string; marge: string; param: string }> = [
  { ca: "M", marge: "N", param: "N" }, // Ambassadeur
  { ca: "O", marge: "P", param: "P" }, // Compagnon
  { ca: "Q", marge: "R", param: "R" }, // Challenger
  { ca: "S", marge: "T", param: "T" }, // Rose
  { ca: "U", marge: "V", param: "V" }, // Prunelier
  { ca: "W", marge: "X", param: "X" }, // Anthylide
  { ca: "Y", marge: "Z", param: "Z" }, // Calendula
];
const FMT_EUR = '#,##0.0 "€";(#,##0.0) "€";" - "';

// Écrit les formules CA + Marges de toutes les typologies sur une ligne produit donnée.
// CA   = E × Tarif × (1 − remise add.) × Nb Offres typologie × (1 − Remise typologie)
// Marge = CA − (E × Coût achat × Nb Offres typologie)
function writeTypoFormulas(ws: ExcelJS.Worksheet, row: number, remiseRow: number, nbOffRow: number) {
  for (const t of TYPO_COLS) {
    const nbOff = `$${t.param}$${nbOffRow}`;
    const rem = `$${t.param}$${remiseRow}`;
    const caCell = ws.getCell(`${t.ca}${row}`);
    const mgCell = ws.getCell(`${t.marge}${row}`);
    caCell.value = { formula: `E${row}*H${row}*(1-I${row})*${nbOff}*(1-${rem})` };
    mgCell.value = { formula: `${t.ca}${row}-E${row}*F${row}*${nbOff}` };
    caCell.numFmt = FMT_EUR;
    mgCell.numFmt = FMT_EUR;
  }
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

    // On mappe au plus 3 paliers sur les 3 blocs existants (on ne crée/supprime aucun bloc).
    const used = Math.min(paliers.length, BLOCKS.length);

    for (let b = 0; b < BLOCKS.length; b++) {
      const blk = BLOCKS[b];
      const pal = b < used ? paliers[b] : null;

      if (pal) {
        // Titre du bloc : "<nom campagne> — <code> — <label>" (le nom de campagne est
        // reporté sur chaque palier ; sans toucher au reste de la ligne).
        const titreParts = [payload.nom, pal.code, pal.label].map(s => (s || "").trim()).filter(Boolean);
        ws.getCell(blk.title, 1).value = titreParts.join(" — ") || "Offre";
        // Nombre d'offres (B) = nb de packs vendus du palier.
        ws.getCell(blk.nbOffresRow, 2).value = pal.qtyPacks || 0;
        // Nombre de produits (B) = somme des qté/pack des produits.
        ws.getCell(blk.nbProduitsRow, 2).value = pal.produits.reduce((s, p) => s + (p.qtyParPack || 0), 0);
      }

      // Correction d'un bug du gabarit d'origine : les formules "Nb Offres" (ligne nbOffRow)
      // de TOUS les blocs pointent sur $B$4 (nb offres du bloc 1). On les recâble sur le
      // "Nombre d'offres" du bloc courant (B<nbOffresRow>), pour chaque typologie.
      const bOffresCell = `$B$${blk.nbOffresRow}`;
      for (const t of TYPO_COLS) {
        const cell = ws.getCell(`${t.param}${blk.nbOffRow}`);
        const v: any = cell.value;
        if (v && typeof v === "object" && typeof v.formula === "string") {
          cell.value = { formula: v.formula.replace(/\$B\$4/g, bOffresCell) };
        }
      }

      // Remplir / vider les 13 lignes "Produit Vente" du bloc.
      for (let i = 0; i < blk.pvCount; i++) {
        const row = blk.pvFirst + i;
        const p = pal ? pal.produits[i] : undefined;

        if (p) {
          ws.getCell(row, 1).value = p.ref || "";                 // A : Code article
          ws.getCell(row, 2).value = p.name || "";                // B : Libellé (Odoo)
          ws.getCell(row, 3).value = p.barcode || "";             // C : EAN (Odoo)
          ws.getCell(row, 5).value = p.qtyParPack || 0;           // E : Pdt dans offre (qté/pack)
          ws.getCell(row, 6).value = round2(p.standardPrice || 0);// F : Coût achat unitaire
          ws.getCell(row, 8).value = round2(p.listPrice || 0);    // H : Tarif revendeur unitaire
          ws.getCell(row, 10).value = round2(p.ppc || 0);         // J : PPC
          // CA + Marges pour les 7 typologies (le gabarit ne calcule que la colonne M).
          writeTypoFormulas(ws, row, blk.remiseRow, blk.nbOffRow);
          // D (typ. prod), G/K/L (formules) : laissés tels quels.
        } else {
          // Pas de produit pour cette ligne (ou bloc non utilisé) → vider le code et les
          // valeurs saisies du gabarit, en gardant le style. On vide aussi les CA/Marges
          // pour ne pas laisser des montants calculés sur une ligne sans produit.
          ws.getCell(row, 1).value = null;  // A
          ws.getCell(row, 2).value = null;  // B
          ws.getCell(row, 3).value = null;  // C
          ws.getCell(row, 5).value = null;  // E
          ws.getCell(row, 6).value = null;  // F
          ws.getCell(row, 8).value = null;  // H
          ws.getCell(row, 10).value = null; // J
          for (const t of TYPO_COLS) {
            ws.getCell(`${t.ca}${row}`).value = null;
            ws.getCell(`${t.marge}${row}`).value = null;
          }
        }
      }

      // Ligne de synthèse : SUM par colonne CA et Marges de toutes les typologies, sur la
      // plage data du bloc. Le gabarit ne fournit la SUM que pour la colonne M (Ambassadeur) ;
      // on ajoute les colonnes manquantes pour que AC (CA total) et AD (Marge total) soient
      // justes. (AC/AD restent les formules d'origine, qui somment ces colonnes.)
      for (const t of TYPO_COLS) {
        const caSyn = ws.getCell(`${t.ca}${blk.synRow}`);
        const mgSyn = ws.getCell(`${t.marge}${blk.synRow}`);
        caSyn.value = { formula: `SUM(${t.ca}${blk.dataFirst}:${t.ca}${blk.dataLast})` };
        mgSyn.value = { formula: `SUM(${t.marge}${blk.dataFirst}:${t.marge}${blk.dataLast})` };
        caSyn.numFmt = FMT_EUR;
        mgSyn.numFmt = FMT_EUR;
      }
      // Synthèse, colonnes de droite (col 28=AB, 29=AC, 30=AD) :
      //   AB = "Nbrs Offres planifiées" → NOMBRE, surtout pas des euros.
      //   AC = CA total  → euros.
      //   AD = Marge totale → euros (le gabarit l'avait en %, hérité de l'ancienne marge %).
      // ATTENTION : dans le gabarit, AB et AC partagent le MÊME objet style (ExcelJS), donc
      // changer le format de l'un change l'autre. On casse le partage en clonant le style
      // d'AB avant de lui appliquer son propre format nombre.
      const abCell = ws.getCell(blk.synRow, 28);
      abCell.style = JSON.parse(JSON.stringify(abCell.style || {}));
      ws.getCell(blk.synRow, 29).numFmt = FMT_EUR;   // AC : CA total (euros)
      ws.getCell(blk.synRow, 30).numFmt = FMT_EUR;   // AD : Marge totale (euros)
      abCell.numFmt = '#,##0;(#,##0);" - "';          // AB : nombre (après clonage)
    }

    // ── GRANDS COMPTES & Besoins logistiques ────────────────────────────────────
    // Ces deux sections reprennent la même liste de produits que le bloc 1 (mêmes codes,
    // mêmes positions ; elles référencent d'ailleurs le bloc 1). On les aligne donc sur
    // le PALIER 1. On remplit uniquement les colonnes demandées, sans toucher au reste :
    //   - GRANDS COMPTES : libellé (B), coût achat (F), tarif revendeur (H), PPC (J).
    //     (K = PPC remisé et L = BRI sont déjà des formules du gabarit → laissées telles quelles.)
    //   - Besoins logistiques : libellé (B) seulement.
    const pal1 = paliers[0]; // bloc 1 = palier 1
    const GC_PV_FIRST = 123, GC_PV_COUNT = 13;   // lignes "Produit Vente" de GRANDS COMPTES
    const LOG_PV_FIRST = 149, LOG_PV_COUNT = 13; // lignes "Produit Vente" de Besoins logistiques

    for (let i = 0; i < GC_PV_COUNT; i++) {
      const row = GC_PV_FIRST + i;
      const p = pal1 ? pal1.produits[i] : undefined;
      if (p) {
        ws.getCell(row, 2).value = p.name || "";                 // B : Libellé
        ws.getCell(row, 6).value = round2(p.standardPrice || 0); // F : Coût achat unitaire
        ws.getCell(row, 8).value = round2(p.listPrice || 0);     // H : Tarif revendeur unitaire
        ws.getCell(row, 10).value = round2(p.ppc || 0);          // J : PPC
      } else {
        ws.getCell(row, 2).value = null;
        ws.getCell(row, 6).value = null;
        ws.getCell(row, 8).value = null;
        ws.getCell(row, 10).value = null;
      }
    }

    for (let i = 0; i < LOG_PV_COUNT; i++) {
      const row = LOG_PV_FIRST + i;
      const p = pal1 ? pal1.produits[i] : undefined;
      ws.getCell(row, 2).value = p ? (p.name || "") : null;      // B : Libellé
    }

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
