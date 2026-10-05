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
import { qtyParPack, totalPacks, ventilationPalier, articlesPalier, qtyPalierOpca, qtyKeyLib } from "@/lib/create-campaign";
import { estOpca } from "@/lib/type-produit";

export const MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

/** Prix d'achat unitaire par réf pour l'export : catalogue Odoo (Mapping), complété par les
 *  prix saisis dans les campagnes (réfs hors Odoo). Une réf de campagne prime sur le catalogue. */
export function coutsAchatParRef(mapping: { ref: string; standardPrice?: number }[] | undefined, paliers: { produits: { ref: string; standardPrice?: number; typProd?: string }[] }[], prixComposants: Record<string, number> = {}): Record<string, number> {
  const out: Record<string, number> = {};
  // Ordre de priorité croissante : catalogue Odoo < prix des composants de kits < articles de campagne.
  for (const m of mapping || []) { const r = (m.ref || "").trim(); if (r && (m.standardPrice || 0) > 0) out[r] = m.standardPrice!; }
  for (const [r, c] of Object.entries(prixComposants)) if (c > 0) out[r] = c;
  for (const pal of paliers) for (const p of pal.produits || []) {
    const r = (p.ref || "").trim();
    if (r && p.typProd !== "OPCA" && (p.standardPrice || 0) > 0) out[r] = p.standardPrice!;
  }
  return out;
}

/** Lettre de colonne Excel (1 → A, 27 → AA). */
function col(n: number): string { let t = ""; while (n > 0) { const r = (n - 1) % 26; t = String.fromCharCode(65 + r) + t; n = (n - 1 - r) / 26; } return t; }

const FMT_UNITES = "#,##0;-#,##0;\"\"";
const FMT_EUR = "#,##0 \"€\";-#,##0 \"€\";\"\"";
const FMT_PRIX = "#,##0.00 \"€\"";

/** Prix d'achat unitaire d'une réf : null pour un kit (assemblé, coût porté par ses
 *  composants) ou si le prix est inconnu (à compléter à la main dans Excel). */
function prixAchat(ref: string, coutByRef: Record<string, number>, kits: Set<string>): number | null {
  if (kits.has(ref)) return null;
  const c = coutByRef[ref];
  return typeof c === "number" && c > 0 ? Math.round(c * 100) / 100 : null;
}

/** Écrit un onglet "Synthèse logistique" (réf × mois) dans le classeur. Partagé entre les
 *  exports simple et multi. nameByRef : libellés de secours ; coutByRef : prix d'achat
 *  unitaire (coût Odoo) par réf.
 *  Mise en page calquée sur le planning achats : quantités par mois, puis montant d'achat
 *  par mois (prix × qté), « Total réf », et en bas « Total mois » + « Total année ». */
export function writeSyntheseLogistiqueSheet(wb: ExcelJS.Workbook, log: SyntheseLogistique, nameByRef: Record<string, string> = {}, coutByRef: Record<string, number> = {}) {
  const existant = wb.getWorksheet("Synthèse logistique");
  if (existant) wb.removeWorksheet(existant.id);
  const sw = wb.addWorksheet("Synthèse logistique", { views: [{ state: "frozen", xSplit: 3, ySplit: 4, showGridLines: false }] });
  ecrireTableauLogistique(sw, log, {
    titre: "Synthèse besoins logistiques — par référence et par mois",
    colonnesTexte: [{ titre: "Réf", largeur: 16 }, { titre: "Produit", largeur: 40 }],
    lignes: log.lignes.map(l => ({ textes: [l.ref, l.name || nameByRef[l.ref] || ""], ref: l.ref, parMois: l.parMois, total: l.total })),
    notesSurvol: true,
    filtres: false,
  }, coutByRef);
  writeDetailLogistiqueSheet(wb, log, nameByRef, coutByRef);
}

/** Onglet « Détail logistique » : une ligne par référence × campagne × offre, même mise en
 *  page que la synthèse. Filtres automatiques : les totaux du bas suivent le filtre
 *  (ex. achats d'une seule campagne). */
function writeDetailLogistiqueSheet(wb: ExcelJS.Workbook, log: SyntheseLogistique, nameByRef: Record<string, string>, coutByRef: Record<string, number>) {
  const NOM = "Détail logistique";
  const existant = wb.getWorksheet(NOM);
  if (existant) wb.removeWorksheet(existant.id);
  const detail = log.detail || [];
  if (!detail.length) return;
  // Ordre : références dans l'ordre de la synthèse (plus gros besoin d'abord), puis offres.
  const rang = new Map(log.lignes.map((l, i) => [l.ref, i]));
  const tri = [...detail].sort((a, b) => ((rang.get(a.ref) ?? 1e9) - (rang.get(b.ref) ?? 1e9)) || (b.total - a.total));
  const ws = wb.addWorksheet(NOM, { views: [{ state: "frozen", xSplit: 5, ySplit: 4, showGridLines: false }] });
  ecrireTableauLogistique(ws, log, {
    titre: "Détail des besoins logistiques — par référence, campagne et offre",
    colonnesTexte: [{ titre: "Réf", largeur: 14 }, { titre: "Produit", largeur: 36 }, { titre: "Campagne", largeur: 28 }, { titre: "Offre", largeur: 28 }],
    lignes: tri.map(d => ({ textes: [d.ref, d.name || nameByRef[d.ref] || "", d.campagne, d.palier], ref: d.ref, parMois: d.parMois, total: d.total })),
    notesSurvol: false,
    filtres: true,
  }, coutByRef);
}

interface LigneTableau { textes: string[]; ref: string; parMois: number[]; total: number; }

/**
 * Tableau commun aux deux onglets :
 *   [textes…] | Prix achat | Quantités : mois… | Total qté | Achats € : mois… | Total réf €
 * Achat du mois = prix × qté du mois (formule : un prix complété à la main met tout à jour).
 * Lignes de pied : TOTAL UNITÉS, TOTAL MOIS (€), TOTAL ANNÉE. Avec filtres, les pieds
 * utilisent SUBTOTAL : ils ne comptent que les lignes visibles.
 */
function ecrireTableauLogistique(ws: ExcelJS.Worksheet, log: SyntheseLogistique, o: {
  titre: string; colonnesTexte: { titre: string; largeur: number }[]; lignes: LigneTableau[];
  notesSurvol: boolean; filtres: boolean;
}, coutByRef: Record<string, number>) {
  const TEAL = "0D9488", TEAL_DARK = "0F766E", BLEU = "1F4E79", BLEU_SOFT = "EAF1F8", DARK = "1A1A2E", WHITE = "FFFFFF";
  const mois = log.moisLabels && log.moisLabels.length ? log.moisLabels : MOIS_FR;
  const kits = new Set(log.refsKit || []);
  const nT = o.colonnesTexte.length, nM = mois.length;
  const C_PRIX = nT + 1, C_Q1 = nT + 2, C_QTOT = C_Q1 + nM, C_E1 = C_QTOT + 1, C_ETOT = C_E1 + nM;
  const nbCols = C_ETOT;
  ws.columns = [...o.colonnesTexte.map(c => ({ width: c.largeur })), { width: 10 }, ...mois.map(() => ({ width: 10 })), { width: 11 }, ...mois.map(() => ({ width: 12 })), { width: 14 }];
  const remplir = (cell: ExcelJS.Cell, couleur: string) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + couleur } }; };

  // Titre (+ alerte campagnes ignorées)
  const titre = ws.addRow([o.titre]);
  ws.mergeCells(titre.number, 1, titre.number, nbCols);
  const tc = ws.getCell(titre.number, 1);
  tc.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
  remplir(tc, TEAL); tc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  titre.height = 24;
  const avert = log.ignorees && log.ignorees.length
    ? `⚠ Campagne(s) NON comptée(s), dates de début/fin manquantes ou invalides : ${log.ignorees.join(", ")}` : "";
  const ligneAvert = ws.addRow([avert]);
  if (avert) {
    ws.mergeCells(ligneAvert.number, 1, ligneAvert.number, nbCols);
    const wc = ws.getCell(ligneAvert.number, 1);
    wc.font = { bold: true, size: 10, color: { argb: "FFB91C1C" }, name: "Calibri" };
    remplir(wc, "FEE2E2"); wc.alignment = { wrapText: true, vertical: "middle", indent: 1 };
    ligneAvert.height = 28;
  }

  // Bandeau de groupes : Quantités | Achats €
  const bande = ws.addRow([]);
  ws.mergeCells(bande.number, C_Q1, bande.number, C_QTOT);
  ws.mergeCells(bande.number, C_E1, bande.number, C_ETOT);
  const bq = ws.getCell(bande.number, C_Q1), be = ws.getCell(bande.number, C_E1);
  bq.value = "Quantités (unités)"; be.value = "Achats (€) = prix × quantité";
  for (const [c, coul] of [[bq, TEAL], [be, BLEU]] as const) {
    c.font = { bold: true, size: 10, color: { argb: "FF" + WHITE }, name: "Calibri" };
    remplir(c, coul); c.alignment = { horizontal: "center", vertical: "middle" };
  }
  const head = ws.addRow([...o.colonnesTexte.map(c => c.titre), "Prix achat", ...mois, "Total", ...mois, "Total réf"]);
  head.height = 30;
  head.eachCell((c, ci) => {
    c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" };
    remplir(c, ci >= C_E1 ? BLEU : TEAL);
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });

  // Notes de survol (synthèse) : offres qui composent chaque case de quantité.
  const detailParRef = new Map<string, LigneDetailLogistique[]>();
  for (const d of log.detail || []) {
    if (!detailParRef.has(d.ref)) detailParRef.set(d.ref, []);
    detailParRef.get(d.ref)!.push(d);
  }
  const fmt = (n: number) => Math.round(n).toLocaleString("fr-FR");
  const noteOffres = (parts: { d: LigneDetailLogistique; q: number }[]) => parts
    .filter(p => p.q > 0).sort((a, b) => b.q - a.q)
    .map(p => `${p.d.campagne} / ${p.d.palier} : ${fmt(p.q)}`).join("\n");

  const premiere = ws.rowCount + 1;
  let prixManquants = 0, refPrec = "", bandeau = false;
  const refsSansPrix = new Set<string>();
  for (const l of o.lignes) {
    if (l.ref !== refPrec) { bandeau = !bandeau; refPrec = l.ref; }
    const estKit = kits.has(l.ref);
    const prix = prixAchat(l.ref, coutByRef, kits);
    if (prix == null && !estKit) refsSansPrix.add(l.ref);
    const row = ws.addRow([...l.textes, estKit ? "kit" : prix, ...l.parMois, l.total]);
    const r = row.number, P = `$${col(C_PRIX)}${r}`;
    for (let i = 0; i < nM; i++) {
      row.getCell(C_E1 + i).value = { formula: `IF(ISNUMBER(${P}),${P}*${col(C_Q1 + i)}${r},0)` };
    }
    row.getCell(C_ETOT).value = { formula: `SUM(${col(C_E1)}${r}:${col(C_E1 + nM - 1)}${r})` };
    for (let c = 1; c <= nbCols; c++) {
      const cell = row.getCell(c);
      cell.font = { size: 10, name: c === 1 ? "Consolas" : "Calibri", color: { argb: "FF" + DARK }, bold: c === C_QTOT || c === C_ETOT };
      if (c >= C_E1) remplir(cell, BLEU_SOFT); else if (bandeau && o.filtres) remplir(cell, "F5F7FA");
      cell.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
      if (c >= C_PRIX) cell.alignment = { horizontal: "right" };
      if (c >= C_Q1 && c <= C_QTOT) cell.numFmt = FMT_UNITES;
      if (c >= C_E1) cell.numFmt = FMT_EUR;
    }
    ws.getCell(r, 1).numFmt = "@";
    const cp = row.getCell(C_PRIX);
    cp.numFmt = FMT_PRIX;
    if (estKit) { cp.font = { size: 9, italic: true, name: "Calibri", color: { argb: "FF6B7280" } }; cp.note = "Kit assemblé : pas acheté en tant que tel, ses composants sont comptés à part."; }
    else if (prix == null) remplir(cp, "FEF3C7");
    if (o.notesSurvol) {
      const parts = detailParRef.get(l.ref) || [];
      if (parts.length) {
        l.parMois.forEach((q, i) => {
          if (q > 0) { const t = noteOffres(parts.map(d => ({ d, q: d.parMois[i] || 0 }))); if (t) row.getCell(C_Q1 + i).note = t; }
        });
        const tTot = noteOffres(parts.map(d => ({ d, q: d.total })));
        if (tTot) row.getCell(C_QTOT).note = tTot;
      }
    }
  }
  prixManquants = refsSansPrix.size;
  const derniere = ws.rowCount;
  if (!o.lignes.length) {
    const empty = ws.addRow(["", "Aucun besoin calculé — renseigne les dates de début et de fin de campagne."]);
    ws.mergeCells(empty.number, 2, empty.number, nbCols);
    ws.getCell(empty.number, 2).font = { italic: true, size: 10, color: { argb: "FF6B7280" }, name: "Calibri" };
  }
  if (o.filtres && o.lignes.length) ws.autoFilter = { from: { row: head.number, column: 1 }, to: { row: derniere, column: nbCols } };

  // Pieds de tableau. Avec filtres : SUBTOTAL(109) ne somme que les lignes visibles.
  const somme = (c: number) => o.filtres
    ? { formula: `SUBTOTAL(109,${col(c)}${premiere}:${col(c)}${derniere})` }
    : { formula: `SUM(${col(c)}${premiere}:${col(c)}${derniere})` };
  const suffixe = o.filtres ? " (lignes filtrées)" : "";
  const pied = (libelle: string, couleur: string, cols: number[], fmtCell: string) => {
    const row = ws.addRow([]);
    row.getCell(Math.min(2, nT)).value = libelle;
    for (let c = 1; c <= nbCols; c++) {
      const cell = row.getCell(c);
      if (cols.includes(c) && o.lignes.length) { cell.value = somme(c); cell.numFmt = fmtCell; cell.alignment = { horizontal: "right" }; }
      cell.font = { bold: true, size: 10, name: "Calibri", color: { argb: "FF" + WHITE } };
      remplir(cell, couleur);
    }
    return row;
  };
  const range = (a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i);
  ws.addRow([]);
  pied(`TOTAL UNITÉS${suffixe}`, TEAL, range(C_Q1, C_QTOT), FMT_UNITES);
  const piedMois = pied(`TOTAL MOIS (€)${suffixe}`, BLEU, range(C_E1, C_ETOT), FMT_EUR);
  // Total année : une seule case, sous « Total réf », comme dans le planning achats.
  const annee = ws.addRow([]);
  annee.getCell(Math.min(2, nT)).value = `TOTAL ANNÉE (€)${suffixe}`;
  annee.getCell(Math.min(2, nT)).font = { bold: true, size: 11, name: "Calibri", color: { argb: "FF" + BLEU } };
  const ca = annee.getCell(C_ETOT);
  ca.value = { formula: `${col(C_ETOT)}${piedMois.number}` };
  ca.numFmt = FMT_EUR; ca.alignment = { horizontal: "right" };
  ca.font = { bold: true, size: 12, name: "Calibri", color: { argb: "FF" + WHITE } };
  remplir(ca, BLEU);
  annee.height = 22;

  // Notes de bas de tableau
  ws.addRow([]);
  const notes = [
    "Profil de livraison : 40 % le mois précédant le début de l'offre, puis 60 % lissé à parts égales jusqu'à 1 mois avant la fin. Les mois s'étendent sur l'année suivante si une offre déborde. Composants de kits : 100 % à M-2 du début de campagne (assemblage avant lancement).",
    "Prix d'achat = coût unitaire Odoo. Les kits ne sont pas valorisés (leurs composants le sont). Modifier un prix met à jour les achats et les totaux.",
  ];
  if (prixManquants) notes.push(`⚠ ${prixManquants} référence(s) sans prix d'achat (cases jaunes) : non comptées dans les achats. Complète le prix pour les inclure.`);
  if (o.notesSurvol && (log.detail || []).length) notes.push("Survole une quantité pour voir les offres qui la composent. Le détail complet (filtrable) est dans l'onglet « Détail logistique ».");
  for (const t of notes) {
    const n = ws.addRow([t]);
    ws.mergeCells(n.number, 1, n.number, Math.min(nbCols, nT + 2 + nM));
    ws.getCell(n.number, 1).font = { italic: true, size: 9, color: { argb: t.startsWith("⚠") ? "FFB45309" : "FF6B7280" }, name: "Calibri" };
    ws.getCell(n.number, 1).alignment = { wrapText: true, vertical: "top" };
    n.height = 26;
  }
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
  // Références des kits (trousses assemblées) : pas achetées en tant que telles, ce sont leurs
  // composants qui le sont → exclues du total d'achat pour ne pas compter deux fois.
  refsKit?: string[];
  // Prix d'achat saisis sur les composants de kits (produits hors Odoo notamment).
  prixComposants?: Record<string, number>;
  // Composants de kits : besoin total à M-2, par composant (détail par kit dans `detail`).
  // Campagnes exclues faute de dates exploitables (manquantes, invalides ou fin < début).
  // Remontées à l'utilisateur : sans ça, leurs besoins disparaissaient en silence.
  ignorees?: string[];
}

const PART_VAGUE = 0.4; // 40% le mois -1
const MOIS_AVANCE_KIT = 2; // composants de kit : besoin à M-2 du début de campagne

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

/** Besoin d'une campagne détaillé par OFFRE : { ref, libellé du palier, quantité totale },
 *  paliers puis Grands Comptes et canaux non B2B.
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
  // Grands Comptes et Besoins non B2B (Maison Dr Hauschka, Eshop…) : quantités saisies par
  // enseigne / canal, à approvisionner en plus des paliers (même profil de livraison).
  const canaux = [
    ...(camp.gcEnseignes || []).map(e => ({ e, groupe: "Grands comptes" })),
    ...(camp.canauxNonB2B || []).map(e => ({ e, groupe: "Non B2B" })),
  ];
  for (const { e, groupe } of canaux) {
    for (const art of camp.articles) {
      const ref = art.ref.trim();
      if (!ref || estOpca(art.typProd)) continue;
      const qte = (e.qties || {})[qtyKeyLib(art, camp.articles)] || 0;
      if (qte > 0) out.push({ ref, palier: `${groupe} — ${e.nom || "(sans nom)"}`, qte });
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
  const refsKit = new Set<string>();
  const prixComposants: Record<string, number> = {};

  for (const camp of campagnes) {
    const ad = absMonth(camp.dateDebut);
    const af = absMonth(camp.dateFin);
    if (ad == null || af == null || af < ad) { ignorees.push(camp.nom || "(sans nom)"); continue; }
    const nameByRef: Record<string, string> = {};
    for (const a of camp.articles) if (a.ref.trim()) nameByRef[a.ref.trim()] = a.name || "";

    // On répartit OFFRE PAR OFFRE, puis on agrège par référence : le total par réf est
    // inchangé, et chaque case de la synthèse devient explicable (onglet Détail + survol).
    // Kits : leurs composants sont à approvisionner en une fois à M-2 du début de campagne
    // (assemblage avant lancement), au prorata du nombre de kits de chaque offre.
    const kitParRef = new Map(camp.articles.filter(a => a.kit && a.ref.trim() && (a.composants || []).length).map(a => [a.ref.trim(), a]));
    for (const r of kitParRef.keys()) refsKit.add(r);
    for (const b of besoinsParOffre(camp)) {
      const kit = kitParRef.get(b.ref);
      if (kit) {
        for (const comp of kit.composants || []) {
          // Produit saisi à la main sans code : la désignation sert de référence.
          const ref = (comp.ref || "").trim() || (comp.name || "").trim(), qte = b.qte * (comp.qty || 0);
          if (!ref || qte <= 0) continue;
          if ((comp.cout || 0) > 0) prixComposants[ref] = comp.cout!;
          const mois = ad - MOIS_AVANCE_KIT;
          const rep = new Map([[mois, qte]]);
          if (!accum[ref]) accum[ref] = { name: comp.name || nameByRef[ref] || "", parMoisAbs: new Map() };
          if (!accum[ref].name && comp.name) accum[ref].name = comp.name;
          detailAbs.push({ ref, name: comp.name || nameByRef[ref] || "", campagne: camp.nom || "(sans nom)", palier: `${b.palier} — composant du kit ${kit.name || kit.ref}`, parMoisAbs: rep });
          accum[ref].parMoisAbs.set(mois, (accum[ref].parMoisAbs.get(mois) || 0) + qte);
          if (mois < minAbs) minAbs = mois;
          if (mois > maxAbs) maxAbs = mois;
        }
      }
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

  return { lignes, totalParMois, totalGeneral, moisLabels, ignorees, detail, refsKit: [...refsKit], prixComposants };
}
