// lib/calc-offre.ts — Couche de calcul PURE de l'offre (mêmes formules que le template Excel).
// Réutilisable par l'aperçu interactif (recalcul live) ET l'export. N'a AUCun effet de bord
// et ne dépend d'aucune lib : simples fonctions sur des nombres.

export const TYPOLOGIES = ["Ambassadeur", "Compagnon", "Challenger", "Rose", "Prunelier", "Anthylide", "Calendula"];
// Valeurs par défaut du gabarit (peuvent être surchargées par palier).
export const DEFAULT_PCTS = [0.5, 0.1, 0.1, 0.1, 0.1, 0.05, 0.05];
export const DEFAULT_REMISES = [0.17, 0.13, 0.08, 0.325, 0.3, 0.28, 0.25];
export const REMISE_ADD_DEFAUT = 0.15; // colonne I par défaut

export interface CalcProduit {
  ref: string; name: string;
  qtyParPack: number;      // E : Pdt dans offre
  standardPrice: number;   // F : coût achat unitaire
  listPrice: number;       // H : tarif revendeur unitaire
  ppc: number;             // J : PPC
  remiseAdd?: number;      // I : remise additionnelle (défaut REMISE_ADD_DEFAUT)
}

export interface CalcPalier {
  code: string; label: string;
  nbPacks: number;                 // B : nombre d'offres
  pcts?: number[];                 // % offres par typologie (7) — sinon DEFAULT_PCTS
  remises?: number[];              // remises par typologie (7) — sinon DEFAULT_REMISES
  remiseAdd?: number;              // remise additionnelle du palier (col I) appliquée à tous
  produits: CalcProduit[];
}

// Résultat par typologie pour un palier.
export interface TypoResult { ca: number; marge: number; nbOffres: number; }
export interface PalierResult {
  parTypo: TypoResult[];           // 7 typologies
  caTotal: number;                 // Σ CA typologies
  margeTotal: number;              // Σ marges typologies
  margePct: number;                // margeTotal / caTotal
  coutTotal: number;               // Σ coûts d'achat écoulés
}

/** Calcule CA/marge d'un palier, par typologie et en total. Formules identiques au template :
 *  Nb offres typo = %offres × nbPacks
 *  CA   = E × H × (1−I) × NbOffresTypo × (1−remiseTypo)
 *  Marge = CA − (E × F × NbOffresTypo)
 *  sommés sur tous les produits. */
export function calcPalier(pal: CalcPalier): PalierResult {
  const pcts = pal.pcts && pal.pcts.length === 7 ? pal.pcts : DEFAULT_PCTS;
  const remises = pal.remises && pal.remises.length === 7 ? pal.remises : DEFAULT_REMISES;
  const parTypo: TypoResult[] = [];
  let caTotal = 0, margeTotal = 0, coutTotal = 0;

  for (let t = 0; t < 7; t++) {
    const nbOff = pcts[t] * (pal.nbPacks || 0);
    let ca = 0, cout = 0;
    for (const p of pal.produits) {
      const E = p.qtyParPack || 0, H = p.listPrice || 0, F = p.standardPrice || 0;
      const I = p.remiseAdd != null ? p.remiseAdd : (pal.remiseAdd != null ? pal.remiseAdd : REMISE_ADD_DEFAUT);
      ca += E * H * (1 - I) * nbOff * (1 - remises[t]);
      cout += E * F * nbOff;
    }
    parTypo.push({ ca, marge: ca - cout, nbOffres: nbOff });
    caTotal += ca; margeTotal += ca - cout; coutTotal += cout;
  }
  return { parTypo, caTotal, margeTotal, margePct: caTotal > 0 ? margeTotal / caTotal : 0, coutTotal };
}

// ── Décomposition (audit) ────────────────────────────────────────────────────
// Détaille, produit par produit, d'où vient le CA / la marge d'un palier. Sert au
// panneau « Détail des calculs » : les valeurs affichées viennent des mêmes formules
// que calcPalier(), il n'y a donc aucun risque de divergence entre écran et export.
export interface DetailProduit {
  ref: string; name: string;
  qtyParPack: number;      // E
  listPrice: number;       // H  tarif revendeur
  standardPrice: number;   // F  coût
  remiseAdd: number;       // I
  prixNetMoyen: number;    // H × (1−I) × (1−remise moyenne pondérée)
  unitesTotal: number;     // E × Σ nbOffres (toutes typologies)
  ca: number;              // Σ typologies
  cout: number;            // Σ typologies
  marge: number;           // ca − cout
  margePct: number;
  partCa: number;          // poids dans le CA du palier
}
export interface DetailPalier {
  produits: DetailProduit[];
  remiseMoyenne: number;   // Σ (%offres × remise typo) — remise moyenne pondérée
  nbOffresTotal: number;   // Σ nbOffres typologies (= Σ%  × nbPacks)
  caTotal: number; coutTotal: number; margeTotal: number; margePct: number;
  caParOffre: number; margeParOffre: number;
}

/** Décompose le CA/marge d'un palier produit par produit (mêmes formules que calcPalier). */
export function detailPalier(pal: CalcPalier): DetailPalier {
  const pcts = pal.pcts && pal.pcts.length === 7 ? pal.pcts : DEFAULT_PCTS;
  const remises = pal.remises && pal.remises.length === 7 ? pal.remises : DEFAULT_REMISES;
  const nbOffres = pcts.map(p => p * (pal.nbPacks || 0));
  const nbOffresTotal = nbOffres.reduce((s, n) => s + n, 0);
  // Remise moyenne pondérée par le nombre d'offres de chaque typologie.
  const remiseMoyenne = nbOffresTotal > 0
    ? remises.reduce((s, r, t) => s + r * nbOffres[t], 0) / nbOffresTotal
    : 0;

  const produits: DetailProduit[] = pal.produits.map(p => {
    const E = p.qtyParPack || 0, H = p.listPrice || 0, F = p.standardPrice || 0;
    const I = p.remiseAdd != null ? p.remiseAdd : (pal.remiseAdd != null ? pal.remiseAdd : REMISE_ADD_DEFAUT);
    let ca = 0, cout = 0;
    for (let t = 0; t < 7; t++) {
      ca += E * H * (1 - I) * nbOffres[t] * (1 - remises[t]);
      cout += E * F * nbOffres[t];
    }
    return {
      ref: p.ref, name: p.name,
      qtyParPack: E, listPrice: H, standardPrice: F, remiseAdd: I,
      prixNetMoyen: H * (1 - I) * (1 - remiseMoyenne),
      unitesTotal: E * nbOffresTotal,
      ca, cout, marge: ca - cout, margePct: ca > 0 ? (ca - cout) / ca : 0,
      partCa: 0,
    };
  });
  const caTotal = produits.reduce((s, p) => s + p.ca, 0);
  const coutTotal = produits.reduce((s, p) => s + p.cout, 0);
  for (const p of produits) p.partCa = caTotal > 0 ? p.ca / caTotal : 0;
  const margeTotal = caTotal - coutTotal;
  const nb = pal.nbPacks || 0;
  return {
    produits, remiseMoyenne, nbOffresTotal,
    caTotal, coutTotal, margeTotal,
    margePct: caTotal > 0 ? margeTotal / caTotal : 0,
    caParOffre: nb > 0 ? caTotal / nb : 0,
    margeParOffre: nb > 0 ? margeTotal / nb : 0,
  };
}

export interface SyntheseGlobale { caTotal: number; margeTotal: number; margePct: number; nbPacks: number; }

/** Synthèse de toute la campagne (somme des paliers). */
export function calcSynthese(paliers: CalcPalier[]): SyntheseGlobale {
  let caTotal = 0, margeTotal = 0, nbPacks = 0;
  for (const pal of paliers) {
    const r = calcPalier(pal);
    caTotal += r.caTotal; margeTotal += r.margeTotal; nbPacks += pal.nbPacks || 0;
  }
  return { caTotal, margeTotal, margePct: caTotal > 0 ? margeTotal / caTotal : 0, nbPacks };
}

/** Besoin total par référence = Σ_paliers (qté/pack × nb packs). */
export function calcBesoinParRef(paliers: CalcPalier[]): Record<string, { name: string; total: number }> {
  const out: Record<string, { name: string; total: number }> = {};
  for (const pal of paliers) for (const p of pal.produits) {
    const ref = (p.ref || "").trim();
    if (!ref) continue;
    if (!out[ref]) out[ref] = { name: p.name || "", total: 0 };
    out[ref].total += (p.qtyParPack || 0) * (pal.nbPacks || 0);
    if (!out[ref].name && p.name) out[ref].name = p.name;
  }
  return out;
}
