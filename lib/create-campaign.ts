// lib/create-campaign.ts — Création d'une campagne "à blanc" (sans campagne N-1 à analyser).
// Les ARTICLES sont communs à toute la campagne (mêmes codes dans tous les paliers).
// Chaque PALIER a son propre nombre de packs → sa propre qté/pack par article.
// L'app récupère la conso N-1 (période bornée) + le pricing Odoo, recommande la qté/pack
// par palier (conso ÷ nbPacks du palier), puis produit le même fichier Proposition.

import * as odoo from "@/lib/odoo";

export interface ArticleCampagne {
  ref: string;            // code article (commun à tous les paliers)
  // résolu après analyse :
  productId?: number;
  name?: string;
  barcode?: string;
  standardPrice?: number;
  listPrice?: number;
  ppc?: number;
  consoN1?: number;       // conso N-1 sur la période
  found?: boolean;
  // Réf libre (PLV/testeur non créé sur Odoo) : l'utilisateur saisit nom + prix à la main.
  // Quand manuel = true, l'analyse Odoo ne doit pas écraser ces valeurs.
  manuel?: boolean;
  // Type de produit (colonne D du template) : Produit Vente / Testeur / Échantillon / UG / PLV.
  typProd?: string;
}

// Valeurs possibles du type de produit (liste déroulante).
export const TYPES_PRODUIT = ["Produit Vente", "Testeur", "Échantillon", "UG", "PLV"];

export interface PalierSaisi {
  code: string;           // ex. "REGE1"
  label: string;          // ex. "Premium"
  nbPacks: number;        // nb de packs cible du palier
  // qté/pack par article, indexée par ref. Vide = on prend la reco (conso ÷ nbPacks).
  qtyParPack: Record<string, number>;
  // Remise statut : standard (même % pour toutes les typologies) ou spécifique (gabarit).
  remiseStandard?: boolean;
  remiseStandardTaux?: number; // ex. 0.17
  // Remise additionnelle (colonne I) appliquée à tous les produits du palier (ex. 0.15).
  remiseAddTaux?: number;
  // Reco % offres par typologie (7 valeurs, ordre TYPOLOGIES) propre à CE palier, calculée
  // sur les ventes N-1 du code offre du palier (qui l'a acheté l'an dernier).
  pctOffresReco?: number[];
}

export interface CampagneCreee {
  id: string;
  nom: string;
  dateDebut: string;
  dateFin: string;
  periodeDebut: string;
  periodeFin: string;
  articles: ArticleCampagne[]; // communs à tous les paliers
  paliers: PalierSaisi[];      // chaque palier porte sa propre reco % offres (par code offre)
  annee?: string;             // année / cycle de rangement, saisi par l'utilisateur (ex. "2027")
  createdAt?: string;
}

// Ordre des 7 typologies du template (= statuts clients Odoo, mêmes noms).
export const TYPOLOGIES = ["Ambassadeur", "Compagnon", "Challenger", "Rose", "Prunelier", "Anthylide", "Calendula"];

export function genId(): string {
  return Math.random().toString(36).slice(2, 10);
}

// Forme de campagne d'analyse (lib/campaigns.ts) pour le pont vers l'outil Analyse, sans dépendance.
export interface CampagneAnalyseLike { id: string; nom: string; offres: string[]; produits: string[]; notes: string[]; }

/**
 * Convertit une campagne CRÉÉE en campagne d'ANALYSE pour suivre sa progression :
 *  - offres  = codes des paliers (packs Odoo),
 *  - produits = réfs des articles (fallback si les packs n'existent pas encore),
 *  - notes    = vide (l'utilisateur peut en ajouter dans l'analyse).
 * L'outil d'analyse existant compte alors les ventes réelles via codes offres → sinon produits.
 */
export function campagneCreeeToAnalyse(camp: CampagneCreee): CampagneAnalyseLike {
  const offres = [...new Set(camp.paliers.map(p => (p.code || "").trim()).filter(Boolean))];
  const produits = [...new Set(camp.articles.map(a => (a.ref || "").trim()).filter(Boolean))];
  return { id: genId(), nom: camp.nom || "Campagne", offres, produits, notes: [] };
}

// Forme minimale d'une préco (lib/preco.ts) pour la conversion, sans dépendance circulaire.
interface PrecoLigneLike { ref: string; name: string; productId: number; qtyParPack: number; conserve: boolean; }
interface PrecoPalierLike { code: string; label: string; qtyPacks: number; produits: PrecoLigneLike[]; }
interface PrecoLike { nom: string; paliers: PrecoPalierLike[]; }

/**
 * Convertit une préco N+1 (analyse de campagne) en campagne créée (à blanc), pour transférer
 * vers l'écran "Créer une campagne". Articles = union des produits conservés (tous paliers).
 * Chaque palier reprend sa qté/pack par article (0 si l'article n'y figure pas).
 */
export function precoToCampagne(preco: PrecoLike): CampagneCreee {
  // Union des articles conservés, dans l'ordre d'apparition.
  const seen = new Set<string>();
  const articles: ArticleCampagne[] = [];
  for (const pal of preco.paliers) {
    for (const p of pal.produits) {
      if (!p.conserve) continue;
      const ref = (p.ref || "").trim();
      if (!ref || seen.has(ref)) continue;
      seen.add(ref);
      articles.push({ ref, name: p.name || "", productId: p.productId });
    }
  }

  const paliers: PalierSaisi[] = preco.paliers.map(pal => {
    const qtyParPack: Record<string, number> = {};
    for (const p of pal.produits) {
      if (!p.conserve) continue;
      const ref = (p.ref || "").trim();
      if (ref) qtyParPack[ref] = p.qtyParPack || 0;
    }
    return { code: pal.code, label: pal.label, nbPacks: pal.qtyPacks || 0, qtyParPack };
  });

  return {
    id: genId(),
    nom: preco.nom || "",
    dateDebut: "", dateFin: "", periodeDebut: "", periodeFin: "",
    articles: articles.length ? articles : [{ ref: "" }],
    paliers: paliers.length ? paliers : [],
  };
}

/** Récupère conso N-1 + pricing Odoo pour tous les articles de la campagne. */
export async function analyseCampagneCreee(session: odoo.OdooSession, camp: CampagneCreee): Promise<CampagneCreee> {
  const refs = [...new Set(camp.articles.map(a => a.ref.trim()).filter(Boolean))];
  if (!refs.length) return camp;

  // 1) Résoudre les produits (id, libellé) — TOUJOURS, indépendamment des dates de conso.
  //    Ainsi le libellé et les prix sont chargés même sans période N-1 renseignée.
  const prods = await odoo.searchProductsByRefs(session, refs);   // { ref -> {id, name} }
  const ids = Object.values(prods).map(p => p.id).filter(Boolean);
  const pricing = await odoo.getProductsPricing(session, ids);
  const pricingByRef: Record<string, odoo.ProductPricing> = {};
  for (const pr of Object.values(pricing)) if (pr.ref) pricingByRef[pr.ref] = pr;

  // 2) Conso N-1 seulement si la période est renseignée.
  const conso = (camp.periodeDebut && camp.periodeFin)
    ? await odoo.getConsumption(session, refs, camp.periodeDebut, camp.periodeFin)
    : {};

  const articles = camp.articles.map(a => {
    const ref = a.ref.trim();
    const resolved = prods[ref];
    const pid = resolved?.id || 0;
    // Réf manuelle OU réf absente d'Odoo → on conserve les valeurs saisies (ne pas écraser).
    if (a.manuel || !pid) {
      const c = conso[ref];
      return { ...a, ref, manuel: true, productId: pid, found: false, consoN1: c?.qty ?? a.consoN1 ?? 0 } as ArticleCampagne;
    }
    const pr = pricingByRef[ref];
    const c = conso[ref];
    return {
      ...a,
      ref,
      manuel: false,
      productId: pid,
      name: pr?.name || resolved?.name || c?.name || "",
      barcode: pr?.barcode || "",
      standardPrice: pr?.standardPrice || 0,
      listPrice: pr?.listPrice || 0,
      ppc: pr?.ppc || 0,
      consoN1: c?.qty || 0,
      found: true,
    } as ArticleCampagne;
  });

  // Reco % offres par typologie = répartition des QUANTITÉS vendues N-1 par statut client,
  // calculée sur les ARTICLES de la campagne. On ne compte QUE les 7 typologies exactes :
  // les statuts hors-liste (GC Concept Store, Siège…) et l'inconnu sont IGNORÉS, et le % de
  // chaque typologie est rapporté au total des seules 7 typologies (donc Σ des 7 = 100%).
  const norm = (s: string) => (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
  let paliers = camp.paliers;
  if (camp.periodeDebut && camp.periodeFin && refs.length) {
    // Quantités par statut (tous articles de la campagne confondus).
    const vent = await odoo.getQtyByStatut(session, refs, camp.periodeDebut, camp.periodeFin);
    const qtyParStatut: Record<string, number> = {};
    for (const v of Object.values(vent)) {
      for (const [statut, q] of Object.entries(v.parStatut)) {
        qtyParStatut[norm(statut)] = (qtyParStatut[norm(statut)] || 0) + q;
      }
    }
    // Quantité de chaque typologie = match EXACT (nom normalisé identique) uniquement.
    const qtyTypo = TYPOLOGIES.map(typo => qtyParStatut[norm(typo)] || 0);
    const totalTypo = qtyTypo.reduce((s, n) => s + n, 0);
    if (totalTypo > 0) {
      const reco = qtyTypo.map(q => q / totalTypo);
      paliers = camp.paliers.map(pal => ({ ...pal, pctOffresReco: reco }));
    }
  }

  return { ...camp, articles, paliers };
}

/** Somme des packs de tous les paliers d'une campagne (dénominateur de la reco). */
export function totalPacks(paliers: PalierSaisi[]): number {
  return paliers.reduce((s, p) => s + (p.nbPacks || 0), 0);
}

/** Qté/pack effective d'un article dans un palier : valeur saisie (>0) sinon reco.
 *
 *  Reco = conso N-1 répartie sur TOUS les paliers, pas recopiée dans chacun :
 *    qté/pack = conso ÷ (total des packs de la campagne)
 *  Ainsi Σ(paliers) qté/pack × nbPacks ≈ conso N-1 (le gâteau est réparti, pas multiplié).
 *  `totalP` = somme des packs de tous les paliers ; si omis (ancien appel), on retombe sur
 *  l'ancien comportement (÷ nbPacks du palier seul) pour rétro-compatibilité.
 *
 *  Un 0 stocké n'écrase PAS la reco : si l'utilisateur veut exclure un article, il le retire. */
export function qtyParPack(art: ArticleCampagne, pal: PalierSaisi, totalP?: number): number {
  const manual = pal.qtyParPack[art.ref];
  if (manual != null && manual > 0) return manual;
  const denom = (totalP != null && totalP > 0) ? totalP : (pal.nbPacks || 0);
  const reco = denom > 0 ? Math.round((art.consoN1 || 0) / denom) : 0;
  // Si pas de reco possible (pas de conso), on respecte une éventuelle saisie 0.
  return reco > 0 ? reco : (manual ?? 0);
}

export interface ExportPayload {
  nom: string;
  paliers: Array<{
    code: string; label: string; qtyPacks: number;
    remiseStandard?: boolean; remiseStandardTaux?: number; remiseAddTaux?: number;
    pctOffres?: number[]; // % offres reco par typologie (7 valeurs), propre à ce palier
    produits: Array<{
      ref: string; name: string; productId: number;
      qtyParPack: number;
      barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
      typProd?: string;
    }>;
  }>;
}

/** Convertit la campagne (enrichie) vers le payload d'export Proposition. */
export function toExportPayload(camp: CampagneCreee): ExportPayload {
  const arts = camp.articles.filter(a => a.ref.trim());
  const totalP = totalPacks(camp.paliers);
  return {
    nom: camp.nom,
    paliers: camp.paliers.map(pal => ({
      code: pal.code,
      label: pal.label,
      qtyPacks: pal.nbPacks || 0,
      remiseStandard: pal.remiseStandard,
      remiseStandardTaux: pal.remiseStandardTaux,
      remiseAddTaux: pal.remiseAddTaux,
      pctOffres: pal.pctOffresReco,
      produits: arts.map(a => {
        // Tout ce qui n'est PAS "Produit Vente" (UG, Testeur, PLV, Échantillon) est gratuit :
        // prix de vente (listPrice) et PPC forcés à 0 → aucun CA généré. Le coût (standardPrice)
        // est conservé car ces produits ont un coût réel pour l'entreprise.
        const typ = a.typProd || "Produit Vente";
        const estVente = typ === "Produit Vente";
        return {
          ref: a.ref.trim(),
          name: a.name || "",
          productId: a.productId || 0,
          qtyParPack: qtyParPack(a, pal, totalP),
          barcode: a.barcode || "",
          standardPrice: a.standardPrice || 0,
          listPrice: estVente ? (a.listPrice || 0) : 0,
          ppc: estVente ? (a.ppc || 0) : 0,
          typProd: typ,
        };
      }),
    })).filter(p => p.produits.length),
  };
}
