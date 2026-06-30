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
}

export interface PalierSaisi {
  code: string;           // ex. "REGE1"
  label: string;          // ex. "Premium"
  nbPacks: number;        // nb de packs cible du palier
  // qté/pack par article, indexée par ref. Vide = on prend la reco (conso ÷ nbPacks).
  qtyParPack: Record<string, number>;
  // Remise statut : standard (même % pour toutes les typologies) ou spécifique (gabarit).
  remiseStandard?: boolean;
  remiseStandardTaux?: number; // ex. 0.17
}

export interface CampagneCreee {
  id: string;
  nom: string;
  dateDebut: string;
  dateFin: string;
  periodeDebut: string;
  periodeFin: string;
  articles: ArticleCampagne[]; // communs à tous les paliers
  paliers: PalierSaisi[];
  createdAt?: string;
}

export function genId(): string {
  return Math.random().toString(36).slice(2, 10);
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

  return { ...camp, articles };
}

/** Qté/pack effective d'un article dans un palier : valeur saisie (>0) sinon reco (conso ÷ nbPacks).
 *  Un 0 stocké n'écrase PAS la reco : si l'utilisateur veut exclure un article, il le retire.
 *  Cela évite qu'un 0 résiduel (transfert/chargement) masque la recommandation. */
export function qtyParPack(art: ArticleCampagne, pal: PalierSaisi): number {
  const manual = pal.qtyParPack[art.ref];
  if (manual != null && manual > 0) return manual;
  const nb = pal.nbPacks || 0;
  const reco = nb > 0 ? Math.round((art.consoN1 || 0) / nb) : 0;
  // Si pas de reco possible (pas de conso), on respecte une éventuelle saisie 0.
  return reco > 0 ? reco : (manual ?? 0);
}

export interface ExportPayload {
  nom: string;
  paliers: Array<{
    code: string; label: string; qtyPacks: number;
    remiseStandard?: boolean; remiseStandardTaux?: number;
    produits: Array<{
      ref: string; name: string; productId: number;
      qtyParPack: number;
      barcode?: string; standardPrice?: number; listPrice?: number; ppc?: number;
    }>;
  }>;
}

/** Convertit la campagne (enrichie) vers le payload d'export Proposition. */
export function toExportPayload(camp: CampagneCreee): ExportPayload {
  const arts = camp.articles.filter(a => a.ref.trim());
  return {
    nom: camp.nom,
    paliers: camp.paliers.map(pal => ({
      code: pal.code,
      label: pal.label,
      qtyPacks: pal.nbPacks || 0,
      remiseStandard: pal.remiseStandard,
      remiseStandardTaux: pal.remiseStandardTaux,
      produits: arts.map(a => ({
        ref: a.ref.trim(),
        name: a.name || "",
        productId: a.productId || 0,
        qtyParPack: qtyParPack(a, pal),
        barcode: a.barcode || "",
        standardPrice: a.standardPrice || 0,
        listPrice: a.listPrice || 0,
        ppc: a.ppc || 0,
      })),
    })).filter(p => p.produits.length),
  };
}
