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
}

export interface PalierSaisi {
  code: string;           // ex. "REGE1"
  label: string;          // ex. "Premium"
  nbPacks: number;        // nb de packs cible du palier
  // qté/pack par article, indexée par ref. Vide = on prend la reco (conso ÷ nbPacks).
  qtyParPack: Record<string, number>;
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

/** Récupère conso N-1 + pricing Odoo pour tous les articles de la campagne. */
export async function analyseCampagneCreee(session: odoo.OdooSession, camp: CampagneCreee): Promise<CampagneCreee> {
  const refs = [...new Set(camp.articles.map(a => a.ref.trim()).filter(Boolean))];
  if (!refs.length) return camp;

  const conso = await odoo.getConsumption(session, refs, camp.periodeDebut, camp.periodeFin);
  const productIds = Object.values(conso).map(c => c.productId).filter(Boolean);
  const pricing = await odoo.getProductsPricing(session, productIds);

  const articles = camp.articles.map(a => {
    const ref = a.ref.trim();
    const c = conso[ref];
    const pid = c?.productId || 0;
    const pr = pid ? pricing[pid] : undefined;
    return {
      ...a,
      ref,
      productId: pid,
      name: pr?.name || c?.name || "",
      barcode: pr?.barcode || "",
      standardPrice: pr?.standardPrice || 0,
      listPrice: pr?.listPrice || 0,
      ppc: pr?.ppc || 0,
      consoN1: c?.qty || 0,
      found: c?.found ?? false,
    } as ArticleCampagne;
  });

  return { ...camp, articles };
}

/** Qté/pack effective d'un article dans un palier : valeur saisie sinon reco (conso ÷ nbPacks). */
export function qtyParPack(art: ArticleCampagne, pal: PalierSaisi): number {
  const manual = pal.qtyParPack[art.ref];
  if (manual != null) return manual;
  const nb = pal.nbPacks || 0;
  return nb > 0 ? Math.round((art.consoN1 || 0) / nb) : 0;
}

export interface ExportPayload {
  nom: string;
  paliers: Array<{
    code: string; label: string; qtyPacks: number;
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
