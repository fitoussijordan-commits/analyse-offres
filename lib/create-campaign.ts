// lib/create-campaign.ts — Création d'une campagne "à blanc" (sans campagne N-1 à analyser).
// L'utilisateur définit ses paliers, ses articles et un nombre de packs cible par palier ;
// l'app va chercher la consommation N-1 sur une période bornée (Odoo) pour recommander la
// qté/pack de chaque article, puis produit le même fichier Proposition que le mode analyse.

import * as odoo from "@/lib/odoo";

export interface ArticleSaisi {
  ref: string;            // code article saisi par l'utilisateur
  // valeurs résolues / recommandées (remplies après analyse conso + pricing Odoo)
  productId?: number;
  name?: string;
  barcode?: string;
  standardPrice?: number;
  listPrice?: number;
  ppc?: number;
  consoN1?: number;       // conso N-1 sur la période (qté vendue)
  qtyParPack?: number;    // reco = round(consoN1 / nbPacks), ajustable par l'utilisateur
  found?: boolean;        // trouvé dans Odoo ?
}

export interface PalierSaisi {
  code: string;           // ex. "REGE1" (sert de code dans le titre du bloc)
  label: string;          // ex. "Premium"
  nbPacks: number;        // nombre de packs cible du palier
  articles: ArticleSaisi[];
}

export interface CampagneCreee {
  id: string;
  nom: string;
  dateDebut: string;      // début campagne (YYYY-MM-DD)
  dateFin: string;        // fin campagne
  periodeDebut: string;   // début fenêtre conso N-1 (YYYY-MM-DD)
  periodeFin: string;     // fin fenêtre conso N-1
  paliers: PalierSaisi[];
  createdAt?: string;
}

export function genId(): string {
  return Math.random().toString(36).slice(2, 10);
}

/**
 * Enrichit une campagne créée : pour chaque article de chaque palier, récupère depuis Odoo
 * la conso N-1 (sur la période bornée) + les données tarifaires, puis calcule la qté/pack
 * recommandée = round(conso / nbPacks). Renvoie une nouvelle campagne (immutable).
 */
export async function analyseCampagneCreee(session: odoo.OdooSession, camp: CampagneCreee): Promise<CampagneCreee> {
  // Toutes les refs, tous paliers confondus (un seul appel Odoo pour conso + pricing).
  const allRefs = [...new Set(camp.paliers.flatMap(p => p.articles.map(a => a.ref.trim()).filter(Boolean)))];
  if (!allRefs.length) return camp;

  const conso = await odoo.getConsumption(session, allRefs, camp.periodeDebut, camp.periodeFin);
  const productIds = Object.values(conso).map(c => c.productId).filter(Boolean);
  const pricing = await odoo.getProductsPricing(session, productIds);

  const paliers = camp.paliers.map(pal => {
    const nbPacks = pal.nbPacks || 0;
    const articles = pal.articles.map(a => {
      const ref = a.ref.trim();
      const c = conso[ref];
      const pid = c?.productId || 0;
      const pr = pid ? pricing[pid] : undefined;
      const consoN1 = c?.qty || 0;
      // Reco qté/pack = conso N-1 / nb de packs (arrondi), min 0.
      const reco = nbPacks > 0 ? Math.round(consoN1 / nbPacks) : 0;
      return {
        ...a,
        ref,
        productId: pid,
        name: pr?.name || c?.name || "",
        barcode: pr?.barcode || "",
        standardPrice: pr?.standardPrice || 0,
        listPrice: pr?.listPrice || 0,
        ppc: pr?.ppc || 0,
        consoN1,
        // On ne garde la reco que si l'utilisateur n'a pas déjà saisi une valeur manuelle.
        qtyParPack: a.qtyParPack != null ? a.qtyParPack : reco,
        found: c?.found ?? false,
      } as ArticleSaisi;
    });
    return { ...pal, articles };
  });

  return { ...camp, paliers };
}

// Payload attendu par /api/export-template (identique au mode préco analyse).
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

/** Convertit une campagne créée (enrichie) vers le payload d'export Proposition. */
export function toExportPayload(camp: CampagneCreee): ExportPayload {
  return {
    nom: camp.nom,
    paliers: camp.paliers
      .map(pal => ({
        code: pal.code,
        label: pal.label,
        qtyPacks: pal.nbPacks || 0,
        produits: pal.articles
          .filter(a => a.ref.trim())
          .map(a => ({
            ref: a.ref.trim(),
            name: a.name || "",
            productId: a.productId || 0,
            qtyParPack: a.qtyParPack || 0,
            barcode: a.barcode || "",
            standardPrice: a.standardPrice || 0,
            listPrice: a.listPrice || 0,
            ppc: a.ppc || 0,
          })),
      }))
      .filter(p => p.produits.length),
  };
}
