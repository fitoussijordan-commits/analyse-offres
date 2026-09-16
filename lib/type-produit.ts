// lib/type-produit.ts — Règles par type de produit, partagées par l'écran, les calculs et l'export.
//
// OPCA (Offre Produit Contre Achat) : article virtuel qui représente le panier d'achat de
// l'institut (ex. « OPCA 500 » = 500 € d'achats tout catalogue). Méthode marketing :
//  - tarif = montant du seuil, payé net par l'institut ;
//  - coût  = 55 % du seuil (marge 45 %, hors gratuits envoyés à tous les instituts) ;
//  - pas de remise additionnelle (le seuil est un montant net) ;
//  - pas de besoin logistique (ce n'est pas un produit à approvisionner).
// Les unités offertes en échange se saisissent à côté, en type UG.

export const TYP_VENTE = "Produit Vente";
export const TYP_OPCA = "OPCA";
export const OPCA_TAUX_COUT = 0.55;

export const estOpca = (typ?: string) => typ === TYP_OPCA;

/** Génère du CA : Produit Vente et OPCA. UG / Testeur / PLV / Échantillon sont gratuits. */
export const genereCA = (typ?: string) => { const t = typ || TYP_VENTE; return t === TYP_VENTE || t === TYP_OPCA; };

/** Coût unitaire d'une OPCA à partir de son seuil. */
export const coutOpca = (seuil?: number) => Math.round((seuil || 0) * OPCA_TAUX_COUT * 100) / 100;

/** Remise additionnelle effective d'une ligne : toujours 0 pour une OPCA. */
export const remiseAddLigne = (typ: string | undefined, remiseAdd: number) => estOpca(typ) ? 0 : remiseAdd;
