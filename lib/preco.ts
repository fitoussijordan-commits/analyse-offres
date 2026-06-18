// lib/preco.ts — Préconisation d'offre N+1 (rôle trade marketing / chef de projet)
// À partir des produits conservés + du croisement produit × statut client de la campagne
// actuelle, génère une recommandation d'offre par statut client.
import type { CampaignResult, ProduitStatut } from "@/lib/analyse-campaign";

export interface PrecoProduit {
  ref: string; name: string; productId: number;
  ca: number; qty: number;          // perf sur le segment concerné
  pctSegment: number;               // part du CA produit dans le CA total du segment
  conserve: boolean;                // fait partie des produits conservés
}

export interface PrecoStatut {
  statut: string;
  caSegment: number;                // CA total du segment (tous produits campagne)
  conserves: PrecoProduit[];        // produits conservés, perf sur ce segment
  candidats: PrecoProduit[];        // produits à AJOUTER : performants sur ce segment, non conservés
  aRetirer: PrecoProduit[];         // produits conservés mais faibles sur ce segment
}

export interface PrecoResult {
  nom: string;
  produitsConserves: { ref: string; name: string; productId: number }[];
  parStatut: PrecoStatut[];
  global: { candidats: PrecoProduit[]; conserves: PrecoProduit[] };
}

// Seuils (ajustables)
const TOP_CANDIDATS = 5;          // nb max de candidats proposés par segment
const SEUIL_FAIBLE = 0.02;        // < 2% du CA segment => produit conservé jugé faible sur ce segment

/**
 * Construit la préconisation N+1.
 * @param result  résultat d'analyse de la campagne actuelle
 * @param conservedIds  ids des produits (productId) que l'utilisateur conserve dans l'offre
 */
export function buildPreco(result: CampaignResult, conservedIds: number[]): PrecoResult {
  const conserved = new Set(conservedIds);
  const pps = result.produitsParStatut;

  // Référentiel produits conservés (nom/ref via le croisement, sinon via produits)
  const refIndex = new Map<number, ProduitStatut>();
  for (const p of pps) refIndex.set(p.productId, p);
  const produitsConserves = conservedIds.map(id => {
    const p = refIndex.get(id) || result.produits.find(x => x.productId === id);
    return { ref: p?.ref || "", name: p?.name || "", productId: id };
  });

  // CA par statut (segment) à partir des stats globales
  const caParStatut = new Map<string, number>();
  for (const s of result.statuts) caParStatut.set(s.name, s.ca);

  // Liste des statuts présents, triés par CA décroissant
  const statutsTries = [...result.statuts].sort((a, b) => b.ca - a.ca).map(s => s.name);

  const parStatut: PrecoStatut[] = statutsTries.map(statut => {
    const caSegment = caParStatut.get(statut) || 0;
    // Tous les produits ayant vendu sur ce segment, triés par CA segment
    const surSegment = pps
      .map(p => {
        const v = p.parStatut[statut];
        return v ? { p, ca: v.ca, qty: v.qty } : null;
      })
      .filter((x): x is { p: ProduitStatut; ca: number; qty: number } => !!x && x.ca > 0)
      .sort((a, b) => b.ca - a.ca);

    const mk = (x: { p: ProduitStatut; ca: number; qty: number }): PrecoProduit => ({
      ref: x.p.ref, name: x.p.name, productId: x.p.productId,
      ca: x.ca, qty: x.qty,
      pctSegment: caSegment > 0 ? x.ca / caSegment : 0,
      conserve: conserved.has(x.p.productId),
    });

    const conserves = surSegment.filter(x => conserved.has(x.p.productId)).map(mk);
    // Candidats à ajouter : non conservés, meilleurs vendeurs du segment
    const candidats = surSegment.filter(x => !conserved.has(x.p.productId)).slice(0, TOP_CANDIDATS).map(mk);
    // Produits conservés mais faibles sur ce segment
    const aRetirer = conserves.filter(c => c.pctSegment < SEUIL_FAIBLE);

    return { statut, caSegment, conserves, candidats, aRetirer };
  });

  // Reco globale (tous segments confondus)
  const sortedAll = [...pps].sort((a, b) => b.ca - a.ca);
  const caTotal = result.caTotal || 0;
  const toGlobal = (p: ProduitStatut): PrecoProduit => ({
    ref: p.ref, name: p.name, productId: p.productId, ca: p.ca, qty: p.qtyVendue,
    pctSegment: caTotal > 0 ? p.ca / caTotal : 0, conserve: conserved.has(p.productId),
  });
  const global = {
    conserves: sortedAll.filter(p => conserved.has(p.productId)).map(toGlobal),
    candidats: sortedAll.filter(p => !conserved.has(p.productId)).slice(0, TOP_CANDIDATS).map(toGlobal),
  };

  return { nom: result.nom, produitsConserves, parStatut, global };
}
