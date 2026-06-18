// lib/preco.ts — Préconisation d'offre N+1 (rôle trade marketing / chef de projet)
// Logique : un PALIER = un code offre de la campagne. L'utilisateur conserve des produits
// par palier ; on en déduit le besoin de commande fournisseur (quantités vendues campagne,
// agrégées par produit, tous paliers confondus).
import type { CampaignResult } from "@/lib/analyse-campaign";

export interface PrecoLigne {
  ref: string; name: string; productId: number;
  ca: number; qty: number;          // qté totale vendue du produit sur ce palier (campagne actuelle)
  qtyParPack: number;               // qté du produit dans 1 pack = round(qty / nb packs)
  conserve: boolean;
}

export interface PrecoPalier {
  code: string; label: string;      // code offre + libellé
  caTotal: number; qtyPacks: number;
  produits: PrecoLigne[];           // composants du palier, triés par CA
}

export interface BesoinFournisseur {
  ref: string; name: string; productId: number;
  qty: number;                      // quantité totale à commander (somme paliers conservés)
  ca: number;                       // CA associé (info)
  paliers: string[];               // codes offre où ce produit est conservé
}

export interface PrecoResult {
  nom: string;
  paliers: PrecoPalier[];
  besoins: BesoinFournisseur[];     // récap commande fournisseur, par produit
  totalQty: number;
}

/**
 * Construit la préconisation N+1 par palier (offre) + le besoin fournisseur.
 * @param result        résultat d'analyse de la campagne actuelle
 * @param conservedIds  ids des produits conservés, par code offre : { [code]: number[] }
 */
export function buildPreco(result: CampaignResult, conservedIds: Record<string, number[]>): PrecoResult {
  const paliers: PrecoPalier[] = result.results
    .filter(r => !r.error && (r.caTotal > 0 || r.qtyTotal > 0))
    .map(r => {
      const conserved = new Set(conservedIds[r.offre.code] ?? r.produits.map(p => p.productId));
      const nbPacks = r.qtyTotal || 0;
      const produits: PrecoLigne[] = [...r.produits]
        .sort((a, b) => b.ca - a.ca)
        .map(p => ({
          ref: p.ref, name: p.name, productId: p.productId, ca: p.ca, qty: p.qtyVendue,
          qtyParPack: nbPacks > 0 ? Math.round(p.qtyVendue / nbPacks) : 0,
          conserve: conserved.has(p.productId),
        }));
      return { code: r.offre.code, label: r.offre.label, caTotal: r.caTotal, qtyPacks: r.qtyTotal, produits };
    });

  // Besoin fournisseur : agrégation par produit des quantités CONSERVÉES, tous paliers
  const besoinMap = new Map<number, BesoinFournisseur>();
  for (const pal of paliers) {
    for (const p of pal.produits) {
      if (!p.conserve) continue;
      let b = besoinMap.get(p.productId);
      if (!b) { b = { ref: p.ref, name: p.name, productId: p.productId, qty: 0, ca: 0, paliers: [] }; besoinMap.set(p.productId, b); }
      b.qty += p.qty; b.ca += p.ca;
      if (!b.paliers.includes(pal.code)) b.paliers.push(pal.code);
    }
  }
  const besoins = [...besoinMap.values()].sort((a, b) => b.qty - a.qty);
  const totalQty = besoins.reduce((s, b) => s + b.qty, 0);

  return { nom: result.nom, paliers, besoins, totalQty };
}
