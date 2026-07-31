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
  // Remises par typologie (7 valeurs, ordre TYPOLOGIES) éditées dans l'Aperçu offre.
  // Si défini, prime sur remiseStandardTaux / le gabarit (permet des remises fines par statut).
  remisesTypo?: number[];
  // Remise additionnelle (colonne I) appliquée à tous les produits du palier (ex. 0.15).
  remiseAddTaux?: number;
  // Reco % offres par typologie (7 valeurs, ordre TYPOLOGIES) propre à CE palier, calculée
  // sur les ventes N-1 du code offre du palier (qui l'a acheté l'an dernier).
  pctOffresReco?: number[];
  // Nombre TOTAL de produits (Produit Vente) que doit contenir un pack de ce palier.
  // Si défini (>0), la reco qté/pack de chaque article = ce total ventilé au prorata de la
  // conso N-1 (au lieu de conso ÷ total packs). Les UG/Testeurs/PLV sont exclus de ce total.
  nbProduitsPack?: number;
  // Descriptif libre du palier (saisi par l'utilisateur) → retranscrit dans l'Excel à côté
  // du nom du palier (colonne D de la ligne titre).
  descriptif?: string;
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
  // Grands Comptes : 6 enseignes, chacune avec sa remise et ses quantités par article (ref→qté).
  // Retranscrites dans le bloc GRANDS COMPTES du fichier Excel (une colonne qté + remise par enseigne).
  gcEnseignes?: GcEnseigne[];
  // Besoins NON B2B : canaux internes / directs (Maison Dr Hauschka, Eshop…), même structure
  // que les GC (remise + quantités par article). Colonnes ajoutées après le bloc GC dans l'Excel.
  canauxNonB2B?: GcEnseigne[];
  createdAt?: string;
}

export interface GcEnseigne {
  nom: string;                     // ex. "BIOCOOP"
  remise: number;                  // ex. 0.25
  qties: Record<string, number>;   // ref (ou clé composite) → quantité
}

// Enseignes GC par défaut (ordre = colonnes M, P, S, V, Y, AB du template).
export const GC_ENSEIGNES_DEFAUT: GcEnseigne[] = [
  { nom: "BIOCOOP", remise: 0.25, qties: {} },
  { nom: "Mlle Bio", remise: 0.20, qties: {} },
  { nom: "Galeries Lafayette", remise: 0.16, qties: {} },
  { nom: "Printemps", remise: 0.185, qties: {} },
  { nom: "Place des tendances", remise: 0.15, qties: {} },
  { nom: "NewPharma", remise: 0.20, qties: {} },
];

// Canaux NON B2B par défaut (2 premiers imposés, l'utilisateur peut en ajouter).
export const CANAUX_NONB2B_DEFAUT: GcEnseigne[] = [
  { nom: "Maison Dr Hauschka", remise: 0, qties: {} },
  { nom: "Eshop", remise: 0, qties: {} },
];

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
interface PrecoPalierLike {
  code: string; label: string; qtyPacks: number; produits: PrecoLigneLike[];
  parStatut?: Record<string, { qty: number; ca: number; nbCommandes: number }>;
}
interface PrecoLike { nom: string; paliers: PrecoPalierLike[]; }

/** Normalise un libellé de statut (casse/accents) pour le comparer aux TYPOLOGIES. */
function normStatut(s: string): string {
  return (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
}

/**
 * Convertit la ventilation par statut d'une offre N-1 en % d'offres par typologie
 * (7 valeurs, ordre TYPOLOGIES). On se base sur le NOMBRE DE COMMANDES : la question
 * métier est « quel statut a pris cette MEA », pas « qui a acheté le plus de volume ».
 * Retourne undefined si aucune typologie connue n'est représentée.
 */
export function statutsToPctOffres(parStatut?: Record<string, { qty: number; ca: number; nbCommandes: number }>): number[] | undefined {
  if (!parStatut) return undefined;
  const parNorm: Record<string, number> = {};
  for (const [name, v] of Object.entries(parStatut)) {
    parNorm[normStatut(name)] = (parNorm[normStatut(name)] || 0) + (v.nbCommandes || 0);
  }
  const vals = TYPOLOGIES.map(t => parNorm[normStatut(t)] || 0);
  const total = vals.reduce((s, n) => s + n, 0);
  if (total <= 0) return undefined;
  return vals.map(v => v / total);
}

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
    // Reco « % offres par typologie » issue des ventes N-1 de CETTE offre (qui l'a prise).
    // Si l'offre n'a pas d'historique par statut, on laisse vide : l'analyse conso N-1
    // recalculera, et à défaut le gabarit s'applique.
    const pctOffresReco = statutsToPctOffres(pal.parStatut);
    return { code: pal.code, label: pal.label, nbPacks: pal.qtyPacks || 0, qtyParPack, ...(pctOffresReco ? { pctOffresReco } : {}) };
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

  // Reco % offres par typologie : calculée PAR PALIER depuis son code offre N-1.
  // On interroge getStatutDistributionByOffer(code) pour chaque palier qui a un code,
  // ce qui donne la vraie répartition des commandes de CE palier par statut client.
  // Si un palier n'a pas de code ou renvoie 0 résultats, on tombe en fallback sur la
  // distribution globale (quantités N-1 tous articles confondus).
  const norm = (s: string) => (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
  let paliers = camp.paliers;
  if (camp.periodeDebut && camp.periodeFin && refs.length) {
    // --- Fallback global : quantités par statut sur tous les articles ---
    let globalReco: number[] | undefined;
    try {
      const vent = await odoo.getQtyByStatut(session, refs, camp.periodeDebut, camp.periodeFin);
      const qtyParStatut: Record<string, number> = {};
      for (const v of Object.values(vent)) {
        for (const [statut, q] of Object.entries(v.parStatut)) {
          qtyParStatut[norm(statut)] = (qtyParStatut[norm(statut)] || 0) + q;
        }
      }
      const qtyTypo = TYPOLOGIES.map(typo => qtyParStatut[norm(typo)] || 0);
      const totalTypo = qtyTypo.reduce((s, n) => s + n, 0);
      if (totalTypo > 0) globalReco = qtyTypo.map(q => q / totalTypo);
    } catch { /* ignore, on utilisera undefined */ }

    // --- Par palier : distribution propre au code offre ---
    paliers = await Promise.all(camp.paliers.map(async pal => {
      const code = (pal.code || "").trim();
      if (code) {
        try {
          const dist = await odoo.getStatutDistributionByOffer(session, code, camp.periodeDebut!, camp.periodeFin!);
          const qtyTypo = TYPOLOGIES.map(typo => dist[norm(typo)] || 0);
          const total = qtyTypo.reduce((s, n) => s + n, 0);
          if (total > 0) {
            return { ...pal, pctOffresReco: qtyTypo.map(q => q / total) };
          }
        } catch { /* ignore : on retombe sur la reco transférée ou le global */ }
      }
      // Reco déjà portée par le transfert depuis l'analyse (ventilation réelle de l'offre
      // N-1) : on la conserve, elle est plus fiable que la distribution globale.
      if (pal.pctOffresReco && pal.pctOffresReco.some(v => v > 0)) return pal;
      // Fallback : distribution globale si pas de code ou pas de données pour ce code.
      return globalReco ? { ...pal, pctOffresReco: globalReco } : pal;
    }));
  }

  return { ...camp, articles, paliers };
}

/** Somme des packs de tous les paliers d'une campagne (dénominateur de la reco). */
export function totalPacks(paliers: PalierSaisi[]): number {
  return paliers.reduce((s, p) => s + (p.nbPacks || 0), 0);
}

/** Un article compte-t-il dans la ventilation "N produits/pack" ? (Produit Vente uniquement) */
function estVenteArticle(a: ArticleCampagne): boolean {
  return (a.typProd || "Produit Vente") === "Produit Vente";
}

/**
 * Ventile un nombre TOTAL de produits/pack entre les articles Produit Vente d'une campagne,
 * au prorata de leur conso N-1. Renvoie une map { ref -> qté entière }. La somme des qtés
 * vaut exactement `nbTotal` grâce à la répartition du reste (méthode du plus grand reste).
 * Les articles sans conso ne reçoivent rien. Les UG/Testeurs/PLV sont exclus.
 */
export function ventilerPack(articles: ArticleCampagne[], nbTotal: number, pal?: PalierSaisi): Record<string, number> {
  const out: Record<string, number> = {};
  if (!nbTotal || nbTotal <= 0) return out;
  const ventes = articles.filter(a => a.ref.trim() && estVenteArticle(a));

  // 1) Les quantités saisies à la main sont figées : elles consomment une part de la cible.
  //    Le reste de la cible est réparti entre les seuls articles laissés en auto.
  let figees = 0;
  const auto: ArticleCampagne[] = [];
  for (const a of ventes) {
    const manual = pal ? pal.qtyParPack[qtyKeyLib(a, articles)] : undefined;
    if (manual != null && manual > 0) { out[a.ref.trim()] = manual; figees += manual; }
    else auto.push(a);
  }

  // Saisies manuelles ≥ cible : rien à répartir, les articles auto passent à 0 (l'UI
  // affichera l'écart pour que l'utilisateur corrige, sans qu'on invente des quantités).
  const restant = nbTotal - figees;
  if (restant <= 0) { for (const a of auto) out[a.ref.trim()] = 0; return out; }
  if (!auto.length) return out;

  // 2) Répartition au prorata de la conso N-1 sur les articles restants.
  //    Si aucun n'a de conso, on répartit à parts égales pour atteindre quand même la cible.
  const somme = auto.reduce((s, a) => s + (a.consoN1 || 0), 0);
  const bruts = auto.map(a => ({
    ref: a.ref.trim(),
    exact: somme > 0 ? (a.consoN1 || 0) / somme * restant : restant / auto.length,
  }));
  let attribue = 0;
  for (const b of bruts) { const base = Math.floor(b.exact); out[b.ref] = (out[b.ref] || 0) + base; attribue += base; }
  // Unités restantes aux plus grands restes fractionnaires (somme exacte = cible).
  let reste = restant - attribue;
  const parReste = [...bruts].sort((a, b) => (b.exact - Math.floor(b.exact)) - (a.exact - Math.floor(a.exact)));
  for (let i = 0; reste > 0 && parReste.length; i = (i + 1) % parReste.length) { out[parReste[i].ref] += 1; reste--; }
  return out;
}

/** Somme des qté/pack effectives des Produit Vente d'un palier (pour l'alerte d'écart UI). */
export function totalProduitsPack(articles: ArticleCampagne[], pal: PalierSaisi, totalP: number): number {
  return articles
    .filter(a => a.ref.trim() && estVenteArticle(a))
    .reduce((s, a) => s + qtyParPack(a, pal, totalP), 0);
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
/** Clé de saisie qté/pack : réf seule si unique, sinon "ref#type" (doublons vendu/UG). */
export function qtyKeyLib(art: ArticleCampagne, articles?: ArticleCampagne[]): string {
  const ref = (art.ref || "").trim();
  if (!articles) return ref;
  const doublon = articles.filter(a => (a.ref || "").trim() === ref).length > 1;
  return doublon ? `${ref}#${art.typProd || "Produit Vente"}` : ref;
}

export function qtyParPack(art: ArticleCampagne, pal: PalierSaisi, totalP?: number, ventilation?: Record<string, number>, articles?: ArticleCampagne[]): number {
  const key = qtyKeyLib(art, articles);
  const manual = pal.qtyParPack[key];
  // Mode "N produits/pack" : la ventilation fait déjà autorité — elle intègre les quantités
  // saisies à la main et ne répartit au prorata que le solde de la cible. On la lit donc
  // AVANT la saisie manuelle, sinon le total dépasserait la cible.
  if (ventilation) {
    const v = ventilation[art.ref.trim()];
    if (v != null) return v; // 0 inclus : la ventilation fait autorité en mode cible
  }
  if (manual != null && manual > 0) return manual;
  const denom = (totalP != null && totalP > 0) ? totalP : (pal.nbPacks || 0);
  const reco = denom > 0 ? Math.round((art.consoN1 || 0) / denom) : 0;
  // Si pas de reco possible (pas de conso), on respecte une éventuelle saisie 0.
  return reco > 0 ? reco : (manual ?? 0);
}

/** Ventilation à utiliser pour un palier : sa map si nbProduitsPack>0, sinon undefined. */
export function ventilationPalier(articles: ArticleCampagne[], pal: PalierSaisi): Record<string, number> | undefined {
  return (pal.nbProduitsPack && pal.nbProduitsPack > 0) ? ventilerPack(articles, pal.nbProduitsPack, pal) : undefined;
}

export interface ExportPayload {
  nom: string;
  gcEnseignes?: GcEnseigne[]; // Grands Comptes : 6 enseignes (qté + remise par article)
  canauxNonB2B?: GcEnseigne[]; // Besoins non B2B (Maison Dr Hauschka, Eshop…)
  paliers: Array<{
    code: string; label: string; qtyPacks: number; descriptif?: string;
    remiseStandard?: boolean; remiseStandardTaux?: number; remiseAddTaux?: number;
    pctOffres?: number[]; // % offres reco par typologie (7 valeurs), propre à ce palier
    remises?: number[];   // remises par typologie (7 valeurs) éditées dans l'aperçu
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
    gcEnseignes: camp.gcEnseignes,
    // Non B2B « d'office » : même sans saisie, les colonnes Maison Dr Hauschka / Eshop
    // apparaissent dans l'Excel (comme les 6 GC fixes du template).
    canauxNonB2B: camp.canauxNonB2B ?? CANAUX_NONB2B_DEFAUT,
    paliers: camp.paliers.map(pal => {
      const vent = ventilationPalier(arts, pal);
      return {
      code: pal.code,
      label: pal.label,
      qtyPacks: pal.nbPacks || 0,
      descriptif: pal.descriptif,
      remiseStandard: pal.remiseStandard,
      remiseStandardTaux: pal.remiseStandardTaux,
      remiseAddTaux: pal.remiseAddTaux,
      pctOffres: pal.pctOffresReco,
      remises: pal.remisesTypo,
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
          qtyParPack: qtyParPack(a, pal, totalP, vent, arts),
          barcode: a.barcode || "",
          standardPrice: a.standardPrice || 0,
          listPrice: estVente ? (a.listPrice || 0) : 0,
          ppc: estVente ? (a.ppc || 0) : 0,
          typProd: typ,
        };
      }),
    };
    }).filter(p => p.produits.length),
  };
}
