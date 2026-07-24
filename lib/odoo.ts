// lib/odoo.ts — Analyse Offres

export interface OdooConfig { url: string; db: string; }
export interface OdooSession { uid: number; name: string; login: string; sessionId: string; config: OdooConfig; }

async function rpc(config: OdooConfig, endpoint: string, params: any, sessionId?: string) {
  const res = await fetch("/api/odoo/proxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ odooUrl: config.url, endpoint, params, sessionId }),
  });
  const data = await res.json();
  if (!res.ok || data.error) throw new Error(data.error || `Erreur ${res.status}`);
  return { result: data.result, sessionId: data.sessionId };
}

export async function authenticate(config: OdooConfig, login: string, password: string): Promise<OdooSession> {
  const { result, sessionId: sid } = await rpc(config, "/web/session/authenticate", { db: config.db, login, password });
  if (!result || !result.uid || result.uid === false) throw new Error("Identifiants incorrects");
  return { uid: result.uid, name: result.name || result.username || login, login: login.toLowerCase(), sessionId: sid || result.session_id || "", config };
}

async function call(session: OdooSession, endpoint: string, params: any) {
  const { result } = await rpc(session.config, endpoint, params, session.sessionId);
  return result;
}

export async function searchRead(session: OdooSession, model: string, domain: any[], fields: string[], limit = 0, order = "", context: Record<string,any> = {}) {
  return call(session, "/web/dataset/call_kw", { model, method: "search_read", args: [domain], kwargs: { fields, limit, order, context } });
}

export interface MeaTemplate { id: number; name: string; active: boolean; }
export interface MeaTemplateLine { productCode: string; productName: string; }

/** Cherche dans les modèles de devis (actifs + archivés) par nom */
/** Normalise pour comparaison : minuscules + suppression des accents */
function norm(s: string): string {
  return (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

export async function searchMeaTemplates(session: OdooSession, query: string): Promise<MeaTemplate[]> {
  const q = query.trim();
  if (!q) return [];
  // On cherche mot par mot pour contourner les problèmes de caractères spéciaux
  const firstWord = q.split(/\s+/)[0];
  const all = await searchRead(session, "sale.order.template", [["name", "ilike", firstWord]], ["id", "name", "active"], 50, "name", { active_test: false });
  if (!all?.length) return [];
  // Filtrage côté client sur tous les mots, insensible aux accents et à la casse
  const words = norm(q).split(/\s+/).filter(Boolean);
  return (all as any[])
    .filter(r => words.every(w => norm(r.name).includes(w)))
    .map(r => ({ id: r.id, name: r.name, active: r.active !== false }));
}

/** Récupère les lignes produits d'un modèle de devis */
export async function getMeaTemplateLines(session: OdooSession, templateId: number): Promise<MeaTemplateLine[]> {
  const lines = await searchRead(session, "sale.order.template.line", [["sale_order_template_id","=",templateId],["product_id","!=",false]], ["product_id","display_type"], 0);
  const productLines = (lines||[]).filter((l:any) => !l.display_type && l.product_id);
  if (!productLines.length) return [];
  const productIds: number[] = [...new Set<number>(productLines.map((l:any) => l.product_id[0] as number))];
  const prods = await searchRead(session, "product.product", [["id","in",productIds]], ["id","default_code","name"], 0);
  const prodMap: Record<number,{code:string;name:string}> = {};
  for (const p of (prods||[]) as any[]) if (p.default_code) prodMap[p.id] = { code: p.default_code, name: p.name };
  const seen = new Set<string>();
  const result: MeaTemplateLine[] = [];
  for (const l of productLines) {
    const entry = prodMap[(l as any).product_id[0]];
    if (entry && !seen.has(entry.code)) { seen.add(entry.code); result.push({ productCode: entry.code, productName: entry.name }); }
  }
  return result;
}

// ── Données tarifaires produit (pour remplir le template Proposition) ──────────
// Champs Odoo (product.product hérite de product.template) :
//   barcode        → EAN / code à barres (char)
//   standard_price → coût d'achat unitaire (float)
//   list_price     → tarif revendeur / prix de vente (float)
//   x_ppc          → PPC, prix public conseillé (champ custom monetary)
export interface ProductPricing {
  productId: number;
  ref: string;            // default_code
  name: string;
  barcode: string;        // EAN
  standardPrice: number;  // coût achat unitaire
  listPrice: number;      // tarif revendeur unitaire
  ppc: number;            // PPC
  description?: string;   // description commerciale (description_sale), pour l'aperçu recherche
}

/**
 * Récupère, pour une liste d'IDs produits Odoo, les champs tarifaires nécessaires
 * au remplissage du template Proposition (EAN, coût achat, tarif revendeur, PPC).
 * Indexé par productId pour un mapping direct depuis la préco.
 */
export async function getProductsPricing(session: OdooSession, productIds: number[]): Promise<Record<number, ProductPricing>> {
  const ids = [...new Set(productIds.filter(Boolean))];
  const out: Record<number, ProductPricing> = {};
  if (!ids.length) return out;
  const prods = await searchRead(
    session, "product.product",
    [["id", "in", ids]],
    ["id", "default_code", "name", "barcode", "standard_price", "list_price", "x_ppc"],
    0
  );
  for (const p of (prods || []) as any[]) {
    out[p.id] = {
      productId: p.id,
      ref: p.default_code || "",
      name: p.name || "",
      barcode: p.barcode || "",
      standardPrice: typeof p.standard_price === "number" ? p.standard_price : 0,
      listPrice: typeof p.list_price === "number" ? p.list_price : 0,
      ppc: typeof p.x_ppc === "number" ? p.x_ppc : 0,
    };
  }
  return out;
}

// ── Répartition des commandes N-1 par statut client (pour la reco % offres par typologie) ──
/**
 * Pour une liste de réfs sur une période, compte le nombre de COMMANDES distinctes par statut
 * client (res.partner.x_statut_client_id). Renvoie { statutName -> nbCommandes }.
 * Sert à recommander les % offres par typologie = part de chaque statut dans les commandes N-1.
 */
export async function getStatutDistribution(session: OdooSession, refs: string[], dateFrom: string, dateTo: string): Promise<Record<string, number>> {
  const out: Record<string, number> = {};
  const clean = [...new Set(refs.map(r => (r || "").trim()).filter(Boolean))];
  if (!clean.length || !dateFrom || !dateTo) return out;

  // 1) Résoudre les réfs → ids produits.
  const prods = await searchRead(session, "product.product", [["default_code", "in", clean]], ["id"], 0);
  const ids = (prods || []).map((p: any) => p.id as number);
  if (!ids.length) return out;

  // 2) Lignes de vente confirmées sur la période → commandes distinctes.
  const lines = await searchRead(
    session, "sale.order.line",
    [
      ["product_id", "in", ids],
      ["order_id.state", "in", ["sale", "done"]],
      ["order_id.date_order", ">=", `${dateFrom} 00:00:00`],
      ["order_id.date_order", "<=", `${dateTo} 23:59:59`],
      ["display_type", "=", false],
    ],
    ["order_id"], 0
  );
  const orderIds = [...new Set((lines || []).map((l: any) => (Array.isArray(l.order_id) ? l.order_id[0] : l.order_id) as number))];
  if (!orderIds.length) return out;

  // 3) Commandes → client (partner).
  const orders = await searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "partner_id"], 0);
  const partnerByOrder: Record<number, number> = {};
  const partnerIds = new Set<number>();
  for (const o of (orders || []) as any[]) {
    if (o.partner_id) { partnerByOrder[o.id] = o.partner_id[0]; partnerIds.add(o.partner_id[0]); }
  }
  if (!partnerIds.size) return out;

  // 4) Clients → statut.
  const partners = await searchRead(session, "res.partner", [["id", "in", [...partnerIds]]], ["id", "x_statut_client_id"], 0);
  const statutByPartner: Record<number, string> = {};
  for (const p of (partners || []) as any[]) {
    if (p.x_statut_client_id) statutByPartner[p.id] = p.x_statut_client_id[1];
  }

  // 5) Compter les commandes par statut (chaque commande comptée une fois).
  for (const oid of orderIds as number[]) {
    const pid = partnerByOrder[oid];
    const statut = (pid && statutByPartner[pid]) || "";
    if (!statut) continue;
    out[statut] = (out[statut] || 0) + 1;
  }
  return out;
}

/**
 * Répartition des commandes N-1 par statut client pour UN code offre (pack) précis, sur la
 * période. Le code offre est un product.product (default_code = code offre). On compte les
 * commandes distinctes ayant acheté ce pack, ventilées par statut du client.
 * Renvoie { statutName -> nbCommandes }.
 */
export async function getStatutDistributionByOffer(session: OdooSession, offerCode: string, dateFrom: string, dateTo: string): Promise<Record<string, number>> {
  const out: Record<string, number> = {};
  const code = (offerCode || "").trim();
  if (!code || !dateFrom || !dateTo) return out;

  // 1) Pack (product.product) du code offre.
  const packs = await searchRead(session, "product.product", [["default_code", "=", code]], ["id"], 1);
  if (!packs?.length) return out;
  const packId = (packs[0] as any).id;

  // 2) Lignes de vente du pack sur la période → commandes distinctes.
  const lines = await searchRead(
    session, "sale.order.line",
    [
      ["product_id", "=", packId],
      ["order_id.state", "in", ["sale", "done"]],
      ["order_id.date_order", ">=", `${dateFrom} 00:00:00`],
      ["order_id.date_order", "<=", `${dateTo} 23:59:59`],
      ["display_type", "=", false],
    ],
    ["order_id"], 0
  );
  const orderIds = [...new Set((lines || []).map((l: any) => (Array.isArray(l.order_id) ? l.order_id[0] : l.order_id) as number))];
  if (!orderIds.length) return out;

  // 3) Commandes → client, 4) clients → statut, 5) comptage.
  const orders = await searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "partner_id"], 0);
  const partnerByOrder: Record<number, number> = {};
  const partnerIds = new Set<number>();
  for (const o of (orders || []) as any[]) if (o.partner_id) { partnerByOrder[o.id] = o.partner_id[0]; partnerIds.add(o.partner_id[0]); }
  if (!partnerIds.size) return out;
  const partners = await searchRead(session, "res.partner", [["id", "in", [...partnerIds]]], ["id", "x_statut_client_id"], 0);
  const statutByPartner: Record<number, string> = {};
  for (const p of (partners || []) as any[]) if (p.x_statut_client_id) statutByPartner[p.id] = p.x_statut_client_id[1];
  for (const oid of orderIds as number[]) {
    const statut = statutByPartner[partnerByOrder[oid]] || "";
    if (statut) out[statut] = (out[statut] || 0) + 1;
  }
  return out;
}

/** Catalogue complet : tous les product.product ayant un default_code (pour l'onglet Mapping). */
export async function getAllProducts(session: OdooSession): Promise<ProductPricing[]> {
  const prods = await searchRead(
    session, "product.product",
    [["default_code", "!=", false]],
    ["id", "default_code", "name", "barcode", "standard_price", "list_price", "x_ppc"],
    0, "default_code"
  );
  const out: ProductPricing[] = [];
  for (const p of (prods || []) as any[]) {
    if (!p.default_code) continue;
    out.push({
      productId: p.id,
      ref: p.default_code,
      name: p.name || "",
      barcode: p.barcode || "",
      standardPrice: typeof p.standard_price === "number" ? p.standard_price : 0,
      listPrice: typeof p.list_price === "number" ? p.list_price : 0,
      ppc: typeof p.x_ppc === "number" ? p.x_ppc : 0,
    });
  }
  return out;
}

/** Recherche dynamique de produits par code OU libellé (pour l'autocomplete « ajouter un
 *  article »). Renvoie au plus `limit` produits ayant un default_code, triés par pertinence
 *  (les codes exacts d'abord). La requête est un OU sur default_code / name / barcode /
 *  descriptions (commerciale + description produit). */
export async function searchProductsByQuery(session: OdooSession, query: string, limit = 20): Promise<ProductPricing[]> {
  const q = (query || "").trim();
  if (q.length < 2) return [];
  // OU à 5 branches : 4 opérateurs "|" pour 5 critères (default_code, name, barcode,
  // description commerciale, description produit).
  const domain: any[] = [
    ["default_code", "!=", false],
    "|", "|", "|", "|",
    ["default_code", "ilike", q],
    ["name", "ilike", q],
    ["barcode", "ilike", q],
    ["description_sale", "ilike", q],
    ["description", "ilike", q],
  ];
  const prods = await searchRead(
    session, "product.product", domain,
    ["id", "default_code", "name", "barcode", "standard_price", "list_price", "x_ppc", "description_sale"],
    limit, "default_code"
  );
  const out: ProductPricing[] = [];
  for (const p of (prods || []) as any[]) {
    if (!p.default_code) continue;
    out.push({
      productId: p.id,
      ref: p.default_code,
      name: p.name || "",
      barcode: p.barcode || "",
      standardPrice: typeof p.standard_price === "number" ? p.standard_price : 0,
      listPrice: typeof p.list_price === "number" ? p.list_price : 0,
      ppc: typeof p.x_ppc === "number" ? p.x_ppc : 0,
      description: typeof p.description_sale === "string" ? p.description_sale : "",
    });
  }
  // Tri pertinence : code exact > code au début > nom au début > code/nom contient >
  // match uniquement dans la description (le moins prioritaire).
  const ql = q.toLowerCase();
  return out.sort((a, b) => {
    const score = (p: ProductPricing) => {
      const c = p.ref.toLowerCase(), n = p.name.toLowerCase(), d = (p.description || "").toLowerCase();
      if (c === ql) return 0;
      if (c.startsWith(ql)) return 1;
      if (n.startsWith(ql)) return 2;
      if (c.includes(ql) || n.includes(ql)) return 3;
      if (d.includes(ql)) return 4; // trouvé seulement dans la description
      return 5;
    };
    return score(a) - score(b) || a.ref.localeCompare(b.ref);
  });
}

/** Résout des codes article (default_code) → { ref: { id, name } } via Odoo. */
export async function searchProductsByRefs(session: OdooSession, refs: string[]): Promise<Record<string, { id: number; name: string }>> {
  const out: Record<string, { id: number; name: string }> = {};
  const clean = [...new Set(refs.map(r => (r || "").trim()).filter(Boolean))];
  if (!clean.length) return out;
  const prods = await searchRead(session, "product.product", [["default_code", "in", clean]], ["id", "default_code", "name"], 0);
  for (const p of (prods || []) as any[]) {
    if (p.default_code) out[p.default_code] = { id: p.id, name: p.name || "" };
  }
  return out;
}

// ── Consommation par article sur une période (pour la création de campagne à blanc) ──
export interface ConsoArticle {
  ref: string;            // code article (default_code)
  productId: number;
  name: string;
  qty: number;            // quantité totale vendue sur la période
  ca: number;             // CA HT associé (info)
  found: boolean;         // l'article a-t-il été trouvé dans Odoo ?
}

/**
 * Somme les quantités RÉELLEMENT LIVRÉES de chaque article entre deux dates, depuis
 * l'historique des mouvements de stock (stock.move). On ne compte que les sorties vers
 * un client (location destination usage="customer"), à l'état "done" (mouvements
 * effectués). C'est la même source que l'onglet Inventaire → Historique des mouvements
 * dans Odoo, filtré sur les livraisons clients terminées.
 *
 * Note : le CA (`ca`) n'est plus renseigné ici — les mouvements de stock ne portent pas
 * de prix de vente. Le champ est conservé à 0 pour compatibilité de l'interface.
 *
 * @param refs     codes article (default_code)
 * @param dateFrom "YYYY-MM-DD" inclus
 * @param dateTo   "YYYY-MM-DD" inclus
 */
export async function getConsumption(session: OdooSession, refs: string[], dateFrom: string, dateTo: string): Promise<Record<string, ConsoArticle>> {
  const out: Record<string, ConsoArticle> = {};
  const clean = [...new Set(refs.map(r => (r || "").trim()).filter(Boolean))];
  if (!clean.length || !dateFrom || !dateTo) return out;

  // 1) Résoudre les refs → product.product (id, code, nom).
  const prods = await searchRead(session, "product.product", [["default_code", "in", clean]], ["id", "default_code", "name"], 0);
  const byId: Record<number, { ref: string; name: string }> = {};
  const ids: number[] = [];
  for (const p of (prods || []) as any[]) {
    if (!p.default_code) continue;
    byId[p.id] = { ref: p.default_code, name: p.name || "" };
    ids.push(p.id);
  }
  // Initialiser toutes les refs demandées (found=false par défaut → article inconnu d'Odoo).
  for (const ref of clean) out[ref] = { ref, productId: 0, name: "", qty: 0, ca: 0, found: false };
  for (const p of (prods || []) as any[]) {
    if (p.default_code && out[p.default_code]) {
      out[p.default_code].productId = p.id;
      out[p.default_code].name = p.name || "";
      out[p.default_code].found = true;
    }
  }
  if (!ids.length) return out;

  // 2) Mouvements de stock LIVRÉS vers un client sur la période.
  //    - state "done"                 → mouvement réellement effectué (livraison faite)
  //    - location_dest_id.usage       → "customer" : destination = emplacement client (= une sortie/livraison)
  //    - date                         → date effective du mouvement (borne la période)
  //    Quantité prise : `product_qty`. Ce champ existe sur toutes les versions d'Odoo et,
  //    une fois le mouvement à l'état "done" (filtré ci-dessus), il reflète la quantité
  //    réellement traitée (livrée). On évite ainsi les champs `qty_done`/`quantity` qui
  //    n'existent que sur certaines versions.
  const moves = await searchRead(
    session, "stock.move",
    [
      ["product_id", "in", ids],
      ["state", "=", "done"],
      ["location_dest_id.usage", "=", "customer"],
      ["date", ">=", `${dateFrom} 00:00:00`],
      ["date", "<=", `${dateTo} 23:59:59`],
    ],
    ["product_id", "product_qty", "location_id", "location_dest_id"],
    0
  );
  for (const m of (moves || []) as any[]) {
    if (!m.product_id) continue;
    const entry = byId[m.product_id[0]];
    if (!entry) continue;
    out[entry.ref].qty += (typeof m.product_qty === "number" ? m.product_qty : 0);
  }
  return out;
}

// ── Sondage conso d'un article (hors campagne) ───────────────────────────────
export interface ConsoSondage {
  ref: string; productId: number; name: string; found: boolean;
  qtyTotal: number;                        // total livré sur la période
  parMois: { mois: string; qty: number }[]; // "YYYY-MM" -> qté, trié chronologiquement
  parStatut: { statut: string; qty: number }[]; // ventilation par statut client (nb commandes non pertinent ici)
}

/**
 * Sonde la consommation d'UN article sur une période, sans l'ajouter à la campagne.
 * Sert à jauger le volume d'un produit similaire (ex. « trousse rituel apaisant » N-1
 * pour dimensionner « trousse rituel mélisse » N+1). Renvoie le total livré, la
 * répartition par mois (saisonnalité) et par statut client.
 */
export async function getConsoSondage(session: OdooSession, ref: string, dateFrom: string, dateTo: string): Promise<ConsoSondage | null> {
  const code = (ref || "").trim();
  if (!code || !dateFrom || !dateTo) return null;

  const prods = await searchRead(session, "product.product", [["default_code", "=ilike", code]], ["id", "default_code", "name"], 1);
  const prod = (prods && prods[0]) as any;
  if (!prod) return { ref: code, productId: 0, name: "", found: false, qtyTotal: 0, parMois: [], parStatut: [] };

  const pid = prod.id;
  // Mouvements de stock livrés (même source que getConsumption), avec la date pour ventiler par mois.
  const moves = await searchRead(
    session, "stock.move",
    [
      ["product_id", "=", pid],
      ["state", "=", "done"],
      ["location_dest_id.usage", "=", "customer"],
      ["date", ">=", `${dateFrom} 00:00:00`],
      ["date", "<=", `${dateTo} 23:59:59`],
    ],
    ["product_qty", "date"], 0
  );
  const moisMap: Record<string, number> = {};
  let qtyTotal = 0;
  for (const m of (moves || []) as any[]) {
    const q = typeof m.product_qty === "number" ? m.product_qty : 0;
    qtyTotal += q;
    const mois = (m.date || "").slice(0, 7); // "YYYY-MM"
    if (mois) moisMap[mois] = (moisMap[mois] || 0) + q;
  }
  const parMois = Object.entries(moisMap).map(([mois, qty]) => ({ mois, qty })).sort((a, b) => a.mois.localeCompare(b.mois));

  // Ventilation par statut client (via les lignes de vente, qui portent le partenaire).
  const parStatut: { statut: string; qty: number }[] = [];
  try {
    const vent = await getQtyByStatut(session, [code], dateFrom, dateTo);
    const v = vent[code];
    if (v) {
      const rows = Object.entries(v.parStatut).map(([statut, qty]) => ({ statut, qty }));
      if (v.qtyInconnu > 0) rows.push({ statut: "— Non renseigné —", qty: v.qtyInconnu });
      parStatut.push(...rows.sort((a, b) => b.qty - a.qty));
    }
  } catch { /* la conso reste utile même sans ventilation statut */ }

  return { ref: prod.default_code || code, productId: pid, name: prod.name || "", found: true, qtyTotal, parMois, parStatut };
}

// ── Diagnostic : ventilation des QUANTITÉS vendues par statut client ──────────
export interface VentilationStatut {
  ref: string;                       // code article
  name: string;
  qtyTotal: number;                  // total unités vendues sur la période
  parStatut: Record<string, number>; // statut client -> unités
  qtyInconnu: number;                // unités dont le client n'a pas de statut renseigné
}

/**
 * Pour chaque article, somme les quantités vendues sur la période, ventilées par statut du
 * client (x_statut_client_id). Sert à VÉRIFIER si les statuts sont exploitables avant de
 * baser la reco dessus. Renvoie aussi la part "inconnu" (client sans statut).
 */
export async function getQtyByStatut(session: OdooSession, refs: string[], dateFrom: string, dateTo: string): Promise<Record<string, VentilationStatut>> {
  const out: Record<string, VentilationStatut> = {};
  const clean = [...new Set(refs.map(r => (r || "").trim()).filter(Boolean))];
  if (!clean.length || !dateFrom || !dateTo) return out;

  // 1) refs → product.product
  const prods = await searchRead(session, "product.product", [["default_code", "in", clean]], ["id", "default_code", "name"], 0);
  const byId: Record<number, { ref: string; name: string }> = {};
  const ids: number[] = [];
  for (const p of (prods || []) as any[]) {
    if (!p.default_code) continue;
    byId[p.id] = { ref: p.default_code, name: p.name || "" };
    ids.push(p.id);
    out[p.default_code] = { ref: p.default_code, name: p.name || "", qtyTotal: 0, parStatut: {}, qtyInconnu: 0 };
  }
  if (!ids.length) return out;

  // 2) lignes de vente (avec order_id pour remonter au client) sur la période.
  const lines = await searchRead(
    session, "sale.order.line",
    [
      ["product_id", "in", ids],
      ["order_id.state", "in", ["sale", "done"]],
      ["order_id.date_order", ">=", `${dateFrom} 00:00:00`],
      ["order_id.date_order", "<=", `${dateTo} 23:59:59`],
      ["display_type", "=", false],
      ["is_downpayment", "=", false],
    ],
    ["product_id", "product_uom_qty", "order_id", "state"],
    0
  );
  const validLines = (lines || []).filter((l: any) => l.state !== "cancel" && l.product_id && l.order_id);
  const orderIds = [...new Set(validLines.map((l: any) => (Array.isArray(l.order_id) ? l.order_id[0] : l.order_id) as number))];
  if (!orderIds.length) return out;

  // 3) commandes → client
  const orders = await searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "partner_id"], 0);
  const partnerByOrder: Record<number, number> = {};
  const partnerIds = new Set<number>();
  for (const o of (orders || []) as any[]) if (o.partner_id) { partnerByOrder[o.id] = o.partner_id[0]; partnerIds.add(o.partner_id[0]); }

  // 4) clients → statut
  const partners = partnerIds.size
    ? await searchRead(session, "res.partner", [["id", "in", [...partnerIds]]], ["id", "x_statut_client_id"], 0)
    : [];
  const statutByPartner: Record<number, string> = {};
  for (const p of (partners || []) as any[]) if (p.x_statut_client_id) statutByPartner[p.id] = p.x_statut_client_id[1];

  // 5) sommer les quantités par statut
  for (const l of validLines as any[]) {
    const entry = byId[l.product_id[0]];
    if (!entry) continue;
    const v = out[entry.ref];
    const qty = l.product_uom_qty || 0;
    v.qtyTotal += qty;
    const oid = Array.isArray(l.order_id) ? l.order_id[0] : l.order_id;
    const statut = statutByPartner[partnerByOrder[oid]] || "";
    if (statut) v.parStatut[statut] = (v.parStatut[statut] || 0) + qty;
    else v.qtyInconnu += qty;
  }
  return out;
}
