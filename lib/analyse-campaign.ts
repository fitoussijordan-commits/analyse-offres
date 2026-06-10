// lib/analyse-campaign.ts — Analyse d'une campagne (offres + produits + notes)
// Produit le détail par offre/note (logique identique à l'analyse d'offre historique)
// ET les totaux dédoublonnés par ligne + synthèses clients (catégorie statistique, adhérent réseau).
import * as odoo from "@/lib/odoo";
import type { Campagne, Offre } from "@/lib/campaigns";

export type StateFilter = "all" | "avenir" | "valide";

export interface ProduitCA { ref: string; name: string; productId: number; qtyVendue: number; ca: number; }
export interface DelegueCA { userId: number; name: string; qtyVendue: number; ca: number; }
export interface ClientStat { id: number; name: string; qtyVendue: number; ca: number; nbCommandes: number; }
export interface DebugOrder { id: number; name: string; partnerName?: string; invoiceStatus?: string; }

export interface OffreAnalyse {
  offre: { code: string; label: string };
  caTotal: number; qtyTotal: number;
  produits: ProduitCA[]; delegues: DelegueCA[]; debugOrders: DebugOrder[];
  error: string | null;
}
export interface CatchallResult { codeInterne: string; data: { caTotal: number; qtyTotal: number; produits: ProduitCA[]; delegues: DelegueCA[]; debugOrders: DebugOrder[] } | null; }

export interface CampaignResult {
  nom: string;
  caTotal: number; qtyTotal: number; nbCommandes: number;
  produits: ProduitCA[]; delegues: DelegueCA[];
  categories: ClientStat[]; adherents: ClientStat[];
  perOffre: { code: string; label: string; caTotal: number; qtyTotal: number; error: string | null }[];
  results: OffreAnalyse[];        // détail par offre (+ produits autonomes)
  catchalls: CatchallResult[];    // détail par note
  split?: { valide: { qty: number; ca: number }; avenir: { qty: number; ca: number } };
  error: string | null;
}

// ── Domaines ──────────────────────────────────────────────────────────────────
function orderLineDomain(f: StateFilter): any[] {
  if (f === "avenir") return [["order_id.state", "=", "sale"], ["order_id.invoice_status", "!=", "invoiced"]];
  if (f === "valide") return [["order_id.state", "in", ["sale", "done"]], ["order_id.invoice_status", "=", "invoiced"]];
  return [["order_id.state", "in", ["sale", "done"]]];
}
function orderDomain(f: StateFilter): any[] {
  if (f === "avenir") return [["state", "=", "sale"], ["invoice_status", "!=", "invoiced"]];
  if (f === "valide") return [["state", "in", ["sale", "done"]], ["invoice_status", "=", "invoiced"]];
  return [["state", "in", ["sale", "done"]]];
}
function noteDomain(notes: string[]): any[] {
  const conds = notes.map(n => ["x_note_interne", "ilike", n.trim()]);
  if (!conds.length) return [];
  return [...Array(conds.length - 1).fill("|"), ...conds];
}

interface LineRec { id: number; orderId: number; productId: number; productName: string; ref: string; qty: number; subtotal: number; }

async function resolveRefs(session: odoo.OdooSession, refs: string[]): Promise<{ ids: number[]; refByPid: Record<number, string> }> {
  const clean = [...new Set(refs.map(r => r.trim()).filter(Boolean))];
  if (!clean.length) return { ids: [], refByPid: {} };
  const prods = await odoo.searchRead(session, "product.product", [["default_code", "in", clean]], ["id", "default_code"], 0);
  const refByPid: Record<number, string> = {};
  for (const p of prods) refByPid[p.id] = p.default_code;
  return { ids: prods.map((p: any) => p.id as number), refByPid };
}

function toRecs(lines: any[], refByPid: Record<number, string>): LineRec[] {
  return lines.filter((l: any) => l.state !== "cancel" && l.product_id).map((l: any) => ({
    id: l.id,
    orderId: Array.isArray(l.order_id) ? l.order_id[0] : l.order_id,
    productId: l.product_id[0],
    productName: l.product_id[1],
    ref: refByPid[l.product_id[0]] || "",
    qty: l.product_uom_qty || 0,
    subtotal: l.price_subtotal || 0,
  }));
}

// ── Analyse d'une offre (logique historique) ──────────────────────────────────
// Retourne le détail + les lignes composants brutes (pour le dédoublonnage campagne)
async function analyseOffre(session: odoo.OdooSession, offre: Offre, filter: StateFilter): Promise<{ res: OffreAnalyse; recs: LineRec[] }> {
  const lineDom = orderLineDomain(filter);
  const base: OffreAnalyse = { offre: { code: offre.code, label: offre.label }, caTotal: 0, qtyTotal: 0, produits: [], delegues: [], debugOrders: [], error: null };

  const packProds = await odoo.searchRead(session, "product.product", [["default_code", "in", [offre.code.trim()]]], ["id"], 1);
  if (!packProds.length) return { res: { ...base, error: `Produit "${offre.code}" introuvable dans Odoo` }, recs: [] };
  const packId = packProds[0].id;

  const packLines = await odoo.searchRead(session, "sale.order.line", [["product_id", "=", packId], ...lineDom, ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_uom_qty", "state"], 0);
  const activePack = packLines.filter((l: any) => l.state !== "cancel");
  const orderIds = [...new Set(activePack.map((l: any) => l.order_id[0] as number))];
  const qtyTotal = activePack.reduce((s: number, l: any) => s + (l.product_uom_qty || 0), 0);
  if (!orderIds.length) return { res: { ...base, qtyTotal }, recs: [] };

  const { ids: compIds, refByPid } = await resolveRefs(session, offre.produits);
  let recs: LineRec[] = [];
  let produits: ProduitCA[] = [];
  let caTotal = 0;
  if (compIds.length) {
    const compLines = await odoo.searchRead(session, "sale.order.line", [["order_id", "in", orderIds], ["product_id", "in", compIds], ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
    recs = toRecs(compLines, refByPid);
    const pm: Record<number, ProduitCA> = {};
    for (const r of recs) { if (!pm[r.productId]) pm[r.productId] = { productId: r.productId, ref: r.ref, name: r.productName, qtyVendue: 0, ca: 0 }; pm[r.productId].qtyVendue += r.qty; pm[r.productId].ca += r.subtotal; }
    produits = Object.values(pm).sort((a, b) => b.ca - a.ca);
    caTotal = produits.reduce((s, p) => s + p.ca, 0);
  }

  const ords = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "name", "user_id", "partner_id", "invoice_status"], 0);
  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = ords.map((o: any) => {
    if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    return { id: o.id, name: o.name, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status };
  });

  // délégués : qté depuis les packs, CA depuis les composants
  const um: Record<number, DelegueCA> = {};
  for (const l of activePack) { const u = userByOrder[l.order_id[0]]; if (!u) continue; if (!um[u.id]) um[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; um[u.id].qtyVendue += l.product_uom_qty || 0; }
  for (const r of recs) { const u = userByOrder[r.orderId]; if (!u) continue; if (!um[u.id]) um[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; um[u.id].ca += r.subtotal; }
  const delegues = Object.values(um).sort((a, b) => b.ca - a.ca);

  return { res: { offre: { code: offre.code, label: offre.label }, caTotal, qtyTotal, produits, delegues, debugOrders, error: null }, recs };
}

// ── Analyse d'une note interne (catchall historique) ──────────────────────────
async function analyseNote(session: odoo.OdooSession, note: string, excludeOrderIds: number[], excludeOfferCodes: string[], filter: StateFilter, produitRefs: string[]): Promise<{ res: CatchallResult; recs: LineRec[] }> {
  const oDom = orderDomain(filter);
  const noteOrders = await odoo.searchRead(session, "sale.order", [["x_note_interne", "ilike", note.trim()], ...oDom], ["id", "name", "user_id", "partner_id", "invoice_status"], 0);
  const exclude = new Set(excludeOrderIds);
  let orphans = noteOrders.filter((o: any) => !exclude.has(o.id));
  if (orphans.length && excludeOfferCodes.length) {
    const offerProds = await odoo.searchRead(session, "product.product", [["default_code", "in", excludeOfferCodes]], ["id"], 0);
    const offerIds = offerProds.map((p: any) => p.id as number);
    if (offerIds.length) {
      const offLines = await odoo.searchRead(session, "sale.order.line", [["order_id", "in", orphans.map((o: any) => o.id)], ["product_id", "in", offerIds], ["display_type", "=", false]], ["order_id"], 0);
      const withOffer = new Set(offLines.map((l: any) => l.order_id[0] as number));
      orphans = orphans.filter((o: any) => !withOffer.has(o.id));
    }
  }
  if (!orphans.length) return { res: { codeInterne: note, data: { caTotal: 0, qtyTotal: 0, produits: [], delegues: [], debugOrders: [] } }, recs: [] };

  const orphanIds = orphans.map((o: any) => o.id as number);
  let filteredPids: Set<number> | null = null;
  let refByPid: Record<number, string> = {};
  if (produitRefs.length) {
    const r = await resolveRefs(session, produitRefs);
    filteredPids = new Set(r.ids); refByPid = r.refByPid;
  }
  const lines = await odoo.searchRead(session, "sale.order.line", [["order_id", "in", orphanIds], ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
  const active = lines.filter((l: any) => l.state !== "cancel" && l.price_subtotal > 0 && l.product_id && (!filteredPids || filteredPids.has(l.product_id[0])));
  const recs = toRecs(active, refByPid);
  const caTotal = recs.reduce((s, r) => s + r.subtotal, 0);

  const pm: Record<number, ProduitCA> = {};
  for (const r of recs) { if (!pm[r.productId]) pm[r.productId] = { productId: r.productId, ref: r.ref || "", name: r.productName, qtyVendue: 0, ca: 0 }; pm[r.productId].qtyVendue += r.qty; pm[r.productId].ca += r.subtotal; }
  const produits = Object.values(pm).sort((a, b) => b.ca - a.ca);

  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = orphans.map((o: any) => {
    if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    return { id: o.id, name: `${o.name} (note)`, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status };
  });
  const um: Record<number, DelegueCA> = {};
  for (const r of recs) { const u = userByOrder[r.orderId]; if (!u) continue; if (!um[u.id]) um[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; um[u.id].qtyVendue += r.qty; um[u.id].ca += r.subtotal; }
  const delegues = Object.values(um).sort((a, b) => b.ca - a.ca);

  return { res: { codeInterne: note, data: { caTotal, qtyTotal: orphanIds.length, produits, delegues, debugOrders } }, recs };
}

// ── Produits autonomes (réfs campagne hors offres/notes) ──────────────────────
async function analyseStandalone(session: odoo.OdooSession, refs: string[], filter: StateFilter): Promise<{ res: OffreAnalyse | null; recs: LineRec[] }> {
  const { ids, refByPid } = await resolveRefs(session, refs);
  if (!ids.length) return { res: null, recs: [] };
  const lineDom = orderLineDomain(filter);
  const lines = await odoo.searchRead(session, "sale.order.line", [["product_id", "in", ids], ...lineDom, ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
  const recs = toRecs(lines, refByPid);
  if (!recs.length) return { res: null, recs: [] };

  const pm: Record<number, ProduitCA> = {};
  for (const r of recs) { if (!pm[r.productId]) pm[r.productId] = { productId: r.productId, ref: r.ref, name: r.productName, qtyVendue: 0, ca: 0 }; pm[r.productId].qtyVendue += r.qty; pm[r.productId].ca += r.subtotal; }
  const produits = Object.values(pm).sort((a, b) => b.ca - a.ca);
  const caTotal = produits.reduce((s, p) => s + p.ca, 0);
  const qtyTotal = produits.reduce((s, p) => s + p.qtyVendue, 0);

  const orderIds = [...new Set(recs.map(r => r.orderId))];
  const ords = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "name", "user_id", "partner_id", "invoice_status"], 0);
  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = ords.map((o: any) => {
    if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    return { id: o.id, name: o.name, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status };
  });
  const um: Record<number, DelegueCA> = {};
  for (const r of recs) { const u = userByOrder[r.orderId]; if (!u) continue; if (!um[u.id]) um[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; um[u.id].qtyVendue += r.qty; um[u.id].ca += r.subtotal; }
  const delegues = Object.values(um).sort((a, b) => b.ca - a.ca);

  return { res: { offre: { code: "PRODUITS", label: "Produits hors offre" }, caTotal, qtyTotal, produits, delegues, debugOrders, error: null }, recs };
}

// ── Analyse complète de la campagne ───────────────────────────────────────────
export async function fetchCampaign(session: odoo.OdooSession, campagne: Campagne, configOffres: Offre[], filter: StateFilter = "all"): Promise<CampaignResult> {
  const offresCfg = campagne.offres
    .map(code => configOffres.find(o => o.code.toLowerCase() === code.toLowerCase()))
    .filter(Boolean) as Offre[];

  const allRecs: LineRec[] = [];
  const results: OffreAnalyse[] = [];

  // 1. Offres (détail + lignes pour dédoublonnage)
  const offerOrderIds = new Set<number>();
  for (const offre of offresCfg) {
    const { res, recs } = await analyseOffre(session, offre, filter);
    results.push(res);
    allRecs.push(...recs);
    for (const d of res.debugOrders) offerOrderIds.add(d.id);
  }

  // 2. Produits autonomes
  const { res: standalone, recs: standaloneRecs } = await analyseStandalone(session, campagne.produits, filter);
  if (standalone) { results.push(standalone); allRecs.push(...standaloneRecs); }

  // 3. Notes (uniquement réfs campagne : offres composants + produits autonomes)
  const campaignRefs = [...new Set([...offresCfg.flatMap(o => o.produits), ...campagne.produits])];
  const catchalls: CatchallResult[] = [];
  for (const note of campagne.notes) {
    const { res, recs } = await analyseNote(session, note, [...offerOrderIds], campagne.offres, filter, campaignRefs);
    catchalls.push(res);
    allRecs.push(...recs);
  }

  // 4. Déduplication par ligne
  const byId = new Map<number, LineRec>();
  for (const r of allRecs) if (!byId.has(r.id)) byId.set(r.id, r);
  const lines = [...byId.values()];

  const perOffre = results.map(r => ({ code: r.offre.code, label: r.offre.label, caTotal: r.caTotal, qtyTotal: r.qtyTotal, error: r.error }));

  if (!lines.length) {
    return { nom: campagne.nom, caTotal: 0, qtyTotal: 0, nbCommandes: 0, produits: [], delegues: [], categories: [], adherents: [], perOffre, results, catchalls, split: { valide: { qty: 0, ca: 0 }, avenir: { qty: 0, ca: 0 } }, error: null };
  }

  // 5. Commandes → user / partner / invoice
  const orderIds = [...new Set(lines.map(l => l.orderId))];
  const orders = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "user_id", "partner_id", "invoice_status"], 0);
  const orderUser: Record<number, { id: number; name: string }> = {};
  const orderPartner: Record<number, number> = {};
  const orderInv: Record<number, string> = {};
  const partnerIds = new Set<number>();
  for (const o of orders) {
    if (o.user_id) orderUser[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    if (o.partner_id) { orderPartner[o.id] = o.partner_id[0]; partnerIds.add(o.partner_id[0]); }
    orderInv[o.id] = o.invoice_status || "";
  }

  // 6. Clients → catégorie statistique + adhérent réseau
  const partnerCat: Record<number, { id: number; name: string }> = {};
  const partnerAdh: Record<number, { id: number; name: string }> = {};
  if (partnerIds.size) {
    try {
      const partners = await odoo.searchRead(session, "res.partner", [["id", "in", [...partnerIds]]], ["id", "x_categorie_statistique_id", "x_adherent_reseau_id"], 0);
      for (const p of partners) {
        if (p.x_categorie_statistique_id) partnerCat[p.id] = { id: p.x_categorie_statistique_id[0], name: p.x_categorie_statistique_id[1] };
        if (p.x_adherent_reseau_id) partnerAdh[p.id] = { id: p.x_adherent_reseau_id[0], name: p.x_adherent_reseau_id[1] };
      }
    } catch { /* champs clients indisponibles */ }
  }

  // 7. Agrégations dédoublonnées
  let caTotal = 0;
  const prodMap: Record<number, ProduitCA> = {};
  const delMap: Record<number, DelegueCA> = {};
  const catMap: Record<string, ClientStat & { orders: Set<number> }> = {};
  const adhMap: Record<string, ClientStat & { orders: Set<number> }> = {};
  const split = { valide: { qty: 0, ca: 0 }, avenir: { qty: 0, ca: 0 } };
  const NONE = { id: 0, name: "— Non renseigné —" };

  for (const l of lines) {
    caTotal += l.subtotal;
    if (!prodMap[l.productId]) prodMap[l.productId] = { productId: l.productId, ref: l.ref, name: l.productName, qtyVendue: 0, ca: 0 };
    prodMap[l.productId].qtyVendue += l.qty; prodMap[l.productId].ca += l.subtotal;
    const u = orderUser[l.orderId];
    if (u) { if (!delMap[u.id]) delMap[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; delMap[u.id].qtyVendue += l.qty; delMap[u.id].ca += l.subtotal; }
    const pid = orderPartner[l.orderId];
    const cat = (pid && partnerCat[pid]) || NONE;
    const ck = String(cat.id);
    if (!catMap[ck]) catMap[ck] = { id: cat.id, name: cat.name, qtyVendue: 0, ca: 0, nbCommandes: 0, orders: new Set() };
    catMap[ck].qtyVendue += l.qty; catMap[ck].ca += l.subtotal; catMap[ck].orders.add(l.orderId);
    const adh = (pid && partnerAdh[pid]) || NONE;
    const ak = String(adh.id);
    if (!adhMap[ak]) adhMap[ak] = { id: adh.id, name: adh.name, qtyVendue: 0, ca: 0, nbCommandes: 0, orders: new Set() };
    adhMap[ak].qtyVendue += l.qty; adhMap[ak].ca += l.subtotal; adhMap[ak].orders.add(l.orderId);
    if (filter === "all") {
      if (orderInv[l.orderId] === "invoiced") { split.valide.qty += l.qty; split.valide.ca += l.subtotal; }
      else { split.avenir.qty += l.qty; split.avenir.ca += l.subtotal; }
    }
  }
  const finalize = (m: Record<string, ClientStat & { orders: Set<number> }>): ClientStat[] =>
    Object.values(m).map(({ orders, ...rest }) => ({ ...rest, nbCommandes: orders.size })).sort((a, b) => b.ca - a.ca);

  // Quantité campagne = offres vendues (packs) + unités produits autonomes + commandes notées
  // (et NON la somme des unités de composants, qui gonfle le chiffre)
  const qtyOffres = results.reduce((s, r) => s + (r.qtyTotal || 0), 0) + catchalls.reduce((s, c) => s + (c.data?.qtyTotal || 0), 0);

  return {
    nom: campagne.nom,
    caTotal, qtyTotal: qtyOffres, nbCommandes: orderIds.length,
    produits: Object.values(prodMap).sort((a, b) => b.ca - a.ca),
    delegues: Object.values(delMap).sort((a, b) => b.ca - a.ca),
    categories: finalize(catMap), adherents: finalize(adhMap),
    perOffre, results, catchalls, split, error: null,
  };
}
