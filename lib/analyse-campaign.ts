// lib/analyse-campaign.ts — Analyse d'une campagne (offres + produits + notes)
// Produit le détail par offre/note (logique identique à l'analyse d'offre historique)
// ET les totaux dédoublonnés par ligne + synthèses clients (catégorie statistique, adhérent réseau).
import * as odoo from "@/lib/odoo";
import type { Campagne, Offre } from "@/lib/campaigns";

export type StateFilter = "all" | "avenir" | "valide";

export interface ProduitCA { ref: string; name: string; productId: number; qtyVendue: number; ca: number; }
export interface DelegueCA { userId: number; name: string; qtyVendue: number; ca: number; }
export interface ClientStat { id: number; name: string; qtyVendue: number; ca: number; nbCommandes: number; }
export interface DebugOrder { id: number; name: string; partnerName?: string; invoiceStatus?: string; ca?: number; invoiced?: boolean; }

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

/**
 * Détermine les commandes réellement facturées : celles ayant au moins une facture
 * (account.move, type out_invoice/out_refund) en état 'posted'. Plus fiable que le
 * champ invoice_status de la commande, qui ne reflète pas toujours la facturation réelle.
 * Retourne un Set des orderId facturés.
 */
async function getInvoicedOrderIds(session: odoo.OdooSession, orderIds: number[]): Promise<Set<number>> {
  const invoiced = new Set<number>();
  if (!orderIds.length) return invoiced;
  // On lit invoice_ids sur les commandes
  const orders = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "invoice_ids"], 0);
  const invToOrders: Record<number, number[]> = {};
  const allInvIds = new Set<number>();
  for (const o of orders as any[]) {
    for (const invId of (o.invoice_ids || [])) {
      allInvIds.add(invId);
      (invToOrders[invId] ||= []).push(o.id);
    }
  }
  if (!allInvIds.size) return invoiced;
  // On ne garde que les factures clients postées (validées)
  const moves = await odoo.searchRead(
    session, "account.move",
    [["id", "in", [...allInvIds]], ["state", "=", "posted"], ["move_type", "in", ["out_invoice", "out_refund"]]],
    ["id"], 0
  );
  for (const m of moves as any[]) for (const oid of (invToOrders[m.id] || [])) invoiced.add(oid);
  return invoiced;
}

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

  const caByOrder: Record<number, number> = {};
  for (const r of recs) caByOrder[r.orderId] = (caByOrder[r.orderId] || 0) + r.subtotal;
  const ords = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "name", "user_id", "partner_id", "invoice_status"], 0);
  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = ords.map((o: any) => {
    if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    return { id: o.id, name: o.name, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status, ca: caByOrder[o.id] || 0 };
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

  // On ne retient que les commandes ayant au moins une ligne dans le périmètre campagne :
  // une commande notée sans aucun produit de la campagne ne doit pas être comptée.
  const matchedOrderIds = new Set(recs.map(r => r.orderId));
  const caByOrder: Record<number, number> = {};
  for (const r of recs) caByOrder[r.orderId] = (caByOrder[r.orderId] || 0) + r.subtotal;
  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = orphans
    .filter((o: any) => matchedOrderIds.has(o.id))
    .map((o: any) => {
      if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
      return { id: o.id, name: `${o.name} (note)`, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status, ca: caByOrder[o.id] || 0 };
    });
  const um: Record<number, DelegueCA> = {};
  for (const r of recs) { const u = userByOrder[r.orderId]; if (!u) continue; if (!um[u.id]) um[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; um[u.id].qtyVendue += r.qty; um[u.id].ca += r.subtotal; }
  const delegues = Object.values(um).sort((a, b) => b.ca - a.ca);

  return { res: { codeInterne: note, data: { caTotal, qtyTotal: matchedOrderIds.size, produits, delegues, debugOrders } }, recs };
}

// ── Produits autonomes (réfs campagne hors offres/notes) ──────────────────────
async function analyseStandalone(session: odoo.OdooSession, refs: string[], filter: StateFilter, excludeLineIds: Set<number>): Promise<{ res: OffreAnalyse | null; recs: LineRec[] }> {
  const { ids, refByPid } = await resolveRefs(session, refs);
  if (!ids.length) return { res: null, recs: [] };
  const lineDom = orderLineDomain(filter);
  const lines = await odoo.searchRead(session, "sale.order.line", [["product_id", "in", ids], ...lineDom, ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
  // exclure les lignes déjà comptées dans une offre (dédoublonnage "hors offre")
  const recs = toRecs(lines, refByPid).filter(r => !excludeLineIds.has(r.id));
  if (!recs.length) return { res: null, recs: [] };

  const pm: Record<number, ProduitCA> = {};
  for (const r of recs) { if (!pm[r.productId]) pm[r.productId] = { productId: r.productId, ref: r.ref, name: r.productName, qtyVendue: 0, ca: 0 }; pm[r.productId].qtyVendue += r.qty; pm[r.productId].ca += r.subtotal; }
  const produits = Object.values(pm).sort((a, b) => b.ca - a.ca);
  const caTotal = produits.reduce((s, p) => s + p.ca, 0);
  const qtyTotal = produits.reduce((s, p) => s + p.qtyVendue, 0);

  const orderIds = [...new Set(recs.map(r => r.orderId))];
  const caByOrder: Record<number, number> = {};
  for (const r of recs) caByOrder[r.orderId] = (caByOrder[r.orderId] || 0) + r.subtotal;
  const ords = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "name", "user_id", "partner_id", "invoice_status"], 0);
  const userByOrder: Record<number, { id: number; name: string }> = {};
  const debugOrders: DebugOrder[] = ords.map((o: any) => {
    if (o.user_id) userByOrder[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    return { id: o.id, name: o.name, partnerName: o.partner_id ? o.partner_id[1] : undefined, invoiceStatus: o.invoice_status, ca: caByOrder[o.id] || 0 };
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

  // 2. Produits autonomes (en excluant les lignes déjà comptées dans une offre)
  const offerLineIds = new Set<number>(allRecs.map(r => r.id));
  const { res: standalone, recs: standaloneRecs } = await analyseStandalone(session, campagne.produits, filter, offerLineIds);
  if (standalone) { results.push(standalone); allRecs.push(...standaloneRecs); }

  // 3. Notes : CA des commandes rattachées par note interne (hors commandes déjà comptées
  //    dans une offre), MAIS restreint aux seuls produits de la campagne. On passe donc la
  //    liste des références campagne (codes d'offres + produits autonomes) pour exclure les
  //    produits hors périmètre présents dans ces commandes.
  const campaignRefs = [...new Set([...campagne.offres, ...campagne.produits].map(r => r.trim()).filter(Boolean))];
  const catchalls: CatchallResult[] = [];
  for (const note of campagne.notes) {
    const { res, recs } = await analyseNote(session, note, [...offerOrderIds], campagne.offres, filter, campaignRefs);
    catchalls.push(res);
    allRecs.push(...recs);
  }

  // 4. Déduplication par ligne. On isole les commandes notées : leur CA sera rattaché
  //    en bloc plus bas (via catchalls), donc on les retire ici pour éviter tout double
  //    comptage ou comptage partiel.
  const noteOrderIdSet = new Set<number>();
  for (const c of catchalls) for (const o of c.data?.debugOrders ?? []) noteOrderIdSet.add(o.id);
  const byId = new Map<number, LineRec>();
  for (const r of allRecs) if (!byId.has(r.id)) byId.set(r.id, r);
  const lines = [...byId.values()].filter(l => !noteOrderIdSet.has(l.orderId));

  const perOffre = results.map(r => ({ code: r.offre.code, label: r.offre.label, caTotal: r.caTotal, qtyTotal: r.qtyTotal, error: r.error }));

  if (!lines.length) {
    return { nom: campagne.nom, caTotal: 0, qtyTotal: 0, nbCommandes: 0, produits: [], delegues: [], categories: [], adherents: [], perOffre, results, catchalls, split: { valide: { qty: 0, ca: 0 }, avenir: { qty: 0, ca: 0 } }, error: null };
  }

  // 5. Commandes → user / partner / invoice
  const orderIds = [...new Set(lines.map(l => l.orderId))];
  const orders = await odoo.searchRead(session, "sale.order", [["id", "in", orderIds]], ["id", "user_id", "partner_id", "invoice_status"], 0);
  const orderUser: Record<number, { id: number; name: string }> = {};
  const orderPartner: Record<number, number> = {};
  const partnerIds = new Set<number>();
  for (const o of orders) {
    if (o.user_id) orderUser[o.id] = { id: o.user_id[0], name: o.user_id[1] };
    if (o.partner_id) { orderPartner[o.id] = o.partner_id[0]; partnerIds.add(o.partner_id[0]); }
  }

  // Facturation réelle : commandes avec une facture posée (toutes commandes confondues,
  // lignes + notes), pour un découpage validé/à venir fiable.
  const debugOrderIds: number[] = [];
  for (const r of results) for (const o of r.debugOrders) debugOrderIds.push(o.id);
  for (const c of catchalls) for (const o of c.data?.debugOrders ?? []) debugOrderIds.push(o.id);
  const allOrderIds = [...new Set([...orderIds, ...noteOrderIdSet, ...debugOrderIds])];
  const invoicedOrderIds = await getInvoicedOrderIds(session, allOrderIds);
  const isInvoiced = (oid: number) => invoicedOrderIds.has(oid);

  // On propage le statut de facturation réel sur les debugOrders (utilisés par l'UI).
  for (const r of results) for (const o of r.debugOrders) o.invoiced = isInvoiced(o.id);
  for (const c of catchalls) for (const o of c.data?.debugOrders ?? []) o.invoiced = isInvoiced(o.id);

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
      if (isInvoiced(l.orderId)) { split.valide.qty += l.qty; split.valide.ca += l.subtotal; }
      else { split.avenir.qty += l.qty; split.avenir.ca += l.subtotal; }
    }
  }
  const finalize = (m: Record<string, ClientStat & { orders: Set<number> }>): ClientStat[] =>
    Object.values(m).map(({ orders, ...rest }) => ({ ...rest, nbCommandes: orders.size })).sort((a, b) => b.ca - a.ca);

  // Les commandes notées sont orphelines (analyseNote les a déjà exclues des offres),
  // donc leur CA n'est jamais dans `lines`. On le rattache systématiquement au total
  // et aux détails, en ne dédoublonnant que les lignes réellement déjà comptées.
  const lineIds = new Set(lines.map(l => l.id));
  let extraNoteOrders = 0;
  for (const c of catchalls) {
    if (!c.data || c.data.caTotal <= 0) continue;
    extraNoteOrders += c.data.debugOrders.length;
    caTotal += c.data.caTotal;
    for (const p of c.data.produits) {
      if (!prodMap[p.productId]) prodMap[p.productId] = { productId: p.productId, ref: p.ref, name: p.name, qtyVendue: 0, ca: 0 };
      prodMap[p.productId].qtyVendue += p.qtyVendue; prodMap[p.productId].ca += p.ca;
    }
    for (const d of c.data.delegues) {
      if (!delMap[d.userId]) delMap[d.userId] = { userId: d.userId, name: d.name, qtyVendue: 0, ca: 0 };
      delMap[d.userId].qtyVendue += d.qtyVendue; delMap[d.userId].ca += d.ca;
    }
    if (filter === "all") {
      // Répartition validé / à venir par commande notée selon sa facturation réelle
      for (const o of c.data.debugOrders) {
        const ca = o.ca ?? 0;
        if (isInvoiced(o.id)) { split.valide.qty += 1; split.valide.ca += ca; }
        else { split.avenir.qty += 1; split.avenir.ca += ca; }
      }
    }
  }
  void lineIds;

  // Le CA total doit toujours égaler la somme du détail affiché (offres + standalone + notes).
  // On le recalcule explicitement pour éviter toute divergence avec les lignes agrégées.
  const caDetail = results.reduce((s, r) => s + (r.caTotal || 0), 0) + catchalls.reduce((s, c) => s + (c.data?.caTotal || 0), 0);
  caTotal = caDetail;

  // Le split validé/à venir doit lui aussi sommer au CA total. On rattache au "à venir"
  // l'écart non couvert par l'agrégation par ligne (notamment le CA des notes).
  if (filter === "all") {
    const splitSum = split.valide.ca + split.avenir.ca;
    const missing = caTotal - splitSum;
    if (Math.abs(missing) > 0.5) split.avenir.ca += missing;
  }

  // Quantité campagne = offres vendues (packs) + unités produits autonomes + commandes notées
  // (et NON la somme des unités de composants, qui gonfle le chiffre)
  const qtyOffres = results.reduce((s, r) => s + (r.qtyTotal || 0), 0) + catchalls.reduce((s, c) => s + (c.data?.qtyTotal || 0), 0);

  return {
    nom: campagne.nom,
    caTotal, qtyTotal: qtyOffres, nbCommandes: orderIds.length + extraNoteOrders,
    produits: Object.values(prodMap).sort((a, b) => b.ca - a.ca),
    delegues: Object.values(delMap).sort((a, b) => b.ca - a.ca),
    categories: finalize(catMap), adherents: finalize(adhMap),
    perOffre, results, catchalls, split, error: null,
  };
}
