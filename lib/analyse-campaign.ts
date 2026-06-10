// lib/analyse-campaign.ts — Analyse d'une campagne (offres + produits + notes), dédoublonnée par ligne
import * as odoo from "@/lib/odoo";
import type { Campagne, Offre } from "@/lib/campaigns";

export type StateFilter = "all" | "avenir" | "valide";

export interface ProduitCA { ref: string; name: string; productId: number; qtyVendue: number; ca: number; }
export interface DelegueCA { userId: number; name: string; qtyVendue: number; ca: number; }
export interface ClientStat { id: number; name: string; qtyVendue: number; ca: number; nbCommandes: number; }
export interface OffreBreakdown { code: string; label: string; caTotal: number; qtyTotal: number; error: string | null; }

export interface CampaignResult {
  nom: string;
  caTotal: number;
  qtyTotal: number;
  nbCommandes: number;
  produits: ProduitCA[];
  delegues: DelegueCA[];
  categories: ClientStat[];   // x_categorie_statistique_id
  adherents: ClientStat[];    // x_adherent_reseau_id
  perOffre: OffreBreakdown[];
  split?: { valide: { qty: number; ca: number }; avenir: { qty: number; ca: number } };
  error: string | null;
}

// ── Domaines selon le filtre d'état ───────────────────────────────────────────
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

// OR de plusieurs conditions ilike sur x_note_interne
function noteDomain(notes: string[]): any[] {
  const conds = notes.map(n => ["x_note_interne", "ilike", n.trim()]);
  if (conds.length === 0) return [];
  const prefix = Array(conds.length - 1).fill("|");
  return [...prefix, ...conds];
}

interface LineRec { id: number; orderId: number; productId: number; productName: string; qty: number; subtotal: number; }

function toRecs(lines: any[]): LineRec[] {
  return lines
    .filter((l: any) => l.state !== "cancel" && l.product_id)
    .map((l: any) => ({
      id: l.id,
      orderId: Array.isArray(l.order_id) ? l.order_id[0] : l.order_id,
      productId: l.product_id[0],
      productName: l.product_id[1],
      qty: l.product_uom_qty || 0,
      subtotal: l.price_subtotal || 0,
    }));
}

// Résout des références produits → map productId -> ref
async function resolveRefs(session: odoo.OdooSession, refs: string[]): Promise<{ ids: number[]; refByPid: Record<number, string> }> {
  const clean = [...new Set(refs.map(r => r.trim()).filter(Boolean))];
  if (!clean.length) return { ids: [], refByPid: {} };
  const prods = await odoo.searchRead(session, "product.product", [["default_code", "in", clean]], ["id", "default_code"], 0);
  const refByPid: Record<number, string> = {};
  for (const p of prods) refByPid[p.id] = p.default_code;
  return { ids: prods.map((p: any) => p.id as number), refByPid };
}

export async function fetchCampaign(
  session: odoo.OdooSession,
  campagne: Campagne,
  configOffres: Offre[],
  filter: StateFilter = "all"
): Promise<CampaignResult> {
  const lineDom = orderLineDomain(filter);
  const empty: CampaignResult = {
    nom: campagne.nom, caTotal: 0, qtyTotal: 0, nbCommandes: 0,
    produits: [], delegues: [], categories: [], adherents: [], perOffre: [],
    split: { valide: { qty: 0, ca: 0 }, avenir: { qty: 0, ca: 0 } }, error: null,
  };

  // Offres de la campagne → leur config (codes composants)
  const offresCfg = campagne.offres
    .map(code => configOffres.find(o => o.code.toLowerCase() === code.toLowerCase()))
    .filter(Boolean) as Offre[];

  // 1. Résolution des produits "composants" (offres) + réfs autonomes
  const offerCompRefs = offresCfg.flatMap(o => o.produits);
  const { ids: compIds, refByPid: compRefMap } = await resolveRefs(session, offerCompRefs);
  const { ids: standaloneIds, refByPid: standaloneRefMap } = await resolveRefs(session, campagne.produits);
  const refByPid: Record<number, string> = { ...compRefMap, ...standaloneRefMap };
  const campaignProductIds = [...new Set([...compIds, ...standaloneIds])];

  // 2. Résolution des produits "pack" (codes d'offres) → commandes par offre
  const { ids: packIds } = await resolveRefs(session, offresCfg.map(o => o.code));
  // map code -> packId
  const packByCode: Record<string, number> = {};
  if (offresCfg.length) {
    const packProds = await odoo.searchRead(session, "product.product", [["default_code", "in", offresCfg.map(o => o.code)]], ["id", "default_code"], 0);
    for (const p of packProds) packByCode[(p.default_code || "").toLowerCase()] = p.id;
  }

  const allLines: LineRec[] = [];
  const perOffre: OffreBreakdown[] = [];

  // 3. Lignes des offres : composants à l'intérieur des commandes contenant le pack
  for (const offre of offresCfg) {
    const packId = packByCode[offre.code.toLowerCase()];
    if (!packId) { perOffre.push({ code: offre.code, label: offre.label, caTotal: 0, qtyTotal: 0, error: `Produit "${offre.code}" introuvable` }); continue; }
    const packLines = await odoo.searchRead(session, "sale.order.line", [["product_id", "=", packId], ...lineDom, ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "state"], 0);
    const orderIds = [...new Set(packLines.filter((l: any) => l.state !== "cancel").map((l: any) => l.order_id[0] as number))];
    const offCompIds = (await resolveRefs(session, offre.produits)).ids;
    let offCa = 0, offQty = 0;
    if (orderIds.length && offCompIds.length) {
      const compLines = await odoo.searchRead(session, "sale.order.line", [["order_id", "in", orderIds], ["product_id", "in", offCompIds], ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
      const recs = toRecs(compLines);
      for (const r of recs) { offCa += r.subtotal; offQty += r.qty; }
      allLines.push(...recs);
    }
    perOffre.push({ code: offre.code, label: offre.label, caTotal: offCa, qtyTotal: offQty, error: null });
  }

  // 4. Réfs produits autonomes : lignes partout (selon filtre)
  if (standaloneIds.length) {
    const lines = await odoo.searchRead(session, "sale.order.line", [["product_id", "in", standaloneIds], ...lineDom, ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
    allLines.push(...toRecs(lines));
  }

  // 5. Notes internes : commandes notées → uniquement lignes des produits de la campagne
  if (campagne.notes.length && campaignProductIds.length) {
    const notedOrders = await odoo.searchRead(session, "sale.order", [...noteDomain(campagne.notes), ...orderDomain(filter)], ["id"], 0);
    const notedIds = notedOrders.map((o: any) => o.id as number);
    if (notedIds.length) {
      const lines = await odoo.searchRead(session, "sale.order.line", [["order_id", "in", notedIds], ["product_id", "in", campaignProductIds], ["display_type", "=", false], ["is_downpayment", "=", false]], ["order_id", "product_id", "product_uom_qty", "price_subtotal", "state"], 0);
      allLines.push(...toRecs(lines));
    }
  }

  // 6. Déduplication par id de ligne
  const lineById = new Map<number, LineRec>();
  for (const r of allLines) if (!lineById.has(r.id)) lineById.set(r.id, r);
  const lines = [...lineById.values()];
  if (!lines.length) return { ...empty, perOffre };

  // 7. Commandes concernées → user_id, partner_id, invoice_status
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

  // 8. Clients → catégorie statistique + adhérent réseau
  const partnerCat: Record<number, { id: number; name: string }> = {};
  const partnerAdh: Record<number, { id: number; name: string }> = {};
  if (partnerIds.size) {
    try {
      const partners = await odoo.searchRead(session, "res.partner", [["id", "in", [...partnerIds]]], ["id", "x_categorie_statistique_id", "x_adherent_reseau_id"], 0);
      for (const p of partners) {
        if (p.x_categorie_statistique_id) partnerCat[p.id] = { id: p.x_categorie_statistique_id[0], name: p.x_categorie_statistique_id[1] };
        if (p.x_adherent_reseau_id) partnerAdh[p.id] = { id: p.x_adherent_reseau_id[0], name: p.x_adherent_reseau_id[1] };
      }
    } catch { /* champs clients indisponibles : synthèses clients vides, le reste de l'analyse continue */ }
  }

  // 9. Agrégations
  let caTotal = 0, qtyTotal = 0;
  const prodMap: Record<number, ProduitCA> = {};
  const delMap: Record<number, DelegueCA> = {};
  const catMap: Record<string, ClientStat & { orders: Set<number> }> = {};
  const adhMap: Record<string, ClientStat & { orders: Set<number> }> = {};
  const split = { valide: { qty: 0, ca: 0 }, avenir: { qty: 0, ca: 0 } };

  const NONE = { id: 0, name: "— Non renseigné —" };

  for (const l of lines) {
    caTotal += l.subtotal; qtyTotal += l.qty;
    // produits
    if (!prodMap[l.productId]) prodMap[l.productId] = { productId: l.productId, ref: refByPid[l.productId] || "", name: l.productName, qtyVendue: 0, ca: 0 };
    prodMap[l.productId].qtyVendue += l.qty; prodMap[l.productId].ca += l.subtotal;
    // délégués
    const u = orderUser[l.orderId];
    if (u) { if (!delMap[u.id]) delMap[u.id] = { userId: u.id, name: u.name, qtyVendue: 0, ca: 0 }; delMap[u.id].qtyVendue += l.qty; delMap[u.id].ca += l.subtotal; }
    // client : catégorie
    const pid = orderPartner[l.orderId];
    const cat = (pid && partnerCat[pid]) || NONE;
    const ck = String(cat.id);
    if (!catMap[ck]) catMap[ck] = { id: cat.id, name: cat.name, qtyVendue: 0, ca: 0, nbCommandes: 0, orders: new Set() };
    catMap[ck].qtyVendue += l.qty; catMap[ck].ca += l.subtotal; catMap[ck].orders.add(l.orderId);
    // client : adhérent
    const adh = (pid && partnerAdh[pid]) || NONE;
    const ak = String(adh.id);
    if (!adhMap[ak]) adhMap[ak] = { id: adh.id, name: adh.name, qtyVendue: 0, ca: 0, nbCommandes: 0, orders: new Set() };
    adhMap[ak].qtyVendue += l.qty; adhMap[ak].ca += l.subtotal; adhMap[ak].orders.add(l.orderId);
    // split validé / à venir
    if (filter === "all") {
      if (orderInv[l.orderId] === "invoiced") { split.valide.qty += l.qty; split.valide.ca += l.subtotal; }
      else { split.avenir.qty += l.qty; split.avenir.ca += l.subtotal; }
    }
  }

  const finalize = (m: Record<string, ClientStat & { orders: Set<number> }>): ClientStat[] =>
    Object.values(m).map(({ orders, ...rest }) => ({ ...rest, nbCommandes: orders.size })).sort((a, b) => b.ca - a.ca);

  return {
    nom: campagne.nom,
    caTotal,
    qtyTotal,
    nbCommandes: orderIds.length,
    produits: Object.values(prodMap).sort((a, b) => b.ca - a.ca),
    delegues: Object.values(delMap).sort((a, b) => b.ca - a.ca),
    categories: finalize(catMap),
    adherents: finalize(adhMap),
    perOffre,
    split,
    error: null,
  };
}
