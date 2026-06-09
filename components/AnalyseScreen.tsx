"use client";
import { useState, useRef, useEffect, useCallback } from "react";
import * as odoo from "@/lib/odoo";
import { supabase } from "@/lib/supabase";

// ── Design tokens ─────────────────────────────────────────────────────────────
const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0", borderFocus: "#3b82f6",
  blue: "#3b82f6", blueDark: "#1d4ed8", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  red: "#ef4444", redSoft: "#fef2f2",
  purple: "#8b5cf6", purpleSoft: "#f5f3ff",
  slate: "#475569",
  shadow: "0 1px 3px rgba(0,0,0,0.06)", shadowMd: "0 4px 16px rgba(0,0,0,0.10)",
};

// ── Types ─────────────────────────────────────────────────────────────────────
interface Offre { id: string; code: string; label: string; produits: string[]; codeInterne?: string; }
interface ProduitCA { ref: string; name: string; productId: number; qtyVendue: number; ca: number; }
interface DelegueCA { userId: number; name: string; qtyVendue: number; ca: number; }
interface OffreAnalyse {
  offre: Offre; loading: boolean; error: string | null;
  caTotal: number; qtyTotal: number;
  produits: ProduitCA[]; delegues: DelegueCA[];
  debugOrders?: { id: number; name: string; partnerName?: string; invoiceStatus?: string }[];
  split?: { valide: { qty: number; ca: number }; avenir: { qty: number; ca: number } };
}
interface Props { session: odoo.OdooSession; onToast: (msg: string, type?: "success" | "error" | "info") => void; }

// ── Extraction refs ───────────────────────────────────────────────────────────
function extractRefs(text: string): string[] {
  const numRefs = text.match(/\d{5,}/g);
  if (numRefs && numRefs.length >= 2) return [...new Set(numRefs)];
  const alphaRefs = text.match(/\b(?=[A-Z0-9]*\d)[A-Z0-9-_]{4,}\b/gi);
  if (alphaRefs && alphaRefs.length >= 1) return [...new Set(alphaRefs)];
  return text.split(/[\n\r,;]+/).map(r => r.trim()).filter(Boolean);
}

// ── LocalStorage ──────────────────────────────────────────────────────────────
// ── Supabase CRUD offres ──────────────────────────────────────────────────────
function rowToOffre(row: any): Offre {
  return { id: row.id, code: row.code, label: row.label || "", produits: row.produits || [], codeInterne: row.code_interne || undefined };
}
function offreToRow(o: Offre) {
  return { id: o.id, code: o.code, label: o.label || "", produits: o.produits, code_interne: o.codeInterne || null };
}
async function dbLoadOffres(): Promise<Offre[]> {
  const { data, error } = await supabase.from("analyse_offres").select("*").order("created_at");
  if (error) throw error;
  return (data || []).map(rowToOffre);
}
async function dbUpsertOffre(o: Offre): Promise<void> {
  const { error } = await supabase.from("analyse_offres").upsert(offreToRow(o));
  if (error) throw error;
}
async function dbDeleteOffre(id: string): Promise<void> {
  const { error } = await supabase.from("analyse_offres").delete().eq("id", id);
  if (error) throw error;
}
// Migration localStorage → Supabase (one-shot)
const LS_OFFRES_KEY = "ao_offres_config";
const LS_OFFRES_MIGRATED = "ao_offres_migrated_v1";
async function migrateOffresFromLS(): Promise<number> {
  if (typeof window === "undefined") return 0;
  if (localStorage.getItem(LS_OFFRES_MIGRATED)) return 0;
  const raw = localStorage.getItem(LS_OFFRES_KEY);
  if (!raw) { localStorage.setItem(LS_OFFRES_MIGRATED, "1"); return 0; }
  try {
    const offres: Offre[] = JSON.parse(raw);
    if (!offres.length) { localStorage.setItem(LS_OFFRES_MIGRATED, "1"); return 0; }
    const { error } = await supabase.from("analyse_offres").upsert(offres.map(offreToRow), { onConflict: "id" });
    if (error) throw error;
    localStorage.setItem(LS_OFFRES_MIGRATED, "1");
    return offres.length;
  } catch(e) { console.error("Migration offres failed", e); return 0; }
}
function genId() { return Math.random().toString(36).slice(2, 10); }

// ── Filtre Odoo ───────────────────────────────────────────────────────────────
type StateFilter = "all" | "avenir" | "valide";
function filterDomain(f: StateFilter): any[] {
  if (f === "avenir") return [["state","=","sale"],["invoice_status","!=","invoiced"]];
  if (f === "valide") return [["state","in",["sale","done"]],["invoice_status","=","invoiced"]];
  return [["state","in",["sale","done"]]];
}
function orderLineDomain(f: StateFilter): any[] {
  if (f === "avenir") return [["order_id.state","=","sale"],["order_id.invoice_status","!=","invoiced"]];
  if (f === "valide") return [["order_id.state","in",["sale","done"]],["order_id.invoice_status","=","invoiced"]];
  return [["order_id.state","in",["sale","done"]]];
}

// ── fetchCAForOffre (identique) ───────────────────────────────────────────────
async function fetchCAForOffre(session: odoo.OdooSession, offre: Offre, filter: StateFilter = "all"): Promise<Omit<OffreAnalyse,"offre"|"loading">> {
  const lineDomain = orderLineDomain(filter);
  let offreProd = await odoo.searchRead(session,"product.product",[["default_code","=ilike",offre.code.trim()]],["id","name","default_code"],1);
  if (!offreProd.length) offreProd = await odoo.searchRead(session,"product.product",[["default_code","ilike",offre.code.trim()]],["id","name","default_code"],1);
  if (!offreProd.length) return { caTotal:0, qtyTotal:0, produits:[], delegues:[], debugOrders:[], error:`Produit "${offre.code}" introuvable dans Odoo` };
  const offreProductId = offreProd[0].id;
  const offreLines = await odoo.searchRead(session,"sale.order.line",[["product_id","=",offreProductId],...lineDomain,["display_type","=",false],["is_downpayment","=",false]],["order_id","product_uom_qty","state"],0);
  const activeOffreLines = offreLines.filter((l:any) => l.state !== "cancel");
  const orderIdsFromLines = new Set<number>(activeOffreLines.map((l:any) => l.order_id[0] as number));
  const orderIds = Array.from(orderIdsFromLines) as number[];
  const qtyTotal = activeOffreLines.reduce((s:number,l:any) => s+(l.product_uom_qty||0),0);
  if (!orderIds.length) return { caTotal:0, qtyTotal:0, produits:[], delegues:[], debugOrders:[], error:null };
  const ords = await odoo.searchRead(session,"sale.order",[["id","in",orderIds]],["id","name","partner_id","invoice_status"],orderIds.length);
  const orderInvoiceMap: Record<number,string> = {};
  for (const o of ords) orderInvoiceMap[o.id] = o.invoice_status ?? "";
  const debugOrders = ords.map((o:any) => ({ id:o.id, name:o.name, partnerName:o.partner_id?o.partner_id[1]:undefined, invoiceStatus:o.invoice_status }));
  const resolveComp = async (ref:string) => {
    let prods = await odoo.searchRead(session,"product.product",[["default_code","=ilike",ref.trim()]],["id","name","default_code"],1);
    if (!prods.length) prods = await odoo.searchRead(session,"product.product",[["default_code","ilike",ref.trim()]],["id","name","default_code"],1);
    if (!prods.length) return null;
    return { ref, productId:prods[0].id, name:prods[0].name };
  };
  const resolved = offre.produits.length ? (await Promise.all(offre.produits.map(resolveComp))).filter(Boolean) as {ref:string;productId:number;name:string}[] : [];
  let produits: ProduitCA[] = resolved.map(r => ({ ...r, qtyVendue:0, ca:0 }));
  let caTotal = 0;
  if (resolved.length > 0) {
    const compIds = resolved.map(r => r.productId);
    const compLines = await odoo.searchRead(session,"sale.order.line",[["order_id","in",orderIds],["product_id","in",compIds],["display_type","=",false],["is_downpayment","=",false]],["product_id","product_uom_qty","price_subtotal","state"],0);
    const activeComp = compLines.filter((l:any) => l.state !== "cancel");
    const prodMap: Record<number,{qty:number;ca:number}> = {};
    for (const l of activeComp) { const pid=l.product_id[0]; if (!prodMap[pid]) prodMap[pid]={qty:0,ca:0}; prodMap[pid].qty+=l.product_uom_qty||0; prodMap[pid].ca+=l.price_subtotal||0; }
    produits = resolved.map(r => ({ ...r, qtyVendue:prodMap[r.productId]?.qty||0, ca:prodMap[r.productId]?.ca||0 }));
    caTotal = produits.reduce((s,p) => s+p.ca, 0);
  }
  const orders = await odoo.searchRead(session,"sale.order",[["id","in",orderIds]],["id","user_id"],orderIds.length);
  const orderUserMap: Record<number,{userId:number;name:string}> = {};
  for (const o of orders) { if (o.user_id) orderUserMap[o.id]={userId:o.user_id[0],name:o.user_id[1]}; }
  const userMap: Record<number,{name:string;qty:number;ca:number}> = {};
  for (const l of activeOffreLines) { const user=orderUserMap[l.order_id[0]]; if (!user) continue; if (!userMap[user.userId]) userMap[user.userId]={name:user.name,qty:0,ca:0}; userMap[user.userId].qty+=l.product_uom_qty||0; }
  if (resolved.length > 0) {
    const compLines2 = await odoo.searchRead(session,"sale.order.line",[["order_id","in",orderIds],["product_id","in",resolved.map(r=>r.productId)],["display_type","=",false],["is_downpayment","=",false]],["order_id","price_subtotal","state"],0);
    for (const l of compLines2.filter((l:any) => l.state !== "cancel")) { const user=orderUserMap[l.order_id[0]]; if (!user) continue; if (!userMap[user.userId]) userMap[user.userId]={name:user.name,qty:0,ca:0}; userMap[user.userId].ca+=l.price_subtotal||0; }
  }
  const delegues: DelegueCA[] = Object.entries(userMap).map(([uid,v]) => ({userId:Number(uid),name:v.name,qtyVendue:v.qty,ca:v.ca})).sort((a,b)=>b.ca-a.ca);
  const splitValide={qty:0,ca:0}; const splitAvenir={qty:0,ca:0};
  if (filter === "all") {
    for (const l of activeOffreLines) { const inv=orderInvoiceMap[l.order_id[0]]; if (inv==="invoiced") splitValide.qty+=l.product_uom_qty||0; else splitAvenir.qty+=l.product_uom_qty||0; }
    if (resolved.length > 0) {
      const compIds=resolved.map(r=>r.productId);
      const splitLines = await odoo.searchRead(session,"sale.order.line",[["order_id","in",orderIds],["product_id","in",compIds],["display_type","=",false],["is_downpayment","=",false]],["order_id","price_subtotal","state"],0);
      for (const l of splitLines.filter((l:any) => l.state !== "cancel")) { const inv=orderInvoiceMap[l.order_id[0]]; if (inv==="invoiced") splitValide.ca+=l.price_subtotal||0; else splitAvenir.ca+=l.price_subtotal||0; }
    }
  }
  return { caTotal, qtyTotal, produits, delegues, debugOrders, split:{valide:splitValide,avenir:splitAvenir}, error:null };
}

// ── fetchCatchall (identique) ─────────────────────────────────────────────────
async function fetchCatchall(session: odoo.OdooSession, codeInterne: string, excludeOrderIds: number[], excludeOfferCodes: string[], filter: StateFilter, produitRefs: string[] = []): Promise<Omit<OffreAnalyse,"offre"|"loading">> {
  const orderDomain = filterDomain(filter);
  const noteOrders = await odoo.searchRead(session,"sale.order",[["x_note_interne","ilike",codeInterne.trim()],...orderDomain],["id","name","user_id","partner_id","invoice_status"],0);
  const excludeSet = new Set(excludeOrderIds);
  let orphans = noteOrders.filter((o:any) => !excludeSet.has(o.id));
  if (orphans.length > 0 && excludeOfferCodes.length > 0) {
    const orphanIds = orphans.map((o:any) => o.id as number);
    const offerProds = await odoo.searchRead(session,"product.product",[["default_code","in",excludeOfferCodes]],["id"],0);
    const offerProdIds = offerProds.map((p:any) => p.id as number);
    if (offerProdIds.length > 0) {
      const offerLines = await odoo.searchRead(session,"sale.order.line",[["order_id","in",orphanIds],["product_id","in",offerProdIds],["display_type","=",false]],["order_id"],0);
      const ordersWithOfferLines = new Set(offerLines.map((l:any) => l.order_id[0] as number));
      orphans = orphans.filter((o:any) => !ordersWithOfferLines.has(o.id));
    }
  }
  if (!orphans.length) return { caTotal:0, qtyTotal:0, produits:[], delegues:[], debugOrders:[], error:null };
  const orphanIds = orphans.map((o:any) => o.id as number);
  const debugOrders = orphans.map((o:any) => ({id:o.id,name:`${o.name} (note)`,partnerName:o.partner_id?o.partner_id[1]:undefined,invoiceStatus:o.invoice_status}));
  let filteredProdIds: Set<number>|null = null;
  const prodRefMap: Record<number,string> = {};
  if (produitRefs.length > 0) {
    const configProds = await odoo.searchRead(session,"product.product",[["default_code","in",produitRefs]],["id","default_code"],0);
    filteredProdIds = new Set(configProds.map((p:any) => p.id as number));
    for (const p of configProds) prodRefMap[p.id] = p.default_code;
  }
  const lines = await odoo.searchRead(session,"sale.order.line",[["order_id","in",orphanIds],["display_type","=",false],["is_downpayment","=",false]],["order_id","product_id","product_uom_qty","price_subtotal","state"],0);
  const activeLines = lines.filter((l:any) => { if (l.state==="cancel") return false; if (l.price_subtotal<=0) return false; if (filteredProdIds&&l.product_id&&!filteredProdIds.has(l.product_id[0])) return false; return true; });
  const caTotal = activeLines.reduce((s:number,l:any) => s+(l.price_subtotal||0),0);
  const prodMap: Record<number,{name:string;ref:string;qty:number;ca:number}> = {};
  for (const l of activeLines) { if (!l.product_id) continue; const pid=l.product_id[0]; const ref=prodRefMap[pid]||l.product_id[1]; if (!prodMap[pid]) prodMap[pid]={name:l.product_id[1],ref,qty:0,ca:0}; prodMap[pid].qty+=l.product_uom_qty||0; prodMap[pid].ca+=l.price_subtotal||0; }
  const produits: ProduitCA[] = Object.entries(prodMap).map(([pid,v]) => ({productId:Number(pid),ref:v.ref,name:v.name,qtyVendue:v.qty,ca:v.ca})).sort((a,b)=>b.ca-a.ca);
  const userMap: Record<number,{name:string;qty:number;ca:number}> = {};
  for (const o of orphans) { if (!o.user_id) continue; const uid=o.user_id[0]; if (!userMap[uid]) userMap[uid]={name:o.user_id[1],qty:0,ca:0}; }
  for (const l of activeLines) { const o=orphans.find((x:any)=>x.id===l.order_id[0]); if (!o?.user_id) continue; const uid=o.user_id[0]; if (userMap[uid]) { userMap[uid].qty+=l.product_uom_qty||0; userMap[uid].ca+=l.price_subtotal||0; } }
  const delegues: DelegueCA[] = Object.entries(userMap).map(([uid,v]) => ({userId:Number(uid),name:v.name,qtyVendue:v.qty,ca:v.ca})).sort((a,b)=>b.ca-a.ca);
  const splitValide={qty:0,ca:0}; const splitAvenir={qty:0,ca:0};
  if (filter === "all") {
    for (const o of orphans) { const inv=o.invoice_status??""; const orderLines=activeLines.filter((l:any)=>l.order_id[0]===o.id); const orderCA=orderLines.reduce((s:number,l:any)=>s+(l.price_subtotal||0),0); if (inv==="invoiced") {splitValide.qty+=1;splitValide.ca+=orderCA;} else {splitAvenir.qty+=1;splitAvenir.ca+=orderCA;} }
  }
  return { caTotal, qtyTotal:orphanIds.length, produits, delegues, debugOrders, split:{valide:splitValide,avenir:splitAvenir}, error:null };
}

function fmt(n: number) { return new Intl.NumberFormat("fr-FR",{style:"currency",currency:"EUR",maximumFractionDigits:0}).format(n); }
function initials(name: string) { return name.split(" ").map((n:string)=>n[0]).slice(0,2).join("").toUpperCase(); }

// ── Export Excel (identique) ──────────────────────────────────────────────────
interface CatchallResult { codeInterne:string; loading:boolean; data:Omit<OffreAnalyse,"offre"|"loading">|null; }
async function exportToExcel(results: OffreAnalyse[], catchalls: CatchallResult[], onToast: Props["onToast"], setExporting: (v:boolean)=>void) {
  setExporting(true);
  try {
    const payload = { results:results.filter(r=>!r.loading&&!r.error).map(r=>({offre:r.offre,caTotal:r.caTotal,qtyTotal:r.qtyTotal,produits:r.produits,delegues:r.delegues,debugOrders:r.debugOrders??[]})), catchalls:catchalls.filter(c=>!c.loading&&c.data).map(c=>({codeInterne:c.codeInterne,data:c.data})) };
    const res = await fetch("/api/export-excel",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
    if (!res.ok) throw new Error(`Erreur ${res.status}`);
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download=`analyse_offres_${new Date().toISOString().slice(0,10)}.xlsx`; a.click(); URL.revokeObjectURL(url);
    onToast("Export téléchargé","success");
  } catch(e:any) { onToast("Erreur export : "+e.message,"error"); }
  finally { setExporting(false); }
}

// ─────────────────────────────────────────────────────────────────────────────
// GESTION DES OFFRES (modal/panneau)
// ─────────────────────────────────────────────────────────────────────────────
function OffresPanel({ onClose, onToast }: { onClose: ()=>void; onToast: Props["onToast"]; }) {
  const [offres, setOffres] = useState<Offre[]>([]);
  const [editId, setEditId] = useState<string|null>(null);
  const [showForm, setShowForm] = useState(false);
  const [formCode, setFormCode] = useState("");
  const [formLabel, setFormLabel] = useState("");
  const [formProduits, setFormProduits] = useState("");
  const [formCodeInterne, setFormCodeInterne] = useState("");

  useEffect(() => {
    migrateOffresFromLS().then(n => {
      if (n > 0) onToast(`${n} offre(s) importée(s)`, "success");
    });
    dbLoadOffres().then(setOffres).catch(e => onToast("Erreur chargement offres : " + e.message, "error"));
  }, []);

  const openNew = () => { setEditId(null); setFormCode(""); setFormLabel(""); setFormProduits(""); setFormCodeInterne(""); setShowForm(true); };
  const openEdit = (o: Offre) => { setEditId(o.id); setFormCode(o.code); setFormLabel(o.label); setFormProduits(o.produits.join("\n")); setFormCodeInterne(o.codeInterne||""); setShowForm(true); };

  const save = () => {
    const code=formCode.trim(); const label=formLabel.trim();
    const produits=formProduits.split(/[\n\r,;]+/).map(r=>r.trim()).filter(Boolean);
    const codeInterne=formCodeInterne.trim()||undefined;
    if (!code) { onToast("Code offre requis","error"); return; }
    if (!produits.length) { onToast("Au moins un produit requis","error"); return; }
    if (editId) {
      const offre = { id: editId, code, label, produits, codeInterne };
      dbUpsertOffre(offre).then(() => {
        setOffres(prev => prev.map(o => o.id === editId ? offre : o));
        setShowForm(false); onToast("Offre mise à jour", "success");
      }).catch(e => onToast("Erreur : " + e.message, "error"));
    } else {
      if (offres.some(o=>o.code.toLowerCase()===code.toLowerCase())) { onToast("Ce code offre existe déjà","error"); return; }
      const offre = { id: genId(), code, label, produits, codeInterne };
      dbUpsertOffre(offre).then(() => {
        setOffres(prev => [...prev, offre]);
        setShowForm(false); onToast("Offre créée", "success");
      }).catch(e => onToast("Erreur : " + e.message, "error"));
    }
  };

  const deleteOffre = (id: string) => {
    if (!confirm("Supprimer cette offre ?")) return;
    dbDeleteOffre(id).then(() => {
      setOffres(prev => prev.filter(o => o.id !== id)); onToast("Offre supprimée", "info");
    }).catch(e => onToast("Erreur suppression : " + e.message, "error"));
  };

  const inputStyle = { width:"100%", padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:13, fontFamily:"inherit", background:C.white, color:C.text, outline:"none" };
  const labelStyle = { fontSize:11, fontWeight:700 as const, color:C.textMuted, textTransform:"uppercase" as const, letterSpacing:"0.07em", display:"block" as const, marginBottom:5 };

  return (
    <div style={{ position:"fixed", inset:0, zIndex:200, display:"flex" }}>
      <div style={{ flex:1, background:"rgba(0,0,0,0.4)" }} onClick={onClose} />
      <div style={{ width:480, background:C.white, display:"flex", flexDirection:"column", boxShadow:"-8px 0 40px rgba(0,0,0,0.15)", overflow:"hidden" }}>
        {/* Header */}
        <div style={{ padding:"20px 24px 16px", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700, color:C.text }}>{showForm ? (editId?"Modifier l'offre":"Nouvelle offre") : "Gestion des offres"}</div>
            {!showForm && <div style={{ fontSize:12, color:C.textMuted, marginTop:2 }}>{offres.length} offre{offres.length!==1?"s":""} configurée{offres.length!==1?"s":""}</div>}
          </div>
          <button onClick={showForm ? ()=>setShowForm(false) : onClose} style={{ width:32, height:32, display:"flex", alignItems:"center", justifyContent:"center", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, cursor:"pointer" }}>
            {showForm
              ? <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.text} strokeWidth="2"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
              : <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.text} strokeWidth="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            }
          </button>
        </div>

        {/* Content */}
        <div style={{ flex:1, overflowY:"auto", padding:"20px 24px" }}>
          {showForm ? (
            <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
              {[
                { label:"Code offre *", value:formCode, set:setFormCode, placeholder:"Ex: 7131482" },
                { label:"Libellé (optionnel)", value:formLabel, set:setFormLabel, placeholder:"Ex: Offre été 2024" },
                { label:"Code interne (x_note_interne)", value:formCodeInterne, set:setFormCodeInterne, placeholder:"Ex: CURE26" },
              ].map(({label,value,set,placeholder}) => (
                <div key={label}>
                  <label style={labelStyle}>{label}</label>
                  <input value={value} onChange={e=>set(e.target.value)} placeholder={placeholder} style={inputStyle} />
                </div>
              ))}
              <div>
                <label style={labelStyle}>Références produits * — une par ligne</label>
                <textarea
                  value={formProduits}
                  onChange={e=>setFormProduits(e.target.value)}
                  onPaste={e=>{const p=e.clipboardData.getData("text");const refs=extractRefs(p);if(refs.length>=2){e.preventDefault();setFormProduits(refs.join("\n"));}}}
                  placeholder={"1010214\n1010302\n1010305"}
                  rows={6}
                  style={{...inputStyle, fontFamily:"'SF Mono', 'Fira Code', monospace", fontSize:12, resize:"vertical"}}
                />
                <div style={{ display:"flex", justifyContent:"space-between", marginTop:5 }}>
                  <span style={{ fontSize:11, color:C.textMuted }}>{formProduits.split(/[\n\r,;]+/).filter(r=>r.trim()).length} référence(s)</span>
                  <button type="button" onClick={()=>setFormProduits(extractRefs(formProduits).join("\n"))}
                    style={{ fontSize:11, padding:"3px 10px", background:"transparent", border:`1px solid ${C.border}`, borderRadius:6, color:C.textMuted, cursor:"pointer", fontFamily:"inherit" }}>
                    🧹 Nettoyer
                  </button>
                </div>
              </div>
              <button onClick={save} style={{ marginTop:4, padding:"11px 0", background:C.blue, color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
                {editId ? "Enregistrer les modifications" : "Créer l'offre"}
              </button>
            </div>
          ) : (
            <>
              <button onClick={openNew} style={{ width:"100%", padding:"10px 0", background:C.blue, color:"#fff", border:"none", borderRadius:10, fontSize:13, fontWeight:700, cursor:"pointer", fontFamily:"inherit", marginBottom:16, display:"flex", alignItems:"center", justifyContent:"center", gap:7 }}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                Nouvelle offre
              </button>
              {offres.length === 0 ? (
                <div style={{ textAlign:"center", padding:"32px 0", color:C.textMuted }}>
                  <div style={{ fontSize:32, marginBottom:10 }}>🗂️</div>
                  <div style={{ fontSize:13, fontWeight:600 }}>Aucune offre configurée</div>
                  <div style={{ fontSize:12, marginTop:4 }}>Créez vos offres et associez leurs références produits</div>
                </div>
              ) : (
                <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                  {offres.map(o => (
                    <div key={o.id} style={{ border:`1px solid ${C.border}`, borderRadius:10, padding:"12px 14px", background:C.white }}>
                      <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:10 }}>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ display:"flex", alignItems:"center", gap:7, flexWrap:"wrap" }}>
                            <span style={{ fontSize:13, fontWeight:700, color:C.text, fontFamily:"'SF Mono','Fira Code',monospace" }}>{o.code}</span>
                            {o.label && <span style={{ fontSize:12, color:C.textMuted }}>{o.label}</span>}
                            {o.codeInterne && <span style={{ fontSize:11, background:C.purpleSoft, color:C.purple, borderRadius:5, padding:"2px 7px", fontWeight:600 }}>{o.codeInterne}</span>}
                          </div>
                          <div style={{ marginTop:7, display:"flex", flexWrap:"wrap", gap:4 }}>
                            {o.produits.map(p => <span key={p} style={{ fontSize:11, fontFamily:"'SF Mono','Fira Code',monospace", background:C.blueSoft, color:C.blue, borderRadius:5, padding:"2px 7px" }}>{p}</span>)}
                          </div>
                          <div style={{ fontSize:11, color:C.textMuted, marginTop:5 }}>{o.produits.length} produit(s)</div>
                        </div>
                        <div style={{ display:"flex", gap:5, flexShrink:0 }}>
                          <button onClick={()=>openEdit(o)} style={{ padding:"5px 9px", background:C.bg, border:`1px solid ${C.border}`, borderRadius:7, cursor:"pointer", fontSize:11, color:C.textSec }}>✏️</button>
                          <button onClick={()=>deleteOffre(o.id)} style={{ padding:"5px 9px", background:C.redSoft, border:`1px solid ${C.red}22`, borderRadius:7, cursor:"pointer", fontSize:11, color:C.red }}>🗑</button>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// ONGLET ANALYSE
// ─────────────────────────────────────────────────────────────────────────────
function AnalyseTab({ session, onToast, filter, sharedCodes, onCodesChange }: { session:odoo.OdooSession; onToast:Props["onToast"]; filter:StateFilter; sharedCodes:string[]; onCodesChange:(c:string[])=>void; }) {
  const [configOffres, setConfigOffres] = useState<Offre[]>([]);
  const [results, setResults] = useState<OffreAnalyse[]>([]);
  const [catchalls, setCatchalls] = useState<CatchallResult[]>([]);
  const [pendingCodes, setPendingCodes] = useState<string[]>([]);
  const [globalLoading, setGlobalLoading] = useState(false);
  const [expandedId, setExpandedId] = useState<string|null>(null);
  const [detailMode, setDetailMode] = useState<Record<string,"produits"|"delegues"|"debug">>({});
  const [exporting, setExporting] = useState(false);
  const [dropdownOpen, setDropdownOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!dropdownOpen) return;
    const handler = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setDropdownOpen(false);
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [dropdownOpen]);

  useEffect(() => { dbLoadOffres().then(setConfigOffres).catch(console.error); }, []);
  useEffect(() => {
    if (sharedCodes.length > 0 && configOffres.length > 0) {
      const valid = sharedCodes.filter(c => configOffres.some(o => o.code.toLowerCase() === c.toLowerCase()));
      if (valid.length > 0) setPendingCodes(valid);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [configOffres]);

  const findOffre = useCallback((code:string): Offre|null => configOffres.find(o=>o.code.toLowerCase()===code.trim().toLowerCase())||null, [configOffres]);
  const alreadyLoadedCodes = results.map(r => r.offre.code.toLowerCase());
  const availableOffres = configOffres.filter(o => !alreadyLoadedCodes.includes(o.code.toLowerCase()) && !pendingCodes.map(c=>c.toLowerCase()).includes(o.code.toLowerCase()));

  const addCode = (code: string) => {
    if (!code) return;
    const offre = findOffre(code);
    if (!offre) { onToast(`Offre "${code}" non configurée`, "error"); return; }
    if (!pendingCodes.includes(code)) setPendingCodes(p => [...p, code]);
  };

  const analyseAll = async () => {
    if (!pendingCodes.length) return;
    const codesToLoad = [...pendingCodes]; setPendingCodes([]); setGlobalLoading(true);
    const placeholders: OffreAnalyse[] = codesToLoad.map(code => ({ offre:findOffre(code)!, loading:true, error:null, caTotal:0, qtyTotal:0, produits:[], delegues:[] }));
    setResults(prev => [...prev, ...placeholders]);
    const localFinished: OffreAnalyse[] = [];
    await Promise.all(codesToLoad.map(async code => {
      const offre = findOffre(code)!;
      try {
        const data = await fetchCAForOffre(session, offre, filter);
        const finished: OffreAnalyse = { offre, loading:false, ...data };
        localFinished.push(finished);
        setResults(prev => prev.map(r => r.offre.code===code ? finished : r));
      } catch(e:any) {
        const failed: OffreAnalyse = { offre, loading:false, error:e.message||"Erreur Odoo", caTotal:0, qtyTotal:0, produits:[], delegues:[] };
        localFinished.push(failed);
        setResults(prev => prev.map(r => r.offre.code===code ? failed : r));
      }
    }));
    const prevLoaded = results.filter(r => !r.loading && !r.error);
    await runCatchalls([...prevLoaded, ...localFinished.filter(r=>!r.error)]);
    setGlobalLoading(false);
    const allLoaded = [...prevLoaded.map(r=>r.offre.code), ...localFinished.filter(r=>!r.error).map(r=>r.offre.code)];
    onCodesChange(allLoaded);
    onToast(`${localFinished.length} offre(s) analysée(s)`,"success");
  };

  const runCatchalls = async (currentResults: OffreAnalyse[]) => {
    const codesInternes = Array.from(new Set(currentResults.filter(r=>r.offre.codeInterne?.trim()).map(r=>r.offre.codeInterne!.trim())));
    if (!codesInternes.length) { setCatchalls([]); return; }
    const allOrderIds = currentResults.flatMap(r=>(r.debugOrders||[]).map(o=>o.id));
    const allOfferCodes = currentResults.map(r=>r.offre.code);
    setCatchalls(codesInternes.map(ci => ({ codeInterne:ci, loading:true, data:null })));
    await Promise.all(codesInternes.map(async ci => {
      try {
        const parentOffre = currentResults.find(r=>r.offre.codeInterne?.trim()===ci);
        const data = await fetchCatchall(session, ci, allOrderIds, allOfferCodes, filter, parentOffre?.offre.produits??[]);
        setCatchalls(prev => prev.map(c => c.codeInterne===ci ? {...c,loading:false,data} : c));
      } catch { setCatchalls(prev => prev.map(c => c.codeInterne===ci ? {...c,loading:false,data:null} : c)); }
    }));
  };

  const removeResult = (code: string) => {
    setResults(prev => {
      const next = prev.filter(r => r.offre.code !== code);
      const remaining = Array.from(new Set(next.filter(r=>r.offre.codeInterne?.trim()).map(r=>r.offre.codeInterne!.trim())));
      setCatchalls(prev => prev.filter(c => remaining.includes(c.codeInterne)));
      return next;
    });
  };

  const clearAll = () => { setResults([]); setCatchalls([]); setPendingCodes([]); setExpandedId(null); onCodesChange([]); };

  const refreshAll = async () => {
    setGlobalLoading(true);
    setResults(prev => prev.map(r => ({ ...r, loading:true, error:null })));
    const refreshed: OffreAnalyse[] = [];
    await Promise.all(results.map(async r => {
      try {
        const data = await fetchCAForOffre(session, r.offre, filter);
        const finished: OffreAnalyse = { ...r, ...data, loading:false };
        refreshed.push(finished);
        setResults(prev => prev.map(x => x.offre.code===r.offre.code ? finished : x));
      } catch(e:any) {
        const failed: OffreAnalyse = { ...r, loading:false, error:e.message||"Erreur" };
        refreshed.push(failed);
        setResults(prev => prev.map(x => x.offre.code===r.offre.code ? failed : x));
      }
    }));
    await runCatchalls(refreshed);
    setGlobalLoading(false);
    onToast("Données actualisées","success");
  };

  // Totaux
  const totalCA = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+r.caTotal,0) + catchalls.filter(c=>!c.loading&&c.data).reduce((s,c)=>s+(c.data?.caTotal??0),0);
  const totalQty = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+r.qtyTotal,0) + catchalls.filter(c=>!c.loading&&c.data).reduce((s,c)=>s+(c.data?.qtyTotal??0),0);
  const splitValideCA = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+(r.split?.valide.ca??0),0) + catchalls.filter(c=>!c.loading&&c.data).reduce((s,c)=>s+(c.data?.split?.valide.ca??0),0);
  const splitAvenirCA = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+(r.split?.avenir.ca??0),0) + catchalls.filter(c=>!c.loading&&c.data).reduce((s,c)=>s+(c.data?.split?.avenir.ca??0),0);
  const splitValideQty = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+(r.split?.valide.qty??0),0);
  const splitAvenirQty = results.filter(r=>!r.loading&&!r.error).reduce((s,r)=>s+(r.split?.avenir.qty??0),0);
  const hasSplit = filter === "all" && (splitValideQty > 0 || splitAvenirQty > 0);
  const hasResults = results.some(r => !r.loading && !r.error);

  const btnBase = { border:"none", borderRadius:8, cursor:"pointer", fontFamily:"inherit", fontWeight:600 as const, display:"flex", alignItems:"center", gap:6, fontSize:12 };

  return (
    <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>
      {/* Barre de sélection */}
      <div style={{ padding:"16px 24px", background:C.white, borderBottom:`1px solid ${C.border}`, flexShrink:0 }}>
        <div style={{ display:"flex", gap:10, alignItems:"center", flexWrap:"wrap" }}>
          {/* Dropdown multi-sélection */}
          <div ref={dropdownRef} style={{ flex:"1 1 240px", position:"relative" }}>
            <button
              onClick={() => setDropdownOpen(o => !o)}
              disabled={availableOffres.length === 0}
              style={{ width:"100%", padding:"9px 36px 9px 12px", border:`1.5px solid ${dropdownOpen?C.blue:C.border}`, borderRadius:9, fontSize:13, fontFamily:"inherit", background:C.white, color:availableOffres.length?C.text:C.textMuted, cursor:availableOffres.length?"pointer":"default", textAlign:"left", outline:"none", boxShadow:dropdownOpen?`0 0 0 3px ${C.blue}18`:"none" }}
            >
              {availableOffres.length === 0 ? "— Toutes les offres ajoutées —" : pendingCodes.length > 0 ? `${pendingCodes.length} offre(s) sélectionnée(s)` : "Sélectionner des offres…"}
            </button>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.textMuted} strokeWidth="2" style={{ position:"absolute", right:11, top:"50%", transform:`translateY(-50%) rotate(${dropdownOpen?180:0}deg)`, pointerEvents:"none", transition:"transform 0.15s" }}>
              <polyline points="6 9 12 15 18 9"/>
            </svg>
            {dropdownOpen && availableOffres.length > 0 && (
              <div style={{ position:"absolute", top:"calc(100% + 6px)", left:0, right:0, background:C.white, border:`1.5px solid ${C.blue}44`, borderRadius:10, boxShadow:"0 8px 32px rgba(0,0,0,0.13)", zIndex:50, overflow:"hidden" }}>
                {/* Tout sélectionner */}
                <div
                  onClick={() => {
                    const allCodes = availableOffres.map(o => o.code);
                    const allSelected = allCodes.every(c => pendingCodes.includes(c));
                    if (allSelected) setPendingCodes(p => p.filter(c => !allCodes.includes(c)));
                    else setPendingCodes(p => [...p, ...allCodes.filter(c => !p.includes(c))]);
                  }}
                  style={{ padding:"10px 14px", display:"flex", alignItems:"center", gap:10, cursor:"pointer", borderBottom:`1px solid ${C.border}`, background:C.bg }}
                >
                  <div style={{ width:16, height:16, border:`2px solid ${availableOffres.every(o=>pendingCodes.includes(o.code))?C.blue:C.border}`, borderRadius:4, background:availableOffres.every(o=>pendingCodes.includes(o.code))?C.blue:"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                    {availableOffres.every(o=>pendingCodes.includes(o.code)) && <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>}
                  </div>
                  <span style={{ fontSize:12, fontWeight:700, color:C.textSec }}>Tout sélectionner</span>
                </div>
                {/* Liste des offres */}
                <div style={{ maxHeight:260, overflowY:"auto" }}>
                  {availableOffres.map(o => {
                    const checked = pendingCodes.includes(o.code);
                    return (
                      <div key={o.id}
                        onClick={() => setPendingCodes(p => checked ? p.filter(c=>c!==o.code) : [...p, o.code])}
                        style={{ padding:"9px 14px", display:"flex", alignItems:"center", gap:10, cursor:"pointer", background:checked?C.blueSoft:"transparent", borderBottom:`1px solid ${C.border}` }}
                        onMouseEnter={e => { if(!checked)(e.currentTarget as HTMLDivElement).style.background=C.bg; }}
                        onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background=checked?C.blueSoft:"transparent"; }}
                      >
                        <div style={{ width:16, height:16, border:`2px solid ${checked?C.blue:C.border}`, borderRadius:4, background:checked?C.blue:"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                          {checked && <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>}
                        </div>
                        <div style={{ minWidth:0 }}>
                          <span style={{ fontSize:12, fontWeight:700, color:C.text, fontFamily:"'SF Mono','Fira Code',monospace" }}>{o.code}</span>
                          {o.label && <span style={{ fontSize:12, color:C.textMuted }}> — {o.label}</span>}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>

          {/* Chips en attente */}
          {pendingCodes.length > 0 && (
            <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
              {pendingCodes.map(code => {
                const o = findOffre(code);
                return (
                  <span key={code} style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"5px 10px", background:C.blueSoft, border:`1px solid ${C.blue}33`, borderRadius:7, fontSize:12, fontWeight:600, color:C.blue }}>
                    {code}{o?.label ? ` · ${o.label}` : ""}
                    <button onClick={()=>setPendingCodes(p=>p.filter(c=>c!==code))} style={{ background:"none", border:"none", cursor:"pointer", color:C.blue, padding:0, lineHeight:1, fontSize:14 }}>×</button>
                  </span>
                );
              })}
            </div>
          )}

          {/* Bouton analyser */}
          <button onClick={analyseAll} disabled={!pendingCodes.length||globalLoading}
            style={{ ...btnBase, padding:"9px 20px", background:pendingCodes.length&&!globalLoading?C.blue:C.border, color:pendingCodes.length&&!globalLoading?"#fff":C.textMuted, cursor:pendingCodes.length&&!globalLoading?"pointer":"default", fontSize:13, fontWeight:700, flexShrink:0 }}>
            {globalLoading
              ? <><span style={{ width:13,height:13,borderRadius:"50%",border:"2px solid rgba(255,255,255,0.3)",borderTopColor:"#fff",display:"inline-block",animation:"spin 0.7s linear infinite" }}/>Calcul…</>
              : <><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polygon points="5 3 19 12 5 21 5 3"/></svg>Analyser</>}
          </button>
        </div>
      </div>

      {/* Contenu scrollable */}
      <div style={{ flex:1, overflowY:"auto", padding:"20px 24px" }}>

        {/* Stats */}
        {hasResults && (
          <div style={{ marginBottom:20 }}>
            <div style={{ display:"grid", gridTemplateColumns:`repeat(${hasSplit?4:2}, 1fr)`, gap:12, marginBottom:14 }}>
              {[
                { label:"CA Total", value:fmt(totalCA), color:C.green, soft:C.greenSoft, icon:<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg> },
                { label:"Quantité vendue", value:Math.round(totalQty).toString(), color:C.amber, soft:C.amberSoft, icon:<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/></svg> },
                ...(hasSplit ? [
                  { label:"Validé", value:fmt(splitValideCA), color:C.green, soft:"#f0fdf4", icon:<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="20 6 9 17 4 12"/></svg> },
                  { label:"À venir", value:fmt(splitAvenirCA), color:C.amber, soft:C.amberSoft, icon:<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> },
                ] : []),
              ].map(stat => (
                <div key={stat.label} style={{ background:stat.soft, border:`1px solid ${stat.color}22`, borderRadius:12, padding:"14px 16px", display:"flex", alignItems:"center", gap:12 }}>
                  <div style={{ width:38,height:38,borderRadius:10,background:`${stat.color}18`,display:"flex",alignItems:"center",justifyContent:"center",color:stat.color,flexShrink:0 }}>{stat.icon}</div>
                  <div>
                    <div style={{ fontSize:11,fontWeight:600,color:stat.color,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:2 }}>{stat.label}</div>
                    <div style={{ fontSize:22,fontWeight:800,color:stat.color,lineHeight:1 }}>{stat.value}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* Actions */}
            <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
              <button onClick={refreshAll} disabled={globalLoading} style={{ ...btnBase, padding:"7px 14px", background:C.white, border:`1px solid ${C.border}`, color:C.textSec }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>
                Actualiser
              </button>
              <button onClick={clearAll} style={{ ...btnBase, padding:"7px 14px", background:C.white, border:`1px solid ${C.red}44`, color:C.red }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
                Vider
              </button>
              <button onClick={()=>exportToExcel(results,catchalls,onToast,setExporting)} disabled={exporting} style={{ ...btnBase, padding:"7px 16px", background:exporting?C.border:C.green, border:"none", color:"#fff" }}>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                {exporting?"Export…":"Excel"}
              </button>
            </div>
          </div>
        )}

        {/* Résultats */}
        {results.length === 0 ? (
          configOffres.length === 0 ? (
            <div style={{ textAlign:"center", padding:"60px 24px", color:C.textMuted }}>
              <div style={{ width:56,height:56,borderRadius:16,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px" }}>
                <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke={C.textMuted} strokeWidth="1.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
              </div>
              <div style={{ fontSize:14, fontWeight:600, color:C.textSec, marginBottom:4 }}>Aucune offre configurée</div>
              <div style={{ fontSize:12 }}>Ouvrez la gestion des offres pour commencer</div>
            </div>
          ) : (
            <div style={{ textAlign:"center", padding:"60px 24px", color:C.textMuted }}>
              <div style={{ width:56,height:56,borderRadius:16,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px" }}>
                <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke={C.textMuted} strokeWidth="1.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
              </div>
              <div style={{ fontSize:14, fontWeight:600, color:C.textSec, marginBottom:4 }}>Sélectionnez des offres à analyser</div>
              <div style={{ fontSize:12 }}>Utilisez le menu déroulant ci-dessus puis cliquez sur Analyser</div>
            </div>
          )
        ) : (
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {results.map(r => {
              const isExpanded = expandedId === r.offre.id;
              const mode = detailMode[r.offre.id];
              return (
                <div key={r.offre.id} style={{ background:C.white, border:`1px solid ${r.error?C.red+"44":C.border}`, borderRadius:12, overflow:"hidden", boxShadow:C.shadow }}>
                  {/* Header ligne */}
                  <div style={{ padding:"14px 18px", display:"flex", alignItems:"center", gap:16 }}>
                    {/* Infos offre */}
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ display:"flex", alignItems:"center", gap:8, flexWrap:"wrap" }}>
                        <span style={{ fontSize:14, fontWeight:700, color:C.text, fontFamily:"'SF Mono','Fira Code',monospace" }}>{r.offre.code}</span>
                        {r.offre.label && <span style={{ fontSize:12, color:C.textMuted }}>{r.offre.label}</span>}
                        {r.offre.codeInterne && <span style={{ fontSize:11, background:C.purpleSoft, color:C.purple, borderRadius:5, padding:"2px 7px", fontWeight:600 }}>{r.offre.codeInterne}</span>}
                      </div>
                      {r.loading && <div style={{ fontSize:12, color:C.textMuted, marginTop:3, display:"flex", alignItems:"center", gap:5 }}><span style={{ width:11,height:11,borderRadius:"50%",border:`2px solid ${C.blue}`,borderTopColor:"transparent",display:"inline-block",animation:"spin 0.8s linear infinite" }}/>Calcul en cours…</div>}
                      {!r.loading && r.error && <div style={{ fontSize:12, color:C.red, marginTop:3 }}>{r.error}</div>}
                      {!r.loading && !r.error && <div style={{ fontSize:11, color:C.textMuted, marginTop:3 }}>{r.offre.produits.length} référence(s)</div>}
                    </div>

                    {/* Métriques */}
                    {!r.loading && !r.error && (
                      <div style={{ display:"flex", gap:24, alignItems:"center", flexShrink:0 }}>
                        <div style={{ textAlign:"right" }}>
                          <div style={{ fontSize:10, fontWeight:600, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.05em" }}>CA</div>
                          <div style={{ fontSize:18, fontWeight:800, color:C.green }}>{fmt(r.caTotal)}</div>
                        </div>
                        <div style={{ textAlign:"right" }}>
                          <div style={{ fontSize:10, fontWeight:600, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.05em" }}>Qté</div>
                          <div style={{ fontSize:18, fontWeight:800, color:C.amber }}>{Math.round(r.qtyTotal)}</div>
                        </div>
                        {/* Boutons détail */}
                        <div style={{ display:"flex", gap:4 }}>
                          {([["produits",`Produits (${r.produits.length})`,C.blue],["delegues",`Délégués (${r.delegues.length})`,C.purple],["debug",`Cdes (${r.debugOrders?.length??0})`,C.slate]] as [string,string,string][]).map(([m,lbl,col]) => (
                            <button key={m} onClick={()=>{ if(isExpanded&&mode===m){setExpandedId(null);}else{setExpandedId(r.offre.id);setDetailMode(dm=>({...dm,[r.offre.id]:m as any}));} }}
                              style={{ padding:"5px 10px", background:isExpanded&&mode===m?col+"18":"transparent", border:`1px solid ${isExpanded&&mode===m?col:C.border}`, borderRadius:7, cursor:"pointer", fontSize:11, fontWeight:600, color:isExpanded&&mode===m?col:C.textMuted, fontFamily:"inherit" }}>
                              {lbl}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}

                    <button onClick={()=>removeResult(r.offre.code)} style={{ width:28,height:28,display:"flex",alignItems:"center",justifyContent:"center",background:"transparent",border:"none",cursor:"pointer",color:C.textMuted,flexShrink:0,borderRadius:6 }}>
                      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                    </button>
                  </div>

                  {/* Panneaux détail */}
                  {!r.loading && !r.error && isExpanded && mode === "produits" && (
                    <div style={{ borderTop:`1px solid ${C.border}` }}>
                      <table style={{ width:"100%", borderCollapse:"collapse" as const, fontSize:12 }}>
                        <thead>
                          <tr style={{ background:C.bg }}>
                            <th style={{ padding:"8px 18px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11 }}>Référence</th>
                            <th style={{ padding:"8px 18px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11 }}>Produit</th>
                            <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>Qté</th>
                            <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>CA</th>
                            <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>%</th>
                          </tr>
                        </thead>
                        <tbody>
                          {r.produits.map((p,i) => (
                            <tr key={p.productId} style={{ borderTop:`1px solid ${C.border}`, background:i%2===0?C.white:C.bg+"80" }}>
                              <td style={{ padding:"9px 18px", fontFamily:"'SF Mono','Fira Code',monospace", fontWeight:600, color:C.blue }}>{p.ref}</td>
                              <td style={{ padding:"9px 18px", color:C.textSec, maxWidth:240, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{p.name}</td>
                              <td style={{ padding:"9px 18px", textAlign:"right", color:C.textSec }}>{Math.round(p.qtyVendue)}</td>
                              <td style={{ padding:"9px 18px", textAlign:"right", fontWeight:700, color:C.green }}>{fmt(p.ca)}</td>
                              <td style={{ padding:"9px 18px", textAlign:"right", color:C.textMuted }}>{Math.round((p.ca/(r.caTotal||1))*100)}%</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}

                  {!r.loading && !r.error && isExpanded && mode === "delegues" && (
                    <div style={{ borderTop:`1px solid ${C.border}` }}>
                      {r.delegues.length===0 ? <div style={{ padding:"16px 18px", color:C.textMuted, fontSize:12 }}>Aucun délégué trouvé</div> : (
                        <table style={{ width:"100%", borderCollapse:"collapse" as const, fontSize:12 }}>
                          <thead>
                            <tr style={{ background:C.bg }}>
                              <th style={{ padding:"8px 18px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11 }}>Délégué</th>
                              <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>Qté</th>
                              <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>CA</th>
                              <th style={{ padding:"8px 18px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}>Part</th>
                            </tr>
                          </thead>
                          <tbody>
                            {r.delegues.map((d,i) => (
                              <tr key={d.userId} style={{ borderTop:`1px solid ${C.border}`, background:i%2===0?C.white:C.bg+"80" }}>
                                <td style={{ padding:"9px 18px" }}>
                                  <div style={{ display:"flex", alignItems:"center", gap:9 }}>
                                    <div style={{ width:28,height:28,borderRadius:14,background:C.purpleSoft,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>
                                      <span style={{ fontSize:10,fontWeight:800,color:C.purple }}>{initials(d.name)}</span>
                                    </div>
                                    <span style={{ fontWeight:600, color:C.textSec }}>{d.name}</span>
                                  </div>
                                </td>
                                <td style={{ padding:"9px 18px", textAlign:"right", color:C.textSec }}>{Math.round(d.qtyVendue)}</td>
                                <td style={{ padding:"9px 18px", textAlign:"right", fontWeight:700, color:C.purple }}>{fmt(d.ca)}</td>
                                <td style={{ padding:"9px 18px", textAlign:"right" }}>
                                  <div style={{ display:"flex", alignItems:"center", gap:7, justifyContent:"flex-end" }}>
                                    <div style={{ width:60,height:5,background:C.border,borderRadius:3,overflow:"hidden" }}>
                                      <div style={{ height:"100%", width:`${Math.round((d.ca/(r.caTotal||1))*100)}%`, background:C.purple, borderRadius:3 }}/>
                                    </div>
                                    <span style={{ fontSize:11,color:C.textMuted,width:30,textAlign:"right" }}>{Math.round((d.ca/(r.caTotal||1))*100)}%</span>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </div>
                  )}

                  {!r.loading && !r.error && isExpanded && mode === "debug" && (
                    <div style={{ borderTop:`1px solid ${C.border}`, padding:"12px 18px" }}>
                      <div style={{ fontSize:11, fontWeight:600, color:C.textMuted, marginBottom:10 }}>{r.debugOrders?.length} commandes incluses</div>
                      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(200px, 1fr))", gap:6 }}>
                        {(r.debugOrders??[]).sort((a,b)=>a.name.localeCompare(b.name)).map(o => (
                          <a key={o.id} href={`${session.config.url}/web#id=${o.id}&model=sale.order&view_type=form`} target="_blank" rel="noreferrer"
                            style={{ display:"flex", flexDirection:"column", padding:"8px 10px", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, textDecoration:"none" }}>
                            <span style={{ fontSize:12,fontWeight:700,color:C.blue,fontFamily:"'SF Mono','Fira Code',monospace" }}>{o.name}</span>
                            {o.partnerName && <span style={{ fontSize:11,color:C.textMuted }}>{o.partnerName}</span>}
                          </a>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {/* Catch-alls */}
        {catchalls.map(c => {
          const d = c.data;
          if (!c.loading && (!d||(d.qtyTotal===0&&d.caTotal===0))) return null;
          const caKey = `catchall_${c.codeInterne}`;
          const isExp = expandedId === caKey;
          const mode = detailMode[caKey] || "delegues";
          return (
            <div key={c.codeInterne} style={{ marginTop:12, border:`1.5px dashed ${C.amber}88`, borderRadius:12, overflow:"hidden", background:C.white }}>
              <div style={{ padding:"12px 18px", background:C.amberSoft, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                <div>
                  <div style={{ fontSize:13,fontWeight:700,color:C.amber }}>{c.codeInterne}</div>
                  <div style={{ fontSize:11,color:C.amber,opacity:0.8 }}>Commandes sans code offre spécifique (note interne)</div>
                </div>
                {c.loading && <div style={{ width:16,height:16,border:`2px solid ${C.amber}`,borderTopColor:"transparent",borderRadius:"50%",animation:"spin 0.8s linear infinite" }}/>}
              </div>
              {!c.loading && d && (
                <div style={{ padding:"12px 18px" }}>
                  <div style={{ display:"flex", gap:16, marginBottom:12, alignItems:"center" }}>
                    <div><span style={{ fontSize:10,fontWeight:600,color:C.textMuted,textTransform:"uppercase",letterSpacing:"0.05em" }}>CA </span><span style={{ fontSize:18,fontWeight:800,color:C.green }}>{fmt(d.caTotal)}</span></div>
                    <div><span style={{ fontSize:10,fontWeight:600,color:C.textMuted,textTransform:"uppercase",letterSpacing:"0.05em" }}>Commandes </span><span style={{ fontSize:18,fontWeight:800,color:C.amber }}>{d.qtyTotal}</span></div>
                    <div style={{ flex:1 }}/>
                    <div style={{ display:"flex", gap:5 }}>
                      {([["produits",`Produits (${d.produits.length})`,C.blue],["delegues",`Délégués (${d.delegues.length})`,C.purple],["debug",`Cdes (${d.debugOrders?.length??0})`,C.slate]] as [string,string,string][]).map(([m,lbl,col]) => (
                        <button key={m} onClick={()=>{setExpandedId(isExp&&mode===m?null:caKey);setDetailMode(dm=>({...dm,[caKey]:m as any}));}}
                          style={{ padding:"5px 10px", background:isExp&&mode===m?col+"18":"transparent", border:`1px solid ${isExp&&mode===m?col:C.border}`, borderRadius:7, cursor:"pointer", fontSize:11, fontWeight:600, color:isExp&&mode===m?col:C.textMuted, fontFamily:"inherit" }}>
                          {lbl}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}
              {!c.loading && d && isExp && mode==="produits" && (
                <div style={{ borderTop:`1px solid ${C.border}` }}>
                  <table style={{ width:"100%",borderCollapse:"collapse",fontSize:12 }}>
                    <thead><tr style={{ background:C.bg }}><th style={{ padding:"8px 18px",textAlign:"left",fontWeight:600,color:C.textMuted,fontSize:11 }}>Produit</th><th style={{ padding:"8px 18px",textAlign:"right",fontWeight:600,color:C.textMuted,fontSize:11 }}>Qté</th><th style={{ padding:"8px 18px",textAlign:"right",fontWeight:600,color:C.textMuted,fontSize:11 }}>CA</th></tr></thead>
                    <tbody>{d.produits.length===0?<tr><td colSpan={3} style={{ padding:"12px 18px",color:C.textMuted,textAlign:"center" }}>Aucun produit</td></tr>:d.produits.map((p,i)=>(
                      <tr key={p.productId} style={{ borderTop:`1px solid ${C.border}`,background:i%2===0?C.white:C.bg+"80" }}>
                        <td style={{ padding:"9px 18px",color:C.textSec }}>{p.name}</td>
                        <td style={{ padding:"9px 18px",textAlign:"right",color:C.textSec }}>{Math.round(p.qtyVendue)}</td>
                        <td style={{ padding:"9px 18px",textAlign:"right",fontWeight:700,color:C.green }}>{fmt(p.ca)}</td>
                      </tr>
                    ))}</tbody>
                  </table>
                </div>
              )}
              {!c.loading && d && isExp && mode==="delegues" && (
                <div style={{ borderTop:`1px solid ${C.border}` }}>
                  {d.delegues.length===0?<div style={{ padding:"12px 18px",color:C.textMuted,fontSize:12 }}>Aucun délégué</div>:(
                    <table style={{ width:"100%",borderCollapse:"collapse",fontSize:12 }}>
                      <thead><tr style={{ background:C.bg }}><th style={{ padding:"8px 18px",textAlign:"left",fontWeight:600,color:C.textMuted,fontSize:11 }}>Délégué</th><th style={{ padding:"8px 18px",textAlign:"right",fontWeight:600,color:C.textMuted,fontSize:11 }}>CA</th><th style={{ padding:"8px 18px",textAlign:"right",fontWeight:600,color:C.textMuted,fontSize:11 }}>%</th></tr></thead>
                      <tbody>{d.delegues.map((del,i)=>(
                        <tr key={del.userId} style={{ borderTop:`1px solid ${C.border}`,background:i%2===0?C.white:C.bg+"80" }}>
                          <td style={{ padding:"9px 18px" }}><div style={{ display:"flex",alignItems:"center",gap:8 }}><div style={{ width:26,height:26,borderRadius:13,background:C.purpleSoft,display:"flex",alignItems:"center",justifyContent:"center" }}><span style={{ fontSize:10,fontWeight:800,color:C.purple }}>{initials(del.name)}</span></div><span style={{ fontWeight:600,color:C.textSec }}>{del.name}</span></div></td>
                          <td style={{ padding:"9px 18px",textAlign:"right",fontWeight:700,color:C.purple }}>{fmt(del.ca)}</td>
                          <td style={{ padding:"9px 18px",textAlign:"right",color:C.textMuted }}>{Math.round((del.ca/(d.caTotal||1))*100)}%</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  )}
                </div>
              )}
              {!c.loading && d && isExp && mode==="debug" && (
                <div style={{ borderTop:`1px solid ${C.border}`,padding:"12px 18px" }}>
                  <div style={{ display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(200px, 1fr))",gap:6 }}>
                    {(d.debugOrders??[]).sort((a,b)=>a.name.localeCompare(b.name)).map(o=>(
                      <a key={o.id} href={`${session.config.url}/web#id=${o.id}&model=sale.order&view_type=form`} target="_blank" rel="noreferrer" style={{ display:"flex",flexDirection:"column",padding:"8px 10px",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,textDecoration:"none" }}>
                        <span style={{ fontSize:12,fontWeight:700,color:C.blue,fontFamily:"'SF Mono','Fira Code',monospace" }}>{o.name}</span>
                        {o.partnerName && <span style={{ fontSize:11,color:C.textMuted }}>{o.partnerName}</span>}
                      </a>
                    ))}
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════════
// COMPOSANT PRINCIPAL
// ═════════════════════════════════════════════════════════════════════════════
export default function AnalyseScreen({ session, onToast }: Props) {
  const [filter, setFilter] = useState<StateFilter>("all");
  const [sharedCodes, setSharedCodes] = useState<string[]>([]);
  const [showOffresPanel, setShowOffresPanel] = useState(false);

  const FILTERS: [StateFilter, string][] = [["all","Tout"],["avenir","À venir"],["valide","Validé"]];

  return (
    <div style={{ flex:1, display:"flex", flexDirection:"column", height:"100%", overflow:"hidden", background:C.bg }}>
      {/* Topbar */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.border}`, padding:"0 24px", display:"flex", alignItems:"center", gap:16, height:56, flexShrink:0 }}>
        <div style={{ fontSize:15, fontWeight:700, color:C.text, flexShrink:0 }}>Analyse des Offres</div>
        <div style={{ flex:1 }}/>

        {/* Filter tabs */}
        <div style={{ display:"flex", gap:2, background:C.bg, border:`1px solid ${C.border}`, borderRadius:9, padding:3 }}>
          {FILTERS.map(([key,label]) => (
            <button key={key} onClick={()=>setFilter(key)}
              style={{ padding:"5px 14px", border:"none", borderRadius:7, cursor:"pointer", fontFamily:"inherit", fontSize:12, fontWeight:600,
                background:filter===key?C.white:"transparent",
                color:filter===key?(key==="valide"?C.green:key==="avenir"?C.amber:C.blue):C.textMuted,
                boxShadow:filter===key?C.shadow:"none" }}>
              {label}
            </button>
          ))}
        </div>

        {/* Gérer les offres */}
        <button onClick={()=>setShowOffresPanel(true)}
          style={{ padding:"7px 14px", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, cursor:"pointer", fontSize:12, fontWeight:600, color:C.textSec, fontFamily:"inherit", display:"flex", alignItems:"center", gap:6 }}>
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="3"/><path d="M19.07 4.93a10 10 0 010 14.14M4.93 4.93a10 10 0 000 14.14"/></svg>
          Gérer les offres
        </button>
      </div>

      {/* Contenu */}
      <AnalyseTab key={filter} session={session} onToast={onToast} filter={filter} sharedCodes={sharedCodes} onCodesChange={setSharedCodes} />

      {/* Panneau offres */}
      {showOffresPanel && <OffresPanel onClose={()=>setShowOffresPanel(false)} onToast={onToast} />}
    </div>
  );
}
