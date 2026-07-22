
"use client";
import { useState, useEffect, useRef } from "react";
import * as odoo from "@/lib/odoo";
import { supabase } from "@/lib/supabase";

import { C } from "@/lib/theme";

interface ComposantKit { ref: string; nom?: string; qtyParKit: number; dateArrivage?: string; }
type ProjetStatus = "planning" | "en_attente" | "stock_ok" | "reserve" | "envoye";
interface ProjetKit {
  id: string; nom: string; refFinale?: string; qtyKits: number;
  composants: ComposantKit[]; photo?: string;
  dateLancement?: string; dateEsat: string;
  notes?: string; status: ProjetStatus; createdAt: string;
}
interface StockInfo { ref: string; nom: string; dispo: number; productId: number; }
interface Props { session: odoo.OdooSession; onToast: (msg: string, type?: "success"|"error"|"info") => void; }

function genId() { return Math.random().toString(36).slice(2, 10); }

// ── Supabase CRUD ─────────────────────────────────────────────────────────────
function rowToProjet(row: any): ProjetKit {
  return {
    id: row.id, nom: row.nom, refFinale: row.ref_finale || undefined,
    qtyKits: row.qty_kits, composants: row.composants || [],
    photo: row.photo || undefined,
    dateLancement: row.date_lancement || undefined, dateEsat: row.date_esat,
    notes: row.notes || undefined, status: row.status as ProjetStatus,
    createdAt: row.created_at,
  };
}
function projetToRow(p: ProjetKit, userLogin: string) {
  return {
    id: p.id, user_login: userLogin, nom: p.nom,
    ref_finale: p.refFinale || null, qty_kits: p.qtyKits,
    composants: p.composants, photo: p.photo || null,
    date_lancement: p.dateLancement || null, date_esat: p.dateEsat,
    notes: p.notes || null, status: p.status, created_at: p.createdAt,
  };
}
async function dbLoad(userLogin: string): Promise<ProjetKit[]> {
  const { data, error } = await supabase
    .from("projets_kits").select("*").order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(rowToProjet);
}
async function dbUpsert(p: ProjetKit, userLogin: string): Promise<void> {
  const { error } = await supabase.from("projets_kits").upsert(projetToRow(p, userLogin));
  if (error) throw error;
}
async function dbUpdate(p: ProjetKit, userLogin: string): Promise<void> {
  const { error } = await supabase.from("projets_kits")
    .update(projetToRow(p, userLogin)).eq("id", p.id);
  if (error) throw error;
}

// Migration localStorage → Supabase (one-shot)
const LS_MIGRATE_KEY = "ao_kits_migrated_v1";
async function migrateFromLocalStorage(userLogin: string): Promise<number> {
  if (typeof window === "undefined") return 0;
  if (localStorage.getItem(LS_MIGRATE_KEY)) return 0;
  const raw = localStorage.getItem("ao_projets_kits");
  if (!raw) { localStorage.setItem(LS_MIGRATE_KEY, "1"); return 0; }
  try {
    const projets: ProjetKit[] = JSON.parse(raw);
    if (!projets.length) { localStorage.setItem(LS_MIGRATE_KEY, "1"); return 0; }
    const rows = projets.map(p => projetToRow(p, userLogin));
    const { error } = await supabase.from("projets_kits").upsert(rows, { onConflict: "id" });
    if (error) throw error;
    localStorage.setItem(LS_MIGRATE_KEY, "1");
    return projets.length;
  } catch(e) { console.error("Migration failed", e); return 0; }
}
function fmt(n: number) { return new Intl.NumberFormat("fr-FR").format(n); }
function fmtDate(iso?: string) {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short", year: "numeric" });
}
function daysUntil(iso: string): number {
  if (!iso) return Infinity;
  return Math.ceil((new Date(iso + "T00:00:00").getTime() - Date.now()) / 86400000);
}

const STATUS_CONFIG: Record<ProjetStatus, { label: string; color: string; bg: string; step: number; }> = {
  planning:   { label: "Planification",    color: C.slate,  bg: "#f1f5f9",   step: 0 },
  en_attente: { label: "En attente stock", color: C.amber,  bg: C.amberSoft, step: 1 },
  stock_ok:   { label: "Stock prêt",  color: C.green,  bg: C.greenSoft, step: 2 },
  reserve:    { label: "Réservé",color: C.blue,   bg: C.blueSoft,  step: 3 },
  envoye:     { label: "Envoyé ESAT", color: C.purple, bg: C.purpleSoft,step: 4 },
};
const STATUSES: ProjetStatus[] = ["planning", "en_attente", "stock_ok", "reserve", "envoye"];

async function fetchStock(session: odoo.OdooSession, refs: string[]): Promise<StockInfo[]> {
  if (!refs.length) return [];
  const prods = await odoo.searchRead(
    session, "product.product",
    [["default_code", "in", refs]],
    ["id", "default_code", "name", "qty_available"], 0
  );
  return prods.map((p: any) => ({
    ref: p.default_code, nom: p.name,
    dispo: Math.floor(p.qty_available || 0), productId: p.id,
  }));
}

// ─── FORMULAIRE ────────────────────────────────────────────────────────────────
function FormulaireProjet({ projet, onSave, onCancel, onToast }: {
  projet?: ProjetKit; onSave: (p: ProjetKit) => void;
  onCancel: () => void; onToast: Props["onToast"];
}) {
  const [nom, setNom] = useState(projet?.nom || "");
  const [refFinale, setRefFinale] = useState(projet?.refFinale || "");
  const [qtyKits, setQtyKits] = useState(String(projet?.qtyKits || ""));
  const [dateLancement, setDateLancement] = useState(projet?.dateLancement || "");
  const [dateEsat, setDateEsat] = useState(projet?.dateEsat || "");
  const [notes, setNotes] = useState(projet?.notes || "");
  const [composants, setComposants] = useState<ComposantKit[]>(projet?.composants || []);
  const [photo, setPhoto] = useState<string|undefined>(projet?.photo);
  const [newRef, setNewRef] = useState("");
  const [newQty, setNewQty] = useState("1");
  const [newDate, setNewDate] = useState("");
  const photoRef = useRef<HTMLInputElement>(null);

  const handlePhoto = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => setPhoto(ev.target?.result as string);
    reader.readAsDataURL(f);
  };
  const addComposant = () => {
    const ref = newRef.trim().toUpperCase(); if (!ref) return;
    if (composants.some(c => c.ref.toLowerCase() === ref.toLowerCase())) { onToast("Référence déjà ajoutée", "error"); return; }
    setComposants(prev => [...prev, { ref, qtyParKit: parseInt(newQty)||1, dateArrivage: newDate||undefined }]);
    setNewRef(""); setNewQty("1"); setNewDate("");
  };
  const removeComposant = (ref: string) => setComposants(prev => prev.filter(c => c.ref !== ref));
  const updateQty = (ref: string, qty: number) => setComposants(prev => prev.map(c => c.ref===ref ? { ...c, qtyParKit: Math.max(1,qty) } : c));
  const updateDate = (ref: string, date: string) => setComposants(prev => prev.map(c => c.ref===ref ? { ...c, dateArrivage: date||undefined } : c));
  const save = () => {
    if (!nom.trim()) { onToast("Nom du projet requis", "error"); return; }
    const qty = parseInt(qtyKits);
    if (!qty || qty <= 0) { onToast("Quantité de kits invalide", "error"); return; }
    if (!composants.length) { onToast("Au moins un composant requis", "error"); return; }
    if (!dateEsat) { onToast("Date ESAT requise", "error"); return; }
    onSave({ id: projet?.id||genId(), nom: nom.trim(), refFinale: refFinale.trim()||undefined,
      qtyKits: qty, composants, photo, dateLancement: dateLancement||undefined, dateEsat,
      notes: notes.trim()||undefined, status: projet?.status||"planning",
      createdAt: projet?.createdAt||new Date().toISOString() });
  };

  const inp: React.CSSProperties = { padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:13,
    fontFamily:"inherit", background:C.surface, color:C.text, outline:"none", width:"100%", boxSizing:"border-box" as const };
  const lbl: React.CSSProperties = { fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase" as const,
    letterSpacing:"0.07em", display:"block" as const, marginBottom:5 };

  return (
    <div style={{ flex:1, overflowY:"auto", padding:"24px 32px" }}>
      <div style={{ maxWidth:920, margin:"0 auto" }}>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:28 }}>
          <button onClick={onCancel} style={{ width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center",background:C.surface,border:`1px solid ${C.border}`,borderRadius:9,cursor:"pointer",flexShrink:0 }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.text} strokeWidth="2"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
          </button>
          <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:C.text }}>{projet ? "Modifier le projet" : "Nouveau projet kit"}</h2>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"148px 1fr", gap:20, marginBottom:20 }}>
          {/* Photo upload */}
          <div>
            <label style={lbl}>Photo produit</label>
            <div onClick={() => photoRef.current?.click()} style={{ width:148, height:148, borderRadius:14, border:`2px dashed ${projet?.photo||photo?"transparent":C.borderDark}`, background:photo?"transparent":C.bg, cursor:"pointer", overflow:"hidden", display:"flex", alignItems:"center", justifyContent:"center" }}>
              {photo
                ? <img src={photo} style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                : <div style={{ textAlign:"center", color:C.textMuted }}>
                    <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>
                    <div style={{ fontSize:10, marginTop:4, lineHeight:1.4 }}>Cliquer pour<br/>ajouter</div>
                  </div>}
            </div>
            <input ref={photoRef} type="file" accept="image/*" onChange={handlePhoto} style={{ display:"none" }} />
            {photo && <button onClick={e=>{e.stopPropagation();setPhoto(undefined);}} style={{ marginTop:6, fontSize:11, color:C.red, background:"none", border:"none", cursor:"pointer", padding:0, fontFamily:"inherit" }}>Supprimer</button>}
          </div>

          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 140px", gap:12 }}>
              <div>
                <label style={lbl}>Nom du projet *</label>
                <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Ex: Vanity Noël 2026" style={inp} />
              </div>
              <div>
                <label style={lbl}>Qté de kits *</label>
                <input type="number" min="1" value={qtyKits} onChange={e=>setQtyKits(e.target.value)} placeholder="500" style={inp} />
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:12 }}>
              <div>
                <label style={lbl}>Référence finale</label>
                <input value={refFinale} onChange={e=>setRefFinale(e.target.value.toUpperCase())} placeholder="Ex: 9050001"
                  style={{ ...inp, fontFamily:"'SF Mono','Fira Code',monospace" }} />
              </div>
              <div>
                <label style={lbl}>Date de lancement</label>
                <input type="date" value={dateLancement} onChange={e=>setDateLancement(e.target.value)} style={inp} />
              </div>
              <div>
                <label style={lbl}>Date envoi ESAT *</label>
                <input type="date" value={dateEsat} onChange={e=>setDateEsat(e.target.value)} style={inp} />
              </div>
            </div>
            <div>
              <label style={lbl}>Notes (optionnel)</label>
              <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Informations complémentaires…" rows={2}
                style={{ ...inp, resize:"vertical" as const }} />
            </div>
          </div>
        </div>

        {/* Composants */}
        <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:12, padding:"18px 20px", marginBottom:16 }}>
          <div style={{ fontSize:12, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:14 }}>Composants du kit</div>
          <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
            <input value={newRef} onChange={e=>setNewRef(e.target.value.toUpperCase())}
              onKeyDown={e=>{ if(e.key==="Enter") addComposant(); }}
              placeholder="Référence (ex: 1010214)"
              style={{ padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:12, fontFamily:"'SF Mono','Fira Code',monospace", flex:"1 1 160px", outline:"none", background:C.surface, color:C.text }} />
            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ fontSize:11, color:C.textMuted }}>Qté/kit</span>
              <input type="number" min="1" value={newQty} onChange={e=>setNewQty(e.target.value)}
                style={{ padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:13, width:70, textAlign:"center" as const, outline:"none", background:C.surface, color:C.text }} />
            </div>
            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ fontSize:11, color:C.textMuted }}>Arrivage</span>
              <input type="date" value={newDate} onChange={e=>setNewDate(e.target.value)}
                style={{ padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:12, outline:"none", background:C.surface, color:C.text }} />
            </div>
            <button onClick={addComposant} style={{ padding:"9px 18px", background:C.blue, color:"#fff", border:"none", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit" }}>
              + Ajouter
            </button>
          </div>
          {composants.length === 0
            ? <div style={{ textAlign:"center", padding:"20px 0", color:C.textMuted, fontSize:12 }}>Aucun composant ajouté</div>
            : <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:C.bg }}>
                    {["Référence","Qté / kit","Qté totale","Date arrivage",""].map((h,i) => (
                      <th key={i} style={{ padding:"7px 12px", textAlign:i>=1&&i<=3?"center" as const:"left" as const, fontWeight:600, color:C.textMuted, fontSize:11 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {composants.map(c => {
                    const qtyTotal = c.qtyParKit * (parseInt(qtyKits)||0);
                    return (
                      <tr key={c.ref} style={{ borderTop:`1px solid ${C.border}` }}>
                        <td style={{ padding:"9px 12px", fontFamily:"'SF Mono','Fira Code',monospace", fontWeight:700, color:C.blue, fontSize:12 }}>{c.ref}</td>
                        <td style={{ padding:"9px 12px", textAlign:"center" }}>
                          <input type="number" min="1" value={c.qtyParKit} onChange={e=>updateQty(c.ref, parseInt(e.target.value)||1)}
                            style={{ width:70, padding:"4px 8px", border:`1px solid ${C.border}`, borderRadius:6, fontSize:12, textAlign:"center" as const, background:C.bg }} />
                        </td>
                        <td style={{ padding:"9px 12px", textAlign:"center", fontWeight:700, color:C.text }}>{qtyTotal > 0 ? fmt(qtyTotal) : "—"}</td>
                        <td style={{ padding:"9px 12px", textAlign:"center" }}>
                          <input type="date" value={c.dateArrivage||""} onChange={e=>updateDate(c.ref, e.target.value)}
                            style={{ padding:"4px 8px", border:`1px solid ${C.border}`, borderRadius:6, fontSize:11, background:C.bg }} />
                        </td>
                        <td style={{ padding:"9px 12px", textAlign:"right" }}>
                          <button onClick={()=>removeComposant(c.ref)} style={{ background:"none", border:"none", cursor:"pointer", color:C.red, padding:4 }}>
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/></svg>
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
          }
        </div>

        <div style={{ display:"flex", gap:10 }}>
          <button onClick={onCancel} style={{ flex:1, padding:"11px 0", background:C.surface, color:C.textSec, border:`1px solid ${C.border}`, borderRadius:10, fontSize:14, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>Annuler</button>
          <button onClick={save} style={{ flex:2, padding:"11px 0", background:C.blue, color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
            {projet ? "Enregistrer les modifications" : "Créer le projet"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── DÉTAIL PROJET ─────────────────────────────────────────────────────────────
function DetailProjet({ projet, session, onBack, onUpdate, onToast }: {
  projet: ProjetKit; session: odoo.OdooSession;
  onBack: () => void; onUpdate: (p: ProjetKit) => void; onToast: Props["onToast"];
}) {
  const [stock, setStock] = useState<StockInfo[]>([]);
  const [loading, setLoading] = useState(false);
  const [lastCheck, setLastCheck] = useState<Date|null>(null);
  const [editing, setEditing] = useState(false);
  const photoRef = useRef<HTMLInputElement>(null);

  useEffect(() => { verifierStock(); }, []);

  const verifierStock = async () => {
    setLoading(true);
    try {
      const refs = projet.composants.map(c => c.ref);
      const s = await fetchStock(session, refs);
      setStock(s);
      setLastCheck(new Date());
      const allOk = projet.composants.every(c => {
        const info = s.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
        return info && info.dispo >= c.qtyParKit * projet.qtyKits;
      });
      if (allOk && projet.status === "en_attente") {
        onUpdate({ ...projet, status: "stock_ok" });
        onToast("⚡ Stock disponible — réservation possible", "success");
      } else if (allOk) {
        onToast("Stock vérifié — tout est disponible ✅", "success");
      } else {
        onToast("Stock vérifié — certains produits manquent", "info");
      }
    } catch(e:any) { onToast("Erreur stock : " + e.message, "error"); }
    finally { setLoading(false); }
  };

  const handlePhoto = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => { onUpdate({ ...projet, photo: ev.target?.result as string }); onToast("Photo mise à jour", "success"); };
    reader.readAsDataURL(f);
  };

  const setStatus = (s: ProjetStatus) => { onUpdate({ ...projet, status: s }); onToast("Statut mis à jour", "success"); };

  if (editing) return (
    <FormulaireProjet projet={projet}
      onSave={p => { onUpdate(p); setEditing(false); onToast("Projet mis à jour", "success"); }}
      onCancel={() => setEditing(false)} onToast={onToast} />
  );

  const cfg = STATUS_CONFIG[projet.status];
  const currentStep = cfg.step;
  const daysEsat = daysUntil(projet.dateEsat);
  const allStockOk = stock.length > 0 && projet.composants.every(c => {
    const info = stock.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
    return info && info.dispo >= c.qtyParKit * projet.qtyKits;
  });
  const datesArr = projet.composants.filter(c=>c.dateArrivage).map(c=>c.dateArrivage!).sort();
  const lastArr = datesArr.at(-1);
  const isLoadingFirst = loading && stock.length === 0;

  return (
    <div style={{ flex:1, overflowY:"auto", background:C.bg }}>
      {/* ── Hero ── */}
      <div style={{ background:C.surface, borderBottom:`1px solid ${C.border}`, padding:"20px 28px 24px" }}>
        {/* Breadcrumb + actions */}
        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:20 }}>
          <button onClick={onBack} style={{ height:30, padding:"0 10px", display:"flex", alignItems:"center", gap:5, background:"transparent", border:`1px solid ${C.border}`, borderRadius:7, cursor:"pointer", fontSize:12, color:C.textSec, fontFamily:"inherit" }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
            Liste
          </button>
          <span style={{ fontSize:12, color:C.border }}>›</span>
          <span style={{ fontSize:12, color:C.textMuted }}>Projets Kits</span>
          <span style={{ fontSize:12, color:C.border }}>›</span>
          <span style={{ fontSize:12, color:C.text, fontWeight:600 }}>{projet.nom}</span>
          <div style={{ flex:1 }}/>
          <button onClick={() => setEditing(true)} style={{ height:32, padding:"0 14px", display:"flex", alignItems:"center", gap:6, background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, cursor:"pointer", fontSize:12, fontWeight:600, color:C.textSec, fontFamily:"inherit" }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            Modifier
          </button>
        </div>

        {/* Hero content */}
        <div style={{ display:"flex", gap:22, alignItems:"flex-start" }}>
          {/* Photo zone */}
          <div style={{ position:"relative", flexShrink:0 }}>
            <div onClick={() => photoRef.current?.click()} style={{ width:110, height:110, borderRadius:16, border:`2px ${projet.photo?"solid transparent":"dashed "+C.borderDark}`, background:projet.photo?"transparent":C.bg, cursor:"pointer", overflow:"hidden", display:"flex", alignItems:"center", justifyContent:"center", boxShadow:projet.photo?"0 4px 20px rgba(0,0,0,0.12)":"none" }}>
              {projet.photo
                ? <img src={projet.photo} style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                : <div style={{ textAlign:"center", color:C.textMuted }}>
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>
                    <div style={{ fontSize:9, marginTop:4, lineHeight:1.4, opacity:0.6 }}>Photo<br/>produit</div>
                  </div>}
            </div>
            {projet.photo && (
              <button onClick={() => photoRef.current?.click()} style={{ position:"absolute", bottom:-6, right:-6, width:24, height:24, borderRadius:"50%", background:C.blue, border:`2px solid ${C.surface}`, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center" }}>
                <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
              </button>
            )}
            <input ref={photoRef} type="file" accept="image/*" onChange={handlePhoto} style={{ display:"none" }} />
          </div>

          {/* Infos */}
          <div style={{ flex:1, minWidth:0 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:6 }}>
              <h1 style={{ margin:0, fontSize:24, fontWeight:800, color:C.text, letterSpacing:"-0.3px" }}>{projet.nom}</h1>
              <span style={{ padding:"4px 12px", borderRadius:20, fontSize:11, fontWeight:700, color:cfg.color, background:cfg.bg, border:`1px solid ${cfg.color}33`, whiteSpace:"nowrap" }}>{cfg.label}</span>
            </div>
            <div style={{ display:"flex", gap:14, alignItems:"center", flexWrap:"wrap", marginBottom:14 }}>
              {projet.refFinale && (
                <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                  <span style={{ fontSize:11, color:C.textMuted, fontWeight:500 }}>Réf finale</span>
                  <span style={{ fontSize:13, fontWeight:700, color:C.text, fontFamily:"'SF Mono','Fira Code',monospace", background:C.bg, border:`1px solid ${C.border}`, borderRadius:6, padding:"2px 8px" }}>{projet.refFinale}</span>
                </div>
              )}
              <span style={{ fontSize:12, color:C.textMuted }}>{fmt(projet.qtyKits)} kits</span>
              <span style={{ fontSize:12, color:C.border }}>·</span>
              <span style={{ fontSize:12, color:C.textMuted }}>{projet.composants.length} composant{projet.composants.length!==1?"s":""}</span>
            </div>

            {/* KPI chips */}
            <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
              {[
                ...(projet.dateLancement ? [{
                  emoji:"🚀", label:"Lancement", value:fmtDate(projet.dateLancement),
                  sub: daysUntil(projet.dateLancement)>0?`J-${daysUntil(projet.dateLancement)}`:"passé",
                  urgent: daysUntil(projet.dateLancement)<=0, warn: daysUntil(projet.dateLancement)<=7&&daysUntil(projet.dateLancement)>0,
                }] : []),
                ...(lastArr ? [{
                  emoji:"📦", label:"Dernier arrivage", value:fmtDate(lastArr),
                  sub: daysUntil(lastArr)>0?`J-${daysUntil(lastArr)}`:"arrivé",
                  urgent: false, warn: daysUntil(lastArr)<=14&&daysUntil(lastArr)>0,
                }] : []),
                {
                  emoji:"🏭", label:"Envoi ESAT", value:fmtDate(projet.dateEsat),
                  sub: daysEsat>0?`J-${daysEsat}`:"passé",
                  urgent: daysEsat<=0, warn: daysEsat<=7&&daysEsat>0,
                },
              ].map(item => (
                <div key={item.label} style={{ display:"flex", alignItems:"center", gap:10, background:item.urgent?C.redSoft:item.warn?C.amberSoft:C.bg, border:`1px solid ${item.urgent?C.red+"33":item.warn?C.amber+"33":C.border}`, borderRadius:10, padding:"8px 14px" }}>
                  <span style={{ fontSize:16 }}>{item.emoji}</span>
                  <div>
                    <div style={{ fontSize:10, color:C.textMuted, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.04em" }}>{item.label}</div>
                    <div style={{ fontSize:13, fontWeight:700, color:item.urgent?C.red:item.warn?C.amber:C.text }}>{item.value}</div>
                    <div style={{ fontSize:10, color:item.urgent?C.red:item.warn?C.amber:C.textMuted, opacity:0.8 }}>{item.sub}</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      <div style={{ padding:"20px 28px", display:"flex", flexDirection:"column", gap:14 }}>

        {/* Alerte stock OK */}
        {(projet.status==="stock_ok"||allStockOk) && projet.status!=="reserve" && projet.status!=="envoye" && (
          <div style={{ background:"linear-gradient(135deg,#ecfdf5,#d1fae5)", border:`1.5px solid ${C.green}55`, borderRadius:14, padding:"14px 20px", display:"flex", alignItems:"center", gap:16 }}>
            <div style={{ width:44,height:44,background:C.green+"22",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={C.green} strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            </div>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:14, fontWeight:700, color:C.greenDark }}>Stock disponible — prêt à réserver</div>
              <div style={{ fontSize:12, color:C.green, marginTop:2 }}>Tous les composants ont le stock nécessaire.</div>
            </div>
            <button onClick={() => setStatus("reserve")} style={{ padding:"10px 20px", background:C.green, color:"#fff", border:"none", borderRadius:10, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit", whiteSpace:"nowrap" }}>
              Marquer réservé
            </button>
          </div>
        )}

        {/* Table composants */}
        <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:14, overflow:"hidden" }}>
          <div style={{ padding:"14px 20px", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
            <span style={{ fontSize:13, fontWeight:700, color:C.text }}>Composants du kit</span>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              {lastCheck && <span style={{ fontSize:11, color:C.textMuted }}>Vérifié à {lastCheck.toLocaleTimeString("fr-FR",{hour:"2-digit",minute:"2-digit"})}</span>}
              <button onClick={verifierStock} disabled={loading}
                style={{ padding:"7px 14px", background:loading?C.bg:`${C.blue}18`, border:`1px solid ${loading?C.border:`${C.blue}44`}`, borderRadius:8, cursor:loading?"default":"pointer", fontSize:12, fontWeight:600, color:loading?C.textMuted:C.blue, fontFamily:"inherit", display:"flex", alignItems:"center", gap:6 }}>
                {loading
                  ? <><span style={{ width:12,height:12,border:`2px solid ${C.blue}`,borderTopColor:"transparent",borderRadius:"50%",display:"inline-block",animation:"spin 0.8s linear infinite" }}/>Vérification…</>
                  : <><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>Vérifier</>
                }
              </button>
            </div>
          </div>

          <table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead>
              <tr style={{ background:"#f8fafc" }}>
                <th style={{ padding:"10px 20px", textAlign:"left" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>RÉFÉRENCE</th>
                <th style={{ padding:"10px 16px", textAlign:"left" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>DÉSIGNATION</th>
                <th style={{ padding:"10px 16px", textAlign:"center" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>ARRIVAGE PRÉVU</th>
                <th style={{ padding:"10px 16px", textAlign:"right" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>QTÉ/KIT</th>
                <th style={{ padding:"10px 16px", textAlign:"right" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>BESOIN TOTAL</th>
                <th style={{ padding:"10px 16px", textAlign:"right" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>DISPO ODOO</th>
                <th style={{ padding:"10px 20px", textAlign:"center" as const, fontSize:11, fontWeight:600, color:C.textMuted, letterSpacing:"0.04em" }}>ÉCART</th>
              </tr>
            </thead>
            <tbody>
              {projet.composants.map((c, i) => {
                const qtyTotale = c.qtyParKit * projet.qtyKits;
                const info = stock.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
                const ecart = info ? info.dispo - qtyTotale : null;
                const ok = ecart !== null && ecart >= 0;
                return (
                  <tr key={c.ref} style={{ borderTop:`1px solid ${C.border}`, background:i%2===0?C.surface:"#fafbfc" }}>
                    <td style={{ padding:"13px 20px" }}>
                      <span style={{ fontFamily:"'SF Mono','Fira Code',monospace", fontSize:12, fontWeight:700, color:C.blue, background:C.blueSoft, padding:"3px 8px", borderRadius:5 }}>{c.ref}</span>
                    </td>
                    <td style={{ padding:"13px 16px", fontSize:13, color:C.textSec }}>
                      {isLoadingFirst
                        ? <span style={{ display:"inline-block", width:160, height:13, background:C.border, borderRadius:4, animation:"pulse 1.5s ease-in-out infinite" }}/>
                        : (info?.nom || c.nom || <span style={{ color:C.textMuted, fontStyle:"italic", fontSize:12 }}>—</span>)}
                    </td>
                    <td style={{ padding:"13px 16px", textAlign:"center" }}>
                      {c.dateArrivage
                        ? <span style={{ fontSize:12, fontWeight:600, color:daysUntil(c.dateArrivage)<=0?C.green:daysUntil(c.dateArrivage)<=14?C.amber:C.textSec }}>{fmtDate(c.dateArrivage)}</span>
                        : <span style={{ fontSize:12, color:C.textMuted }}>—</span>}
                    </td>
                    <td style={{ padding:"13px 16px", textAlign:"right", fontSize:13, color:C.textSec }}>{fmt(c.qtyParKit)}</td>
                    <td style={{ padding:"13px 16px", textAlign:"right", fontSize:13, fontWeight:700, color:C.text }}>{fmt(qtyTotale)}</td>
                    <td style={{ padding:"13px 16px", textAlign:"right" }}>
                      {isLoadingFirst
                        ? <span style={{ display:"inline-block", width:50, height:13, background:C.border, borderRadius:4, animation:"pulse 1.5s ease-in-out infinite", float:"right" as const }}/>
                        : info
                          ? <span style={{ fontSize:13, fontWeight:700, color:ok?C.green:C.red }}>{fmt(info.dispo)}</span>
                          : <span style={{ fontSize:12, color:C.textMuted }}>—</span>}
                    </td>
                    <td style={{ padding:"13px 20px", textAlign:"center" }}>
                      {isLoadingFirst
                        ? <span style={{ display:"inline-block", width:60, height:22, background:C.border, borderRadius:20, animation:"pulse 1.5s ease-in-out infinite" }}/>
                        : !info
                          ? <span style={{ fontSize:11, color:C.textMuted }}>—</span>
                          : ok
                          ? <span style={{ fontSize:12, fontWeight:700, color:C.green, background:C.greenSoft, border:`1px solid ${C.green}33`, padding:"4px 10px", borderRadius:20 }}>+{fmt(ecart!)}</span>
                          : <span style={{ fontSize:12, fontWeight:700, color:C.red, background:C.redSoft, border:`1px solid ${C.red}33`, padding:"4px 10px", borderRadius:20 }}>−{fmt(Math.abs(ecart!))}</span>
                      }
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Progression stepper */}
        <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:14, padding:"18px 28px" }}>
          <div style={{ fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:20 }}>Progression</div>
          <div style={{ display:"flex", alignItems:"flex-start", gap:0 }}>
            {STATUSES.map((s, i) => {
              const scfg = STATUS_CONFIG[s];
              const isActive = projet.status === s;
              const isDone = currentStep > i;
              return (
                <div key={s} style={{ display:"flex", alignItems:"center", flex: i < STATUSES.length-1 ? 1 : "none" }}>
                  <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:8 }}>
                    <button onClick={() => setStatus(s)}
                      style={{ width:38,height:38,borderRadius:"50%",border:`2.5px solid ${isActive||isDone?scfg.color:C.border}`,background:isActive?scfg.color:isDone?scfg.bg:C.surface,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",transition:"all 0.15s",boxShadow:isActive?`0 4px 14px ${scfg.color}44`:"none" }}>
                      {isDone
                        ? <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={scfg.color} strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
                        : <div style={{ width:8,height:8,borderRadius:"50%",background:isActive?"#fff":C.border }}/>}
                    </button>
                    <span style={{ fontSize:10,fontWeight:isActive?700:500,color:isActive?scfg.color:isDone?scfg.color:C.textMuted,whiteSpace:"nowrap",textAlign:"center" as const,maxWidth:80,lineHeight:1.3 }}>{scfg.label}</span>
                  </div>
                  {i < STATUSES.length-1 && (
                    <div style={{ flex:1,height:2.5,background:isDone?`linear-gradient(90deg,${scfg.color}66,${STATUS_CONFIG[STATUSES[i+1]].color}44)`:C.border,margin:"0 6px",marginBottom:26,borderRadius:2 }}/>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Notes */}
        {projet.notes && (
          <div style={{ background:C.amberSoft, border:`1px solid ${C.amber}33`, borderRadius:12, padding:"13px 18px", fontSize:13, color:C.textSec, display:"flex", gap:10, alignItems:"flex-start" }}>
            <span style={{ fontSize:16, flexShrink:0 }}>📝</span>
            <div><span style={{ fontWeight:700, color:C.amber }}>Note : </span>{projet.notes}</div>
          </div>
        )}
      </div>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}`}</style>
    </div>
  );
}

// ─── LISTE ──────────────────────────────────────────────────────────────────────
function ListeProjets({ projets, onSelect, onNew }: { projets: ProjetKit[]; onSelect: (p:ProjetKit)=>void; onNew:()=>void; }) {
  const actifs = projets.filter(p => p.status!=="envoye");
  const envoyes = projets.filter(p => p.status==="envoye");

  const Card = ({ p }: { p: ProjetKit }) => {
    const cfg = STATUS_CONFIG[p.status];
    const daysE = daysUntil(p.dateEsat);
    const datesArr = p.composants.filter(c=>c.dateArrivage).map(c=>c.dateArrivage!).sort();
    const lastArr = datesArr.at(-1);
    const daysLast = lastArr ? daysUntil(lastArr) : Infinity;
    return (
      <div onClick={() => onSelect(p)}
        style={{ background:C.surface, border:`1.5px solid ${p.status==="stock_ok"?C.green+"55":C.border}`, borderRadius:14, padding:"0", cursor:"pointer", overflow:"hidden", boxShadow:p.status==="stock_ok"?`0 4px 20px ${C.green}18`:"0 1px 4px rgba(0,0,0,0.05)", transition:"box-shadow 0.15s, transform 0.1s" }}>
        <div style={{ display:"flex", gap:0 }}>
          {/* Photo thumbnail */}
          <div style={{ width:76, minHeight:76, background:C.bg, flexShrink:0, display:"flex", alignItems:"center", justifyContent:"center", overflow:"hidden", borderRight:`1px solid ${C.border}` }}>
            {p.photo
              ? <img src={p.photo} style={{ width:"100%", height:"100%", objectFit:"cover", minHeight:76 }} />
              : <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={C.border} strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>}
          </div>
          <div style={{ flex:1, padding:"14px 16px" }}>
            <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:8, marginBottom:4 }}>
              <div>
                <div style={{ fontSize:14, fontWeight:700, color:C.text, marginBottom:2 }}>{p.nom}</div>
                {p.refFinale && <div style={{ fontSize:11, color:C.textMuted, fontFamily:"'SF Mono','Fira Code',monospace" }}>Réf: {p.refFinale}</div>}
              </div>
              <span style={{ fontSize:11, fontWeight:700, color:cfg.color, background:cfg.bg, border:`1px solid ${cfg.color}33`, padding:"3px 9px", borderRadius:20, whiteSpace:"nowrap", flexShrink:0 }}>{cfg.label}</span>
            </div>
            <div style={{ fontSize:12, color:C.textMuted, marginBottom:8 }}>{fmt(p.qtyKits)} kits · {p.composants.length} composants</div>
            <div style={{ display:"flex", gap:12, fontSize:11, color:C.textMuted, flexWrap:"wrap" }}>
              {lastArr && <span style={{ color:daysLast<=7&&daysLast>0?C.amber:daysLast<=0?C.green:C.textMuted }}>
                📦 {fmtDate(lastArr)}{daysLast>0&&daysLast<=30?` (J-${daysLast})`:daysLast<=0?" ✓":""}
              </span>}
              <span style={{ color:daysE<=7&&daysE>0?C.amber:daysE<=0?C.red:C.textMuted }}>
                🏭 ESAT {fmtDate(p.dateEsat)}{daysE>0&&daysE<=30?` (J-${daysE})`:""}
              </span>
            </div>
          </div>
        </div>
        {p.status==="stock_ok" && (
          <div style={{ padding:"8px 16px 10px", background:C.greenSoft, borderTop:`1px solid ${C.green}22`, fontSize:12, fontWeight:600, color:C.green, display:"flex", alignItems:"center", gap:6 }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            Stock disponible — réservation possible
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={{ flex:1, overflowY:"auto", padding:"24px 28px" }}>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:24 }}>
        <div>
          <h2 style={{ margin:0, fontSize:20, fontWeight:800, color:C.text }}>Projets Kits</h2>
          <p style={{ margin:"3px 0 0", fontSize:13, color:C.textMuted }}>{projets.length} projet{projets.length!==1?"s":""}</p>
        </div>
        <button onClick={onNew} style={{ padding:"9px 18px", background:C.blue, color:"#fff", border:"none", borderRadius:9, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit", display:"flex", alignItems:"center", gap:7 }}>
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Nouveau projet
        </button>
      </div>

      {projets.length === 0
        ? <div style={{ textAlign:"center", padding:"80px 24px" }}>
            <div style={{ width:64,height:64,borderRadius:16,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px" }}>
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={C.textMuted} strokeWidth="1.5"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/></svg>
            </div>
            <div style={{ fontSize:15, fontWeight:600, color:C.textSec, marginBottom:6 }}>Aucun projet kit</div>
            <div style={{ fontSize:13, color:C.textMuted, marginBottom:20 }}>Créez votre premier projet pour suivre vos kits</div>
            <button onClick={onNew} style={{ padding:"10px 22px", background:C.blue, color:"#fff", border:"none", borderRadius:9, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit" }}>Créer un projet</button>
          </div>
        : <>
            {actifs.length > 0 && (
              <div style={{ marginBottom:28 }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:12 }}>Projets en cours</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(400px,1fr))", gap:12 }}>
                  {actifs.map(p => <Card key={p.id} p={p} />)}
                </div>
              </div>
            )}
            {envoyes.length > 0 && (
              <div>
                <div style={{ fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:12 }}>Envoyés à l’ESAT</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(400px,1fr))", gap:12, opacity:0.6 }}>
                  {envoyes.map(p => <Card key={p.id} p={p} />)}
                </div>
              </div>
            )}
          </>}
    </div>
  );
}

// ─── COMPOSANT PRINCIPAL ────────────────────────────────────────────────────────
export default function ProjetScreen({ session, onToast }: Props) {
  const [projets, setProjets] = useState<ProjetKit[]>([]);
  const [view, setView] = useState<"liste"|"detail"|"nouveau">("liste");
  const [selectedId, setSelectedId] = useState<string|null>(null);
  const [dbLoading, setDbLoading] = useState(true);

  useEffect(() => { loadAll(); }, []);

  const loadAll = async () => {
    setDbLoading(true);
    try {
      const migrated = await migrateFromLocalStorage(session.login);
      if (migrated > 0) onToast(`${migrated} projet(s) importé(s) depuis la version précédente`, "success");
      const data = await dbLoad(session.login);
      setProjets(data);
    } catch (e: any) {
      onToast("Erreur chargement : " + e.message, "error");
    } finally {
      setDbLoading(false);
    }
  };

  const selected = projets.find(p => p.id === selectedId) || null;

  const handleSave = async (p: ProjetKit) => {
    try {
      await dbUpsert(p, session.login);
      const exists = projets.find(x => x.id === p.id);
      setProjets(exists ? projets.map(x => x.id === p.id ? p : x) : [p, ...projets]);
      setSelectedId(p.id); setView("detail");
      onToast("Projet enregistré", "success");
    } catch (e: any) {
      onToast("Erreur sauvegarde : " + e.message, "error");
    }
  };
  const handleUpdate = async (p: ProjetKit) => {
    try {
      await dbUpdate(p, session.login);
      setProjets(prev => prev.map(x => x.id === p.id ? p : x));
    } catch (e: any) {
      onToast("Erreur mise à jour : " + e.message, "error");
    }
  };

  return (
    <div style={{ flex:1, display:"flex", flexDirection:"column", height:"100%", overflow:"hidden", background:C.bg }}>
      <div style={{ background:C.surface, borderBottom:`1px solid ${C.border}`, padding:"0 24px", display:"flex", alignItems:"center", height:54, flexShrink:0 }}>
        <div style={{ fontSize:14, fontWeight:700, color:C.text }}>Projets Kits</div>
        <div style={{ flex:1 }}/>
        {(view==="detail"||view==="nouveau") && (
          <button onClick={() => { setView("liste"); setSelectedId(null); }} style={{ fontSize:12, color:C.textMuted, background:"none", border:"none", cursor:"pointer", fontFamily:"inherit" }}>
            ← Retour à la liste
          </button>
        )}
      </div>
      {dbLoading
        ? <div style={{ flex:1, display:"flex", alignItems:"center", justifyContent:"center", color:C.textMuted, gap:10, fontSize:13 }}>
            <span style={{ width:16,height:16,border:`2px solid ${C.blue}`,borderTopColor:"transparent",borderRadius:"50%",display:"inline-block",animation:"spin 0.8s linear infinite" }}/>
            Chargement des projets…
            <style>{`@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}`}</style>
          </div>
        : <>
            {view==="liste" && <ListeProjets projets={projets} onSelect={p=>{ setSelectedId(p.id); setView("detail"); }} onNew={()=>setView("nouveau")} />}
            {view==="nouveau" && <FormulaireProjet onSave={handleSave} onCancel={()=>setView("liste")} onToast={onToast} />}
            {view==="detail" && selected && <DetailProjet projet={selected} session={session} onBack={()=>setView("liste")} onUpdate={handleUpdate} onToast={onToast} />}
          </>
      }
    </div>
  );
}
