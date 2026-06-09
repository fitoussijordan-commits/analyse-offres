"use client";
import { useState, useEffect, useRef } from "react";
import * as odoo from "@/lib/odoo";

// ── Design tokens ─────────────────────────────────────────────────────────────
const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0",
  blue: "#3b82f6", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5", greenDark: "#065f46",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  red: "#ef4444", redSoft: "#fef2f2",
  purple: "#8b5cf6", purpleSoft: "#f5f3ff",
  slate: "#475569",
  shadow: "0 1px 3px rgba(0,0,0,0.06)", shadowMd: "0 4px 16px rgba(0,0,0,0.10)",
};

// ── Types ─────────────────────────────────────────────────────────────────────
interface ComposantKit { ref: string; nom?: string; qtyParKit: number; dateArrivage?: string; }
type ProjetStatus = "planning" | "en_attente" | "stock_ok" | "reserve" | "envoye";
interface ProjetKit {
  id: string; nom: string; qtyKits: number;
  composants: ComposantKit[];
  dateLancement?: string; dateEsat: string;
  notes?: string; status: ProjetStatus; createdAt: string;
}
interface StockInfo { ref: string; nom: string; dispo: number; productId: number; }
interface Props { session: odoo.OdooSession; onToast: (msg: string, type?: "success"|"error"|"info") => void; }

// ── LocalStorage ──────────────────────────────────────────────────────────────
const LS_KEY = "ao_projets_kits";
function loadProjets(): ProjetKit[] { try { const r = localStorage.getItem(LS_KEY); return r ? JSON.parse(r) : []; } catch { return []; } }
function saveProjets(p: ProjetKit[]) { localStorage.setItem(LS_KEY, JSON.stringify(p)); }
function genId() { return Math.random().toString(36).slice(2, 10); }

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmt(n: number) { return new Intl.NumberFormat("fr-FR").format(n); }
function fmtDate(iso: string) {
  if (!iso) return "—";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("fr-FR", { day: "2-digit", month: "short", year: "numeric" });
}
function daysUntil(iso: string): number {
  if (!iso) return Infinity;
  const diff = new Date(iso + "T00:00:00").getTime() - Date.now();
  return Math.ceil(diff / 86400000);
}

const STATUS_CONFIG: Record<ProjetStatus, { label: string; color: string; bg: string; }> = {
  planning:   { label: "Planification", color: C.slate,  bg: "#f1f5f9" },
  en_attente: { label: "En attente stock", color: C.amber,  bg: C.amberSoft },
  stock_ok:   { label: "⚡ Stock prêt !",  color: C.green,  bg: C.greenSoft },
  reserve:    { label: "Réservé",           color: C.blue,   bg: C.blueSoft },
  envoye:     { label: "Envoyé ESAT",       color: C.purple, bg: C.purpleSoft },
};

// ── Stock fetch ───────────────────────────────────────────────────────────────
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

// ═════════════════════════════════════════════════════════════════════════════
// FORMULAIRE PROJET
// ═════════════════════════════════════════════════════════════════════════════
function FormulaireProjet({ projet, onSave, onCancel, onToast }: {
  projet?: ProjetKit; onSave: (p: ProjetKit) => void;
  onCancel: () => void; onToast: Props["onToast"];
}) {
  const [nom, setNom] = useState(projet?.nom || "");
  const [qtyKits, setQtyKits] = useState(String(projet?.qtyKits || ""));
  const [dateLancement, setDateLancement] = useState(projet?.dateLancement || "");
  const [dateEsat, setDateEsat] = useState(projet?.dateEsat || "");
  const [notes, setNotes] = useState(projet?.notes || "");
  const [composants, setComposants] = useState<ComposantKit[]>(projet?.composants || []);
  const [newRef, setNewRef] = useState("");
  const [newQty, setNewQty] = useState("1");
  const [newDate, setNewDate] = useState("");

  const addComposant = () => {
    const ref = newRef.trim().toUpperCase();
    if (!ref) return;
    if (composants.some(c => c.ref.toLowerCase() === ref.toLowerCase())) { onToast("Référence déjà ajoutée", "error"); return; }
    const qty = parseInt(newQty) || 1;
    setComposants(prev => [...prev, { ref, qtyParKit: qty, dateArrivage: newDate || undefined }]);
    setNewRef(""); setNewQty("1"); setNewDate("");
  };

  const removeComposant = (ref: string) => setComposants(prev => prev.filter(c => c.ref !== ref));
  const updateQty = (ref: string, qty: number) => setComposants(prev => prev.map(c => c.ref === ref ? { ...c, qtyParKit: Math.max(1, qty) } : c));
  const updateDate = (ref: string, date: string) => setComposants(prev => prev.map(c => c.ref === ref ? { ...c, dateArrivage: date || undefined } : c));

  const save = () => {
    if (!nom.trim()) { onToast("Nom du projet requis", "error"); return; }
    const qty = parseInt(qtyKits);
    if (!qty || qty <= 0) { onToast("Quantité de kits invalide", "error"); return; }
    if (!composants.length) { onToast("Au moins un composant requis", "error"); return; }
    if (!dateEsat) { onToast("Date ESAT requise", "error"); return; }
    const p: ProjetKit = {
      id: projet?.id || genId(), nom: nom.trim(), qtyKits: qty,
      composants, dateLancement: dateLancement || undefined, dateEsat,
      notes: notes.trim() || undefined,
      status: projet?.status || "planning",
      createdAt: projet?.createdAt || new Date().toISOString(),
    };
    onSave(p);
  };

  const inp = { padding:"9px 12px", border:`1.5px solid ${C.border}`, borderRadius:8, fontSize:13, fontFamily:"inherit", background:C.white, color:C.text, outline:"none" };
  const lbl = { fontSize:11, fontWeight:700 as const, color:C.textMuted, textTransform:"uppercase" as const, letterSpacing:"0.07em", display:"block" as const, marginBottom:5 };

  return (
    <div style={{ flex:1, overflowY:"auto", padding:"20px 24px" }}>
      <div style={{ maxWidth:900, margin:"0 auto" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:24 }}>
          <button onClick={onCancel} style={{ width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,cursor:"pointer" }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.text} strokeWidth="2"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
          </button>
          <h2 style={{ margin:0, fontSize:18, fontWeight:700, color:C.text }}>{projet ? "Modifier le projet" : "Nouveau projet kit"}</h2>
        </div>

        <div style={{ display:"flex", flexDirection:"column", gap:18 }}>
          {/* Infos générales */}
          <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:12, padding:"18px 20px" }}>
            <div style={{ fontSize:12, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:14 }}>Informations générales</div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr auto", gap:12, alignItems:"start" }}>
              <div>
                <label style={lbl}>Nom du projet *</label>
                <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Ex: Vanity Noël 2026" style={{ ...inp, width:"100%" }} />
              </div>
              <div>
                <label style={lbl}>Quantité de kits *</label>
                <input type="number" min="1" value={qtyKits} onChange={e=>setQtyKits(e.target.value)} placeholder="500" style={{ ...inp, width:120 }} />
              </div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginTop:12 }}>
              <div>
                <label style={lbl}>Date de lancement</label>
                <input type="date" value={dateLancement} onChange={e=>setDateLancement(e.target.value)} style={{ ...inp, width:"100%" }} />
              </div>
              <div>
                <label style={lbl}>Date envoi ESAT *</label>
                <input type="date" value={dateEsat} onChange={e=>setDateEsat(e.target.value)} style={{ ...inp, width:"100%" }} />
              </div>
            </div>
            <div style={{ marginTop:12 }}>
              <label style={lbl}>Notes (optionnel)</label>
              <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Informations complémentaires…" rows={2}
                style={{ ...inp, width:"100%", resize:"vertical", fontFamily:"inherit" }} />
            </div>
          </div>

          {/* Composants */}
          <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:12, padding:"18px 20px" }}>
            <div style={{ fontSize:12, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:14 }}>Composants du kit</div>

            {/* Ajout */}
            <div style={{ display:"flex", gap:8, marginBottom:14, flexWrap:"wrap", alignItems:"center" }}>
              <input value={newRef} onChange={e=>setNewRef(e.target.value.toUpperCase())}
                onKeyDown={e=>{ if(e.key==="Enter") addComposant(); }}
                placeholder="Référence (ex: 1010214)" style={{ ...inp, flex:"1 1 160px", fontFamily:"'SF Mono','Fira Code',monospace", fontSize:12 }} />
              <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                <span style={{ fontSize:11, color:C.textMuted, whiteSpace:"nowrap" }}>Qté/kit</span>
                <input type="number" min="1" value={newQty} onChange={e=>setNewQty(e.target.value)}
                  style={{ ...inp, width:70, textAlign:"center" }} />
              </div>
              <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                <span style={{ fontSize:11, color:C.textMuted, whiteSpace:"nowrap" }}>Arrivage</span>
                <input type="date" value={newDate} onChange={e=>setNewDate(e.target.value)}
                  style={{ ...inp, fontSize:12 }} />
              </div>
              <button onClick={addComposant} style={{ padding:"9px 16px", background:C.blue, color:"#fff", border:"none", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit", whiteSpace:"nowrap" }}>
                + Ajouter
              </button>
            </div>

            {/* Liste */}
            {composants.length === 0 ? (
              <div style={{ textAlign:"center", padding:"20px 0", color:C.textMuted, fontSize:12 }}>Aucun composant ajouté</div>
            ) : (
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:C.bg }}>
                    <th style={{ padding:"7px 12px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11 }}>Référence</th>
                    <th style={{ padding:"7px 12px", textAlign:"center", fontWeight:600, color:C.textMuted, fontSize:11 }}>Qté / kit</th>
                    <th style={{ padding:"7px 12px", textAlign:"center", fontWeight:600, color:C.textMuted, fontSize:11 }}>Qté totale</th>
                    <th style={{ padding:"7px 12px", textAlign:"center", fontWeight:600, color:C.textMuted, fontSize:11 }}>Date arrivage</th>
                    <th style={{ padding:"7px 12px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {composants.map((c, i) => {
                    const qtyTotal = c.qtyParKit * (parseInt(qtyKits) || 0);
                    return (
                      <tr key={c.ref} style={{ borderTop:`1px solid ${C.border}` }}>
                        <td style={{ padding:"9px 12px", fontFamily:"'SF Mono','Fira Code',monospace", fontWeight:700, color:C.blue }}>{c.ref}</td>
                        <td style={{ padding:"9px 12px", textAlign:"center" }}>
                          <input type="number" min="1" value={c.qtyParKit}
                            onChange={e=>updateQty(c.ref, parseInt(e.target.value)||1)}
                            style={{ width:70, padding:"4px 8px", border:`1px solid ${C.border}`, borderRadius:6, fontSize:12, textAlign:"center", fontFamily:"inherit", background:C.bg }} />
                        </td>
                        <td style={{ padding:"9px 12px", textAlign:"center", fontWeight:700, color:C.text }}>{qtyTotal > 0 ? fmt(qtyTotal) : "—"}</td>
                        <td style={{ padding:"9px 12px", textAlign:"center" }}>
                          <input type="date" value={c.dateArrivage || ""}
                            onChange={e=>updateDate(c.ref, e.target.value)}
                            style={{ padding:"4px 8px", border:`1px solid ${C.border}`, borderRadius:6, fontSize:11, fontFamily:"inherit", background:C.bg, color:c.dateArrivage?C.text:C.textMuted }} />
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
            )}
          </div>

          {/* Actions */}
          <div style={{ display:"flex", gap:10 }}>
            <button onClick={onCancel} style={{ flex:1, padding:"11px 0", background:C.bg, color:C.textSec, border:`1px solid ${C.border}`, borderRadius:10, fontSize:14, fontWeight:600, cursor:"pointer", fontFamily:"inherit" }}>
              Annuler
            </button>
            <button onClick={save} style={{ flex:2, padding:"11px 0", background:C.blue, color:"#fff", border:"none", borderRadius:10, fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>
              {projet ? "Enregistrer les modifications" : "Créer le projet"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════════
// DÉTAIL PROJET
// ═════════════════════════════════════════════════════════════════════════════
function DetailProjet({ projet, session, onBack, onUpdate, onToast }: {
  projet: ProjetKit; session: odoo.OdooSession;
  onBack: () => void; onUpdate: (p: ProjetKit) => void; onToast: Props["onToast"];
}) {
  const [stock, setStock] = useState<StockInfo[]>([]);
  const [loading, setLoading] = useState(false);
  const [lastCheck, setLastCheck] = useState<Date|null>(null);
  const [editing, setEditing] = useState(false);

  useEffect(() => { verifierStock(); }, []);

  const verifierStock = async () => {
    setLoading(true);
    try {
      const refs = projet.composants.map(c => c.ref);
      const s = await fetchStock(session, refs);
      setStock(s);
      setLastCheck(new Date());
      // Vérifier si tout est dispo
      const allOk = projet.composants.every(c => {
        const info = s.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
        return info && info.dispo >= c.qtyParKit * projet.qtyKits;
      });
      if (allOk && projet.status === "en_attente") {
        const updated = { ...projet, status: "stock_ok" as ProjetStatus };
        onUpdate(updated);
        onToast("⚡ Tout le stock est disponible ! Vous pouvez réserver.", "success");
      } else if (allOk) {
        onToast("Stock vérifié — tout est disponible ✅", "success");
      } else {
        onToast("Stock vérifié — certains produits manquent encore", "info");
      }
    } catch(e:any) { onToast("Erreur stock : " + e.message, "error"); }
    finally { setLoading(false); }
  };

  const setStatus = (s: ProjetStatus) => {
    const updated = { ...projet, status: s };
    onUpdate(updated);
    onToast("Statut mis à jour", "success");
  };

  if (editing) {
    return <FormulaireProjet projet={projet} onSave={p => { onUpdate(p); setEditing(false); onToast("Projet mis à jour", "success"); }} onCancel={() => setEditing(false)} onToast={onToast} />;
  }

  const cfg = STATUS_CONFIG[projet.status];
  const daysEsat = daysUntil(projet.dateEsat);
  const allStockOk = stock.length > 0 && projet.composants.every(c => {
    const info = stock.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
    return info && info.dispo >= c.qtyParKit * projet.qtyKits;
  });

  return (
    <div style={{ flex:1, overflowY:"auto", padding:"20px 24px" }}>
      <div>
        {/* Header */}
        <div style={{ display:"flex", alignItems:"flex-start", gap:12, marginBottom:24 }}>
          <button onClick={onBack} style={{ width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8,cursor:"pointer",flexShrink:0,marginTop:3 }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.text} strokeWidth="2"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
          </button>
          <div style={{ flex:1 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
              <h2 style={{ margin:0, fontSize:20, fontWeight:800, color:C.text }}>{projet.nom}</h2>
              <span style={{ fontSize:12, fontWeight:700, color:cfg.color, background:cfg.bg, padding:"3px 10px", borderRadius:6 }}>{cfg.label}</span>
            </div>
            <div style={{ fontSize:13, color:C.textMuted, marginTop:4 }}>{fmt(projet.qtyKits)} kits · {projet.composants.length} composant(s)</div>
          </div>
          <button onClick={() => setEditing(true)} style={{ padding:"7px 14px", background:C.bg, border:`1px solid ${C.border}`, borderRadius:8, cursor:"pointer", fontSize:12, fontWeight:600, color:C.textSec, fontFamily:"inherit", display:"flex", alignItems:"center", gap:6 }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
            Modifier
          </button>
        </div>

        {/* Alerte stock OK */}
        {(projet.status === "stock_ok" || allStockOk) && projet.status !== "reserve" && projet.status !== "envoye" && (
          <div style={{ background:C.greenSoft, border:`1.5px solid ${C.green}44`, borderRadius:12, padding:"16px 20px", marginBottom:20, display:"flex", alignItems:"center", gap:16 }}>
            <div style={{ width:40,height:40,background:C.green+"22",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke={C.green} strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            </div>
            <div style={{ flex:1 }}>
              <div style={{ fontSize:14, fontWeight:700, color:C.greenDark }}>Stock disponible pour ce projet</div>
              <div style={{ fontSize:12, color:C.green, marginTop:2 }}>Tous les composants ont le stock nécessaire. Vous pouvez procéder à la réservation dans Odoo.</div>
            </div>
            <button onClick={() => setStatus("reserve")} style={{ padding:"9px 18px", background:C.green, color:"#fff", border:"none", borderRadius:9, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit", whiteSpace:"nowrap" }}>
              Marquer comme réservé
            </button>
          </div>
        )}

        {/* Dates + infos */}
        {(() => {
          const datesArr = projet.composants.filter(c=>c.dateArrivage).map(c=>c.dateArrivage!).sort();
          const lastArr = datesArr.at(-1);
          const daysLast = lastArr ? daysUntil(lastArr) : Infinity;
          return (
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(200px,1fr))", gap:12, marginBottom:20 }}>
              {[
                { label:"Qté de kits", value:fmt(projet.qtyKits), color:C.blue, icon:<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/></svg> },
                ...(projet.dateLancement ? [{ label:`Lancement${daysUntil(projet.dateLancement)<Infinity?(daysUntil(projet.dateLancement)>0?` (J-${daysUntil(projet.dateLancement)})`:" ✓"):""}`, value:fmtDate(projet.dateLancement), color:daysUntil(projet.dateLancement)<=0?C.green:daysUntil(projet.dateLancement)<=7?C.amber:C.text, icon:<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> }] : []),
                ...(lastArr ? [{ label:`Dernier arrivage${daysLast<Infinity?(daysLast>0?` (J-${daysLast})`:" ✓"):""}`, value:fmtDate(lastArr), color:daysLast<=0?C.green:daysLast<=7?C.amber:C.text, icon:<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> }] : []),
                { label:`Envoi ESAT${daysEsat<Infinity?(daysEsat>0?` (J-${daysEsat})`:" (passé)"):""}`, value:fmtDate(projet.dateEsat), color:daysEsat<=7?C.red:C.text, icon:<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg> },
              ].map(item => (
                <div key={item.label} style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:10, padding:"12px 16px", display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ color:item.color, opacity:0.7, flexShrink:0 }}>{item.icon}</div>
                  <div>
                    <div style={{ fontSize:10, fontWeight:600, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.05em" }}>{item.label}</div>
                    <div style={{ fontSize:15, fontWeight:700, color:item.color, marginTop:1 }}>{item.value}</div>
                  </div>
                </div>
              ))}
            </div>
          );
        })()}

        {/* Composants + stock */}
        <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:12, overflow:"hidden", marginBottom:16 }}>
          <div style={{ padding:"14px 20px", borderBottom:`1px solid ${C.border}`, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
            <div style={{ fontSize:13, fontWeight:700, color:C.text }}>Composants du kit</div>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              {lastCheck && <span style={{ fontSize:11, color:C.textMuted }}>Vérifié {lastCheck.toLocaleTimeString("fr-FR", { hour:"2-digit", minute:"2-digit" })}</span>}
              <button onClick={verifierStock} disabled={loading}
                style={{ padding:"7px 14px", background:loading?C.bg:C.blue+"18", border:`1px solid ${loading?C.border:C.blue+"44"}`, borderRadius:8, cursor:loading?"default":"pointer", fontSize:12, fontWeight:600, color:loading?C.textMuted:C.blue, fontFamily:"inherit", display:"flex", alignItems:"center", gap:6 }}>
                {loading
                  ? <><span style={{ width:12,height:12,border:`2px solid ${C.blue}`,borderTopColor:"transparent",borderRadius:"50%",display:"inline-block",animation:"spin 0.8s linear infinite" }}/>Vérification…</>
                  : <><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>Vérifier le stock</>}
              </button>
            </div>
          </div>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
            <thead>
              <tr style={{ background:C.bg }}>
                <th style={{ padding:"10px 16px", width:52 }}></th>
                <th style={{ padding:"10px 16px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11, width:110 }}>Référence</th>
                <th style={{ padding:"10px 16px", textAlign:"left", fontWeight:600, color:C.textMuted, fontSize:11 }}>Désignation</th>
                <th style={{ padding:"10px 16px", textAlign:"center", fontWeight:600, color:C.textMuted, fontSize:11, width:130 }}>Arrivage prévu</th>
                <th style={{ padding:"10px 16px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11, width:80 }}>Qté/kit</th>
                <th style={{ padding:"10px 16px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11, width:110 }}>Besoin total</th>
                <th style={{ padding:"10px 16px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11, width:110 }}>Dispo Odoo</th>
                <th style={{ padding:"10px 16px", textAlign:"right", fontWeight:600, color:C.textMuted, fontSize:11, width:120 }}>Écart</th>
              </tr>
            </thead>
            <tbody>
              {projet.composants.map((c, i) => {
                const qtyTotale = c.qtyParKit * projet.qtyKits;
                const info = stock.find(x => x.ref.toLowerCase() === c.ref.toLowerCase());
                const ok = info && info.dispo >= qtyTotale;
                const partiel = info && info.dispo > 0 && info.dispo < qtyTotale;
                const manquant = info && info.dispo < qtyTotale;
                return (
                  <tr key={c.ref} style={{ borderTop:`1px solid ${C.border}`, background:i%2===0?C.white:C.bg+"80" }}>
                    <td style={{ padding:"11px 20px", fontFamily:"'SF Mono','Fira Code',monospace", fontWeight:700, color:C.blue }}>{c.ref}</td>
                    <td style={{ padding:"11px 20px", color:C.textSec }}>{info?.nom || c.nom || <span style={{ color:C.textMuted, fontStyle:"italic" }}>—</span>}</td>
                    <td style={{ padding:"11px 20px", textAlign:"right", color:C.textSec }}>{fmt(c.qtyParKit)}</td>
                    <td style={{ padding:"11px 20px", textAlign:"right", fontWeight:700, color:C.text }}>{fmt(qtyTotale)}</td>
                    {stock.length > 0 && <>
                      <td style={{ padding:"11px 16px", textAlign:"right", fontWeight:700, color:ok?C.green:partiel?C.amber:C.red }}>
                        {info ? fmt(info.dispo) : <span style={{ color:C.red, fontSize:12 }}>Introuvable</span>}
                      </td>
                      <td style={{ padding:"11px 16px", textAlign:"center" }}>
                        {!info
                          ? <span style={{ fontSize:11, color:C.red, background:C.redSoft, padding:"3px 8px", borderRadius:5, fontWeight:600 }}>Réf. inconnue</span>
                          : ok
                          ? <span style={{ fontSize:11, color:C.green, background:C.greenSoft, padding:"3px 8px", borderRadius:5, fontWeight:600 }}>✓ OK</span>
                          : partiel
                          ? <span style={{ fontSize:11, color:C.amber, background:C.amberSoft, padding:"3px 8px", borderRadius:5, fontWeight:600 }}>⚠ {fmt(qtyTotale - info!.dispo)} manquants</span>
                          : <span style={{ fontSize:11, color:C.red, background:C.redSoft, padding:"3px 8px", borderRadius:5, fontWeight:600 }}>✗ {fmt(qtyTotale - (info?.dispo||0))} manquants</span>
                        }
                      </td>
                    </>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Progression statut */}
        <div style={{ background:C.white, border:`1px solid ${C.border}`, borderRadius:12, padding:"16px 20px", marginBottom:16 }}>
          <div style={{ fontSize:12, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:14 }}>Progression du projet</div>
          <div style={{ display:"flex", gap:6, alignItems:"center" }}>
            {(["planning","en_attente","stock_ok","reserve","envoye"] as ProjetStatus[]).map((s, i, arr) => {
              const scfg = STATUS_CONFIG[s];
              const isActive = projet.status === s;
              const isPast = arr.indexOf(projet.status) > i;
              return (
                <div key={s} style={{ display:"flex", alignItems:"center", flex: i < arr.length-1 ? 1 : "none" }}>
                  <button onClick={() => setStatus(s)}
                    style={{ padding:"6px 12px", background:isActive?scfg.color:isPast?scfg.bg:"transparent", color:isActive?"#fff":isPast?scfg.color:C.textMuted, border:`1.5px solid ${isActive?scfg.color:isPast?scfg.color+"44":C.border}`, borderRadius:7, cursor:"pointer", fontSize:11, fontWeight:600, fontFamily:"inherit", whiteSpace:"nowrap" }}>
                    {scfg.label}
                  </button>
                  {i < arr.length - 1 && <div style={{ flex:1, height:1, background:isPast?scfg.color+"44":C.border, margin:"0 4px" }}/>}
                </div>
              );
            })}
          </div>
        </div>

        {/* Notes */}
        {projet.notes && (
          <div style={{ background:C.amberSoft, border:`1px solid ${C.amber}33`, borderRadius:10, padding:"12px 16px", fontSize:13, color:C.textSec }}>
            <span style={{ fontWeight:600, color:C.amber }}>Note : </span>{projet.notes}
          </div>
        )}
      </div>
      <style>{`@keyframes spin { from { transform:rotate(0deg); } to { transform:rotate(360deg); } } @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }`}</style>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════════
// LISTE DES PROJETS
// ═════════════════════════════════════════════════════════════════════════════
function ListeProjets({ projets, onSelect, onNew }: { projets: ProjetKit[]; onSelect: (p:ProjetKit) => void; onNew: () => void; }) {
  const grouped = {
    actif: projets.filter(p => !["envoye"].includes(p.status)),
    envoye: projets.filter(p => p.status === "envoye"),
  };

  const CardProjet = ({ p }: { p: ProjetKit }) => {
    const cfg = STATUS_CONFIG[p.status];
    const daysE = daysUntil(p.dateEsat);
    return (
      <div onClick={() => onSelect(p)} style={{ background:C.white, border:`1.5px solid ${p.status==="stock_ok"?C.green+"44":C.border}`, borderRadius:12, padding:"16px 18px", cursor:"pointer", boxShadow:p.status==="stock_ok"?`0 4px 20px ${C.green}18`:C.shadow, transition:"all 0.1s" }}>
        <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:10, marginBottom:10 }}>
          <div>
            <div style={{ fontSize:14, fontWeight:700, color:C.text }}>{p.nom}</div>
            <div style={{ fontSize:12, color:C.textMuted, marginTop:2 }}>{fmt(p.qtyKits)} kits · {p.composants.length} composants</div>
          </div>
          <span style={{ fontSize:11, fontWeight:700, color:cfg.color, background:cfg.bg, padding:"3px 9px", borderRadius:5, whiteSpace:"nowrap", flexShrink:0 }}>{cfg.label}</span>
        </div>
        {(() => {
          const datesArrivage = p.composants.filter(c=>c.dateArrivage).map(c=>c.dateArrivage!).sort();
          const lastArrivage = datesArrivage.at(-1);
          const daysLast = lastArrivage ? daysUntil(lastArrivage) : Infinity;
          return (
            <div style={{ display:"flex", gap:12, fontSize:11, color:C.textMuted, flexWrap:"wrap" }}>
              {lastArrivage && <span style={{ color:daysLast<=7&&daysLast>0?C.amber:daysLast<=0?C.green:C.textMuted }}>
                📦 Dernier arrivage : {fmtDate(lastArrivage)}{daysLast>0&&daysLast<=30?` (J-${daysLast})`:daysLast<=0?" ✓":""}
              </span>}
              <span style={{ color:daysE<=7&&daysE>0?C.amber:daysE<=0?C.red:C.textMuted }}>
                🏭 ESAT : {fmtDate(p.dateEsat)}{daysE>0&&daysE<=30?` (J-${daysE})`:""}
              </span>
            </div>
          );
        })()}
        {p.status === "stock_ok" && (
          <div style={{ marginTop:10, padding:"7px 12px", background:C.greenSoft, borderRadius:7, fontSize:12, fontWeight:600, color:C.green, display:"flex", alignItems:"center", gap:6 }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><polyline points="20 6 9 17 4 12"/></svg>
            Stock disponible — réservation possible
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={{ flex:1, overflowY:"auto", padding:"20px 24px" }}>
      <div>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:24 }}>
          <div>
            <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:C.text }}>Projets Kits</h2>
            <p style={{ margin:"4px 0 0", fontSize:13, color:C.textMuted }}>{projets.length} projet{projets.length!==1?"s":""}</p>
          </div>
          <button onClick={onNew} style={{ padding:"9px 18px", background:C.blue, color:"#fff", border:"none", borderRadius:9, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit", display:"flex", alignItems:"center", gap:7 }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
            Nouveau projet
          </button>
        </div>

        {projets.length === 0 ? (
          <div style={{ textAlign:"center", padding:"80px 24px" }}>
            <div style={{ width:64,height:64,borderRadius:16,background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px" }}>
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={C.textMuted} strokeWidth="1.5"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>
            </div>
            <div style={{ fontSize:15, fontWeight:600, color:C.textSec, marginBottom:6 }}>Aucun projet kit</div>
            <div style={{ fontSize:13, color:C.textMuted, marginBottom:20 }}>Créez votre premier projet pour suivre vos kits et leur stock</div>
            <button onClick={onNew} style={{ padding:"10px 22px", background:C.blue, color:"#fff", border:"none", borderRadius:9, cursor:"pointer", fontSize:13, fontWeight:700, fontFamily:"inherit" }}>
              Créer un projet
            </button>
          </div>
        ) : (
          <>
            {grouped.actif.length > 0 && (
              <div style={{ marginBottom:24 }}>
                <div style={{ fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:10 }}>Projets en cours</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(380px, 1fr))", gap:12 }}>
                  {grouped.actif.map(p => <CardProjet key={p.id} p={p} />)}
                </div>
              </div>
            )}
            {grouped.envoye.length > 0 && (
              <div>
                <div style={{ fontSize:11, fontWeight:700, color:C.textMuted, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:10 }}>Envoyés à l'ESAT</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(380px, 1fr))", gap:12, opacity:0.65 }}>
                  {grouped.envoye.map(p => <CardProjet key={p.id} p={p} />)}
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ═════════════════════════════════════════════════════════════════════════════
// COMPOSANT PRINCIPAL
// ═════════════════════════════════════════════════════════════════════════════
export default function ProjetScreen({ session, onToast }: Props) {
  const [projets, setProjets] = useState<ProjetKit[]>([]);
  const [view, setView] = useState<"liste"|"detail"|"nouveau">("liste");
  const [selectedId, setSelectedId] = useState<string|null>(null);

  useEffect(() => { setProjets(loadProjets()); }, []);

  const selected = projets.find(p => p.id === selectedId) || null;

  const handleSave = (p: ProjetKit) => {
    const updated = projets.find(x => x.id === p.id)
      ? projets.map(x => x.id === p.id ? p : x)
      : [...projets, p];
    saveProjets(updated); setProjets(updated);
    setSelectedId(p.id); setView("detail");
    onToast("Projet enregistré", "success");
  };

  const handleUpdate = (p: ProjetKit) => {
    const updated = projets.map(x => x.id === p.id ? p : x);
    saveProjets(updated); setProjets(updated);
  };

  return (
    <div style={{ flex:1, display:"flex", flexDirection:"column", height:"100%", overflow:"hidden", background:C.bg }}>
      {/* Topbar */}
      <div style={{ background:C.white, borderBottom:`1px solid ${C.border}`, padding:"0 24px", display:"flex", alignItems:"center", height:56, flexShrink:0 }}>
        <div style={{ fontSize:15, fontWeight:700, color:C.text }}>Projets Kits</div>
        <div style={{ flex:1 }}/>
        {(view === "detail" || view === "nouveau") && (
          <button onClick={() => { setView("liste"); setSelectedId(null); }} style={{ fontSize:12, color:C.textMuted, background:"none", border:"none", cursor:"pointer", fontFamily:"inherit" }}>
            ← Retour à la liste
          </button>
        )}
      </div>

      {view === "liste" && <ListeProjets projets={projets} onSelect={p=>{ setSelectedId(p.id); setView("detail"); }} onNew={()=>setView("nouveau")} />}
      {view === "nouveau" && <FormulaireProjet onSave={handleSave} onCancel={()=>setView("liste")} onToast={onToast} />}
      {view === "detail" && selected && <DetailProjet projet={selected} session={session} onBack={()=>setView("liste")} onUpdate={handleUpdate} onToast={onToast} />}
    </div>
  );
}
