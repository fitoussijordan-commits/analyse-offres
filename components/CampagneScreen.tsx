"use client";
import { useState, useEffect, useRef } from "react";
import * as odoo from "@/lib/odoo";
import * as cp from "@/lib/campaigns";
import { fetchCampaign, type CampaignResult, type StateFilter } from "@/lib/analyse-campaign";

const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0",
  blue: "#3b82f6", blueDark: "#1d4ed8", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  red: "#ef4444", redSoft: "#fef2f2",
  purple: "#8b5cf6", purpleSoft: "#f5f3ff",
  teal: "#0d9488", tealSoft: "#f0fdfa",
  shadow: "0 1px 3px rgba(0,0,0,0.06)", shadowMd: "0 4px 16px rgba(0,0,0,0.10)",
};
const fmtEur = (n: number) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n || 0);
const fmtNum = (n: number) => new Intl.NumberFormat("fr-FR").format(n || 0);

function extractRefs(text: string): string[] {
  const numRefs = text.match(/\d{5,}/g);
  if (numRefs && numRefs.length >= 2) return [...new Set(numRefs)];
  return text.split(/[\n\r,;]+/).map(r => r.trim()).filter(Boolean);
}

interface Props { session: odoo.OdooSession; onToast: (msg: string, type?: "success" | "error" | "info") => void; }

const inputStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", padding: "9px 12px", border: `1.5px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontFamily: "inherit", background: C.white, color: C.text, outline: "none" };
const labelStyle: React.CSSProperties = { fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.07em", display: "block", marginBottom: 5 };

// ════════════════════════════════════════════════════════════════════════════
// Panneau : gestion des OFFRES
// ════════════════════════════════════════════════════════════════════════════
function OffrePanel({ onClose, onToast, onChanged }: { onClose: () => void; onToast: Props["onToast"]; onChanged: () => void }) {
  const [offres, setOffres] = useState<cp.Offre[]>([]);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [code, setCode] = useState(""); const [label, setLabel] = useState(""); const [produits, setProduits] = useState(""); const [codeInterne, setCodeInterne] = useState("");

  const reload = () => cp.loadOffres().then(setOffres).catch(e => onToast("Erreur : " + e.message, "error"));
  useEffect(() => { reload(); }, []);

  const openNew = () => { setEditId(null); setCode(""); setLabel(""); setProduits(""); setCodeInterne(""); setShowForm(true); };
  const openEdit = (o: cp.Offre) => { setEditId(o.id); setCode(o.code); setLabel(o.label); setProduits(o.produits.join("\n")); setCodeInterne(o.codeInterne || ""); setShowForm(true); };

  const save = async () => {
    const c = code.trim(); if (!c) { onToast("Code offre requis", "error"); return; }
    const refs = produits.split(/[\n\r,;]+/).map(r => r.trim()).filter(Boolean);
    const o: cp.Offre = { id: editId || cp.genId(), code: c, label: label.trim(), produits: refs, codeInterne: codeInterne.trim() || undefined };
    try { await cp.upsertOffre(o); setShowForm(false); reload(); onChanged(); onToast("Offre enregistrée", "success"); }
    catch (e: any) { onToast("Erreur : " + e.message, "error"); }
  };
  const remove = async (id: string) => { if (!confirm("Supprimer cette offre ?")) return; try { await cp.deleteOffre(id); reload(); onChanged(); onToast("Offre supprimée", "info"); } catch (e: any) { onToast("Erreur : " + e.message, "error"); } };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 200, display: "flex" }}>
      <div style={{ flex: 1, background: "rgba(0,0,0,0.4)" }} onClick={onClose} />
      <div style={{ width: 480, background: C.white, display: "flex", flexDirection: "column", boxShadow: "-8px 0 40px rgba(0,0,0,0.15)" }}>
        <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>{showForm ? (editId ? "Modifier l'offre" : "Nouvelle offre") : "Gestion des offres"}</div>
          <button onClick={showForm ? () => setShowForm(false) : onClose} style={{ width: 32, height: 32, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 16 }}>{showForm ? "←" : "✕"}</button>
        </div>
        <div style={{ flex: 1, overflowY: "auto", padding: "20px 24px" }}>
          {showForm ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div><label style={labelStyle}>Code offre *</label><input value={code} onChange={e => setCode(e.target.value)} placeholder="Ex: 7131482" style={inputStyle} /></div>
              <div><label style={labelStyle}>Libellé</label><input value={label} onChange={e => setLabel(e.target.value)} placeholder="Ex: Offre Noël" style={inputStyle} /></div>
              <div><label style={labelStyle}>Note interne (x_note_interne)</label><input value={codeInterne} onChange={e => setCodeInterne(e.target.value)} placeholder="Ex: NOEL26" style={inputStyle} /></div>
              <div>
                <label style={labelStyle}>Références produits — une par ligne</label>
                <textarea value={produits} onChange={e => setProduits(e.target.value)} onPaste={e => { const refs = extractRefs(e.clipboardData.getData("text")); if (refs.length >= 2) { e.preventDefault(); setProduits(refs.join("\n")); } }} rows={6} placeholder={"1010214\n1010302"} style={{ ...inputStyle, fontFamily: "monospace", fontSize: 12, resize: "vertical" }} />
                <div style={{ fontSize: 11, color: C.textMuted, marginTop: 4 }}>{produits.split(/[\n\r,;]+/).filter(r => r.trim()).length} référence(s)</div>
              </div>
              <button onClick={save} style={{ padding: "11px 0", background: C.blue, color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{editId ? "Enregistrer" : "Créer l'offre"}</button>
            </div>
          ) : (
            <>
              <button onClick={openNew} style={{ width: "100%", padding: "10px 0", background: C.blue, color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginBottom: 16 }}>+ Nouvelle offre</button>
              {offres.length === 0 ? <div style={{ textAlign: "center", padding: 32, color: C.textMuted, fontSize: 13 }}>Aucune offre configurée</div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {offres.map(o => (
                    <div key={o.id} style={{ border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" }}>
                            <span style={{ fontSize: 13, fontWeight: 700, color: C.text, fontFamily: "monospace" }}>{o.code}</span>
                            {o.label && <span style={{ fontSize: 12, color: C.textMuted }}>{o.label}</span>}
                            {o.codeInterne && <span style={{ fontSize: 11, background: C.purpleSoft, color: C.purple, borderRadius: 5, padding: "2px 7px", fontWeight: 600 }}>{o.codeInterne}</span>}
                          </div>
                          <div style={{ fontSize: 11, color: C.textMuted, marginTop: 5 }}>{o.produits.length} produit(s)</div>
                        </div>
                        <div style={{ display: "flex", gap: 5 }}>
                          <button onClick={() => openEdit(o)} style={{ padding: "5px 9px", background: C.bg, border: `1px solid ${C.border}`, borderRadius: 7, cursor: "pointer", fontSize: 11 }}>✏️</button>
                          <button onClick={() => remove(o.id)} style={{ padding: "5px 9px", background: C.redSoft, border: `1px solid ${C.red}22`, borderRadius: 7, cursor: "pointer", fontSize: 11, color: C.red }}>🗑</button>
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

// ════════════════════════════════════════════════════════════════════════════
// Panneau : gestion des CAMPAGNES
// ════════════════════════════════════════════════════════════════════════════
function CampagnePanel({ onClose, onToast, offres, onChanged }: { onClose: () => void; onToast: Props["onToast"]; offres: cp.Offre[]; onChanged: () => void }) {
  const [campagnes, setCampagnes] = useState<cp.Campagne[]>([]);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [nom, setNom] = useState(""); const [selOffres, setSelOffres] = useState<string[]>([]); const [produits, setProduits] = useState(""); const [notes, setNotes] = useState("");

  const reload = () => cp.loadCampagnes().then(setCampagnes).catch(e => onToast("Erreur : " + e.message, "error"));
  useEffect(() => { reload(); }, []);

  const openNew = () => { setEditId(null); setNom(""); setSelOffres([]); setProduits(""); setNotes(""); setShowForm(true); };
  const openEdit = (c: cp.Campagne) => { setEditId(c.id); setNom(c.nom); setSelOffres(c.offres); setProduits(c.produits.join("\n")); setNotes(c.notes.join("\n")); setShowForm(true); };
  const toggleOffre = (code: string) => setSelOffres(p => p.includes(code) ? p.filter(c => c !== code) : [...p, code]);

  const save = async () => {
    const n = nom.trim(); if (!n) { onToast("Nom de campagne requis", "error"); return; }
    const refs = produits.split(/[\n\r,;]+/).map(r => r.trim()).filter(Boolean);
    const nts = notes.split(/[\n\r]+/).map(r => r.trim()).filter(Boolean);
    if (!selOffres.length && !refs.length && !nts.length) { onToast("Ajoute au moins une offre, un produit ou une note", "error"); return; }
    const c: cp.Campagne = { id: editId || cp.genId(), nom: n, offres: selOffres, produits: refs, notes: nts };
    try { await cp.upsertCampagne(c); setShowForm(false); reload(); onChanged(); onToast("Campagne enregistrée", "success"); }
    catch (e: any) { onToast("Erreur : " + e.message, "error"); }
  };
  const remove = async (id: string) => { if (!confirm("Supprimer cette campagne ?")) return; try { await cp.deleteCampagne(id); reload(); onChanged(); onToast("Campagne supprimée", "info"); } catch (e: any) { onToast("Erreur : " + e.message, "error"); } };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 200, display: "flex" }}>
      <div style={{ flex: 1, background: "rgba(0,0,0,0.4)" }} onClick={onClose} />
      <div style={{ width: 520, background: C.white, display: "flex", flexDirection: "column", boxShadow: "-8px 0 40px rgba(0,0,0,0.15)" }}>
        <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.text }}>{showForm ? (editId ? "Modifier la campagne" : "Nouvelle campagne") : "Gestion des campagnes"}</div>
          <button onClick={showForm ? () => setShowForm(false) : onClose} style={{ width: 32, height: 32, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 16 }}>{showForm ? "←" : "✕"}</button>
        </div>
        <div style={{ flex: 1, overflowY: "auto", padding: "20px 24px" }}>
          {showForm ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div><label style={labelStyle}>Nom de la campagne *</label><input value={nom} onChange={e => setNom(e.target.value)} placeholder="Ex: Campagne Noël 2026" style={inputStyle} /></div>
              <div>
                <label style={labelStyle}>Offres incluses ({selOffres.length})</label>
                {offres.length === 0 ? <div style={{ fontSize: 12, color: C.textMuted }}>Aucune offre — créez-en via « Gérer les offres ».</div> : (
                  <div style={{ display: "flex", flexDirection: "column", gap: 4, maxHeight: 200, overflowY: "auto", border: `1px solid ${C.border}`, borderRadius: 8, padding: 6 }}>
                    {offres.map(o => {
                      const checked = selOffres.includes(o.code);
                      return (
                        <div key={o.id} onClick={() => toggleOffre(o.code)} style={{ display: "flex", alignItems: "center", gap: 9, padding: "7px 9px", borderRadius: 6, cursor: "pointer", background: checked ? C.blueSoft : "transparent" }}>
                          <div style={{ width: 16, height: 16, borderRadius: 4, border: `2px solid ${checked ? C.blue : C.border}`, background: checked ? C.blue : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, color: "#fff", fontSize: 11 }}>{checked && "✓"}</div>
                          <span style={{ fontSize: 12, fontWeight: 700, fontFamily: "monospace", color: C.text }}>{o.code}</span>
                          {o.label && <span style={{ fontSize: 12, color: C.textMuted }}>{o.label}</span>}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
              <div>
                <label style={labelStyle}>Références produits autonomes — une par ligne</label>
                <textarea value={produits} onChange={e => setProduits(e.target.value)} onPaste={e => { const refs = extractRefs(e.clipboardData.getData("text")); if (refs.length >= 2) { e.preventDefault(); setProduits(refs.join("\n")); } }} rows={4} placeholder={"Coffrets saisis sans offre…\n1099001"} style={{ ...inputStyle, fontFamily: "monospace", fontSize: 12, resize: "vertical" }} />
              </div>
              <div>
                <label style={labelStyle}>Notes internes — une par ligne</label>
                <textarea value={notes} onChange={e => setNotes(e.target.value)} rows={3} placeholder={"Noel\nNoël 2026"} style={{ ...inputStyle, fontSize: 12, resize: "vertical" }} />
              </div>
              <button onClick={save} style={{ padding: "11px 0", background: C.teal, color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{editId ? "Enregistrer" : "Créer la campagne"}</button>
            </div>
          ) : (
            <>
              <button onClick={openNew} style={{ width: "100%", padding: "10px 0", background: C.teal, color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginBottom: 16 }}>+ Nouvelle campagne</button>
              {campagnes.length === 0 ? <div style={{ textAlign: "center", padding: 32, color: C.textMuted, fontSize: 13 }}>Aucune campagne</div> : (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {campagnes.map(c => (
                    <div key={c.id} style={{ border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 14px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{c.nom}</div>
                          <div style={{ fontSize: 11, color: C.textMuted, marginTop: 4 }}>{c.offres.length} offre(s) · {c.produits.length} produit(s) · {c.notes.length} note(s)</div>
                        </div>
                        <div style={{ display: "flex", gap: 5 }}>
                          <button onClick={() => openEdit(c)} style={{ padding: "5px 9px", background: C.bg, border: `1px solid ${C.border}`, borderRadius: 7, cursor: "pointer", fontSize: 11 }}>✏️</button>
                          <button onClick={() => remove(c.id)} style={{ padding: "5px 9px", background: C.redSoft, border: `1px solid ${C.red}22`, borderRadius: 7, cursor: "pointer", fontSize: 11, color: C.red }}>🗑</button>
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

// ════════════════════════════════════════════════════════════════════════════
// ÉCRAN PRINCIPAL
// ════════════════════════════════════════════════════════════════════════════
type Tab = "produits" | "delegues" | "categories" | "adherents" | "offres";
const FILTERS: [StateFilter, string][] = [["all", "Tout"], ["avenir", "À venir"], ["valide", "Validé"]];

export default function CampagneScreen({ session, onToast }: Props) {
  const [offres, setOffres] = useState<cp.Offre[]>([]);
  const [campagnes, setCampagnes] = useState<cp.Campagne[]>([]);
  const [selId, setSelId] = useState<string>("");
  const [filter, setFilter] = useState<StateFilter>("all");
  const [result, setResult] = useState<CampaignResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [tab, setTab] = useState<Tab>("produits");
  const [showCampPanel, setShowCampPanel] = useState(false);
  const [showOffrePanel, setShowOffrePanel] = useState(false);
  const [ddOpen, setDdOpen] = useState(false);
  const ddRef = useRef<HTMLDivElement>(null);

  const reloadAll = () => {
    cp.loadOffres().then(setOffres).catch(() => {});
    cp.loadCampagnes().then(setCampagnes).catch(() => {});
  };
  useEffect(() => { reloadAll(); }, []);
  useEffect(() => {
    if (!ddOpen) return;
    const h = (e: MouseEvent) => { if (ddRef.current && !ddRef.current.contains(e.target as Node)) setDdOpen(false); };
    document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h);
  }, [ddOpen]);

  const selected = campagnes.find(c => c.id === selId) || null;

  const analyser = async () => {
    if (!selected) { onToast("Sélectionne une campagne", "error"); return; }
    setLoading(true); setResult(null);
    try {
      const res = await fetchCampaign(session, selected, offres, filter);
      setResult(res); setTab("produits");
      onToast(`Analyse terminée : ${fmtEur(res.caTotal)}`, "success");
    } catch (e: any) { onToast("Erreur analyse : " + e.message, "error"); }
    finally { setLoading(false); }
  };

  // relance auto si on change de filtre alors qu'un résultat existe
  useEffect(() => { if (result && selected) analyser(); /* eslint-disable-next-line */ }, [filter]);

  const exportExcel = async () => {
    if (!result) return;
    setExporting(true);
    try {
      const filterLabel = FILTERS.find(f => f[0] === filter)?.[1];
      const res = await fetch("/api/export-campaign", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ...result, filterLabel }) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `campagne_${result.nom.replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Export téléchargé", "success");
    } catch (e: any) { onToast("Erreur export : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  const kpi = (label: string, value: string, color: string) => (
    <div style={{ flex: 1, minWidth: 150, background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "14px 16px", boxShadow: C.shadow }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>{label}</div>
      <div style={{ fontSize: 22, fontWeight: 800, color, marginTop: 4 }}>{value}</div>
    </div>
  );

  const TABS: [Tab, string][] = [["produits", "Produits"], ["delegues", "Délégués"], ["categories", "Catégorie statistique"], ["adherents", "Adhérent réseau"], ["offres", "Par offre"]];

  return (
    <div style={{ flex: 1, display: "flex", flexDirection: "column", height: "100%", overflow: "hidden", background: C.bg }}>
      {/* Topbar */}
      <div style={{ background: C.white, borderBottom: `1px solid ${C.border}`, padding: "0 24px", display: "flex", alignItems: "center", gap: 14, height: 56, flexShrink: 0 }}>
        <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>Analyse des Campagnes</div>
        <div style={{ flex: 1 }} />
        <div style={{ display: "flex", gap: 2, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 9, padding: 3 }}>
          {FILTERS.map(([key, label]) => (
            <button key={key} onClick={() => setFilter(key)} style={{ padding: "5px 14px", border: "none", borderRadius: 7, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600, background: filter === key ? C.white : "transparent", color: filter === key ? (key === "valide" ? C.green : key === "avenir" ? C.amber : C.blue) : C.textMuted, boxShadow: filter === key ? C.shadow : "none" }}>{label}</button>
          ))}
        </div>
        <button onClick={() => setShowOffrePanel(true)} style={{ padding: "7px 12px", background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.textSec, fontFamily: "inherit" }}>Gérer les offres</button>
        <button onClick={() => setShowCampPanel(true)} style={{ padding: "7px 12px", background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.textSec, fontFamily: "inherit" }}>Gérer les campagnes</button>
      </div>

      {/* Barre de sélection */}
      <div style={{ padding: "16px 24px", display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", borderBottom: `1px solid ${C.border}`, background: C.white }}>
        <div ref={ddRef} style={{ position: "relative", flex: "1 1 280px", maxWidth: 360 }}>
          <button onClick={() => setDdOpen(o => !o)} disabled={!campagnes.length} style={{ width: "100%", padding: "9px 36px 9px 12px", border: `1.5px solid ${ddOpen ? C.teal : C.border}`, borderRadius: 9, fontSize: 13, fontFamily: "inherit", background: C.white, color: selected ? C.text : C.textMuted, cursor: campagnes.length ? "pointer" : "default", textAlign: "left" }}>
            {campagnes.length === 0 ? "— Aucune campagne —" : selected ? selected.nom : "Sélectionner une campagne…"}
          </button>
          <span style={{ position: "absolute", right: 12, top: "50%", transform: `translateY(-50%) rotate(${ddOpen ? 180 : 0}deg)`, color: C.textMuted, pointerEvents: "none" }}>▾</span>
          {ddOpen && campagnes.length > 0 && (
            <div style={{ position: "absolute", top: "calc(100% + 6px)", left: 0, right: 0, background: C.white, border: `1.5px solid ${C.teal}44`, borderRadius: 10, boxShadow: C.shadowMd, zIndex: 50, maxHeight: 280, overflowY: "auto" }}>
              {campagnes.map(c => (
                <div key={c.id} onClick={() => { setSelId(c.id); setDdOpen(false); }} style={{ padding: "10px 14px", cursor: "pointer", borderBottom: `1px solid ${C.border}`, background: c.id === selId ? C.tealSoft : "transparent" }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{c.nom}</div>
                  <div style={{ fontSize: 11, color: C.textMuted, marginTop: 2 }}>{c.offres.length} offre(s) · {c.produits.length} produit(s) · {c.notes.length} note(s)</div>
                </div>
              ))}
            </div>
          )}
        </div>
        <button onClick={analyser} disabled={!selected || loading} style={{ padding: "9px 20px", background: !selected || loading ? C.border : C.teal, color: !selected || loading ? C.textMuted : "#fff", border: "none", borderRadius: 9, fontSize: 13, fontWeight: 700, cursor: !selected || loading ? "default" : "pointer", fontFamily: "inherit" }}>{loading ? "Analyse…" : "Analyser"}</button>
        {result && <button onClick={exportExcel} disabled={exporting} style={{ padding: "9px 16px", background: C.greenSoft, border: `1px solid ${C.green}44`, borderRadius: 9, fontSize: 13, fontWeight: 600, color: C.green, cursor: exporting ? "default" : "pointer", fontFamily: "inherit" }}>{exporting ? "Export…" : "⬇ Export Excel"}</button>}
      </div>

      {/* Contenu */}
      <div style={{ flex: 1, overflowY: "auto", padding: "20px 24px" }}>
        {loading ? (
          <div style={{ textAlign: "center", padding: 60, color: C.textMuted }}>
            <div style={{ width: 24, height: 24, border: `3px solid ${C.teal}`, borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 14px" }} />
            Analyse de la campagne en cours…
            <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
          </div>
        ) : !result ? (
          <div style={{ textAlign: "center", padding: "60px 20px", color: C.textMuted }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📊</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: C.text, marginBottom: 6 }}>Sélectionne une campagne et lance l'analyse</div>
            <div style={{ fontSize: 13 }}>Les ventes des offres, produits et notes de la campagne seront agrégées sans doublons.</div>
          </div>
        ) : (
          <div style={{ maxWidth: 1100, margin: "0 auto" }}>
            {/* KPIs */}
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 18 }}>
              {kpi("CA total", fmtEur(result.caTotal), C.teal)}
              {kpi("Quantité", fmtNum(result.qtyTotal), C.blue)}
              {kpi("Commandes", fmtNum(result.nbCommandes), C.purple)}
              {result.split && filter === "all" && kpi("CA validé", fmtEur(result.split.valide.ca), C.green)}
              {result.split && filter === "all" && kpi("CA à venir", fmtEur(result.split.avenir.ca), C.amber)}
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 4, marginBottom: 12, flexWrap: "wrap" }}>
              {TABS.map(([k, l]) => (
                <button key={k} onClick={() => setTab(k)} style={{ padding: "7px 14px", border: `1px solid ${tab === k ? C.teal : C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, background: tab === k ? C.tealSoft : C.white, color: tab === k ? C.teal : C.textSec, fontFamily: "inherit" }}>{l}</button>
              ))}
            </div>

            {/* Tables */}
            <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
              {tab === "produits" && <Tbl head={["Réf", "Produit", "Qté", "CA"]} aligns={["left", "left", "right", "right"]} rows={result.produits.map(p => [p.ref, p.name, fmtNum(p.qtyVendue), fmtEur(p.ca)])} />}
              {tab === "delegues" && <Tbl head={["Délégué", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right"]} rows={result.delegues.map(d => [d.name, fmtNum(d.qtyVendue), fmtEur(d.ca), result.caTotal ? ((d.ca / result.caTotal) * 100).toFixed(1) + " %" : "—"])} />}
              {tab === "categories" && <Tbl head={["Catégorie statistique", "Cmd", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right", "right"]} rows={result.categories.map(c => [c.name, fmtNum(c.nbCommandes), fmtNum(c.qtyVendue), fmtEur(c.ca), result.caTotal ? ((c.ca / result.caTotal) * 100).toFixed(1) + " %" : "—"])} />}
              {tab === "adherents" && <Tbl head={["Adhérent réseau", "Cmd", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right", "right"]} rows={result.adherents.map(a => [a.name, fmtNum(a.nbCommandes), fmtNum(a.qtyVendue), fmtEur(a.ca), result.caTotal ? ((a.ca / result.caTotal) * 100).toFixed(1) + " %" : "—"])} />}
              {tab === "offres" && <Tbl head={["Code offre", "Libellé", "Qté", "CA"]} aligns={["left", "left", "right", "right"]} rows={result.perOffre.map(o => [o.code, o.label || (o.error ? "⚠ " + o.error : ""), fmtNum(o.qtyTotal), fmtEur(o.caTotal)])} />}
            </div>
          </div>
        )}
      </div>

      {showOffrePanel && <OffrePanel onClose={() => setShowOffrePanel(false)} onToast={onToast} onChanged={reloadAll} />}
      {showCampPanel && <CampagnePanel onClose={() => setShowCampPanel(false)} onToast={onToast} offres={offres} onChanged={reloadAll} />}
    </div>
  );
}

// Petit composant table
function Tbl({ head, aligns, rows }: { head: string[]; aligns: ("left" | "right" | "center")[]; rows: (string | number)[][] }) {
  if (!rows.length) return <div style={{ padding: 30, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucune donnée</div>;
  return (
    <table style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead>
        <tr style={{ background: C.bg }}>
          {head.map((h, i) => <th key={i} style={{ padding: "10px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", textAlign: aligns[i], borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
        </tr>
      </thead>
      <tbody>
        {rows.map((r, ri) => (
          <tr key={ri} style={{ background: ri % 2 ? "#f9fafb" : C.white }}>
            {r.map((c, ci) => <td key={ci} style={{ padding: "9px 14px", fontSize: 13, color: ci === 0 ? C.text : C.textSec, fontWeight: ci === 0 ? 600 : 400, textAlign: aligns[ci], borderBottom: `1px solid ${C.border}` }}>{c}</td>)}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
