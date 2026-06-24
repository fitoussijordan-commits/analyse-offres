"use client";
import { useState, useEffect, useRef } from "react";
import * as odoo from "@/lib/odoo";
import * as cp from "@/lib/campaigns";
import { fetchCampaign, type CampaignResult, type StateFilter } from "@/lib/analyse-campaign";
import { buildPreco } from "@/lib/preco";
import { ChartCard, HBarChart, PieChart, SplitBar, fmtEurShort } from "./CampagneCharts";

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
function OffrePanel({ onClose, onToast, onChanged, session }: { onClose: () => void; onToast: Props["onToast"]; onChanged: () => void; session: odoo.OdooSession }) {
  const [offres, setOffres] = useState<cp.Offre[]>([]);
  const [showForm, setShowForm] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [code, setCode] = useState(""); const [label, setLabel] = useState(""); const [produits, setProduits] = useState(""); const [codeInterne, setCodeInterne] = useState("");

  // MEA search
  const [meaQuery, setMeaQuery] = useState("");
  const [meaSuggestions, setMeaSuggestions] = useState<odoo.MeaTemplate[]>([]);
  const [meaLoading, setMeaLoading] = useState(false);
  const [meaDropOpen, setMeaDropOpen] = useState(false);
  const meaRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!meaDropOpen) return;
    const handler = (e: MouseEvent) => { if (meaRef.current && !meaRef.current.contains(e.target as Node)) setMeaDropOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [meaDropOpen]);

  const searchMea = async (q: string) => {
    if (!q.trim()) { setMeaSuggestions([]); setMeaDropOpen(false); return; }
    setMeaLoading(true); setMeaDropOpen(true);
    try { const results = await odoo.searchMeaTemplates(session, q); setMeaSuggestions(results); }
    catch (e: any) { setMeaSuggestions([]); onToast("Recherche MEA : " + (e?.message || "erreur Odoo"), "error"); }
    finally { setMeaLoading(false); }
  };

  const loadMea = async (tpl: odoo.MeaTemplate) => {
    setMeaDropOpen(false); setMeaQuery(tpl.name); setMeaLoading(true);
    try {
      const lines = await odoo.getMeaTemplateLines(session, tpl.id);
      if (!lines.length) { onToast("Aucun produit trouvé dans ce modèle", "error"); return; }
      setProduits(lines.map(l => l.productCode).join("\n"));
      if (!label) setLabel(tpl.name);
      onToast(`${lines.length} produit(s) chargé(s) depuis "${tpl.name}"`, "success");
    } catch (e: any) { onToast("Erreur chargement MEA : " + e.message, "error"); }
    finally { setMeaLoading(false); }
  };

  const reload = () => cp.loadOffres().then(setOffres).catch(e => onToast("Erreur : " + e.message, "error"));
  useEffect(() => { reload(); }, []);

  const openNew = () => { setEditId(null); setCode(""); setLabel(""); setProduits(""); setCodeInterne(""); setMeaQuery(""); setMeaSuggestions([]); setShowForm(true); };
  const openEdit = (o: cp.Offre) => { setEditId(o.id); setCode(o.code); setLabel(o.label); setProduits(o.produits.join("\n")); setCodeInterne(o.codeInterne || ""); setMeaQuery(""); setMeaSuggestions([]); setShowForm(true); };

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
              {/* Recherche MEA Odoo */}
              <div ref={meaRef} style={{ position: "relative" }}>
                <label style={labelStyle}>Charger depuis Odoo (modèle de devis)</label>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  <input value={meaQuery} onChange={e => { setMeaQuery(e.target.value); searchMea(e.target.value); }} placeholder="Rechercher une MEA par nom…" style={{ ...inputStyle, flex: 1 }} />
                  {meaLoading && <span style={{ width: 16, height: 16, borderRadius: "50%", border: `2px solid ${C.blue}`, borderTopColor: "transparent", display: "inline-block", animation: "spin 0.7s linear infinite", flexShrink: 0 }} />}
                </div>
                {meaDropOpen && meaSuggestions.length > 0 && (
                  <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: C.white, border: `1.5px solid ${C.blue}44`, borderRadius: 10, boxShadow: "0 8px 32px rgba(0,0,0,0.13)", zIndex: 100, maxHeight: 220, overflowY: "auto" }}>
                    {meaSuggestions.map(tpl => (
                      <div key={tpl.id} onClick={() => loadMea(tpl)} style={{ padding: "10px 14px", cursor: "pointer", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 8 }}
                        onMouseEnter={e => (e.currentTarget as HTMLDivElement).style.background = C.bg}
                        onMouseLeave={e => (e.currentTarget as HTMLDivElement).style.background = C.white}>
                        <span style={{ fontSize: 13, fontWeight: 600, color: C.text }}>{tpl.name}</span>
                        {!tpl.active && <span style={{ fontSize: 10, background: C.amberSoft, color: C.amber, borderRadius: 4, padding: "1px 6px", fontWeight: 600 }}>Archivé</span>}
                      </div>
                    ))}
                  </div>
                )}
                {meaDropOpen && !meaLoading && meaSuggestions.length === 0 && meaQuery.trim() && (
                  <div style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, background: C.white, border: `1.5px solid ${C.border}`, borderRadius: 10, padding: "12px 14px", zIndex: 100, fontSize: 12, color: C.textMuted }}>Aucun modèle trouvé</div>
                )}
              </div>
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
type Tab = "produits" | "delegues" | "categories" | "adherents" | "statuts" | "offres" | "commandes" | "preco";
const FILTERS: [StateFilter, string][] = [["all", "Tout"], ["avenir", "À venir"], ["valide", "Validé"]];

// URL vers la fiche commande Odoo (vue formulaire)
function odooOrderUrl(baseUrl: string, orderId: number): string {
  const base = (baseUrl || "").replace(/\/+$/, "");
  return `${base}/web#id=${orderId}&model=sale.order&view_type=form`;
}

// Construit la liste dédoublonnée des commandes (offres + notes), comme l'onglet Excel "Toutes Commandes"
interface CmdRow { id: number; name: string; partner: string; code: string; label: string; type: "Offre" | "Note"; ca: number; avenir: boolean; orderTotal: number; }
function buildCommandes(result: CampaignResult): CmdRow[] {
  const seen = new Set<string>();
  const rows: CmdRow[] = [];
  for (const r of result.results) {
    for (const o of r.debugOrders) {
      const n = o.name.replace(" (note)", "");
      if (seen.has(n)) continue;
      seen.add(n);
      rows.push({ id: o.id, name: n, partner: o.partnerName ?? "", code: r.offre.code, label: r.offre.label, type: "Offre", ca: o.ca ?? 0, avenir: !o.invoiced, orderTotal: o.orderTotal ?? 0 });
    }
  }
  for (const c of result.catchalls) {
    for (const o of c.data?.debugOrders ?? []) {
      const n = o.name.replace(" (note)", "");
      if (seen.has(n)) continue;
      seen.add(n);
      rows.push({ id: o.id, name: n, partner: o.partnerName ?? "", code: c.codeInterne, label: "Note interne", type: "Note", ca: o.ca ?? 0, avenir: !o.invoiced, orderTotal: o.orderTotal ?? 0 });
    }
  }
  return rows.sort((a, b) => a.name.localeCompare(b.name));
}

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

  const TABS: [Tab, string][] = [["produits", "Produits"], ["delegues", "Délégués"], ["categories", "Catégorie statistique"], ["adherents", "Adhérent réseau"], ["statuts", "Statut client"], ["offres", "Par offre"], ["commandes", "Commandes"], ["preco", "Préco N+1"]];

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

            {/* Graphiques */}
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 18 }}>
              {filter === "all" && result.split && (result.split.valide.ca > 0 || result.split.avenir.ca > 0) && (
                <ChartCard title="Validé vs à venir">
                  <SplitBar valide={result.split.valide.ca} avenir={result.split.avenir.ca} />
                </ChartCard>
              )}
              {result.perOffre.filter(o => o.caTotal > 0).length > 0 && (
                <ChartCard title="CA par offre">
                  <HBarChart data={result.perOffre.filter(o => o.caTotal > 0).sort((a, b) => b.caTotal - a.caTotal).map(o => ({ label: o.label || o.code, value: o.caTotal }))} color={C.teal} valueFmt={fmtEurShort} />
                </ChartCard>
              )}
              {result.categories.filter(c => c.ca > 0).length > 0 && (
                <ChartCard title="Répartition par catégorie statistique">
                  <PieChart data={result.categories.map(c => ({ label: c.name, value: c.ca }))} />
                </ChartCard>
              )}
              {result.adherents.filter(a => a.ca > 0).length > 0 && (
                <ChartCard title="Répartition par adhérent réseau">
                  <PieChart data={result.adherents.map(a => ({ label: a.name, value: a.ca }))} />
                </ChartCard>
              )}
              {result.statuts.filter(s => s.ca > 0).length > 0 && (
                <ChartCard title="Répartition par statut client">
                  <PieChart data={result.statuts.map(s => ({ label: s.name, value: s.ca }))} />
                </ChartCard>
              )}
              {result.delegues.filter(d => d.ca > 0).length > 0 && (
                <ChartCard title="Top délégués (CA)">
                  <HBarChart data={result.delegues.slice(0, 8).map(d => ({ label: d.name, value: d.ca }))} color={C.purple} valueFmt={fmtEurShort} />
                </ChartCard>
              )}
              {result.produits.filter(p => p.ca > 0).length > 0 && (
                <ChartCard title="Top produits (CA)">
                  <HBarChart data={result.produits.slice(0, 8).map(p => ({ label: p.name, value: p.ca }))} color={C.blue} valueFmt={fmtEurShort} />
                </ChartCard>
              )}
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", gap: 4, marginBottom: 12, flexWrap: "wrap" }}>
              {TABS.map(([k, l]) => (
                <button key={k} onClick={() => setTab(k)} style={{ padding: "7px 14px", border: `1px solid ${tab === k ? C.teal : C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, background: tab === k ? C.tealSoft : C.white, color: tab === k ? C.teal : C.textSec, fontFamily: "inherit" }}>{l}</button>
              ))}
            </div>

            {/* Tables */}
            {tab === "offres" ? (
              <OffresDrillDown result={result} />
            ) : tab === "commandes" ? (
              <CommandesTab result={result} baseUrl={session.config.url} />
            ) : tab === "preco" ? (
              <PrecoTab result={result} onToast={onToast} session={session} />
            ) : (
              <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
                {tab === "produits" && <Tbl
                  head={["Réf", "Produit", "Qté", "CA", "% CA"]} aligns={["left", "left", "right", "right", "right"]}
                  rows={result.produits.map(p => [p.ref, p.name, fmtNum(p.qtyVendue), fmtEur(p.ca), pctOf(p.ca, result.caTotal)])}
                  total={["TOTAL", "", fmtNum(result.produits.reduce((s, p) => s + p.qtyVendue, 0)), fmtEur(result.caTotal), "100,0 %"]}
                />}
                {tab === "delegues" && <Tbl
                  head={["Délégué", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right"]}
                  rows={result.delegues.map(d => [d.name, fmtNum(d.qtyVendue), fmtEur(d.ca), pctOf(d.ca, result.caTotal)])}
                  total={["TOTAL", fmtNum(result.delegues.reduce((s, d) => s + d.qtyVendue, 0)), fmtEur(result.caTotal), "100,0 %"]}
                />}
                {tab === "categories" && <Tbl
                  head={["Catégorie statistique", "Cmd", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right", "right"]}
                  rows={result.categories.map(c => [c.name, fmtNum(c.nbCommandes), fmtNum(c.qtyVendue), fmtEur(c.ca), pctOf(c.ca, result.caTotal)])}
                  total={["TOTAL", fmtNum(result.categories.reduce((s, c) => s + c.nbCommandes, 0)), fmtNum(result.categories.reduce((s, c) => s + c.qtyVendue, 0)), fmtEur(result.caTotal), "100,0 %"]}
                />}
                {tab === "adherents" && <Tbl
                  head={["Adhérent réseau", "Cmd", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right", "right"]}
                  rows={result.adherents.map(a => [a.name, fmtNum(a.nbCommandes), fmtNum(a.qtyVendue), fmtEur(a.ca), pctOf(a.ca, result.caTotal)])}
                  total={["TOTAL", fmtNum(result.adherents.reduce((s, a) => s + a.nbCommandes, 0)), fmtNum(result.adherents.reduce((s, a) => s + a.qtyVendue, 0)), fmtEur(result.caTotal), "100,0 %"]}
                />}
                {tab === "statuts" && <Tbl
                  head={["Statut client", "Cmd", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right", "right"]}
                  rows={result.statuts.map(s => [s.name, fmtNum(s.nbCommandes), fmtNum(s.qtyVendue), fmtEur(s.ca), pctOf(s.ca, result.caTotal)])}
                  total={["TOTAL", fmtNum(result.statuts.reduce((s, x) => s + x.nbCommandes, 0)), fmtNum(result.statuts.reduce((s, x) => s + x.qtyVendue, 0)), fmtEur(result.caTotal), "100,0 %"]}
                />}
              </div>
            )}
          </div>
        )}
      </div>

      {showOffrePanel && <OffrePanel onClose={() => setShowOffrePanel(false)} onToast={onToast} onChanged={reloadAll} session={session} />}
      {showCampPanel && <CampagnePanel onClose={() => setShowCampPanel(false)} onToast={onToast} offres={offres} onChanged={reloadAll} />}
    </div>
  );
}

// ════════════════════════════════════════════════════════════════════════════
// Onglet "Préco N+1" : produits conservés PAR PALIER (offre) → besoin fournisseur
// ════════════════════════════════════════════════════════════════════════════
function PrecoTab({ result, onToast, session }: { result: CampaignResult; onToast: Props["onToast"]; session: odoo.OdooSession }) {
  // État : produits conservés par code offre (palier). Par défaut, tout est conservé.
  const initial = () => {
    const m: Record<string, Set<number>> = {};
    for (const r of result.results) if (!r.error) m[r.offre.code] = new Set(r.produits.map(p => p.productId));
    return m;
  };
  const [conserved, setConserved] = useState<Record<string, Set<number>>>(initial);
  const [exporting, setExporting] = useState(false);
  const [exportingTpl, setExportingTpl] = useState(false);

  const toggle = (code: string, id: number) => setConserved(prev => {
    const n = { ...prev }; const set = new Set(n[code] ?? []);
    set.has(id) ? set.delete(id) : set.add(id); n[code] = set; return n;
  });

  const conservedIds: Record<string, number[]> = {};
  for (const k of Object.keys(conserved)) conservedIds[k] = [...conserved[k]];
  const preco = buildPreco(result, conservedIds);

  const exportPreco = async () => {
    setExporting(true);
    try {
      const res = await fetch("/api/export-preco", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(preco) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `preco_${result.nom.replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Préco exportée", "success");
    } catch (e: any) { onToast("Erreur export préco : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  // Export au format du template "Proposition template" (Fichier trade) :
  // on enrichit la préco avec les données tarifaires Odoo (EAN, coût achat,
  // tarif revendeur, PPC) puis on délègue le remplissage du gabarit à l'API.
  const exportTemplate = async () => {
    setExportingTpl(true);
    try {
      // 1) Ne garder que les produits conservés de chaque palier.
      const palierProds = preco.paliers.map(pal => ({
        ...pal,
        produits: pal.produits.filter(p => p.conserve),
      })).filter(pal => pal.produits.length);
      if (!palierProds.length) { onToast("Aucun produit conservé à exporter", "error"); return; }

      // 2) Récupérer le pricing Odoo pour tous les produits concernés (1 seul appel).
      const allIds = [...new Set(palierProds.flatMap(p => p.produits.map(x => x.productId)))];
      const pricing = await odoo.getProductsPricing(session, allIds);

      // 3) Construire le payload enrichi.
      const payload = {
        nom: preco.nom,
        paliers: palierProds.map(pal => ({
          code: pal.code,
          label: pal.label,
          qtyPacks: pal.qtyPacks,
          produits: pal.produits.map(p => {
            const pr = pricing[p.productId];
            return {
              ref: p.ref, name: p.name, productId: p.productId,
              qtyParPack: p.qtyParPack, conserve: true,
              typProd: "Produit Vente",
              barcode: pr?.barcode || "",
              standardPrice: pr?.standardPrice || 0,
              listPrice: pr?.listPrice || 0,
              ppc: pr?.ppc || 0,
            };
          }),
        })),
      };

      const res = await fetch("/api/export-template", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `proposition_${result.nom.replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Template Proposition exporté", "success");
    } catch (e: any) { onToast("Erreur export template : " + e.message, "error"); }
    finally { setExportingTpl(false); }
  };

  if (!preco.paliers.length) return <div style={{ padding: 30, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucun palier (offre) à analyser pour la préconisation.</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {/* Barre d'action */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Construction de l'offre N+1 par palier</span>
        <span style={{ fontSize: 12, color: C.textMuted }}>Décoche les produits que tu ne reconduis pas dans chaque offre.</span>
        <div style={{ flex: 1 }} />
        <button onClick={exportPreco} disabled={exporting} style={{ padding: "7px 16px", background: C.teal, border: "none", borderRadius: 8, cursor: exporting ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exporting ? 0.6 : 1 }}>{exporting ? "Export…" : "⬇ Exporter (préco + besoin fournisseur)"}</button>
        <button onClick={exportTemplate} disabled={exportingTpl} style={{ padding: "7px 16px", background: C.blue, border: "none", borderRadius: 8, cursor: exportingTpl ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exportingTpl ? 0.6 : 1 }}>{exportingTpl ? "Export…" : "⬇ Exporter template Proposition"}</button>
      </div>

      {/* Un bloc par palier (offre) */}
      {preco.paliers.map(pal => {
        const nbConserves = pal.produits.filter(p => p.conserve).length;
        return (
          <div key={pal.code} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
            <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.border}`, background: C.blueSoft, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <span style={{ fontSize: 11, fontWeight: 800, fontFamily: "monospace", background: C.blue, color: "#fff", borderRadius: 5, padding: "2px 8px" }}>{pal.code}</span>
              <span style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{pal.label || "Offre"}</span>
              <span style={{ fontSize: 12, color: C.textMuted }}>{fmtEur(pal.caTotal)} · {fmtNum(pal.qtyPacks)} pack(s)</span>
              <div style={{ flex: 1 }} />
              <span style={{ fontSize: 12, color: C.textMuted }}>{nbConserves}/{pal.produits.length} conservé(s)</span>
            </div>
            <div style={{ maxHeight: 320, overflowY: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                  <tr style={{ background: C.bg }}>
                    {["", "Réf", "Produit", "Qté vendue", "Qté / pack", "CA"].map((h, i) => <th key={i} style={{ padding: "9px 14px", fontSize: 11, fontWeight: 700, color: i === 4 ? C.blue : C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", textAlign: i >= 3 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {pal.produits.map((p, ri) => (
                    <tr key={p.productId} onClick={() => toggle(pal.code, p.productId)} style={{ cursor: "pointer", background: p.conserve ? C.tealSoft : ri % 2 ? "#f9fafb" : C.white, opacity: p.conserve ? 1 : 0.55 }}>
                      <td style={{ padding: "8px 14px", width: 36, borderBottom: `1px solid ${C.border}` }}>
                        <input type="checkbox" checked={p.conserve} onChange={() => toggle(pal.code, p.productId)} style={{ cursor: "pointer" }} />
                      </td>
                      <td style={{ padding: "8px 14px", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.ref}</td>
                      <td style={{ padding: "8px 14px", fontSize: 13, color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.name}</td>
                      <td style={{ padding: "8px 14px", fontSize: 13, fontWeight: 700, color: C.text, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtNum(p.qty)}</td>
                      <td style={{ padding: "8px 14px", fontSize: 13, fontWeight: 800, color: C.blue, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtNum(p.qtyParPack)}</td>
                      <td style={{ padding: "8px 14px", fontSize: 13, color: C.textSec, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtEur(p.ca)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {/* Composition type d'1 pack (produits conservés) */}
            {(() => {
              const compo = pal.produits.filter(p => p.conserve && p.qtyParPack > 0);
              const totalUnites = compo.reduce((s, p) => s + p.qtyParPack, 0);
              if (!compo.length) return null;
              return (
                <div style={{ padding: "12px 16px", borderTop: `1px solid ${C.border}`, background: "#fafbff" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.blue, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>
                    🧱 Composition type d'1 pack — {fmtNum(totalUnites)} unités
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                    {compo.map(p => (
                      <div key={p.productId} style={{ display: "flex", alignItems: "center", gap: 6, background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, padding: "5px 10px", fontSize: 12 }}>
                        <span style={{ fontWeight: 800, color: C.blue }}>{fmtNum(p.qtyParPack)}×</span>
                        <span style={{ fontFamily: "monospace", fontWeight: 600, color: C.text }}>{p.ref}</span>
                        <span style={{ color: C.textMuted, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name.replace(/^\[\d+\]\s*/, "")}</span>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          </div>
        );
      })}

      {/* Récap besoin fournisseur (par produit, total) */}
      <div style={{ background: C.white, border: `2px solid ${C.teal}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadowMd }}>
        <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.border}`, background: C.tealSoft, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: 14, fontWeight: 800, color: C.teal }}>📦 Besoin en commande fournisseur</span>
          <span style={{ fontSize: 12, color: C.textMuted }}>Quantités à commander pour l'offre N+1 (base : ventes campagne actuelle)</span>
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 13, fontWeight: 800, color: C.teal }}>Total : {fmtNum(preco.totalQty)} u</span>
        </div>
        {preco.besoins.length === 0 ? (
          <div style={{ padding: 24, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucun produit conservé — coche des produits dans les paliers ci-dessus.</div>
        ) : (
          <div style={{ maxHeight: 480, overflowY: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                <tr style={{ background: C.bg }}>
                  {["Réf", "Produit", "Paliers", "Qté à commander", "CA associé"].map((h, i) => <th key={i} style={{ padding: "10px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", textAlign: i >= 3 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {preco.besoins.map((b, ri) => (
                  <tr key={b.productId} style={{ background: ri % 2 ? "#f9fafb" : C.white }}>
                    <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{b.ref}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{b.name}</td>
                    <td style={{ padding: "9px 14px", fontSize: 12, color: C.textMuted, fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{b.paliers.join(", ")}</td>
                    <td style={{ padding: "9px 14px", fontSize: 14, fontWeight: 800, color: C.teal, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtNum(b.qty)}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtEur(b.ca)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

// % d'une valeur sur un total, formaté fr-FR
function pctOf(val: number, total: number): string {
  if (!total) return "—";
  return ((val / total) * 100).toFixed(1).replace(".", ",") + " %";
}

// Petit composant table (avec ligne TOTAL optionnelle)
function Tbl({ head, aligns, rows, total }: { head: string[]; aligns: ("left" | "right" | "center")[]; rows: (string | number)[][]; total?: (string | number)[] }) {
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
      {total && (
        <tfoot>
          <tr style={{ background: C.tealSoft }}>
            {total.map((c, ci) => <td key={ci} style={{ padding: "10px 14px", fontSize: 13, fontWeight: 800, color: C.teal, textAlign: aligns[ci], borderTop: `2px solid ${C.teal}44` }}>{c}</td>)}
          </tr>
        </tfoot>
      )}
    </table>
  );
}

// ── Onglet "Par offre" : drill-down (produits composants + délégués par offre) ──
function OffresDrillDown({ result }: { result: CampaignResult }) {
  const [open, setOpen] = useState<string | null>(result.results[0]?.offre.code ?? null);
  // sources = offres analysées + notes (catchalls avec données)
  type Src = { key: string; code: string; label: string; ca: number; qty: number; cmd: number; produits: { ref: string; name: string; qtyVendue: number; ca: number }[]; delegues: { name: string; qtyVendue: number; ca: number }[]; error: string | null; note: boolean };
  const sources: Src[] = [
    ...result.results.map(r => ({ key: "o:" + r.offre.code, code: r.offre.code, label: r.offre.label, ca: r.caTotal, qty: r.qtyTotal, cmd: r.debugOrders.length, produits: r.produits, delegues: r.delegues, error: r.error, note: false })),
    ...result.catchalls.filter(c => c.data && (c.data.debugOrders.length > 0 || c.data.caTotal > 0)).map(c => ({ key: "n:" + c.codeInterne, code: c.codeInterne, label: "Note interne", ca: c.data!.caTotal, qty: c.data!.qtyTotal, cmd: c.data!.debugOrders.length, produits: c.data!.produits, delegues: c.data!.delegues, error: null, note: true })),
  ];
  if (!sources.length) return <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: 30, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucune offre</div>;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {sources.map(s => {
        const isOpen = open === s.key;
        return (
          <div key={s.key} style={{ background: C.white, border: `1px solid ${isOpen ? C.teal + "66" : C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
            <div onClick={() => setOpen(isOpen ? null : s.key)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 16px", cursor: "pointer", background: isOpen ? C.tealSoft : C.white }}>
              <span style={{ transform: `rotate(${isOpen ? 90 : 0}deg)`, transition: "transform .15s", color: C.textMuted, fontSize: 13 }}>▶</span>
              <span style={{ fontSize: 13, fontWeight: 700, fontFamily: "monospace", color: C.text }}>{s.code}</span>
              <span style={{ fontSize: 12, color: C.textMuted, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {s.error ? <span style={{ color: C.amber }}>⚠ {s.error}</span> : s.label}
                {s.note && <span style={{ marginLeft: 6, fontSize: 10, background: "#fff7ed", color: "#f97316", borderRadius: 4, padding: "1px 6px", fontWeight: 700 }}>NOTE</span>}
              </span>
              <span style={{ fontSize: 11, color: C.textMuted }}>{fmtNum(s.cmd)} cmd · {fmtNum(s.qty)} qté</span>
              <span style={{ fontSize: 15, fontWeight: 800, color: s.note ? "#f97316" : C.teal, minWidth: 90, textAlign: "right" }}>{fmtEur(s.ca)}</span>
            </div>
            {isOpen && (
              <div style={{ padding: "4px 16px 16px", borderTop: `1px solid ${C.border}` }}>
                {s.produits.length > 0 && (
                  <div style={{ marginTop: 14 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: C.blue, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Produits composants</div>
                    <Tbl head={["Réf", "Produit", "Qté", "CA", "% CA"]} aligns={["left", "left", "right", "right", "right"]} rows={s.produits.map(p => [p.ref, p.name, fmtNum(p.qtyVendue), fmtEur(p.ca), pctOf(p.ca, s.ca)])} />
                  </div>
                )}
                {s.delegues.length > 0 && (
                  <div style={{ marginTop: 18 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: C.purple, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>Par délégué</div>
                    <Tbl head={["Délégué", "Qté", "CA", "% CA"]} aligns={["left", "right", "right", "right"]} rows={s.delegues.map(d => [d.name, fmtNum(d.qtyVendue), fmtEur(d.ca), pctOf(d.ca, s.ca)])} />
                  </div>
                )}
                {!s.produits.length && !s.delegues.length && <div style={{ padding: 16, textAlign: "center", color: C.textMuted, fontSize: 12 }}>Aucun détail disponible</div>}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ── Onglet "Commandes" : liste dédoublonnée avec lien Odoo ─────────────────────
function CommandesTab({ result, baseUrl }: { result: CampaignResult; baseUrl: string }) {
  const [q, setQ] = useState("");
  const all = buildCommandes(result);
  const filtered = q.trim() ? all.filter(r => (r.name + " " + r.partner + " " + r.code + " " + r.label).toLowerCase().includes(q.trim().toLowerCase())) : all;
  const nbOffre = all.filter(r => r.type === "Offre").length;
  const nbNote = all.filter(r => r.type === "Note").length;
  const avenir = all.filter(r => r.avenir).sort((a, b) => b.ca - a.ca);
  const caAvenir = avenir.reduce((s, r) => s + r.ca, 0);
  return (
   <>
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
      <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="Rechercher (commande, client, offre…)" style={{ flex: "1 1 240px", maxWidth: 360, padding: "7px 11px", border: `1.5px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontFamily: "inherit", outline: "none", color: C.text }} />
        <span style={{ fontSize: 12, color: C.textMuted }}>{fmtNum(filtered.length)} commande(s) · <span style={{ color: C.teal, fontWeight: 600 }}>{fmtNum(nbOffre)} offre</span> · <span style={{ color: "#f97316", fontWeight: 600 }}>{fmtNum(nbNote)} note</span></span>
      </div>
      {filtered.length === 0 ? (
        <div style={{ padding: 30, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucune commande</div>
      ) : (
        <div style={{ maxHeight: 560, overflowY: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
              <tr style={{ background: C.bg }}>
                {["Commande", "Client", "Source", "Libellé", "CA offre", "Total cmd", "Poids", "Type", ""].map((h, i) => <th key={i} style={{ padding: "10px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", textAlign: (i >= 4 && i <= 6) || i === 7 ? (i === 7 ? "center" : "right") : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {filtered.map((r, ri) => {
                const isNote = r.type === "Note";
                return (
                  <tr key={r.name} style={{ background: isNote ? "#fff7ed" : ri % 2 ? "#f9fafb" : C.white }}>
                    <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}`, fontFamily: "monospace" }}>{r.name}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{r.partner || "—"}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: isNote ? "#f97316" : C.teal, borderBottom: `1px solid ${C.border}`, fontFamily: "monospace" }}>{r.code}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{r.label}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: C.text, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtEur(r.ca)}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{r.orderTotal > 0 ? fmtEur(r.orderTotal) : "—"}</td>
                    <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 700, textAlign: "right", borderBottom: `1px solid ${C.border}`, color: r.orderTotal > 0 ? (r.ca / r.orderTotal >= 0.999 ? C.teal : C.textSec) : C.textMuted }}>{r.orderTotal > 0 ? pctOf(r.ca, r.orderTotal) : "—"}</td>
                    <td style={{ padding: "9px 14px", fontSize: 11, textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                      <span style={{ fontWeight: 700, padding: "2px 8px", borderRadius: 5, background: isNote ? "#ffedd5" : C.tealSoft, color: isNote ? "#f97316" : C.teal }}>{r.type}</span>
                    </td>
                    <td style={{ padding: "9px 14px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                      <a href={odooOrderUrl(baseUrl, r.id)} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12, fontWeight: 600, color: C.blue, textDecoration: "none", whiteSpace: "nowrap" }} title="Ouvrir dans Odoo">Odoo ↗</a>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>

    {/* ── Section CA à venir ──────────────────────────────────────────────── */}
    <div style={{ marginTop: 18, background: C.white, border: `1px solid ${C.amber}55`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
      <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", background: C.amberSoft }}>
        <span style={{ fontSize: 13, fontWeight: 700, color: "#b45309" }}>Commandes — CA à venir</span>
        <span style={{ fontSize: 12, color: C.textMuted }}>{fmtNum(avenir.length)} commande(s)</span>
        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 15, fontWeight: 800, color: C.amber }}>{fmtEur(caAvenir)}</span>
      </div>
      {avenir.length === 0 ? (
        <div style={{ padding: 24, textAlign: "center", color: C.textMuted, fontSize: 13 }}>Aucune commande à venir — tout est facturé</div>
      ) : (
        <div style={{ maxHeight: 420, overflowY: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
              <tr style={{ background: C.bg }}>
                {["Commande", "Client", "Source", "CA à venir", ""].map((h, i) => <th key={i} style={{ padding: "10px 14px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", textAlign: i === 3 ? "right" : i === 4 ? "center" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {avenir.map((r, ri) => (
                <tr key={r.name} style={{ background: ri % 2 ? "#f9fafb" : C.white }}>
                  <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}`, fontFamily: "monospace" }}>{r.name}</td>
                  <td style={{ padding: "9px 14px", fontSize: 13, color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{r.partner || "—"}</td>
                  <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 600, color: r.type === "Note" ? "#f97316" : C.teal, borderBottom: `1px solid ${C.border}`, fontFamily: "monospace" }}>{r.code}</td>
                  <td style={{ padding: "9px 14px", fontSize: 13, fontWeight: 700, color: C.amber, textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtEur(r.ca)}</td>
                  <td style={{ padding: "9px 14px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                    <a href={odooOrderUrl(baseUrl, r.id)} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12, fontWeight: 600, color: C.blue, textDecoration: "none", whiteSpace: "nowrap" }} title="Ouvrir dans Odoo">Odoo ↗</a>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
   </>
  );
}
