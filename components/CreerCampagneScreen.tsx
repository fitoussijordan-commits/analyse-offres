"use client";
import { useState, useEffect } from "react";
import * as odoo from "@/lib/odoo";
import {
  CampagneCreee, PalierSaisi, ArticleSaisi, genId,
  analyseCampagneCreee, toExportPayload,
} from "@/lib/create-campaign";
import { loadCampagnesCreees, upsertCampagneCreee, deleteCampagneCreee } from "@/lib/campaigns";

const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0",
  blue: "#3b82f6", blueDark: "#1d4ed8", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  red: "#ef4444", redSoft: "#fef2f2",
  teal: "#0d9488", tealSoft: "#f0fdfa",
  shadow: "0 1px 3px rgba(0,0,0,0.06)", shadowMd: "0 4px 16px rgba(0,0,0,0.10)",
};
const fmtNum = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n || 0));
const fmtEur = (n: number) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n || 0);

interface Props {
  session: odoo.OdooSession;
  onToast: (msg: string, type?: "success" | "error" | "info") => void;
}

function emptyArticle(): ArticleSaisi { return { ref: "" }; }
function emptyPalier(n: number): PalierSaisi {
  return { code: `REGE${n}`, label: n === 1 ? "Premium" : n === 2 ? "Standard" : n === 3 ? "Essentiel" : `Palier ${n}`, nbPacks: 0, articles: [emptyArticle()] };
}
function emptyCampagne(): CampagneCreee {
  return { id: genId(), nom: "", dateDebut: "", dateFin: "", periodeDebut: "", periodeFin: "", paliers: [emptyPalier(1)] };
}

const inputStyle: React.CSSProperties = {
  padding: "7px 10px", border: `1px solid ${C.border}`, borderRadius: 7, fontSize: 13,
  fontFamily: "inherit", color: C.text, background: C.white, outline: "none",
};
const labelStyle: React.CSSProperties = { fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 4, display: "block" };

export default function CreerCampagneScreen({ session, onToast }: Props) {
  const [camp, setCamp] = useState<CampagneCreee>(emptyCampagne);
  const [saved, setSaved] = useState<CampagneCreee[]>([]);
  const [analysing, setAnalysing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [saving, setSaving] = useState(false);
  const [analysed, setAnalysed] = useState(false);

  useEffect(() => { void reload(); }, []);
  async function reload() {
    try { setSaved(await loadCampagnesCreees()); } catch (e: any) { onToast("Erreur chargement : " + e.message, "error"); }
  }

  // ── Mutations du formulaire ───────────────────────────────────────────────
  const setField = (k: keyof CampagneCreee, v: any) => setCamp(c => ({ ...c, [k]: v }));
  const setPalier = (pi: number, patch: Partial<PalierSaisi>) =>
    setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i === pi ? { ...p, ...patch } : p) }));
  const setArticle = (pi: number, ai: number, patch: Partial<ArticleSaisi>) =>
    setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i !== pi ? p : { ...p, articles: p.articles.map((a, j) => j === ai ? { ...a, ...patch } : a) }) }));
  const addPalier = () => setCamp(c => ({ ...c, paliers: [...c.paliers, emptyPalier(c.paliers.length + 1)] }));
  const removePalier = (pi: number) => setCamp(c => ({ ...c, paliers: c.paliers.filter((_, i) => i !== pi) }));
  const addArticle = (pi: number) => setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i === pi ? { ...p, articles: [...p.articles, emptyArticle()] } : p) }));
  const removeArticle = (pi: number, ai: number) => setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i !== pi ? p : { ...p, articles: p.articles.filter((_, j) => j !== ai) }) }));

  // ── Analyse conso N-1 + pricing Odoo ──────────────────────────────────────
  const analyser = async () => {
    if (!camp.periodeDebut || !camp.periodeFin) { onToast("Renseigne la période N-1 (début et fin)", "error"); return; }
    setAnalysing(true);
    try {
      const enriched = await analyseCampagneCreee(session, camp);
      setCamp(enriched);
      setAnalysed(true);
      const nbFound = enriched.paliers.flatMap(p => p.articles).filter(a => a.found).length;
      const nbTotal = enriched.paliers.flatMap(p => p.articles).filter(a => a.ref.trim()).length;
      onToast(`Conso N-1 récupérée (${nbFound}/${nbTotal} articles trouvés dans Odoo)`, nbFound === nbTotal ? "success" : "info");
    } catch (e: any) { onToast("Erreur analyse : " + e.message, "error"); }
    finally { setAnalysing(false); }
  };

  // ── Sauvegarde Supabase ───────────────────────────────────────────────────
  const sauvegarder = async () => {
    if (!camp.nom.trim()) { onToast("Donne un nom à la campagne", "error"); return; }
    setSaving(true);
    try { await upsertCampagneCreee(camp); await reload(); onToast("Campagne sauvegardée", "success"); }
    catch (e: any) { onToast("Erreur sauvegarde : " + e.message, "error"); }
    finally { setSaving(false); }
  };

  const charger = (c: CampagneCreee) => { setCamp(c); setAnalysed(false); onToast(`Campagne « ${c.nom} » chargée`, "info"); };
  const supprimer = async (id: string) => { try { await deleteCampagneCreee(id); await reload(); onToast("Supprimée", "success"); } catch (e: any) { onToast(e.message, "error"); } };
  const nouvelle = () => { setCamp(emptyCampagne()); setAnalysed(false); };

  // ── Export Proposition ─────────────────────────────────────────────────────
  const exporter = async () => {
    const payload = toExportPayload(camp);
    if (!payload.paliers.length) { onToast("Aucun article à exporter", "error"); return; }
    setExporting(true);
    try {
      const res = await fetch("/api/export-template", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `proposition_${(camp.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Fichier Proposition exporté", "success");
    } catch (e: any) { onToast("Erreur export : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
    <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Créer une campagne</h1>
        <span style={{ fontSize: 13, color: C.textMuted }}>Construis une campagne de zéro et recommande les quantités selon la conso N-1.</span>
        <div style={{ flex: 1 }} />
        <button onClick={nouvelle} style={{ padding: "7px 14px", background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit" }}>+ Nouvelle</button>
      </div>

      {/* Campagnes sauvegardées */}
      {saved.length > 0 && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {saved.map(s => (
            <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, padding: "5px 6px 5px 12px", fontSize: 12, boxShadow: C.shadow }}>
              <button onClick={() => charger(s)} style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>{s.nom || "(sans nom)"}</button>
              <button onClick={() => supprimer(s.id)} title="Supprimer" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.textMuted, fontSize: 14 }}>×</button>
            </div>
          ))}
        </div>
      )}

      {/* Infos campagne */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: 18, boxShadow: C.shadow }}>
        <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr 1fr", gap: 14, marginBottom: 14 }}>
          <div><label style={labelStyle}>Nom de la campagne</label><input style={{ ...inputStyle, width: "100%" }} value={camp.nom} onChange={e => setField("nom", e.target.value)} placeholder="Ex. Régénérants 2026" /></div>
          <div><label style={labelStyle}>Début campagne</label><input type="date" style={{ ...inputStyle, width: "100%" }} value={camp.dateDebut} onChange={e => setField("dateDebut", e.target.value)} /></div>
          <div><label style={labelStyle}>Fin campagne</label><input type="date" style={{ ...inputStyle, width: "100%" }} value={camp.dateFin} onChange={e => setField("dateFin", e.target.value)} /></div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.5fr", gap: 14, alignItems: "end" }}>
          <div><label style={labelStyle}>Période N-1 — début</label><input type="date" style={{ ...inputStyle, width: "100%" }} value={camp.periodeDebut} onChange={e => setField("periodeDebut", e.target.value)} /></div>
          <div><label style={labelStyle}>Période N-1 — fin</label><input type="date" style={{ ...inputStyle, width: "100%" }} value={camp.periodeFin} onChange={e => setField("periodeFin", e.target.value)} /></div>
          <div style={{ fontSize: 12, color: C.textMuted, paddingBottom: 8 }}>L'app sommera les ventes de chaque article sur cette fenêtre pour recommander la qté/pack.</div>
        </div>
      </div>

      {/* Paliers */}
      {camp.paliers.map((pal, pi) => (
        <div key={pi} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
          <div style={{ padding: "12px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <input style={{ ...inputStyle, width: 90, fontWeight: 700 }} value={pal.code} onChange={e => setPalier(pi, { code: e.target.value })} placeholder="Code" />
            <input style={{ ...inputStyle, width: 150 }} value={pal.label} onChange={e => setPalier(pi, { label: e.target.value })} placeholder="Libellé" />
            <span style={{ fontSize: 12, color: C.textMuted }}>Nb packs cible</span>
            <input type="number" style={{ ...inputStyle, width: 90 }} value={pal.nbPacks || ""} onChange={e => setPalier(pi, { nbPacks: parseInt(e.target.value) || 0 })} placeholder="0" />
            <div style={{ flex: 1 }} />
            {camp.paliers.length > 1 && <button onClick={() => removePalier(pi)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.red, fontSize: 13, fontWeight: 600, fontFamily: "inherit" }}>Supprimer le palier</button>}
          </div>
          <div style={{ padding: "8px 16px 14px" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {["Code article", "Libellé (Odoo)", "Conso N-1", "Qté / pack (reco)", ""].map((h, i) => (
                    <th key={i} style={{ padding: "6px 8px", fontSize: 11, fontWeight: 700, color: i === 3 ? C.blue : C.textMuted, textTransform: "uppercase", letterSpacing: "0.03em", textAlign: i >= 2 && i < 4 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pal.articles.map((a, ai) => (
                  <tr key={ai}>
                    <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}` }}>
                      <input style={{ ...inputStyle, width: 120, fontFamily: "monospace" }} value={a.ref} onChange={e => setArticle(pi, ai, { ref: e.target.value, name: undefined, found: undefined })} placeholder="Code" />
                    </td>
                    <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontSize: 13, color: a.found === false && a.ref ? C.red : C.textSec }}>
                      {a.name || (a.found === false && a.ref ? "Introuvable dans Odoo" : "—")}
                    </td>
                    <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, color: C.textSec }}>{analysed ? fmtNum(a.consoN1 || 0) : "—"}</td>
                    <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>
                      <input type="number" style={{ ...inputStyle, width: 80, textAlign: "right", fontWeight: 700, color: C.blue }} value={a.qtyParPack ?? ""} onChange={e => setArticle(pi, ai, { qtyParPack: e.target.value === "" ? undefined : (parseInt(e.target.value) || 0) })} placeholder="—" />
                    </td>
                    <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "center" }}>
                      {pal.articles.length > 1 && <button onClick={() => removeArticle(pi, ai)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.textMuted, fontSize: 15 }}>×</button>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button onClick={() => addArticle(pi)} style={{ marginTop: 8, padding: "5px 12px", background: C.blueSoft, border: `1px solid ${C.blue}`, borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>+ Ajouter un article</button>
          </div>
        </div>
      ))}

      <button onClick={addPalier} style={{ alignSelf: "flex-start", padding: "8px 16px", background: C.white, border: `1px dashed ${C.blue}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>+ Ajouter un palier</button>

      {/* Barre d'action */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", position: "sticky", bottom: -24, background: C.bg, padding: "14px 24px", margin: "0 -24px -24px", borderTop: `1px solid ${C.border}`, zIndex: 5 }}>
        <button onClick={analyser} disabled={analysing} style={{ padding: "9px 18px", background: C.teal, border: "none", borderRadius: 8, cursor: analysing ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: analysing ? 0.6 : 1 }}>{analysing ? "Analyse…" : "📊 Analyser conso N-1 (Odoo)"}</button>
        <button onClick={sauvegarder} disabled={saving} style={{ padding: "9px 18px", background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: saving ? "default" : "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit", opacity: saving ? 0.6 : 1 }}>{saving ? "…" : "💾 Sauvegarder"}</button>
        <div style={{ flex: 1 }} />
        <button onClick={exporter} disabled={exporting} style={{ padding: "9px 18px", background: C.blue, border: "none", borderRadius: 8, cursor: exporting ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exporting ? 0.6 : 1 }}>{exporting ? "Export…" : "⬇ Exporter template Proposition"}</button>
      </div>
    </div>
    </div>
  );
}
