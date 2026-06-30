"use client";
import { useState, useEffect } from "react";
import * as odoo from "@/lib/odoo";
import {
  CampagneCreee, PalierSaisi, ArticleCampagne, genId,
  analyseCampagneCreee, toExportPayload, qtyParPack, TYPES_PRODUIT,
} from "@/lib/create-campaign";
import { loadCampagnesCreees, upsertCampagneCreee, deleteCampagneCreee } from "@/lib/campaigns";
import { buildSyntheseLogistique } from "@/lib/logistique";

const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0",
  blue: "#3b82f6", blueDark: "#1d4ed8", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5",
  red: "#ef4444", redSoft: "#fef2f2",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  teal: "#0d9488", tealSoft: "#f0fdfa",
  shadow: "0 1px 3px rgba(0,0,0,0.06)", shadowMd: "0 4px 16px rgba(0,0,0,0.10)",
};
const fmtNum = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n || 0));

interface Props {
  session: odoo.OdooSession;
  onToast: (msg: string, type?: "success" | "error" | "info") => void;
  initialDraft?: CampagneCreee | null;     // brouillon transféré depuis l'analyse (préco)
  onDraftConsumed?: () => void;
}

function emptyArticle(): ArticleCampagne { return { ref: "" }; }
function emptyPalier(n: number): PalierSaisi {
  return { code: `REGE${n}`, label: n === 1 ? "Premium" : n === 2 ? "Standard" : n === 3 ? "Essentiel" : `Palier ${n}`, nbPacks: 0, qtyParPack: {} };
}
function emptyCampagne(): CampagneCreee {
  return { id: genId(), nom: "", dateDebut: "", dateFin: "", periodeDebut: "", periodeFin: "", articles: [emptyArticle()], paliers: [emptyPalier(1)] };
}

const inputStyle: React.CSSProperties = {
  padding: "7px 10px", border: `1px solid ${C.border}`, borderRadius: 7, fontSize: 13,
  fontFamily: "inherit", color: C.text, background: C.white, outline: "none",
};
const labelStyle: React.CSSProperties = { fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 4, display: "block" };

export default function CreerCampagneScreen({ session, onToast, initialDraft, onDraftConsumed }: Props) {
  const [camp, setCamp] = useState<CampagneCreee>(emptyCampagne);
  const [saved, setSaved] = useState<CampagneCreee[]>([]);
  const [analysing, setAnalysing] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportingMulti, setExportingMulti] = useState(false);
  const [saving, setSaving] = useState(false);
  const [analysed, setAnalysed] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  useEffect(() => { void reload(); }, []);
  // Charger le brouillon transféré depuis l'analyse (préco N+1), une seule fois.
  useEffect(() => {
    if (initialDraft) {
      setCamp(initialDraft);
      setAnalysed(false);
      onToast("Préco transférée — renseigne les dates puis « Analyser »", "info");
      onDraftConsumed?.();
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [initialDraft]);
  async function reload() {
    try { setSaved(await loadCampagnesCreees()); } catch (e: any) { onToast("Erreur chargement : " + e.message, "error"); }
  }

  const setField = (k: keyof CampagneCreee, v: any) => setCamp(c => ({ ...c, [k]: v }));
  const setArticle = (ai: number, patch: Partial<ArticleCampagne>) =>
    setCamp(c => ({ ...c, articles: c.articles.map((a, i) => i === ai ? { ...a, ...patch } : a) }));
  const addArticle = () => setCamp(c => ({ ...c, articles: [...c.articles, emptyArticle()] }));
  const removeArticle = (ai: number) => setCamp(c => ({ ...c, articles: c.articles.filter((_, i) => i !== ai) }));

  const setPalier = (pi: number, patch: Partial<PalierSaisi>) =>
    setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i === pi ? { ...p, ...patch } : p) }));
  const setPalierQty = (pi: number, ref: string, val: number | null) =>
    setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => {
      if (i !== pi) return p;
      const q = { ...p.qtyParPack };
      if (val == null) delete q[ref]; else q[ref] = val;
      return { ...p, qtyParPack: q };
    }) }));
  const addPalier = () => setCamp(c => ({ ...c, paliers: [...c.paliers, emptyPalier(c.paliers.length + 1)] }));
  const removePalier = (pi: number) => setCamp(c => ({ ...c, paliers: c.paliers.filter((_, i) => i !== pi) }));

  const analyser = async () => {
    if (!camp.periodeDebut || !camp.periodeFin) { onToast("Renseigne la période N-1 (début et fin)", "error"); return; }
    setAnalysing(true);
    try {
      const enriched = await analyseCampagneCreee(session, camp);
      setCamp(enriched); setAnalysed(true);
      const nbFound = enriched.articles.filter(a => a.found).length;
      const nbTotal = enriched.articles.filter(a => a.ref.trim()).length;
      // Diagnostic % offres : nb de paliers ayant reçu une reco par statut (calculée par code offre).
      const nbReco = enriched.paliers.filter(p => p.pctOffresReco && p.pctOffresReco.some(v => v > 0)).length;
      const recoMsg = nbReco > 0
        ? ` | % offres recalculés par statut sur ${nbReco}/${enriched.paliers.length} palier(s)`
        : " | ⚠ aucune répartition par statut trouvée (code offre absent des ventes N-1 ?) — % du gabarit conservés";
      onToast(`Conso N-1 récupérée (${nbFound}/${nbTotal} articles trouvés)${recoMsg}`, nbFound === nbTotal ? "success" : "info");
    } catch (e: any) { onToast("Erreur analyse : " + e.message, "error"); }
    finally { setAnalysing(false); }
  };

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

  const exporter = async () => {
    const payload: any = toExportPayload(camp);
    if (!payload.paliers.length) { onToast("Aucun article à exporter", "error"); return; }
    setExporting(true);
    try {
      // Charger tout le catalogue Odoo pour remplir l'onglet Mapping (VLOOKUP sur n'importe quelle réf).
      try {
        const catalogue = await odoo.getAllProducts(session);
        payload.mapping = catalogue.map(p => ({ ref: p.ref, name: p.name, barcode: p.barcode, standardPrice: p.standardPrice, listPrice: p.listPrice, ppc: p.ppc }));
      } catch { /* si le catalogue échoue, le Mapping se limite aux articles de la campagne */ }
      const res = await fetch("/api/export-template", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `proposition_${(camp.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Fichier Proposition exporté", "success");
    } catch (e: any) { onToast("Erreur export : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  // ── Export multi-campagnes (sélection cochée) + synthèse logistique ────────
  const toggleSelect = (id: string) => setSelected(s => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });

  const exporterMulti = async () => {
    const choisies = saved.filter(s => selected.has(s.id));
    if (!choisies.length) { onToast("Coche au moins une campagne", "error"); return; }
    setExportingMulti(true);
    try {
      // Enrichir chaque campagne (conso N-1 + pricing + libellés) si incomplet, puis payload.
      // On ré-analyse dès qu'un article n'a pas de libellé OU pas de prix, pour garantir
      // que la synthèse logistique affiche bien les libellés.
      const enriched: CampagneCreee[] = [];
      for (const c of choisies) {
        const complet = c.articles.length > 0 && c.articles.every(a => !a.ref.trim() || (a.name && a.listPrice != null));
        enriched.push(complet ? c : await analyseCampagneCreee(session, c));
      }
      const campagnes = enriched.map(c => toExportPayload(c)).map((p, i) => ({ nom: enriched[i].nom, paliers: p.paliers }));
      const logistique = buildSyntheseLogistique(enriched);
      const res = await fetch("/api/export-multi", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ campagnes, logistique }) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `campagnes_annee.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast(`Export de ${choisies.length} campagne(s) + synthèse logistique`, "success");
    } catch (e: any) { onToast("Erreur export multi : " + e.message, "error"); }
    finally { setExportingMulti(false); }
  };

  const articlesValides = camp.articles.filter(a => a.ref.trim());

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
    <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Créer une campagne</h1>
        <span style={{ fontSize: 13, color: C.textMuted }}>Mêmes articles dans tous les paliers ; seules les quantités changent.</span>
        <div style={{ flex: 1 }} />
        <button onClick={nouvelle} style={{ padding: "7px 14px", background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit" }}>+ Nouvelle</button>
      </div>

      {saved.length > 0 && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", boxShadow: C.shadow }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Mes campagnes sauvegardées</span>
            <span style={{ fontSize: 12, color: C.textMuted }}>Coche celles à inclure dans l'export annuel.</span>
            <div style={{ flex: 1 }} />
            <button onClick={exporterMulti} disabled={exportingMulti || selected.size === 0} style={{ padding: "7px 14px", background: selected.size ? C.teal : C.border, border: "none", borderRadius: 8, cursor: exportingMulti || !selected.size ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exportingMulti ? 0.6 : 1 }}>{exportingMulti ? "Export…" : `⬇ Exporter sélection (${selected.size}) + synthèse logistique`}</button>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {saved.map(s => (
              <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, background: selected.has(s.id) ? C.tealSoft : C.white, border: `1px solid ${selected.has(s.id) ? C.teal : C.border}`, borderRadius: 8, padding: "5px 8px 5px 8px", fontSize: 12 }}>
                <input type="checkbox" checked={selected.has(s.id)} onChange={() => toggleSelect(s.id)} style={{ cursor: "pointer" }} />
                <button onClick={() => charger(s)} style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>{s.nom || "(sans nom)"}</button>
                <button onClick={() => supprimer(s.id)} title="Supprimer" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.textMuted, fontSize: 14 }}>×</button>
              </div>
            ))}
          </div>
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
          <div style={{ fontSize: 12, color: C.textMuted, paddingBottom: 8 }}>L'app sommera les ventes de chaque article sur cette fenêtre.</div>
        </div>
      </div>

      {/* Articles de la campagne (communs à tous les paliers) */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: 18, boxShadow: C.shadow }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: C.text, marginBottom: 10 }}>Articles de la campagne</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {["Code", "Libellé", "Type", "EAN", "Coût", "Tarif rev.", "PPC", "Conso N-1", ""].map((h, i) => (
                <th key={i} style={{ padding: "6px 6px", fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.03em", textAlign: i >= 4 && i <= 7 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {camp.articles.map((a, ai) => {
              // Réf inconnue d'Odoo (après analyse) → surlignée pour signaler une saisie 100% libre.
              const inconnue = analysed && a.found === false && !!a.ref.trim();
              // Tous les champs (désignation, EAN, prix) sont TOUJOURS éditables : modifier une
              // valeur marque l'article comme manuel pour que l'analyse ne l'écrase plus.
              const numCell = (val: number | undefined, key: "standardPrice" | "listPrice" | "ppc") => (
                <input type="number" step="0.01" style={{ ...inputStyle, width: 75, textAlign: "right", padding: "5px 6px" }} value={val ?? ""} onChange={e => setArticle(ai, { manuel: true, [key]: e.target.value === "" ? 0 : parseFloat(e.target.value) } as any)} placeholder="—" />
              );
              return (
              <tr key={ai} style={{ background: inconnue ? C.amberSoft : "transparent" }}>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <input style={{ ...inputStyle, width: 110, fontFamily: "monospace", padding: "5px 6px" }} value={a.ref} onChange={e => setArticle(ai, { ref: e.target.value, name: undefined, found: undefined, manuel: undefined })} placeholder="Code" />
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <input style={{ ...inputStyle, width: 200, padding: "5px 6px" }} value={a.name ?? ""} onChange={e => setArticle(ai, { manuel: true, name: e.target.value })} placeholder="Désignation" />
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <select style={{ ...inputStyle, width: 130, padding: "5px 6px" }} value={a.typProd ?? "Produit Vente"} onChange={e => setArticle(ai, { typProd: e.target.value })}>
                    {TYPES_PRODUIT.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <input style={{ ...inputStyle, width: 120, padding: "5px 6px", fontFamily: "monospace" }} value={a.barcode ?? ""} onChange={e => setArticle(ai, { manuel: true, barcode: e.target.value })} placeholder="EAN" />
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>{numCell(a.standardPrice, "standardPrice")}</td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>{numCell(a.listPrice, "listPrice")}</td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>{numCell(a.ppc, "ppc")}</td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, color: C.textSec }}>{analysed ? fmtNum(a.consoN1 || 0) : "—"}</td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}`, textAlign: "center" }}>
                  {camp.articles.length > 1 && <button onClick={() => removeArticle(ai)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.textMuted, fontSize: 15 }}>×</button>}
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 8 }}>
          <button onClick={addArticle} style={{ padding: "5px 12px", background: C.blueSoft, border: `1px solid ${C.blue}`, borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>+ Ajouter un article</button>
          <span style={{ fontSize: 11, color: C.textMuted }}>Tous les champs (désignation, EAN, prix) sont éditables : tu peux saisir/ajuster les prix à la main, même pour une réf inconnue d'Odoo (surlignée).</span>
        </div>
      </div>

      {/* Paliers : nb packs + qté/pack par article */}
      {camp.paliers.map((pal, pi) => (
        <div key={pi} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
          <div style={{ padding: "12px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <input style={{ ...inputStyle, width: 90, fontWeight: 700 }} value={pal.code} onChange={e => setPalier(pi, { code: e.target.value })} placeholder="Code" />
            <input style={{ ...inputStyle, width: 150 }} value={pal.label} onChange={e => setPalier(pi, { label: e.target.value })} placeholder="Libellé" />
            <span style={{ fontSize: 12, color: C.textMuted }}>Nb packs cible</span>
            <input type="number" style={{ ...inputStyle, width: 90 }} value={pal.nbPacks || ""} onChange={e => setPalier(pi, { nbPacks: parseInt(e.target.value) || 0 })} placeholder="0" />
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: C.textSec, cursor: "pointer", marginLeft: 6 }}>
              <input type="checkbox" checked={!!pal.remiseStandard} onChange={e => setPalier(pi, { remiseStandard: e.target.checked })} style={{ cursor: "pointer" }} />
              Remise statut standard
            </label>
            {pal.remiseStandard && (
              <input type="number" step="0.1" style={{ ...inputStyle, width: 70 }} value={pal.remiseStandardTaux != null ? Math.round(pal.remiseStandardTaux * 1000) / 10 : ""} onChange={e => setPalier(pi, { remiseStandardTaux: e.target.value === "" ? undefined : (parseFloat(e.target.value) || 0) / 100 })} placeholder="17" title="% remise pour toutes les typologies" />
            )}
            {pal.remiseStandard && <span style={{ fontSize: 12, color: C.textMuted }}>%</span>}
            <span style={{ fontSize: 12, color: C.textMuted, marginLeft: 6 }}>Remise add.</span>
            <input type="number" step="0.1" style={{ ...inputStyle, width: 70 }} value={pal.remiseAddTaux != null ? Math.round(pal.remiseAddTaux * 1000) / 10 : ""} onChange={e => setPalier(pi, { remiseAddTaux: e.target.value === "" ? undefined : (parseFloat(e.target.value) || 0) / 100 })} placeholder="—" title="% remise additionnelle (colonne I) appliquée à tous les produits du palier" />
            <span style={{ fontSize: 12, color: C.textMuted }}>%</span>
            <div style={{ flex: 1 }} />
            {camp.paliers.length > 1 && <button onClick={() => removePalier(pi)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.red, fontSize: 13, fontWeight: 600, fontFamily: "inherit" }}>Supprimer le palier</button>}
          </div>
          <div style={{ padding: "8px 16px 14px" }}>
            {articlesValides.length === 0 ? (
              <div style={{ fontSize: 13, color: C.textMuted, padding: "8px 0" }}>Ajoute des articles à la campagne ci-dessus.</div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    {["Code", "Libellé", "Conso N-1", "Reco", "Qté / pack"].map((h, i) => (
                      <th key={i} style={{ padding: "6px 8px", fontSize: 11, fontWeight: 700, color: i === 4 ? C.blue : C.textMuted, textTransform: "uppercase", letterSpacing: "0.03em", textAlign: i >= 2 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {articlesValides.map((a) => {
                    const reco = pal.nbPacks > 0 ? Math.round((a.consoN1 || 0) / pal.nbPacks) : 0;
                    const m = pal.qtyParPack[a.ref];
                    // Valeur affichée : saisie >0 sinon reco. Un 0 résiduel n'écrase pas la reco.
                    const manual = (m != null && m > 0) ? m : undefined;
                    return (
                    <tr key={a.ref}>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontFamily: "monospace", fontSize: 13, color: C.text }}>{a.ref}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontSize: 13, color: C.textSec }}>{a.name || "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, color: C.textMuted }}>{analysed ? fmtNum(a.consoN1 || 0) : "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, fontWeight: 700, color: pal.nbPacks > 0 ? C.teal : C.textMuted }}>{pal.nbPacks > 0 ? fmtNum(reco) : "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>
                        <input type="number" style={{ ...inputStyle, width: 80, textAlign: "right", fontWeight: 700, color: C.blue }}
                          value={manual != null ? manual : (pal.nbPacks > 0 ? reco : "")}
                          onChange={e => setPalierQty(pi, a.ref, e.target.value === "" ? null : (parseInt(e.target.value) || 0))}
                          placeholder={pal.nbPacks > 0 ? String(reco) : "—"} />
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
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
