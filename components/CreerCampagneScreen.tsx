"use client";
import { useState, useEffect } from "react";
import * as odoo from "@/lib/odoo";
import {
  CampagneCreee, PalierSaisi, ArticleCampagne, genId,
  analyseCampagneCreee, toExportPayload, qtyParPack, totalPacks, ventilationPalier, TYPES_PRODUIT,
  GcEnseigne, GC_ENSEIGNES_DEFAUT,
} from "@/lib/create-campaign";
import { loadCampagnesCreees, upsertCampagneCreee, deleteCampagneCreee, loadCampagnesCreeesCorbeille, restoreCampagneCreee, hardDeleteCampagneCreee } from "@/lib/campaigns";
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
  return { code: "", label: "", nbPacks: 0, qtyParPack: {} };
}
function emptyCampagne(): CampagneCreee {
  return { id: genId(), nom: "", dateDebut: "", dateFin: "", periodeDebut: "", periodeFin: "", articles: [emptyArticle()], paliers: [emptyPalier(1)] };
}

// Clé de saisie qté/pack d'un article dans un palier. Si la réf est unique dans la campagne,
// on garde la réf seule (compatible avec les campagnes existantes). Si la réf apparaît
// plusieurs fois (ex. stick vendu + stick UG), on distingue par le type → "ref#type".
function qtyKey(art: ArticleCampagne, articles: ArticleCampagne[]): string {
  const ref = (art.ref || "").trim();
  const doublon = articles.filter(a => (a.ref || "").trim() === ref).length > 1;
  return doublon ? `${ref}#${art.typProd || "Produit Vente"}` : ref;
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
  const [filterYear, setFilterYear] = useState<string>("all");
  const [corbeille, setCorbeille] = useState<CampagneCreee[]>([]);
  const [showCorbeille, setShowCorbeille] = useState(false);
  const [openInfoPalier, setOpenInfoPalier] = useState<number | null>(null);

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
    try {
      setSaved(await loadCampagnesCreees());
      setCorbeille(await loadCampagnesCreeesCorbeille());
    } catch (e: any) { onToast("Erreur chargement : " + e.message, "error"); }
  }

  const setField = (k: keyof CampagneCreee, v: any) => setCamp(c => ({ ...c, [k]: v }));
  const setArticle = (ai: number, patch: Partial<ArticleCampagne>) =>
    setCamp(c => ({ ...c, articles: c.articles.map((a, i) => i === ai ? { ...a, ...patch } : a) }));
  const addArticle = () => setCamp(c => ({ ...c, articles: [...c.articles, emptyArticle()] }));
  const removeArticle = (ai: number) => setCamp(c => ({ ...c, articles: c.articles.filter((_, i) => i !== ai) }));

  const setPalier = (pi: number, patch: Partial<PalierSaisi>) =>
    setCamp(c => ({ ...c, paliers: c.paliers.map((p, i) => i === pi ? { ...p, ...patch } : p) }));

  // ── Grands Comptes (enseignes dynamiques, toutes exportées dans l'Excel).
  const GC_MAX = 10; // limite UI raisonnable
  // undefined → nouvelles campagnes → afficher les 6 défauts ; [] → l'utilisateur a tout retiré.
  const gcEnseignes: GcEnseigne[] = camp.gcEnseignes != null ? camp.gcEnseignes : GC_ENSEIGNES_DEFAUT;
  const ensureGc = (c: CampagneCreee): GcEnseigne[] => c.gcEnseignes != null ? c.gcEnseignes.map(e => ({ ...e, qties: { ...e.qties } })) : GC_ENSEIGNES_DEFAUT.map(e => ({ ...e, qties: {} }));
  const setGcQty = (ei: number, key: string, v: number) => setCamp(c => { const g = ensureGc(c); g[ei] = { ...g[ei], qties: { ...g[ei].qties, [key]: v } }; return { ...c, gcEnseignes: g }; });
  const setGcNom = (ei: number, nom: string) => setCamp(c => { const g = ensureGc(c); g[ei] = { ...g[ei], nom }; return { ...c, gcEnseignes: g }; });
  const setGcRemise = (ei: number, remise: number) => setCamp(c => { const g = ensureGc(c); g[ei] = { ...g[ei], remise }; return { ...c, gcEnseignes: g }; });
  const addGc = () => setCamp(c => { const g = ensureGc(c); if (g.length >= GC_MAX) return c; return { ...c, gcEnseignes: [...g, { nom: `GC${g.length + 1}`, remise: 0, qties: {} }] }; });
  const removeGc = (ei: number) => setCamp(c => { const g = ensureGc(c); return { ...c, gcEnseignes: g.filter((_, i) => i !== ei) }; });
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
      // Diagnostic % offres : montre la répartition par typologie trouvée (ou prévient si rien).
      const TYPOS = ["Ambassadeur", "Compagnon", "Challenger", "Rose", "Prunelier", "Anthylide", "Calendula"];
      const reco = enriched.paliers.find(p => p.pctOffresReco && p.pctOffresReco.some(v => v > 0))?.pctOffresReco;
      const recoMsg = reco
        ? " | % offres reco : " + reco.map((v, i) => v > 0 ? `${TYPOS[i]} ${Math.round(v * 100)}%` : null).filter(Boolean).join(", ")
        : " | ⚠ aucun statut client trouvé sur les ventes N-1 (champ statut vide dans Odoo ?) — % du gabarit conservés";
      onToast(`Conso N-1 (${nbFound}/${nbTotal} articles)${recoMsg}`, nbFound === nbTotal ? "success" : "info");
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
  const supprimer = async (id: string) => { try { await deleteCampagneCreee(id); await reload(); onToast("Déplacée dans la corbeille", "success"); } catch (e: any) { onToast(e.message, "error"); } };
  const restaurer = async (id: string) => { try { await restoreCampagneCreee(id); await reload(); onToast("Campagne restaurée", "success"); } catch (e: any) { onToast(e.message, "error"); } };
  const supprimerDefinitif = async (c: CampagneCreee) => {
    if (!window.confirm(`Supprimer DÉFINITIVEMENT « ${c.nom || "(sans nom)"} » ? Cette action est irréversible.`)) return;
    try { await hardDeleteCampagneCreee(c.id); await reload(); onToast("Supprimée définitivement", "success"); } catch (e: any) { onToast(e.message, "error"); }
  };
  const nouvelle = () => { setCamp(emptyCampagne()); setAnalysed(false); };

  // Vide les quantités saisies de tous les paliers → laisse la reco automatique
  // (conso ÷ total des packs) s'appliquer. Ne modifie que l'état local : rien n'est
  // enregistré tant que l'utilisateur ne clique pas sur « Sauvegarder ».
  const reinitialiserRecos = () => {
    const nbSaisies = camp.paliers.reduce((s, p) => s + Object.keys(p.qtyParPack || {}).length, 0);
    if (nbSaisies === 0) { onToast("Aucune quantité saisie : la reco est déjà active", "info"); return; }
    if (!window.confirm(`Effacer ${nbSaisies} quantité(s) saisie(s) et laisser la reco automatique ? (non enregistré tant que tu ne sauvegardes pas)`)) return;
    setCamp(c => ({ ...c, paliers: c.paliers.map(p => ({ ...p, qtyParPack: {} })) }));
    onToast("Recos réinitialisées — vérifie puis sauvegarde", "success");
  };

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
      // Synthèse logistique (réf × mois) de cette campagne → onglet dans le même fichier.
      payload.logistique = buildSyntheseLogistique([camp]);
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
  // Dénominateur de la reco : total des packs de TOUS les paliers (la conso est répartie,
  // pas recopiée dans chaque palier).
  const totalP = totalPacks(camp.paliers);

  // Année / cycle de rangement = le champ saisi par l'utilisateur (sinon "—").
  const yearOf = (c: CampagneCreee): string => (c.annee || "").trim() || "—";
  const annees = [...new Set(saved.map(yearOf))].sort((a, b) => b.localeCompare(a));
  const savedFiltered = filterYear === "all" ? saved : saved.filter(c => yearOf(c) === filterYear);

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
    <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Créer une campagne</h1>
        <span style={{ fontSize: 13, color: C.textMuted }}>Mêmes articles dans tous les paliers ; seules les quantités changent.</span>
        <div style={{ flex: 1 }} />
        <button onClick={reinitialiserRecos} title="Efface les quantités saisies pour laisser la reco automatique (conso ÷ total packs)" style={{ padding: "7px 14px", background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit" }}>↻ Réinitialiser les recos</button>
        <button onClick={() => setShowCorbeille(v => !v)} style={{ padding: "7px 14px", background: showCorbeille ? C.bg : C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit" }}>🗑 Corbeille{corbeille.length ? ` (${corbeille.length})` : ""}</button>
        <button onClick={nouvelle} style={{ padding: "7px 14px", background: C.white, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.text, fontFamily: "inherit" }}>+ Nouvelle</button>
      </div>

      {showCorbeille && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", boxShadow: C.shadow }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>🗑 Corbeille</span>
            <span style={{ fontSize: 12, color: C.textMuted }}>Restaure une campagne ou supprime-la définitivement.</span>
          </div>
          {corbeille.length === 0 ? (
            <span style={{ fontSize: 12, color: C.textMuted }}>La corbeille est vide.</span>
          ) : (
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {corbeille.map(s => (
                <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 8, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: "5px 8px", fontSize: 12 }}>
                  <span style={{ fontWeight: 600, color: C.text }}>{s.nom || "(sans nom)"}</span>
                  <span style={{ fontSize: 10, color: C.textMuted, background: C.white, borderRadius: 4, padding: "1px 5px" }}>{yearOf(s)}</span>
                  <button onClick={() => restaurer(s.id)} title="Restaurer" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.teal, fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>↩ Restaurer</button>
                  <button onClick={() => supprimerDefinitif(s)} title="Supprimer définitivement" style={{ border: "none", background: "transparent", cursor: "pointer", color: "#c0392b", fontSize: 12, fontWeight: 700, fontFamily: "inherit" }}>Supprimer définitivement</button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {saved.length > 0 && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", boxShadow: C.shadow }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Mes campagnes sauvegardées</span>
            <select value={filterYear} onChange={e => setFilterYear(e.target.value)} style={{ ...inputStyle, width: 120, padding: "5px 8px" }}>
              <option value="all">Toutes années</option>
              {annees.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <span style={{ fontSize: 12, color: C.textMuted }}>Coche celles à inclure dans l'export annuel.</span>
            <div style={{ flex: 1 }} />
            <button onClick={exporterMulti} disabled={exportingMulti || selected.size === 0} style={{ padding: "7px 14px", background: selected.size ? C.teal : C.border, border: "none", borderRadius: 8, cursor: exportingMulti || !selected.size ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exportingMulti ? 0.6 : 1 }}>{exportingMulti ? "Export…" : `⬇ Exporter sélection (${selected.size}) + synthèse logistique`}</button>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {savedFiltered.map(s => (
              <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 6, background: selected.has(s.id) ? C.tealSoft : C.white, border: `1px solid ${selected.has(s.id) ? C.teal : C.border}`, borderRadius: 8, padding: "5px 8px 5px 8px", fontSize: 12 }}>
                <input type="checkbox" checked={selected.has(s.id)} onChange={() => toggleSelect(s.id)} style={{ cursor: "pointer" }} />
                <button onClick={() => charger(s)} style={{ border: "none", background: "transparent", cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>{s.nom || "(sans nom)"}</button>
                <span style={{ fontSize: 10, color: C.textMuted, background: C.bg, borderRadius: 4, padding: "1px 5px" }}>{yearOf(s)}</span>
                <button onClick={() => supprimer(s.id)} title="Mettre à la corbeille" style={{ border: "none", background: "transparent", cursor: "pointer", color: C.textMuted, fontSize: 14 }}>×</button>
              </div>
            ))}
            {savedFiltered.length === 0 && <span style={{ fontSize: 12, color: C.textMuted }}>Aucune campagne pour {filterYear}.</span>}
          </div>
        </div>
      )}

      {/* Infos campagne */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: 18, boxShadow: C.shadow }}>
        <div style={{ display: "grid", gridTemplateColumns: "1.5fr 0.7fr 1fr 1fr", gap: 14, marginBottom: 14 }}>
          <div><label style={labelStyle}>Nom de la campagne</label><input style={{ ...inputStyle, width: "100%" }} value={camp.nom} onChange={e => setField("nom", e.target.value)} placeholder="Ex. Régénérants 2026" /></div>
          <div><label style={labelStyle}>Année / cycle</label><input style={{ ...inputStyle, width: "100%" }} value={camp.annee ?? ""} onChange={e => setField("annee", e.target.value)} placeholder="Ex. 2027" /></div>
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
              const estVenteArt = (a.typProd ?? "Produit Vente") === "Produit Vente";
              const numCell = (val: number | undefined, key: "standardPrice" | "listPrice" | "ppc") => {
                // Tarif de vente et PPC sont forcés à 0 (lecture seule) pour les non-Produit Vente.
                const bloque = !estVenteArt && (key === "listPrice" || key === "ppc");
                return (
                  <input type="number" step="0.01" disabled={bloque}
                    style={{ ...inputStyle, width: 75, textAlign: "right", padding: "5px 6px", background: bloque ? C.bg : undefined, color: bloque ? C.textMuted : undefined }}
                    value={bloque ? 0 : (val ?? "")}
                    onChange={e => setArticle(ai, { manuel: true, [key]: e.target.value === "" ? 0 : parseFloat(e.target.value) } as any)}
                    placeholder="—" title={bloque ? "Gratuit (non Produit Vente) → 0" : undefined} />
                );
              };
              return (
              <tr key={ai} style={{ background: inconnue ? C.amberSoft : "transparent" }}>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <input style={{ ...inputStyle, width: 110, fontFamily: "monospace", padding: "5px 6px" }} value={a.ref} onChange={e => setArticle(ai, { ref: e.target.value, name: undefined, found: undefined, manuel: undefined })} placeholder="Code" />
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <input style={{ ...inputStyle, width: 200, padding: "5px 6px" }} value={a.name ?? ""} onChange={e => setArticle(ai, { manuel: true, name: e.target.value })} placeholder="Désignation" />
                </td>
                <td style={{ padding: "5px 6px", borderBottom: `1px solid ${C.border}` }}>
                  <select style={{ ...inputStyle, width: 130, padding: "5px 6px" }} value={a.typProd ?? "Produit Vente"} onChange={e => {
                    const t = e.target.value;
                    // Non "Produit Vente" (UG/Testeur/PLV/Échantillon) = gratuit → tarif de vente et PPC à 0.
                    const patch: any = { typProd: t };
                    if (t !== "Produit Vente") { patch.listPrice = 0; patch.ppc = 0; }
                    setArticle(ai, patch);
                  }}>
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
            <span style={{ fontSize: 12, color: C.textMuted }}>Nb offres cible</span>
            <input type="number" style={{ ...inputStyle, width: 90 }} value={pal.nbPacks || ""} onChange={e => setPalier(pi, { nbPacks: parseInt(e.target.value) || 0 })} placeholder="0" />
            <span style={{ fontSize: 12, color: C.textMuted, marginLeft: 6 }} title="Nombre total de produits (Produit Vente) par offre. La reco répartit ce total entre les articles au prorata des ventes N-1. Laisse vide pour la reco automatique.">Produits/offre</span>
            <input type="number" style={{ ...inputStyle, width: 80 }} value={pal.nbProduitsPack || ""} onChange={e => setPalier(pi, { nbProduitsPack: e.target.value === "" ? undefined : (parseInt(e.target.value) || 0) })} placeholder="auto" />
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: C.textSec, cursor: "pointer", marginLeft: 6 }}>
              <input type="checkbox" checked={!!pal.remiseStandard} onChange={e => setPalier(pi, { remiseStandard: e.target.checked })} style={{ cursor: "pointer" }} />
              Remise statut unique
            </label>
            {pal.remiseStandard && (
              <input type="number" step="0.1" style={{ ...inputStyle, width: 70 }} value={pal.remiseStandardTaux != null ? Math.round(pal.remiseStandardTaux * 1000) / 10 : ""} onChange={e => setPalier(pi, { remiseStandardTaux: e.target.value === "" ? undefined : (parseFloat(e.target.value) || 0) / 100 })} placeholder="17" title="% remise pour toutes les typologies" />
            )}
            {pal.remiseStandard && <span style={{ fontSize: 12, color: C.textMuted }}>%</span>}
            <span style={{ fontSize: 12, color: C.textMuted, marginLeft: 6 }}>Remise add.</span>
            <input type="number" step="0.1" style={{ ...inputStyle, width: 70 }} value={pal.remiseAddTaux != null ? Math.round(pal.remiseAddTaux * 1000) / 10 : ""} onChange={e => setPalier(pi, { remiseAddTaux: e.target.value === "" ? undefined : (parseFloat(e.target.value) || 0) / 100 })} placeholder="—" title="% remise additionnelle (colonne I) appliquée à tous les produits du palier" />
            <span style={{ fontSize: 12, color: C.textMuted }}>%</span>
            <div style={{ flex: 1 }} />
            <button
              onClick={() => setOpenInfoPalier(openInfoPalier === pi ? null : pi)}
              title="Voir le détail des calculs de ce palier"
              style={{ padding: "3px 8px", background: openInfoPalier === pi ? C.teal : "transparent", border: `1px solid ${openInfoPalier === pi ? C.teal : C.border}`, borderRadius: 6, cursor: "pointer", color: openInfoPalier === pi ? "#fff" : C.textMuted, fontSize: 13, fontWeight: 700, fontFamily: "inherit" }}
            >ℹ</button>
            {camp.paliers.length > 1 && <button onClick={() => removePalier(pi)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.red, fontSize: 13, fontWeight: 600, fontFamily: "inherit" }}>Supprimer le palier</button>}
          </div>
          {openInfoPalier === pi && (() => {
            const vent = ventilationPalier(articlesValides, pal);
            const modeVent = !!(pal.nbProduitsPack && pal.nbProduitsPack > 0);
            const TYPOS = ["Ambassadeur", "Compagnon", "Challenger", "Rose", "Prunelier", "Anthylide", "Calendula"];
            const rows: { label: string; value: string; note?: string }[] = [];
            // ── Mode de calcul ──
            rows.push({ label: "Mode reco", value: modeVent ? `N produits/offre (${pal.nbProduitsPack})` : "Standard (conso ÷ total offres)" });
            if (modeVent) {
              const totalConsoVente = articlesValides.filter(a => (a.typProd || "Produit Vente") === "Produit Vente" && (a.consoN1 || 0) > 0).reduce((s, a) => s + (a.consoN1 || 0), 0);
              rows.push({ label: "Formule", value: `qté/offre = round(consoN1 / ${fmtNum(totalConsoVente)} × ${pal.nbProduitsPack})`, note: "Répartition au prorata, méthode du plus grand reste" });
            } else {
              rows.push({ label: "Formule", value: `qté/offre = round(consoN1 ÷ ${fmtNum(totalP)})`, note: `${fmtNum(totalP)} = total offres sur toute la campagne (${camp.paliers.map(p => `${p.label || p.code} ×${p.nbPacks}`).join(" + ")})` });
            }
            rows.push({ label: "Nb offres palier", value: fmtNum(pal.nbPacks) });
            rows.push({ label: "Total offres campagne", value: fmtNum(totalP) });
            // ── Remises ──
            if (pal.remiseStandard && pal.remiseStandardTaux != null) rows.push({ label: "Remise statut unique", value: `${Math.round(pal.remiseStandardTaux * 1000) / 10} %`, note: "Appliquée à toutes les typologies (colonne remise unifiée)" });
            if (pal.remiseAddTaux != null && pal.remiseAddTaux > 0) rows.push({ label: "Remise additionnelle", value: `${Math.round(pal.remiseAddTaux * 1000) / 10} %`, note: "Colonne I de la Proposition (s'ajoute à la remise par typo)" });
            // ── % typologies ──
            if (pal.pctOffresReco && pal.pctOffresReco.some(v => v > 0)) {
              rows.push({ label: "% offres par typo", value: pal.pctOffresReco.map((v, i) => v > 0 ? `${TYPOS[i]} ${Math.round(v * 100)}%` : null).filter(Boolean).join("  ·  "), note: "Calculé sur les ventes N-1 analysées (répartition réelle des commandes par statut client)" });
            }
            return (
              <div style={{ margin: "0 16px 0", background: "#f8fafc", border: `1px solid ${C.teal}`, borderTop: "none", borderRadius: "0 0 8px 8px", padding: "10px 14px", display: "flex", flexDirection: "column", gap: 6 }}>
                <div style={{ fontSize: 11, fontWeight: 800, color: C.teal, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 2 }}>Détail des calculs — {pal.label || pal.code}</div>
                {rows.map((r, ri) => (
                  <div key={ri} style={{ display: "flex", gap: 10, alignItems: "flex-start", fontSize: 12 }}>
                    <span style={{ minWidth: 170, color: C.textMuted, fontWeight: 600, flexShrink: 0 }}>{r.label}</span>
                    <span style={{ color: C.text, fontFamily: r.label === "Formule" ? "monospace" : "inherit", fontWeight: r.label === "Formule" ? 700 : 400 }}>{r.value}</span>
                    {r.note && <span style={{ color: C.textMuted, fontSize: 11, marginLeft: 4 }}>— {r.note}</span>}
                  </div>
                ))}
                {!analysed && <div style={{ fontSize: 11, color: "#b45309", marginTop: 4 }}>⚠ Lance « Analyser conso N-1 » pour voir les % typologies et les consos réelles.</div>}
              </div>
            );
          })()}
          <div style={{ padding: "10px 16px 0", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12, color: C.textMuted, whiteSpace: "nowrap" }}>Descriptif</span>
            <input style={{ ...inputStyle, flex: 1 }} value={pal.descriptif ?? ""} onChange={e => setPalier(pi, { descriptif: e.target.value })} placeholder="Ex. Panachage 50 offres // Remise 15% pour BRI conso 10€…" title="Texte libre retranscrit dans l'Excel à côté du nom du palier" />
          </div>
          <div style={{ padding: "8px 16px 14px" }}>
            {articlesValides.length === 0 ? (
              <div style={{ fontSize: 13, color: C.textMuted, padding: "8px 0" }}>Ajoute des articles à la campagne ci-dessus.</div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    {["Code", "Libellé", "Type", "Conso N-1", "Reco", "Qté / offre"].map((h, i) => (
                      <th key={i} style={{ padding: "6px 8px", fontSize: 11, fontWeight: 700, color: i === 5 ? C.blue : C.textMuted, textTransform: "uppercase", letterSpacing: "0.03em", textAlign: i >= 3 ? "right" : "left", borderBottom: `1px solid ${C.border}` }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {articlesValides.map((a, ai) => {
                    // Clé composite : distingue les doublons de réf (ex. stick vendu vs stick UG).
                    const k = qtyKey(a, articlesValides);
                    // Ventilation "N produits/pack" si le palier a un total saisi, sinon undefined.
                    const vent = ventilationPalier(articlesValides, pal);
                    // Reco = ventilation prorata si mode N produits/pack, sinon conso ÷ total packs.
                    // La ventilation exclut les non-vente (UG/Testeur/PLV) → reco 0 pour eux.
                    const hasConso = (a.consoN1 || 0) > 0;
                    const estVenteLigne = (a.typProd || "Produit Vente") === "Produit Vente";
                    const reco = (hasConso && estVenteLigne) ? qtyParPack({ ...a, ref: a.ref }, pal, totalP, vent, articlesValides) : 0;
                    // Reco affichable si conso, Produit Vente ET (mode N produits/pack OU nb packs défini).
                    const recoDispo = hasConso && estVenteLigne && (!!vent || pal.nbPacks > 0);
                    const m = pal.qtyParPack[k];
                    // Valeur affichée : saisie >0 sinon reco. Un 0 résiduel n'écrase pas la reco.
                    const manual = (m != null && m > 0) ? m : undefined;
                    return (
                    <tr key={ai}>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontFamily: "monospace", fontSize: 13, color: C.text }}>{a.ref}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontSize: 13, color: C.textSec }}>{a.name || "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, fontSize: 12 }}>
                        {(() => {
                          const t = a.typProd || "Produit Vente";
                          const vente = t === "Produit Vente";
                          return <span title={vente ? "Génère du CA" : "Gratuit : prix de vente et PPC = 0"} style={{ fontWeight: 600, color: vente ? C.textMuted : "#b45309", background: vente ? "transparent" : "#fef3c7", borderRadius: 4, padding: vente ? 0 : "1px 6px" }}>{t}</span>;
                        })()}
                      </td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, color: C.textMuted }}>{analysed ? fmtNum(a.consoN1 || 0) : "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, fontWeight: 700, color: recoDispo ? C.teal : C.textMuted }}>{recoDispo ? fmtNum(reco) : "—"}</td>
                      <td style={{ padding: "5px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right" }}>
                        <input type="number" style={{ ...inputStyle, width: 80, textAlign: "right", fontWeight: 700, color: C.blue }}
                          value={manual != null ? manual : (recoDispo ? reco : "")}
                          onChange={e => setPalierQty(pi, k, e.target.value === "" ? null : (parseInt(e.target.value) || 0))}
                          placeholder={recoDispo ? String(reco) : "—"} />
                      </td>
                    </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
            {/* Alerte écart : mode N produits/pack → total réparti vs cible saisie. */}
            {pal.nbProduitsPack && pal.nbProduitsPack > 0 && (() => {
              const vent = ventilationPalier(articlesValides, pal);
              const totalReparti = articlesValides
                .filter(a => (a.typProd || "Produit Vente") === "Produit Vente")
                .reduce((s, a) => s + qtyParPack(a, pal, totalP, vent, articlesValides), 0);
              const cible = pal.nbProduitsPack;
              const ecart = totalReparti - cible;
              const ok = ecart === 0;
              return (
                <div style={{ marginTop: 8, padding: "6px 10px", borderRadius: 7, fontSize: 12, fontWeight: 600,
                  background: ok ? C.tealSoft : "#fef3c7", color: ok ? C.teal : "#b45309", display: "inline-block" }}>
                  {ok
                    ? `✓ ${totalReparti} produits/offre répartis (= cible ${cible})`
                    : `⚠ ${totalReparti} produits/offre répartis vs cible ${cible} (écart ${ecart > 0 ? "+" : ""}${ecart}) — ajuste les quantités`}
                </div>
              );
            })()}
          </div>
        </div>
      ))}

      <button onClick={addPalier} style={{ alignSelf: "flex-start", padding: "8px 16px", background: C.white, border: `1px dashed ${C.blue}`, borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.blueDark, fontFamily: "inherit" }}>+ Ajouter un palier</button>

      {/* Bloc GRANDS COMPTES : 6 enseignes (qté + remise par article), sauvegardé avec la campagne. */}
      {articlesValides.length > 0 && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
          <div style={{ padding: "12px 16px", background: "#fff7ed", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 11, fontWeight: 800, fontFamily: "monospace", background: "#b45309", color: "#fff", borderRadius: 5, padding: "2px 8px" }}>GRANDS COMPTES</span>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Quantités par enseigne (nom et remise éditables, sauvegardés)</span>
            <span style={{ fontSize: 11, color: C.textMuted, marginLeft: 4 }}>{gcEnseignes.length} enseigne{gcEnseignes.length > 1 ? "s" : ""}</span>
          </div>
          <div style={{ padding: "0 16px 14px", overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr>
                  <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }} colSpan={2}>Enseigne →</th>
                  {gcEnseignes.map((e, ei) => (
                    <th key={ei} style={{ padding: "4px 4px", borderBottom: `1px solid ${C.border}` }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
                        <input style={{ ...inputStyle, width: 84, textAlign: "center", fontWeight: 700, color: "#b45309", padding: "4px 6px" }} value={e.nom} onChange={ev => setGcNom(ei, ev.target.value)} />
                        <button onClick={() => removeGc(ei)} title="Supprimer cette enseigne" style={{ padding: "2px 5px", background: "none", border: `1px solid #fca5a5`, borderRadius: 5, cursor: "pointer", color: "#ef4444", fontSize: 11, fontWeight: 700, lineHeight: 1 }}>×</button>
                      </div>
                    </th>
                  ))}
                  <th style={{ padding: "4px 4px", borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" }}>
                    <button onClick={addGc} disabled={gcEnseignes.length >= GC_MAX} title={gcEnseignes.length >= GC_MAX ? `Maximum ${GC_MAX} enseignes (limite du template)` : "Ajouter une enseigne GC"} style={{ padding: "4px 8px", background: gcEnseignes.length >= GC_MAX ? C.border : "#fff7ed", border: `1px dashed #b45309`, borderRadius: 6, cursor: gcEnseignes.length >= GC_MAX ? "default" : "pointer", color: gcEnseignes.length >= GC_MAX ? C.textMuted : "#b45309", fontSize: 13, fontWeight: 700, opacity: gcEnseignes.length >= GC_MAX ? 0.4 : 1 }}>+</button>
                  </th>
                </tr>
                <tr>
                  <th style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted, fontWeight: 600, borderBottom: `1px solid ${C.border}` }} colSpan={2}>Remise %</th>
                  {gcEnseignes.map((e, ei) => (
                    <th key={ei} style={{ padding: "3px 4px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                      <input type="number" step="0.1" style={{ ...inputStyle, width: 64, textAlign: "center", padding: "4px 6px" }} value={Math.round(e.remise * 1000) / 10} onChange={ev => setGcRemise(ei, (parseFloat(ev.target.value) || 0) / 100)} />
                    </th>
                  ))}
                  <th style={{ borderBottom: `1px solid ${C.border}` }} />
                </tr>
                <tr>
                  <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Réf</th>
                  <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Libellé</th>
                  {gcEnseignes.map((e, ei) => <th key={ei} style={{ padding: "5px 4px", textAlign: "center", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Qté</th>)}
                  <th style={{ borderBottom: `1px solid ${C.border}` }} />
                </tr>
              </thead>
              <tbody>
                {articlesValides.map((a, ai) => {
                  const k = qtyKey(a, articlesValides);
                  return (
                    <tr key={ai}>
                      <td style={{ padding: "4px 8px", fontFamily: "monospace", fontSize: 13, borderBottom: `1px solid ${C.border}` }}>{a.ref}</td>
                      <td style={{ padding: "4px 8px", color: C.textSec, fontSize: 13, borderBottom: `1px solid ${C.border}` }}>{a.name || "—"}{(a.typProd && a.typProd !== "Produit Vente") ? ` (${a.typProd})` : ""}</td>
                      {gcEnseignes.map((e, ei) => (
                        <td key={ei} style={{ padding: "3px 4px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                          <input type="number" style={{ ...inputStyle, width: 60, textAlign: "center", color: "#b45309", padding: "4px 6px" }} value={e.qties[k] || ""} onChange={ev => setGcQty(ei, k, parseInt(ev.target.value) || 0)} placeholder="0" />
                        </td>
                      ))}
                      <td style={{ borderBottom: `1px solid ${C.border}` }} />
                    </tr>
                  );
                })}
                <tr style={{ background: "#fff7ed" }}>
                  <td style={{ padding: "5px 8px", fontWeight: 800, color: "#b45309" }} colSpan={2}>TOTAL</td>
                  {gcEnseignes.map((e, ei) => {
                    const tot = articlesValides.reduce((s, a) => s + (e.qties[qtyKey(a, articlesValides)] || 0), 0);
                    return <td key={ei} style={{ padding: "5px 4px", textAlign: "center", fontWeight: 800, color: "#b45309" }}>{fmtNum(tot)}</td>;
                  })}
                  <td />
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

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
