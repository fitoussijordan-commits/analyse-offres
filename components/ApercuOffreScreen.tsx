"use client";
import { useState, useEffect, useMemo } from "react";
import * as odoo from "@/lib/odoo";
import { loadCampagnesCreees, upsertCampagne } from "@/lib/campaigns";
import { CampagneCreee, GcEnseigne, GC_ENSEIGNES_DEFAUT, qtyParPack, totalPacks, ventilationPalier, toExportPayload, campagneCreeeToAnalyse } from "@/lib/create-campaign";
import {
  TYPOLOGIES, DEFAULT_PCTS, DEFAULT_REMISES, REMISE_ADD_DEFAUT,
  CalcPalier, calcPalier, calcSynthese, calcBesoinParRef, detailPalier, calcGrandsComptes,
} from "@/lib/calc-offre";
import { buildSyntheseLogistique } from "@/lib/logistique";

import { C } from "@/lib/theme";
const fmtEur = (n: number) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n || 0);
const fmtNum = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n || 0));
const fmtPct = (n: number) => `${(Math.round((n || 0) * 1000) / 10).toFixed(1)} %`;
// Prix : arrondi à 2 décimales (évite les artefacts flottants type 2.7600000000000002).
const fmtPrix = (n: number) => (Math.round((n || 0) * 100) / 100).toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtDate = (d: string) => { const m = (d || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : "—"; };

const input: React.CSSProperties = { padding: "4px 7px", border: `1px solid ${C.border}`, borderRadius: 6, fontSize: 12, fontFamily: "inherit", color: C.text, outline: "none", width: 70, textAlign: "right" };

interface Props {
  session: odoo.OdooSession;
  onToast: (msg: string, type?: "success" | "error" | "info") => void;
  onGoAnalyse?: () => void; // bascule vers l'outil Analyse des campagnes
}

interface PalierEdit {
  code: string; label: string; nbPacks: number; descriptif?: string;
  pcts: number[]; remises: number[]; remiseAdd: number;
  produits: { ref: string; name: string; barcode: string; qtyParPack: number; standardPrice: number; listPrice: number; ppc: number; typProd?: string }[];
}

function toPaliersEdit(camp: CampagneCreee): PalierEdit[] {
  const arts = camp.articles.filter(a => a.ref.trim());
  const totalP = totalPacks(camp.paliers);
  return camp.paliers.map(pal => {
    const vent = ventilationPalier(arts, pal);
    return {
    code: pal.code, label: pal.label, nbPacks: pal.nbPacks || 0, descriptif: pal.descriptif,
    pcts: (pal as any).pctOffresReco && (pal as any).pctOffresReco.length === 7 ? [...(pal as any).pctOffresReco] : [...DEFAULT_PCTS],
    // Remise unique du palier (Créer campagne) → répliquée sur les 7 typologies.
    // Dès qu'un taux est défini (>=0), on l'applique aux 7, même si le flag remiseStandard
    // n'a pas été (re)stocké. Sinon on garde les remises par défaut du gabarit.
    remises: pal.remiseStandardTaux != null ? new Array(7).fill(pal.remiseStandardTaux) : [...DEFAULT_REMISES],
    remiseAdd: pal.remiseAddTaux != null ? pal.remiseAddTaux : REMISE_ADD_DEFAUT,
    produits: arts.map(a => {
      // Réplique la gratuité : UG / Testeur / PLV / Échantillon → tarif de vente et PPC = 0
      // (aucun CA). Le coût reste réel. Même règle que l'export.
      const estVente = (a.typProd || "Produit Vente") === "Produit Vente";
      return {
        ref: a.ref.trim(), name: a.name || "", barcode: a.barcode || "",
        qtyParPack: qtyParPack(a, pal, totalP, vent, arts),
        standardPrice: a.standardPrice || 0,
        listPrice: estVente ? (a.listPrice || 0) : 0,
        ppc: estVente ? (a.ppc || 0) : 0,
        typProd: a.typProd || "Produit Vente",
      };
    }),
    };
  });
}

type SubTab = "offre" | "logistique" | "synthese";

export default function ApercuOffreScreen({ session, onToast, onGoAnalyse }: Props) {
  const [saved, setSaved] = useState<CampagneCreee[]>([]);
  const [camp, setCamp] = useState<CampagneCreee | null>(null);
  const [paliers, setPaliers] = useState<PalierEdit[]>([]);
  const [exporting, setExporting] = useState(false);
  const [tab, setTab] = useState<SubTab>("offre");
  // Grands Comptes : enseignes dynamiques. Template Excel = 6 colonnes max, UI = 10 max.
  const GC_MAX = 10;
  const [gcEnseignes, setGcEnseignes] = useState<GcEnseigne[]>(GC_ENSEIGNES_DEFAUT.map(e => ({ ...e, qties: {} })));

  useEffect(() => { void (async () => {
    try { const list = await loadCampagnesCreees(); setSaved(list); if (list.length) selectCamp(list[0]); }
    catch (e: any) { onToast("Erreur chargement : " + e.message, "error"); }
  })(); }, []);

  function selectCamp(c: CampagneCreee) {
    setCamp(c); setPaliers(toPaliersEdit(c));
    setGcEnseignes(c.gcEnseignes != null ? c.gcEnseignes.map(e => ({ ...e, qties: { ...e.qties } })) : GC_ENSEIGNES_DEFAUT.map(e => ({ ...e, qties: {} })));
  }
  const setGcQty = (ei: number, key: string, v: number) => setGcEnseignes(es => es.map((e, i) => i === ei ? { ...e, qties: { ...e.qties, [key]: v } } : e));
  const setGcNom = (ei: number, nom: string) => setGcEnseignes(es => es.map((e, i) => i === ei ? { ...e, nom } : e));
  const setGcRemise = (ei: number, remise: number) => setGcEnseignes(es => es.map((e, i) => i === ei ? { ...e, remise } : e));
  const addGc = () => setGcEnseignes(es => es.length >= GC_MAX ? es : [...es, { nom: `GC${es.length + 1}`, remise: 0, qties: {} }]);
  const removeGc = (ei: number) => setGcEnseignes(es => es.filter((_, i) => i !== ei));
  const nom = camp?.nom || "";

  // ── Mutations ──────────────────────────────────────────────────────────────
  const setPalier = (pi: number, patch: Partial<PalierEdit>) => setPaliers(ps => ps.map((p, i) => i === pi ? { ...p, ...patch } : p));
  const setPct = (pi: number, ti: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, pcts: p.pcts.map((x, j) => j === ti ? v : x) }));
  const setRemise = (pi: number, ti: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, remises: p.remises.map((x, j) => j === ti ? v : x) }));
  const setQty = (pi: number, ri: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, produits: p.produits.map((pr, j) => j === ri ? { ...pr, qtyParPack: v } : pr) }));

  // ── Calculs live ──────────────────────────────────────────────────────────
  const calcPaliers: CalcPalier[] = useMemo(() => paliers.map(p => ({
    code: p.code, label: p.label, nbPacks: p.nbPacks, pcts: p.pcts, remises: p.remises, remiseAdd: p.remiseAdd, produits: p.produits,
  })), [paliers]);
  const synthese = useMemo(() => calcSynthese(calcPaliers), [calcPaliers]);
  const besoin = useMemo(() => calcBesoinParRef(calcPaliers), [calcPaliers]);

  // CA / marge Grands Comptes (clé article : réf seule si unique, sinon réf#type).
  const gc = useMemo(() => {
    const pal1 = paliers[0];
    if (!pal1) return calcGrandsComptes([], []);
    const count: Record<string, number> = {};
    for (const p of pal1.produits) { const r = (p.ref || "").trim(); if (r) count[r] = (count[r] || 0) + 1; }
    const keyOf = (p: any) => { const r = (p.ref || "").trim(); return (r && count[r] > 1) ? `${r}#${p.typProd || "Produit Vente"}` : r; };
    const info = pal1.produits.map(p => ({
      key: keyOf(p),
      listPrice: (p.typProd || "Produit Vente") === "Produit Vente" ? (p.listPrice || 0) : 0,
      standardPrice: p.standardPrice || 0,
      remiseAdd: pal1.remiseAdd || 0,
    }));
    return calcGrandsComptes(gcEnseignes.map(e => ({ nom: e.nom, remise: e.remise, qties: e.qties })), info);
  }, [paliers, gcEnseignes]);

  // Besoin logistique par mois (réutilise la logique testée, avec les qtés éditées + dates campagne).
  const logistique = useMemo(() => {
    if (!camp) return null;
    const virtual: CampagneCreee = {
      ...camp,
      articles: (paliers[0]?.produits || []).map(p => ({ ref: p.ref, name: p.name, barcode: p.barcode })),
      paliers: paliers.map(p => ({
        code: p.code, label: p.label, nbPacks: p.nbPacks,
        qtyParPack: Object.fromEntries(p.produits.map(pr => [pr.ref, pr.qtyParPack])),
      })) as any,
    };
    return buildSyntheseLogistique([virtual]);
  }, [camp, paliers]);

  // ── Export ─────────────────────────────────────────────────────────────────
  const exporter = async () => {
    if (!paliers.length) { onToast("Rien à exporter", "error"); return; }
    setExporting(true);
    try {
      const payload: any = {
        nom,
        gcEnseignes,
        paliers: paliers.map(p => ({
          code: p.code, label: p.label, qtyPacks: p.nbPacks, descriptif: p.descriptif,
          pctOffres: p.pcts, remises: p.remises, remiseAddTaux: p.remiseAdd,
          produits: p.produits.map(pr => ({ ref: pr.ref, name: pr.name, productId: 0, qtyParPack: pr.qtyParPack, barcode: pr.barcode, standardPrice: pr.standardPrice, listPrice: pr.listPrice, ppc: pr.ppc, typProd: pr.typProd })),
        })),
      };
      if (logistique) payload.logistique = logistique;
      try {
        const catalogue = await odoo.getAllProducts(session);
        payload.mapping = catalogue.map(p => ({ ref: p.ref, name: p.name, barcode: p.barcode, standardPrice: p.standardPrice, listPrice: p.listPrice, ppc: p.ppc }));
      } catch { /* mapping limité */ }
      const res = await fetch("/api/export-template", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `proposition_${(nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Fichier Proposition exporté", "success");
    } catch (e: any) { onToast("Erreur export : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  // Gros export annuel : toutes les campagnes → 1 onglet/campagne + synthèse logistique cumulée
  // + synthèse CA annuelle. Utilise les campagnes SAUVEGARDÉES (pas les éditions non enregistrées).
  const [exportingAnnuel, setExportingAnnuel] = useState(false);
  const exporterAnnuel = async () => {
    if (!saved.length) { onToast("Aucune campagne à exporter", "error"); return; }
    setExportingAnnuel(true);
    try {
      const campagnes = saved.map(c => toExportPayload(c));
      const logistique = buildSyntheseLogistique(saved);
      const body: any = { campagnes, logistique };
      // Mapping catalogue (une seule fois pour tout le classeur).
      try {
        const catalogue = await odoo.getAllProducts(session);
        body.mapping = catalogue.map(p => ({ ref: p.ref, name: p.name, barcode: p.barcode, standardPrice: p.standardPrice, listPrice: p.listPrice, ppc: p.ppc }));
      } catch { /* mapping limité */ }
      const res = await fetch("/api/export-multi", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `campagnes_annee.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast(`Export annuel : ${saved.length} campagne(s)`, "success");
    } catch (e: any) { onToast("Erreur export annuel : " + e.message, "error"); }
    finally { setExportingAnnuel(false); }
  };

  // Valider la campagne → la rendre analysable dans « Analyse des campagnes » (suivi de progression).
  const validerPourAnalyse = async () => {
    if (!camp) return;
    try {
      const analyse = campagneCreeeToAnalyse(camp);
      await upsertCampagne(analyse); // enregistre dans la table campagnes (outil Analyse)
      onToast(`« ${camp.nom} » ajoutée à l'Analyse des campagnes`, "success");
      onGoAnalyse?.();
    } catch (e: any) { onToast("Erreur : " + e.message, "error"); }
  };

  if (!saved.length) return <div style={{ padding: 40, textAlign: "center", color: C.textMuted }}>Aucune campagne sauvegardée. Crée-en une dans « Créer une campagne » d'abord.</div>;

  const TABS: [SubTab, string][] = [["offre", "Offre"], ["logistique", "Besoin logistique"], ["synthese", "Synthèse détaillée"]];

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
    <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Aperçu interactif de l'offre</h1>
        <span style={{ fontSize: 13, color: C.textMuted }}>Édite les paramètres → tout se recalcule en direct.</span>
        <div style={{ flex: 1 }} />
        <select value={camp?.id || ""} onChange={e => { const c = saved.find(s => s.id === e.target.value); if (c) selectCamp(c); }} style={{ ...input, width: 240, textAlign: "left", fontSize: 13, padding: "7px 10px" }}>
          {saved.map(s => <option key={s.id} value={s.id}>{s.nom || "(sans nom)"}</option>)}
        </select>
        <button onClick={validerPourAnalyse} style={{ padding: "8px 16px", background: C.teal, border: "none", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit" }}>✓ Valider → suivre la progression</button>
        <button onClick={exporter} disabled={exporting} style={{ padding: "8px 16px", background: C.blue, border: "none", borderRadius: 8, cursor: exporting ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exporting ? 0.6 : 1 }}>{exporting ? "Export…" : "Exporter Excel"}</button>
        <button onClick={exporterAnnuel} disabled={exportingAnnuel} title="Toutes les campagnes : 1 onglet par campagne + synthèse logistique cumulée + synthèse CA annuelle" style={{ padding: "8px 16px", background: C.blueDark, border: "none", borderRadius: 8, cursor: exportingAnnuel ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exportingAnnuel ? 0.6 : 1 }}>{exportingAnnuel ? "Export…" : "📚 Export annuel (toutes campagnes)"}</button>
      </div>

      {/* Bannière : nom + dates de campagne */}
      {camp && (
        <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap", background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 18px", boxShadow: C.shadow }}>
          <span style={{ fontSize: 18, fontWeight: 800, color: C.text }}>{camp.nom || "(sans nom)"}</span>
          {camp.annee && <span style={{ fontSize: 12, fontWeight: 700, color: C.blueDark, background: C.blueSoft, borderRadius: 6, padding: "3px 10px" }}>{camp.annee}</span>}
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 13, color: C.textMuted }}>📅 Campagne</span>
          <span style={{ fontSize: 15, fontWeight: 700, color: C.text }}>{fmtDate(camp.dateDebut)} → {fmtDate(camp.dateFin)}</span>
        </div>
      )}

      {/* KPI synthèse globale */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
        {gc.caTotal > 0 ? (
          <>
            <Kpi label="CA Retail + Institut" value={fmtEur(synthese.caTotal)} color={C.teal} />
            <Kpi label="CA Grands Comptes" value={fmtEur(gc.caTotal)} color={C.blue} />
            <Kpi label="CA total (RI + GC)" value={fmtEur(synthese.caTotal + gc.caTotal)} color={C.blueDark} />
            <Kpi label="Marge totale (RI + GC)" value={fmtEur(synthese.margeTotal + gc.margeTotal)} color={C.green} />
            <Kpi label="Marge % globale" value={fmtPct((synthese.caTotal + gc.caTotal) > 0 ? (synthese.margeTotal + gc.margeTotal) / (synthese.caTotal + gc.caTotal) : 0)} color={C.green} />
          </>
        ) : (
          <>
            <Kpi label="CA total campagne" value={fmtEur(synthese.caTotal)} color={C.blue} />
            <Kpi label="Marge totale" value={fmtEur(synthese.margeTotal)} color={C.teal} />
            <Kpi label="Marge %" value={fmtPct(synthese.margePct)} color={C.green} />
            <Kpi label="Nb offres (tous paliers)" value={fmtNum(synthese.nbPacks)} color={C.amber} />
          </>
        )}
      </div>

      {/* Sous-onglets */}
      <div style={{ display: "flex", gap: 6, borderBottom: `1px solid ${C.border}` }}>
        {TABS.map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ padding: "8px 16px", border: "none", background: "transparent", cursor: "pointer", fontSize: 13, fontWeight: 700, fontFamily: "inherit", color: tab === id ? C.blueDark : C.textMuted, borderBottom: tab === id ? `2px solid ${C.blue}` : "2px solid transparent" }}>{label}</button>
        ))}
      </div>

      {tab === "offre" && <>
        <OffreTab paliers={paliers} calcPaliers={calcPaliers} setPalier={setPalier} setPct={setPct} setRemise={setRemise} setQty={setQty} besoin={besoin} />
        <GrandsComptesBloc produits={paliers[0]?.produits || []} enseignes={gcEnseignes} gcMax={GC_MAX} setGcQty={setGcQty} setGcNom={setGcNom} setGcRemise={setGcRemise} addGc={addGc} removeGc={removeGc} />
      </>}
      {tab === "logistique" && <LogistiqueTab log={logistique} />}
      {tab === "synthese" && <SyntheseTab paliers={paliers} calcPaliers={calcPaliers} gc={gc} />}

      <div style={{ fontSize: 12, color: C.textMuted, fontStyle: "italic", paddingBottom: 20 }}>
        Aperçu en lecture/édition — les modifications ne sont pas sauvegardées. Utilise « Exporter Excel » pour récupérer le fichier avec tes valeurs.
      </div>
    </div>
    </div>
  );
}

// ── Onglet OFFRE ──────────────────────────────────────────────────────────────
function OffreTab({ paliers, calcPaliers, setPalier, setPct, setRemise, setQty, besoin }: any) {
  const [openDetail, setOpenDetail] = useState<number | null>(null);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {paliers.map((pal: PalierEdit, pi: number) => {
        const r = calcPalier(calcPaliers[pi]);
        const d = detailPalier(calcPaliers[pi]);
        return (
          <div key={pi} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
            <div style={{ padding: "12px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <span style={{ fontSize: 11, fontWeight: 800, fontFamily: "monospace", background: C.blue, color: "#fff", borderRadius: 5, padding: "2px 8px" }}>{pal.code}</span>
              <span style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{pal.label}</span>
              <span style={{ fontSize: 12, color: C.textMuted }}>Nb offres</span>
              <input type="number" style={{ ...input, width: 80 }} value={pal.nbPacks || ""} onChange={e => setPalier(pi, { nbPacks: parseInt(e.target.value) || 0 })} />
              <span style={{ fontSize: 12, color: C.textMuted }}>Remise add.</span>
              <input type="number" step="0.1" style={{ ...input, width: 60 }} value={Math.round(pal.remiseAdd * 1000) / 10} onChange={e => setPalier(pi, { remiseAdd: (parseFloat(e.target.value) || 0) / 100 })} />
              <span style={{ fontSize: 12, color: C.textMuted }}>%</span>
              <div style={{ flex: 1 }} />
              <span style={{ fontSize: 13, fontWeight: 700, color: C.blue }}>CA {fmtEur(r.caTotal)}</span>
              <span style={{ fontSize: 13, fontWeight: 700, color: C.green }}>Marge {fmtEur(r.margeTotal)} ({fmtPct(r.margePct)})</span>
              <button onClick={() => setOpenDetail(openDetail === pi ? null : pi)}
                style={{ padding: "4px 10px", background: openDetail === pi ? C.blueSoft : "transparent", border: `1px solid ${openDetail === pi ? C.blue : C.border}`, borderRadius: 6, cursor: "pointer", color: openDetail === pi ? C.blueDark : C.textMuted, fontSize: 11.5, fontWeight: 600, fontFamily: "inherit" }}>
                Détail calcul
              </button>
            </div>

            {openDetail === pi && (
              <div style={{ background: C.bg, borderBottom: `1px solid ${C.border}`, padding: "12px 16px", display: "flex", flexDirection: "column", gap: 10 }}>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: C.blueDark, textTransform: "uppercase", letterSpacing: "0.06em" }}>Détail des calculs — {pal.label || pal.code}</div>

                <div style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
                  {[
                    { l: "CA par produit", v: "qté/offre × tarif × (1 − remise add.) × nb offres typo × (1 − remise typo)", mono: true },
                    { l: "Marge par produit", v: "CA − (qté/offre × coût × nb offres typo)", mono: true },
                    { l: "Nb offres typo", v: `% offres × ${fmtNum(pal.nbPacks)} offres du palier`, mono: true },
                  ].map((x, i) => (
                    <div key={i} style={{ display: "flex", gap: 10, alignItems: "baseline" }}>
                      <span style={{ minWidth: 130, color: C.textMuted, fontWeight: 600, flexShrink: 0 }}>{x.l}</span>
                      <span style={{ color: C.text, fontFamily: "ui-monospace, monospace", fontSize: 11.5 }}>{x.v}</span>
                    </div>
                  ))}
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 8 }}>
                  {[
                    { l: "Remise moy. pondérée", v: fmtPct(d.remiseMoyenne), n: "par nb d'offres de chaque typo" },
                    { l: "Coût d'achat total", v: fmtEur(d.coutTotal) },
                    { l: "CA / offre", v: fmtEur(d.caParOffre) },
                    { l: "Marge / offre", v: fmtEur(d.margeParOffre) },
                    { l: "Coef. multiplicateur", v: d.coutTotal > 0 ? (d.caTotal / d.coutTotal).toFixed(2) + " ×" : "—", n: "CA ÷ coût" },
                  ].map((k, i) => (
                    <div key={i} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 7, padding: "7px 10px" }}>
                      <div style={{ fontSize: 9.5, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600 }}>{k.l}</div>
                      <div style={{ fontSize: 14, fontWeight: 600, color: C.text, fontVariantNumeric: "tabular-nums", marginTop: 1 }}>{k.v}</div>
                      {k.n && <div style={{ fontSize: 9.5, color: C.textMuted, marginTop: 1 }}>{k.n}</div>}
                    </div>
                  ))}
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5, fontVariantNumeric: "tabular-nums" }}>
                    <thead>
                      <tr>
                        {["Réf", "Qté/offre", "Unités", "Tarif", "Prix net moy.", "Coût", "CA", "Marge", "Marge %", "Part CA"].map((h, hi) => (
                          <th key={hi} style={{ padding: "4px 8px", fontSize: 9.5, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", textAlign: hi === 0 ? "left" : "right", borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {d.produits.filter(p => p.qtyParPack > 0).map((p, li) => (
                        <tr key={li} style={{ background: li % 2 ? C.white : "transparent" }}>
                          <td style={{ padding: "3px 8px", fontFamily: "ui-monospace, monospace", color: C.textSec, whiteSpace: "nowrap" }}>{p.ref}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", fontWeight: 600 }}>{p.qtyParPack}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted }}>{fmtNum(p.unitesTotal)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted }}>{fmtPrix(p.listPrice)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.textSec }}>{fmtPrix(p.prixNetMoyen)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted }}>{fmtPrix(p.standardPrice)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.blueDark, fontWeight: 600 }}>{fmtEur(p.ca)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: p.marge >= 0 ? C.green : C.red, fontWeight: 600 }}>{fmtEur(p.marge)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: p.margePct >= 0.3 ? C.green : C.amber }}>{fmtPct(p.margePct)}</td>
                          <td style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted }}>{fmtPct(p.partCa)}</td>
                        </tr>
                      ))}
                      <tr>
                        <td colSpan={6} style={{ padding: "5px 8px", borderTop: `1px solid ${C.borderDark}`, color: C.textMuted, fontWeight: 600 }}>Total palier</td>
                        <td style={{ padding: "5px 8px", textAlign: "right", borderTop: `1px solid ${C.borderDark}`, color: C.blueDark, fontWeight: 700 }}>{fmtEur(d.caTotal)}</td>
                        <td style={{ padding: "5px 8px", textAlign: "right", borderTop: `1px solid ${C.borderDark}`, color: C.green, fontWeight: 700 }}>{fmtEur(d.margeTotal)}</td>
                        <td style={{ padding: "5px 8px", textAlign: "right", borderTop: `1px solid ${C.borderDark}`, color: C.green, fontWeight: 700 }}>{fmtPct(d.margePct)}</td>
                        <td style={{ padding: "5px 8px", borderTop: `1px solid ${C.borderDark}` }} />
                      </tr>
                    </tbody>
                  </table>
                </div>
                <div style={{ fontSize: 10.5, color: C.textMuted }}>
                  « Prix net moy. » = tarif × (1 − remise add. {fmtPct(pal.remiseAdd)}) × (1 − remise moy. {fmtPct(d.remiseMoyenne)}). « Unités » = qté/offre × {fmtNum(d.nbOffresTotal)} offres réparties.
                </div>
              </div>
            )}
            <div style={{ padding: "8px 16px 0", display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 12, color: C.textMuted, whiteSpace: "nowrap" }}>Descriptif</span>
              <input style={{ ...input, flex: 1, textAlign: "left", width: "auto" }} value={pal.descriptif ?? ""} onChange={e => setPalier(pi, { descriptif: e.target.value })} placeholder="Texte retranscrit dans l'Excel à côté du nom du palier" />
            </div>
            <div style={{ padding: "8px 16px", overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: "left", padding: "4px 8px", color: C.textMuted, borderBottom: `1px solid ${C.border}` }}></th>
                    {TYPOLOGIES.map(t => <th key={t} style={{ padding: "4px 8px", color: C.textSec, fontWeight: 700, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{t}</th>)}
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>% Offres</td>
                    {pal.pcts.map((v: number, ti: number) => <td key={ti} style={{ padding: "3px 4px", textAlign: "center" }}>
                      <input type="number" step="1" style={{ ...input, width: 52, textAlign: "center" }} value={Math.round(v * 1000) / 10} onChange={e => setPct(pi, ti, (parseFloat(e.target.value) || 0) / 100)} />
                    </td>)}
                  </tr>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>Remise</td>
                    {pal.remises.map((v: number, ti: number) => <td key={ti} style={{ padding: "3px 4px", textAlign: "center" }}>
                      <input type="number" step="0.1" style={{ ...input, width: 52, textAlign: "center" }} value={Math.round(v * 1000) / 10} onChange={e => setRemise(pi, ti, (parseFloat(e.target.value) || 0) / 100)} />
                    </td>)}
                  </tr>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>Nb offres</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", color: C.textSec }}>{fmtNum(tr.nbOffres)}</td>)}
                  </tr>
                  <tr style={{ background: C.blueSoft }}>
                    <td style={{ padding: "4px 8px", color: C.blueDark, fontWeight: 700 }}>CA</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", fontWeight: 600, color: C.text, whiteSpace: "nowrap" }}>{fmtEur(tr.ca)}</td>)}
                  </tr>
                  <tr style={{ background: C.tealSoft }}>
                    <td style={{ padding: "4px 8px", color: C.teal, fontWeight: 700 }}>Marge</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", fontWeight: 600, color: tr.marge >= 0 ? C.text : C.red, whiteSpace: "nowrap" }}>{fmtEur(tr.marge)}</td>)}
                  </tr>
                </tbody>
              </table>
            </div>
            <div style={{ padding: "0 16px 14px", overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>{["Réf", "Produit", "EAN", "Qté/offre", "Coût", "Tarif", "PPC"].map((h, i) => <th key={i} style={{ padding: "5px 8px", textAlign: i >= 3 ? "right" : "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr>
                </thead>
                <tbody>
                  {pal.produits.map((p: any, ri: number) => (
                    <tr key={ri}>
                      <td style={{ padding: "4px 8px", fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.ref}</td>
                      <td style={{ padding: "4px 8px", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.name}</td>
                      <td style={{ padding: "4px 8px", color: C.textMuted, fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.barcode || "—"}</td>
                      <td style={{ padding: "3px 8px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>
                        <input type="number" style={{ ...input, width: 60, fontWeight: 700, color: C.blue }} value={p.qtyParPack || ""} onChange={e => setQty(pi, ri, parseInt(e.target.value) || 0)} />
                      </td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{fmtPrix(p.standardPrice)}</td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{fmtPrix(p.listPrice)}</td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{fmtPrix(p.ppc)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── Bloc GRANDS COMPTES : mêmes articles × enseignes dynamiques (max 6, limite template) ──
function GrandsComptesBloc({ produits, enseignes, gcMax, setGcQty, setGcNom, setGcRemise, addGc, removeGc }: {
  produits: any[]; enseignes: GcEnseigne[]; gcMax: number;
  setGcQty: (ei: number, key: string, v: number) => void;
  setGcNom: (ei: number, nom: string) => void;
  setGcRemise: (ei: number, remise: number) => void;
  addGc: () => void;
  removeGc: (ei: number) => void;
}) {
  if (!produits.length) return null;
  // Clé composite : distingue les doublons de réf (stick vendu vs stick UG).
  const refCount: Record<string, number> = {};
  for (const p of produits) { const r = (p.ref || "").trim(); if (r) refCount[r] = (refCount[r] || 0) + 1; }
  const keyOf = (p: any) => { const r = (p.ref || "").trim(); return (r && refCount[r] > 1) ? `${r}#${p.typProd || "Produit Vente"}` : r; };
  return (
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
      <div style={{ padding: "12px 16px", background: "#fff7ed", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ fontSize: 11, fontWeight: 800, fontFamily: "monospace", background: "#b45309", color: "#fff", borderRadius: 5, padding: "2px 8px" }}>GRANDS COMPTES</span>
        <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>Quantités par enseigne (nom et remise éditables)</span>
        <span style={{ fontSize: 11, color: C.textMuted, marginLeft: 4 }}>{enseignes.length} enseigne{enseignes.length > 1 ? "s" : ""}</span>
      </div>
      <div style={{ padding: "0 16px 14px", overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            {/* Ligne 1 : noms d'enseignes éditables + bouton + */}
            <tr>
              <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }} colSpan={2}>Enseigne →</th>
              {enseignes.map((e, ei) => (
                <th key={ei} style={{ padding: "4px 4px", borderBottom: `1px solid ${C.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
                    <input style={{ ...input, width: 80, textAlign: "center", fontWeight: 700, color: "#b45309" }} value={e.nom} onChange={ev => setGcNom(ei, ev.target.value)} />
                    <button onClick={() => removeGc(ei)} title="Supprimer cette enseigne" style={{ padding: "2px 5px", background: "none", border: "1px solid #fca5a5", borderRadius: 5, cursor: "pointer", color: "#ef4444", fontSize: 11, fontWeight: 700, lineHeight: 1 }}>×</button>
                  </div>
                </th>
              ))}
              <th style={{ padding: "4px 4px", borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" }}>
                <button onClick={addGc} disabled={enseignes.length >= gcMax} title={enseignes.length >= gcMax ? `Maximum ${gcMax} enseignes (limite du template)` : "Ajouter une enseigne GC"} style={{ padding: "4px 8px", background: enseignes.length >= gcMax ? C.border : "#fff7ed", border: "1px dashed #b45309", borderRadius: 6, cursor: enseignes.length >= gcMax ? "default" : "pointer", color: enseignes.length >= gcMax ? C.textMuted : "#b45309", fontSize: 13, fontWeight: 700, opacity: enseignes.length >= gcMax ? 0.4 : 1 }}>+</button>
              </th>
            </tr>
            {/* Ligne 2 : remises éditables (%) */}
            <tr>
              <th style={{ padding: "3px 8px", textAlign: "right", color: C.textMuted, fontWeight: 600, borderBottom: `1px solid ${C.border}` }} colSpan={2}>Remise %</th>
              {enseignes.map((e, ei) => (
                <th key={ei} style={{ padding: "3px 4px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                  <input type="number" step="0.1" style={{ ...input, width: 60, textAlign: "center" }} value={Math.round(e.remise * 1000) / 10} onChange={ev => setGcRemise(ei, (parseFloat(ev.target.value) || 0) / 100)} />
                </th>
              ))}
              <th style={{ borderBottom: `1px solid ${C.border}` }} />
            </tr>
            {/* Ligne 3 : entêtes fixes */}
            <tr>
              <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Réf</th>
              <th style={{ padding: "5px 8px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Produit</th>
              {enseignes.map((e, ei) => <th key={ei} style={{ padding: "5px 4px", textAlign: "center", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Qté</th>)}
              <th style={{ borderBottom: `1px solid ${C.border}` }} />
            </tr>
          </thead>
          <tbody>
            {produits.map((p: any, ri: number) => {
              const k = keyOf(p);
              return (
                <tr key={ri}>
                  <td style={{ padding: "4px 8px", fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.ref}</td>
                  <td style={{ padding: "4px 8px", color: C.textSec, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{p.name}</td>
                  {enseignes.map((e, ei) => (
                    <td key={ei} style={{ padding: "3px 4px", textAlign: "center", borderBottom: `1px solid ${C.border}` }}>
                      <input type="number" style={{ ...input, width: 60, textAlign: "center", color: "#b45309" }} value={e.qties[k] || ""} onChange={ev => setGcQty(ei, k, parseInt(ev.target.value) || 0)} placeholder="0" />
                    </td>
                  ))}
                  <td style={{ borderBottom: `1px solid ${C.border}` }} />
                </tr>
              );
            })}
            {/* Ligne total par enseigne */}
            <tr style={{ background: "#fff7ed" }}>
              <td style={{ padding: "5px 8px", fontWeight: 800, color: "#b45309" }} colSpan={2}>TOTAL</td>
              {enseignes.map((e, ei) => {
                const tot = produits.reduce((s, p) => s + (e.qties[keyOf(p)] || 0), 0);
                return <td key={ei} style={{ padding: "5px 4px", textAlign: "center", fontWeight: 800, color: "#b45309" }}>{fmtNum(tot)}</td>;
              })}
              <td />
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── Onglet BESOIN LOGISTIQUE (réf × mois) ────────────────────────────────────
function LogistiqueTab({ log }: { log: any }) {
  if (!log || !log.lignes?.length) return <div style={{ padding: 30, textAlign: "center", color: C.textMuted }}>Renseigne les dates de campagne (dans « Créer une campagne ») pour calculer le planning par mois.</div>;
  return (
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", boxShadow: C.shadow }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: C.tealSoft }}>
            <th style={{ padding: "8px 10px", textAlign: "left", color: C.teal, fontWeight: 700, borderBottom: `1px solid ${C.border}`, position: "sticky", left: 0, background: C.tealSoft }}>Réf</th>
            <th style={{ padding: "8px 10px", textAlign: "left", color: C.teal, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>Produit</th>
            {log.moisLabels.map((m: string) => <th key={m} style={{ padding: "8px 10px", textAlign: "right", color: C.textSec, fontWeight: 700, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{m}</th>)}
            <th style={{ padding: "8px 10px", textAlign: "right", color: C.teal, fontWeight: 800, borderBottom: `1px solid ${C.border}` }}>Total</th>
          </tr>
        </thead>
        <tbody>
          {log.lignes.map((l: any) => (
            <tr key={l.ref}>
              <td style={{ padding: "6px 10px", fontFamily: "monospace", borderBottom: `1px solid ${C.border}`, position: "sticky", left: 0, background: C.white }}>{l.ref}</td>
              <td style={{ padding: "6px 10px", color: C.textSec, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{l.name}</td>
              {l.parMois.map((q: number, i: number) => <td key={i} style={{ padding: "6px 10px", textAlign: "right", color: q ? C.text : C.textMuted, borderBottom: `1px solid ${C.border}` }}>{q ? fmtNum(q) : "—"}</td>)}
              <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: 700, color: C.teal, borderBottom: `1px solid ${C.border}` }}>{fmtNum(l.total)}</td>
            </tr>
          ))}
          <tr style={{ background: C.tealSoft }}>
            <td style={{ padding: "8px 10px", fontWeight: 800, color: C.teal, position: "sticky", left: 0, background: C.tealSoft }}>TOTAL</td>
            <td style={{ background: C.tealSoft }} />
            {log.totalParMois.map((q: number, i: number) => <td key={i} style={{ padding: "8px 10px", textAlign: "right", fontWeight: 700, color: C.teal }}>{fmtNum(q)}</td>)}
            <td style={{ padding: "8px 10px", textAlign: "right", fontWeight: 800, color: C.teal }}>{fmtNum(log.totalGeneral)}</td>
          </tr>
        </tbody>
      </table>
      <div style={{ padding: "10px 14px", fontSize: 11, color: C.textMuted, fontStyle: "italic" }}>Profil : 40 % le mois précédant le début de l'offre, puis 60 % lissé jusqu'à 1 mois avant la fin. Les mois débordent sur l'année suivante si l'offre l'exige.</div>
    </div>
  );
}

// ── Onglet SYNTHÈSE DÉTAILLÉE (CA/Marge par offre + par statut) ──────────────
function SyntheseTab({ paliers, calcPaliers, gc }: { paliers: PalierEdit[]; calcPaliers: CalcPalier[]; gc: ReturnType<typeof calcGrandsComptes> }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      {/* CA/Marge par offre */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
        <div style={{ padding: "10px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, fontSize: 13, fontWeight: 800, color: C.blueDark }}>CA / Marge par offre</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr>{["Offre", "Nb offres", "CA", "Marge €", "Marge %"].map((h, i) => <th key={i} style={{ padding: "8px 14px", textAlign: i >= 1 ? "right" : "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
          <tbody>
            {paliers.map((pal, pi) => {
              const r = calcPalier(calcPaliers[pi]);
              return (
                <tr key={pi}>
                  <td style={{ padding: "7px 14px", fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>{pal.code} — {pal.label}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtNum(pal.nbPacks)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", fontWeight: 600, color: C.blue, borderBottom: `1px solid ${C.border}` }}>{fmtEur(r.caTotal)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", fontWeight: 600, color: C.teal, borderBottom: `1px solid ${C.border}` }}>{fmtEur(r.margeTotal)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtPct(r.margePct)}</td>
                </tr>
              );
            })}
            {(() => {
              const caRI = calcPaliers.reduce((s, c) => s + calcPalier(c).caTotal, 0);
              const mgRI = calcPaliers.reduce((s, c) => s + calcPalier(c).margeTotal, 0);
              const nbTot = paliers.reduce((s, p) => s + (p.nbPacks || 0), 0);
              const totRow = (label: string, nb: string, ca: number, mg: number, color: string, bg: string, top = false) => (
                <tr style={{ background: bg }}>
                  <td style={{ padding: "8px 14px", fontWeight: 800, color, borderTop: top ? `2px solid ${C.borderDark}` : undefined }}>{label}</td>
                  <td style={{ padding: "8px 14px", textAlign: "right", color, fontWeight: 700, borderTop: top ? `2px solid ${C.borderDark}` : undefined }}>{nb}</td>
                  <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 800, color, borderTop: top ? `2px solid ${C.borderDark}` : undefined }}>{fmtEur(ca)}</td>
                  <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 800, color, borderTop: top ? `2px solid ${C.borderDark}` : undefined }}>{fmtEur(mg)}</td>
                  <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 700, color, borderTop: top ? `2px solid ${C.borderDark}` : undefined }}>{fmtPct(ca > 0 ? mg / ca : 0)}</td>
                </tr>
              );
              return <>
                {totRow("Total Retail + Institut", fmtNum(nbTot), caRI, mgRI, C.teal, C.tealSoft, true)}
                {gc.caTotal > 0 && totRow("Total Grands Comptes", "—", gc.caTotal, gc.margeTotal, C.blueDark, C.blueSoft)}
                {gc.caTotal > 0 && totRow("Total général (RI + GC)", "—", caRI + gc.caTotal, mgRI + gc.margeTotal, C.text, C.bg)}
              </>;
            })()}
          </tbody>
        </table>
      </div>

      {/* Détail Grands Comptes par enseigne */}
      {gc.parEnseigne.some(e => e.ca > 0) && (
        <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
          <div style={{ padding: "10px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, fontSize: 13, fontWeight: 800, color: C.blueDark }}>Grands Comptes par enseigne</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr>{["Enseigne", "Qté", "CA", "Marge €", "Marge %"].map((h, i) => <th key={i} style={{ padding: "8px 14px", textAlign: i >= 1 ? "right" : "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
            <tbody>
              {gc.parEnseigne.filter(e => e.ca > 0).map((e, i) => (
                <tr key={i}>
                  <td style={{ padding: "7px 14px", fontWeight: 600, borderBottom: `1px solid ${C.border}` }}>{e.nom}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtNum(e.qty)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", fontWeight: 600, color: C.blue, borderBottom: `1px solid ${C.border}` }}>{fmtEur(e.ca)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", fontWeight: 600, color: C.green, borderBottom: `1px solid ${C.border}` }}>{fmtEur(e.marge)}</td>
                  <td style={{ padding: "7px 14px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>{fmtPct(e.ca > 0 ? e.marge / e.ca : 0)}</td>
                </tr>
              ))}
              <tr style={{ background: C.blueSoft }}>
                <td style={{ padding: "8px 14px", fontWeight: 800, color: C.blueDark }}>Total GC</td>
                <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 700, color: C.blueDark }}>{fmtNum(gc.qtyTotal)}</td>
                <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 800, color: C.blueDark }}>{fmtEur(gc.caTotal)}</td>
                <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 800, color: C.blueDark }}>{fmtEur(gc.margeTotal)}</td>
                <td style={{ padding: "8px 14px", textAlign: "right", fontWeight: 700, color: C.blueDark }}>{fmtPct(gc.margePct)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      {/* Ventilation par statut (typologie) : nb offres, CA, marge — agrégés tous paliers */}
      <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "auto", boxShadow: C.shadow }}>
        <div style={{ padding: "10px 16px", background: C.tealSoft, borderBottom: `1px solid ${C.border}`, fontSize: 13, fontWeight: 800, color: C.teal }}>Ventilation par statut client (tous paliers)</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead><tr>
            <th style={{ padding: "6px 10px", textAlign: "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}></th>
            {TYPOLOGIES.map(t => <th key={t} style={{ padding: "6px 10px", textAlign: "right", color: C.textSec, fontWeight: 700, borderBottom: `1px solid ${C.border}`, whiteSpace: "nowrap" }}>{t}</th>)}
            <th style={{ padding: "6px 10px", textAlign: "right", color: C.text, fontWeight: 800, borderBottom: `1px solid ${C.border}` }}>Total</th>
          </tr></thead>
          <tbody>
            {(() => {
              // Agrège par typologie sur tous les paliers.
              const nbOff = new Array(7).fill(0), ca = new Array(7).fill(0), marge = new Array(7).fill(0);
              for (const cp of calcPaliers) {
                const r = calcPalier(cp);
                for (let t = 0; t < 7; t++) { nbOff[t] += r.parTypo[t].nbOffres; ca[t] += r.parTypo[t].ca; marge[t] += r.parTypo[t].marge; }
              }
              const row = (label: string, arr: number[], fmt: (n: number) => string, color: string, bg?: string) => (
                <tr style={{ background: bg }}>
                  <td style={{ padding: "6px 10px", fontWeight: 700, color }}>{label}</td>
                  {arr.map((v, i) => <td key={i} style={{ padding: "6px 10px", textAlign: "right", whiteSpace: "nowrap" }}>{fmt(v)}</td>)}
                  <td style={{ padding: "6px 10px", textAlign: "right", fontWeight: 800 }}>{fmt(arr.reduce((s, x) => s + x, 0))}</td>
                </tr>
              );
              return <>
                {row("Nb offres", nbOff, fmtNum, C.textSec)}
                {row("CA", ca, fmtEur, C.blueDark, C.blueSoft)}
                {row("Marge", marge, fmtEur, C.teal, C.tealSoft)}
              </>;
            })()}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Kpi({ label, value, color }: { label: string; value: string; color: string }) {
  return (
    <div style={{ flex: "1 1 180px", background: C.white, border: `1px solid ${C.border}`, borderLeft: `3px solid ${color}`, borderRadius: 10, padding: "12px 16px", boxShadow: C.shadow }}>
      <div style={{ fontSize: 11, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 700 }}>{label}</div>
      <div style={{ fontSize: 20, fontWeight: 800, color, marginTop: 3 }}>{value}</div>
    </div>
  );
}
