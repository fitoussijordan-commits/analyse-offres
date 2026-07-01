"use client";
import { useState, useEffect, useMemo } from "react";
import * as odoo from "@/lib/odoo";
import { loadCampagnesCreees } from "@/lib/campaigns";
import { CampagneCreee, qtyParPack } from "@/lib/create-campaign";
import {
  TYPOLOGIES, DEFAULT_PCTS, DEFAULT_REMISES, REMISE_ADD_DEFAUT,
  CalcPalier, calcPalier, calcSynthese, calcBesoinParRef,
} from "@/lib/calc-offre";

const C = {
  bg: "#f1f5f9", white: "#ffffff",
  text: "#0f172a", textSec: "#334155", textMuted: "#64748b",
  border: "#e2e8f0",
  blue: "#3b82f6", blueDark: "#1d4ed8", blueSoft: "#eff6ff",
  green: "#10b981", greenSoft: "#ecfdf5",
  amber: "#f59e0b", amberSoft: "#fffbeb",
  red: "#ef4444",
  teal: "#0d9488", tealSoft: "#f0fdfa",
  shadow: "0 1px 3px rgba(0,0,0,0.06)",
};
const fmtEur = (n: number) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n || 0);
const fmtNum = (n: number) => new Intl.NumberFormat("fr-FR").format(Math.round(n || 0));
const fmtPct = (n: number) => `${(Math.round((n || 0) * 1000) / 10).toFixed(1)} %`;

const input: React.CSSProperties = { padding: "4px 7px", border: `1px solid ${C.border}`, borderRadius: 6, fontSize: 12, fontFamily: "inherit", color: C.text, outline: "none", width: 70, textAlign: "right" };

interface Props {
  session: odoo.OdooSession;
  onToast: (msg: string, type?: "success" | "error" | "info") => void;
}

// État éditable d'un palier dans l'aperçu (paramètres pilotables).
interface PalierEdit {
  code: string; label: string; nbPacks: number;
  pcts: number[]; remises: number[]; remiseAdd: number;
  produits: { ref: string; name: string; barcode: string; qtyParPack: number; standardPrice: number; listPrice: number; ppc: number }[];
}

// Convertit une CampagneCreee (sauvegardée) en paliers éditables pour l'aperçu.
function toPaliersEdit(camp: CampagneCreee): PalierEdit[] {
  const arts = camp.articles.filter(a => a.ref.trim());
  return camp.paliers.map(pal => ({
    code: pal.code, label: pal.label, nbPacks: pal.nbPacks || 0,
    pcts: (pal as any).pctOffresReco && (pal as any).pctOffresReco.length === 7 ? [...(pal as any).pctOffresReco] : [...DEFAULT_PCTS],
    remises: pal.remiseStandard && pal.remiseStandardTaux != null ? new Array(7).fill(pal.remiseStandardTaux) : [...DEFAULT_REMISES],
    remiseAdd: pal.remiseAddTaux != null ? pal.remiseAddTaux : REMISE_ADD_DEFAUT,
    produits: arts.map(a => ({
      ref: a.ref.trim(), name: a.name || "", barcode: a.barcode || "",
      qtyParPack: qtyParPack(a, pal),
      standardPrice: a.standardPrice || 0, listPrice: a.listPrice || 0, ppc: a.ppc || 0,
    })),
  }));
}

export default function ApercuOffreScreen({ session, onToast }: Props) {
  const [saved, setSaved] = useState<CampagneCreee[]>([]);
  const [campId, setCampId] = useState<string>("");
  const [paliers, setPaliers] = useState<PalierEdit[]>([]);
  const [nom, setNom] = useState("");
  const [exporting, setExporting] = useState(false);

  useEffect(() => { void (async () => {
    try { const list = await loadCampagnesCreees(); setSaved(list); if (list.length && !campId) selectCamp(list[0], list); }
    catch (e: any) { onToast("Erreur chargement : " + e.message, "error"); }
  })(); }, []);

  function selectCamp(c: CampagneCreee, list = saved) {
    setCampId(c.id); setNom(c.nom); setPaliers(toPaliersEdit(c));
  }

  // ── Mutations (édition live) ──────────────────────────────────────────────
  const setPalier = (pi: number, patch: Partial<PalierEdit>) => setPaliers(ps => ps.map((p, i) => i === pi ? { ...p, ...patch } : p));
  const setPct = (pi: number, ti: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, pcts: p.pcts.map((x, j) => j === ti ? v : x) }));
  const setRemise = (pi: number, ti: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, remises: p.remises.map((x, j) => j === ti ? v : x) }));
  const setQty = (pi: number, ri: number, v: number) => setPaliers(ps => ps.map((p, i) => i !== pi ? p : { ...p, produits: p.produits.map((pr, j) => j === ri ? { ...pr, qtyParPack: v } : pr) }));

  // ── Calculs live ──────────────────────────────────────────────────────────
  const calcPaliers: CalcPalier[] = useMemo(() => paliers.map(p => ({
    code: p.code, label: p.label, nbPacks: p.nbPacks, pcts: p.pcts, remises: p.remises, remiseAdd: p.remiseAdd,
    produits: p.produits,
  })), [paliers]);
  const synthese = useMemo(() => calcSynthese(calcPaliers), [calcPaliers]);
  const besoin = useMemo(() => calcBesoinParRef(calcPaliers), [calcPaliers]);

  // ── Export Excel depuis l'aperçu (avec les valeurs éditées) ────────────────
  const exporter = async () => {
    if (!paliers.length) { onToast("Rien à exporter", "error"); return; }
    setExporting(true);
    try {
      const payload: any = {
        nom,
        paliers: paliers.map(p => ({
          code: p.code, label: p.label, qtyPacks: p.nbPacks,
          pctOffres: p.pcts,                    // % offres édités
          remises: p.remises,                   // remises éditées (par typologie)
          remiseAddTaux: p.remiseAdd,
          produits: p.produits.map(pr => ({
            ref: pr.ref, name: pr.name, productId: 0, qtyParPack: pr.qtyParPack,
            barcode: pr.barcode, standardPrice: pr.standardPrice, listPrice: pr.listPrice, ppc: pr.ppc,
          })),
        })),
      };
      try {
        const catalogue = await odoo.getAllProducts(session);
        payload.mapping = catalogue.map(p => ({ ref: p.ref, name: p.name, barcode: p.barcode, standardPrice: p.standardPrice, listPrice: p.listPrice, ppc: p.ppc }));
      } catch { /* mapping limité aux articles */ }
      const res = await fetch("/api/export-template", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `Erreur ${res.status}`);
      const blob = await res.blob(); const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = `proposition_${(nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx`; a.click(); URL.revokeObjectURL(url);
      onToast("Fichier Proposition exporté", "success");
    } catch (e: any) { onToast("Erreur export : " + e.message, "error"); }
    finally { setExporting(false); }
  };

  if (!saved.length) return <div style={{ padding: 40, textAlign: "center", color: C.textMuted }}>Aucune campagne sauvegardée. Crée-en une dans « Créer une campagne » d'abord.</div>;

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
    <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Aperçu interactif de l'offre</h1>
        <span style={{ fontSize: 13, color: C.textMuted }}>Édite les paramètres → CA et marges se recalculent en direct.</span>
        <div style={{ flex: 1 }} />
        <select value={campId} onChange={e => { const c = saved.find(s => s.id === e.target.value); if (c) selectCamp(c); }} style={{ ...input, width: 240, textAlign: "left", fontSize: 13, padding: "7px 10px" }}>
          {saved.map(s => <option key={s.id} value={s.id}>{s.nom || "(sans nom)"}</option>)}
        </select>
        <button onClick={exporter} disabled={exporting} style={{ padding: "8px 16px", background: C.blue, border: "none", borderRadius: 8, cursor: exporting ? "default" : "pointer", fontSize: 13, fontWeight: 700, color: "#fff", fontFamily: "inherit", opacity: exporting ? 0.6 : 1 }}>{exporting ? "Export…" : "⬇ Exporter Excel"}</button>
      </div>

      {/* Synthèse globale (live) — figée en haut au scroll */}
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", position: "sticky", top: 0, zIndex: 10, background: C.bg, paddingTop: 4, paddingBottom: 8 }}>
        <Kpi label="CA total campagne" value={fmtEur(synthese.caTotal)} color={C.blue} />
        <Kpi label="Marge totale" value={fmtEur(synthese.margeTotal)} color={C.teal} />
        <Kpi label="Marge %" value={fmtPct(synthese.margePct)} color={C.green} />
        <Kpi label="Nb packs (tous paliers)" value={fmtNum(synthese.nbPacks)} color={C.amber} />
      </div>

      {/* Un bloc par palier */}
      {paliers.map((pal, pi) => {
        const r = calcPalier(calcPaliers[pi]);
        return (
          <div key={pi} style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
            {/* En-tête palier — figé au scroll, sous la barre KPI */}
            <div style={{ padding: "12px 16px", background: C.blueSoft, borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", position: "sticky", top: 92, zIndex: 8 }}>
              <span style={{ fontSize: 11, fontWeight: 800, fontFamily: "monospace", background: C.blue, color: "#fff", borderRadius: 5, padding: "2px 8px" }}>{pal.code}</span>
              <span style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{pal.label}</span>
              <span style={{ fontSize: 12, color: C.textMuted }}>Nb packs</span>
              <input type="number" style={{ ...input, width: 80 }} value={pal.nbPacks || ""} onChange={e => setPalier(pi, { nbPacks: parseInt(e.target.value) || 0 })} />
              <span style={{ fontSize: 12, color: C.textMuted }}>Remise add.</span>
              <input type="number" step="0.1" style={{ ...input, width: 60 }} value={Math.round(pal.remiseAdd * 1000) / 10} onChange={e => setPalier(pi, { remiseAdd: (parseFloat(e.target.value) || 0) / 100 })} />
              <span style={{ fontSize: 12, color: C.textMuted }}>%</span>
              <div style={{ flex: 1 }} />
              <span style={{ fontSize: 13, fontWeight: 800, color: C.blue }}>CA {fmtEur(r.caTotal)}</span>
              <span style={{ fontSize: 13, fontWeight: 800, color: C.teal }}>Marge {fmtEur(r.margeTotal)} ({fmtPct(r.margePct)})</span>
            </div>

            {/* Grille typologies (% offres + remises éditables, CA/marge live) */}
            <div style={{ padding: "8px 16px", overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: "left", padding: "4px 8px", color: C.textMuted, borderBottom: `1px solid ${C.border}` }}></th>
                    {TYPOLOGIES.map(t => <th key={t} style={{ padding: "4px 8px", color: C.textSec, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{t}</th>)}
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>% Offres</td>
                    {pal.pcts.map((v, ti) => <td key={ti} style={{ padding: "3px 4px", textAlign: "center" }}>
                      <input type="number" step="1" style={{ ...input, width: 52, textAlign: "center" }} value={Math.round(v * 1000) / 10} onChange={e => setPct(pi, ti, (parseFloat(e.target.value) || 0) / 100)} />
                    </td>)}
                  </tr>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>Remise</td>
                    {pal.remises.map((v, ti) => <td key={ti} style={{ padding: "3px 4px", textAlign: "center" }}>
                      <input type="number" step="0.1" style={{ ...input, width: 52, textAlign: "center" }} value={Math.round(v * 1000) / 10} onChange={e => setRemise(pi, ti, (parseFloat(e.target.value) || 0) / 100)} />
                    </td>)}
                  </tr>
                  <tr>
                    <td style={{ padding: "4px 8px", color: C.textMuted, fontWeight: 600 }}>Nb offres</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", color: C.textSec }}>{fmtNum(tr.nbOffres)}</td>)}
                  </tr>
                  <tr style={{ background: C.blueSoft }}>
                    <td style={{ padding: "4px 8px", color: C.blueDark, fontWeight: 700 }}>CA</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", fontWeight: 600, color: C.text }}>{fmtEur(tr.ca)}</td>)}
                  </tr>
                  <tr style={{ background: C.tealSoft }}>
                    <td style={{ padding: "4px 8px", color: C.teal, fontWeight: 700 }}>Marge</td>
                    {r.parTypo.map((tr, ti) => <td key={ti} style={{ padding: "4px 8px", textAlign: "center", fontWeight: 600, color: tr.marge >= 0 ? C.text : C.red }}>{fmtEur(tr.marge)}</td>)}
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Produits du palier (qté/pack éditable) */}
            <div style={{ padding: "0 16px 14px" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr>
                    {["Réf", "Produit", "EAN", "Qté/pack", "Coût", "Tarif", "PPC"].map((h, i) => <th key={i} style={{ padding: "5px 8px", textAlign: i >= 3 ? "right" : "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {pal.produits.map((p, ri) => (
                    <tr key={ri}>
                      <td style={{ padding: "4px 8px", fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.ref}</td>
                      <td style={{ padding: "4px 8px", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.name}</td>
                      <td style={{ padding: "4px 8px", color: C.textMuted, fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{p.barcode || "—"}</td>
                      <td style={{ padding: "3px 8px", textAlign: "right", borderBottom: `1px solid ${C.border}` }}>
                        <input type="number" style={{ ...input, width: 60, fontWeight: 700, color: C.blue }} value={p.qtyParPack || ""} onChange={e => setQty(pi, ri, parseInt(e.target.value) || 0)} />
                      </td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.standardPrice}</td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.listPrice}</td>
                      <td style={{ padding: "4px 8px", textAlign: "right", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{p.ppc}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        );
      })}

      {/* Besoin logistique (total par réf, live) */}
      <div style={{ background: C.white, border: `2px solid ${C.teal}`, borderRadius: 12, overflow: "hidden", boxShadow: C.shadow }}>
        <div style={{ padding: "12px 16px", background: C.tealSoft, borderBottom: `1px solid ${C.border}` }}>
          <span style={{ fontSize: 14, fontWeight: 800, color: C.teal }}>📦 Besoin total par référence</span>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr>{["Réf", "Produit", "Besoin total"].map((h, i) => <th key={i} style={{ padding: "8px 14px", textAlign: i === 2 ? "right" : "left", color: C.textMuted, fontWeight: 700, borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
          <tbody>
            {Object.entries(besoin).sort((a, b) => b[1].total - a[1].total).map(([ref, v]) => (
              <tr key={ref}>
                <td style={{ padding: "7px 14px", fontFamily: "monospace", borderBottom: `1px solid ${C.border}` }}>{ref}</td>
                <td style={{ padding: "7px 14px", color: C.textSec, borderBottom: `1px solid ${C.border}` }}>{v.name}</td>
                <td style={{ padding: "7px 14px", textAlign: "right", fontWeight: 700, color: C.teal, borderBottom: `1px solid ${C.border}` }}>{fmtNum(v.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 12, color: C.textMuted, fontStyle: "italic", paddingBottom: 20 }}>
        Aperçu en lecture/édition — les modifications ici ne sont pas sauvegardées. Pour exporter l'Excel, utilise « Créer une campagne ».
      </div>
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
