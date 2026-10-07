"use client";
// components/CampagneCharts.tsx — Graphiques SVG maison (sans dépendance) pour l'analyse de campagne
import React from "react";

import { C, PALETTE } from "@/lib/theme";

const fmtEur = (n: number) => new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 }).format(n || 0);
const fmtEurShort = (n: number) => {
  const v = n || 0;
  if (Math.abs(v) >= 1_000_000) return (v / 1_000_000).toFixed(1).replace(".0", "") + " M€";
  if (Math.abs(v) >= 1_000) return Math.round(v / 1000) + " k€";
  return Math.round(v) + " €";
};

// ── Conteneur carte ───────────────────────────────────────────────────────────
export function ChartCard({ title, children, full, aside }: { title: string; children: React.ReactNode; full?: boolean; aside?: React.ReactNode }) {
  return (
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 18px", boxShadow: C.shadow, minWidth: 0, ...(full ? { gridColumn: "1 / -1" } : {}) }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 14 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", flex: 1 }}>{title}</div>
        {aside}
      </div>
      {children}
    </div>
  );
}

/** Grille des graphiques : colonnes régulières qui s'adaptent à la largeur de l'écran. */
export function ChartGrid({ children }: { children: React.ReactNode }) {
  return <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(360px, 1fr))", gap: 12, marginBottom: 18, alignItems: "start" }}>{children}</div>;
}

// ── Barres horizontales (CA par offre, top délégués, top produits, répartitions) ─
// En HTML (et non en SVG étiré) : le texte garde une taille lisible quelle que soit la
// largeur de la carte. Libellé coupé selon la place réelle, complet au survol.
export interface BarItem { label: string; sub?: string; value: number; }
export function HBarChart({ data, color = C.teal, valueFmt = fmtEur, total, limite = 8 }: {
  data: BarItem[]; color?: string; valueFmt?: (n: number) => string;
  total?: number;     // base du % affiché (défaut : somme des valeurs)
  limite?: number;    // nb de lignes avant « Voir tout »
}) {
  const [tout, setTout] = React.useState(false);
  if (!data.length) return <Empty />;
  const max = Math.max(...data.map(d => d.value), 1);
  const base = total ?? data.reduce((s, d) => s + d.value, 0);
  const visibles = tout ? data : data.slice(0, limite);
  return (
    <div>
      <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
        {visibles.map((d, i) => {
          const pct = base > 0 ? (d.value / base) * 100 : 0;
          return (
            <div key={i} title={`${d.label}${d.sub ? ` · ${d.sub}` : ""} : ${valueFmt(d.value)} (${pct.toFixed(1).replace(".", ",")} %)`}
              style={{ display: "grid", gridTemplateColumns: "minmax(110px, 38%) 1fr auto", alignItems: "center", gap: 10 }}>
              <div style={{ minWidth: 0, lineHeight: 1.25 }}>
                {/* 2 lignes max : la fin du libellé (souvent la partie utile, ex. « TG VIP Retail ») reste visible. */}
                <div style={{ fontSize: 12.5, color: C.text, overflow: "hidden", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", wordBreak: "break-word" }}>{d.label}</div>
                {d.sub && <div style={{ fontSize: 10.5, color: C.textMuted, fontFamily: "ui-monospace, monospace" }}>{d.sub}</div>}
              </div>
              <div style={{ height: 10, background: C.bg, borderRadius: 5, overflow: "hidden" }}>
                <div style={{ width: `${Math.max((d.value / max) * 100, d.value > 0 ? 1.5 : 0)}%`, height: "100%", background: color, borderRadius: 5 }} />
              </div>
              <div style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", whiteSpace: "nowrap" }}>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: C.text }}>{valueFmt(d.value)}</span>
                <span style={{ fontSize: 11, color: C.textMuted, marginLeft: 6, display: "inline-block", minWidth: 38 }}>{pct.toFixed(pct < 10 ? 1 : 0).replace(".", ",")} %</span>
              </div>
            </div>
          );
        })}
      </div>
      {data.length > limite && (
        <button onClick={() => setTout(t => !t)}
          style={{ marginTop: 10, padding: "4px 10px", background: "transparent", border: `1px solid ${C.border}`, borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.textSec, fontFamily: "inherit" }}>
          {tout ? "Réduire" : `Voir tout (${data.length})`}
        </button>
      )}
    </div>
  );
}

// ── Répartition (catégorie, adhérent, statut) ─────────────────────────────────
// Barres triées, % du CA total. La part « non renseignée » est sortie du graphique et
// indiquée à part, pour que les vraies catégories restent lisibles.
const NON_RENSEIGNE = /non renseign/i;
export function RepartitionChart({ data, color = C.blue, valueFmt = fmtEurShort, libelleVide = "sans valeur renseignée" }: {
  data: { label: string; value: number }[]; color?: string; valueFmt?: (n: number) => string; libelleVide?: string;
}) {
  const positifs = data.filter(d => d.value > 0).sort((a, b) => b.value - a.value);
  if (!positifs.length) return <Empty />;
  const total = positifs.reduce((s, d) => s + d.value, 0);
  const vides = positifs.filter(d => NON_RENSEIGNE.test(d.label));
  const reels = positifs.filter(d => !NON_RENSEIGNE.test(d.label));
  const partVide = total > 0 ? vides.reduce((s, d) => s + d.value, 0) / total * 100 : 0;
  return (
    <div>
      <HBarChart data={reels} color={color} valueFmt={valueFmt} total={total} limite={7} />
      {partVide > 0 && (
        <div style={{ marginTop: 10, fontSize: 12, color: C.textMuted }}>
          + <strong style={{ color: C.textSec }}>{partVide.toFixed(1).replace(".", ",")} %</strong> du CA {libelleVide}
        </div>
      )}
    </div>
  );
}

// ── Camembert (répartition catégorie / adhérent) ──────────────────────────────
export function PieChart({ data }: { data: { label: string; value: number }[] }) {
  const filtered = data.filter(d => d.value > 0);
  if (!filtered.length) return <Empty />;
  const total = filtered.reduce((s, d) => s + d.value, 0);
  // top 6 + regroupement "Autres"
  const sorted = [...filtered].sort((a, b) => b.value - a.value);
  let slices = sorted;
  if (sorted.length > 7) {
    const top = sorted.slice(0, 6);
    const rest = sorted.slice(6).reduce((s, d) => s + d.value, 0);
    slices = [...top, { label: "Autres", value: rest }];
  }
  const cx = 90, cy = 90, r = 82;
  let angle = -Math.PI / 2;
  const arcs = slices.map((d, i) => {
    const frac = d.value / total;
    const a0 = angle, a1 = angle + frac * Math.PI * 2;
    angle = a1;
    const large = frac > 0.5 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const path = frac >= 0.9999
      ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy} Z`
      : `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`;
    return { path, color: PALETTE[i % PALETTE.length], label: d.label, value: d.value, frac };
  });
  return (
    <div style={{ display: "flex", gap: 16, alignItems: "center", flexWrap: "wrap" }}>
      <svg viewBox="0 0 180 180" width={180} height={180} style={{ flexShrink: 0 }}>
        {arcs.map((a, i) => <path key={i} d={a.path} fill={a.color} stroke={C.white} strokeWidth={1.5} />)}
      </svg>
      <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 140 }}>
        {arcs.map((a, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12 }}>
            <span style={{ width: 11, height: 11, borderRadius: 3, background: a.color, flexShrink: 0 }} />
            <span style={{ color: C.textSec, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.label}</span>
            <span style={{ color: C.textMuted, fontWeight: 600 }}>{(a.frac * 100).toFixed(1)}%</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── Barre empilée Validé / À venir ────────────────────────────────────────────
export function SplitBar({ valide, avenir }: { valide: number; avenir: number }) {
  const total = valide + avenir;
  if (total <= 0) return <Empty />;
  const pV = (valide / total) * 100;
  const pA = 100 - pV;
  return (
    <div>
      <div style={{ display: "flex", height: 38, borderRadius: 8, overflow: "hidden", border: `1px solid ${C.border}` }}>
        {pV > 0 && <div style={{ width: `${pV}%`, background: C.green, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 12, fontWeight: 700 }}>{pV >= 12 ? `${pV.toFixed(0)}%` : ""}</div>}
        {pA > 0 && <div style={{ width: `${pA}%`, background: C.amber, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 12, fontWeight: 700 }}>{pA >= 12 ? `${pA.toFixed(0)}%` : ""}</div>}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ width: 11, height: 11, borderRadius: 3, background: C.green }} />
          <div>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600 }}>Validé (facturé)</div>
            <div style={{ fontSize: 16, fontWeight: 800, color: C.green }}>{fmtEur(valide)}</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ width: 11, height: 11, borderRadius: 3, background: C.amber }} />
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600 }}>À venir</div>
            <div style={{ fontSize: 16, fontWeight: 800, color: C.amber }}>{fmtEur(avenir)}</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function Empty() {
  return <div style={{ padding: 24, textAlign: "center", color: C.textMuted, fontSize: 12 }}>Aucune donnée</div>;
}

export { fmtEurShort };

// ── Carte de barres agrandissable ─────────────────────────────────────────────
// Repliée : top de la mesure principale. Cliquée : pleine largeur, liste complète,
// choix de la mesure (CA / unités / commandes / marge) qui retrie, et toutes les mesures
// en colonnes. Un nouveau clic sur l'en-tête (ou « Réduire ») la replie.
export interface MesureCarte { cle: string; label: string; fmt: (n: number) => string; }
export interface LigneCarte { label: string; sub?: string; valeurs: Record<string, number>; }

export function CarteBarres({ title, lignes, mesures, color = C.teal, repartition, libelleVide = "sans valeur renseignée" }: {
  title: string; lignes: LigneCarte[]; mesures: MesureCarte[]; color?: string;
  repartition?: boolean;   // met de côté la part « Non renseigné » (catégories, statuts, réseaux)
  libelleVide?: string;
}) {
  const [agrandi, setAgrandi] = React.useState(false);
  const [mesure, setMesure] = React.useState(mesures[0].cle);
  const ref = React.useRef<HTMLDivElement>(null);
  const m = mesures.find(x => x.cle === mesure) || mesures[0];
  const totaux: Record<string, number> = {};
  for (const ms of mesures) totaux[ms.cle] = lignes.reduce((s, l) => s + Math.max(0, l.valeurs[ms.cle] || 0), 0);
  const estVide = (l: LigneCarte) => !!repartition && NON_RENSEIGNE.test(l.label);
  const reelles = lignes.filter(l => !estVide(l) && (l.valeurs[m.cle] || 0) !== 0).sort((a, b) => (b.valeurs[m.cle] || 0) - (a.valeurs[m.cle] || 0));
  const vides = lignes.filter(estVide);
  const partVide = totaux[m.cle] > 0 ? vides.reduce((s, l) => s + Math.max(0, l.valeurs[m.cle] || 0), 0) / totaux[m.cle] * 100 : 0;
  const basculer = () => {
    setAgrandi(a => !a);
    setTimeout(() => ref.current?.scrollIntoView({ behavior: "smooth", block: "nearest" }), 50);
  };
  if (!lignes.length) return null;

  const max = Math.max(...reelles.map(l => Math.abs(l.valeurs[m.cle] || 0)), 1);
  const pct = (v: number, cle: string) => totaux[cle] > 0 ? (v / totaux[cle]) * 100 : 0;
  const fmtPct = (p: number) => `${p.toFixed(p < 10 ? 1 : 0).replace(".", ",")} %`;

  return (
    <div ref={ref} style={{ background: C.white, border: `1px solid ${agrandi ? color + "88" : C.border}`, borderRadius: 12, padding: "16px 18px", boxShadow: agrandi ? C.shadowMd : C.shadow, minWidth: 0, gridColumn: agrandi ? "1 / -1" : undefined, transition: "box-shadow .2s, border-color .2s" }}>
      <div onClick={basculer} role="button" aria-expanded={agrandi} title={agrandi ? "Réduire" : "Agrandir : liste complète et choix de la mesure"}
        style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, cursor: "pointer", userSelect: "none" }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", flex: 1 }}>{title}</div>
        {!agrandi && <span style={{ fontSize: 11, color: C.textMuted }}>{reelles.length > 8 ? `${reelles.length} lignes · ` : ""}agrandir ⤢</span>}
        {agrandi && <span style={{ fontSize: 11, fontWeight: 600, color: C.textSec, border: `1px solid ${C.border}`, borderRadius: 6, padding: "2px 8px" }}>Réduire ⤡</span>}
      </div>

      {agrandi && mesures.length > 1 && (
        <div style={{ display: "inline-flex", gap: 3, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, padding: 3, marginBottom: 12 }}>
          {mesures.map(ms => (
            <button key={ms.cle} onClick={() => setMesure(ms.cle)}
              style={{ padding: "4px 12px", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit",
                background: mesure === ms.cle ? C.white : "transparent", color: mesure === ms.cle ? C.text : C.textMuted, boxShadow: mesure === ms.cle ? C.shadow : "none" }}>
              {ms.label}
            </button>
          ))}
        </div>
      )}

      {!agrandi ? (
        <HBarChart data={reelles.map(l => ({ label: l.label, sub: l.sub, value: l.valeurs[m.cle] || 0 }))} color={color} valueFmt={m.fmt} total={totaux[m.cle]} limite={8} />
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontVariantNumeric: "tabular-nums" }}>
            <thead>
              <tr>
                <th style={{ textAlign: "left", fontSize: 11, color: C.textMuted, fontWeight: 700, padding: "6px 8px", borderBottom: `1px solid ${C.border}`, width: "28%" }}>#  Libellé</th>
                <th style={{ fontSize: 11, color: C.textMuted, fontWeight: 700, padding: "6px 8px", borderBottom: `1px solid ${C.border}` }}>{m.label}</th>
                {mesures.map(ms => (
                  <th key={ms.cle} onClick={() => setMesure(ms.cle)} style={{ textAlign: "right", fontSize: 11, color: ms.cle === m.cle ? C.text : C.textMuted, fontWeight: 700, padding: "6px 8px", borderBottom: `1px solid ${C.border}`, cursor: "pointer", whiteSpace: "nowrap" }}>
                    {ms.label}{ms.cle === m.cle ? " ↓" : ""}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {reelles.map((l, i) => {
                const v = l.valeurs[m.cle] || 0;
                return (
                  <tr key={i} onMouseEnter={e => (e.currentTarget.style.background = C.bg)} onMouseLeave={e => (e.currentTarget.style.background = "transparent")}>
                    <td style={{ padding: "6px 8px", borderBottom: `1px solid ${C.border}`, minWidth: 180 }}>
                      <span style={{ fontSize: 11, color: C.textMuted, marginRight: 8 }}>{i + 1}</span>
                      <span style={{ fontSize: 13, color: C.text }}>{l.label}</span>
                      {l.sub && <span style={{ fontSize: 11, color: C.textMuted, marginLeft: 6, fontFamily: "ui-monospace, monospace" }}>{l.sub}</span>}
                    </td>
                    <td style={{ padding: "6px 8px", borderBottom: `1px solid ${C.border}`, minWidth: 140 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ flex: 1, height: 10, background: C.bg, borderRadius: 5, overflow: "hidden" }}>
                          <div style={{ width: `${Math.max(Math.abs(v) / max * 100, v ? 1.5 : 0)}%`, height: "100%", background: v < 0 ? C.red : color, borderRadius: 5 }} />
                        </div>
                        <span style={{ fontSize: 11, color: C.textMuted, minWidth: 42, textAlign: "right" }}>{fmtPct(pct(v, m.cle))}</span>
                      </div>
                    </td>
                    {mesures.map(ms => {
                      const val = l.valeurs[ms.cle] || 0;
                      return <td key={ms.cle} style={{ padding: "6px 8px", borderBottom: `1px solid ${C.border}`, textAlign: "right", fontSize: 13, whiteSpace: "nowrap", fontWeight: ms.cle === m.cle ? 700 : 400, color: val < 0 ? C.red : ms.cle === m.cle ? C.text : C.textSec }}>{ms.fmt(val)}</td>;
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {partVide > 0 && (
        <div style={{ marginTop: 10, fontSize: 12, color: C.textMuted }}>
          + <strong style={{ color: C.textSec }}>{partVide.toFixed(1).replace(".", ",")} %</strong> {m.cle === "ca" ? "du CA" : `des ${m.label.toLowerCase()}`} {libelleVide}
        </div>
      )}
    </div>
  );
}
