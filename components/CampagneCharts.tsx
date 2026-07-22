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
export function ChartCard({ title, children, full }: { title: string; children: React.ReactNode; full?: boolean }) {
  return (
    <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 18px", boxShadow: C.shadow, flex: full ? "1 1 100%" : "1 1 320px", minWidth: 0 }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 14 }}>{title}</div>
      {children}
    </div>
  );
}

// ── Barres horizontales (CA par offre, top délégués, top produits) ─────────────
export function HBarChart({ data, color = C.teal, valueFmt = fmtEur }: { data: { label: string; value: number }[]; color?: string; valueFmt?: (n: number) => string }) {
  if (!data.length) return <Empty />;
  const max = Math.max(...data.map(d => d.value), 1);
  const rowH = 22, gap = 5;
  const height = data.length * (rowH + gap);
  const labelW = 130, valueW = 78;
  return (
    <svg viewBox={`0 0 600 ${height}`} width="100%" style={{ display: "block" }} preserveAspectRatio="xMinYMin meet">
      {data.map((d, i) => {
        const y = i * (rowH + gap);
        const barMaxW = 600 - labelW - valueW - 8;
        const w = Math.max((d.value / max) * barMaxW, d.value > 0 ? 3 : 0);
        const c = color || PALETTE[i % PALETTE.length];
        return (
          <g key={i}>
            <text x={0} y={y + rowH / 2} dominantBaseline="middle" fontSize={11} fill={C.textSec} style={{ fontFamily: "inherit" }}>
              {d.label.length > 18 ? d.label.slice(0, 17) + "…" : d.label}
            </text>
            <rect x={labelW} y={y + 3} width={barMaxW} height={rowH - 6} rx={4} fill={C.bg} />
            <rect x={labelW} y={y + 3} width={w} height={rowH - 6} rx={4} fill={c} />
            <text x={labelW + barMaxW + valueW} y={y + rowH / 2} dominantBaseline="middle" textAnchor="end" fontSize={11} fontWeight={700} fill={C.text} style={{ fontFamily: "inherit" }}>
              {valueFmt(d.value)}
            </text>
          </g>
        );
      })}
    </svg>
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
