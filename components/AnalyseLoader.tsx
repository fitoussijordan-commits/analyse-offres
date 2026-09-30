"use client";
// Écran d'attente de l'analyse d'une campagne : trois produits Dr. Hauschka qui sautent à
// tour de rôle, une barre de progression et l'étape en cours.
// Repris de l'écran de lecture de bon de l'app commande (commande-app/components/BonLoader).
//
// La progression suit les vraies étapes de l'analyse (lib/analyse-campaign : onEtape) :
// à l'intérieur d'une étape, la barre approche son plafond sans jamais l'atteindre,
// puis saute au palier suivant quand l'étape change vraiment.
// Les visuels sont des PNG détourés servis depuis public/.
import { useEffect, useRef, useState } from "react";
import type { EtapeAnalyse } from "@/lib/analyse-campaign";
import { C } from "@/lib/theme";

// Hauteurs proches des proportions réelles (le tube est plus grand).
const PRODUCTS = [
  { src: "/serum-hydratant.png", height: 118 },
  { src: "/tube-purifiant.png", height: 150 },
  { src: "/lotion-tonifiante.png", height: 128 },
];
const JUMP_S = 1.1;
const TAU_MS = 2500;

// Petites phrases qui tournent pendant l'attente.
const TIPS = [
  "Dédoublonnage des lignes de commande…",
  "Séparation validé / à venir…",
  "Ventilation par statut client…",
  "Calcul de la marge ligne par ligne…",
];

export default function AnalyseLoader({ etape, nom }: { etape: EtapeAnalyse | null; nom?: string }) {
  const courante = etape ?? { label: "Préparation de l'analyse", de: 0, a: 3 };
  const [pct, setPct] = useState(courante.de);
  const [tip, setTip] = useState(0);
  const debut = useRef(Date.now());
  const cle = `${courante.label}|${courante.de}`;

  useEffect(() => { debut.current = Date.now(); }, [cle]);

  // Approche exponentielle du plafond de l'étape : rapide au début, puis ralentit.
  useEffect(() => {
    const t = setInterval(() => {
      const cible = courante.de + (courante.a - courante.de) * (1 - Math.exp(-(Date.now() - debut.current) / TAU_MS));
      setPct(p => Math.max(p, cible));
    }, 120);
    return () => clearInterval(t);
  }, [cle, courante.de, courante.a]);

  useEffect(() => {
    const t = setInterval(() => setTip(i => (i + 1) % TIPS.length), 3200);
    return () => clearInterval(t);
  }, []);

  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", flex: 1, minHeight: 440, padding: "20px" }}>
      <style>{`
        @keyframes aoTubeJump {
          0%   { transform: translateY(0) scale(1.12, 0.86); }
          12%  { transform: translateY(0) scale(0.94, 1.08); }
          45%  { transform: translateY(-62px) scale(1, 1) rotate(var(--tilt-a)); }
          55%  { transform: translateY(-66px) scale(1, 1) rotate(var(--tilt-b)); }
          88%  { transform: translateY(0) scale(0.96, 1.05); }
          100% { transform: translateY(0) scale(1.12, 0.86); }
        }
        @keyframes aoTubeShadow {
          0%, 100% { transform: scaleX(1.1); opacity: 0.22; }
          50%      { transform: scaleX(0.55); opacity: 0.1; }
        }
        @keyframes aoBarShine {
          from { transform: translateX(-80px); }
          to   { transform: translateX(460px); }
        }
        @media (prefers-reduced-motion: reduce) {
          .ao-tube, .ao-shadow, .ao-shine { animation: none !important; }
        }
      `}</style>

      {/* Produits + ombres, décalés dans le temps pour sauter à tour de rôle */}
      <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "center", gap: 18 }}>
        {PRODUCTS.map(({ src, height }, i) => {
          const n = PRODUCTS.length;
          const delay = `-${(i * JUMP_S) / n}s`;
          // Au centre : se dandine. Sur un côté : penche vers le centre en l'air.
          const side = i - (n - 1) / 2;
          const [a, b] = side === 0 ? ["-6deg", "4deg"] : side < 0 ? ["9deg", "3deg"] : ["-9deg", "-3deg"];
          return (
            <div key={src} style={{ position: "relative", height: 170, width: 90, display: "flex", alignItems: "flex-end", justifyContent: "center" }}>
              <div className="ao-shadow" style={{
                position: "absolute", bottom: 2, width: 64, height: 12, borderRadius: "50%",
                background: C.text, animation: `aoTubeShadow ${JUMP_S}s ease-in-out infinite`, animationDelay: delay,
              }} />
              <div className="ao-tube" style={{
                transformOrigin: "50% 100%", marginBottom: 8,
                animation: `aoTubeJump ${JUMP_S}s cubic-bezier(.45,.05,.55,.95) infinite`, animationDelay: delay,
                ["--tilt-a" as any]: a, ["--tilt-b" as any]: b,
              }}>
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src={src} alt="" style={{ height, maxWidth: 110, objectFit: "contain", display: "block" }} />
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ fontSize: 17, fontWeight: 700, color: C.text, marginTop: 16 }}>{courante.label}…</div>
      <div style={{ fontSize: 13, color: C.textMuted, marginTop: 4, minHeight: 18, textAlign: "center" }}>
        {courante.detail || TIPS[tip]}
      </div>

      {/* Barre */}
      <div style={{ width: "100%", maxWidth: 420, height: 8, borderRadius: 999, background: C.tealSoft, border: `1px solid ${C.border}`, marginTop: 18, overflow: "hidden" }}>
        <div style={{
          position: "relative", overflow: "hidden", height: "100%", width: `${pct}%`, borderRadius: 999,
          transition: "width 0.25s linear", background: C.teal,
        }}>
          <div className="ao-shine" style={{
            position: "absolute", inset: 0, width: 80,
            background: "linear-gradient(90deg, transparent, rgba(255,255,255,.45), transparent)",
            animation: "aoBarShine 1.4s linear infinite",
          }} />
        </div>
      </div>
      <div style={{ fontSize: 12, color: C.textMuted, marginTop: 8, fontVariantNumeric: "tabular-nums" }}>
        {Math.round(pct)} %{nom ? ` · ${nom}` : ""}
      </div>
    </div>
  );
}
