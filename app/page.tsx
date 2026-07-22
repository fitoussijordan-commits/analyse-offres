"use client";
import { useState, useEffect } from "react";
import * as odoo from "@/lib/odoo";
import CampagneScreen from "@/components/CampagneScreen";
import CreerCampagneScreen from "@/components/CreerCampagneScreen";
import ApercuOffreScreen from "@/components/ApercuOffreScreen";
import ProjetScreen from "@/components/ProjetScreen";
import PlanningTab from "@/components/PlanningTab";

// ── Design tokens (palette centralisée) ───────────────────────────────────────
import { D } from "@/lib/theme";

// ── LocalStorage helpers ───────────────────────────────────────────────────────
const LS_CFG = "ao_config";
const LS_SESSION = "ao_session";
function loadCfg() { try { const c = localStorage.getItem(LS_CFG); return c ? JSON.parse(c) : null; } catch { return null; } }
function saveCfg(url: string, db: string, login: string) { try { localStorage.setItem(LS_CFG, JSON.stringify({ url, db, login })); } catch {} }
function loadSession(): odoo.OdooSession | null { try { const s = localStorage.getItem(LS_SESSION); return s ? JSON.parse(s) : null; } catch { return null; } }
function saveSession(s: odoo.OdooSession) { try { localStorage.setItem(LS_SESSION, JSON.stringify(s)); } catch {} }
function clearSession() { try { localStorage.removeItem(LS_SESSION); } catch {} }

interface Toast { msg: string; type: "success" | "error" | "info"; id: number; }
type AppView = "hub" | "analyse" | "creer" | "apercu" | "projets" | "planning";

// ── Nav item ──────────────────────────────────────────────────────────────────
function NavItem({ icon, label, active, locked, onClick }: { icon: React.ReactNode; label: string; active?: boolean; locked?: boolean; onClick?: () => void; }) {
  return (
    <button
      onClick={locked ? undefined : onClick}
      style={{
        width: "100%", padding: "8px 12px", display: "flex", alignItems: "center", gap: 10,
        background: active ? "#262933" : "transparent",
        border: "none", borderLeft: active ? "2px solid #7c88e0" : "2px solid transparent",
        borderRadius: "0 6px 6px 0", cursor: locked ? "default" : "pointer",
        color: active ? "#e7e9f2" : locked ? "#475569" : D.sidebarText,
        fontSize: 12.5, fontWeight: active ? 600 : 400, fontFamily: "inherit",
        marginBottom: 1, textAlign: "left" as const, opacity: locked ? 0.5 : 1,
      }}
    >
      <span style={{ flexShrink: 0, opacity: 0.85 }}>{icon}</span>
      <span style={{ flex: 1 }}>{label}</span>
      {locked && <span style={{ fontSize: 10, background: "#262933", padding: "2px 6px", borderRadius: 4, color: "#475569" }}>Bientôt</span>}
    </button>
  );
}

// ── Hub view ──────────────────────────────────────────────────────────────────
function HubView({ onNavigate }: { onNavigate: (v: AppView) => void }) {
  const tools = [
    {
      id: "analyse" as AppView, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#4650a8" strokeWidth="1.8">
          <line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>
        </svg>
      ),
      label: "Analyse des Campagnes", desc: "CA, délégués, produits et clients par campagne (offres + produits + notes), sans doublons",
      active: true,
    },
    {
      id: "creer" as AppView, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#4650a8" strokeWidth="1.8">
          <line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>
        </svg>
      ),
      label: "Créer une Campagne", desc: "Construis une campagne de zéro : paliers, articles, dates, et reco des quantités selon la conso N-1",
      active: true,
    },
    {
      id: "apercu" as AppView, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#337a74" strokeWidth="1.8">
          <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
        </svg>
      ),
      label: "Aperçu offre", desc: "Aperçu interactif d'une campagne : édite qté/remises/%, CA et marges recalculés en direct",
      active: true,
    },
    {
      id: "projets" as AppView, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#71589e" strokeWidth="1.8">
          <path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>
        </svg>
      ),
      label: "Projets Kits", desc: "Suivi des kits : composants, stock, dates arrivage et envoi ESAT",
      active: true,
    },
    {
      id: "planning" as AppView, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#337a74" strokeWidth="1.8">
          <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
        </svg>
      ),
      label: "Planning", desc: "Planification des quantités par référence et par mois",
      active: true,
    },
    {
      id: null, icon: (
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#6b6e78" strokeWidth="1.8">
          <circle cx="12" cy="12" r="3"/><path d="M12 1v4M12 19v4M4.22 4.22l2.83 2.83M16.95 16.95l2.83 2.83M1 12h4M19 12h4M4.22 19.78l2.83-2.83M16.95 7.05l2.83-2.83"/>
        </svg>
      ),
      label: "Outil 3", desc: "Prochainement disponible", active: false,
    },
  ];

  return (
    <div style={{ flex: 1, padding: "48px 48px", overflowY: "auto" as const }}>
      <div style={{ marginBottom: 40 }}>
        <h1 style={{ fontSize: 26, fontWeight: 800, color: D.text, margin: 0, letterSpacing: "-0.02em" }}>Tableau de bord</h1>
        <p style={{ fontSize: 14, color: D.textMuted, marginTop: 6 }}>Sélectionnez un outil pour commencer.</p>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 300px))", gap: 16 }}>
        {tools.map((t, i) => (
          <div
            key={i}
            onClick={t.active ? () => onNavigate(t.id!) : undefined}
            style={{
              background: D.white, border: `1.5px solid ${t.active ? D.accent + "44" : D.border}`,
              borderRadius: 14, padding: 24, cursor: t.active ? "pointer" : "default",
              boxShadow: t.active ? `0 4px 20px ${D.accent}18` : D.shadow,
              transition: "all 0.15s", opacity: t.active ? 1 : 0.55,
              display: "flex", flexDirection: "column" as const, gap: 14,
            }}
          >
            <div style={{ width: 52, height: 52, borderRadius: 12, background: t.active ? D.blueSoft : "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center" }}>
              {t.icon}
            </div>
            <div>
              <div style={{ fontSize: 15, fontWeight: 700, color: t.active ? D.text : "#475569", marginBottom: 4 }}>{t.label}</div>
              <div style={{ fontSize: 12, color: D.textMuted, lineHeight: 1.5 }}>{t.desc}</div>
            </div>
            {t.active && (
              <div style={{ marginTop: 4, fontSize: 12, fontWeight: 600, color: D.accent, display: "flex", alignItems: "center", gap: 4 }}>
                Ouvrir <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// PAGE PRINCIPALE
// ═══════════════════════════════════════════════════════════════════════════════
export default function Home() {
  const [session, setSession] = useState<odoo.OdooSession | null>(null);
  const [url, setUrl] = useState("");
  const [db, setDb] = useState("");
  const [login, setLogin] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [toasts, setToasts] = useState<Toast[]>([]);
  const [ready, setReady] = useState(false);
  const [view, setView] = useState<AppView>("hub");
  // Brouillon de campagne transféré depuis l'analyse (préco N+1) vers "Créer une campagne".
  const [transferDraft, setTransferDraft] = useState<import("@/lib/create-campaign").CampagneCreee | null>(null);

  useEffect(() => {
    const cfg = loadCfg();
    if (cfg) { setUrl(cfg.url); setDb(cfg.db); setLogin(cfg.login); }
    const sess = loadSession();
    if (sess) setSession(sess);
    setReady(true);
  }, []);

  const showToast = (msg: string, type: "success" | "error" | "info" = "info") => {
    const id = Date.now();
    setToasts(t => [...t, { msg, type, id }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 3500);
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault(); setError(""); setLoading(true);
    try {
      const sess = await odoo.authenticate({ url: url.trim().replace(/\/$/, ""), db: db.trim() }, login.trim(), password);
      saveCfg(url.trim().replace(/\/$/, ""), db.trim(), login.trim());
      saveSession(sess); setSession(sess);
    } catch (e: any) { setError(e.message || "Connexion impossible"); }
    finally { setLoading(false); }
  };

  const handleLogout = () => { clearSession(); setSession(null); setPassword(""); setView("hub"); };

  if (!ready) return null;

  // ── Login ───────────────────────────────────────────────────────────────────
  if (!session) {
    return (
      <div style={{ minHeight: "100vh", background: D.sidebar, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
        <div style={{ width: "100%", maxWidth: 380 }}>
          <div style={{ textAlign: "center", marginBottom: 32 }}>
            <div style={{ width: 52, height: 52, borderRadius: 14, background: D.accent, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px" }}>
              <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
            </div>
            <div style={{ fontSize: 24, fontWeight: 800, color: "#fff", letterSpacing: "-0.02em" }}>ANALYSE</div>
            <div style={{ fontSize: 13, color: D.sidebarText, marginTop: 6 }}>Connexion à votre instance Odoo</div>
          </div>
          <form onSubmit={handleLogin} style={{ background: "#262933", borderRadius: 16, padding: "28px 24px", border: "1px solid #3a3d47", display: "flex", flexDirection: "column", gap: 14 }}>
            {[
              { label: "URL Odoo", value: url, set: setUrl, placeholder: "https://monentreprise.odoo.com", type: "url" },
              { label: "Base de données", value: db, set: setDb, placeholder: "ma-base", type: "text" },
              { label: "Login", value: login, set: setLogin, placeholder: "prenom.nom@entreprise.fr", type: "email" },
              { label: "Mot de passe", value: password, set: setPassword, placeholder: "••••••••", type: "password" },
            ].map(({ label, value, set, placeholder, type }) => (
              <div key={label}>
                <label style={{ fontSize: 11, fontWeight: 600, color: "#64748b", textTransform: "uppercase" as const, letterSpacing: "0.07em", display: "block", marginBottom: 6 }}>{label}</label>
                <input type={type} value={value} onChange={e => set(e.target.value)} placeholder={placeholder} required
                  style={{ width: "100%", boxSizing: "border-box" as const, padding: "10px 13px", border: "1.5px solid #3a3d47", borderRadius: 10, fontSize: 14, fontFamily: "inherit", background: "#0f172a", color: "#e2e8f0", outline: "none" }} />
              </div>
            ))}
            {error && <div style={{ padding: "10px 13px", background: D.redSoft, border: "1px solid #fecaca", borderRadius: 10, fontSize: 13, color: D.red, fontWeight: 600 }}>{error}</div>}
            <button type="submit" disabled={loading} style={{ marginTop: 4, padding: "12px 0", background: loading ? "#3a3d47" : D.accent, color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: loading ? "default" : "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
              {loading ? <><span style={{ width: 14, height: 14, borderRadius: "50%", border: "2px solid rgba(255,255,255,0.3)", borderTopColor: "#fff", display: "inline-block", animation: "spin 0.7s linear infinite" }} />Connexion…</> : "Se connecter"}
            </button>
          </form>
        </div>
        <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  // ── App ─────────────────────────────────────────────────────────────────────
  return (
    <div style={{ display: "flex", height: "100vh", overflow: "hidden", background: D.bg, fontFamily: "-apple-system, BlinkMacSystemFont, \'Segoe UI\', sans-serif" }}>

      {/* Toasts */}
      <div style={{ position: "fixed", top: 16, right: 16, zIndex: 9999, display: "flex", flexDirection: "column" as const, gap: 8, pointerEvents: "none" }}>
        {toasts.map(t => (
          <div key={t.id} style={{
            padding: "11px 16px", borderRadius: 10, fontSize: 13, fontWeight: 600, maxWidth: 340,
            background: t.type === "success" ? "#f0fdf4" : t.type === "error" ? D.redSoft : D.blueSoft,
            color: t.type === "success" ? "#166534" : t.type === "error" ? D.red : D.accent,
            border: `1px solid ${t.type === "success" ? "#bbf7d0" : t.type === "error" ? "#fecaca" : "#bfdbfe"}`,
            boxShadow: D.shadowMd,
          }}>{t.msg}</div>
        ))}
      </div>

      {/* Sidebar */}
      <aside style={{ width: 220, background: D.sidebar, display: "flex", flexDirection: "column" as const, flexShrink: 0, zIndex: 10 }}>
        {/* Logo */}
        <div style={{ padding: "18px 16px 14px", borderBottom: "1px solid #262933" }}>
          <button onClick={() => setView("hub")} style={{ display: "flex", alignItems: "center", gap: 10, background: "transparent", border: "none", cursor: "pointer", padding: 0, width: "100%", textAlign: "left" as const }}>
            <div style={{ width: 28, height: 28, background: D.accent, borderRadius: 7, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
            </div>
            <div>
              <div style={{ fontSize: 13.5, fontWeight: 700, color: "#f2f2f4", letterSpacing: "0.03em" }}>ANALYSE</div>
              <div style={{ fontSize: 10, color: "#565b6b", marginTop: 1 }}>Gestion campagnes</div>
            </div>
          </button>
        </div>

        {/* Nav */}
        <nav style={{ flex: 1, padding: "12px 8px", overflowY: "auto" as const }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: "#565b6b", textTransform: "uppercase" as const, letterSpacing: "0.08em", padding: "0 8px 8px" }}>Outils</div>
          <NavItem
            icon={<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>}
            label="Analyse des Campagnes" active={view === "analyse"} onClick={() => setView("analyse")}
          />
          <NavItem
            icon={<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>}
            label="Projets Kits" active={view === "projets"} onClick={() => setView("projets")}
          />
          <NavItem
            icon={<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>}
            label="Planning" active={view === "planning"} onClick={() => setView("planning")}
          />
          <NavItem
            icon={<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>}
            label="Créer une Campagne" active={view === "creer"} onClick={() => setView("creer")}
          />
          <NavItem
            icon={<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>}
            label="Aperçu offre" active={view === "apercu"} onClick={() => setView("apercu")}
          />
        </nav>

        {/* User */}
        <div style={{ padding: "12px 14px", borderTop: "1px solid #262933" }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "#c9cbd2", marginBottom: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" as const }}>{session.name}</div>
          <div style={{ fontSize: 11, color: "#475569", marginBottom: 10, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" as const }}>{session.config.url.replace(/^https?:\/\//, "")}</div>
          <button onClick={handleLogout} style={{ width: "100%", padding: "7px 10px", background: "transparent", border: "1px solid #3a3d47", borderRadius: 7, cursor: "pointer", fontSize: 12, fontWeight: 600, color: "#64748b", fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Déconnexion
          </button>
        </div>
      </aside>

      {/* Main */}
      <main style={{ flex: 1, display: "flex", flexDirection: "column" as const, overflow: "hidden", minWidth: 0 }}>
        {view === "hub" && <HubView onNavigate={setView} />}
        {view === "analyse" && <CampagneScreen session={session} onToast={showToast} onTransferToCreer={(draft) => { setTransferDraft(draft); setView("creer"); }} />}
        {view === "creer" && <CreerCampagneScreen session={session} onToast={showToast} initialDraft={transferDraft} onDraftConsumed={() => setTransferDraft(null)} />}
        {view === "apercu" && <ApercuOffreScreen session={session} onToast={showToast} onGoAnalyse={() => setView("analyse")} />}
        {view === "projets" && <ProjetScreen session={session} onToast={showToast} />}
        {view === "planning" && <PlanningTab onToast={showToast} />}
      </main>

      <style>{`* { box-sizing: border-box; } body { -webkit-font-smoothing: antialiased; } @keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
