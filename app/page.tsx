"use client";

import { useState, useEffect } from "react";
import * as odoo from "@/lib/odoo";
import AnalyseScreen from "@/components/AnalyseScreen";

// ── Couleurs ──────────────────────────────────────────────────────────────────
const C = {
  bg: "#f8fafc", white: "#ffffff", text: "#1a1a2e", textMuted: "#6b7280",
  border: "#e5e7eb", blue: "#3b82f6", blueSoft: "#eff6ff",
  red: "#ef4444", redSoft: "#fef2f2",
  shadow: "0 1px 4px rgba(0,0,0,0.07)",
  shadowMd: "0 4px 16px rgba(0,0,0,0.12)",
};

// ── LocalStorage helpers ───────────────────────────────────────────────────────
const LS_CFG = "ao_config";
const LS_SESSION = "ao_session";

function loadCfg(): { url: string; db: string; login: string } | null {
  try { const c = localStorage.getItem(LS_CFG); return c ? JSON.parse(c) : null; } catch { return null; }
}
function saveCfg(url: string, db: string, login: string) {
  try { localStorage.setItem(LS_CFG, JSON.stringify({ url, db, login })); } catch {}
}
function loadSession(): odoo.OdooSession | null {
  try { const s = localStorage.getItem(LS_SESSION); return s ? JSON.parse(s) : null; } catch { return null; }
}
function saveSession(s: odoo.OdooSession) {
  try { localStorage.setItem(LS_SESSION, JSON.stringify(s)); } catch {}
}
function clearSession() {
  try { localStorage.removeItem(LS_SESSION); } catch {}
}

// ── Toast ─────────────────────────────────────────────────────────────────────
interface Toast { msg: string; type: "success" | "error" | "info"; id: number }

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

  // Charger la config et la session sauvegardées
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
    e.preventDefault();
    setError(""); setLoading(true);
    try {
      const sess = await odoo.authenticate({ url: url.trim().replace(/\/$/, ""), db: db.trim() }, login.trim(), password);
      saveCfg(url.trim().replace(/\/$/, ""), db.trim(), login.trim());
      saveSession(sess);
      setSession(sess);
    } catch (e: any) {
      setError(e.message || "Connexion impossible");
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = () => {
    clearSession();
    setSession(null);
    setPassword("");
  };

  if (!ready) return null;

  // ── App principale ──────────────────────────────────────────────────────────
  if (session) {
    return (
      <div style={{ minHeight: "100vh", background: C.bg }}>
        {/* Toast */}
        <div style={{ position: "fixed", top: 16, right: 16, zIndex: 9999, display: "flex", flexDirection: "column", gap: 8 }}>
          {toasts.map(t => (
            <div key={t.id} style={{
              padding: "10px 16px", borderRadius: 10, fontSize: 13, fontWeight: 600,
              background: t.type === "success" ? "#f0fdf4" : t.type === "error" ? C.redSoft : C.blueSoft,
              color: t.type === "success" ? "#166534" : t.type === "error" ? C.red : C.blue,
              border: `1px solid ${t.type === "success" ? "#bbf7d0" : t.type === "error" ? "#fecaca" : "#bfdbfe"}`,
              boxShadow: C.shadowMd, maxWidth: 320,
            }}>{t.msg}</div>
          ))}
        </div>

        {/* Header */}
        <div style={{ background: C.white, borderBottom: `1px solid ${C.border}`, padding: "12px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100, boxShadow: C.shadow }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 30, height: 30, borderRadius: 8, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
            </div>
            <div>
              <div style={{ fontSize: 15, fontWeight: 800, color: C.text }}>Analyse Offres</div>
              <div style={{ fontSize: 11, color: C.textMuted }}>{session.name} · {session.config.url.replace(/^https?:\/\//, "")}</div>
            </div>
          </div>
          <button onClick={handleLogout} style={{ padding: "6px 12px", background: C.bg, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.textMuted, fontFamily: "inherit", display: "flex", alignItems: "center", gap: 6 }}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Déconnexion
          </button>
        </div>

        <AnalyseScreen session={session} onToast={showToast} />
      </div>
    );
  }

  // ── Login ───────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ width: "100%", maxWidth: 400 }}>
        {/* Logo */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 56, height: 56, borderRadius: 16, background: C.blue, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
          </div>
          <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>Analyse Offres</div>
          <div style={{ fontSize: 13, color: C.textMuted, marginTop: 4 }}>Connexion à Odoo requise</div>
        </div>

        {/* Formulaire */}
        <form onSubmit={handleLogin} style={{ background: C.white, borderRadius: 18, padding: "28px 24px", boxShadow: C.shadowMd, border: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 14 }}>
          {[
            { label: "URL Odoo", value: url, set: setUrl, placeholder: "https://monentreprise.odoo.com", type: "url" },
            { label: "Base de données", value: db, set: setDb, placeholder: "ma-base", type: "text" },
            { label: "Login", value: login, set: setLogin, placeholder: "prenom.nom@entreprise.fr", type: "email" },
            { label: "Mot de passe", value: password, set: setPassword, placeholder: "••••••••", type: "password" },
          ].map(({ label, value, set, placeholder, type }) => (
            <div key={label}>
              <label style={{ fontSize: 11, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", display: "block", marginBottom: 6 }}>{label}</label>
              <input
                type={type}
                value={value}
                onChange={e => set(e.target.value)}
                placeholder={placeholder}
                required
                style={{ width: "100%", boxSizing: "border-box", padding: "10px 12px", border: `1.5px solid ${C.border}`, borderRadius: 10, fontSize: 14, fontFamily: "inherit", background: C.bg, color: C.text, outline: "none" }}
              />
            </div>
          ))}

          {error && (
            <div style={{ padding: "10px 12px", background: C.redSoft, border: `1px solid #fecaca`, borderRadius: 10, fontSize: 13, color: C.red, fontWeight: 600 }}>
              {error}
            </div>
          )}

          <button type="submit" disabled={loading} style={{ marginTop: 4, padding: "13px 0", background: loading ? C.border : C.blue, color: loading ? C.textMuted : "#fff", border: "none", borderRadius: 12, fontSize: 15, fontWeight: 700, cursor: loading ? "default" : "pointer", fontFamily: "inherit", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            {loading ? (
              <><span style={{ width: 14, height: 14, borderRadius: "50%", border: "2px solid #fff4", borderTopColor: C.blue, display: "inline-block", animation: "spin 0.7s linear infinite" }} />Connexion…</>
            ) : "Se connecter"}
          </button>
        </form>
      </div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
