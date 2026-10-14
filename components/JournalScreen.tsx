"use client";
// Journal des modifications (réservé aux admins) : qui a créé / modifié / supprimé quoi,
// filtrable par utilisateur, par fonction de l'app et par action, avec le détail
// « avant → après » des champs modifiés.
import { useEffect, useMemo, useState } from "react";
import * as odoo from "@/lib/odoo";
import { C } from "@/lib/theme";

interface Ligne {
  id: number; cree_le: string; utilisateur: string; login: string;
  table_cible: string; action: string; objet_id: string; objet_nom: string | null;
  avant: any; apres: any;
}

// Fonction de l'app correspondant à chaque table.
export const FONCTIONS: Record<string, string> = {
  campagnes_creees: "Créer une campagne",
  campagnes: "Analyse des campagnes",
  analyse_offres: "Offres (analyse)",
  projets_kits: "Projets kits",
  planning_produits: "Planning — produits",
  planning_quantites: "Planning — quantités",
  parametres_app: "Paramètres (clients exclus)",
};
const ACTIONS: Record<string, { label: string; couleur: string; fond: string }> = {
  "création": { label: "Création", couleur: C.green, fond: C.greenSoft },
  modification: { label: "Modification", couleur: C.blue, fond: C.blueSoft },
  corbeille: { label: "Corbeille", couleur: C.amber, fond: C.amberSoft },
  restauration: { label: "Restauration", couleur: C.teal, fond: C.tealSoft },
  suppression: { label: "Suppression", couleur: C.red, fond: C.redSoft },
};

const sessionId = () => { try { return JSON.parse(localStorage.getItem("ao_session") || "null")?.sessionId || ""; } catch { return ""; } };
export async function accesJournal(): Promise<boolean> {
  try { const r = await fetch("/api/journal?acces=1", { headers: { "x-odoo-session": sessionId() } }); return r.ok && (await r.json()).acces === true; }
  catch { return false; }
}

const fmtDate = (iso: string) => new Date(iso).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
const court = (v: any) => {
  if (v == null || v === "") return "—";
  const t = typeof v === "object" ? JSON.stringify(v) : String(v);
  return t.length > 160 ? t.slice(0, 157) + "…" : t;
};

/** Listes d'objets (articles, paliers, enseignes…) : détail par élément identifié par sa
 *  réf / code / libellé / nom, au lieu d'une liste entière illisible. */
function diffListe(champ: string, a: any, b: any): { champ: string; avant: any; apres: any }[] {
  const cle = (x: any, i: number) => x && typeof x === "object" ? String(x.ref || x.code || x.label || x.nom || `#${i + 1}`) : `#${i + 1}`;
  if (!Array.isArray(a) || !Array.isArray(b) || ![...a, ...b].every(x => x && typeof x === "object")) return [{ champ, avant: a, apres: b }];
  const ma = new Map(a.map((x, i) => [cle(x, i), x])), mb = new Map(b.map((x, i) => [cle(x, i), x]));
  const out: { champ: string; avant: any; apres: any }[] = [];
  for (const id of new Set([...ma.keys(), ...mb.keys()])) {
    const x = ma.get(id), y = mb.get(id);
    if (!x) out.push({ champ: `${champ} › ${id} (ajouté)`, avant: null, apres: y });
    else if (!y) out.push({ champ: `${champ} › ${id} (retiré)`, avant: x, apres: null });
    else if (JSON.stringify(x) !== JSON.stringify(y)) {
      // Seulement les propriétés qui changent dans l'élément.
      for (const p of new Set([...Object.keys(x), ...Object.keys(y)])) {
        if (JSON.stringify(x[p]) !== JSON.stringify(y[p])) out.push({ champ: `${champ} › ${id} › ${p}`, avant: x[p], apres: y[p] });
      }
    }
  }
  return out;
}

/** Champs modifiés entre deux versions (1er niveau, et 1er niveau de `data` pour les campagnes). */
function differences(avant: any, apres: any): { champ: string; avant: any; apres: any }[] {
  if (!avant || !apres) return [];
  const out: { champ: string; avant: any; apres: any }[] = [];
  const ignorer = new Set(["updated_at"]);
  for (const k of new Set([...Object.keys(avant), ...Object.keys(apres)])) {
    if (ignorer.has(k)) continue;
    const a = avant[k], b = apres[k];
    if (k === "data" && a && b && typeof a === "object" && typeof b === "object") {
      for (const d of new Set([...Object.keys(a), ...Object.keys(b)])) {
        if (JSON.stringify(a[d]) !== JSON.stringify(b[d])) out.push(...diffListe(d, a[d], b[d]));
      }
    } else if (JSON.stringify(a) !== JSON.stringify(b)) out.push(...diffListe(k, a, b));
  }
  return out;
}

export default function JournalScreen({ onToast }: { session: odoo.OdooSession; onToast: (m: string, t?: "success" | "error" | "info") => void }) {
  const [lignes, setLignes] = useState<Ligne[]>([]);
  const [utilisateurs, setUtilisateurs] = useState<{ login: string; nom: string }[]>([]);
  const [login, setLogin] = useState(""); const [table, setTable] = useState(""); const [action, setAction] = useState(""); const [q, setQ] = useState("");
  const [chargement, setChargement] = useState(false);
  const [ouvert, setOuvert] = useState<number | null>(null);

  useEffect(() => {
    const t = setTimeout(async () => {
      setChargement(true);
      try {
        const p = new URLSearchParams({ limit: "500" });
        if (login) p.set("login", login); if (table) p.set("table", table); if (action) p.set("action", action); if (q.trim()) p.set("q", q.trim());
        const r = await fetch(`/api/journal?${p}`, { headers: { "x-odoo-session": sessionId() } });
        const j = await r.json();
        if (!r.ok) throw new Error(j.error || `Erreur ${r.status}`);
        setLignes(j.lignes); setUtilisateurs(j.utilisateurs);
      } catch (e: any) { onToast("Journal : " + e.message, "error"); }
      finally { setChargement(false); }
    }, 250);
    return () => clearTimeout(t);
  }, [login, table, action, q]);

  const parJour = useMemo(() => {
    const m = new Map<string, Ligne[]>();
    for (const l of lignes) { const j = new Date(l.cree_le).toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }); (m.get(j) || m.set(j, []).get(j)!).push(l); }
    return [...m];
  }, [lignes]);

  const select: React.CSSProperties = { padding: "7px 10px", border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 13, fontFamily: "inherit", color: C.text, background: C.white };

  return (
    <div style={{ flex: 1, height: "100%", overflowY: "auto", padding: 24 }}>
      <div style={{ maxWidth: 1800, margin: "0 auto", display: "flex", flexDirection: "column", gap: 14 }}>
        <div>
          <h1 style={{ fontSize: 20, fontWeight: 800, color: C.text, margin: 0 }}>Journal des modifications</h1>
          <div style={{ fontSize: 13, color: C.textMuted, marginTop: 3 }}>Qui a créé, modifié ou supprimé quoi. Visible uniquement par toi.</div>
        </div>

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 14px" }}>
          <select value={login} onChange={e => setLogin(e.target.value)} style={select}>
            <option value="">Tous les utilisateurs</option>
            {utilisateurs.map(u => <option key={u.login} value={u.login}>{u.nom}</option>)}
          </select>
          <select value={table} onChange={e => setTable(e.target.value)} style={select}>
            <option value="">Toutes les fonctions</option>
            {Object.entries(FONCTIONS).map(([t, l]) => <option key={t} value={t}>{l}</option>)}
          </select>
          <select value={action} onChange={e => setAction(e.target.value)} style={select}>
            <option value="">Toutes les actions</option>
            {Object.entries(ACTIONS).map(([a, v]) => <option key={a} value={a}>{v.label}</option>)}
          </select>
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Rechercher un objet (campagne, offre, réf…)" style={{ ...select, flex: "1 1 240px" }} />
          <span style={{ fontSize: 12, color: C.textMuted }}>{chargement ? "Chargement…" : `${lignes.length} modification(s)`}</span>
        </div>

        {!chargement && !lignes.length && (
          <div style={{ padding: 40, textAlign: "center", color: C.textMuted, fontSize: 13, background: C.white, border: `1px solid ${C.border}`, borderRadius: 12 }}>
            Aucune modification enregistrée pour ces filtres. Le journal démarre à sa mise en service : les changements plus anciens ne sont pas connus.
          </div>
        )}

        {parJour.map(([jour, ls]) => (
          <div key={jour}>
            <div style={{ fontSize: 12, fontWeight: 700, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", margin: "4px 0 6px" }}>{jour}</div>
            <div style={{ background: C.white, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
              {ls.map(l => {
                const a = ACTIONS[l.action] || { label: l.action, couleur: C.textSec, fond: C.bg };
                const diffs = l.action === "modification" ? differences(l.avant, l.apres) : [];
                const estOuvert = ouvert === l.id;
                return (
                  <div key={l.id} style={{ borderBottom: `1px solid ${C.border}` }}>
                    <div onClick={() => setOuvert(estOuvert ? null : l.id)} style={{ display: "grid", gridTemplateColumns: "60px 180px 120px 200px 1fr auto", gap: 12, alignItems: "center", padding: "10px 14px", cursor: "pointer" }}>
                      <span style={{ fontSize: 12, color: C.textMuted, fontVariantNumeric: "tabular-nums" }}>{new Date(l.cree_le).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}</span>
                      <span style={{ fontSize: 13, fontWeight: 600, color: C.text }} title={l.login}>{l.utilisateur}</span>
                      <span><span style={{ fontSize: 11, fontWeight: 700, color: a.couleur, background: a.fond, borderRadius: 5, padding: "2px 8px" }}>{a.label}</span></span>
                      <span style={{ fontSize: 12, color: C.textSec }}>{FONCTIONS[l.table_cible] || l.table_cible}</span>
                      <span style={{ fontSize: 13, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                        {l.objet_nom || l.objet_id || "—"}
                        {diffs.length > 0 && <span style={{ fontSize: 12, color: C.textMuted }}> · {[...new Set(diffs.map(d => d.champ.split(" › ")[0]))].slice(0, 4).join(", ")}{diffs.length > 4 ? "…" : ""}</span>}
                      </span>
                      <span style={{ fontSize: 12, color: C.textMuted }}>{estOuvert ? "▲" : "▼"}</span>
                    </div>
                    {estOuvert && (
                      <div style={{ padding: "4px 14px 14px 86px", fontSize: 12 }}>
                        <div style={{ color: C.textMuted, marginBottom: 8 }}>{fmtDate(l.cree_le)} · {l.login} · id {l.objet_id}</div>
                        {l.action === "modification" ? (
                          diffs.length ? (
                            <table style={{ borderCollapse: "collapse", width: "100%" }}>
                              <thead><tr>{["Champ", "Avant", "Après"].map(h => <th key={h} style={{ textAlign: "left", fontSize: 11, color: C.textMuted, padding: "4px 8px", borderBottom: `1px solid ${C.border}` }}>{h}</th>)}</tr></thead>
                              <tbody>
                                {diffs.map(d => (
                                  <tr key={d.champ}>
                                    <td style={{ padding: "5px 8px", fontWeight: 600, color: C.text, verticalAlign: "top", whiteSpace: "nowrap" }}>{d.champ}</td>
                                    <td style={{ padding: "5px 8px", color: C.red, verticalAlign: "top", fontFamily: "ui-monospace, monospace", wordBreak: "break-word" }}>{court(d.avant)}</td>
                                    <td style={{ padding: "5px 8px", color: C.green, verticalAlign: "top", fontFamily: "ui-monospace, monospace", wordBreak: "break-word" }}>{court(d.apres)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          ) : <div style={{ color: C.textMuted }}>Enregistré sans changement.</div>
                        ) : (
                          <pre style={{ margin: 0, padding: 10, background: C.bg, borderRadius: 8, maxHeight: 280, overflow: "auto", fontSize: 11 }}>{JSON.stringify(l.action === "suppression" ? l.avant : l.apres, null, 2)}</pre>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
