// lib/server/odoo-auth.ts — Vérifie qu'une requête vient d'un utilisateur connecté à Odoo.
// L'app n'utilise pas Supabase Auth : la seule preuve d'identité est la session Odoo.
// On la valide auprès du serveur Odoo (ODOO_URL, jamais une URL fournie par le client,
// sinon un faux serveur pourrait répondre « connecté »).

import { NextRequest, NextResponse } from "next/server";
import { fetchT } from "@/lib/fetchTimeout";

export const SESSION_HEADER = "x-odoo-session";
const TTL_MS = 5 * 60_000;
export interface UtilisateurOdoo { uid: number; nom: string; login: string; }
const cache = new Map<string, { user: UtilisateurOdoo; exp: number }>();

async function validerSession(sessionId: string): Promise<UtilisateurOdoo | null> {
  const hit = cache.get(sessionId);
  if (hit && hit.exp > Date.now()) return hit.user;

  const base = (process.env.ODOO_URL || "").replace(/\/$/, "");
  if (!base) throw new Error("ODOO_URL manquante côté serveur");
  const res = await fetchT(`${base}/web/session/get_session_info`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Cookie: `session_id=${sessionId}` },
    body: JSON.stringify({ jsonrpc: "2.0", method: "call", id: Date.now(), params: {} }),
  });
  const data = await res.json().catch(() => null);
  const uid = data?.result?.uid;
  if (!res.ok || data?.error || typeof uid !== "number" || uid <= 0) { cache.delete(sessionId); return null; }

  const user: UtilisateurOdoo = { uid, nom: String(data.result.name || data.result.username || `Utilisateur ${uid}`), login: String(data.result.username || "") };
  cache.set(sessionId, { user, exp: Date.now() + TTL_MS });
  if (cache.size > 500) for (const [k, v] of cache) if (v.exp <= Date.now()) cache.delete(k);
  return user;
}

/** Utilisateur Odoo de la requête (nom, login), ou une réponse d'erreur 401/500. */
export async function requireOdooUser(req: NextRequest): Promise<UtilisateurOdoo | NextResponse> {
  const sessionId = (req.headers.get(SESSION_HEADER) || "").trim();
  if (!sessionId || !/^[A-Za-z0-9_\-]{10,200}$/.test(sessionId)) {
    return NextResponse.json({ error: "Non connecté" }, { status: 401 });
  }
  try {
    const user = await validerSession(sessionId);
    return user ?? NextResponse.json({ error: "Session Odoo expirée, reconnecte-toi" }, { status: 401 });
  } catch (e: any) {
    console.error("[odoo-auth]", e?.message);
    return NextResponse.json({ error: "Vérification de session impossible" }, { status: 500 });
  }
}

/** Accès au journal des modifications : logins listés dans JOURNAL_ADMINS (séparés par des
 *  virgules) ; à défaut, le compte de Jordan Fitoussi (login contenant « fitoussi »). */
export function estAdminJournal(user: UtilisateurOdoo): boolean {
  const liste = (process.env.JOURNAL_ADMINS || "").split(",").map(x => x.trim().toLowerCase()).filter(Boolean);
  const login = user.login.toLowerCase();
  return liste.length ? liste.includes(login) : login.includes("fitoussi");
}

/** Renvoie l'uid Odoo, ou une réponse d'erreur (401/500) à retourner telle quelle. */
export async function requireOdooSession(req: NextRequest): Promise<number | NextResponse> {
  const u = await requireOdooUser(req);
  return u instanceof NextResponse ? u : u.uid;
}
