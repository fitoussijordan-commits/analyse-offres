// lib/server/odoo-auth.ts — Vérifie qu'une requête vient d'un utilisateur connecté à Odoo.
// L'app n'utilise pas Supabase Auth : la seule preuve d'identité est la session Odoo.
// On la valide auprès du serveur Odoo (ODOO_URL, jamais une URL fournie par le client,
// sinon un faux serveur pourrait répondre « connecté »).

import { NextRequest, NextResponse } from "next/server";
import { fetchT } from "@/lib/fetchTimeout";

export const SESSION_HEADER = "x-odoo-session";
const TTL_MS = 5 * 60_000;
const cache = new Map<string, { uid: number; exp: number }>();

async function validerSession(sessionId: string): Promise<number | null> {
  const hit = cache.get(sessionId);
  if (hit && hit.exp > Date.now()) return hit.uid;

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

  cache.set(sessionId, { uid, exp: Date.now() + TTL_MS });
  if (cache.size > 500) for (const [k, v] of cache) if (v.exp <= Date.now()) cache.delete(k);
  return uid;
}

/** Renvoie l'uid Odoo, ou une réponse d'erreur (401/500) à retourner telle quelle. */
export async function requireOdooSession(req: NextRequest): Promise<number | NextResponse> {
  const sessionId = (req.headers.get(SESSION_HEADER) || "").trim();
  if (!sessionId || !/^[A-Za-z0-9_\-]{10,200}$/.test(sessionId)) {
    return NextResponse.json({ error: "Non connecté" }, { status: 401 });
  }
  try {
    const uid = await validerSession(sessionId);
    return uid ?? NextResponse.json({ error: "Session Odoo expirée, reconnecte-toi" }, { status: 401 });
  } catch (e: any) {
    console.error("[odoo-auth]", e?.message);
    return NextResponse.json({ error: "Vérification de session impossible" }, { status: 500 });
  }
}
