// lib/server/supabase-admin.ts — Accès Supabase CÔTÉ SERVEUR uniquement.
// La clé secrète (SUPABASE_SECRET_KEY) ne doit jamais être importée dans un composant client :
// elle contourne le RLS. Le navigateur passe par /api/db, qui vérifie la session Odoo.


export const SUPABASE_REST = `${(process.env.SUPABASE_URL || "https://fcjtntvuuhmrqgafdsjl.supabase.co").replace(/\/$/, "")}/rest/v1`;

/** En-têtes d'authentification Supabase avec la clé secrète. Lève une erreur si absente. */
export function adminHeaders(): Record<string, string> {
  const key = process.env.SUPABASE_SECRET_KEY;
  if (!key) throw new Error("SUPABASE_SECRET_KEY manquante côté serveur");
  // Nouvelles clés (sb_secret_…) : header apikey seul, elles ne sont pas des JWT.
  // Ancienne clé service_role (JWT) : apikey + Bearer.
  return key.startsWith("sb_") ? { apikey: key } : { apikey: key, Authorization: `Bearer ${key}` };
}

/** GET PostgREST avec la clé secrète (pour les routes serveur). */
export async function sbAdminGet<T>(path: string): Promise<T> {
  const res = await fetch(`${SUPABASE_REST}${path}`, { headers: adminHeaders(), cache: "no-store" });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.json();
}
