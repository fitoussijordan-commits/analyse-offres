// app/api/journal/route.ts — Lecture du journal des modifications, réservée aux admins
// (lib/server/odoo-auth : estAdminJournal). ?acces=1 indique seulement si l'utilisateur y a droit.
import { NextRequest, NextResponse } from "next/server";
import { requireOdooUser, estAdminJournal } from "@/lib/server/odoo-auth";
import { SUPABASE_REST, adminHeaders } from "@/lib/server/supabase-admin";

export async function GET(req: NextRequest) {
  const user = await requireOdooUser(req);
  if (user instanceof NextResponse) return user;
  const admin = estAdminJournal(user);
  const p = req.nextUrl.searchParams;
  if (p.get("acces")) return NextResponse.json({ acces: admin });
  if (!admin) return NextResponse.json({ error: "Accès réservé" }, { status: 403 });

  const f = new URLSearchParams({ select: "*", order: "cree_le.desc", limit: String(Math.min(Number(p.get("limit")) || 300, 1000)) });
  const login = p.get("login"), table = p.get("table"), action = p.get("action"), q = (p.get("q") || "").trim();
  if (login) f.append("login", `eq.${login}`);
  if (table) f.append("table_cible", `eq.${table}`);
  if (action) f.append("action", `eq.${action}`);
  if (q) f.append("or", `(objet_nom.ilike.*${q.replace(/[*(),]/g, " ")}*,objet_id.ilike.*${q.replace(/[*(),]/g, " ")}*)`);
  try {
    const [lignes, users] = await Promise.all([
      fetch(`${SUPABASE_REST}/journal_modifications?${f}`, { headers: adminHeaders(), cache: "no-store" }).then(r => r.json()),
      // Liste des utilisateurs ayant déjà agi (pour le filtre).
      fetch(`${SUPABASE_REST}/journal_modifications?select=login,utilisateur&order=cree_le.desc&limit=5000`, { headers: adminHeaders(), cache: "no-store" }).then(r => r.json()),
    ]);
    const vus = new Map<string, string>();
    for (const u of Array.isArray(users) ? users : []) if (u.login && !vus.has(u.login)) vus.set(u.login, u.utilisateur || u.login);
    return NextResponse.json({ lignes: Array.isArray(lignes) ? lignes : [], utilisateurs: [...vus].map(([login, nom]) => ({ login, nom })) });
  } catch (e: any) {
    console.error("[api/journal]", e?.message);
    return NextResponse.json({ error: "Lecture du journal impossible" }, { status: 500 });
  }
}
