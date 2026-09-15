// app/api/db/[table]/route.ts — Passerelle navigateur → Supabase.
// Le navigateur n'a plus de clé Supabase : il appelle cette route avec sa session Odoo,
// on la vérifie, puis on relaie la requête PostgREST avec la clé secrète (côté serveur).
// Garde-fous : tables autorisées uniquement, pas de jointure imbriquée, pas de
// modification/suppression sans filtre, rate limit par IP.

import { NextRequest, NextResponse } from "next/server";
import { requireOdooSession } from "@/lib/server/odoo-auth";
import { SUPABASE_REST, adminHeaders } from "@/lib/server/supabase-admin";
import { checkRateLimit, getClientIp } from "@/lib/rateLimiter";

const TABLES = new Set([
  "campagnes", "campagnes_creees", "analyse_offres", "projets_kits",
  "planning_produits", "planning_quantites", "planning_historique",
]);
const PREFER_OK = /^(resolution=(merge|ignore)-duplicates|return=(minimal|representation))(,\s*(resolution=(merge|ignore)-duplicates|return=(minimal|representation)))*$/;
// Paramètres PostgREST qui ne sont pas des filtres de lignes.
const NON_FILTRES = new Set(["select", "order", "limit", "offset", "on_conflict", "columns"]);

async function relay(req: NextRequest, table: string, method: "GET" | "POST" | "PATCH" | "DELETE") {
  const rl = checkRateLimit(`db:${getClientIp(req)}`, 300, 60_000);
  if (!rl.allowed) return NextResponse.json({ error: "Trop de requêtes" }, { status: 429 });

  if (!TABLES.has(table)) return NextResponse.json({ error: "Table non autorisée" }, { status: 403 });

  const auth = await requireOdooSession(req);
  if (auth instanceof NextResponse) return auth;

  const params = req.nextUrl.searchParams;
  // Jointures imbriquées (select=*,autre_table(*)) interdites : elles ouvriraient d'autres tables.
  if ((params.get("select") || "").includes("(")) return NextResponse.json({ error: "Sélection non autorisée" }, { status: 400 });
  if ((method === "PATCH" || method === "DELETE") && ![...params.keys()].some(k => !NON_FILTRES.has(k))) {
    return NextResponse.json({ error: "Filtre obligatoire" }, { status: 400 });
  }

  let headers: Record<string, string>;
  try { headers = { ...adminHeaders(), "Content-Type": "application/json" }; }
  catch (e: any) { console.error("[api/db]", e.message); return NextResponse.json({ error: "Configuration serveur incomplète" }, { status: 500 }); }
  const prefer = req.headers.get("prefer");
  if (prefer && PREFER_OK.test(prefer.trim())) headers.Prefer = prefer.trim();

  const qs = params.toString();
  const res = await fetch(`${SUPABASE_REST}/${table}${qs ? `?${qs}` : ""}`, {
    method,
    headers,
    body: method === "GET" || method === "DELETE" ? undefined : await req.text(),
    cache: "no-store",
  });
  const body = await res.text();
  return new NextResponse(body || null, {
    status: res.status,
    headers: { "Content-Type": res.headers.get("content-type") || "application/json" },
  });
}

type Ctx = { params: Promise<{ table: string }> | { table: string } };
const tableOf = async (ctx: Ctx) => (await ctx.params).table;

export async function GET(req: NextRequest, ctx: Ctx) { return relay(req, await tableOf(ctx), "GET"); }
export async function POST(req: NextRequest, ctx: Ctx) { return relay(req, await tableOf(ctx), "POST"); }
export async function PATCH(req: NextRequest, ctx: Ctx) { return relay(req, await tableOf(ctx), "PATCH"); }
export async function DELETE(req: NextRequest, ctx: Ctx) { return relay(req, await tableOf(ctx), "DELETE"); }
