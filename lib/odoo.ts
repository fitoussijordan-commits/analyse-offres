// lib/odoo.ts — Analyse Offres

export interface OdooConfig { url: string; db: string; }
export interface OdooSession { uid: number; name: string; login: string; sessionId: string; config: OdooConfig; }

async function rpc(config: OdooConfig, endpoint: string, params: any, sessionId?: string) {
  const res = await fetch("/api/odoo/proxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ odooUrl: config.url, endpoint, params, sessionId }),
  });
  const data = await res.json();
  if (!res.ok || data.error) throw new Error(data.error || `Erreur ${res.status}`);
  return { result: data.result, sessionId: data.sessionId };
}

export async function authenticate(config: OdooConfig, login: string, password: string): Promise<OdooSession> {
  const { result, sessionId: sid } = await rpc(config, "/web/session/authenticate", { db: config.db, login, password });
  if (!result || !result.uid || result.uid === false) throw new Error("Identifiants incorrects");
  return { uid: result.uid, name: result.name || result.username || login, login: login.toLowerCase(), sessionId: sid || result.session_id || "", config };
}

async function call(session: OdooSession, endpoint: string, params: any) {
  const { result } = await rpc(session.config, endpoint, params, session.sessionId);
  return result;
}

export async function searchRead(session: OdooSession, model: string, domain: any[], fields: string[], limit = 0, order = "", context: Record<string,any> = {}) {
  return call(session, "/web/dataset/call_kw", { model, method: "search_read", args: [domain], kwargs: { fields, limit, order, context } });
}

export interface MeaTemplate { id: number; name: string; active: boolean; }
export interface MeaTemplateLine { productCode: string; productName: string; }

/** Cherche dans les modèles de devis (actifs + archivés) par nom */
/** Normalise pour comparaison : minuscules + suppression des accents */
function norm(s: string): string {
  return (s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

export async function searchMeaTemplates(session: OdooSession, query: string): Promise<MeaTemplate[]> {
  const q = query.trim();
  if (!q) return [];
  // On cherche mot par mot pour contourner les problèmes de caractères spéciaux
  const firstWord = q.split(/\s+/)[0];
  const all = await searchRead(session, "sale.order.template", [["name", "ilike", firstWord]], ["id", "name", "active"], 50, "name", { active_test: false });
  if (!all?.length) return [];
  // Filtrage côté client sur tous les mots, insensible aux accents et à la casse
  const words = norm(q).split(/\s+/).filter(Boolean);
  return (all as any[])
    .filter(r => words.every(w => norm(r.name).includes(w)))
    .map(r => ({ id: r.id, name: r.name, active: r.active !== false }));
}

/** Récupère les lignes produits d'un modèle de devis */
export async function getMeaTemplateLines(session: OdooSession, templateId: number): Promise<MeaTemplateLine[]> {
  const lines = await searchRead(session, "sale.order.template.line", [["sale_order_template_id","=",templateId],["product_id","!=",false]], ["product_id","display_type"], 0);
  const productLines = (lines||[]).filter((l:any) => !l.display_type && l.product_id);
  if (!productLines.length) return [];
  const productIds: number[] = [...new Set<number>(productLines.map((l:any) => l.product_id[0] as number))];
  const prods = await searchRead(session, "product.product", [["id","in",productIds]], ["id","default_code","name"], 0);
  const prodMap: Record<number,{code:string;name:string}> = {};
  for (const p of (prods||[]) as any[]) if (p.default_code) prodMap[p.id] = { code: p.default_code, name: p.name };
  const seen = new Set<string>();
  const result: MeaTemplateLine[] = [];
  for (const l of productLines) {
    const entry = prodMap[(l as any).product_id[0]];
    if (entry && !seen.has(entry.code)) { seen.add(entry.code); result.push({ productCode: entry.code, productName: entry.name }); }
  }
  return result;
}
