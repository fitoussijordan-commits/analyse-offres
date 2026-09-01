// app/api/export-campaign/route.ts — Export Excel d'une campagne
// Reprend la richesse de l'export par offre (onglet par offre, Synthèse Articles,
// Commandes Note, Toutes Commandes) + synthèses campagne (délégués, catégorie, adhérent).
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;

const TEAL = "0D9488", TEAL_S = "F0FDFA";
const ORANGE = "F97316", ORANGE_S = "FFF7ED";
const PURPLE = "7C3AED", PURPLE_S = "F5F3FF";
const BLUE = "3B82F6", BLUE_S = "EFF6FF";
const AMBER = "F59E0B";
const GREEN = "22C55E";
const DARK = "1A1A2E", GRAY = "6B7280", LGRAY = "F1F5F9", WHITE = "FFFFFF";

function hdr(ws: ExcelJS.Worksheet, row: number, col: number, value: string, bg: string, span?: number) {
  const c = ws.getCell(row, col); c.value = value;
  c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 11, name: "Calibri" };
  c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } };
  c.alignment = { horizontal: "center", vertical: "middle" };
  if (span) ws.mergeCells(row, col, row, col + span - 1);
}
function border(): any { return { top: { style: "thin", color: { argb: "FFE5E7EB" } }, bottom: { style: "thin", color: { argb: "FFE5E7EB" } }, left: { style: "thin", color: { argb: "FFE5E7EB" } }, right: { style: "thin", color: { argb: "FFE5E7EB" } } }; }
// Date Odoo "YYYY-MM-DD" (ou datetime) → "JJ/MM/AAAA" ; vide si non renseignée.
function fmtDateFr(d?: string): string { const m = (d || "").match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : ""; }
function dataCell(cell: ExcelJS.Cell, value: any, bg: string, align: "left" | "center" | "right" = "left") {
  cell.value = value;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } };
  cell.alignment = { horizontal: align, vertical: "middle" };
  cell.border = border();
  cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } };
}
const eur = (c: ExcelJS.Cell) => { c.numFmt = '#,##0 "€"'; };
const pct = (c: ExcelJS.Cell) => { c.numFmt = "0.0%"; };
function headRow(ws: ExcelJS.Worksheet, cols: string[], bg: string) {
  const r = ws.addRow(cols); r.height = 22;
  r.eachCell(c => { c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } }; c.alignment = { horizontal: "center", vertical: "middle" }; c.border = border(); });
  return r;
}

interface ProduitCA { ref: string; name: string; qtyVendue: number; ca: number; }
interface DelegueCA { name: string; qtyVendue: number; ca: number; }
interface ClientStat { name: string; qtyVendue: number; ca: number; nbCommandes: number; }
interface DebugOrder { id: number; name: string; partnerName?: string; ca?: number; invoiced?: boolean; dateExpedition?: string; }
interface OffreAnalyse { offre: { code: string; label: string }; caTotal: number; qtyTotal: number; produits: ProduitCA[]; delegues: DelegueCA[]; debugOrders: DebugOrder[]; error: string | null; }
interface CatchallResult { codeInterne: string; data: { caTotal: number; qtyTotal: number; produits: ProduitCA[]; delegues: DelegueCA[]; debugOrders: DebugOrder[] } | null; }
interface Payload {
  nom: string; caTotal: number; qtyTotal: number; nbCommandes: number;
  produits: ProduitCA[]; delegues: DelegueCA[]; categories: ClientStat[]; adherents: ClientStat[]; statuts?: ClientStat[];
  results: OffreAnalyse[]; catchalls: CatchallResult[];
  split?: { valide: { qty: number; ca: number }; avenir: { qty: number; ca: number } };
  filterLabel?: string;
}

// ── Récapitulatif campagne ────────────────────────────────────────────────────
function buildRecap(wb: ExcelJS.Workbook, p: Payload) {
  const ws = wb.addWorksheet("Récapitulatif", { views: [{ showGridLines: false }] });
  ws.mergeCells("A1:F1");
  const t = ws.getCell("A1"); t.value = `Campagne : ${p.nom}`;
  t.font = { bold: true, size: 16, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } }; t.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 34;
  ws.mergeCells("A2:F2");
  const s = ws.getCell("A2"); s.value = `Exporté le ${new Date().toLocaleDateString("fr-FR")} — CA HT${p.filterLabel ? " — Filtre : " + p.filterLabel : ""}`;
  s.font = { italic: true, size: 10, color: { argb: "FF" + GRAY } }; s.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } }; s.alignment = { horizontal: "center" };

  let r = 4;
  const kpis: [string, number, boolean][] = [["CA total (sans doublons)", Math.round(p.caTotal), true], ["Offres / unités vendues", p.qtyTotal, false], ["Nb commandes", p.nbCommandes, false]];
  if (p.split) { kpis.push(["CA validé (facturé)", Math.round(p.split.valide.ca), true]); kpis.push(["CA à venir", Math.round(p.split.avenir.ca), true]); }
  for (const [label, val, isEur] of kpis) {
    ws.getCell(r, 1).value = label; ws.getCell(r, 1).font = { bold: true, size: 11, color: { argb: "FF334155" } };
    const v = ws.getCell(r, 3); v.value = val; v.font = { bold: true, size: 12, color: { argb: "FF" + TEAL } }; if (isEur) eur(v); else v.numFmt = "#,##0";
    r++;
  }

  r += 1;
  ws.getCell(r, 1).value = "Détail par offre / source"; ws.getCell(r, 1).font = { bold: true, size: 12, color: { argb: "FF" + DARK } }; r++;
  const startHeaderRow = r;
  while (ws.rowCount < r - 1) ws.addRow([]);
  headRow(ws, ["Code", "Libellé", "CA HT", "Qté (offres)", "Commandes", "Délégués"], DARK);
  const all = [
    ...p.results.map(x => ({ code: x.offre.code, label: x.offre.label, ca: x.caTotal, qty: x.qtyTotal, cmd: x.debugOrders.length, del: x.delegues.length, note: false })),
    ...p.catchalls.filter(c => c.data).map(c => ({ code: c.codeInterne, label: "Note interne", ca: c.data!.caTotal, qty: c.data!.qtyTotal, cmd: c.data!.debugOrders.length, del: c.data!.delegues.length, note: true })),
  ];
  all.forEach((x, i) => {
    const bg = i % 2 === 0 ? WHITE : LGRAY;
    const row = ws.addRow([x.code, x.label, x.ca, x.qty, x.cmd, x.del]);
    row.height = 20;
    row.eachCell((cell, col) => { dataCell(cell, cell.value, bg, col <= 2 ? "left" : "center"); });
    const caC = row.getCell(3); caC.font = { bold: true, color: { argb: "FF" + (x.note ? ORANGE : TEAL) }, size: 10 }; eur(caC);
    row.getCell(1).numFmt = "@";
  });
  void startHeaderRow;
  ws.columns = [{ width: 14 }, { width: 34 }, { width: 16 }, { width: 13 }, { width: 12 }, { width: 10 }];
}

// ── Feuille par offre ──────────────────────────────────────────────────────────
function buildOffreSheet(wb: ExcelJS.Workbook, r: OffreAnalyse) {
  const ca = r.caTotal, qty = r.qtyTotal;
  const ws = wb.addWorksheet(String(r.offre.code).slice(0, 31), { views: [{ showGridLines: false }] });
  ws.mergeCells("A1:G1");
  const t = ws.getCell("A1"); t.value = `${r.offre.code}${r.offre.label ? " — " + r.offre.label : ""}`;
  t.font = { bold: true, size: 15, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } }; t.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 36;
  ws.mergeCells("A2:B2"); const kl = ws.getCell("A2"); kl.value = "CA HT Total"; kl.font = { bold: true, color: { argb: "FF" + TEAL }, size: 10 }; kl.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } }; kl.alignment = { horizontal: "center", vertical: "middle" };
  ws.mergeCells("C2:D2"); const kv = ws.getCell("C2"); kv.value = ca; kv.font = { bold: true, color: { argb: "FF" + TEAL }, size: 16 }; kv.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } }; kv.alignment = { horizontal: "center", vertical: "middle" }; eur(kv);
  const ql = ws.getCell("E2"); ql.value = "Qté"; ql.font = { bold: true, color: { argb: "FF" + ORANGE }, size: 10 }; ql.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + ORANGE_S } }; ql.alignment = { horizontal: "center", vertical: "middle" };
  const qv = ws.getCell("F2"); qv.value = qty; qv.font = { bold: true, color: { argb: "FF" + ORANGE }, size: 16 }; qv.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + ORANGE_S } }; qv.alignment = { horizontal: "center", vertical: "middle" };
  const ol = ws.getCell("G2"); ol.value = `${r.debugOrders.length} commandes`; ol.font = { bold: true, color: { argb: "FF" + GRAY }, size: 10 }; ol.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + LGRAY } }; ol.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(2).height = 34; ws.getRow(3).height = 8;
  let cursor = 4;
  if (r.error) { ws.getCell(cursor, 1).value = "⚠ " + r.error; ws.getCell(cursor, 1).font = { color: { argb: "FF" + ORANGE }, bold: true }; }

  if (r.produits.length) {
    hdr(ws, cursor, 1, "Produits composants", BLUE, 5); ws.getRow(cursor).height = 22; cursor++;
    const ph = ws.addRow(["Référence", "Nom produit", "Qté vendue", "CA HT (€)", "% CA"]); ph.height = 18;
    ph.eachCell(c => { c.font = { bold: true, color: { argb: "FF1D4ED8" }, size: 10 }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFBFDBFE" } }; c.alignment = { horizontal: "center" }; c.border = border(); });
    cursor++;
    r.produits.forEach((p, i) => { const bg = i % 2 ? BLUE_S : WHITE; const row = ws.addRow([p.ref, p.name, p.qtyVendue, p.ca, ca > 0 ? p.ca / ca : 0]); row.height = 18; row.eachCell((c, col) => dataCell(c, c.value, bg, col <= 2 ? "left" : "center")); eur(row.getCell(4)); pct(row.getCell(5)); cursor++; });
    cursor++;
  }
  if (r.delegues.length) {
    hdr(ws, cursor, 1, "Par délégué", PURPLE, 4); ws.getRow(cursor).height = 22; cursor++;
    const dh = ws.addRow(["Délégué", "Qté vendue", "CA HT (€)", "% CA"]); dh.height = 18;
    dh.eachCell(c => { c.font = { bold: true, color: { argb: "FF" + PURPLE }, size: 10 }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDE9FE" } }; c.alignment = { horizontal: "center" }; c.border = border(); });
    cursor++;
    r.delegues.forEach((d, i) => { const bg = i % 2 ? PURPLE_S : WHITE; const row = ws.addRow([d.name, d.qtyVendue, d.ca, ca > 0 ? d.ca / ca : 0]); row.height = 18; row.eachCell((c, col) => dataCell(c, c.value, bg, col === 1 ? "left" : "center")); eur(row.getCell(3)); pct(row.getCell(4)); cursor++; });
  }
  ws.columns = [{ width: 14 }, { width: 36 }, { width: 14 }, { width: 16 }, { width: 10 }, { width: 3 }, { width: 16 }];
}

// ── Commandes Note ──────────────────────────────────────────────────────────────
function buildCommandesNote(wb: ExcelJS.Workbook, catchalls: CatchallResult[]) {
  const ws = wb.addWorksheet("Commandes Note", { views: [{ showGridLines: false }] });
  ws.mergeCells("A1:C1"); const t = ws.getCell("A1"); t.value = "Commandes notées (sans offre)";
  t.font = { bold: true, size: 14, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + ORANGE } }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 32; ws.getRow(2).height = 6;
  headRow(ws, ["Commande", "Client", "Note interne"], DARK);
  let i = 0;
  for (const c of catchalls) for (const o of (c.data?.debugOrders ?? [])) {
    const bg = i % 2 ? ORANGE_S : WHITE; const row = ws.addRow([o.name.replace(" (note)", ""), o.partnerName ?? "", c.codeInterne]); row.height = 18; row.eachCell(cell => dataCell(cell, cell.value, bg, "left")); i++;
  }
  ws.columns = [{ width: 16 }, { width: 42 }, { width: 16 }];
}

// ── Toutes Commandes ────────────────────────────────────────────────────────────
function buildToutesCommandes(wb: ExcelJS.Workbook, results: OffreAnalyse[], catchalls: CatchallResult[]) {
  const ws = wb.addWorksheet("Toutes Commandes", { views: [{ showGridLines: false }] });
  ws.mergeCells("A1:F1"); const t = ws.getCell("A1"); t.value = "Toutes les commandes — Offres + Notes";
  t.font = { bold: true, size: 14, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + DARK } }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 32; ws.getRow(2).height = 6;
  headRow(ws, ["Commande", "Client", "Source", "Libellé", "Type", "Expédition prévue"], DARK);
  const seen = new Set<string>();
  const rows: { name: string; partner: string; code: string; label: string; type: string; exp: string }[] = [];
  for (const r of results) for (const o of r.debugOrders) { const n = o.name.replace(" (note)", ""); if (seen.has(n)) continue; seen.add(n); rows.push({ name: n, partner: o.partnerName ?? "", code: r.offre.code, label: r.offre.label, type: "Offre", exp: fmtDateFr(o.dateExpedition) }); }
  for (const c of catchalls) for (const o of (c.data?.debugOrders ?? [])) { const n = o.name.replace(" (note)", ""); if (seen.has(n)) continue; seen.add(n); rows.push({ name: n, partner: o.partnerName ?? "", code: c.codeInterne, label: "Note interne", type: "Note", exp: fmtDateFr(o.dateExpedition) }); }
  rows.sort((a, b) => a.name.localeCompare(b.name));
  rows.forEach((o, i) => {
    const isNote = o.type === "Note"; const bg = isNote ? ORANGE_S : (i % 2 ? LGRAY : WHITE);
    const row = ws.addRow([o.name, o.partner, o.code, o.label, o.type, o.exp]); row.height = 18;
    row.eachCell((cell, col) => { dataCell(cell, cell.value, bg, col === 3 || col === 5 || col === 6 ? "center" : "left"); if (col === 3) { cell.font = { bold: true, color: { argb: "FF" + (isNote ? ORANGE : TEAL) }, size: 10 }; cell.numFmt = "@"; } if (col === 5 && isNote) cell.font = { bold: true, color: { argb: "FF" + ORANGE }, size: 10 }; });
  });
  ws.columns = [{ width: 16 }, { width: 42 }, { width: 14 }, { width: 34 }, { width: 10 }, { width: 18 }];
}

// ── Synthèse Articles ───────────────────────────────────────────────────────────
function buildSyntheseArticles(wb: ExcelJS.Workbook, produits: ProduitCA[], total: number) {
  const ws = wb.addWorksheet("Synthèse Articles", { views: [{ showGridLines: false, state: "frozen", ySplit: 3 }] });
  ws.mergeCells("A1:E1"); const t = ws.getCell("A1"); t.value = "Synthèse — Total vendu par article (sans doublons)";
  t.font = { bold: true, size: 14, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 32; ws.getRow(2).height = 6;
  headRow(ws, ["Référence", "Nom article", "Qté vendue", "CA HT total", "% CA"], DARK);
  produits.forEach((a, i) => { const bg = i % 2 ? TEAL_S : WHITE; const row = ws.addRow([a.ref, a.name, a.qtyVendue, a.ca, total > 0 ? a.ca / total : 0]); row.height = 18; row.eachCell((c, col) => dataCell(c, c.value, bg, col <= 2 ? "left" : "center")); const caC = row.getCell(4); caC.font = { bold: true, color: { argb: "FF" + TEAL }, size: 10 }; eur(caC); pct(row.getCell(5)); });
  ws.columns = [{ width: 16 }, { width: 44 }, { width: 14 }, { width: 16 }, { width: 10 }];
}

// ── Synthèse délégués (campagne, dédoublonnée) ──────────────────────────────────
function buildSyntheseDelegues(wb: ExcelJS.Workbook, delegues: DelegueCA[], total: number) {
  const ws = wb.addWorksheet("Synthèse délégués", { views: [{ showGridLines: false, state: "frozen", ySplit: 3 }] });
  ws.mergeCells("A1:D1"); const t = ws.getCell("A1"); t.value = "Synthèse délégués — campagne";
  t.font = { bold: true, size: 14, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + PURPLE } }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 32; ws.getRow(2).height = 6;
  headRow(ws, ["Délégué", "Qté vendue", "CA HT", "% CA"], DARK);
  delegues.forEach((d, i) => { const bg = i % 2 ? PURPLE_S : WHITE; const row = ws.addRow([d.name, d.qtyVendue, d.ca, total > 0 ? d.ca / total : 0]); row.height = 18; row.eachCell((c, col) => dataCell(c, c.value, bg, col === 1 ? "left" : "center")); eur(row.getCell(3)); pct(row.getCell(4)); });
  const tr = ws.addRow(["TOTAL", delegues.reduce((s, d) => s + d.qtyVendue, 0), Math.round(total), 1]); tr.height = 22;
  tr.eachCell((c, col) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + PURPLE } }; c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10 }; c.alignment = { horizontal: col === 1 ? "left" : "center", vertical: "middle" }; c.border = border(); }); eur(tr.getCell(3)); pct(tr.getCell(4));
  ws.columns = [{ width: 34 }, { width: 14 }, { width: 16 }, { width: 10 }];
}

// ── Synthèse clients (catégorie / adhérent) ─────────────────────────────────────
function buildSyntheseClient(wb: ExcelJS.Workbook, title: string, color: string, firstCol: string, rows: ClientStat[], total: number) {
  const ws = wb.addWorksheet(title, { views: [{ showGridLines: false, state: "frozen", ySplit: 3 }] });
  ws.mergeCells("A1:E1"); const t = ws.getCell("A1"); t.value = title;
  t.font = { bold: true, size: 14, color: { argb: "FF" + WHITE } }; t.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + color } }; t.alignment = { horizontal: "center", vertical: "middle" }; ws.getRow(1).height = 32; ws.getRow(2).height = 6;
  headRow(ws, [firstCol, "Nb commandes", "Qté", "CA HT", "% CA"], DARK);
  rows.forEach((rw, i) => { const bg = i % 2 ? LGRAY : WHITE; const row = ws.addRow([rw.name, rw.nbCommandes, rw.qtyVendue, rw.ca, total > 0 ? rw.ca / total : 0]); row.height = 18; row.eachCell((c, col) => dataCell(c, c.value, bg, col === 1 ? "left" : "center")); const caC = row.getCell(4); caC.font = { bold: true, color: { argb: "FF" + color }, size: 10 }; eur(caC); pct(row.getCell(5)); });
  const tr = ws.addRow(["TOTAL", rows.reduce((s, r) => s + r.nbCommandes, 0), rows.reduce((s, r) => s + r.qtyVendue, 0), Math.round(total), 1]); tr.height = 22;
  tr.eachCell((c, col) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + color } }; c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10 }; c.alignment = { horizontal: col === 1 ? "left" : "center", vertical: "middle" }; c.border = border(); }); eur(tr.getCell(4)); pct(tr.getCell(5));
  ws.columns = [{ width: 36 }, { width: 14 }, { width: 12 }, { width: 16 }, { width: 10 }];
}

export async function POST(req: NextRequest) {
  try {
    const p: Payload = await req.json();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Analyse Campagnes"; wb.created = new Date();

    buildRecap(wb, p);
    buildSyntheseArticles(wb, p.produits, p.caTotal);
    buildSyntheseDelegues(wb, p.delegues, p.caTotal);
    buildSyntheseClient(wb, "Catégorie statistique", AMBER, "Catégorie statistique", p.categories, p.caTotal);
    buildSyntheseClient(wb, "Adhérent réseau", GREEN, "Adhérent réseau", p.adherents, p.caTotal);
    if (p.statuts?.length) buildSyntheseClient(wb, "Statut client", AMBER, "Statut client", p.statuts, p.caTotal);
    for (const r of p.results) buildOffreSheet(wb, r);
    if (p.catchalls.some(c => (c.data?.debugOrders?.length ?? 0) > 0)) buildCommandesNote(wb, p.catchalls);
    buildToutesCommandes(wb, p.results, p.catchalls);

    const buf = await wb.xlsx.writeBuffer();
    const safe = (p.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_");
    return new NextResponse(buf, { headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": `attachment; filename="campagne_${safe}.xlsx"` } });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
