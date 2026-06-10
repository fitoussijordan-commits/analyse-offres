// app/api/export-campaign/route.ts — Export Excel d'une analyse de campagne
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;

const BLUE = "3B82F6";
const TEAL = "0D9488";
const TEAL_S = "F0FDFA";
const PURPLE = "7C3AED";
const AMBER = "F59E0B";
const GREEN = "22C55E";
const DARK = "1A1A2E";
const LGRAY = "F1F5F9";
const WHITE = "FFFFFF";
const BORDER = "FFE5E7EB";

interface ProduitCA { ref: string; name: string; qtyVendue: number; ca: number; }
interface DelegueCA { name: string; qtyVendue: number; ca: number; }
interface ClientStat { name: string; qtyVendue: number; ca: number; nbCommandes: number; }
interface OffreBreakdown { code: string; label: string; caTotal: number; qtyTotal: number; }
interface Payload {
  nom: string; caTotal: number; qtyTotal: number; nbCommandes: number;
  produits: ProduitCA[]; delegues: DelegueCA[]; categories: ClientStat[]; adherents: ClientStat[];
  perOffre: OffreBreakdown[];
  split?: { valide: { qty: number; ca: number }; avenir: { qty: number; ca: number } };
  filterLabel?: string;
}

function thin() {
  return {
    top: { style: "thin" as const, color: { argb: BORDER } },
    bottom: { style: "thin" as const, color: { argb: BORDER } },
    left: { style: "thin" as const, color: { argb: BORDER } },
    right: { style: "thin" as const, color: { argb: BORDER } },
  };
}
function headerRow(ws: ExcelJS.Worksheet, row: number, cols: string[], bg: string, aligns: ("left" | "center" | "right")[]) {
  cols.forEach((c, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = c;
    cell.font = { bold: true, size: 10, color: { argb: "FF" + WHITE }, name: "Calibri" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } };
    cell.alignment = { horizontal: aligns[i] || "left", vertical: "middle", wrapText: true };
    cell.border = thin();
  });
  ws.getRow(row).height = 24;
}
function dataRow(ws: ExcelJS.Worksheet, row: number, vals: any[], aligns: ("left" | "center" | "right")[], eurCols: number[] = [], pctCols: number[] = [], even = false) {
  vals.forEach((v, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = v;
    cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } };
    cell.alignment = { horizontal: aligns[i] || "left", vertical: "middle" };
    cell.border = thin();
    if (even) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9FAFB" } };
    if (eurCols.includes(i)) cell.numFmt = '#,##0 "€"';
    else if (pctCols.includes(i)) cell.numFmt = "0.0%";
    else if (typeof v === "number") cell.numFmt = "#,##0";
  });
}

function buildSynthese(wb: ExcelJS.Workbook, title: string, color: string, rows: ClientStat[], total: number, firstColHeader: string) {
  const ws = wb.addWorksheet(title, { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
  headerRow(ws, 1, [firstColHeader, "Nb commandes", "Quantité", "CA (€)", "% du CA"], color, ["left", "center", "center", "right", "right"]);
  rows.forEach((r, i) => {
    dataRow(ws, i + 2, [r.name, r.nbCommandes, r.qtyVendue, Math.round(r.ca), total > 0 ? r.ca / total : 0], ["left", "center", "center", "right", "right"], [3], [4], i % 2 === 1);
  });
  const tr = rows.length + 2;
  dataRow(ws, tr, ["TOTAL", rows.reduce((s, r) => s + r.nbCommandes, 0), rows.reduce((s, r) => s + r.qtyVendue, 0), Math.round(total), 1], ["left", "center", "center", "right", "right"], [3], [4]);
  ws.getRow(tr).eachCell(c => { c.font = { bold: true, size: 10, color: { argb: "FF" + DARK } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } }; });
  [34, 14, 12, 16, 11].forEach((w, i) => (ws.getColumn(i + 1).width = w));
}

export async function POST(req: NextRequest) {
  try {
    const p: Payload = await req.json();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Analyse Campagnes";
    wb.created = new Date();

    // ── Récapitulatif ──────────────────────────────────────────────────────────
    const ws = wb.addWorksheet("Récapitulatif", { views: [{ showGridLines: false }] });
    ws.mergeCells("A1:D1");
    const t = ws.getCell("A1");
    t.value = `Campagne : ${p.nom}`;
    t.font = { bold: true, size: 16, color: { argb: "FF" + DARK } };
    ws.getRow(1).height = 28;
    if (p.filterLabel) {
      ws.mergeCells("A2:D2");
      const f = ws.getCell("A2"); f.value = `Filtre : ${p.filterLabel}`;
      f.font = { size: 11, italic: true, color: { argb: "FF64748B" } };
    }

    const kpis: [string, any, boolean][] = [
      ["CA total", Math.round(p.caTotal), true],
      ["Quantité totale", p.qtyTotal, false],
      ["Nb commandes", p.nbCommandes, false],
    ];
    if (p.split) {
      kpis.push(["CA validé (facturé)", Math.round(p.split.valide.ca), true]);
      kpis.push(["CA à venir", Math.round(p.split.avenir.ca), true]);
    }
    let r = 4;
    kpis.forEach(([label, val, eur]) => {
      const lc = ws.getCell(r, 1); lc.value = label; lc.font = { bold: true, size: 11, color: { argb: "FF334155" } };
      const vc = ws.getCell(r, 2); vc.value = val; vc.font = { bold: true, size: 12, color: { argb: "FF" + TEAL } };
      if (eur) vc.numFmt = '#,##0 "€"'; else vc.numFmt = "#,##0";
      r++;
    });

    // détail par offre
    if (p.perOffre?.length) {
      r += 1;
      const h = ws.getCell(r, 1); h.value = "Détail par offre"; h.font = { bold: true, size: 12, color: { argb: "FF" + DARK } }; r++;
      headerRow(ws, r, ["Code offre", "Libellé", "Quantité", "CA (€)"], BLUE, ["left", "left", "center", "right"]); r++;
      p.perOffre.forEach((o, i) => { dataRow(ws, r, [o.code, o.label || "", o.qtyTotal, Math.round(o.caTotal)], ["left", "left", "center", "right"], [3], [], i % 2 === 1); r++; });
    }
    [26, 28, 14, 16].forEach((w, i) => (ws.getColumn(i + 1).width = w));

    // ── Produits ────────────────────────────────────────────────────────────────
    const wp = wb.addWorksheet("Produits", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
    headerRow(wp, 1, ["Réf", "Produit", "Quantité", "CA (€)", "% du CA"], TEAL, ["left", "left", "center", "right", "right"]);
    p.produits.forEach((pr, i) => {
      dataRow(wp, i + 2, [pr.ref, pr.name, pr.qtyVendue, Math.round(pr.ca), p.caTotal > 0 ? pr.ca / p.caTotal : 0], ["left", "left", "center", "right", "right"], [3], [4], i % 2 === 1);
    });
    [14, 40, 12, 16, 11].forEach((w, i) => (wp.getColumn(i + 1).width = w));

    // ── Synthèse Délégués ─────────────────────────────────────────────────────────
    const wd = wb.addWorksheet("Synthèse délégués", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
    headerRow(wd, 1, ["Délégué", "Quantité", "CA (€)", "% du CA"], PURPLE, ["left", "center", "right", "right"]);
    p.delegues.forEach((d, i) => {
      dataRow(wd, i + 2, [d.name, d.qtyVendue, Math.round(d.ca), p.caTotal > 0 ? d.ca / p.caTotal : 0], ["left", "center", "right", "right"], [2], [3], i % 2 === 1);
    });
    const dtr = p.delegues.length + 2;
    dataRow(wd, dtr, ["TOTAL", p.delegues.reduce((s, d) => s + d.qtyVendue, 0), Math.round(p.caTotal), 1], ["left", "center", "right", "right"], [2], [3]);
    wd.getRow(dtr).eachCell(c => { c.font = { bold: true, size: 10, color: { argb: "FF" + DARK } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } }; });
    [34, 12, 16, 11].forEach((w, i) => (wd.getColumn(i + 1).width = w));

    // ── Synthèses clients ─────────────────────────────────────────────────────────
    buildSynthese(wb, "Catégorie statistique", AMBER, p.categories, p.caTotal, "Catégorie statistique");
    buildSynthese(wb, "Adhérent réseau", GREEN, p.adherents, p.caTotal, "Adhérent réseau");

    const buf = await wb.xlsx.writeBuffer();
    const safe = (p.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_");
    return new NextResponse(buf, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="campagne_${safe}.xlsx"`,
      },
    });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
