// app/api/export-preco/route.ts — Export Excel de la préco N+1 par palier + besoin fournisseur
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;

const TEAL = "0D9488";
const BLUE = "3B82F6";
const DARK = "1A1A2E", WHITE = "FFFFFF", GRAY = "9CA3AF";

function border(): any { return { top: { style: "thin", color: { argb: "FFE5E7EB" } }, bottom: { style: "thin", color: { argb: "FFE5E7EB" } }, left: { style: "thin", color: { argb: "FFE5E7EB" } }, right: { style: "thin", color: { argb: "FFE5E7EB" } } }; }
function headRow(ws: ExcelJS.Worksheet, cols: string[], bg: string) {
  const r = ws.addRow(cols); r.height = 22;
  r.eachCell(c => { c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } }; c.alignment = { horizontal: "center", vertical: "middle" }; c.border = border(); });
  return r;
}
function titleRow(ws: ExcelJS.Worksheet, text: string, bg: string, span: number) {
  const r = ws.addRow([text]); r.height = 24;
  ws.mergeCells(r.number, 1, r.number, span);
  const c = ws.getCell(r.number, 1);
  c.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
  c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + bg } };
  c.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
  return r;
}
function dataRow(ws: ExcelJS.Worksheet, vals: any[], muted = false) {
  const r = ws.addRow(vals);
  r.eachCell(cell => { cell.border = border(); cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + (muted ? GRAY : DARK) } }; });
  return r;
}
const eur = (c: ExcelJS.Cell) => { c.numFmt = '#,##0 "€"'; };
const num = (c: ExcelJS.Cell) => { c.numFmt = "#,##0"; };

interface PrecoLigne { ref: string; name: string; productId: number; ca: number; qty: number; conserve: boolean; }
interface PrecoPalier { code: string; label: string; caTotal: number; qtyPacks: number; produits: PrecoLigne[]; }
interface BesoinFournisseur { ref: string; name: string; productId: number; qty: number; ca: number; paliers: string[]; }
interface PrecoResult { nom: string; paliers: PrecoPalier[]; besoins: BesoinFournisseur[]; totalQty: number; }

export async function POST(req: NextRequest) {
  try {
    const p: PrecoResult = await req.json();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Analyse Offres"; wb.created = new Date();

    // ── Onglet 1 : Besoin fournisseur (la finalité) ─────────────────────────────
    const fw = wb.addWorksheet("Besoin fournisseur", { views: [{ showGridLines: false }] });
    fw.columns = [{ width: 16 }, { width: 46 }, { width: 18 }, { width: 18 }, { width: 14 }];
    titleRow(fw, `Besoin en commande fournisseur — offre N+1 (base « ${p.nom} »)`, TEAL, 5);
    fw.addRow([]);
    headRow(fw, ["Réf", "Produit", "Paliers", "Qté à commander", "CA associé"], TEAL);
    for (const b of p.besoins) {
      const r = dataRow(fw, [b.ref, b.name, b.paliers.join(", "), b.qty, b.ca]);
      num(r.getCell(4)); eur(r.getCell(5));
      r.getCell(4).font = { bold: true, size: 11, name: "Calibri", color: { argb: "FF" + DARK } };
    }
    // Ligne total
    const tr = fw.addRow(["", "TOTAL", "", p.totalQty, ""]);
    tr.eachCell(c => { c.border = border(); c.font = { bold: true, size: 11, name: "Calibri", color: { argb: "FF" + WHITE } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } }; });
    num(tr.getCell(4));
    if (!p.besoins.length) fw.addRow(["—", "Aucun produit conservé"]);

    // ── Un onglet par palier (offre) ────────────────────────────────────────────
    for (const pal of p.paliers) {
      const safe = (pal.code || "Offre").replace(/[\\/?*[\]:]/g, "").slice(0, 28);
      const sw = wb.addWorksheet(safe, { views: [{ showGridLines: false }] });
      sw.columns = [{ width: 16 }, { width: 46 }, { width: 14 }, { width: 14 }, { width: 12 }];
      titleRow(sw, `Palier ${pal.code} — ${pal.label || "Offre"}`, BLUE, 5);
      const info = sw.addRow([`CA palier : `, "", pal.caTotal, "", `${pal.qtyPacks} pack(s)`]);
      eur(info.getCell(3)); info.getCell(1).font = { bold: true, size: 10, name: "Calibri" };
      sw.addRow([]);
      headRow(sw, ["Réf", "Produit", "Conservé", "Qté vendue", "CA"], BLUE);
      for (const c of pal.produits) {
        const r = dataRow(sw, [c.ref, c.name, c.conserve ? "OUI" : "—", c.qty, c.ca], !c.conserve);
        num(r.getCell(4)); eur(r.getCell(5));
        r.getCell(3).alignment = { horizontal: "center" };
        if (c.conserve) r.getCell(3).font = { bold: true, size: 10, name: "Calibri", color: { argb: "FF" + TEAL } };
      }
    }

    const buf = await wb.xlsx.writeBuffer();
    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="preco.xlsx"`,
      },
    });
  } catch (e: any) {
    console.error("Export préco error:", e);
    return NextResponse.json({ error: e?.message || "Erreur export" }, { status: 500 });
  }
}
