// app/api/export-preco/route.ts — Export Excel de la préconisation d'offre N+1
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;

const TEAL = "0D9488";
const AMBER = "F59E0B";
const PURPLE = "7C3AED";
const RED = "EF4444";
const DARK = "1A1A2E", WHITE = "FFFFFF";

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
const eur = (c: ExcelJS.Cell) => { c.numFmt = '#,##0 "€"'; };
const pct = (c: ExcelJS.Cell) => { c.numFmt = "0.0%"; };

interface PrecoProduit { ref: string; name: string; productId: number; ca: number; qty: number; pctSegment: number; conserve: boolean; }
interface PrecoStatut { statut: string; caSegment: number; conserves: PrecoProduit[]; candidats: PrecoProduit[]; aRetirer: PrecoProduit[]; }
interface PrecoResult {
  nom: string;
  produitsConserves: { ref: string; name: string; productId: number }[];
  parStatut: PrecoStatut[];
  global: { candidats: PrecoProduit[]; conserves: PrecoProduit[] };
}

export async function POST(req: NextRequest) {
  try {
    const p: PrecoResult = await req.json();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Analyse Offres"; wb.created = new Date();

    // ── Onglet Synthèse ───────────────────────────────────────────────────────
    const ws = wb.addWorksheet("Préco N+1", { views: [{ showGridLines: false }] });
    ws.columns = [{ width: 16 }, { width: 44 }, { width: 14 }, { width: 12 }, { width: 12 }];

    titleRow(ws, `Préconisation offre N+1 — basée sur « ${p.nom} »`, PURPLE, 5);
    ws.addRow([]);

    // Produits conservés
    titleRow(ws, "Produits conservés dans l'offre", TEAL, 5);
    headRow(ws, ["Réf", "Produit", "", "", ""], TEAL);
    for (const c of p.produitsConserves) {
      const r = ws.addRow([c.ref, c.name]);
      r.eachCell(cell => { cell.border = border(); cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } }; });
    }
    if (!p.produitsConserves.length) ws.addRow(["—", "Aucun produit conservé"]);
    ws.addRow([]);

    // Reco globale — candidats à ajouter
    titleRow(ws, "Candidats à ajouter (top ventes campagne)", AMBER, 5);
    headRow(ws, ["Réf", "Produit", "CA", "Qté", "% CA"], AMBER);
    for (const c of p.global.candidats) {
      const r = ws.addRow([c.ref, c.name, c.ca, c.qty, c.pctSegment]);
      r.eachCell(cell => { cell.border = border(); cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } }; });
      eur(r.getCell(3)); pct(r.getCell(5));
    }
    ws.addRow([]);

    // ── Un onglet par statut client ───────────────────────────────────────────
    for (const s of p.parStatut.filter(x => x.caSegment > 0)) {
      const safe = s.statut.replace(/[\\/?*[\]:]/g, "").slice(0, 28) || "Statut";
      const sw = wb.addWorksheet(safe, { views: [{ showGridLines: false }] });
      sw.columns = [{ width: 16 }, { width: 44 }, { width: 14 }, { width: 12 }, { width: 12 }];

      titleRow(sw, `Statut client : ${s.statut}`, PURPLE, 5);
      const caRow = sw.addRow(["CA du segment", "", s.caSegment]); eur(caRow.getCell(3));
      caRow.getCell(1).font = { bold: true, size: 10, name: "Calibri" };
      sw.addRow([]);

      titleRow(sw, "✓ Produits conservés — perf sur ce statut", TEAL, 5);
      headRow(sw, ["Réf", "Produit", "CA", "Qté", "% segment"], TEAL);
      for (const c of s.conserves) {
        const r = sw.addRow([c.ref, c.name, c.ca, c.qty, c.pctSegment]);
        r.eachCell(cell => { cell.border = border(); cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + (c.pctSegment < 0.02 ? RED : DARK) } }; });
        eur(r.getCell(3)); pct(r.getCell(5));
      }
      if (!s.conserves.length) sw.addRow(["—", "Aucun produit conservé ne vend sur ce segment"]);
      sw.addRow([]);

      titleRow(sw, "★ À ajouter — top ventes du segment", AMBER, 5);
      headRow(sw, ["Réf", "Produit", "CA", "Qté", "% segment"], AMBER);
      for (const c of s.candidats) {
        const r = sw.addRow([c.ref, c.name, c.ca, c.qty, c.pctSegment]);
        r.eachCell(cell => { cell.border = border(); cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } }; });
        eur(r.getCell(3)); pct(r.getCell(5));
      }
      if (!s.candidats.length) sw.addRow(["—", "Pas de candidat hors produits conservés"]);

      if (s.aRetirer.length) {
        sw.addRow([]);
        const r = sw.addRow([`⚠ Faibles sur ce statut (<2%) : ${s.aRetirer.map(x => x.ref).join(", ")}`]);
        sw.mergeCells(r.number, 1, r.number, 5);
        r.getCell(1).font = { italic: true, size: 10, color: { argb: "FF" + RED }, name: "Calibri" };
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
