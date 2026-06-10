// app/api/export-planning/route.ts
// Export Excel du Planning, calqué sur le fichier d'origine "Planning 2026_avec_suivi.xlsm"
// Génération via ExcelJS (déjà utilisé pour l'export Analyse).

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;

const SUPABASE_URL = "https://fcjtntvuuhmrqgafdsjl.supabase.co/rest/v1";
const SUPABASE_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZjanRudHZ1dWhtcnFnYWZkc2psIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ1MTI2OTYsImV4cCI6MjA5MDA4ODY5Nn0.dx8b_rkv7Lt-9K-xGq9-z9OnLsolFNnWJfoTTA8re7M";

const H = { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` };

async function sbGet<T>(path: string): Promise<T> {
  const res = await fetch(`${SUPABASE_URL}${path}`, { headers: H });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return res.json();
}

// ── Couleurs ──────────────────────────────────────────────────────────────────
const TEAL = "0D9488";
const TEAL_S = "F0FDFA";
const DARK = "1A1A2E";
const LGRAY = "F1F5F9";
const WHITE = "FFFFFF";
const BORDER = "FFE5E7EB";

const MONTH_LETTER = ["J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D"];
const MONTH_FULL = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];

interface Produit {
  ref: string; material_text: string; material_id: string | null;
  product_range: string; package_size: string; content: string;
  tray_quantity: number; conso_moyenne: number;
}
interface Quantite { ref: string; mois: string; quantite: number }
interface HistEntry {
  ref: string; produit_label: string; mois: string;
  ancienne_qte: number | null; nouvelle_qte: number; variation: number;
  raison: string; modifie_par: string; modifie_le: string;
}

function thin() {
  return {
    top: { style: "thin" as const, color: { argb: BORDER } },
    bottom: { style: "thin" as const, color: { argb: BORDER } },
    left: { style: "thin" as const, color: { argb: BORDER } },
    right: { style: "thin" as const, color: { argb: BORDER } },
  };
}

export async function GET(req: NextRequest) {
  try {
    const year = Number(new URL(req.url).searchParams.get("year")) || new Date().getFullYear();

    const produits = await sbGet<Produit[]>(
      "/planning_produits?order=product_range.asc,material_text.asc&actif=eq.true"
    );
    const quantites = await sbGet<Quantite[]>(
      `/planning_quantites?mois=gte.${year}-01&mois=lte.${year}-12`
    );

    // Map ref -> { mois: qte }
    const qmap: Record<string, Record<string, number>> = {};
    for (const q of quantites) {
      (qmap[q.ref] ||= {})[q.mois] = q.quantite;
    }
    const months = Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);

    const wb = new ExcelJS.Workbook();
    wb.creator = "Analyse Odoo";
    wb.created = new Date();

    // ════════════════════════════════════════════════════════════════════════
    // Feuille Planning
    // ════════════════════════════════════════════════════════════════════════
    const ws = wb.addWorksheet(`Planning ${year}`, { views: [{ showGridLines: false, state: "frozen", xSplit: 7, ySplit: 2 }] });

    const FIXED = ["Product Range", "Package Size", "Material Text", "Material", "Tray Quantity", "Content", "REF FR", "Conso moyenne"];
    const nFixed = FIXED.length; // 8

    // Ligne 1 : lettres de mois au-dessus des colonnes mois
    MONTH_LETTER.forEach((l, i) => {
      const cell = ws.getCell(1, nFixed + 1 + i);
      cell.value = l;
      cell.alignment = { horizontal: "center" };
      cell.font = { bold: true, size: 9, color: { argb: "FF" + TEAL } };
    });

    // Ligne 2 : en-têtes
    const headerRow = [...FIXED, ...MONTH_FULL, "Total par REF"];
    headerRow.forEach((h, i) => {
      const cell = ws.getCell(2, i + 1);
      cell.value = h;
      cell.font = { bold: true, size: 10, color: { argb: "FF" + WHITE }, name: "Calibri" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
      cell.alignment = { horizontal: i < nFixed ? "left" : "center", vertical: "middle", wrapText: true };
      cell.border = thin();
    });
    ws.getRow(2).height = 28;

    // Lignes de données
    const monthTotals = new Array(12).fill(0);
    let grand = 0;
    produits.forEach((p, idx) => {
      const r = 3 + idx;
      const qtys = months.map((m) => qmap[p.ref]?.[m] ?? 0);
      const rowTotal = qtys.reduce((a, b) => a + b, 0);
      qtys.forEach((q, i) => (monthTotals[i] += q));
      grand += rowTotal;

      const vals = [
        p.product_range, p.package_size, p.material_text,
        p.material_id ?? "", p.tray_quantity ?? 0, p.content,
        p.ref, p.conso_moyenne ?? 0, ...qtys, rowTotal,
      ];
      vals.forEach((v, i) => {
        const cell = ws.getCell(r, i + 1);
        cell.value = v as any;
        cell.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK } };
        cell.alignment = { horizontal: i < nFixed ? "left" : "center", vertical: "middle" };
        cell.border = thin();
        if (idx % 2 === 1 && i < nFixed) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9FAFB" } };
        if (i >= nFixed && i < nFixed + 12 && typeof v === "number" && v > 0) cell.numFmt = "#,##0";
        if (i === headerRow.length - 1) {
          cell.font = { size: 10, bold: true, color: { argb: "FF" + TEAL } };
          cell.numFmt = "#,##0";
        }
      });
    });

    // Ligne TOTAL
    const totalRow = 3 + produits.length;
    const tc0 = ws.getCell(totalRow, 1);
    tc0.value = `TOTAL (${produits.length} produits)`;
    ws.mergeCells(totalRow, 1, totalRow, nFixed);
    tc0.font = { bold: true, size: 10, color: { argb: "FF" + DARK } };
    tc0.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + LGRAY } };
    tc0.alignment = { horizontal: "left", vertical: "middle" };
    monthTotals.forEach((t, i) => {
      const cell = ws.getCell(totalRow, nFixed + 1 + i);
      cell.value = t;
      cell.numFmt = "#,##0";
      cell.font = { bold: true, size: 10, color: { argb: "FF" + TEAL } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = thin();
    });
    const gc = ws.getCell(totalRow, nFixed + 13);
    gc.value = grand;
    gc.numFmt = "#,##0";
    gc.font = { bold: true, size: 11, color: { argb: "FF" + TEAL } };
    gc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL_S } };
    gc.alignment = { horizontal: "center", vertical: "middle" };
    gc.border = thin();

    // Largeurs
    const widths = [16, 14, 34, 12, 11, 12, 10, 13, ...new Array(12).fill(8.5), 13];
    widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));

    // ════════════════════════════════════════════════════════════════════════
    // Feuille Historique
    // ════════════════════════════════════════════════════════════════════════
    try {
      const hist = await sbGet<HistEntry[]>("/planning_historique?order=modifie_le.desc&limit=1000");
      const hs = wb.addWorksheet("Historique", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
      const hHeaders = ["Date / Heure", "Réf FR", "Produit", "Mois", "Ancienne Qté", "Nouvelle Qté", "Variation", "Raison"];
      hHeaders.forEach((h, i) => {
        const cell = hs.getCell(1, i + 1);
        cell.value = h;
        cell.font = { bold: true, size: 10, color: { argb: "FF" + WHITE } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF7C3AED" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = thin();
      });
      hist.forEach((h, idx) => {
        const r = 2 + idx;
        const [y, m] = h.mois.split("-");
        const moisLabel = `${MONTH_FULL[parseInt(m) - 1]} ${y}`;
        const vals = [
          new Date(h.modifie_le).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }),
          h.ref, h.produit_label, moisLabel,
          h.ancienne_qte ?? 0, h.nouvelle_qte, h.variation, h.raison,
        ];
        vals.forEach((v, i) => {
          const cell = hs.getCell(r, i + 1);
          cell.value = v as any;
          cell.font = { size: 10, color: { argb: "FF" + DARK } };
          cell.alignment = { horizontal: i >= 4 && i <= 6 ? "center" : "left", vertical: "middle" };
          cell.border = thin();
          if (i === 6 && typeof v === "number")
            cell.font = { size: 10, bold: true, color: { argb: v > 0 ? "FF16A34A" : v < 0 ? "FFDC2626" : "FF6B7280" } };
        });
      });
      [18, 11, 32, 14, 12, 12, 10, 40].forEach((w, i) => (hs.getColumn(i + 1).width = w));
    } catch { /* historique optionnel */ }

    const buf = await wb.xlsx.writeBuffer();
    return new NextResponse(buf, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="Planning_${year}.xlsx"`,
      },
    });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
