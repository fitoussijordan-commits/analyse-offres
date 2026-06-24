// app/api/export-template/route.ts
// Exporte UNE campagne au format Proposition (1 onglet). Le remplissage est délégué au
// helper partagé lib/fill-proposition.ts (réutilisé par l'export multi-campagnes).

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import { fillPropositionSheet, PROP_SHEET, PropPayload } from "@/lib/fill-proposition";

export const maxDuration = 60;
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template-clean.xlsx");

export async function POST(req: NextRequest) {
  try {
    const payload: PropPayload = await req.json();
    const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
    if (!paliers.length) return NextResponse.json({ error: "Aucun palier à exporter" }, { status: 400 });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_PATH);
    const ws = wb.getWorksheet(PROP_SHEET);
    if (!ws) return NextResponse.json({ error: `Onglet "${PROP_SHEET}" introuvable dans le gabarit` }, { status: 500 });

    fillPropositionSheet(ws, payload);

    const buf = await wb.xlsx.writeBuffer();
    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="proposition_${(payload.nom || "campagne").replace(/[^a-zA-Z0-9_-]+/g, "_")}.xlsx"`,
      },
    });
  } catch (e: any) {
    console.error("Export template error:", e);
    return NextResponse.json({ error: e?.message || "Erreur export template" }, { status: 500 });
  }
}
