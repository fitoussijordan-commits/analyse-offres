// app/api/export-template/route.ts
// Exporte UNE campagne au format Proposition (1 onglet). Le remplissage est délégué au
// helper partagé lib/fill-proposition.ts (réutilisé par l'export multi-campagnes).

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import { fillPropositionWorkbook, PROP_SHEET, PropPayload } from "@/lib/fill-proposition";
import { writeSyntheseLogistiqueSheet, SyntheseLogistique } from "@/lib/logistique";
import { writeSyntheseDetailleeSheet } from "@/lib/synthese-detaillee";

export const maxDuration = 60;
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template-v2.xlsx");

export async function POST(req: NextRequest) {
  try {
    const payload: PropPayload & { logistique?: SyntheseLogistique } = await req.json();
    const paliers = (payload.paliers || []).filter(p => p.produits && p.produits.length);
    if (!paliers.length) return NextResponse.json({ error: "Aucun palier à exporter" }, { status: 400 });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_PATH);
    const ws = wb.getWorksheet(PROP_SHEET);
    if (!ws) return NextResponse.json({ error: `Onglet "${PROP_SHEET}" introuvable dans le gabarit` }, { status: 500 });

    fillPropositionWorkbook(wb, payload);

    // Onglet "Synthèse détaillée" (CA/marge par offre + par statut).
    if (paliers.length) writeSyntheseDetailleeSheet(wb, paliers);

    // Onglet "Synthèse logistique" (réf × mois) si la campagne fournit ses besoins.
    if (payload.logistique && payload.logistique.lignes?.length) {
      const nameByRef: Record<string, string> = {};
      for (const pal of payload.paliers) for (const p of pal.produits) { const r = (p.ref || "").trim(); if (r && p.name && !nameByRef[r]) nameByRef[r] = p.name; }
      writeSyntheseLogistiqueSheet(wb, payload.logistique, nameByRef);
    }

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
