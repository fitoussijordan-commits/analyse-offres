// app/api/export-multi/route.ts
// Export multi-campagnes : un onglet "Proposition" par campagne sélectionnée + un onglet
// "Synthèse logistique" agrégeant les besoins par référence et par mois (jan-déc).

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import { fillPropositionSheet, PROP_SHEET, PropPayload } from "@/lib/fill-proposition";
import { writeSyntheseLogistiqueSheet, SyntheseLogistique } from "@/lib/logistique";

export const maxDuration = 120;
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template-v2.xlsx");

interface MultiPayload {
  campagnes: PropPayload[];        // une PropPayload par campagne (nom + paliers enrichis)
  logistique: SyntheseLogistique;  // synthèse pré-calculée côté client (axe mois dynamique)
}

function sanitizeSheetName(name: string, fallback: string): string {
  // Excel : max 31 car, pas de \ / ? * [ ] :
  const clean = (name || fallback).replace(/[\\/?*[\]:]/g, "").trim().slice(0, 28) || fallback;
  return clean;
}

export async function POST(req: NextRequest) {
  try {
    const payload: MultiPayload = await req.json();
    const campagnes = (payload.campagnes || []).filter(c => (c.paliers || []).some(p => p.produits?.length));
    if (!campagnes.length) return NextResponse.json({ error: "Aucune campagne à exporter" }, { status: 400 });

    // Charger le gabarit pour récupérer le modèle de la feuille Proposition.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(TEMPLATE_PATH);
    const tpl = wb.getWorksheet(PROP_SHEET);
    if (!tpl) return NextResponse.json({ error: `Onglet "${PROP_SHEET}" introuvable` }, { status: 500 });
    const tplModel = JSON.parse(JSON.stringify(tpl.model));
    const tplId = tpl.id;

    // 1) Un onglet Proposition par campagne (clone du gabarit + remplissage).
    const usedNames = new Set<string>();
    campagnes.forEach((camp, idx) => {
      let name = sanitizeSheetName(camp.nom, `Campagne ${idx + 1}`);
      let n = name, k = 2;
      while (usedNames.has(n)) { n = `${name.slice(0, 25)} (${k++})`; }
      usedNames.add(n);
      const ws = wb.addWorksheet(n);
      ws.model = { ...tplModel, name: n };
      ws.name = n;
      fillPropositionSheet(ws, camp);
    });

    // 2) Onglet Synthèse logistique (réf × mois) — fonction partagée (axe des mois dynamique,
    //    déborde sur N+1 si une offre finit l'année suivante).
    const nameByRef: Record<string, string> = {};
    for (const camp of campagnes) for (const pal of camp.paliers) for (const p of pal.produits) {
      const ref = (p.ref || "").trim();
      if (ref && p.name && !nameByRef[ref]) nameByRef[ref] = p.name;
    }
    writeSyntheseLogistiqueSheet(wb, payload.logistique, nameByRef);

    // 3) Supprimer les onglets du gabarit de base (Proposition modèle, Synthese, Mapping) :
    //    on ne garde que les onglets campagnes + la synthèse logistique.
    for (const sheet of [...wb.worksheets]) {
      if (sheet.id !== tplId && sheet.name !== "Synthese" && sheet.name !== "Mapping") continue;
      try { wb.removeWorksheet(sheet.id); } catch { /* ignore */ }
    }

    const buf = await wb.xlsx.writeBuffer();
    return new NextResponse(buf, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="campagnes_annee.xlsx"`,
      },
    });
  } catch (e: any) {
    console.error("Export multi error:", e);
    return NextResponse.json({ error: e?.message || "Erreur export multi" }, { status: 500 });
  }
}
