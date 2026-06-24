// app/api/export-multi/route.ts
// Export multi-campagnes : un onglet "Proposition" par campagne sélectionnée + un onglet
// "Synthèse logistique" agrégeant les besoins par référence et par mois (jan-déc).

import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import path from "path";
import { fillPropositionSheet, PROP_SHEET, PropPayload } from "@/lib/fill-proposition";
import { MOIS_FR } from "@/lib/logistique";

export const maxDuration = 120;
const TEMPLATE_PATH = path.join(process.cwd(), "lib", "templates", "proposition-template-clean.xlsx");

interface MultiPayload {
  campagnes: PropPayload[];        // une PropPayload par campagne (nom + paliers enrichis)
  logistique: {                    // synthèse pré-calculée côté client
    lignes: { ref: string; name: string; parMois: number[]; total: number }[];
    totalParMois: number[];
    totalGeneral: number;
  };
}

// Couleurs de l'onglet synthèse.
const TEAL = "0D9488", DARK = "1A1A2E", WHITE = "FFFFFF", LGRAY = "F1F5F9";

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

    // 2) Onglet Synthèse logistique (réf × mois).
    const log = payload.logistique;
    const sw = wb.addWorksheet("Synthèse logistique", { views: [{ showGridLines: false }] });
    sw.columns = [{ width: 16 }, { width: 42 }, ...MOIS_FR.map(() => ({ width: 10 })), { width: 12 }];

    // Titre
    const titleRow = sw.addRow(["Synthèse besoins logistiques — par référence et par mois"]);
    sw.mergeCells(titleRow.number, 1, titleRow.number, 15);
    const tc = sw.getCell(titleRow.number, 1);
    tc.font = { bold: true, size: 13, color: { argb: "FF" + WHITE }, name: "Calibri" };
    tc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
    tc.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
    titleRow.height = 24;
    sw.addRow([]);

    // En-tête colonnes
    const head = sw.addRow(["Réf", "Produit", ...MOIS_FR, "Total"]);
    head.height = 20;
    head.eachCell(c => {
      c.font = { bold: true, color: { argb: "FF" + WHITE }, size: 10, name: "Calibri" };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
      c.alignment = { horizontal: "center", vertical: "middle" };
    });

    // Fallback libellé : si une ligne logistique n'a pas de nom, on le récupère depuis les
    // produits des campagnes (par réf), qui portent le libellé Odoo.
    const nameByRef: Record<string, string> = {};
    for (const camp of campagnes) for (const pal of camp.paliers) for (const p of pal.produits) {
      const ref = (p.ref || "").trim();
      if (ref && p.name && !nameByRef[ref]) nameByRef[ref] = p.name;
    }

    const numFmt = "#,##0";
    for (const l of log.lignes) {
      const libelle = l.name || nameByRef[l.ref] || "";
      const row = sw.addRow([l.ref, libelle, ...l.parMois, l.total]);
      row.eachCell((c, col) => {
        c.font = { size: 10, name: "Calibri", color: { argb: "FF" + DARK }, bold: col === 15 };
        c.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
        if (col >= 3) { c.numFmt = numFmt; c.alignment = { horizontal: "right" }; }
        if (col === 1) c.font = { ...c.font, name: "Consolas" };
      });
    }
    // Ligne total
    const totRow = sw.addRow(["", "TOTAL", ...log.totalParMois, log.totalGeneral]);
    totRow.eachCell((c, col) => {
      c.font = { bold: true, size: 10, name: "Calibri", color: { argb: "FF" + WHITE } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + TEAL } };
      if (col >= 3) { c.numFmt = numFmt; c.alignment = { horizontal: "right" }; }
    });

    // Note profil
    sw.addRow([]);
    const note = sw.addRow(["Profil de livraison : 40 % le mois précédant le début de l'offre, puis 60 % lissé à parts égales jusqu'à 1 mois avant la fin."]);
    sw.mergeCells(note.number, 1, note.number, 15);
    sw.getCell(note.number, 1).font = { italic: true, size: 9, color: { argb: "FF6B7280" }, name: "Calibri" };

    // 3) Supprimer le gabarit modèle (on ne garde que campagnes + synthèse).
    wb.removeWorksheet(tplId);

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
