// app/api/export-debug/route.ts — Fichier « debug » de l'analyse d'une campagne : toutes les
// lignes de commande retenues avec leur origine, et un récap qui reconstitue le CA total.
// Réservé à l'admin du journal (même contrôle serveur).
import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";
import { requireOdooUser, estAdminJournal } from "@/lib/server/odoo-auth";
import type { LigneDebug } from "@/lib/analyse-campaign";

export async function POST(req: NextRequest) {
  const user = await requireOdooUser(req);
  if (user instanceof NextResponse) return user;
  if (!estAdminJournal(user)) return NextResponse.json({ error: "Accès réservé" }, { status: 403 });

  const { nom, caTotal, margeTotal, lignes, exclus, config } = await req.json() as {
    nom: string; caTotal: number; margeTotal: number; lignes: LigneDebug[]; exclus: string[];
    config: { offres: string[]; produits: string[]; notes: string[]; periode: string };
  };
  const wb = new ExcelJS.Workbook();
  const EUR = '#,##0.00 "€"', ENTETE = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF1F4E79" } };
  const titre = (ws: ExcelJS.Worksheet, cols: string[]) => {
    const r = ws.addRow(cols);
    r.eachCell(c => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = ENTETE; });
    return r;
  };

  // ── Récap ─────────────────────────────────────────────────────────────────
  const rc = wb.addWorksheet("Récap");
  rc.columns = [{ width: 42 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 16 }, { width: 16 }];
  rc.addRow([`Debug analyse — ${nom}`]).font = { bold: true, size: 14 };
  rc.addRow([`Généré le ${new Date().toLocaleString("fr-FR")} par ${user.nom}`]).font = { italic: true, color: { argb: "FF6B7280" } };
  rc.addRow([]);
  rc.addRow(["Configuration"]).font = { bold: true };
  rc.addRow(["Offres (code du pack)", config.offres.join(", ") || "—"]);
  rc.addRow(["Articles seuls", config.produits.join(", ") || "—"]);
  rc.addRow(["Période (articles seuls)", config.periode || "—"]);
  rc.addRow(["Notes internes", config.notes.join(", ") || "—"]);
  rc.addRow(["Clients exclus", exclus.join(", ") || "—"]);
  rc.addRow([]);
  rc.addRow(["CA par origine (lignes dédoublonnées)"]).font = { bold: true };
  titre(rc, ["Origine", "Lignes", "Unités", "CA HT", "Coût", "Marge"]);
  const parSource = new Map<string, { n: number; q: number; ca: number; cout: number }>();
  for (const l of lignes) { const k = l.source || "(inconnue)"; const a = parSource.get(k) || { n: 0, q: 0, ca: 0, cout: 0 }; a.n++; a.q += l.qte; a.ca += l.ca; a.cout += l.cout; parSource.set(k, a); }
  for (const [k, a] of [...parSource].sort((x, y) => y[1].ca - x[1].ca)) {
    const r = rc.addRow([k, a.n, a.q, a.ca, a.cout, a.ca - a.cout]);
    [4, 5, 6].forEach(c => (r.getCell(c).numFmt = EUR));
  }
  const tot = lignes.reduce((s, l) => ({ n: s.n + 1, q: s.q + l.qte, ca: s.ca + l.ca, cout: s.cout + l.cout }), { n: 0, q: 0, ca: 0, cout: 0 });
  const rt = rc.addRow(["TOTAL", tot.n, tot.q, tot.ca, tot.cout, tot.ca - tot.cout]);
  rt.font = { bold: true }; [4, 5, 6].forEach(c => (rt.getCell(c).numFmt = EUR));
  rc.addRow([]);
  const ok = Math.abs(tot.ca - caTotal) < 0.5;
  const rv = rc.addRow(["Contrôle : CA total affiché dans l'app", "", "", caTotal, "", margeTotal]);
  [4, 6].forEach(c => (rv.getCell(c).numFmt = EUR));
  const re = rc.addRow([ok ? "✓ Le total des lignes = CA affiché" : `⚠ Écart de ${(tot.ca - caTotal).toFixed(2)} € avec le CA affiché`]);
  re.font = { bold: true, color: { argb: ok ? "FF15803D" : "FFB91C1C" } };
  rc.addRow([]);
  for (const t of [
    "Règles : CA = montant HT après remise de chaque ligne. Une ligne de commande n'est comptée qu'une fois ;",
    "« Aussi dans » = autres offres où elle a été trouvée (non recomptée). Offres et notes : sans limite de date ;",
    "articles seuls : commandes de la période uniquement. Coût = coût de la ligne de commande, sinon coût actuel de la fiche.",
  ]) rc.addRow([t]).font = { italic: true, size: 9, color: { argb: "FF6B7280" } };

  // ── Par produit ───────────────────────────────────────────────────────────
  const rp = wb.addWorksheet("Par produit × origine");
  rp.columns = [{ width: 12 }, { width: 42 }, { width: 30 }, { width: 10 }, { width: 14 }];
  titre(rp, ["Réf", "Produit", "Origine", "Unités", "CA HT"]);
  const pp = new Map<string, { ref: string; produit: string; source: string; q: number; ca: number }>();
  for (const l of lignes) { const k = `${l.ref}|${l.source}`; const a = pp.get(k) || { ref: l.ref, produit: l.produit, source: l.source, q: 0, ca: 0 }; a.q += l.qte; a.ca += l.ca; pp.set(k, a); }
  for (const a of [...pp.values()].sort((x, y) => x.ref.localeCompare(y.ref) || y.ca - x.ca)) rp.addRow([a.ref, a.produit, a.source, a.q, a.ca]).getCell(5).numFmt = EUR;
  rp.autoFilter = { from: "A1", to: "E1" };

  // ── Toutes les lignes ─────────────────────────────────────────────────────
  const wl = wb.addWorksheet("Lignes");
  wl.columns = [24, 26, 14, 11, 32, 22, 12, 40, 8, 13, 11, 22, 12, 9, 9].map(w => ({ width: w }));
  titre(wl, ["Origine", "Aussi dans", "Commande", "Date", "Client", "Délégué", "Réf", "Produit", "Qté", "CA HT", "Coût unit.", "Source du coût", "Coût", "Facturée", "Réassort"]);
  for (const l of [...lignes].sort((a, b) => a.date.localeCompare(b.date) || a.commande.localeCompare(b.commande))) {
    const r = wl.addRow([l.source, l.aussiDans, l.commande, l.date, l.client, l.delegue, l.ref, l.produit, l.qte, l.ca, l.coutUnitaire, l.coutSource, l.cout, l.facture ? "oui" : "non", l.reassort ? "oui" : ""]);
    [10, 11, 13].forEach(c => (r.getCell(c).numFmt = EUR));
    if (l.ca === 0) r.getCell(10).font = { color: { argb: "FFB91C1C" } };
  }
  const n = lignes.length + 1;
  const rs = wl.addRow(["TOTAL (lignes filtrées)", "", "", "", "", "", "", "", { formula: `SUBTOTAL(109,I2:I${n})` }, { formula: `SUBTOTAL(109,J2:J${n})` }, "", "", { formula: `SUBTOTAL(109,M2:M${n})` }]);
  rs.font = { bold: true }; [10, 13].forEach(c => (rs.getCell(c).numFmt = EUR));
  wl.autoFilter = { from: "A1", to: `O${n}` };
  wl.views = [{ state: "frozen", ySplit: 1 }];

  const buf = await wb.xlsx.writeBuffer();
  return new NextResponse(buf, { status: 200, headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } });
}
