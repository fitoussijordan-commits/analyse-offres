// lib/server/excel-notes.ts — Taille des notes (bulles de survol) dans les exports Excel.
// ExcelJS donne à toute note une bulle fixe de 2 colonnes × 4 lignes, sans option pour la
// changer : une note de plusieurs lignes (détail logistique par offre) est coupée à l'écran.
// On remplace le calcul de l'ancre VML pour dimensionner chaque bulle selon son texte.
// À importer (effet de bord) dans les routes d'export, jamais côté navigateur.

// eslint-disable-next-line @typescript-eslint/no-var-requires
const VmlAnchorXform = require("exceljs/lib/xlsx/xform/comment/vml-anchor-xform");

const LARGEUR_COLS = 5;        // ~ 60 caractères par ligne avec des colonnes de 12
const CAR_PAR_LIGNE = 55;
const HAUTEUR_MAX = 40;

const proto = VmlAnchorXform.prototype;
if (!proto.__tailleAuContenu) {
  const renderOrigine = proto.render;
  proto.render = function (xmlStream: any, model: any) {
    if (model && !model.anchor && model.refAddress && model.note?.texts) {
      const texte: string = model.note.texts.map((t: any) => t.text || "").join("");
      const lignes = texte.split("\n").reduce((n, l) => n + Math.max(1, Math.ceil(l.length / CAR_PAR_LIGNE)), 0);
      const hauteur = Math.min(HAUTEUR_MAX, Math.max(4, lignes + 1));
      const l = model.refAddress.col, t = Math.max(model.refAddress.row - 2, 0);
      xmlStream.leafNode("x:Anchor", null, [l, 6, t, 14, l + LARGEUR_COLS, 2, t + hauteur, 16].join(", "));
      return;
    }
    return renderOrigine.call(this, xmlStream, model);
  };
  proto.__tailleAuContenu = true;
}

export {};
