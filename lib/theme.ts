// lib/theme.ts — Palette unique de l'application (refonte sobre & confortable).
// Un seul accent (indigo), neutres chauds agréables à l'œil, couleurs sémantiques
// désaturées. Tous les écrans importent C d'ici : une retouche = toute l'app suit.

export const C = {
  // Fonds & surfaces
  bg: "#f6f5f2",            // fond de page chaud, doux pour un usage prolongé
  white: "#ffffff",
  surface: "#ffffff",

  // Textes
  text: "#1d1f26",
  textSec: "#43454e",
  textMuted: "#7d7a72",
  slate: "#4c4f58",

  // Bordures
  border: "#e7e4de",
  borderDark: "#d4d0c8",
  borderFocus: "#4650a8",

  // Accent unique (indigo sobre)
  blue: "#4650a8",
  blueDark: "#343d85",
  blueSoft: "#eef0fa",

  // Sémantique désaturée (réservée aux statuts / marges / alertes)
  green: "#2e7d4f",
  greenSoft: "#eef6f0",
  greenDark: "#1d5637",
  amber: "#b3701c",
  amberSoft: "#faf4e9",
  red: "#c23636",
  redSoft: "#fbf0ef",
  orange: "#bb5c22",
  orangeSoft: "#faf2ec",

  // Secondaires désaturés
  purple: "#71589e",
  purpleSoft: "#f3f0f9",
  teal: "#337a74",
  tealSoft: "#edf5f4",

  // Ombres discrètes
  shadow: "0 1px 2px rgba(28,25,20,0.05)",
  shadowMd: "0 6px 20px rgba(28,25,20,0.09)",
};

// Série pour les graphiques (harmonisée, désaturée)
export const PALETTE = ["#4650a8", "#337a74", "#71589e", "#b3701c", "#2e7d4f", "#c23636", "#4a7fa5", "#a85980", "#5b8a72", "#8a6f3c"];

// Tokens du shell (sidebar / login) — dérivés de C
export const D = {
  sidebar: "#1b1d24", sidebarHover: "#262933", sidebarActive: "#343d85",
  sidebarBorder: "#262933", sidebarText: "#8b8e99", sidebarTextActive: "#ffffff",
  accent: C.blue, accentDark: C.blueDark,
  bg: C.bg, white: C.white,
  text: C.text, textMuted: C.textMuted,
  border: C.border,
  shadow: C.shadow, shadowMd: C.shadowMd,
  red: C.red, redSoft: C.redSoft, green: C.green,
  blueSoft: C.blueSoft,
};
