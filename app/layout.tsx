import type { Metadata } from "next";
export const metadata: Metadata = {
  title: "ANALYSE",
  description: "Plateforme d'analyse commerciale Odoo",
};
export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="fr">
      <body style={{ margin: 0, padding: 0, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
        {children}
      </body>
    </html>
  );
}
