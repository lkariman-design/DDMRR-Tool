import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "DDMR Tool — Digital Diagnostic and Maturity Report",
  description: "Executive-grade company assessment and scoring platform",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full">{children}</body>
    </html>
  );
}
