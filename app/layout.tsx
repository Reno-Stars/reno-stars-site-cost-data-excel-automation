import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Site Cost Data - Excel Automation",
  description: "Upload input and output Excel files to automatically migrate site cost data",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="bg-gray-50 text-gray-900 antialiased">{children}</body>
    </html>
  );
}
