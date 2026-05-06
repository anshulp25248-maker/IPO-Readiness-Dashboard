import type { Metadata } from "next";
import { ScoutProvider } from "./_components/ScoutProvider";
import "./globals.css";

export const metadata: Metadata = {
  title: "Scout Smarter",
  description: "An AI Powered Company Search Engine",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col">
        <ScoutProvider>{children}</ScoutProvider>
      </body>
    </html>
  );
}
