import type { Metadata, Viewport } from "next";
import { ServiceWorkerRegistration } from "./_components/ServiceWorkerRegistration";
import { ScoutProvider } from "./_components/ScoutProvider";
import "./globals.css";

export const metadata: Metadata = {
  title: {
    default: "Smart Scouter",
    template: "%s | Smart Scouter",
  },
  description: "Company Search and Investment Readiness Engine",
  applicationName: "Smart Scouter",
  manifest: "/manifest.webmanifest",
  appleWebApp: {
    capable: true,
    statusBarStyle: "black-translucent",
    title: "Smart Scouter",
  },
  formatDetection: {
    telephone: false,
  },
  icons: {
    icon: [
      { url: "/icons/icon-192.png", sizes: "192x192", type: "image/png" },
      { url: "/icons/icon-512.png", sizes: "512x512", type: "image/png" },
    ],
    apple: [{ url: "/icons/apple-touch-icon.png", sizes: "180x180", type: "image/png" }],
  },
};

export const viewport: Viewport = {
  themeColor: "#0f172a",
  width: "device-width",
  initialScale: 1,
  viewportFit: "cover",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="h-full antialiased">
      <body className="min-h-full flex flex-col">
        <ServiceWorkerRegistration />
        <ScoutProvider>{children}</ScoutProvider>
      </body>
    </html>
  );
}
