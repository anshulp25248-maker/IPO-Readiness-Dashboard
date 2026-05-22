import type { MetadataRoute } from "next";

export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Smart Scouter",
    short_name: "Smart Scouter",
    description: "Company Search and Investment Readiness Engine",
    start_url: "/",
    scope: "/",
    display: "standalone",
    background_color: "#dfe9b8",
    theme_color: "#0f172a",
    orientation: "portrait-primary",
    categories: ["business", "finance", "productivity"],
    icons: [
      {
        src: "/icons/smart-scouter.svg",
        sizes: "any",
        type: "image/svg+xml",
        purpose: "any",
      },
      {
        src: "/icons/icon-192.png",
        sizes: "192x192",
        type: "image/png",
        purpose: "any",
      },
      {
        src: "/icons/icon-512.png",
        sizes: "512x512",
        type: "image/png",
        purpose: "any",
      },
      {
        src: "/icons/maskable-512.png",
        sizes: "512x512",
        type: "image/png",
        purpose: "maskable",
      },
    ],
  };
}
