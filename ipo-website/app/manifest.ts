import type { MetadataRoute } from "next";

export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Smart Scouter",
    short_name: "Smart Scouter",
    description: "An AI Powered Company Search Engine",
    start_url: "/",
    scope: "/",
    display: "standalone",
    background_color: "#7dd3fc",
    theme_color: "#0f172a",
    orientation: "portrait-primary",
    categories: ["business", "finance", "productivity"],
    icons: [
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
