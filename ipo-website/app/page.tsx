"use client";

import type { CSSProperties } from "react";
import { FeatureCard } from "./_components/FeatureCard";
import { Drawer } from "./_components/AppShell";
import { useScout } from "./_components/ScoutProvider";
import { SmartScouterLogo } from "./_components/SmartScouterBrand";

const primaryCards = [
  { title: "Dashboard", href: "/dashboard", icon: "dashboard", color: "#2563eb", soft: "#dbeafe" },
  { title: "CDR", href: "/cdr", icon: "report", color: "#7c3aed", soft: "#ede9fe" },
  { title: "Other Companies", href: "/top-10-companies", icon: "ranking", color: "#0f9f6e", soft: "#d1fae5" },
] as const;

const secondaryCards = [
  { title: "Director Profile", href: "/director-profile", icon: "profile", color: "#ea580c", soft: "#ffedd5" },
  { title: "Competitors", href: "/competitors", icon: "competitors", color: "#db2777", soft: "#fce7f3" },
] as const;

const homeCards = [...primaryCards, ...secondaryCards] as const;

function HamburgerButton() {
  const { setDrawerOpen } = useScout();

  return (
    <button
      type="button"
      aria-label="Open menu"
      onClick={() => setDrawerOpen(true)}
      className="absolute left-5 top-5 z-20 flex h-12 w-12 cursor-pointer items-center justify-center rounded-2xl border border-blue-300 bg-blue-50 shadow-lg transition duration-300 hover:-translate-y-0.5 hover:bg-blue-100 hover:shadow-xl focus:outline-none focus:ring-4 focus:ring-blue-100"
    >
      <span className="flex w-5 flex-col gap-1.5">
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
      </span>
    </button>
  );
}

export default function SmartScouterHome() {
  return (
    <main className="relative flex min-h-screen overflow-hidden bg-white px-4 py-8 text-slate-950 sm:px-8 sm:py-10">
      <div
        className="research-candle-backdrop"
        style={
          {
            "--theme-accent": "#2563eb",
            "--theme-accent-dark": "#1e3a8a",
            "--theme-shadow": "rgba(37, 99, 235, 0.18)",
          } as CSSProperties
        }
        aria-hidden="true"
      >
        {Array.from({ length: 18 }).map((_, index) => (
          <span key={index} />
        ))}
      </div>
      <Drawer />
      <HamburgerButton />

      <section className="relative z-10 mx-auto flex w-full max-w-6xl animate-page-enter flex-col items-center justify-start gap-5 pt-20 sm:gap-7 sm:pt-16 lg:pt-10">
        <header className="space-y-3 text-center sm:space-y-4">
          <div className="flex justify-center">
            <SmartScouterLogo />
          </div>
          <h1 className="font-serif text-4xl font-bold tracking-normal text-slate-950 sm:text-6xl lg:text-7xl">
            Smart Scouter
          </h1>
          <p className="text-sm font-medium text-slate-800/80 sm:text-lg">
            Company Search and Investment Readiness Engine
          </p>
        </header>

        <nav className="grid w-full grid-cols-2 gap-3 rounded-[1.25rem] border border-slate-200 bg-white p-3 shadow-[0_24px_80px_rgba(15,23,42,0.12)] sm:grid-cols-5 sm:gap-4 sm:p-4 lg:gap-5 lg:p-5">
          {homeCards.map((card) => (
            <FeatureCard key={card.href} {...card} />
          ))}
        </nav>

        <div className="text-center">
          <p className="font-serif text-2xl font-black tracking-normal text-slate-950 sm:text-3xl">
            GreenFlow Ventures Ltd.
          </p>
          <p className="mt-2 text-xs font-black uppercase tracking-[0.28em] text-slate-800/70">
            Venture Intelligence Dashboard
          </p>
        </div>
      </section>
    </main>
  );
}
