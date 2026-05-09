"use client";

import { FeatureCard } from "./_components/FeatureCard";
import { Drawer } from "./_components/AppShell";
import { useScout } from "./_components/ScoutProvider";

const primaryCards = [
  { title: "Dashboard", href: "/dashboard" },
  { title: "CDR", href: "/cdr" },
  { title: "Top 10 Companies", href: "/top-10-companies" },
];

const secondaryCards = [
  { title: "Director Profile", href: "/director-profile" },
  { title: "Competitors", href: "/competitors" },
];

function HamburgerButton() {
  const { setDrawerOpen } = useScout();

  return (
    <button
      type="button"
      aria-label="Open menu"
      onClick={() => setDrawerOpen(true)}
      className="absolute left-5 top-5 z-10 flex h-12 w-12 cursor-pointer items-center justify-center rounded-2xl border border-white/35 bg-white/15 shadow-lg backdrop-blur-xl transition duration-300 hover:-translate-y-0.5 hover:bg-white/25 hover:shadow-xl focus:outline-none focus:ring-4 focus:ring-white/40"
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
    <main className="relative flex min-h-screen overflow-hidden bg-[linear-gradient(135deg,rgba(31,182,255,0.60),rgba(52,211,153,0.60),rgba(250,204,21,0.60))] px-5 py-10 text-slate-950 sm:px-8">
      <Drawer />
      <HamburgerButton />

      <section className="mx-auto flex w-full max-w-6xl animate-page-enter flex-col items-center justify-center gap-12">
        <header className="space-y-5 text-center">
          <h1 className="bg-gradient-to-r from-slate-950 via-indigo-950 to-purple-900 bg-clip-text font-serif text-5xl font-bold tracking-normal text-transparent sm:text-6xl lg:text-7xl">
            Smart Scouter
          </h1>
          <p className="text-base font-medium text-slate-800/80 sm:text-xl">
            An AI Powered Company Search Engine
          </p>
        </header>

        <nav className="grid w-full gap-6">
          <div className="grid gap-6 md:grid-cols-3">
            {primaryCards.map((card) => (
              <FeatureCard key={card.href} {...card} />
            ))}
          </div>

          <div className="mx-auto grid w-full max-w-3xl gap-6 md:grid-cols-2">
            {secondaryCards.map((card) => (
              <FeatureCard key={card.href} {...card} />
            ))}
          </div>
        </nav>
      </section>
    </main>
  );
}
