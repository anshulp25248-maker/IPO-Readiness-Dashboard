"use client";

import type { CSSProperties } from "react";
import Link from "next/link";
import { usePathname, useRouter } from "next/navigation";
import { factorKeys, factorLabels } from "../_data/companies";
import { useScout } from "./ScoutProvider";
import { SmartScouterLogo } from "./SmartScouterBrand";

type AppIconName = "dashboard" | "report" | "ranking" | "profile" | "competitors" | "upload" | "weights";

const navItems = [
  { label: "Dashboard", href: "/dashboard", icon: "dashboard", color: "#2563eb", soft: "#dbeafe" },
  { label: "CDR", href: "/cdr", icon: "report", color: "#7c3aed", soft: "#ede9fe" },
  { label: "Other Companies", href: "/top-10-companies", icon: "ranking", color: "#0f9f6e", soft: "#d1fae5" },
  { label: "Director Profile", href: "/director-profile", icon: "profile", color: "#ea580c", soft: "#ffedd5" },
  { label: "Competitors", href: "/competitors", icon: "competitors", color: "#db2777", soft: "#fce7f3" },
] as const;

const pageThemes = {
  "/dashboard": {
    accent: "#2563eb",
    accentDark: "#1e3a8a",
    panel: "#dbeafe",
    panelStrong: "#bfdbfe",
    border: "#60a5fa",
    shadow: "rgba(37, 99, 235, 0.22)",
  },
  "/cdr": {
    accent: "#7c3aed",
    accentDark: "#4c1d95",
    panel: "#ede9fe",
    panelStrong: "#ddd6fe",
    border: "#a78bfa",
    shadow: "rgba(124, 58, 237, 0.22)",
  },
  "/top-10-companies": {
    accent: "#0f9f6e",
    accentDark: "#065f46",
    panel: "#d1fae5",
    panelStrong: "#a7f3d0",
    border: "#34d399",
    shadow: "rgba(15, 159, 110, 0.2)",
  },
  "/director-profile": {
    accent: "#ea580c",
    accentDark: "#9a3412",
    panel: "#ffedd5",
    panelStrong: "#fed7aa",
    border: "#fb923c",
    shadow: "rgba(234, 88, 12, 0.2)",
  },
  "/competitors": {
    accent: "#db2777",
    accentDark: "#9d174d",
    panel: "#fce7f3",
    panelStrong: "#fbcfe8",
    border: "#f472b6",
    shadow: "rgba(219, 39, 119, 0.2)",
  },
};

function AppIcon({ name, className = "h-5 w-5" }: { name: AppIconName; className?: string }) {
  const iconClassName = `${className} drop-shadow-[0_5px_8px_rgba(15,23,42,0.18)]`;
  if (name === "dashboard") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 13h6V4H4v9Z" />
        <path d="M14 20h6V4h-6v16Z" />
        <path d="M4 20h6v-3H4v3Z" />
      </svg>
    );
  }

  if (name === "report") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M7 3h7l4 4v14H7V3Z" />
        <path d="M14 3v5h5" />
        <path d="M10 12h6" />
        <path d="M10 16h4" />
      </svg>
    );
  }

  if (name === "ranking") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M5 20v-6h4v6H5Z" />
        <path d="M10 20V4h4v16h-4Z" />
        <path d="M15 20v-9h4v9h-4Z" />
      </svg>
    );
  }

  if (name === "profile") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z" />
        <path d="M4 21a8 8 0 0 1 16 0" />
      </svg>
    );
  }

  if (name === "upload") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 16V4" />
        <path d="M7 9l5-5 5 5" />
        <path d="M5 20h14" />
      </svg>
    );
  }

  if (name === "weights") {
    return (
      <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M5 6h14" />
        <path d="M5 12h14" />
        <path d="M5 18h14" />
        <path d="M8 4v4" />
        <path d="M15 10v4" />
        <path d="M11 16v4" />
      </svg>
    );
  }

  return (
    <svg className={iconClassName} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
      <path d="M8 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z" />
      <path d="M16 13a3 3 0 1 0 0-6 3 3 0 0 0 0 6Z" />
      <path d="M2 21a6 6 0 0 1 12 0" />
      <path d="M14 20a5 5 0 0 1 8 0" />
    </svg>
  );
}

function Hamburger() {
  const { setDrawerOpen } = useScout();

  return (
    <button
      type="button"
      aria-label="Main menu"
      onClick={() => setDrawerOpen(true)}
      className="flex h-11 w-11 items-center justify-center rounded-xl border border-[var(--panel-border)] bg-[var(--panel)] shadow-[0_14px_30px_var(--theme-shadow)] transition hover:-translate-y-0.5 hover:bg-[var(--panel-strong)]"
    >
      <span className="flex w-5 flex-col gap-1.5">
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
      </span>
    </button>
  );
}

export function Drawer() {
  const router = useRouter();
  const {
    drawerOpen,
    setDrawerOpen,
    pendingFactorMarks,
    setPendingFactorWeight,
    runScoring,
    handleUpload,
    resetCompanies,
    uploadStatus,
    scoringStatus,
    companies,
    aiProgress,
  } = useScout();
  const selectedFactorCount = factorKeys.filter((key) => (pendingFactorMarks[key] ?? 0) > 0).length;
  const marksAreReady = selectedFactorCount > 0;

  return (
    <>
      <div
        className={`fixed inset-0 z-40 bg-slate-950/35 backdrop-blur-sm transition-opacity duration-300 ${
          drawerOpen ? "opacity-100" : "pointer-events-none opacity-0"
        }`}
        onClick={() => setDrawerOpen(false)}
      />
      <aside
        className={`fixed left-0 top-0 z-50 h-full w-[min(92vw,390px)] overflow-y-auto border-r border-[var(--panel-border)] bg-white p-5 shadow-[0_28px_90px_var(--theme-shadow)] transition-transform duration-300 ${
          drawerOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-start justify-between gap-4">
          <SmartScouterLogo />
          <button
            type="button"
            onClick={() => setDrawerOpen(false)}
            className="rounded-xl border border-[var(--panel-border)] bg-[var(--panel)] px-3 py-2 text-sm font-black shadow transition hover:bg-[var(--panel-strong)]"
          >
            Close
          </button>
        </div>

        <div className="mt-6 rounded-2xl border border-[var(--panel-border)] bg-[var(--panel)] p-4 shadow-lg">
          <p className="flex items-center gap-2 text-sm font-black text-slate-950">
            <AppIcon name="upload" className="h-5 w-5" />
            Upload Company File
          </p>
          <p className="mt-1 text-xs font-semibold leading-5 text-slate-800/75">
            Upload MCA Excel, CSV, TSV, or JSON company files.
          </p>
          <label className="mt-4 flex cursor-pointer items-center justify-center rounded-xl bg-[var(--theme-accent)] px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-[var(--theme-accent-dark)]">
            Upload File
            <input
              type="file"
              accept=".csv,.tsv,.xlsx,.xls,.xlsm,.ods,.json"
              className="hidden"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) void handleUpload(file);
                event.currentTarget.value = "";
              }}
            />
          </label>
          <button
            type="button"
            onClick={resetCompanies}
            className="mt-3 w-full rounded-xl border border-[var(--panel-border)] bg-[var(--panel-strong)] px-4 py-3 text-sm font-black text-slate-950 shadow transition hover:bg-[var(--panel)]"
          >
            Reset Sample Data
          </button>
          <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">{uploadStatus}</p>
          <p className="mt-2 text-xs font-black text-[var(--theme-accent-dark)]">{companies.length} companies in workspace</p>
        </div>

        <div className="mt-5 rounded-2xl border border-[var(--panel-border)] bg-[var(--panel)] p-4 shadow-lg">
          <div className="flex items-center justify-between gap-3">
            <p className="flex items-center gap-2 text-sm font-black text-slate-950">
              <AppIcon name="weights" className="h-5 w-5" />
              V2 Factor Marks
            </p>
            <span className="rounded-full bg-slate-950 px-3 py-1 text-xs font-black text-white">
              /10
            </span>
          </div>
          <p className="mt-1 text-xs font-semibold text-slate-800/75">
            Mark each factor from 0 to 10. 0 removes the factor; scoring is normalized to 100 automatically.
          </p>
          <div className="mt-4 grid gap-3">
            {factorKeys.map((key) => {
              const value = pendingFactorMarks[key] ?? 0;
              const fill = value * 10;

              return (
                <label
                  key={key}
                  className="rounded-xl border border-[var(--panel-border)] bg-[var(--panel-strong)] px-4 py-3 shadow"
                >
                  <span className="flex items-center justify-between gap-3">
                    <span className="text-sm font-bold text-slate-950">{factorLabels[key]}</span>
                    <span className="rounded-full bg-slate-950 px-3 py-1 font-mono text-sm font-black text-white">
                      {value}/10
                    </span>
                  </span>
                  <input
                    type="range"
                    min="0"
                    max="10"
                    step="1"
                    value={value}
                    onChange={(event) => setPendingFactorWeight(key, Number(event.target.value))}
                    className="mt-3 h-2 w-full cursor-pointer appearance-none rounded-full accent-slate-950"
                    style={{
                      background: `linear-gradient(90deg, var(--theme-accent) 0%, var(--theme-accent) ${fill}%, var(--panel) ${fill}%, var(--panel) 100%)`,
                    }}
                  />
                  <span className="mt-2 flex items-center justify-between text-[11px] font-black uppercase tracking-[0.12em] text-slate-800/65">
                    <span>0</span>
                    <span>10</span>
                  </span>
                </label>
              );
            })}
          </div>
          <p className="mt-3 text-xs font-black text-slate-950">
            {marksAreReady ? `${selectedFactorCount} factors selected. Final score is calculated out of 100.` : "Select at least one factor to score."}
          </p>
          <button
            type="button"
            disabled={!marksAreReady || aiProgress.running}
            onClick={() => {
              setDrawerOpen(false);
              router.push("/dashboard");
              void runScoring();
            }}
            className="mt-4 w-full rounded-xl bg-[var(--theme-accent)] px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-[var(--theme-accent-dark)] disabled:cursor-not-allowed disabled:bg-slate-500 disabled:hover:translate-y-0"
          >
            {aiProgress.running ? "Scoring" : "Run Scoring"}
          </button>
          <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">{scoringStatus}</p>
          {aiProgress.total ? (
            <div className="mt-3">
              <div className="h-2 overflow-hidden rounded-full bg-[var(--panel)]">
                <div
                  className="h-full rounded-full bg-[var(--theme-accent)] animate-fill-sweep"
                  style={{ width: `${Math.round((aiProgress.completed / aiProgress.total) * 100)}%` }}
                />
              </div>
              <p className="mt-2 text-xs font-black text-slate-950">
                Screening {aiProgress.completed} of {aiProgress.total}
              </p>
            </div>
          ) : null}
        </div>

        <nav className="mt-5 grid gap-2 rounded-2xl border border-[var(--panel-border)] bg-white p-3 shadow-lg">
          {navItems.map((item) => (
            <Link
              key={item.href}
              href={item.href}
              onClick={() => setDrawerOpen(false)}
              className="flex items-center gap-3 rounded-xl px-4 py-3 text-sm font-black transition hover:-translate-y-0.5"
              style={{ backgroundColor: item.soft, color: item.color }}
            >
              <AppIcon name={item.icon} className="h-5 w-5" />
              {item.label}
            </Link>
          ))}
        </nav>
      </aside>
    </>
  );
}

type AppShellProps = {
  title: string;
  eyebrow?: string;
  children: React.ReactNode;
};

export function AppShell({ title, eyebrow = "Smart Scouter", children }: AppShellProps) {
  const pathname = usePathname();
  const theme = pageThemes[pathname as keyof typeof pageThemes] ?? pageThemes["/dashboard"];
  const shellStyle = {
    "--theme-accent": theme.accent,
    "--theme-accent-dark": theme.accentDark,
    "--panel": theme.panel,
    "--panel-strong": theme.panelStrong,
    "--panel-border": theme.border,
    "--theme-shadow": theme.shadow,
  } as CSSProperties;

  return (
    <main className="relative min-h-screen overflow-hidden bg-white px-4 py-5 text-slate-950 sm:px-6 lg:px-8" style={shellStyle}>
      <div className="research-candle-backdrop" aria-hidden="true">
        {Array.from({ length: 18 }).map((_, index) => (
          <span key={index} />
        ))}
      </div>
      <div className="relative z-10 min-h-screen transition duration-300">
        <Drawer />
        <header className="mx-auto flex max-w-7xl items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <Hamburger />
            <div className="flex items-center gap-4">
              <SmartScouterLogo compact />
              <div>
                <p className="text-xs font-bold uppercase tracking-[0.22em] text-slate-800/70">
                  {eyebrow}
                </p>
              <h1 className="font-serif text-3xl font-bold text-slate-950 sm:text-4xl">
                {title}
              </h1>
              </div>
            </div>
          </div>
          <nav className="hidden items-center gap-2 rounded-2xl border border-slate-200 bg-white p-1 shadow-[0_18px_42px_rgba(15,23,42,0.12)] lg:flex">
            {navItems.map((item) => (
              <Link
                key={item.href}
                href={item.href}
                className={`flex items-center gap-2 rounded-xl border px-4 py-2 text-sm font-bold transition hover:-translate-y-0.5 ${
                  pathname === item.href ? "text-white shadow-lg" : "border-transparent"
                }`}
                style={{
                  backgroundColor: pathname === item.href ? item.color : item.soft,
                  borderColor: pathname === item.href ? item.color : "transparent",
                  color: pathname === item.href ? "#ffffff" : item.color,
                }}
              >
                <AppIcon name={item.icon} className="h-4 w-4" />
                {item.label}
              </Link>
            ))}
          </nav>
        </header>

        <section className="mx-auto mt-7 max-w-7xl animate-page-enter">{children}</section>
      </div>
    </main>
  );
}
