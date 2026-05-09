"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { factorKeys, factorLabels } from "../_data/companies";
import { useScout } from "./ScoutProvider";

type AppIconName = "dashboard" | "report" | "ranking" | "profile" | "competitors" | "upload" | "weights";

const navItems = [
  { label: "Dashboard", href: "/dashboard", icon: "dashboard" },
  { label: "CDR", href: "/cdr", icon: "report" },
  { label: "Top Companies", href: "/top-10-companies", icon: "ranking" },
  { label: "Director Profile", href: "/director-profile", icon: "profile" },
  { label: "Competitors", href: "/competitors", icon: "competitors" },
] as const;

function AppIcon({ name, className = "h-5 w-5" }: { name: AppIconName; className?: string }) {
  if (name === "dashboard") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 13h6V4H4v9Z" />
        <path d="M14 20h6V4h-6v16Z" />
        <path d="M4 20h6v-3H4v3Z" />
      </svg>
    );
  }

  if (name === "report") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M7 3h7l4 4v14H7V3Z" />
        <path d="M14 3v5h5" />
        <path d="M10 12h6" />
        <path d="M10 16h4" />
      </svg>
    );
  }

  if (name === "ranking") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M5 20v-6h4v6H5Z" />
        <path d="M10 20V4h4v16h-4Z" />
        <path d="M15 20v-9h4v9h-4Z" />
      </svg>
    );
  }

  if (name === "profile") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z" />
        <path d="M4 21a8 8 0 0 1 16 0" />
      </svg>
    );
  }

  if (name === "upload") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 16V4" />
        <path d="M7 9l5-5 5 5" />
        <path d="M5 20h14" />
      </svg>
    );
  }

  if (name === "weights") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
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
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
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
      className="flex h-11 w-11 items-center justify-center rounded-xl border border-white/45 bg-white/50 shadow-lg backdrop-blur-xl transition hover:-translate-y-0.5 hover:bg-white/65"
    >
      <span className="flex w-5 flex-col gap-1.5">
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
        <span className="h-0.5 rounded-full bg-slate-950" />
      </span>
    </button>
  );
}

function ScoringOverlay({ completed, total }: { completed: number; total: number }) {
  const progress = total > 0 ? Math.min(100, Math.round((completed / total) * 100)) : 0;
  const circumference = 2 * Math.PI * 48;
  const offset = circumference - (progress / 100) * circumference;

  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center bg-slate-950/20 backdrop-blur-[2px]"
      role="status"
      aria-live="polite"
      aria-label={`AI scoring ${progress}% complete`}
    >
      <div className="relative flex h-44 w-44 items-center justify-center rounded-full border border-white/45 bg-white/70 shadow-[0_28px_90px_rgba(15,23,42,0.28)] backdrop-blur-2xl">
        <svg className="absolute inset-0 h-full w-full -rotate-90 p-4" viewBox="0 0 120 120" aria-hidden="true">
          <circle
            cx="60"
            cy="60"
            r="48"
            fill="none"
            stroke="rgba(255,255,255,0.76)"
            strokeWidth="10"
          />
          <circle
            cx="60"
            cy="60"
            r="48"
            fill="none"
            stroke="rgb(15,23,42)"
            strokeLinecap="round"
            strokeWidth="10"
            strokeDasharray={circumference}
            strokeDashoffset={offset}
            className="transition-all duration-300"
          />
        </svg>
        <span className="absolute inset-4 animate-spin rounded-full border-2 border-transparent border-t-emerald-600" />
        <span className="font-mono text-4xl font-black text-slate-950">{progress}%</span>
      </div>
    </div>
  );
}

export function Drawer() {
  const router = useRouter();
  const {
    drawerOpen,
    setDrawerOpen,
    pendingFactorWeights,
    setPendingFactorWeight,
    runScoring,
    handleUpload,
    resetCompanies,
    uploadStatus,
    scoringStatus,
    companies,
    aiProgress,
  } = useScout();
  const pendingWeightTotal = factorKeys.reduce((sum, key) => sum + (pendingFactorWeights[key] ?? 0), 0);
  const remainingWeight = Math.max(0, 100 - pendingWeightTotal);
  const weightsAreReady = Math.abs(pendingWeightTotal - 100) < 0.001;

  return (
    <>
      <div
        className={`fixed inset-0 z-40 bg-slate-950/35 backdrop-blur-sm transition-opacity duration-300 ${
          drawerOpen ? "opacity-100" : "pointer-events-none opacity-0"
        }`}
        onClick={() => setDrawerOpen(false)}
      />
      <aside
        className={`fixed left-0 top-0 z-50 h-full w-[min(92vw,390px)] overflow-y-auto border-r border-white/45 bg-[linear-gradient(135deg,rgba(125,211,252,0.78),rgba(153,246,228,0.74),rgba(220,252,231,0.78))] p-5 shadow-[0_28px_90px_rgba(15,23,42,0.24)] backdrop-blur-2xl transition-transform duration-300 ${
          drawerOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.22em] text-slate-800/70">Controls</p>
            <h2 className="font-serif text-3xl font-bold text-slate-950">Smart Scouter</h2>
          </div>
          <button
            type="button"
            onClick={() => setDrawerOpen(false)}
            className="rounded-xl border border-white/45 bg-white/50 px-3 py-2 text-sm font-black shadow backdrop-blur-xl transition hover:bg-white/65"
          >
            Close
          </button>
        </div>

        <div className="mt-6 rounded-2xl border border-white/45 bg-white/50 p-4 shadow-lg backdrop-blur-xl">
          <p className="flex items-center gap-2 text-sm font-black text-slate-950">
            <AppIcon name="upload" className="h-5 w-5" />
            Upload Company File
          </p>
          <p className="mt-1 text-xs font-semibold leading-5 text-slate-800/75">
            Upload MCA Excel, CSV, TSV, or JSON company files.
          </p>
          <label className="mt-4 flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950">
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
            className="mt-3 w-full rounded-xl border border-white/45 bg-white/45 px-4 py-3 text-sm font-black text-slate-950 shadow backdrop-blur-xl transition hover:bg-white/65"
          >
            Reset Sample Data
          </button>
          <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">{uploadStatus}</p>
          <p className="mt-2 text-xs font-black text-indigo-950">{companies.length} companies in workspace</p>
        </div>

        <div className="mt-5 rounded-2xl border border-white/45 bg-white/50 p-4 shadow-lg backdrop-blur-xl">
          <div className="flex items-center justify-between gap-3">
            <p className="flex items-center gap-2 text-sm font-black text-slate-950">
              <AppIcon name="weights" className="h-5 w-5" />
              V2 Factor Weights
            </p>
            <span
              className={`rounded-full px-3 py-1 text-xs font-black ${
                Math.abs(pendingWeightTotal - 100) < 0.001 ? "bg-slate-950 text-white" : "bg-rose-100 text-rose-950"
              }`}
            >
              {pendingWeightTotal}%
            </span>
          </div>
          <p className="mt-1 text-xs font-semibold text-slate-800/75">
            Each factor can be 0% to 30%. Total is capped at 100%.
          </p>
          <div className="mt-4 grid gap-3">
            {factorKeys.map((key) => {
              const value = pendingFactorWeights[key] ?? 0;
              const maxForSlider = Math.min(30, value + remainingWeight);
              const fill = maxForSlider > 0 ? Math.round((value / maxForSlider) * 100) : 0;

              return (
                <label
                  key={key}
                  className="rounded-xl border border-white/45 bg-white/45 px-4 py-3 shadow backdrop-blur-xl"
                >
                  <span className="flex items-center justify-between gap-3">
                    <span className="text-sm font-bold text-slate-950">{factorLabels[key]}</span>
                    <span className="rounded-full bg-slate-950 px-3 py-1 font-mono text-sm font-black text-white">
                      {value}%
                    </span>
                  </span>
                  <input
                    type="range"
                    min="0"
                    max={maxForSlider}
                    step="1"
                    value={value}
                    onChange={(event) => setPendingFactorWeight(key, Number(event.target.value))}
                    className="mt-3 h-2 w-full cursor-pointer appearance-none rounded-full accent-slate-950"
                    style={{
                      background: `linear-gradient(90deg, rgb(15,23,42) 0%, rgb(15,23,42) ${fill}%, rgba(255,255,255,0.7) ${fill}%, rgba(255,255,255,0.7) 100%)`,
                    }}
                  />
                  <span className="mt-2 flex items-center justify-between text-[11px] font-black uppercase tracking-[0.12em] text-slate-800/65">
                    <span>0%</span>
                    <span>Max {Math.round(maxForSlider)}%</span>
                  </span>
                </label>
              );
            })}
          </div>
          <p className="mt-3 text-xs font-black text-slate-950">
            {weightsAreReady ? "Total 100% matched" : `${remainingWeight}% left to match 100%`}
          </p>
          <button
            type="button"
            disabled={!weightsAreReady || aiProgress.running}
            onClick={() => {
              setDrawerOpen(false);
              router.push("/dashboard");
              void runScoring();
            }}
            className="mt-4 w-full rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950 disabled:cursor-not-allowed disabled:bg-slate-500 disabled:hover:translate-y-0"
          >
            {aiProgress.running ? "AI Running" : "Run Scoring"}
          </button>
          <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">{scoringStatus}</p>
          {aiProgress.total ? (
            <div className="mt-3">
              <div className="h-2 overflow-hidden rounded-full bg-white/40">
                <div
                  className="h-full rounded-full bg-slate-950"
                  style={{ width: `${Math.round((aiProgress.completed / aiProgress.total) * 100)}%` }}
                />
              </div>
              <p className="mt-2 text-xs font-black text-slate-950">
                AI screening {aiProgress.completed} of {aiProgress.total}
              </p>
            </div>
          ) : null}
        </div>

        <nav className="mt-5 grid gap-2 rounded-2xl border border-white/45 bg-white/50 p-3 shadow-lg backdrop-blur-xl">
          {navItems.map((item) => (
            <Link
              key={item.href}
              href={item.href}
              onClick={() => setDrawerOpen(false)}
              className="flex items-center gap-3 rounded-xl px-4 py-3 text-sm font-black text-slate-950 transition hover:bg-white/65"
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
  const { aiProgress } = useScout();
  const scoringOverlayOpen = aiProgress.running && aiProgress.total > 0;

  return (
    <main className="min-h-screen bg-[linear-gradient(135deg,rgba(125,211,252,0.76),rgba(153,246,228,0.72),rgba(220,252,231,0.76))] px-4 py-5 text-slate-950 sm:px-6 lg:px-8">
      <div
        className={`min-h-screen transition duration-300 ${
          scoringOverlayOpen ? "pointer-events-none select-none blur-md" : ""
        }`}
      >
        <Drawer />
        <header className="mx-auto flex max-w-7xl items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <Hamburger />
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.22em] text-slate-800/70">
                {eyebrow}
              </p>
              <h1 className="font-serif text-3xl font-bold text-slate-950 sm:text-4xl">
                {title}
              </h1>
            </div>
          </div>
          <nav className="hidden items-center gap-2 rounded-2xl border border-white/45 bg-white/50 p-1 shadow-lg backdrop-blur-xl lg:flex">
            {navItems.map((item) => (
              <Link
                key={item.href}
                href={item.href}
                className="flex items-center gap-2 rounded-xl px-4 py-2 text-sm font-semibold text-slate-900 transition hover:bg-white/65"
              >
                <AppIcon name={item.icon} className="h-4 w-4" />
                {item.label}
              </Link>
            ))}
          </nav>
        </header>

        <section className="mx-auto mt-7 max-w-7xl animate-page-enter">{children}</section>
      </div>
      {scoringOverlayOpen ? (
        <ScoringOverlay completed={aiProgress.completed} total={aiProgress.total} />
      ) : null}
    </main>
  );
}
