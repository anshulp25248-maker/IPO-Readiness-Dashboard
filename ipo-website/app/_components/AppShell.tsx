"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { factorKeys, factorLabels } from "../_data/companies";
import { useScout } from "./ScoutProvider";

const navItems = [
  { label: "Dashboard", href: "/dashboard" },
  { label: "CDR", href: "/cdr" },
  { label: "Top Companies", href: "/top-10-companies" },
  { label: "Director Profile", href: "/director-profile" },
  { label: "Competitors", href: "/competitors" },
];

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

export function Drawer() {
  const router = useRouter();
  const {
    drawerOpen,
    setDrawerOpen,
    pendingFactorSelection,
    includedFactorCount,
    toggleFactor,
    runScoring,
    handleUpload,
    resetCompanies,
    uploadStatus,
    scoringStatus,
    companies,
  } = useScout();

  return (
    <>
      <div
        className={`fixed inset-0 z-40 bg-slate-950/35 backdrop-blur-sm transition-opacity duration-300 ${
          drawerOpen ? "opacity-100" : "pointer-events-none opacity-0"
        }`}
        onClick={() => setDrawerOpen(false)}
      />
      <aside
        className={`fixed left-0 top-0 z-50 h-full w-[min(92vw,390px)] overflow-y-auto border-r border-white/45 bg-[linear-gradient(135deg,rgba(31,182,255,0.60),rgba(52,211,153,0.60),rgba(250,204,21,0.60))] p-5 shadow-[0_28px_90px_rgba(15,23,42,0.28)] backdrop-blur-2xl transition-transform duration-300 ${
          drawerOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-start justify-between gap-4">
          <div>
            <p className="text-xs font-bold uppercase tracking-[0.22em] text-slate-800/70">Controls</p>
            <h2 className="font-serif text-3xl font-bold text-slate-950">Scout Smarter</h2>
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
          <p className="text-sm font-black text-slate-950">Upload Company File</p>
          <p className="mt-1 text-xs font-semibold leading-5 text-slate-800/75">
            Upload MCA Excel, CSV, PDF, JSON, HTML, TSV, or other pandas-readable company files.
          </p>
          <label className="mt-4 flex cursor-pointer items-center justify-center rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950">
            Upload File
            <input
              type="file"
              accept=".csv,.tsv,.txt,.xlsx,.xls,.xlsm,.ods,.json,.html,.htm,.pdf,.parquet,.feather,.pkl,.pickle"
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
            <p className="text-sm font-black text-slate-950">Company Scoring Factors</p>
            <span className="rounded-full bg-slate-950 px-3 py-1 text-xs font-black text-white">
              {includedFactorCount}/7
            </span>
          </div>
          <p className="mt-1 text-xs font-semibold text-slate-800/75">Enabled factors share equal weight.</p>
          <div className="mt-4 grid gap-3">
            {factorKeys.map((key) => (
              <button
                key={key}
                type="button"
                onClick={() => toggleFactor(key)}
                className="flex items-center justify-between rounded-xl border border-white/45 bg-white/45 px-4 py-3 text-left shadow backdrop-blur-xl transition hover:bg-white/65"
              >
                <span className="text-sm font-bold text-slate-950">{factorLabels[key]}</span>
                <span
                  className={`h-6 w-11 rounded-full p-1 transition ${
                    pendingFactorSelection[key] ? "bg-slate-950" : "bg-white/35"
                  }`}
                >
                  <span
                    className={`block h-4 w-4 rounded-full bg-white transition ${
                      pendingFactorSelection[key] ? "translate-x-5" : "translate-x-0"
                    }`}
                  />
                </span>
              </button>
            ))}
          </div>
          <button
            type="button"
            onClick={async () => {
              await runScoring();
              setDrawerOpen(false);
              router.push("/dashboard");
            }}
            className="mt-4 w-full rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950"
          >
            Run Scoring
          </button>
          <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">{scoringStatus}</p>
        </div>

        <nav className="mt-5 grid gap-2 rounded-2xl border border-white/45 bg-white/50 p-3 shadow-lg backdrop-blur-xl">
          {navItems.map((item) => (
            <Link
              key={item.href}
              href={item.href}
              onClick={() => setDrawerOpen(false)}
              className="rounded-xl px-4 py-3 text-sm font-black text-slate-950 transition hover:bg-white/65"
            >
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

export function AppShell({ title, eyebrow = "Scout Smarter", children }: AppShellProps) {
  return (
    <main className="min-h-screen bg-[linear-gradient(135deg,rgba(31,182,255,0.60),rgba(52,211,153,0.60),rgba(250,204,21,0.60))] px-4 py-5 text-slate-950 sm:px-6 lg:px-8">
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
              className="rounded-xl px-4 py-2 text-sm font-semibold text-slate-900 transition hover:bg-white/65"
            >
              {item.label}
            </Link>
          ))}
        </nav>
      </header>

      <section className="mx-auto mt-7 max-w-7xl animate-page-enter">{children}</section>
    </main>
  );
}
