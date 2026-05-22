"use client";

import { useMemo, useState } from "react";
import { AppShell } from "../_components/AppShell";
import { CompanyTable, GlassPanel } from "../_components/CompanyWidgets";
import { useScout } from "../_components/ScoutProvider";
import { Company } from "../_data/companies";

type AiSearchResponse = {
  report?: string;
  companies?: Company[];
  sources?: Array<{ title?: string; url?: string }>;
  generatedAt?: string;
  error?: string;
};

export default function Top10CompaniesPage() {
  const { rankedCompanies, includedFactorCount, replaceCompanies } = useScout();
  const [query, setQuery] = useState("");
  const [city, setCity] = useState("All");
  const [state, setState] = useState("All");
  const [aiQuery, setAiQuery] = useState("");
  const [aiCity, setAiCity] = useState("");
  const [aiState, setAiState] = useState("");
  const [aiCin, setAiCin] = useState("");
  const [aiReport, setAiReport] = useState("");
  const [aiSources, setAiSources] = useState<Array<{ title?: string; url?: string }>>([]);
  const [aiStatus, setAiStatus] = useState("Ready for company research.");
  const [isSearching, setIsSearching] = useState(false);

  const cities = useMemo(
    () => ["All", ...Array.from(new Set(rankedCompanies.map((company) => company.city))).sort()],
    [rankedCompanies],
  );
  const states = useMemo(
    () => ["All", ...Array.from(new Set(rankedCompanies.map((company) => company.state))).sort()],
    [rankedCompanies],
  );

  const filteredCompanies = useMemo(() => {
    const search = query.trim().toLowerCase();
    return rankedCompanies.filter((company) => {
      const matchesSearch =
        !search ||
        company.name.toLowerCase().includes(search) ||
        company.cin.toLowerCase().includes(search);
      const matchesCity = city === "All" || company.city === city;
      const matchesState = state === "All" || company.state === state;
      return matchesSearch && matchesCity && matchesState;
    });
  }, [city, query, rankedCompanies, state]);

  async function runAiSearch() {
    setIsSearching(true);
    setAiStatus("Generating company analysis from the supplied query and AI model...");
    setAiReport("");
    setAiSources([]);

    try {
      const response = await fetch("/api/company-search", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          query: aiQuery,
          city: aiCity,
          state: aiState,
          cin: aiCin,
          companyName: aiQuery,
        }),
      });
      const data = (await response.json()) as AiSearchResponse;

      if (!response.ok) {
        throw new Error(data.error || "Company research failed.");
      }

      const companies = data.companies || [];
      if (companies.length) {
        replaceCompanies(
          companies,
          `${companies.length} AI-researched companies loaded. Use drawer factors and Run Scoring to rescore.`,
        );
        setAiStatus(`${companies.length} companies found and loaded into the dashboard workspace.`);
      } else {
        setAiStatus("Company research completed but no company candidates were found.");
      }
      setAiReport(data.report || "No report returned.");
      setAiSources(data.sources || []);
    } catch (error) {
      setAiStatus(error instanceof Error ? error.message : "Company research failed.");
    } finally {
      setIsSearching(false);
    }
  }

  return (
    <AppShell title="Other Companies">
      <GlassPanel>
        <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
          <div>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Ranked Output</p>
            <h2 className="mt-2 font-serif text-4xl font-bold text-slate-950">
              Best Companies From Uploaded Universe
            </h2>
          </div>
          <p className="rounded-xl bg-white/45 px-4 py-3 text-sm font-bold text-slate-900">
            {includedFactorCount} equal-weight factors included
          </p>
        </div>

        <div className="mt-6 rounded-2xl border border-white/40 bg-white/50 p-4 shadow-lg backdrop-blur-xl">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">
                Research + Feed Parser
              </p>
              <h3 className="mt-2 font-serif text-3xl font-bold text-slate-950">
                Search real-time company feeds
              </h3>
            </div>
            <button
              type="button"
              onClick={runAiSearch}
              disabled={isSearching}
              className="rounded-xl bg-slate-950 px-5 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950 disabled:cursor-not-allowed disabled:opacity-60"
            >
              {isSearching ? "Searching..." : "Generate Company Report"}
            </button>
          </div>

          <div className="mt-5 grid gap-3 lg:grid-cols-[1.2fr_0.7fr_0.7fr_0.8fr]">
            <label className="grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                Company Name / Theme
              </span>
              <input
                value={aiQuery}
                onChange={(event) => setAiQuery(event.target.value)}
                placeholder="e.g. semiconductor, SaaS, Reliance Retail"
                className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
              />
            </label>
            <label className="grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                City
              </span>
              <input
                value={aiCity}
                onChange={(event) => setAiCity(event.target.value)}
                placeholder="Mumbai"
                className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
              />
            </label>
            <label className="grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                State
              </span>
              <input
                value={aiState}
                onChange={(event) => setAiState(event.target.value)}
                placeholder="Maharashtra"
                className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
              />
            </label>
            <label className="grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                CIN
              </span>
              <input
                value={aiCin}
                onChange={(event) => setAiCin(event.target.value)}
                placeholder="U..."
                className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
              />
            </label>
          </div>

          <p className="mt-4 rounded-xl bg-white/45 px-4 py-3 text-sm font-bold text-slate-900">
            {aiStatus}
          </p>

          {aiReport ? (
            <div className="mt-4 grid gap-4 lg:grid-cols-[1.35fr_0.65fr]">
              <article className="max-h-96 overflow-y-auto whitespace-pre-wrap rounded-2xl border border-white/40 bg-white/65 p-5 text-sm font-semibold leading-6 text-slate-900 shadow-inner">
                {aiReport}
              </article>
              <div className="rounded-2xl border border-white/40 bg-white/55 p-4">
                <p className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                  Sources
                </p>
                <div className="mt-3 grid gap-2">
                  {aiSources.slice(0, 6).map((source, index) => (
                    <a
                      key={`${source.url || source.title}-${index}`}
                      href={source.url}
                      target="_blank"
                      rel="noreferrer"
                      className="rounded-xl bg-white/55 px-3 py-2 text-xs font-bold leading-5 text-indigo-950 transition hover:bg-white/80"
                    >
                      {source.title || source.url || "Source"}
                    </a>
                  ))}
                </div>
              </div>
            </div>
          ) : null}
        </div>

        <div className="mt-6 grid gap-3 rounded-2xl border border-white/35 bg-white/45 p-4 backdrop-blur-xl lg:grid-cols-[1.4fr_0.8fr_0.8fr]">
          <label className="grid gap-2">
            <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
              Search Company / CIN
            </span>
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Type company name or CIN"
              className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
            />
          </label>

          <label className="grid gap-2">
            <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
              City
            </span>
            <select
              value={city}
              onChange={(event) => setCity(event.target.value)}
              className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
            >
              {cities.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>
          </label>

          <label className="grid gap-2">
            <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
              State
            </span>
            <select
              value={state}
              onChange={(event) => setState(event.target.value)}
              className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
            >
              {states.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>
          </label>
        </div>

        <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
          <p className="rounded-xl bg-white/45 px-4 py-2 text-sm font-black text-slate-900">
            Showing {filteredCompanies.length} of {rankedCompanies.length} companies
          </p>
          <button
            type="button"
            onClick={() => {
              setQuery("");
              setCity("All");
              setState("All");
            }}
            className="rounded-xl border border-white/45 bg-white/45 px-4 py-2 text-sm font-black text-slate-950 shadow transition hover:bg-white/65"
          >
            Clear Search
          </button>
        </div>

        <div className="mt-6">
          <CompanyTable companies={filteredCompanies} />
        </div>
      </GlassPanel>
    </AppShell>
  );
}
