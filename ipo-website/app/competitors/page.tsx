"use client";

import { useState } from "react";
import { AppShell } from "../_components/AppShell";
import { GlassPanel } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";

export default function CompetitorsPage() {
  const { rankedCompanies, topCompany } = useScout();
  const [selectedCompanyId, setSelectedCompanyId] = useState(topCompany.id);
  const selectedCompany = rankedCompanies.find((company) => company.id === selectedCompanyId) || topCompany;
  const [status, setStatus] = useState("Ready to analyze competitors for the selected company.");
  const [report, setReport] = useState("");
  const [sources, setSources] = useState<Array<{ title?: string; url?: string }>>([]);
  const [isGenerating, setIsGenerating] = useState(false);

  async function generateCompetitors() {
    setIsGenerating(true);
    setStatus("Running competitor Groq + live-feed lane...");
    setReport("");
    setSources([]);

    try {
      const response = await fetch("/api/competitor-research", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ company: selectedCompany, peers: rankedCompanies }),
      });
      const data = (await response.json()) as {
        report?: string;
        sources?: Array<{ title?: string; url?: string }>;
        peerCount?: number;
        error?: string;
      };
      if (!response.ok) throw new Error(data.error || "Competitor research failed.");
      setReport(data.report || "No competitor report returned.");
      setSources(data.sources || []);
      setStatus("Competitor diligence generated from uploaded peers, Groq analysis, and live-feed evidence.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Competitor research failed.");
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <AppShell title="Competitors">
      <GlassPanel className="mb-5">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Competitor Diligence</p>
            <h2 className="mt-2 font-serif text-4xl font-bold text-slate-950">{selectedCompany.name}</h2>
            <p className="mt-2 text-sm font-semibold text-slate-800">{selectedCompany.activity}</p>
          </div>
          <button
            type="button"
            onClick={generateCompetitors}
            disabled={isGenerating}
            className="rounded-xl bg-slate-950 px-5 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {isGenerating ? "Generating..." : "Generate Competitor Report"}
          </button>
        </div>
        <label className="mt-5 grid gap-2">
          <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
            Choose Company
          </span>
          <select
            value={selectedCompany.id}
            onChange={(event) => {
              setSelectedCompanyId(event.target.value);
              setReport("");
              setSources([]);
              setStatus("Ready to analyze competitors for the selected company.");
            }}
            className="h-12 rounded-xl border border-amber-400/55 bg-white/60 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none transition focus:ring-4 focus:ring-white/50"
          >
            {rankedCompanies.map((company) => (
              <option key={company.id} value={company.id}>
                {company.name} - {company.cin}
              </option>
            ))}
          </select>
        </label>
        <p className="mt-4 rounded-xl bg-white/45 px-4 py-3 text-sm font-bold text-slate-900">{status}</p>
        {report ? (
          <div className="mt-4 grid gap-4 lg:grid-cols-[1.35fr_0.65fr]">
            <article className="max-h-[620px] overflow-y-auto rounded-2xl bg-white/35 p-4">
              <ReportViewer report={report} />
            </article>
            <div className="rounded-2xl bg-white/45 p-4">
              <p className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">Sources</p>
              <div className="mt-3 grid gap-2">
                {sources.slice(0, 8).map((source, index) => (
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
      </GlassPanel>

      <div className="grid gap-5 lg:grid-cols-3">
        {rankedCompanies.slice(0, 12).map((company) => (
          <GlassPanel key={company.id}>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">{company.sector}</p>
            <h2 className="mt-3 min-h-20 font-serif text-3xl font-bold leading-tight text-slate-950">
              {company.name}
            </h2>
            <div className="mt-5 grid gap-3">
              {company.competitors.map((competitor) => (
                <div key={competitor} className="rounded-xl border border-white/40 bg-white/45 px-4 py-3 font-bold text-slate-950">
                  {competitor}
                </div>
              ))}
            </div>
          </GlassPanel>
        ))}
      </div>
    </AppShell>
  );
}
