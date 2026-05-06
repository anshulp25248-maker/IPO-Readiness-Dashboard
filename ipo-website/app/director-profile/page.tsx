"use client";

import { useState } from "react";
import { AppShell } from "../_components/AppShell";
import { GlassPanel } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";
import { Company } from "../_data/companies";

type ResearchState = {
  status: string;
  report: string;
  sources: Array<{ title?: string; url?: string }>;
};

function SearchLinks({ query }: { query: string }) {
  const encoded = encodeURIComponent(query);
  return (
    <div className="mt-3 flex flex-wrap gap-2">
      <a
        href={`https://www.linkedin.com/search/results/people/?keywords=${encoded}`}
        target="_blank"
        rel="noreferrer"
        className="inline-flex items-center gap-2 rounded-xl bg-[#0A66C2] px-3 py-2 text-xs font-black text-white shadow transition hover:-translate-y-0.5"
      >
        <span className="rounded bg-white px-1 font-black text-[#0A66C2]">in</span>
        LinkedIn
      </a>
      <a
        href={`https://www.google.com/search?q=${encoded}`}
        target="_blank"
        rel="noreferrer"
        className="rounded-xl bg-white/65 px-3 py-2 text-xs font-black text-slate-950 shadow transition hover:-translate-y-0.5 hover:bg-white/85"
      >
        Google
      </a>
      <a
        href={`https://www.zaubacorp.com/companysearchresults/${encoded}`}
        target="_blank"
        rel="noreferrer"
        className="rounded-xl bg-white/65 px-3 py-2 text-xs font-black text-slate-950 shadow transition hover:-translate-y-0.5 hover:bg-white/85"
      >
        Zauba
      </a>
    </div>
  );
}

export default function DirectorProfilePage() {
  const { rankedCompanies } = useScout();
  const topFive = rankedCompanies.slice(0, 5);
  const [cin, setCin] = useState("");
  const [research, setResearch] = useState<Record<string, ResearchState>>({});
  const [manualStatus, setManualStatus] = useState("Enter CIN to research directors for a company outside the top 5.");
  const [manualReport, setManualReport] = useState("");

  async function researchDirectors(company?: Company, manualCin?: string) {
    const key = company?.id || "manual";
    setResearch((current) => ({
      ...current,
      [key]: { status: "Searching MCA, Zauba, Tofler, LinkedIn/public feed...", report: "", sources: [] },
    }));
    if (!company) {
      setManualStatus("Searching MCA, Zauba, Tofler, LinkedIn/public feed...");
      setManualReport("");
    }

    try {
      const response = await fetch("/api/director-research", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ company, cin: manualCin }),
      });
      const data = (await response.json()) as ResearchState & { error?: string };
      if (!response.ok) throw new Error(data.error || "Director research failed.");

      if (company) {
        setResearch((current) => ({
          ...current,
          [key]: { status: "Director diligence generated.", report: data.report, sources: data.sources || [] },
        }));
      } else {
        setManualStatus("Director diligence generated.");
        setManualReport(data.report);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : "Director research failed.";
      if (company) {
        setResearch((current) => ({ ...current, [key]: { status: message, report: "", sources: [] } }));
      } else {
        setManualStatus(message);
      }
    }
  }

  return (
    <AppShell title="Director Profile">
      <GlassPanel className="mb-5">
        <div className="flex flex-col gap-3 lg:flex-row lg:items-end lg:justify-between">
          <div>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Manual Director Search</p>
            <h2 className="mt-2 font-serif text-3xl font-bold text-slate-950">Generate director research by CIN</h2>
          </div>
          <div className="grid gap-3 sm:grid-cols-[1fr_auto]">
            <input
              value={cin}
              onChange={(event) => setCin(event.target.value)}
              placeholder="Enter CIN"
              className="h-12 rounded-xl border border-white/45 bg-white/60 px-4 font-mono text-sm font-bold text-slate-950 outline-none focus:ring-4 focus:ring-white/50"
            />
            <button
              type="button"
              onClick={() => void researchDirectors(undefined, cin)}
              className="rounded-xl bg-slate-950 px-5 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950"
            >
              Find Directors
            </button>
          </div>
        </div>
        <p className="mt-3 rounded-xl bg-white/45 px-4 py-3 text-sm font-bold text-slate-900">{manualStatus}</p>
        {manualReport ? (
          <article className="mt-4 max-h-[520px] overflow-y-auto rounded-2xl bg-white/35 p-4">
            <ReportViewer report={manualReport} />
            <SearchLinks query={`${cin} directors`} />
          </article>
        ) : null}
      </GlassPanel>

      <div className="grid gap-5 lg:grid-cols-2">
        {topFive.map((company) => (
          <GlassPanel key={company.id}>
            <div className="flex items-start justify-between gap-4">
              <div>
                <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">{company.name}</p>
                <h2 className="mt-3 font-serif text-3xl font-bold text-slate-950">{company.director.name}</h2>
                <p className="mt-1 text-sm font-bold text-indigo-950">{company.director.role}</p>
              </div>
              <span className="rounded-xl bg-slate-950 px-3 py-2 text-sm font-black text-white">
                {company.factors.directorProfile.toFixed(1)}
              </span>
            </div>
            <div className="mt-5 grid gap-3">
              <div className="rounded-xl bg-white/45 p-4">
                <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-800/65">Education</p>
                <p className="mt-2 font-semibold text-slate-950">{company.director.education}</p>
              </div>
              <div className="rounded-xl bg-white/45 p-4">
                <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-800/65">Directorships</p>
                <p className="mt-2 font-semibold text-slate-950">{company.director.directorships}</p>
              </div>
              <div className="rounded-xl bg-white/45 p-4">
                <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-800/65">Credibility Signal</p>
                <p className="mt-2 font-semibold leading-6 text-slate-950">{company.director.credibility}</p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => void researchDirectors(company)}
              className="mt-5 w-full rounded-xl bg-slate-950 px-4 py-3 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950"
            >
              Generate AI Director Diligence
            </button>
            {research[company.id]?.status ? (
              <p className="mt-3 rounded-xl bg-white/45 px-4 py-3 text-sm font-bold text-slate-900">
                {research[company.id].status}
              </p>
            ) : null}
            {research[company.id]?.report ? (
              <article className="mt-4 max-h-[520px] overflow-y-auto rounded-2xl bg-white/35 p-4">
                <ReportViewer report={research[company.id].report} />
                <SearchLinks query={`${company.director.name !== "NA" ? company.director.name : "directors"} ${company.name} ${company.cin}`} />
              </article>
            ) : null}
          </GlassPanel>
        ))}
      </div>
    </AppShell>
  );
}
