"use client";

import { useEffect, useMemo, useState } from "react";
import { AppShell } from "../_components/AppShell";
import { FactorGrid, GlassPanel, ScoreBadge } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";
import { Company } from "../_data/companies";

const cdrTabs = [
  {
    key: "sector-analysis",
    label: "Sector Analysis",
    description: "Macro sector, government stance, thematic tailwinds, headwinds, and capital-market appetite.",
    sections: [
      "Sector definition and investable universe",
      "Government stance, policy, PLI, regulatory direction",
      "Market size, growth drivers, cyclicality, and demand pools",
      "Listed and unlisted sector benchmarks",
      "Sector-level red flags and diligence requests",
    ],
  },
  {
    key: "industry-analysis",
    label: "Industry Analysis",
    description: "NIC/basic-industry level outlook using industry reports, association publications, and public datasets.",
    sections: [
      "NIC/basic industry classification",
      "Industry structure and value chain",
      "Customer concentration and supplier power",
      "Margin profile, working capital pattern, and capex intensity",
      "Industry outlook and data gaps to verify",
    ],
  },
  {
    key: "competitor-analysis",
    label: "Competitor Analysis",
    description: "Peer set, positioning, differentiation, benchmark multiples, and substitution risk.",
    sections: [
      "Peer universe and closest comparables",
      "Relative scale, operating model, and product overlap",
      "Competitive moat and business-model uniqueness",
      "Valuation and funding benchmarks where public",
      "Threats, displacement risks, and diligence questions",
    ],
  },
  {
    key: "director-profile",
    label: "Director Profile",
    description: "Promoter/director credibility, directorship history, public background, conflicts, and governance flags.",
    sections: [
      "Director and promoter identity verification",
      "DIN, appointment, directorships, and disqualification checks",
      "Education, operating track record, and LinkedIn/public footprint",
      "Related-party, litigation, and governance concerns",
      "Management-quality investment view",
    ],
  },
  {
    key: "company-analysis",
    label: "Company Analysis",
    description: "Full company investment view with capital consistency, score reconciliation, risks, and final thesis.",
    sections: [
      "Company snapshot and capital source of truth",
      "Business model and operating model",
      "Financial profile and capital structure",
      "Quantitative Scout Smarter score reconciliation",
      "Investment thesis, monitorables, and recommendation",
    ],
  },
  {
    key: "docx-generation",
    label: "DOCX Generation",
    description: "Generate and export one detailed CDR built from all five section-specific CDR lanes.",
    sections: [
      "Sector, industry, competitor, director, and company sections",
      "Dedicated research and analysis lanes for each section",
      "Bold short headers with detailed body text",
      "Red formatting for risk flags",
      "Source feed appendix",
    ],
  },
] as const;

type CdrTabKey = (typeof cdrTabs)[number]["key"];
type CdrTabState = {
  report: string;
  sources: Array<{ title?: string; url?: string }>;
  status: string;
  isGenerating: boolean;
};

function makeInitialCdrState(): Record<CdrTabKey, CdrTabState> {
  return cdrTabs.reduce(
    (states, tab) => {
      states[tab.key] = {
        report: "",
        sources: [],
        status: "Ready to generate CDR from uploaded MCA data and live feed evidence.",
        isGenerating: false,
      };
      return states;
    },
    {} as Record<CdrTabKey, CdrTabState>,
  );
}

function makeCinOnlyCompany(cin: string): Company {
  const cleanCin = cin.trim().toUpperCase();
  return {
    id: `cin-only-${cleanCin || "company"}`,
    name: `Company linked to ${cleanCin || "searched CIN"}`,
    cin: cleanCin || "Data Not Available",
    sector: "Data Not Available",
    city: "Data Not Available",
    state: "Data Not Available",
    status: "Active",
    paidUpCapital: "Data Not Available",
    authorizedCapital: "Data Not Available",
    paidUpCapitalValue: 0,
    authorizedCapitalValue: 0,
    nicCode: "Data Not Available",
    activity: "Data Not Available",
    lastFiling: "Data Not Available",
    director: {
      name: "Data Not Available",
      role: "Data Not Available",
      education: "Data Not Available",
      directorships: 0,
      credibility: "Public verification required",
    },
    factors: {
      paidUpCapital: 0,
      sector: 0,
      geography: 0,
      businessModel: 0,
      directorProfile: 0,
      capitalRatio: 0,
      filingCompliance: 0,
    },
    competitors: [],
  };
}

export default function CdrPage() {
  const { topCompany, rankedCompanies, scoreCompany } = useScout();
  const [cdrState, setCdrState] = useState<Record<CdrTabKey, CdrTabState>>(makeInitialCdrState);
  const [internalInfo, setInternalInfo] = useState("");
  const [cdrSearch, setCdrSearch] = useState("");
  const [activeTab, setActiveTab] = useState<CdrTabKey>("sector-analysis");
  const [selectedCompanyId, setSelectedCompanyId] = useState(topCompany.id);
  const [externalCompany, setExternalCompany] = useState<Company | null>(null);

  function cleanStatus(message: string) {
    return message
      .replace(/\bGroq\b/gi, "analysis")
      .replace(/\bTavily\b/gi, "live feed")
      .replace(/\bGemini\b/gi, "analysis")
      .replace(/\bOpenRouter\b/gi, "analysis");
  }

  const searchResults = useMemo(() => {
    const search = cdrSearch.trim().toLowerCase();
    if (!search) return rankedCompanies.slice(0, 6);
    return rankedCompanies
      .filter((company) => company.name.toLowerCase().includes(search) || company.cin.toLowerCase().includes(search))
      .slice(0, 8);
  }, [cdrSearch, rankedCompanies]);

  const selectedCompany = useMemo(() => {
    return externalCompany || rankedCompanies.find((company) => company.id === selectedCompanyId) || topCompany;
  }, [externalCompany, rankedCompanies, selectedCompanyId, topCompany]);

  const activeSpec = useMemo(() => cdrTabs.find((tab) => tab.key === activeTab) || cdrTabs[0], [activeTab]);
  const activeState = cdrState[activeTab];
  const { report, sources, status, isGenerating } = activeState;

  function updateTabState(tabKey: CdrTabKey, updates: Partial<CdrTabState>) {
    setCdrState((current) => ({
      ...current,
      [tabKey]: {
        ...current[tabKey],
        ...updates,
      },
    }));
  }

  useEffect(() => {
    setCdrState(makeInitialCdrState());
  }, [selectedCompany.id]);

  useEffect(() => {
    if (!rankedCompanies.some((company) => company.id === selectedCompanyId)) {
      setSelectedCompanyId(topCompany.id);
      setExternalCompany(null);
    }
  }, [rankedCompanies, selectedCompanyId, topCompany.id]);

  async function generateCdr() {
    const tabKey = activeTab;
    const tabSpec = activeSpec;
    const company = selectedCompany;
    const companyScore = scoreCompany(company);
    const reportType = tabKey === "docx-generation" ? "comprehensive-cdr" : tabKey;

    updateTabState(tabKey, {
      isGenerating: true,
      report: "",
      sources: [],
      status:
        tabKey === "docx-generation"
        ? "Generating the full CDR from all section-specific research and analysis lanes..."
        : `Generating real-time ${tabSpec.label} with its dedicated research and analysis lane. You can switch tabs while this continues...`,
    });

    try {
      const response = await fetch("/api/cdr-report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ company, score: companyScore, internalInfo, reportType }),
      });
      const data = (await response.json()) as {
        report?: string;
        sources?: Array<{ title?: string; url?: string }>;
        error?: string;
      };
      if (!response.ok) throw new Error(data.error || "CDR generation failed.");
      updateTabState(tabKey, {
        report: data.report || "No CDR report returned.",
        sources: data.sources || [],
        status:
          tabKey === "docx-generation"
          ? "Comprehensive CDR generated from all section reports and live feed evidence."
          : `${tabSpec.label} generated from uploaded data and real-time live feed evidence.`,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "";
      const detail = message && !message.includes("CDR generation failed") ? ` Details: ${cleanStatus(message).slice(0, 180)}` : "";
      updateTabState(tabKey, {
        status: message.includes("429")
          ? "Report generation is temporarily rate-limited. Please retry after a short while."
          : `${tabSpec.label} could not be completed. Please retry with the same company.${detail}`,
      });
    } finally {
      updateTabState(tabKey, { isGenerating: false });
    }
  }

  async function downloadDocx() {
    if (!report) {
      updateTabState(activeTab, { status: "Generate the CDR before exporting DOCX." });
      return;
    }

    updateTabState(activeTab, { status: "Preparing DOCX export..." });
    try {
      const response = await fetch("/api/cdr-docx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ company: selectedCompany, score: scoreCompany(selectedCompany), report, sources }),
      });
      if (!response.ok) {
        const data = (await response.json()) as { error?: string };
        throw new Error(data.error || "DOCX export failed.");
      }
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `${selectedCompany.name.replace(/[^a-z0-9]+/gi, "_")}-CDR.docx`;
      anchor.click();
      URL.revokeObjectURL(url);
      updateTabState(activeTab, { status: "DOCX report generated." });
    } catch {
      updateTabState(activeTab, { status: "DOCX export could not be completed. Please retry after generating the CDR." });
    }
  }

  return (
    <AppShell title="Company Detailed Report">
      <div className="grid gap-5 lg:grid-cols-[0.95fr_1.05fr]">
        <GlassPanel>
          <div className="flex items-start justify-between gap-5">
            <div>
              <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Selected Company</p>
              <h2 className="mt-3 font-serif text-4xl font-bold text-slate-950">{selectedCompany.name}</h2>
              <p className="mt-3 text-sm font-medium leading-6 text-slate-800">{selectedCompany.activity}</p>
            </div>
            <ScoreBadge score={scoreCompany(selectedCompany)} />
          </div>

          <div className="mt-6 rounded-2xl border border-white/40 bg-white/45 p-4">
            <label className="grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                Search CDR Company by CIN / Name
              </span>
              <input
                value={cdrSearch}
                onChange={(event) => setCdrSearch(event.target.value)}
                placeholder="Enter CIN or company name"
                className="h-12 rounded-xl border border-white/45 bg-white/65 px-4 font-mono text-sm font-bold text-slate-950 outline-none focus:ring-4 focus:ring-white/50"
              />
            </label>
            <label className="mt-3 grid gap-2">
              <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
                Choose From All Parsed Companies
              </span>
              <select
                value={externalCompany ? "" : selectedCompanyId}
                onChange={(event) => {
                  setSelectedCompanyId(event.target.value);
                  setExternalCompany(null);
                }}
                className="h-12 rounded-xl border border-amber-400/55 bg-white/65 px-4 text-sm font-bold text-slate-950 shadow-inner outline-none focus:ring-4 focus:ring-white/50"
              >
                {rankedCompanies.map((company) => (
                  <option key={company.id} value={company.id}>
                    {company.name} - {company.cin}
                  </option>
                ))}
              </select>
            </label>
            <div className="mt-3 grid gap-2">
              {searchResults.map((company) => (
                <button
                  key={company.id}
                  type="button"
                  onClick={() => {
                    setSelectedCompanyId(company.id);
                    setExternalCompany(null);
                  }}
                  className={`rounded-xl px-4 py-3 text-left text-sm font-bold shadow transition hover:bg-white/75 ${
                    selectedCompany.id === company.id ? "bg-slate-950 text-white" : "bg-white/55 text-slate-950"
                  }`}
                >
                  <span className="block break-words">{company.name}</span>
                  <span className="mt-1 block break-all font-mono text-xs opacity-80">{company.cin}</span>
                </button>
              ))}
              {cdrSearch.trim() && searchResults.length === 0 ? (
                <button
                  type="button"
                  onClick={() => setExternalCompany(makeCinOnlyCompany(cdrSearch))}
                  className="rounded-xl border border-white/45 bg-white/55 px-4 py-3 text-left text-sm font-black text-slate-950 shadow transition hover:bg-white/75"
                >
                  Generate CDR using this CIN only: <span className="font-mono">{cdrSearch.trim().toUpperCase()}</span>
                </button>
              ) : null}
            </div>
          </div>

          <div className="mt-7">
            <FactorGrid company={selectedCompany} />
          </div>
        </GlassPanel>

        <GlassPanel>
          <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">CDR Builder</p>
          <div className="mt-5 grid grid-cols-2 gap-2 md:grid-cols-3">
            {cdrTabs.map((tab) => (
              <button
                key={tab.key}
                type="button"
                onClick={() => setActiveTab(tab.key)}
                className={`min-h-14 rounded-xl px-3 py-2 text-left text-xs font-black leading-5 shadow transition hover:-translate-y-0.5 ${
                  activeTab === tab.key
                    ? "bg-slate-950 text-white"
                    : "border border-white/45 bg-white/55 text-slate-950 hover:bg-white/80"
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>
          <div className="mt-5 rounded-2xl border border-white/40 bg-white/45 p-4">
            <p className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">{activeSpec.label}</p>
            <p className="mt-2 text-sm font-semibold leading-6 text-slate-900">{activeSpec.description}</p>
            <div className="mt-4 grid gap-2">
              {activeSpec.sections.map((section, index) => (
                <div key={section} className="flex items-center gap-3 rounded-xl bg-white/45 px-3 py-2">
                  <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-slate-950 text-xs font-black text-white">
                    {index + 1}
                  </span>
                  <span className="text-sm font-bold leading-5 text-slate-950">{section}</span>
                </div>
              ))}
            </div>
          </div>
          <label className="mt-5 grid gap-2">
            <span className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
              Internal Information Optional
            </span>
            <textarea
              value={internalInfo}
              onChange={(event) => setInternalInfo(event.target.value)}
              rows={5}
              placeholder="Add your private notes: revenue hints, management calls, customer names, margins, capex, concerns, thesis, or diligence questions."
              className="resize-y rounded-xl border border-white/45 bg-white/55 px-4 py-3 text-sm font-semibold leading-6 text-slate-950 outline-none transition placeholder:text-slate-700/60 focus:border-indigo-950/50 focus:ring-4 focus:ring-white/50"
            />
          </label>
          <button
            type="button"
            onClick={generateCdr}
            disabled={isGenerating}
            className="mt-6 w-full rounded-xl bg-slate-950 px-5 py-4 text-sm font-black text-white shadow-xl transition hover:-translate-y-0.5 hover:bg-indigo-950 disabled:cursor-not-allowed disabled:opacity-60"
          >
            {isGenerating
              ? "Generating..."
              : activeTab === "docx-generation"
                ? "Generate Comprehensive CDR"
                : `Generate ${activeSpec.label}`}
          </button>
          {activeTab === "docx-generation" ? (
            <button
              type="button"
              onClick={downloadDocx}
              disabled={!report}
              className="mt-3 w-full rounded-xl border border-white/45 bg-white/55 px-5 py-4 text-sm font-black text-slate-950 shadow-lg transition hover:-translate-y-0.5 hover:bg-white/75 disabled:cursor-not-allowed disabled:opacity-50"
            >
              Generate DOCX File
            </button>
          ) : null}
          {status ? (
            <p className="mt-3 rounded-xl bg-white/45 px-4 py-3 text-sm font-bold leading-6 text-slate-900">
              {status}
            </p>
          ) : null}
        </GlassPanel>
      </div>

      {report ? (
        <GlassPanel className="mt-5">
          <div className="flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
            <div>
              <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Generated Report</p>
              <h2 className="mt-2 font-serif text-3xl font-bold text-slate-950">{activeSpec.label}</h2>
            </div>
            <p className="rounded-xl bg-white/55 px-4 py-2 text-sm font-black text-slate-900">
              {sources.length} live sources
            </p>
          </div>
          <article className="mt-5 max-h-[900px] overflow-y-auto rounded-2xl border border-amber-400/45 bg-white/45 p-6 shadow-inner lg:p-8">
            <ReportViewer report={report} />
          </article>
          <div className="mt-4 rounded-2xl border border-amber-400/45 bg-white/45 p-4">
            <p className="text-xs font-black uppercase tracking-[0.18em] text-slate-800/70">
              Live Feed Sources
            </p>
            <div className="mt-3 flex flex-wrap gap-2">
              {sources.length ? (
                sources.slice(0, 10).map((source, index) => (
                  <a
                    key={`${source.url || source.title}-${index}`}
                    href={source.url}
                    target="_blank"
                    rel="noreferrer"
                    className="max-w-sm rounded-xl bg-white/60 px-3 py-2 text-xs font-bold leading-5 text-indigo-950 transition hover:bg-white/85"
                  >
                    {source.title || source.url || "Source"}
                  </a>
                ))
              ) : (
                <p className="text-sm font-semibold text-slate-800">No live sources returned.</p>
              )}
            </div>
          </div>
        </GlassPanel>
      ) : null}
    </AppShell>
  );
}
