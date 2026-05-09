"use client";

import { AppShell } from "../_components/AppShell";
import { FactorGrid, GlassPanel, ScoreBadge, CompanyTable } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";

type DashboardIconName = "building" | "sector" | "location" | "capital" | "eligible" | "score" | "weights" | "filing";

function DashboardIcon({ name }: { name: DashboardIconName }) {
  const className = "h-5 w-5 text-slate-950";

  if (name === "building") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M6 21V4h12v17" />
        <path d="M9 8h2" />
        <path d="M13 8h2" />
        <path d="M9 12h2" />
        <path d="M13 12h2" />
        <path d="M10 21v-5h4v5" />
      </svg>
    );
  }

  if (name === "sector") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 19V9l5 4V9l5 4V7h6v12H4Z" />
      </svg>
    );
  }

  if (name === "location") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 21s7-5.4 7-12a7 7 0 1 0-14 0c0 6.6 7 12 7 12Z" />
        <path d="M12 11a2 2 0 1 0 0-4 2 2 0 0 0 0 4Z" />
      </svg>
    );
  }

  if (name === "capital") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 8h16" />
        <path d="M6 8V6h12v2" />
        <path d="M7 8v11" />
        <path d="M17 8v11" />
        <path d="M4 19h16" />
      </svg>
    );
  }

  if (name === "eligible") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M20 6 9 17l-5-5" />
      </svg>
    );
  }

  if (name === "score") {
    return (
      <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 14a8 8 0 1 1 16 0" />
        <path d="m12 14 4-4" />
        <path d="M8 20h8" />
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
      <path d="M7 3h7l4 4v14H7V3Z" />
      <path d="M14 3v5h5" />
      <path d="M10 13h6" />
    </svg>
  );
}

export default function DashboardPage() {
  const {
    rankedCompanies,
    topCompany,
    scoreCompany,
    includedFactorCount,
    aiScoreReport,
    aiCompanyInsights,
    scoringStatus,
    parserSummary,
    aiProgress,
  } = useScout();
  const eligibleCount = rankedCompanies.filter((company) => company.status === "Active").length;
  const topInsight = aiCompanyInsights[topCompany.id];

  return (
    <AppShell title="Dashboard">
      <div className="grid gap-5 lg:grid-cols-[1.15fr_0.85fr]">
        <GlassPanel className="min-h-[420px]">
          <div className="flex flex-col gap-5 sm:flex-row sm:items-start sm:justify-between">
            <div>
              <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">
                Highest Scored Company
              </p>
              <h2 className="mt-3 max-w-2xl font-serif text-4xl font-bold leading-tight text-slate-950">
                {topCompany.name}
              </h2>
              <p className="mt-3 max-w-2xl text-sm font-medium leading-6 text-slate-800">
                {topCompany.activity}
              </p>
            </div>
            <ScoreBadge score={scoreCompany(topCompany)} />
          </div>

          <div className="mt-7 grid gap-3 sm:grid-cols-4">
            {[
              ["CIN", topCompany.cin, "building"],
              ["Sector", topCompany.sector, "sector"],
              ["Location", `${topCompany.city}, ${topCompany.state}`, "location"],
              ["Paid-up", topCompany.paidUpCapital, "capital"],
            ].map(([label, value, icon]) => (
              <div key={label} className="min-w-0 overflow-hidden rounded-xl border border-white/45 bg-white/55 p-4">
                <div className="flex items-center gap-2">
                  <DashboardIcon name={icon as DashboardIconName} />
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-800/65">{label}</p>
                </div>
                <p className="mt-2 break-words text-sm font-bold leading-6 text-slate-950">{value}</p>
              </div>
            ))}
          </div>

          <div className="mt-7">
            <FactorGrid company={topCompany} />
          </div>
        </GlassPanel>

        <div className="grid gap-5">
          <GlassPanel>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Portfolio Pulse</p>
            <div className="mt-5 grid gap-3">
              {[
                ["Eligible Companies", eligibleCount.toString(), "eligible"],
                ["Prime Score", `${scoreCompany(topCompany)}/100`, "score"],
                ["Weighted Factors", `${includedFactorCount} / 7`, "weights"],
                ["Latest Filing", topCompany.lastFiling, "filing"],
              ].map(([label, value, icon]) => (
                <div key={label} className="flex items-center justify-between gap-3 rounded-xl bg-white/55 px-4 py-3">
                  <span className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                    <DashboardIcon name={icon as DashboardIconName} />
                    {label}
                  </span>
                  <span className="font-bold text-slate-950">{value}</span>
                </div>
              ))}
            </div>
          </GlassPanel>

          <GlassPanel>
            <div className="flex items-start justify-between gap-3">
              <div>
                <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">AI Investment Layer</p>
                <h3 className="mt-2 font-serif text-3xl font-bold text-slate-950">
                  {topCompany.ipoReadinessBand || topInsight?.recommendation || "Run Scoring"}
                </h3>
              </div>
              {topInsight?.aiScore ? (
                <span className="rounded-xl bg-slate-950 px-3 py-2 text-sm font-black text-white">
                  AI {topInsight.aiScore}/100
                </span>
              ) : null}
            </div>
            <p className="mt-4 rounded-xl bg-white/55 px-4 py-3 text-sm font-semibold leading-6 text-slate-900">
              {topCompany.ipoReadinessMessage || topInsight?.rationale || aiScoreReport || scoringStatus || "Upload a file to run parser and AI screening."}
            </p>
            {topInsight?.redFlags?.length ? (
              <div className="mt-3 grid gap-2">
                {topInsight.redFlags.slice(0, 3).map((flag) => (
                  <p key={flag} className="rounded-xl border border-rose-200 bg-rose-50/80 px-4 py-2 text-sm font-bold text-rose-950">
                    {flag}
                  </p>
                ))}
              </div>
            ) : null}
          </GlassPanel>

          <GlassPanel>
            <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Capital Structure</p>
            <div className="mt-5 space-y-4">
              <div>
                <p className="text-sm text-slate-800">Authorized Capital</p>
                <p className="text-2xl font-black text-slate-950">{topCompany.authorizedCapital}</p>
              </div>
              <div>
                <p className="text-sm text-slate-800">NIC Code</p>
                <p className="font-mono text-2xl font-black text-slate-950">{topCompany.nicCode}</p>
              </div>
            </div>
          </GlassPanel>
        </div>
      </div>

      {parserSummary ? (
        <GlassPanel className="mt-5">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Parser Rejection Summary</p>
              <p className="mt-2 text-sm font-semibold text-slate-900">
                Parser rejected {parserSummary.rejectedTotal} companies before Groq scoring: {parserSummary.rejectedCapital} below Rs 5 lakh paid-up capital, {parserSummary.rejectedCommunityService} community-service profiles, and {parserSummary.rejectedGovernment} government/public-sector profiles. {parserSummary.passingToAi} companies passed to AI screening.
              </p>
            </div>
            {aiProgress.total ? (
              <div className="min-w-[260px]">
                <div className="h-3 overflow-hidden rounded-full bg-white/40">
                  <div
                    className="h-full rounded-full bg-slate-950"
                    style={{ width: `${Math.round((aiProgress.completed / aiProgress.total) * 100)}%` }}
                  />
                </div>
                <p className="mt-2 text-xs font-black text-slate-950">
                  AI screening {aiProgress.completed} of {aiProgress.total} companies
                </p>
              </div>
            ) : null}
          </div>
          <div className="mt-4 grid gap-3 sm:grid-cols-5">
            {[
              ["Uploaded", parserSummary.totalUploaded],
              ["Rejected <5L", parserSummary.rejectedCapital],
              ["Community", parserSummary.rejectedCommunityService],
              ["Government", parserSummary.rejectedGovernment],
              ["Passed to AI", parserSummary.passingToAi],
            ].map(([label, value]) => (
              <div key={label} className="rounded-xl bg-white/55 px-4 py-3">
                <p className="text-xs font-bold uppercase tracking-[0.16em] text-slate-700">{label}</p>
                <p className="mt-1 text-2xl font-black text-slate-950">{value}</p>
              </div>
            ))}
          </div>
        </GlassPanel>
      ) : null}

      <div className="mt-5">
        <CompanyTable limit={5} />
      </div>

      {aiScoreReport ? (
        <GlassPanel className="mt-5">
          <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">
            First AI Layer Report
          </p>
          <div className="mt-5 max-h-[680px] overflow-y-auto rounded-2xl border border-white/35 bg-white/35 p-4 shadow-inner">
            <ReportViewer report={aiScoreReport} />
          </div>
        </GlassPanel>
      ) : null}
    </AppShell>
  );
}
