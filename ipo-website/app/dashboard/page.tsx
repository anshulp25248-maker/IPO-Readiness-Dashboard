"use client";

import { AppShell } from "../_components/AppShell";
import { FactorGrid, GlassPanel, ScoreBadge, CompanyTable } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";

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
              ["CIN", topCompany.cin],
              ["Sector", topCompany.sector],
              ["Location", `${topCompany.city}, ${topCompany.state}`],
              ["Paid-up", topCompany.paidUpCapital],
            ].map(([label, value]) => (
              <div key={label} className="min-w-0 overflow-hidden rounded-xl border border-white/40 bg-white/45 p-4">
                <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-800/65">{label}</p>
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
                ["Eligible Companies", eligibleCount.toString()],
                ["Prime Score", `${scoreCompany(topCompany)}/100`],
                ["Weighted Factors", `${includedFactorCount} / 7`],
                ["Latest Filing", topCompany.lastFiling],
              ].map(([label, value]) => (
                <div key={label} className="flex items-center justify-between rounded-xl bg-white/45 px-4 py-3">
                  <span className="text-sm font-semibold text-slate-800">{label}</span>
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
            <p className="mt-4 rounded-xl bg-white/45 px-4 py-3 text-sm font-semibold leading-6 text-slate-900">
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
                Parser rejected {parserSummary.rejectedTotal} companies below Rs 5 lakh paid-up capital. {parserSummary.passingToAi} companies passed to AI screening.
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
          <div className="mt-4 grid gap-3 sm:grid-cols-4">
            {[
              ["Uploaded", parserSummary.totalUploaded],
              ["Rejected <5L", parserSummary.rejectedCapital],
              ["Passed to AI", parserSummary.passingToAi],
              ["Ignored Filters", "Geo / NIC"],
            ].map(([label, value]) => (
              <div key={label} className="rounded-xl bg-white/45 px-4 py-3">
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
