"use client";

import { AppShell } from "../_components/AppShell";
import { FactorGrid, GlassPanel, ScoreBadge, CompanyTable } from "../_components/CompanyWidgets";
import { ReportViewer } from "../_components/ReportViewer";
import { useScout } from "../_components/ScoutProvider";

type DashboardIconName = "building" | "sector" | "location" | "capital" | "eligible" | "score" | "weights" | "filing";

function DashboardIcon({ name }: { name: DashboardIconName }) {
  const className = "h-5 w-5 text-current drop-shadow-[0_5px_8px_rgba(15,23,42,0.18)]";

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
  } = useScout();
  const eligibleCount = rankedCompanies.filter((company) => company.status === "Active").length;
  const topInsight = aiCompanyInsights[topCompany.id];
  const companyCards = [
    { label: "CIN", value: topCompany.cin, icon: "building", bg: "#dbeafe", border: "#60a5fa", color: "#1d4ed8" },
    { label: "Sector", value: topCompany.sector, icon: "sector", bg: "#dcfce7", border: "#4ade80", color: "#15803d" },
    { label: "Location", value: `${topCompany.city}, ${topCompany.state}`, icon: "location", bg: "#fef3c7", border: "#f59e0b", color: "#b45309" },
    { label: "Paid-up", value: topCompany.paidUpCapital, icon: "capital", bg: "#fae8ff", border: "#d946ef", color: "#a21caf" },
  ] as const;
  const pulseCards = [
    { label: "Eligible Companies", value: eligibleCount.toString(), icon: "eligible", bg: "#ccfbf1", color: "#0f766e" },
    { label: "Prime Score", value: `${scoreCompany(topCompany)}/100`, icon: "score", bg: "#dbeafe", color: "#1d4ed8" },
    { label: "Weighted Factors", value: `${includedFactorCount} / 7`, icon: "weights", bg: "#ede9fe", color: "#6d28d9" },
    { label: "Latest Filing", value: topCompany.lastFiling, icon: "filing", bg: "#ffedd5", color: "#c2410c" },
  ] as const;

  return (
    <AppShell title="Dashboard">
      <div className="grid gap-5 lg:grid-cols-[1.15fr_0.85fr]">
        <GlassPanel className="min-h-[420px] animate-soft-float">
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
            {companyCards.map(({ label, value, icon, bg, border, color }) => (
              <div
                key={label}
                className="min-w-0 overflow-hidden rounded-xl border p-4 shadow-[0_10px_24px_rgba(15,23,42,0.10)] transition duration-300 hover:-translate-y-1"
                style={{ backgroundColor: bg, borderColor: border }}
              >
                <div className="flex items-center gap-2" style={{ color }}>
                  <DashboardIcon name={icon} />
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
              {pulseCards.map(({ label, value, icon, bg, color }) => (
                <div
                  key={label}
                  className="flex items-center justify-between gap-3 rounded-xl px-4 py-3 shadow-sm transition duration-300 hover:-translate-y-0.5"
                  style={{ backgroundColor: bg }}
                >
                  <span className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                    <span style={{ color }}>
                      <DashboardIcon name={icon} />
                    </span>
                    {label}
                  </span>
                  <span className="font-bold" style={{ color }}>{value}</span>
                </div>
              ))}
            </div>
          </GlassPanel>

          <GlassPanel>
            <div className="flex items-start justify-between gap-3">
              <div>
                <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">Investment Layer</p>
                <h3 className="mt-2 font-serif text-3xl font-bold text-slate-950">
                  {topCompany.ipoReadinessBand || topInsight?.recommendation || "Run Scoring"}
                </h3>
              </div>
              {topInsight?.aiScore ? (
                <span className="rounded-xl bg-[#2563eb] px-3 py-2 text-sm font-black text-white shadow-lg shadow-blue-500/20">
                  {topInsight.aiScore}/100
                </span>
              ) : null}
            </div>
            <p className="mt-4 rounded-xl bg-[#dcfce7] px-4 py-3 text-sm font-semibold leading-6 text-slate-900">
              {topCompany.ipoReadinessMessage || topInsight?.rationale || aiScoreReport || scoringStatus || "Upload a file to run parser and screening."}
            </p>
            {topInsight?.redFlags?.length ? (
              <div className="mt-3 grid gap-2">
                {topInsight.redFlags.slice(0, 3).map((flag) => (
                  <p key={flag} className="rounded-xl border border-[#e8a8b2] bg-[#f8dce1] px-4 py-2 text-sm font-bold text-[#6d1525]">
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

      <div className="mt-5">
        <CompanyTable />
      </div>

      <div className="mt-6 flex justify-center">
        <div className="w-full max-w-3xl rounded-2xl border border-[#60a5fa] bg-[#dbeafe] px-6 py-5 text-center shadow-[0_18px_50px_rgba(37,99,235,0.18)] animate-glow-pulse">
          <p className="font-serif text-3xl font-black tracking-normal text-slate-950 sm:text-4xl">
            GreenFlow Ventures Ltd.
          </p>
          <p className="mt-2 text-xs font-black uppercase tracking-[0.24em] text-slate-800/70">
            Smart Scouter Investment Console
          </p>
        </div>
      </div>

      {aiScoreReport ? (
        <GlassPanel className="mt-5">
          <p className="text-sm font-bold uppercase tracking-[0.2em] text-slate-800/70">
            First Screening Layer Report
          </p>
          <div className="mt-5 max-h-[680px] overflow-y-auto rounded-2xl border border-[#60a5fa] bg-white p-4 shadow-inner">
            <ReportViewer report={aiScoreReport} />
          </div>
        </GlassPanel>
      ) : null}
    </AppShell>
  );
}
