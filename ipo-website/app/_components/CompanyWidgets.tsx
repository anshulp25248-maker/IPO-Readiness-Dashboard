"use client";

import { Fragment, useState } from "react";
import { Company, factorKeys, factorLabels } from "../_data/companies";
import { useScout } from "./ScoutProvider";
import { flagMessages } from "../_lib/scout-v2";

export function GlassPanel({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <section
      className={`rounded-2xl border border-[var(--panel-border)] bg-[var(--panel)] p-5 shadow-[0_20px_60px_var(--theme-shadow)] transition duration-300 hover:-translate-y-0.5 hover:bg-[var(--panel-strong)] hover:shadow-[0_24px_70px_var(--theme-shadow)] ${className}`}
    >
      {children}
    </section>
  );
}

export function ScoreBadge({ score }: { score: number }) {
  return (
    <div className="inline-flex min-w-24 self-start items-center justify-center rounded-2xl border border-[var(--panel-border)] bg-[var(--theme-accent)] px-4 py-3 text-center text-white shadow-[0_14px_28px_var(--theme-shadow)] animate-glow-pulse">
      <span className="text-3xl font-black">{score}</span>
      <span className="ml-1 text-sm font-bold text-white/80">/100</span>
    </div>
  );
}

export function FactorGrid({ company }: { company: Company }) {
  const { factorMarks } = useScout();
  const visibleFactorKeys = factorKeys.filter((key) => (factorMarks[key] ?? 0) > 0 && key in company.factors).slice(0, 6);
  const displayKeys = visibleFactorKeys.length ? visibleFactorKeys : factorKeys.slice(0, 6);

  return (
    <div className="grid gap-4 md:grid-cols-3">
      {displayKeys.map((key) => {
        const value = Number(company.factors[key] ?? 0);

        return (
        <div key={key} className="min-h-32 rounded-xl border border-[var(--panel-border)] bg-white p-5 shadow-[0_10px_24px_var(--theme-shadow)] transition duration-300 hover:-translate-y-1 hover:bg-[var(--panel-strong)]">
          <div className="flex items-center justify-between gap-3">
            <p className="text-sm font-black leading-5 text-slate-900">
              {factorLabels[key]}
            </p>
            <p className="font-mono text-sm font-black text-[var(--theme-accent-dark)]">
              {Math.round(value)}/10
            </p>
          </div>
          <div className="mt-3 h-2 overflow-hidden rounded-full bg-[var(--panel)]">
            <div
              className="h-full rounded-full bg-[var(--theme-accent)] animate-fill-sweep"
              style={{ width: `${value * 10}%` }}
            />
          </div>
          {company.factorReasoning?.[key] ? (
            <p className="mt-3 line-clamp-3 text-xs font-semibold leading-5 text-slate-800">
              {company.factorReasoning[key]}
            </p>
          ) : null}
        </div>
      )})}
    </div>
  );
}

function BandBadge({ band }: { band?: Company["ipoReadinessBand"] }) {
  const styles =
    band === "IPO Ready"
      ? "bg-[var(--theme-accent)] text-white"
      : band === "Near Ready"
        ? "bg-[var(--panel-strong)] text-[var(--theme-accent-dark)]"
        : band === "Development Stage"
          ? "bg-[var(--panel)] text-[var(--theme-accent-dark)]"
          : "bg-[#f5c8cf] text-[#6d1525]";
  return <span className={`inline-flex rounded-full px-3 py-1 text-xs font-black ${styles}`}>{band || "Pending"}</span>;
}

export function CompanyTable({ limit, companies }: { limit?: number; companies?: Company[] }) {
  const { rankedCompanies, scoreCompany } = useScout();
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const sourceCompanies = companies ?? rankedCompanies;
  const visibleCompanies = sourceCompanies.slice(0, limit ?? sourceCompanies.length);

  return (
    <div className="overflow-hidden rounded-2xl border border-[var(--panel-border)] bg-white shadow-[0_20px_60px_var(--theme-shadow)]">
      <div className="overflow-x-auto">
        <table className="w-full min-w-[1120px] table-fixed text-left">
          <thead>
            <tr className="border-b border-[var(--panel-border)] bg-[var(--panel)] text-xs uppercase tracking-[0.18em] text-slate-800/70">
              <th className="w-[28%] px-5 py-4">Company</th>
              <th className="w-[14%] px-5 py-4">Sector</th>
              <th className="w-[12%] px-5 py-4">City</th>
              <th className="w-[12%] px-5 py-4">Paid-up</th>
              <th className="w-[10%] px-5 py-4 text-right">Score</th>
              <th className="w-[12%] px-5 py-4">Band</th>
              <th className="w-[6%] px-5 py-4 text-center">RF</th>
              <th className="w-[6%] px-5 py-4 text-center">YF</th>
            </tr>
          </thead>
          <tbody>
            {visibleCompanies.map((company) => {
              const expanded = expandedId === company.id;
              return (
                <Fragment key={company.id}>
                  <tr
                    className="cursor-pointer border-b border-slate-100 last:border-0 transition hover:bg-[var(--panel)]"
                    onClick={() => setExpandedId(expanded ? null : company.id)}
                  >
                    <td className="px-5 py-4">
                      <p className="break-words font-bold leading-5 text-slate-950">{company.name}</p>
                      <p className="mt-1 break-all font-mono text-xs leading-4 text-slate-800/75">{company.cin}</p>
                      <p className="mt-1 text-xs font-bold text-slate-700">{company.status}</p>
                    </td>
                    <td className="px-5 py-4 font-semibold leading-5 text-slate-900">{company.sector}</td>
                    <td className="px-5 py-4 leading-5 text-slate-900">{company.city}</td>
                    <td className="whitespace-nowrap px-5 py-4 font-semibold text-slate-900">{company.paidUpCapital}</td>
                    <td className="whitespace-nowrap px-5 py-4 text-right">
                      <span className="inline-flex min-w-12 justify-center rounded-full bg-[var(--theme-accent)] px-3 py-1 text-sm font-bold text-white">
                        {company.status === "Scoring Failed" ? "NULL" : scoreCompany(company)}
                      </span>
                    </td>
                    <td className="px-5 py-4"><BandBadge band={company.ipoReadinessBand} /></td>
                    <td className="px-5 py-4 text-center">
                      <span className="inline-flex h-7 min-w-7 items-center justify-center rounded-full bg-[#f5c8cf] px-2 text-xs font-black text-[#6d1525]">
                        {company.redFlags?.length ?? 0}
                      </span>
                    </td>
                    <td className="px-5 py-4 text-center">
                      <span className="inline-flex h-7 min-w-7 items-center justify-center rounded-full bg-[var(--panel-strong)] px-2 text-xs font-black text-[var(--theme-accent-dark)]">
                        {company.yellowFlags?.length ?? 0}
                      </span>
                    </td>
                  </tr>
                  {expanded ? (
                    <tr key={`${company.id}-detail`} className="border-b border-slate-100">
                      <td colSpan={8} className="px-5 py-5">
                        <div className="grid gap-4 xl:grid-cols-[1.25fr_0.75fr]">
                          <FactorGrid company={company} />
                          <div className="grid gap-3">
                            <div className="rounded-xl bg-[var(--panel)] p-4">
                              <p className="text-xs font-black uppercase tracking-[0.16em] text-slate-700">IPO Readiness</p>
                              <p className="mt-2 text-lg font-black text-slate-950">{company.ipoReadinessBand || "Pending"}</p>
                              <p className="mt-2 text-sm font-semibold leading-6 text-slate-800">{company.ipoReadinessMessage || company.aiScoringError || "Scoring is pending."}</p>
                            </div>
                            <div className="rounded-xl bg-[var(--panel)] p-4">
                              <p className="text-xs font-black uppercase tracking-[0.16em] text-slate-700">Flags</p>
                              {[...(company.redFlags ?? []), ...(company.yellowFlags ?? [])].length ? (
                                <div className="mt-3 grid gap-2">
                                  {[...(company.redFlags ?? []), ...(company.yellowFlags ?? [])].map((flag) => (
                                    <p
                                      key={flag}
                                      className={`rounded-lg px-3 py-2 text-xs font-bold leading-5 ${
                                        flag.startsWith("RF") ? "bg-[#f8dce1] text-[#6d1525]" : "bg-[var(--panel-strong)] text-[var(--theme-accent-dark)]"
                                      }`}
                                    >
                                      {flag}: {flagMessages[flag] || "Review required."}
                                    </p>
                                  ))}
                                </div>
                              ) : (
                                <p className="mt-3 text-sm font-bold text-slate-800">No flags detected. Clean screening profile.</p>
                              )}
                            </div>
                            <div className="rounded-xl bg-[var(--panel)] p-4">
                              <p className="text-xs font-black uppercase tracking-[0.16em] text-slate-700">Status Verification</p>
                              <p className="mt-2 text-sm font-semibold leading-6 text-slate-800">
                                {company.statusVerification
                                  ? `${company.statusVerification.source}: ${company.statusVerification.statusFound}`
                                  : "Status verification pending."}
                              </p>
                            </div>
                          </div>
                        </div>
                      </td>
                    </tr>
                  ) : null}
                </Fragment>
              );
            })}
            {visibleCompanies.length === 0 ? (
              <tr>
                <td className="px-5 py-8 text-center font-bold text-slate-800" colSpan={8}>
                  No companies match this search.
                </td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
    </div>
  );
}
