"use client";

import { Fragment, useState } from "react";
import { Company, factorLabels } from "../_data/companies";
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
      className={`rounded-2xl border border-white/50 bg-white/60 p-5 shadow-[0_20px_60px_rgba(15,23,42,0.12)] backdrop-blur-xl ${className}`}
    >
      {children}
    </section>
  );
}

export function ScoreBadge({ score }: { score: number }) {
  return (
    <div className="inline-flex min-w-24 items-center justify-center rounded-2xl border border-white/50 bg-white/55 px-4 py-3 text-center shadow-lg backdrop-blur-xl">
      <span className="text-3xl font-black text-slate-950">{score}</span>
      <span className="ml-1 text-sm font-bold text-slate-800">/100</span>
    </div>
  );
}

export function FactorGrid({ company }: { company: Company }) {
  const { factorWeights } = useScout();

  return (
    <div className="grid gap-3 sm:grid-cols-2">
      {Object.entries(company.factors).map(([key, value]) => (
        <div key={key} className="rounded-xl border border-white/40 bg-white/55 p-4 transition">
          <div className="flex items-center justify-between gap-3">
            <p className="text-sm font-bold text-slate-900">
              {factorLabels[key as keyof typeof factorLabels]}
            </p>
            <p className="font-mono text-sm font-black text-indigo-950">
              {value.toFixed(1)} x {factorWeights[key as keyof typeof factorWeights]}%
            </p>
          </div>
          <div className="mt-3 h-2 rounded-full bg-white/30">
            <div
              className="h-full rounded-full bg-gradient-to-r from-indigo-950 to-emerald-700"
              style={{ width: `${value * 10}%` }}
            />
          </div>
          {company.factorReasoning?.[key as keyof typeof company.factorReasoning] ? (
            <p className="mt-3 text-xs font-semibold leading-5 text-slate-800">
              {company.factorReasoning[key as keyof typeof company.factorReasoning]}
            </p>
          ) : null}
        </div>
      ))}
    </div>
  );
}

function BandBadge({ band }: { band?: Company["ipoReadinessBand"] }) {
  const styles =
    band === "IPO Ready"
      ? "bg-emerald-900 text-white"
      : band === "Near Ready"
        ? "bg-emerald-100 text-emerald-950"
        : band === "Development Stage"
          ? "bg-amber-100 text-amber-950"
          : "bg-rose-100 text-rose-950";
  return <span className={`inline-flex rounded-full px-3 py-1 text-xs font-black ${styles}`}>{band || "Pending"}</span>;
}

export function CompanyTable({ limit, companies }: { limit?: number; companies?: Company[] }) {
  const { rankedCompanies, scoreCompany } = useScout();
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const sourceCompanies = companies ?? rankedCompanies;
  const visibleCompanies = sourceCompanies.slice(0, limit ?? sourceCompanies.length);

  return (
    <div className="overflow-hidden rounded-2xl border border-white/45 bg-white/55 backdrop-blur-xl">
      <div className="overflow-x-auto">
        <table className="w-full min-w-[1120px] table-fixed text-left">
          <thead>
            <tr className="border-b border-white/25 text-xs uppercase tracking-[0.18em] text-slate-800/70">
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
                    className="cursor-pointer border-b border-white/15 last:border-0 transition hover:bg-white/25"
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
                      <span className="inline-flex min-w-12 justify-center rounded-full bg-slate-950 px-3 py-1 text-sm font-bold text-white">
                        {company.status === "Scoring Failed" ? "NULL" : scoreCompany(company)}
                      </span>
                    </td>
                    <td className="px-5 py-4"><BandBadge band={company.ipoReadinessBand} /></td>
                    <td className="px-5 py-4 text-center">
                      <span className="inline-flex h-7 min-w-7 items-center justify-center rounded-full bg-rose-100 px-2 text-xs font-black text-rose-950">
                        {company.redFlags?.length ?? 0}
                      </span>
                    </td>
                    <td className="px-5 py-4 text-center">
                      <span className="inline-flex h-7 min-w-7 items-center justify-center rounded-full bg-amber-100 px-2 text-xs font-black text-amber-950">
                        {company.yellowFlags?.length ?? 0}
                      </span>
                    </td>
                  </tr>
                  {expanded ? (
                    <tr key={`${company.id}-detail`} className="border-b border-white/15">
                      <td colSpan={8} className="px-5 py-5">
                        <div className="grid gap-4 lg:grid-cols-[1fr_0.85fr]">
                          <FactorGrid company={company} />
                          <div className="grid gap-3">
                            <div className="rounded-xl bg-white/55 p-4">
                              <p className="text-xs font-black uppercase tracking-[0.16em] text-slate-700">IPO Readiness</p>
                              <p className="mt-2 text-lg font-black text-slate-950">{company.ipoReadinessBand || "Pending"}</p>
                              <p className="mt-2 text-sm font-semibold leading-6 text-slate-800">{company.ipoReadinessMessage || company.aiScoringError || "AI scoring is pending."}</p>
                            </div>
                            <div className="rounded-xl bg-white/55 p-4">
                              <p className="text-xs font-black uppercase tracking-[0.16em] text-slate-700">Flags</p>
                              {[...(company.redFlags ?? []), ...(company.yellowFlags ?? [])].length ? (
                                <div className="mt-3 grid gap-2">
                                  {[...(company.redFlags ?? []), ...(company.yellowFlags ?? [])].map((flag) => (
                                    <p
                                      key={flag}
                                      className={`rounded-lg px-3 py-2 text-xs font-bold leading-5 ${
                                        flag.startsWith("RF") ? "bg-rose-50 text-rose-950" : "bg-amber-50 text-amber-950"
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
                            <div className="rounded-xl bg-white/55 p-4">
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
