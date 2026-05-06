"use client";

import { Company, factorLabels } from "../_data/companies";
import { useScout } from "./ScoutProvider";

export function GlassPanel({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <section
      className={`rounded-2xl border border-white/45 bg-white/50 p-5 shadow-[0_20px_60px_rgba(15,23,42,0.14)] backdrop-blur-xl ${className}`}
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
  const { factorSelection } = useScout();

  return (
    <div className="grid gap-3 sm:grid-cols-2">
      {Object.entries(company.factors).map(([key, value]) => (
        <div
          key={key}
          className={`rounded-xl border border-white/35 bg-white/45 p-4 transition ${
            factorSelection[key as keyof typeof factorSelection] ? "" : "opacity-45"
          }`}
        >
          <div className="flex items-center justify-between gap-3">
            <p className="text-sm font-bold text-slate-900">
              {factorLabels[key as keyof typeof factorLabels]}
            </p>
            <p className="font-mono text-sm font-black text-indigo-950">{value.toFixed(1)}</p>
          </div>
          <div className="mt-3 h-2 rounded-full bg-white/30">
            <div
              className="h-full rounded-full bg-gradient-to-r from-indigo-950 to-emerald-700"
              style={{ width: `${value * 10}%` }}
            />
          </div>
        </div>
      ))}
    </div>
  );
}

export function CompanyTable({ limit, companies }: { limit?: number; companies?: Company[] }) {
  const { rankedCompanies, scoreCompany } = useScout();
  const sourceCompanies = companies ?? rankedCompanies;
  const visibleCompanies = sourceCompanies.slice(0, limit ?? sourceCompanies.length);

  return (
    <div className="overflow-hidden rounded-2xl border border-white/40 bg-white/45 backdrop-blur-xl">
      <div className="overflow-x-auto">
        <table className="w-full min-w-[900px] table-fixed text-left">
          <thead>
            <tr className="border-b border-white/25 text-xs uppercase tracking-[0.18em] text-slate-800/70">
              <th className="w-[38%] px-5 py-4">Company</th>
              <th className="w-[18%] px-5 py-4">Sector</th>
              <th className="w-[16%] px-5 py-4">City</th>
              <th className="w-[16%] px-5 py-4">Paid-up</th>
              <th className="w-[12%] px-5 py-4 text-right">Score</th>
            </tr>
          </thead>
          <tbody>
            {visibleCompanies.map((company) => (
              <tr key={company.id} className="border-b border-white/15 last:border-0">
                <td className="px-5 py-4">
                  <p className="break-words font-bold leading-5 text-slate-950">{company.name}</p>
                  <p className="mt-1 break-all font-mono text-xs leading-4 text-slate-800/75">{company.cin}</p>
                </td>
                <td className="px-5 py-4 font-semibold leading-5 text-slate-900">{company.sector}</td>
                <td className="px-5 py-4 leading-5 text-slate-900">{company.city}</td>
                <td className="whitespace-nowrap px-5 py-4 font-semibold text-slate-900">{company.paidUpCapital}</td>
                <td className="whitespace-nowrap px-5 py-4 text-right">
                  <span className="inline-flex min-w-12 justify-center rounded-full bg-slate-950 px-3 py-1 text-sm font-bold text-white">
                    {scoreCompany(company)}
                  </span>
                </td>
              </tr>
            ))}
            {visibleCompanies.length === 0 ? (
              <tr>
                <td className="px-5 py-8 text-center font-bold text-slate-800" colSpan={5}>
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
