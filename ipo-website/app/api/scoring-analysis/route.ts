import { NextResponse } from "next/server";
import { Company, FactorKey, FactorWeights, defaultFactorWeights, factorKeys } from "../../_data/companies";
import {
  assignReadinessBand,
  bandMessages,
  flagMessages,
  normalizeWeights,
  scoreCompanyDeterministically,
} from "../../_lib/scout-v2";
import { envValue, generateAiText } from "../_lib/ai";

export const runtime = "nodejs";

type SearchResult = { title?: string; url?: string; content?: string };
type AiFactor = { score?: number; reasoning?: string; ratio_percentage?: number; cluster_match?: boolean };
type AiScoringJson = {
  composite_score?: number;
  adjusted_score?: number;
  status_verification?: {
    source?: string;
    status_found?: string;
    verified_active?: boolean;
    rf08_applied?: boolean;
    rf09_applied?: boolean;
  };
  factors?: {
    sector_strength?: AiFactor;
    business_model?: AiFactor;
    paid_up_capital?: AiFactor;
    director_profile?: AiFactor;
    filing_compliance?: AiFactor;
    auth_paidup_ratio?: AiFactor;
    geography?: AiFactor;
  };
  red_flags?: string[];
  yellow_flags?: string[];
  ipo_readiness_band?: string;
  ipo_readiness_reasoning?: string;
  override_applied?: boolean;
  override_reason?: string;
};

const factorMap: Record<string, FactorKey> = {
  sector_strength: "sector",
  business_model: "businessModel",
  paid_up_capital: "paidUpCapital",
  director_profile: "directorProfile",
  filing_compliance: "filingCompliance",
  auth_paidup_ratio: "capitalRatio",
  geography: "geography",
};

const tavilyUrl = "https://api.tavily.com/search";

async function tavilySearch(query: string, maxResults = 5) {
  const apiKey = envValue("TAVILY_API_KEY");
  if (!apiKey) return { results: [], status: "Tavily is not configured." };
  const response = await fetch(tavilyUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: apiKey,
      query,
      search_depth: "advanced",
      include_answer: false,
      include_raw_content: false,
      max_results: maxResults,
    }),
    cache: "no-store",
  });
  if (!response.ok) throw new Error(`Tavily search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return { results: data.results ?? [], status: `Tavily returned ${(data.results ?? []).length} snippets.` };
}

async function searchCompany(company: Company) {
  const sectorTerms = `${company.sector} ${company.nicCode} ${company.activity}`.trim();
  const queries = [
    `${company.name} ${company.cin} MCA status active Zauba Tofler`,
    `${company.name} ${company.cin} site:zaubacorp.com company status directors filing`,
    `${company.name} ${company.cin} site:tofler.in company status financials directors`,
    `${company.name} ${company.cin} MCA company master data annual filing`,
    `${sectorTerms} India sector report latest growth forecast PLI SME IPO`,
    `${sectorTerms} India thematic report market size CAGR outlook`,
  ].filter((query) => query.length > 20);

  const batches = await Promise.all(
    queries.map((query) =>
      tavilySearch(query).catch((error) => ({
        results: [],
        status: error instanceof Error ? error.message : "Search failed.",
      })),
    ),
  );
  const seen = new Set<string>();
  const results = batches
    .flatMap((batch) => batch.results)
    .filter((item) => {
      const key = item.url || `${item.title}-${item.content}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 14);
  return {
    results,
    status: results.length ? `Public feed returned ${results.length} snippets.` : batches.map((item) => item.status).join(" | "),
  };
}

function sourceContext(results: SearchResult[]) {
  return (
    results
      .slice(0, 12)
      .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 500)}`)
      .join("\n\n") || "No public feed results returned."
  );
}

function extractJsonObject(content: string) {
  const cleaned = content
    .replace(/```json/gi, "```")
    .replace(/```/g, "")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .trim();
  const start = cleaned.indexOf("{");
  if (start === -1) return cleaned;
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < cleaned.length; index += 1) {
    const char = cleaned[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      escaped = true;
      continue;
    }
    if (char === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (char === "{") depth += 1;
    if (char === "}") depth -= 1;
    if (depth === 0) return cleaned.slice(start, index + 1);
  }
  return cleaned.slice(start);
}

function parseAiJson(content: string) {
  const raw = extractJsonObject(content).replace(/,\s*([}\]])/g, "$1");
  return JSON.parse(raw) as AiScoringJson;
}

function clampFactor(value: unknown, fallback: number) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(10, Math.round(numeric * 10) / 10));
}

function clampScore(value: unknown, fallback: number) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(100, Math.round(numeric)));
}

function normalizeFlags(flags: unknown) {
  if (!Array.isArray(flags)) return [];
  return [...new Set(flags.map((flag) => String(flag).trim().toUpperCase()).filter((flag) => /^(RF|YF)-\d{2}$/.test(flag)))];
}

function normalizeBand(value: unknown, score: number, redFlags: string[], yellowFlags: string[]) {
  const raw = String(value || "").toLowerCase();
  if (raw.includes("ipo ready")) return "IPO Ready" as const;
  if (raw.includes("near")) return "Near Ready" as const;
  if (raw.includes("development")) return "Development Stage" as const;
  if (raw.includes("not")) return "Not Recommended" as const;
  return assignReadinessBand(score, redFlags, yellowFlags);
}

function normalizeAiCompany(company: Company, weights: FactorWeights, ai: AiScoringJson) {
  const fallback = scoreCompanyDeterministically(company, weights);
  const factors = { ...fallback.factors };
  const factorReasoning = { ...(fallback.factorReasoning ?? {}) };

  Object.entries(ai.factors ?? {}).forEach(([key, factor]) => {
    const localKey = factorMap[key];
    if (!localKey) return;
    factors[localKey] = clampFactor(factor?.score, factors[localKey]);
    factorReasoning[localKey] = String(factor?.reasoning || factorReasoning[localKey] || "");
  });

  const aiRedFlags = normalizeFlags(ai.red_flags).filter((flag) => flag.startsWith("RF-"));
  const aiYellowFlags = normalizeFlags(ai.yellow_flags).filter((flag) => flag.startsWith("YF-"));
  const redFlags = [...new Set([...(fallback.redFlags ?? []), ...aiRedFlags])];
  const yellowFlags = [...new Set([...(fallback.yellowFlags ?? []), ...aiYellowFlags])];
  if (ai.status_verification?.rf08_applied) redFlags.push("RF-08");
  if (ai.status_verification?.rf09_applied) redFlags.push("RF-09");
  const uniqueRedFlags = [...new Set(redFlags)];
  const compositeFallback = clampScore(fallback.compositeScore, 0);
  const compositeScore = clampScore(ai.composite_score, compositeFallback);
  const adjustedScore = clampScore(
    ai.adjusted_score,
    Math.max(0, compositeScore - (uniqueRedFlags.includes("RF-08") ? 30 : 0)),
  );
  const band = normalizeBand(ai.ipo_readiness_band, adjustedScore, uniqueRedFlags, yellowFlags);
  const statusVerification = {
    source: ai.status_verification?.source || "Public feed / AI search",
    statusFound: ai.status_verification?.status_found || (ai.status_verification?.rf09_applied ? "Unverified" : "Active or not contradicted"),
    verifiedActive: Boolean(ai.status_verification?.verified_active),
    rf08Applied: Boolean(ai.status_verification?.rf08_applied),
    rf09Applied: Boolean(ai.status_verification?.rf09_applied),
    checkedAt: new Date().toISOString().slice(0, 10),
  };

  return {
    ...company,
    status: statusVerification.rf08Applied ? "Non-Active" : statusVerification.rf09Applied ? "Unverified" : "Active",
    factors,
    factorReasoning,
    compositeScore,
    adjustedScore,
    redFlags: uniqueRedFlags,
    yellowFlags,
    ipoReadinessBand: band,
    ipoReadinessMessage: ai.ipo_readiness_reasoning || bandMessages[band],
    statusVerification,
  } satisfies Company;
}

function insightFromCompany(company: Company) {
  return {
    companyId: company.id,
    companyName: company.name,
    aiScore: Math.round(company.adjustedScore ?? 0),
    recommendation: company.ipoReadinessBand || "Development Stage",
    rationale: company.ipoReadinessMessage || bandMessages[company.ipoReadinessBand || "Development Stage"],
    strengths: factorKeys
      .filter((key) => company.factors[key] >= 8)
      .map((key) => `${key}: ${company.factors[key]}/10.`)
      .slice(0, 4),
    redFlags: (company.redFlags ?? []).map((flag) => `${flag}: ${flagMessages[flag] || "Review required."}`),
    missingData: (company.yellowFlags ?? []).map((flag) => `${flag}: ${flagMessages[flag] || "Review required."}`),
  };
}

function buildPrompt(company: Company, weights: FactorWeights, publicFeed: string, sourceStatus: string) {
  return `Score this company for SME IPO and Pre-IPO equity screening eligibility.

COMPANY DATA:
Name: ${company.name}
CIN: ${company.cin}
NIC Code: ${company.nicCode || "NA"}
Description: ${company.activity || "NA"}
Paid-up Capital: ${company.paidUpCapital}
Authorised Capital: ${company.authorizedCapital}
Incorporation Date: ${company.incorporationDate || "NA"}
Last Filing Date: ${company.lastFiling || "NA"}
City: ${company.city}
State: ${company.state}
Director Name: ${company.director.name}
Director DIN: ${company.director.din || "NA"}
Director Directorships: ${company.director.directorships || "NA"}

WEIGHTS TO APPLY:
Sector Strength: ${weights.sector}%
Business Model: ${weights.businessModel}%
Paid-up Capital: ${weights.paidUpCapital}%
Director Profile: ${weights.directorProfile}%
Filing Compliance: ${weights.filingCompliance}%
Auth/Paid-up Ratio: ${weights.capitalRatio}%
Geography: ${weights.geography}%

PUBLIC FEED STATUS:
${sourceStatus}

PUBLIC FEED:
${publicFeed}

TASKS:
1. Use the public feed as the search evidence for Zauba Corp, Tofler, MCA portal, and sector/thematic reports.
2. Verify whether the company appears Active, Non-Active, or Unverified. If no source verifies status, apply RF-09 only.
3. Apply sector cluster geography validation using NIC code.
4. Score all seven factors using the Scout Smarter V2 investment rules.
5. Detect all red flags and yellow flags.
6. Assign IPO readiness band using adjusted score after any RF-08 penalty.

Return this exact JSON structure with double quotes only:
{
  "composite_score": 0,
  "adjusted_score": 0,
  "status_verification": {
    "source": "",
    "status_found": "",
    "verified_active": true,
    "rf08_applied": false,
    "rf09_applied": false
  },
  "factors": {
    "sector_strength": { "score": 0, "reasoning": "" },
    "business_model": { "score": 0, "reasoning": "" },
    "paid_up_capital": { "score": 0, "reasoning": "" },
    "director_profile": { "score": 0, "reasoning": "" },
    "filing_compliance": { "score": 0, "reasoning": "" },
    "auth_paidup_ratio": { "score": 0, "ratio_percentage": 0, "reasoning": "" },
    "geography": { "score": 0, "cluster_match": false, "reasoning": "" }
  },
  "red_flags": [],
  "yellow_flags": [],
  "ipo_readiness_band": "IPO Ready",
  "ipo_readiness_reasoning": "",
  "override_applied": false,
  "override_reason": ""
}`;
}

async function aiScore(company: Company, weights: FactorWeights, publicFeed: string, sourceStatus: string) {
  const result = await generateAiText({
    task: "scoring",
    system:
      "You are an expert investment screening analyst for a SEBI-registered Category I AIF focused on SME IPO and Pre-IPO equity in India. You have deep knowledge of MCA company data, SEBI listing requirements, BSE SME and NSE Emerge eligibility criteria, Indian sector dynamics, NIC code mapping, and investment due diligence standards. Score objectively. Flag ruthlessly. Never miss a red flag to make a score look better. Your output will be used by an investment committee. Return ONLY valid JSON. No preamble. No explanation outside the JSON structure. No markdown backticks.",
    prompt: buildPrompt(company, weights, publicFeed, sourceStatus),
    temperature: 0.05,
    maxTokens: 2200,
    responseJson: true,
  });
  return parseAiJson(result.text);
}

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as {
      company?: Company;
      companies?: Company[];
      weights?: Partial<FactorWeights>;
    };
    const weights = normalizeWeights(body.weights ?? defaultFactorWeights);
    const company = body.company ?? body.companies?.[0];
    if (!company) return NextResponse.json({ error: "No company supplied." }, { status: 400 });

    let results: SearchResult[] = [];
    let sourceStatus = "";
    try {
      const search = await searchCompany(company);
      results = search.results;
      sourceStatus = search.status;
    } catch (error) {
      sourceStatus = error instanceof Error ? error.message : "Public-source search failed.";
    }

    let scored: Company;
    let fallback = false;
    try {
      const parsed = await aiScore(company, weights, sourceContext(results), sourceStatus);
      scored = normalizeAiCompany(company, weights, parsed);
    } catch (error) {
      fallback = true;
      scored = scoreCompanyDeterministically(company, weights);
      scored = {
        ...scored,
        statusVerification: {
          source: "Fallback deterministic scoring",
          statusFound: sourceStatus || "AI unavailable",
          verifiedActive: false,
          rf08Applied: false,
          rf09Applied: !results.length,
          checkedAt: new Date().toISOString().slice(0, 10),
        },
        redFlags: [...new Set([...(scored.redFlags ?? []), ...(!results.length ? ["RF-09"] : [])])],
        aiScoringError: error instanceof Error ? error.message : "AI provider returned invalid scoring JSON.",
      };
      scored = {
        ...scored,
        status: !results.length ? "Unverified" : scored.status,
        ipoReadinessBand: assignReadinessBand(scored.adjustedScore ?? 0, scored.redFlags ?? [], scored.yellowFlags ?? []),
      };
      scored.ipoReadinessMessage = bandMessages[scored.ipoReadinessBand || "Development Stage"];
    }

    return NextResponse.json({
      company: scored,
      insight: insightFromCompany(scored),
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      sourceStatus,
      fallback,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "AI scoring layer failed." }, { status: 500 });
  }
}
