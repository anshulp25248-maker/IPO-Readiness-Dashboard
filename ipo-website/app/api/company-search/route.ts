import { NextResponse } from "next/server";
import { Company, FactorKey } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";
import { scoringJsonReportFormat } from "../_lib/report-format";

export const runtime = "nodejs";

type SearchBody = {
  query?: string;
  city?: string;
  state?: string;
  cin?: string;
  companyName?: string;
};

type SearchResult = {
  title?: string;
  url?: string;
  content?: string;
};

const factorKeys: FactorKey[] = [
  "paidUpCapital",
  "sector",
  "geography",
  "businessModel",
  "directorProfile",
  "capitalRatio",
  "filingCompliance",
];

function clampScore(value: unknown, fallback = 5) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(10, numeric));
}

function slug(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "company";
}

function normalizeStatus(value: unknown): Company["status"] {
  const normalized = String(value || "Active").toLowerCase();
  if (/\b(inactive|strike|struck|closed|liquidation|liquidated|dissolved|dormant|amalgamated|rejected)\b/.test(normalized)) {
    return "Rejected";
  }
  return "Active";
}

function normalizeCompany(input: Record<string, unknown>, index: number): Company {
  const name = String(input.name || input.companyName || `Researched Company ${index + 1}`);
  const factors = factorKeys.reduce<Record<FactorKey, number>>((scores, key) => {
    const rawFactors = input.factors as Record<string, unknown> | undefined;
    scores[key] = clampScore(rawFactors?.[key], 5);
    return scores;
  }, {} as Record<FactorKey, number>);

  return {
    id: `ai-${slug(name)}-${index}`,
    name,
    cin: String(input.cin || "NA"),
    sector: String(input.sector || "Other"),
    city: String(input.city || "NA"),
    state: String(input.state || "NA"),
    status: normalizeStatus(input.status),
    paidUpCapital: String(input.paidUpCapital || "NA"),
    authorizedCapital: String(input.authorizedCapital || "NA"),
    nicCode: String(input.nicCode || "NA"),
    activity: String(input.activity || "NA"),
    lastFiling: String(input.lastFiling || "NA"),
    director: {
      name: String((input.director as Record<string, unknown> | undefined)?.name || "NA"),
      role: String((input.director as Record<string, unknown> | undefined)?.role || "Director"),
      education: String((input.director as Record<string, unknown> | undefined)?.education || "NA"),
      directorships: Number((input.director as Record<string, unknown> | undefined)?.directorships || 0),
      credibility: String((input.director as Record<string, unknown> | undefined)?.credibility || "Research pending verification"),
    },
    factors,
    competitors: Array.isArray(input.competitors)
      ? input.competitors.map(String).filter(Boolean).slice(0, 3)
      : [],
  };
}

function buildSearchQuery(body: SearchBody) {
  const parts = [
    body.companyName,
    body.cin,
    body.query,
    body.city,
    body.state,
    "India private limited company MCA CIN directors paid up capital activity sector",
  ].filter(Boolean);
  return parts.join(" ");
}

function sourceContext(results: SearchResult[]) {
  if (!results.length) return "Dashboard public search is disabled. Use uploaded/parser data and mark unavailable fields as NA.";
  return results
    .map((item, index) => {
      return [
        `${index + 1}. ${item.title || "Untitled"}`,
        `URL: ${item.url || "NA"}`,
        `Snippet: ${(item.content || "").slice(0, 650)}`,
      ].join("\n");
    })
    .join("\n\n");
}

function cleanProviderDetails(message: string) {
  return message
    .replace(/\bGroq\b/gi, "analysis service")
    .replace(/\bTavily\b/gi, "live feed")
    .replace(/\bGemini\b/gi, "analysis service")
    .replace(/\bOpenRouter\b/gi, "analysis service")
    .replace(/No AI provider configured\.[^|]*/gi, "Analysis service is not configured.");
}

async function groqAnalyze(body: SearchBody, results: SearchResult[]) {
  const prompt = `You are Scout Smarter, an unlisted-company origination analyst.

SEARCH REQUEST
- Company/name/CIN query: ${body.query || body.companyName || body.cin || "NA"}
- CIN: ${body.cin || "NA"}
- City: ${body.city || "NA"}
- State: ${body.state || "NA"}

PUBLIC SEARCH CONTEXT
${sourceContext(results)}

Task:
1. Identify companies relevant to the query/city/state/CIN/name only when the supplied fields or existing context support them.
2. Prefer active Indian private/public unlisted companies when available.
3. Generate a concise analyst report with source-aware caveats.
4. Score every company on these equal factors from 0-10 using this exact logic:
- Reject inactive companies and companies with confirmed paid-up capital below Rs 10 lakh.
- Paid-up Capital: log-normalized against the maximum eligible paid-up capital in the returned set.
- Sector: use NIC sector tier first, then live sentiment from feed evidence only. High-conviction sectors include semiconductors, defence, pharma, medical devices, renewable energy, EV/batteries, AI/software automation, and robotics.
- Geography: Tier 1 Indian metro/business hubs = 10, Tier 2 = 7, Tier 3/other = 4.
- Business Model: if one company has the same NIC/activity, score 10; 2-10 score 10/N; 11-100 score 100/N; above 100 score 1.
- Director Profile: directorships 3+ = 10, 2 = 7, 1 = 4. Professional degree plus multiple directorships = 10. Professional degree only = 5. If fewer than two data points are found, mark credibility "Unverified" and use "NA" in the report.
- Paid-up / Authorized Ratio: <20% = 1, 20-50% = 6, 50-80% = 8, >80% = 9, exactly 100% = 10.
- Filing Compliance: within 6 months = 10, 6-12 months = 7, 12-24 months = 4, confirmed >24 months = reject.
5. Do not hallucinate. Use "NA" where data is not found and state which fields were unavailable.

Return strict JSON only:
{
  "report": "professional markdown report",
  "companies": [
    {
      "name": "...",
      "cin": "...",
      "sector": "...",
      "city": "...",
      "state": "...",
      "status": "Active",
      "paidUpCapital": "...",
      "authorizedCapital": "...",
      "nicCode": "...",
      "activity": "...",
      "lastFiling": "...",
      "director": {"name":"...", "role":"...", "education":"...", "directorships": 0, "credibility":"..."},
      "factors": {
        "paidUpCapital": 0,
        "sector": 0,
        "geography": 0,
        "businessModel": 0,
        "directorProfile": 0,
        "capitalRatio": 0,
        "filingCompliance": 0
      },
      "competitors": ["...", "..."]
    }
  ]
}

${scoringJsonReportFormat}`;

  const result = await generateAiText({
    task: "company-search",
    provider: "groq",
    apiKey:
      envValue("CDR_OTHER_COMPANIES_GROQ_API_KEY") ||
      envValue("GROQ_API_KEY_CDR_OTHER_COMPANIES") ||
      envValue("CDR_COMPREHENSIVE_CDR_GROQ_API_KEY") ||
      envValue("GROQ_API_KEY_CDR_COMPREHENSIVE_CDR") ||
      envValue("GROQ_API_KEY"),
    providerLabel: "Other Companies",
    system: "Return only valid JSON. You are careful, source-aware, and mark missing information as NA.",
    prompt,
    temperature: 0.12,
    maxTokens: 2500,
    responseJson: true,
  });

  const content = result.text || "{}";
  const match = content.match(/\{[\s\S]*\}/);
  const parsed = JSON.parse(match ? match[0] : content) as {
    report?: string;
    companies?: Array<Record<string, unknown>>;
  };

  return {
    report: parsed.report || "No report returned.",
    companies: (parsed.companies || []).map(normalizeCompany),
  };
}

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as SearchBody;
    buildSearchQuery(body);
    const results: SearchResult[] = [];
    const analysis = await groqAnalyze(body, results);

    return NextResponse.json({
      ...analysis,
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? cleanProviderDetails(error.message) : "Company research failed" },
      { status: 500 },
    );
  }
}
