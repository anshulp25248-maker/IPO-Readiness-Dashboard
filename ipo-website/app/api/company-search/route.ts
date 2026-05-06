import { NextResponse } from "next/server";
import { Company, FactorKey } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";

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

const tavilyUrl = "https://api.tavily.com/search";
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
  const name = String(input.name || input.companyName || `AI Company ${index + 1}`);
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
      credibility: String((input.director as Record<string, unknown> | undefined)?.credibility || "AI research pending verification"),
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

async function tavilySearch(query: string) {
  const apiKey = envValue("TAVILY_API_KEY");
  if (!apiKey) return [];

  const response = await fetch(tavilyUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: apiKey,
      query,
      search_depth: "advanced",
      include_answer: false,
      include_raw_content: false,
      max_results: 8,
    }),
    cache: "no-store",
  });

  if (!response.ok) {
    throw new Error(`Tavily search failed with ${response.status}`);
  }

  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

function sourceContext(results: SearchResult[]) {
  if (!results.length) return "No live feed results returned.";
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

async function groqAnalyze(body: SearchBody, results: SearchResult[]) {
  const prompt = `You are Scout Smarter, an unlisted-company origination analyst.

SEARCH REQUEST
- Company/name/CIN query: ${body.query || body.companyName || body.cin || "NA"}
- CIN: ${body.cin || "NA"}
- City: ${body.city || "NA"}
- State: ${body.state || "NA"}

LIVE FEED RESULTS
${sourceContext(results)}

Task:
1. Identify real companies relevant to the query/city/state/CIN/name from the feed.
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
}`;

  const result = await generateAiText({
    task: "company-search",
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
    const query = buildSearchQuery(body);
    const results = await tavilySearch(query);
    const analysis = await groqAnalyze(body, results);

    return NextResponse.json({
      ...analysis,
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "AI search failed" },
      { status: 500 },
    );
  }
}
