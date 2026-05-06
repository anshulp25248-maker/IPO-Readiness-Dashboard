import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";

export const runtime = "nodejs";

type SearchResult = { title?: string; url?: string; content?: string };

const tavilyUrl = "https://api.tavily.com/search";

async function tavilySearch(query: string) {
  const apiKey = envValue("TAVILY_API_KEY");
  if (!apiKey) return [];
  const response = await fetch(tavilyUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ api_key: apiKey, query, search_depth: "advanced", max_results: 10 }),
    cache: "no-store",
  });
  if (!response.ok) throw new Error(`Tavily scoring search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

function sourceContext(results: SearchResult[]) {
  return results
    .slice(0, 4)
    .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 300)}`)
    .join("\n\n") || "No public feed results returned.";
}

function parseJson(content: string) {
  const match = content.match(/\{[\s\S]*\}/);
  return JSON.parse(match ? match[0] : content) as {
    report?: string;
    insights?: Array<{
      companyId?: string;
      companyName?: string;
      aiScore?: number;
      recommendation?: string;
      rationale?: string;
      strengths?: string[];
      redFlags?: string[];
      missingData?: string[];
    }>;
  };
}

function fallbackInsights(companies: Company[], scores?: Record<string, number>, reason = "AI provider unavailable") {
  return companies.map((company) => {
    const score = scores?.[company.id] ?? 0;
    const redFlags = [
      company.lastFiling === "NA" || company.lastFiling === "Data Not Available" ? "Filing date is not available in uploaded data." : "",
      company.director.credibility?.toLowerCase().includes("limited") ? "Director credibility requires public-source verification." : "",
      company.paidUpCapital === "Data Not Available" ? "Paid-up capital is missing from uploaded data." : "",
    ].filter(Boolean);

    return {
      companyId: company.id,
      companyName: company.name,
      aiScore: Math.max(0, Math.min(100, Math.round(score))),
      recommendation: score >= 85 ? "Watchlist" : score >= 70 ? "Data Insufficient" : "Reject",
      rationale: `${reason}. Showing a conservative fallback view from uploaded factor scores only; run again after the Groq quota resets for live public-source AI analysis.`,
      strengths: [
        `Deterministic Scout Score is ${Math.round(score)}/100.`,
        `Sector factor is ${company.factors.sector.toFixed(1)}/10 and business-model factor is ${company.factors.businessModel.toFixed(1)}/10.`,
      ],
      redFlags,
      missingData: ["Live Groq qualitative analysis was not returned."],
    };
  });
}

async function aiJson(prompt: string) {
  const result = await generateAiText({
    system: "Return only valid compact JSON. Never hallucinate.",
    prompt,
    temperature: 0.05,
    maxTokens: 1800,
    responseJson: true,
  });
  return result.text || "{}";
}

function scoringTruth(company: Company, score: number | undefined) {
  return {
    id: company.id,
    name: company.name,
    cin: company.cin,
    uploadedPaidUpCapital: company.paidUpCapital,
    uploadedAuthorizedCapital: company.authorizedCapital,
    uploadedPaidUpCapitalRaw: company.paidUpCapitalValue ?? null,
    uploadedAuthorizedCapitalRaw: company.authorizedCapitalValue ?? null,
    sector: company.sector,
    nicCode: company.nicCode,
    activity: company.activity,
    director: company.director,
    deterministicScore: score ?? null,
    deterministicFactorScores: company.factors,
  };
}

export async function POST(request: Request) {
  try {
    const { companies, scores } = (await request.json()) as {
      companies?: Company[];
      scores?: Record<string, number>;
    };
    const selectedCompanies = (companies || []).slice(0, 3);
    if (!selectedCompanies.length) return NextResponse.json({ error: "No companies supplied." }, { status: 400 });

    const query = `${selectedCompanies.slice(0, 5).map((company) => `${company.name} ${company.cin}`).join(" OR ")} MCA Zauba Tofler sector directors India`;
    const results = await tavilySearch(query);

    const prompt = `You are the AI scoring layer for Scout Smarter. Analyze uploaded Indian company screening data plus public feed snippets. Keep every statement source-aware and conservative.

Do not invent facts. Uploaded paid-up capital, authorized capital, CIN, sector/NIC/activity, factor scores, and deterministic total score are the source of truth.

COMPANY DATA
${JSON.stringify(selectedCompanies.map((company) => scoringTruth(company, scores?.[company.id])), null, 2)}

PUBLIC FEED
${sourceContext(results)}

Return strict JSON only:
{
  "report": "short markdown portfolio-level analyst note",
  "insights": [
    {
      "companyId": "must match input id",
      "companyName": "...",
      "aiScore": 0,
      "recommendation": "Invest / Watchlist / Reject / Data Insufficient",
      "rationale": "2-4 sentence source-aware explanation",
      "strengths": ["verified or clearly inferred strength"],
      "redFlags": ["bold-worthy risks or missing public data"],
      "missingData": ["fields not available publicly"]
    }
  ]
}

AI score should reflect public-data confidence, sector attractiveness, business-model uniqueness, management/compliance evidence, red flags, and missing-data penalties.`;

    let parsed: ReturnType<typeof parseJson>;
    try {
      parsed = parseJson(await aiJson(prompt));
    } catch (error) {
      const message = error instanceof Error ? error.message : "AI provider unavailable";
      return NextResponse.json({
        report: `${message}. Conservative fallback generated from uploaded scoring data.`,
        insights: fallbackInsights(selectedCompanies, scores, message),
        sources: results.map((item) => ({ title: item.title, url: item.url })),
        fallback: true,
      });
    }

    return NextResponse.json({
      report: parsed.report || "AI scoring layer completed.",
      insights: parsed.insights || [],
      sources: results.map((item) => ({ title: item.title, url: item.url })),
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "AI scoring layer failed." }, { status: 500 });
  }
}
