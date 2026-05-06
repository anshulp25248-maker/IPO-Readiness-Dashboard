import fs from "node:fs";
import path from "node:path";
import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";

export const runtime = "nodejs";

type SearchResult = { title?: string; url?: string; content?: string };

const groqUrl = "https://api.groq.com/openai/v1/chat/completions";
const tavilyUrl = "https://api.tavily.com/search";

function readEnvFile(filePath: string) {
  if (!fs.existsSync(filePath)) return {};
  return fs.readFileSync(filePath, "utf8").split(/\r?\n/).reduce<Record<string, string>>((values, line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) return values;
    const equalsAt = trimmed.indexOf("=");
    if (equalsAt === -1) return values;
    values[trimmed.slice(0, equalsAt).trim()] = trimmed.slice(equalsAt + 1).trim().replace(/^["']|["']$/g, "");
    return values;
  }, {});
}

function envValue(key: string) {
  if (process.env[key]) return process.env[key];
  for (const candidate of [path.join(process.cwd(), ".env.local"), path.join(process.cwd(), ".env"), path.join(process.cwd(), "..", ".env")]) {
    const value = readEnvFile(candidate)[key];
    if (value) return value;
  }
  return "";
}

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
    .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 650)}`)
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
    const selectedCompanies = (companies || []).slice(0, 10);
    if (!selectedCompanies.length) return NextResponse.json({ error: "No companies supplied." }, { status: 400 });

    const groqKey = envValue("GROQ_API_KEY");
    if (!groqKey) return NextResponse.json({ error: "GROQ_API_KEY is missing in .env." }, { status: 500 });

    const query = `${selectedCompanies.slice(0, 5).map((company) => `${company.name} ${company.cin}`).join(" OR ")} MCA Zauba Tofler sector directors India`;
    const results = await tavilySearch(query);

    const prompt = `You are the AI scoring layer for Scout Smarter. The strict quantitative factor score has already been calculated from the uploaded Excel/MCA parser. Your job is to analyze the top companies with Groq using public feed evidence and provide an evidence-aware investment readiness layer.

Do not invent facts. Use Data Not Available when the feed does not verify a point. Do not override or restate the deterministic uploaded data with conflicting public snippets. The uploaded paid-up capital, authorized capital, CIN, sector/NIC/activity, deterministic factor scores, and deterministic total score are the source of truth. If a public source conflicts with those fields, report it as a discrepancy and recommend manual verification.

SOURCE-OF-TRUTH COMPANY DATA AND BASE SCORES
${JSON.stringify(selectedCompanies.map((company) => scoringTruth(company, scores?.[company.id])), null, 2)}

LIVE FEED
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

AI score should reflect public-data confidence, real-time sector attractiveness, business-model uniqueness from uploaded NIC/activity plus public evidence, management/funding/compliance evidence, red flags, and missing-data penalties. If public evidence is thin, use Data Insufficient and a conservative AI score.`;

    const response = await fetch(groqUrl, {
      method: "POST",
      headers: { Authorization: `Bearer ${groqKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: envValue("GROQ_MODEL") || "llama-3.3-70b-versatile",
        messages: [
          { role: "system", content: "Return only valid JSON. You are a skeptical investment analyst. Never hallucinate." },
          { role: "user", content: prompt },
        ],
        temperature: 0.1,
        max_tokens: 5500,
      }),
      cache: "no-store",
    });

    if (!response.ok) throw new Error(`Groq scoring analysis failed with ${response.status}`);
    const data = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> };
    const parsed = parseJson(data.choices?.[0]?.message?.content || "{}");

    return NextResponse.json({
      report: parsed.report || "AI scoring layer completed.",
      insights: parsed.insights || [],
      sources: results.map((item) => ({ title: item.title, url: item.url })),
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "AI scoring layer failed." }, { status: 500 });
  }
}
