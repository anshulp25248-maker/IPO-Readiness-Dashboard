import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";
import { scoringJsonReportFormat } from "../_lib/report-format";

export const runtime = "nodejs";

type SearchResult = { title?: string; url?: string; content?: string };
type ParsedInsight = {
  companyId?: string;
  companyName?: string;
  aiScore?: number;
  recommendation?: string;
  rationale?: string;
  strengths?: string[];
  redFlags?: string[];
  missingData?: string[];
};

const tavilyUrl = "https://api.tavily.com/search";

async function tavilySearch(query: string) {
  const apiKey = envValue("TAVILY_API_KEY");
  if (!apiKey) return { results: [], status: "Tavily is not configured; add TAVILY_API_KEY for public-source enrichment." };
  const response = await fetch(tavilyUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ api_key: apiKey, query, search_depth: "advanced", max_results: 10 }),
    cache: "no-store",
  });
  if (!response.ok) throw new Error(`Tavily scoring search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  const results = data.results ?? [];
  return {
    results,
    status: results.length ? `Tavily returned ${results.length} public-source snippets.` : "Tavily returned no public-source snippets.",
  };
}

function sourceContext(results: SearchResult[]) {
  return results
    .slice(0, 4)
    .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 300)}`)
    .join("\n\n") || "No public feed results returned.";
}

function extractJsonObject(content: string) {
  const cleaned = content
    .replace(/```json/gi, "```")
    .replace(/```/g, "")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
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

function repairJson(content: string) {
  return extractJsonObject(content)
    .replace(/,\s*([}\]])/g, "$1")
    .replace(/([}\]"])\s*;\s*([{\["])/g, "$1,$2")
    .replace(/"\s*\n\s*"/g, '",\n"')
    .replace(/]\s*\n\s*"/g, '],\n"')
    .replace(/}\s*\n\s*"/g, '},\n"');
}

function parseJson(content: string) {
  const candidates = [extractJsonObject(content), repairJson(content)];
  let lastError: unknown;
  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate) as {
    report?: string;
        insights?: ParsedInsight[];
      };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError instanceof Error ? lastError : new Error("AI returned invalid JSON");
}

function extractQuotedValue(content: string, key: string) {
  const pattern = new RegExp(`"${key}"\\s*:\\s*"((?:\\\\.|[^"\\\\])*)"`,"s");
  const match = content.match(pattern);
  if (!match) return "";
  try {
    return JSON.parse(`"${match[1]}"`) as string;
  } catch {
    return match[1].replace(/\\"/g, '"').replace(/\\n/g, "\n");
  }
}

function extractStringArray(content: string, key: string) {
  const pattern = new RegExp(`"${key}"\\s*:\\s*\\[([\\s\\S]*?)\\]`, "s");
  const match = content.match(pattern);
  if (!match) return [];
  return [...match[1].matchAll(/"((?:\\.|[^"\\])*)"/g)]
    .map((item) => {
      try {
        return JSON.parse(`"${item[1]}"`) as string;
      } catch {
        return item[1].replace(/\\"/g, '"').replace(/\\n/g, " ");
      }
    })
    .filter(Boolean)
    .slice(0, 4);
}

function clampScore(value: unknown, fallback: number) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.min(100, Math.round(numeric)));
}

function normalizeInsights(companies: Company[], scores: Record<string, number> | undefined, insights: ParsedInsight[]) {
  return companies.map((company, index) => {
    const deterministicScore = scores?.[company.id] ?? 0;
    const insight =
      insights.find((item) => item.companyId === company.id) ||
      insights.find((item) => item.companyName?.toLowerCase().trim() === company.name.toLowerCase().trim()) ||
      insights[index] ||
      {};

    return {
      companyId: company.id,
      companyName: insight.companyName || company.name,
      aiScore: clampScore(deterministicScore, deterministicScore),
      recommendation:
        insight.recommendation ||
        (deterministicScore >= 85 ? "Watchlist" : deterministicScore >= 70 ? "Data Insufficient" : "Reject"),
      rationale:
        insight.rationale ||
        "AI returned an incomplete structured insight, so this view is anchored to the uploaded deterministic score and public-source availability.",
      strengths: Array.isArray(insight.strengths) ? insight.strengths.map(String).slice(0, 4) : [],
      redFlags: Array.isArray(insight.redFlags) ? insight.redFlags.map(String).slice(0, 4) : [],
      missingData: Array.isArray(insight.missingData) ? insight.missingData.map(String).slice(0, 5) : [],
    };
  });
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
      rationale: `${reason}. Showing a conservative fallback view from uploaded factor scores only; run again after the active AI provider quota resets for live public-source analysis.`,
      strengths: [
        `Deterministic Scout Score is ${Math.round(score)}/100.`,
        `Sector factor is ${company.factors.sector.toFixed(1)}/10 and business-model factor is ${company.factors.businessModel.toFixed(1)}/10.`,
      ],
      redFlags,
      missingData: ["Live AI qualitative analysis was not returned."],
    };
  });
}

function narrativeInsights(companies: Company[], scores: Record<string, number> | undefined, aiText: string, sourceStatus: string) {
  const extractedReport = extractQuotedValue(aiText, "report");
  const narrative = (extractedReport || aiText)
    .replace(/[{}[\]"]/g, " ")
    .replace(/\s+/g, " ")
    .slice(0, 700);
  return companies.map((company) => {
    const score = scores?.[company.id] ?? 0;
    const extractedRationale = extractQuotedValue(aiText, "rationale");
    const strengths = extractStringArray(aiText, "strengths");
    const redFlags = extractStringArray(aiText, "redFlags");
    const missingData = extractStringArray(aiText, "missingData");
    return {
      companyId: company.id,
      companyName: company.name,
      aiScore: Math.max(0, Math.min(100, Math.round(score))),
      recommendation: score >= 85 ? "Watchlist" : score >= 70 ? "Data Insufficient" : "Reject",
      rationale:
        extractedRationale ||
        (narrative
          ? `${narrative} The parser score remains the controlling score because uploaded data is the source of truth.`
          : `The AI provider returned an incomplete response. ${sourceStatus}`),
      strengths: strengths.length
        ? strengths
        : [
            `POSITIVE: Parsed Scout Score is ${Math.round(score)}/100, derived from the selected scoring factors rather than model opinion.`,
            `POSITIVE: Business-model score is ${company.factors.businessModel.toFixed(1)}/10 based on uploaded NIC/activity scarcity logic.`,
          ],
      redFlags: redFlags.length
        ? redFlags
        : [
            company.lastFiling === "NA" || company.lastFiling === "Data Not Available"
              ? "RED FLAG: Filing date is not available in uploaded data and should be verified before investment underwriting."
              : "",
          ].filter(Boolean),
      missingData: missingData.length ? missingData : ["Structured AI JSON was incomplete and was repaired into a usable insight."],
    };
  });
}

async function aiJson(prompt: string) {
  const result = await generateAiText({
    task: "scoring",
    system: "Return only valid compact JSON. Never hallucinate.",
    prompt,
    temperature: 0.05,
    maxTokens: 1800,
    responseJson: true,
  });
  return result;
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
    let results: SearchResult[] = [];
    let sourceStatus = "";
    try {
      const tavily = await tavilySearch(query);
      results = tavily.results;
      sourceStatus = tavily.status;
    } catch (error) {
      sourceStatus = error instanceof Error ? error.message : "Tavily public-source enrichment failed.";
    }

    const prompt = `You are the AI scoring layer for Scout Smarter. Analyze uploaded Indian company screening data plus public feed snippets. Keep every statement source-aware and conservative.

Do not invent facts. Uploaded paid-up capital, authorized capital, CIN, sector/NIC/activity, factor scores, and deterministic total score are the source of truth.

COMPANY DATA
${JSON.stringify(selectedCompanies.map((company) => scoringTruth(company, scores?.[company.id])), null, 2)}

PUBLIC FEED
${sourceStatus}

${sourceContext(results)}

Return strict JSON only. Do not wrap JSON in markdown. Do not include a second JSON object. Do not add comments. Escape every newline inside strings as \\n.
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

AI score should reflect public-data confidence, sector attractiveness, business-model uniqueness, management/compliance evidence, red flags, and missing-data penalties.
The aiScore field must equal the deterministicScore supplied in COMPANY DATA. The AI layer may explain risk and public-source confidence, but it must not change the parser score.

${scoringJsonReportFormat}`;

    let parsed: ReturnType<typeof parseJson>;
    try {
      const ai = await aiJson(prompt);
      try {
        parsed = parseJson(ai.text);
      } catch {
        return NextResponse.json({
          report: ai.text || "AI scoring layer returned narrative analysis.",
          insights: narrativeInsights(selectedCompanies, scores, ai.text, sourceStatus),
          sources: results.map((item) => ({ title: item.title, url: item.url })),
          sourceStatus,
          provider: ai.provider,
          model: ai.model,
          fallback: true,
        });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : "AI provider unavailable";
      return NextResponse.json({
        report: `${message}. Conservative fallback generated from uploaded scoring data.`,
        insights: fallbackInsights(selectedCompanies, scores, message),
        sources: results.map((item) => ({ title: item.title, url: item.url })),
        sourceStatus,
        fallback: true,
      });
    }

    return NextResponse.json({
      report: parsed.report || "AI scoring layer completed.",
      insights: normalizeInsights(selectedCompanies, scores, parsed.insights || []),
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      sourceStatus,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "AI scoring layer failed." }, { status: 500 });
  }
}
