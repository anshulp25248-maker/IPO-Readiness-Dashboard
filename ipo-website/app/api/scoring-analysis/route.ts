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

async function tavilySearch(query: string, maxResults = 8) {
  const apiKey = envValue("TAVILY_API_KEY");
  if (!apiKey) return { results: [], status: "Tavily is not configured; add TAVILY_API_KEY for public-source enrichment." };
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
  if (!response.ok) throw new Error(`Tavily scoring search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  const results = data.results ?? [];
  return {
    results,
    status: results.length ? `Tavily returned ${results.length} public-source snippets.` : "Tavily returned no public-source snippets.",
  };
}

async function tavilySearchMany(queries: string[]) {
  const batches = await Promise.all(queries.map((query) => tavilySearch(query, 5).catch((error) => ({
    results: [],
    status: error instanceof Error ? error.message : "Tavily search failed.",
  }))));
  const seen = new Set<string>();
  const results = batches
    .flatMap((batch) => batch.results)
    .filter((item) => {
      const key = item.url || `${item.title}-${item.content}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 16);

  const statuses = batches.map((batch) => batch.status).filter(Boolean);
  return {
    results,
    status: results.length
      ? `Tavily returned ${results.length} public-source snippets across Zauba/Tofler/MCA and sector/thematic searches.`
      : statuses.join(" | ") || "Tavily returned no public-source snippets.",
  };
}

function sourceContext(results: SearchResult[]) {
  return results
    .slice(0, 12)
    .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 420)}`)
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
        cleanDisplayText(insight.rationale || "") ||
        "AI returned an incomplete structured insight, so this view is anchored to the uploaded deterministic score and public-source availability.",
      strengths: Array.isArray(insight.strengths) ? insight.strengths.map((item) => cleanDisplayText(String(item))).slice(0, 4) : [],
      redFlags: Array.isArray(insight.redFlags) ? insight.redFlags.map((item) => cleanDisplayText(String(item))).slice(0, 4) : [],
      missingData: Array.isArray(insight.missingData) ? insight.missingData.map((item) => cleanDisplayText(String(item))).slice(0, 5) : [],
    };
  });
}

function cleanDisplayText(value: string) {
  return value
    .replace(/^\s*["{]?\s*(report|rationale|analysis|insight)\s*["']?\s*:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
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
        cleanDisplayText(extractedRationale || "") ||
        (narrative
          ? `${cleanDisplayText(narrative)} The parser score remains the controlling score because uploaded data is the source of truth.`
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

function thematicSearchQueries(companies: Company[]) {
  const top = companies.slice(0, 5);
  const identityQuery = top
    .map((company) => `${company.name} ${company.cin}`)
    .join(" OR ");
  const sectorTerms = [...new Set(top.map((company) => `${company.sector} ${company.nicCode} ${company.activity}`))]
    .slice(0, 4)
    .join(" OR ");

  return [
    `${identityQuery} site:zaubacorp.com paid up capital authorized capital directors filing status`,
    `${identityQuery} site:tofler.in paid up capital financials directors annual filings`,
    `${identityQuery} MCA company master data charges annual return directors filing compliance`,
    `${sectorTerms} India sector report IBEF FICCI government outlook policy market size CAGR`,
    `${sectorTerms} industry thematic report India investment outlook regulatory stance`,
    `${sectorTerms} competitors listed peers India business activity outlook`,
  ].filter((query) => query.trim().length > 25);
}

export async function POST(request: Request) {
  try {
    const { companies, scores } = (await request.json()) as {
      companies?: Company[];
      scores?: Record<string, number>;
    };
    const selectedCompanies = (companies || []).slice(0, 3);
    if (!selectedCompanies.length) return NextResponse.json({ error: "No companies supplied." }, { status: 400 });

    let results: SearchResult[] = [];
    let sourceStatus = "";
    try {
      const tavily = await tavilySearchMany(thematicSearchQueries(selectedCompanies));
      results = tavily.results;
      sourceStatus = tavily.status;
    } catch (error) {
      sourceStatus = error instanceof Error ? error.message : "Tavily public-source enrichment failed.";
    }

    const prompt = `You are the AI scoring layer for Scout Smarter. Analyze uploaded Indian company screening data plus public feed snippets. Keep every statement source-aware and conservative.

Do not invent facts. Uploaded paid-up capital, authorized capital, CIN, sector/NIC/activity, factor scores, and deterministic total score are the source of truth.
The deterministic parser score has already applied the original dashboard logic, including paid-up capital comparison across all companies in the uploaded file, sector strength, geography, business-model scarcity, director profile, authorised/paid-up ratio, and filing compliance. The AI layer must not recalculate or override those factor scores. It must explain the parser score and improve the investment insight by reading public-source evidence.

For filing compliance and company master details, use the public feed snippets from Zauba, Tofler, MCA, and similar sources where available. If the feed does not verify filing dates, charges, directors, authorised capital, or paid-up capital, write a diligence-quality explanation of the missing evidence and request MCA master data, annual returns, and financial statements.

For sector and industry, generate a thematic analysis using public sector reports, government stance, industry outlook, policy support, demand drivers, cyclicality, competitive intensity, and business activity context. This must be written as a proper report, not in small words or short labels.

COMPANY DATA
${JSON.stringify(selectedCompanies.map((company) => scoringTruth(company, scores?.[company.id])), null, 2)}

PUBLIC FEED
${sourceStatus}

${sourceContext(results)}

Return strict JSON only. Do not wrap JSON in markdown. Do not include a second JSON object. Do not add comments. Escape every newline inside strings as \\n.
{
  "report": "detailed investment-banking style memo with complete paragraphs covering parser score reconciliation, Zauba/Tofler/MCA filing-compliance evidence, sector and industry thematic report, business-model uniqueness, director quality, red flags, positives, and final investment implication",
  "insights": [
    {
      "companyId": "must match input id",
      "companyName": "...",
      "aiScore": 0,
      "recommendation": "Invest / Watchlist / Reject / Data Insufficient",
      "rationale": "one detailed paragraph explaining the parser score, public-source evidence, sector/industry outlook, business model, filing compliance, and investment implication",
      "strengths": ["complete-sentence positive paragraph"],
      "redFlags": ["RED FLAG: complete-sentence risk paragraph"],
      "missingData": ["complete-sentence missing evidence paragraph"]
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
