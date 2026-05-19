import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";

export const runtime = "nodejs";

type SearchResult = {
  title?: string;
  url?: string;
  content?: string;
};

type ReportType =
  | "sector-analysis"
  | "industry-analysis"
  | "competitor-analysis"
  | "director-profile"
  | "company-analysis"
  | "comprehensive-cdr";

type SectionReportType = Exclude<ReportType, "comprehensive-cdr">;

const comprehensiveSectionTypes: SectionReportType[] = [
  "sector-analysis",
  "industry-analysis",
  "competitor-analysis",
  "director-profile",
  "company-analysis",
];

const tavilyUrl = "https://api.tavily.com/search";

const cdrTaskConfig: Record<ReportType, { label: string; envPrefix: string; maxTokens: number }> = {
  "sector-analysis": {
    label: "Sector Analysis",
    envPrefix: "CDR_SECTOR_ANALYSIS",
    maxTokens: 1300,
  },
  "industry-analysis": {
    label: "Industry Analysis",
    envPrefix: "CDR_INDUSTRY_ANALYSIS",
    maxTokens: 1200,
  },
  "competitor-analysis": {
    label: "Competitor Analysis",
    envPrefix: "CDR_COMPETITOR_ANALYSIS",
    maxTokens: 1400,
  },
  "director-profile": {
    label: "Director Profile",
    envPrefix: "CDR_DIRECTOR_PROFILE",
    maxTokens: 1400,
  },
  "company-analysis": {
    label: "Company Analysis",
    envPrefix: "CDR_COMPANY_ANALYSIS",
    maxTokens: 1300,
  },
  "comprehensive-cdr": {
    label: "Comprehensive CDR",
    envPrefix: "CDR_COMPREHENSIVE_CDR",
    maxTokens: 1200,
  },
};

function cdrEnvValue(type: ReportType, suffix: "GROQ_API_KEY" | "TAVILY_API_KEY") {
  const primary = `${cdrTaskConfig[type].envPrefix}_${suffix}`;
  const legacy = `${suffix.replace("_API_KEY", "")}_API_KEY_${cdrTaskConfig[type].envPrefix}`;
  return envValue(primary) || envValue(legacy) || envValue(suffix);
}

async function tavilySearch(query: string, type: ReportType) {
  const apiKey = cdrEnvValue(type, "TAVILY_API_KEY");
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
      max_results: 4,
    }),
    cache: "no-store",
  });

  if (!response.ok) throw new Error(`Tavily search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

async function tavilySearchMany(queries: string[], type: ReportType) {
  const batches = await Promise.all(queries.map((query) => tavilySearch(query, type).catch(() => [])));
  const seen = new Set<string>();
  return batches
    .flat()
    .filter((item) => {
      const key = item.url || `${item.title}-${item.content}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 8);
}

function sourceContext(results: SearchResult[]) {
  if (!results.length) return "No live feed results returned.";
  return results
    .map((item, index) =>
      [`${index + 1}. ${item.title || "Untitled"}`, `URL: ${item.url || "NA"}`, `Snippet: ${(item.content || "").slice(0, 320)}`].join("\n"),
    )
    .join("\n\n");
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function retryDelayMs(message: string) {
  const match = message.match(/try again in\s+([0-9.]+)s/i);
  const seconds = match ? Number(match[1]) : 8;
  return Math.max(3000, Math.min(15000, Math.ceil((Number.isFinite(seconds) ? seconds : 8) * 1000) + 1000));
}

async function generateAiTextWithRetry(options: Parameters<typeof generateAiText>[0], retries = 2) {
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await generateAiText(options);
    } catch (error) {
      const message = error instanceof Error ? error.message : "";
      if (attempt >= retries || !/429|rate limit|try again/i.test(message)) throw error;
      await sleep(retryDelayMs(message));
    }
  }

  return generateAiText(options);
}

function companyTruth(company: Company, score: number | undefined) {
  return {
    name: company.name,
    cin: company.cin,
    sector: company.sector,
    nicCode: company.nicCode,
    activity: company.activity,
    city: company.city,
    state: company.state,
    status: company.status,
    uploadedPaidUpCapital: company.paidUpCapital,
    uploadedAuthorizedCapital: company.authorizedCapital,
    uploadedPaidUpCapitalRaw: company.paidUpCapitalValue ?? null,
    uploadedAuthorizedCapitalRaw: company.authorizedCapitalValue ?? null,
    uploadedLastFiling: company.lastFiling,
    uploadedDirector: company.director,
    uploadedCompetitors: company.competitors,
    deterministicScoutScore: score ?? null,
    deterministicFactorScores: company.factors,
  };
}

function reportInstructions(type: ReportType) {
  const common = `Formatting rules:
- Use only short bold section headings. Each heading must be one or two lines at most, for example "**Director Profile**".
- Under each heading, write detailed analyst body paragraphs. Use 4-7 sentences per paragraph where evidence is available.
- Do not create large markdown titles, numbered lists, long bullet lists, or tables.
- Keep the report focused on the requested tab only.
- Take public data from the live feed, Google-indexed public snippets, MCA, Zauba, Tofler, company websites, LinkedIn public snippets, credible news, government sources, and industry bodies when available.
- Distinguish verified public facts from inference. If evidence is thin, write "Data Not Available" and explain which document or source should be requested.
- Red flags must begin with "RED FLAG:" and positives must begin with "POSITIVE:".`;

  const instructions: Record<ReportType, string> = {
    "sector-analysis": `Create a real-time sector analysis using this order: Sector Definition, Government & Policy Stance, Market Growth, Listed Benchmarks, Sector Risks, Diligence Questions. ${common}`,
    "industry-analysis": `Create a real-time NIC/basic-industry analysis using this order: Industry Classification, Value Chain, Buyer & Supplier Power, Margin & Working Capital Pattern, Compliance Requirements, Industry Outlook. ${common}`,
    "competitor-analysis": `Create a real-time competitor analysis using this order: Peer Universe, Relative Scale, Product Overlap, Competitive Moat, Public Benchmarks, Threats & Questions. Use only evidence-backed competitors; do not invent peer names. ${common}`,
    "director-profile": `Create a real-time director/promoter profile using this order: Director Profile, DIN & Appointment Signals, Public Background, Directorship & Disqualification Checks, Governance Concerns, Management Quality View. Keep the "Director Profile" heading to one line. ${common}`,
    "company-analysis": `Create a real-time company analysis using this order: Company Snapshot, Business Model, Capital Structure, Filing & Charge Signals, Score Reconciliation, Investment View. ${common}`,
    "comprehensive-cdr": `Create the final recommendation for a comprehensive CDR after reading the generated section reports. Do not repeat every section. Reconcile Sector Analysis, Industry Analysis, Competitor Analysis, Director Profile, and Company Analysis into a final investment view. ${common}`,
  };

  return instructions[type];
}

function searchQueries(company: Company, type: ReportType) {
  const identity = `${company.name} ${company.cin}`;
  const sector = `${company.sector} ${company.nicCode} ${company.activity}`;
  const common = [
    `${identity} MCA Zauba Tofler paid up capital authorized capital directors charges filings`,
    `${identity} news litigation funding contracts India`,
    `${identity} Google public search company profile directors competitors financials`,
  ];

  const byType: Record<ReportType, string[]> = {
    "sector-analysis": [
      `${sector} India sector report IBEF FICCI government outlook PLI policy market size CAGR`,
      `${company.sector} India ministry policy annual report market outlook investment`,
      `${company.sector} ${company.activity} India market size listed peers risks tailwinds`,
    ],
    "industry-analysis": [
      `${sector} NIC industry analysis India value chain margin working capital association report`,
      `${company.activity} India industry report government outlook regulatory compliance`,
    ],
    "competitor-analysis": [
      `${identity} competitors peers India revenue funding Tofler Zauba`,
      `${company.activity} India private company competitors listed peers`,
      `${company.name} similar companies India competitors Google Zauba Tofler LinkedIn`,
    ],
    "director-profile": [
      `${identity} directors DIN LinkedIn promoter MCA disqualification`,
      `${company.director?.name || ""} ${company.name} director DIN directorships`,
      `${company.director?.name || ""} ${company.cin} Zauba Tofler MCA director profile directorships`,
    ],
    "company-analysis": [
      `${identity} company profile financials revenue directors charges Tofler Zauba`,
      `${identity} annual return balance sheet MCA filings credit rating litigation`,
    ],
    "comprehensive-cdr": [
      `${sector} India sector report IBEF FICCI government outlook PLI policy market size CAGR`,
      `${company.activity} India industry value chain competitors listed peers`,
      `${identity} competitors directors financials charges litigation news Tofler Zauba MCA`,
    ],
  };

  return [...common, ...byType[type]].filter(Boolean);
}

function reportHeadings(type: ReportType) {
  const structures: Record<ReportType, string[]> = {
    "sector-analysis": [
      "Sector Definition",
      "Government & Policy Stance",
      "Market Growth",
      "Listed Benchmarks",
      "Sector Risks",
      "Diligence Questions",
    ],
    "industry-analysis": [
      "Industry Classification",
      "Value Chain",
      "Buyer & Supplier Power",
      "Margin & Working Capital Pattern",
      "Compliance Requirements",
      "Industry Outlook",
    ],
    "competitor-analysis": [
      "Peer Universe",
      "Relative Scale",
      "Product Overlap",
      "Competitive Moat",
      "Public Benchmarks",
      "Threats & Questions",
    ],
    "director-profile": [
      "Director Profile",
      "DIN & Appointment Signals",
      "Public Background",
      "Directorship & Disqualification Checks",
      "Governance Concerns",
      "Management Quality View",
    ],
    "company-analysis": [
      "Company Snapshot",
      "Business Model",
      "Capital Structure",
      "Filing & Charge Signals",
      "Score Reconciliation",
      "Investment View",
    ],
    "comprehensive-cdr": [
      "Company Snapshot",
      "Sector Analysis",
      "Industry Analysis",
      "Competitor Analysis",
      "Director Profile",
      "Company Analysis",
      "Final Recommendation",
    ],
  };

  return structures[type];
}

function buildHeadingPrompt(
  company: Company,
  score: number | undefined,
  internalInfo: string | undefined,
  selectedReportType: ReportType,
  heading: string,
  results: SearchResult[],
  compact = false,
) {
  const sourceTruth = compact
    ? {
        name: company.name,
        cin: company.cin,
        sector: company.sector,
        nicCode: company.nicCode,
        activity: company.activity,
        paidUpCapital: company.paidUpCapital,
        authorizedCapital: company.authorizedCapital,
        director: company.director,
        score,
        factors: company.factors,
      }
    : companyTruth(company, score);
  const sourceList = compact
    ? "Use uploaded parser data, MCA, Zauba, Tofler, Google-indexed public snippets, company websites, credible news, public LinkedIn snippets, and sector/government sources."
    : `SOURCES TO SEARCH AND USE, IN ORDER:
1. MCA21 Portal (mca.gov.in) for incorporation data, directors, charges, filings
2. Zauba Corp / Tofler / Veratech for company profile and public financial snippets
3. Google-indexed public snippets and company website pages
4. RoC filings for annual returns and auditor reports
5. MOSPI / NIC Code Directory for sector classification
6. IBEF / FICCI / Industry Association Reports for sector outlook
7. Business news sources for recent developments
8. LinkedIn public snippets for business model and management background
9. Court records / MCA charge search for litigation and charges`;

  return `You are a senior investment analyst specializing in unlisted/private companies in India. Research ${company.name} (CIN if known: ${company.cin}) and produce only one detailed subsection for the requested CDR tab. Do not hallucinate or assume any data. Every factual data point must come from uploaded MCA/feed-parser data or the live feed below. If data is unavailable, explicitly state "Data Not Available" for that field.

${sourceList}

SOURCE OF TRUTH FROM UPLOADED EXCEL/MCA PARSER
${JSON.stringify(sourceTruth, null, compact ? 0 : 2)}

STRICT CONSISTENCY RULE
The uploaded Excel/MCA parser values above are the source of truth for CIN, legal name, paid-up capital, authorized capital, sector/NIC/activity, filing date, uploaded director fields, factor scores, and deterministic Scout Score. Do not replace paid-up capital, authorized capital, CIN, or score with a public-source value. If Zauba, Tofler, MCA, or another source appears to conflict, state the discrepancy in a paragraph titled "Public-Source Discrepancy" and recommend manual verification.

OPTIONAL INTERNAL INFORMATION FROM USER
${internalInfo?.trim() || "NA"}

LIVE FEED
${sourceContext(compact ? results.slice(0, 5) : results)}

REQUESTED REPORT TYPE
${selectedReportType}

REQUESTED SUBSECTION
${heading}

REPORT INSTRUCTIONS
${reportInstructions(selectedReportType)}

OUTPUT STRUCTURE
Use exactly this bold heading, on its own line:
**${heading}**

FINAL INSTRUCTIONS:
- Do not invent any number, name, director, DIN, date, filing, or fact.
- If a section has no data available from public sources, write: "Data Not Available - recommend requesting directly from company".
- Cite the source platform for every major data point.
- Keep the tone objective, like a SEBI-registered research analyst.
- Keep the heading short. Put detail in the body under the heading.
- Write 2 detailed analyst paragraphs where evidence exists. If evidence is thin, write 1 detailed paragraph explaining the limitation and what diligence should be requested.
- Maintain uploaded paid-up capital and authorized capital exactly as shown in the source-of-truth block.
- Assess business-model uniqueness using uploaded NIC/activity data plus live public evidence; if public evidence is thin, say so and keep the conclusion conservative.
- This subsection will be stitched with the other CDR subsections, so do not repeat unrelated headings.`;
}

async function generateCdrHeading(
  company: Company,
  score: number | undefined,
  internalInfo: string | undefined,
  selectedReportType: ReportType,
  heading: string,
  results: SearchResult[],
  baseOptions: {
    task: "cdr";
    provider: "groq";
    apiKey: string;
    providerLabel: string;
    system: string;
    temperature: number;
  },
) {
  try {
    const ai = await generateAiTextWithRetry({
      ...baseOptions,
      prompt: buildHeadingPrompt(company, score, internalInfo, selectedReportType, heading, results),
      maxTokens: 650,
    });
    return ai.text || `**${heading}**\nData Not Available - no subsection returned.`;
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (!/413|Request too large|TPM|tokens per minute/i.test(message)) throw error;
    const ai = await generateAiTextWithRetry({
      ...baseOptions,
      prompt: buildHeadingPrompt(company, score, internalInfo, selectedReportType, heading, results.slice(0, 3), true),
      maxTokens: 450,
    });
    return ai.text || `**${heading}**\nData Not Available - no subsection returned.`;
  }
}

async function generateCdrSection(
  company: Company,
  score: number | undefined,
  internalInfo: string | undefined,
  selectedReportType: ReportType,
) {
  const taskConfig = cdrTaskConfig[selectedReportType];
  const results = await tavilySearchMany(searchQueries(company, selectedReportType), selectedReportType);
  const baseOptions = {
    task: "cdr" as const,
    provider: "groq" as const,
    apiKey: cdrEnvValue(selectedReportType, "GROQ_API_KEY"),
    providerLabel: `Groq ${taskConfig.label}`,
    system: "You are Scout Smarter, an investment banking analyst preparing detailed CDR sections from uploaded parser data and live Tavily evidence. Be source-aware, skeptical, and never hallucinate.",
    temperature: 0.12,
  };
  const parts = [];

  for (const heading of reportHeadings(selectedReportType)) {
    parts.push(await generateCdrHeading(company, score, internalInfo, selectedReportType, heading, results, baseOptions));
  }

  return {
    reportType: selectedReportType,
    label: taskConfig.label,
    report: parts.join("\n\n") || `**${taskConfig.label}**\nData Not Available - no CDR report returned.`,
    sources: results.map((item) => ({ title: item.title, url: item.url })),
    provider: baseOptions.providerLabel,
    model: envValue("GROQ_MODEL_CDR") || envValue("GROQ_MODEL") || "llama-3.1-8b-instant",
  };
}

function dedupeSources(sources: Array<{ title?: string; url?: string }>) {
  const seen = new Set<string>();
  return sources.filter((source) => {
    const key = source.url || source.title;
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function generateComprehensiveCdr(company: Company, score: number | undefined, internalInfo: string | undefined) {
  const sectionResults = await Promise.all(
    comprehensiveSectionTypes.map((type) => generateCdrSection(company, score, internalInfo, type)),
  );
  const sectionDigest = sectionResults
    .map((section) => `SECTION: ${section.label}\n${section.report.slice(0, 1200)}`)
    .join("\n\n");
  const taskConfig = cdrTaskConfig["comprehensive-cdr"];
  const finalAi = await generateAiTextWithRetry({
    task: "cdr",
    provider: "groq",
    apiKey: cdrEnvValue("comprehensive-cdr", "GROQ_API_KEY"),
    providerLabel: `Groq ${taskConfig.label}`,
    system: "You are Scout Smarter, preparing only the final recommendation for a comprehensive CDR. Use the section reports as evidence and do not invent new facts.",
    prompt: `Create the final recommendation section for this comprehensive CDR.

SOURCE OF TRUTH
${JSON.stringify(companyTruth(company, score), null, 2)}

OPTIONAL INTERNAL INFORMATION
${internalInfo?.trim() || "NA"}

GENERATED CDR SECTION REPORTS
${sectionDigest}

Use this structure only:
**Final Recommendation**

Write 2-4 detailed paragraphs. Reconcile sector attractiveness, industry structure, competitor position, director quality, capital/filing quality, public-data confidence, red flags, positives, and the final Invest / Watchlist / Reject / Data Insufficient view.`,
    temperature: 0.1,
    maxTokens: taskConfig.maxTokens,
  });
  const report = [
    `**Comprehensive CDR**\nThis document combines the live CDR section reports generated from the same section-specific Groq and Tavily lanes used by the CDR tabs. Uploaded MCA/parser data remains the source of truth for identity, capital, score, sector, activity, and director fields.`,
    ...sectionResults.map((section) => `**${section.label}**\n${section.report}`),
    finalAi.text || "**Final Recommendation**\nData Not Available - final recommendation could not be generated.",
  ].join("\n\n");

  return {
    report,
    sources: dedupeSources(sectionResults.flatMap((section) => section.sources)),
    provider: [...sectionResults.map((section) => section.provider), finalAi.provider].join(", "),
    model: [...new Set([...sectionResults.map((section) => section.model), finalAi.model])].join(", "),
  };
}

export async function POST(request: Request) {
  try {
    const { company, score, internalInfo, reportType } = (await request.json()) as {
      company?: Company;
      score?: number;
      internalInfo?: string;
      reportType?: ReportType;
    };
    if (!company) {
      return NextResponse.json({ error: "No company supplied." }, { status: 400 });
    }

    const selectedReportType: ReportType = reportType || "company-analysis";
    const result =
      selectedReportType === "comprehensive-cdr"
        ? await generateComprehensiveCdr(company, score, internalInfo)
        : await generateCdrSection(company, score, internalInfo, selectedReportType);

    return NextResponse.json({
      report: result.report || "No CDR report returned.",
      sources: result.sources,
      provider: result.provider,
      model: result.model,
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "CDR generation failed." },
      { status: 500 },
    );
  }
}
