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

const tavilyUrl = "https://api.tavily.com/search";

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

  if (!response.ok) throw new Error(`Tavily search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

async function tavilySearchMany(queries: string[]) {
  const batches = await Promise.all(queries.map((query) => tavilySearch(query).catch(() => [])));
  const seen = new Set<string>();
  return batches
    .flat()
    .filter((item) => {
      const key = item.url || `${item.title}-${item.content}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 18);
}

function sourceContext(results: SearchResult[]) {
  if (!results.length) return "No live feed results returned.";
  return results
    .map((item, index) =>
      [`${index + 1}. ${item.title || "Untitled"}`, `URL: ${item.url || "NA"}`, `Snippet: ${(item.content || "").slice(0, 700)}`].join("\n"),
    )
    .join("\n\n");
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
  const common = `Write in complete analyst paragraphs, not bullet points. Use clear section headings, but do not use numbered markdown lists. Every section must distinguish verified facts from inference. Red flags must begin with "RED FLAG:" and investment positives must begin with "POSITIVE:".`;

  const instructions: Record<ReportType, string> = {
    "sector-analysis": `Create a detailed investment-banking sector report. Cover sector definition, market size, growth, policy stance, PLI or government incentives, regulatory bodies, listed benchmarks, funding appetite, margin pattern, cyclicality, demand pools, downside risks, and diligence questions. ${common}`,
    "industry-analysis": `Create a detailed NIC/basic-industry report. Cover value chain, buyer and supplier power, fragmentation, pricing power, working-capital cycle, capex intensity, compliance requirements, industry publications, association reports, government outlook, and how this industry affects the target company's score. ${common}`,
    "competitor-analysis": `Create a detailed competitor report. Identify closest public or private peers only when source evidence supports them. Compare positioning, scale indicators, business model, customer segments, pricing power, differentiation, substitution threat, and valuation or funding benchmarks when public. Judge business-model uniqueness from the live feed plus uploaded activity/NIC data. ${common}`,
    "director-profile": `Create a detailed director and promoter profile. Cover DIN/name evidence where available, appointment and directorship signals, LinkedIn/public profile, education, operating track record, disqualification checks, governance flags, related-party risk, and credibility impact on investment readiness. ${common}`,
    "company-analysis": `Create a full company investment report. Cover company snapshot, business model, capital structure, paid-up and authorized capital consistency, financial public data, filings, charges, score reconciliation, risks, positives, monitorables, and final Invest/Watchlist/Reject/Data Insufficient recommendation. ${common}`,
    "comprehensive-cdr": `Create one comprehensive CDR that combines Sector Analysis, Industry Analysis, Competitor Analysis, Director Profile, and Company Analysis. The document must be long-form, paragraph-heavy, investment-banking style, and detailed enough for board review. Use no markdown tables unless absolutely necessary for a scorecard; otherwise use paragraphs. ${common}`,
  };

  return instructions[type];
}

function searchQueries(company: Company, type: ReportType) {
  const identity = `${company.name} ${company.cin}`;
  const sector = `${company.sector} ${company.nicCode} ${company.activity}`;
  const common = [
    `${identity} MCA Zauba Tofler paid up capital authorized capital directors charges filings`,
    `${identity} news litigation funding contracts India`,
  ];

  const byType: Record<ReportType, string[]> = {
    "sector-analysis": [
      `${sector} India sector report IBEF FICCI government outlook PLI policy market size CAGR`,
      `${company.sector} India ministry policy annual report market outlook investment`,
    ],
    "industry-analysis": [
      `${sector} NIC industry analysis India value chain margin working capital association report`,
      `${company.activity} India industry report government outlook regulatory compliance`,
    ],
    "competitor-analysis": [
      `${identity} competitors peers India revenue funding Tofler Zauba`,
      `${company.activity} India private company competitors listed peers`,
    ],
    "director-profile": [
      `${identity} directors DIN LinkedIn promoter MCA disqualification`,
      `${company.director?.name || ""} ${company.name} director DIN directorships`,
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
    const results = await tavilySearchMany(searchQueries(company, selectedReportType));
    const prompt = `You are a senior investment analyst specializing in unlisted/private companies in India. Your task is to research the company named ${company.name} (CIN if known: ${company.cin}) and produce a thorough, data-driven Investment Research Report. You must NOT generate, hallucinate, or assume any data. Every single data point must be sourced from publicly available platforms listed below or from the uploaded MCA/feed-parser data. If data is unavailable, explicitly state "Data Not Available" for that field.

SOURCES TO SEARCH AND USE, IN ORDER:
1. MCA21 Portal (mca.gov.in) for incorporation data, directors, charges, filings
2. Tofler / Zaubacorp / Veratech for financials, balance sheet, P&L
3. RoC filings for annual returns and auditor reports
4. MOSPI / NIC Code Directory for sector classification
5. IBEF / FICCI / Industry Association Reports for sector outlook
6. News Search (ET, Business Standard, Mint and other credible business news) for recent developments
7. LinkedIn / Company Website for business model, team, products
8. CIBIL / Credit Rating disclosures if available
9. Court records / MCA charge search for litigation and charges

SOURCE OF TRUTH FROM UPLOADED EXCEL/MCA PARSER
${JSON.stringify(companyTruth(company, score), null, 2)}

STRICT CONSISTENCY RULE
The uploaded Excel/MCA parser values above are the source of truth for CIN, legal name, paid-up capital, authorized capital, sector/NIC/activity, filing date, uploaded director fields, factor scores, and deterministic Scout Score. Do not replace paid-up capital, authorized capital, CIN, or score with a public-source value. If Zauba, Tofler, MCA, or another source appears to conflict, state the discrepancy in a paragraph titled "Public-Source Discrepancy" and recommend manual verification.

OPTIONAL INTERNAL INFORMATION FROM USER
${internalInfo?.trim() || "NA"}

LIVE FEED
${sourceContext(results)}

REQUESTED REPORT TYPE
${selectedReportType}

REPORT INSTRUCTIONS
${reportInstructions(selectedReportType)}

BASE REPORT STRUCTURE TO GENERATE:

SECTION 1 - COMPANY SNAPSHOT
Full Legal Name; CIN Number; Date of Incorporation; Company Type; Registered Office Address; Company Status; Authorized Capital vs Paid-Up Capital; ROC Circle/Jurisdiction; AGM and Balance Sheet filing dates for last 3 years; Website/Contact if public.

SECTION 2 - BUSINESS MODEL & OPERATIONS
Products/services; revenue model; key customers/industries served if disclosed; geographies of operation; NIC Code and meaning; licenses/certifications/regulatory approvals; manufacturing/service delivery model.

SECTION 3 - FINANCIAL ANALYSIS (LAST 3-5 YEARS)
Pull only from Tofler / Zaubacorp / RoC filings if available. Include a table for FY21-FY24 covering Revenue from Operations, EBITDA/Operating Profit, Net Profit/Loss, Total Assets, Total Liabilities, Net Worth/Equity, Debt, Cash, Debt-to-Equity, ROE, Current Ratio. Include Revenue CAGR, profit trend, working capital, auditor qualifications, pending statutory dues. If unavailable, state "Public data not available - recommend requesting directly from company."

SECTION 4 - SECTOR & INDUSTRY ANALYSIS
Based on NIC Code: sector/sub-sector; market size; projected CAGR; tailwinds; headwinds; government policy/PLI/regulation; competitive intensity; listed peers; positioning.

SECTION 5 - DIRECTORS & MANAGEMENT
For each director if found: Full Name and DIN; designation; appointment date; other directorships; MCA disqualification status; education/professional background from LinkedIn if available; promoter/non-promoter classification; director changes in last 2 years.

SECTION 6 - SHAREHOLDING PATTERN
Promoter holding; institutional/PE/VC investors; ESOP pool; share pledges; recent shareholding changes; known investor exits. Use Data Not Available where not found.

SECTION 7 - CHARGES & BORROWINGS
Active charges; satisfied charges; total secured debt; lender quality; unusually large recent charges. Use MCA charge search/feed only.

SECTION 8 - LEGAL & COMPLIANCE RISK
NCLT/NCLAT/High Court cases; winding-up petitions; GST status; EPFO/ESIC if traceable; auditor quality; auditor changes; audit qualifications; ROC penalties/late filings. Bold every red flag.

SECTION 9 - RECENT NEWS & DEVELOPMENTS
Last 24 months: fundraising, contracts/tenders, regulatory action, raids, controversies, management exits, expansion, acquisitions/mergers.

SECTION 10 - INVESTMENT THESIS
Bull Case with 3-5 data-backed positives; Bear Case with 3-5 data-backed risks; Key Monitorables; Overall Investment Attractiveness table with Business Quality, Financial Health, Management Quality, Sector Tailwinds, Risk Level, Overall Investability rated 1-10 with justification; Suggested Due Diligence Checklist.

SECTION 11 - QUANTITATIVE SCOUT SMARTER SCORECARD
Explain the uploaded-file factor score, each factor, missing data penalties, and how it compares with AI qualitative view.

SECTION 12 - FINAL RECOMMENDATION
Invest / Watchlist / Reject / Data Insufficient with objective SEBI-style tone.

FINAL INSTRUCTIONS:
- Do not invent any number, name, director, DIN, date, filing, or fact.
- If a section has no data available from public sources, write: "Public data not available - recommend requesting directly from company".
- Cite the source platform for every major data point.
- Keep the tone objective, like a SEBI-registered research analyst.
- Flag every red flag in bold so it is immediately visible.
- Maintain uploaded paid-up capital and authorized capital exactly as shown in the source-of-truth block.
- Assess business-model uniqueness using uploaded NIC/activity data plus live public evidence; if public evidence is thin, say so and keep the conclusion conservative.
- The final report should be detailed enough that a reader can make an informed decision on whether to invest, lend to, or partner with this company without needing follow-up questions.`;

    const ai = await generateAiText({
      task: "cdr",
      system: "You are Scout Smarter, an investment banking analyst preparing a first-pass company diligence report. Be detailed, source-aware, skeptical, and never hallucinate.",
      prompt,
      temperature: 0.12,
      maxTokens: selectedReportType === "comprehensive-cdr" ? 4500 : 3200,
    });

    return NextResponse.json({
      report: ai.text || "No CDR report returned.",
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      provider: ai.provider,
      model: ai.model,
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "CDR generation failed." },
      { status: 500 },
    );
  }
}
