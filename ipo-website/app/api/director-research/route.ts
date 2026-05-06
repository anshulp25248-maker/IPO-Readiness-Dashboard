import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";
import { investmentBankingReportFormat } from "../_lib/report-format";

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
  if (!response.ok) throw new Error(`Tavily director search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

function sourceContext(results: SearchResult[]) {
  return results
    .map((item, index) => `${index + 1}. ${item.title || "Untitled"}\nURL: ${item.url || "NA"}\nSnippet: ${(item.content || "").slice(0, 750)}`)
    .join("\n\n") || "No public feed results returned.";
}

export async function POST(request: Request) {
  try {
    const { company, cin } = (await request.json()) as { company?: Company; cin?: string };
    const target = company?.cin || cin || company?.name || "";
    if (!target) return NextResponse.json({ error: "Provide a company or CIN." }, { status: 400 });

    const query = `${company?.name || ""} ${target} directors DIN MCA Zauba Tofler LinkedIn founder promoter directorships education`;
    const results = await tavilySearch(query);
    const prompt = `Prepare director diligence for this company/CIN. Use public feed results only. You may reference MCA, Zauba, Tofler, LinkedIn public snippets, news, and official pages if returned. Do not scrape private LinkedIn content. Do not hallucinate.

COMPANY DATA
${JSON.stringify(company || { cin }, null, 2)}

PUBLIC FEED
${sourceContext(results)}

${investmentBankingReportFormat}

Write the director report with these major section headings and detailed paragraphs under each: Director Identity and Source Confidence; DIN and Directorship Signals; Education and Operating Track Record; Promoter Credibility; Governance and Related-Party Concerns; Red Flags and Negative News; Verification Gaps; Director Score Recommendation; Investment View.`;

    const ai = await generateAiText({
      task: "director",
      system: "You are a cautious investment-banking KYC and promoter diligence analyst. Separate facts from inference.",
      prompt,
      temperature: 0.1,
      maxTokens: 3000,
    });

    return NextResponse.json({
      report: ai.text || "No director research returned.",
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      provider: ai.provider,
      model: ai.model,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Director research failed." }, { status: 500 });
  }
}
