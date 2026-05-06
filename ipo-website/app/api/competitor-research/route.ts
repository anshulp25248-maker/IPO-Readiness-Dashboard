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
  if (!response.ok) throw new Error(`Tavily competitor search failed with ${response.status}`);
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
    const { company, peers } = (await request.json()) as { company?: Company; peers?: Company[] };
    if (!company) return NextResponse.json({ error: "No company supplied." }, { status: 400 });

    const similarPeers = (peers || [])
      .filter((peer) => peer.id !== company.id)
      .filter((peer) => peer.nicCode === company.nicCode || peer.activity === company.activity || peer.sector === company.sector)
      .slice(0, 25);
    const query = `${company.name} ${company.activity} ${company.nicCode} ${company.sector} competitors similar companies India private limited`;
    const results = await tavilySearch(query);

    const prompt = `Build a competitor and business-model diligence report. Use uploaded peers first, then live feed. Do not invent facts.

TARGET COMPANY
${JSON.stringify(company, null, 2)}

SIMILAR COMPANIES FROM UPLOADED LIST
${JSON.stringify(similarPeers, null, 2)}

LIVE FEED
${sourceContext(results)}

Return markdown with:
1. Business model interpretation
2. Same NIC/activity peers from uploaded list
3. Public competitors from live feed
4. Scarcity/saturation analysis
5. Differentiation hypotheses
6. Red flags and commoditization risk
7. Competitor score implication for investment screening
8. Suggested diligence questions`;

    const ai = await generateAiText({
      system: "You are a competition and business-model diligence analyst for private-company investing.",
      prompt,
      temperature: 0.12,
      maxTokens: 3000,
    });

    return NextResponse.json({
      report: ai.text || "No competitor research returned.",
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      peerCount: similarPeers.length,
      provider: ai.provider,
      model: ai.model,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Competitor research failed." }, { status: 500 });
  }
}
