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
    const groqKey = envValue("GROQ_API_KEY");
    if (!groqKey) return NextResponse.json({ error: "GROQ_API_KEY is missing in .env." }, { status: 500 });

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

    const response = await fetch(groqUrl, {
      method: "POST",
      headers: { Authorization: `Bearer ${groqKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: envValue("GROQ_MODEL") || "llama-3.3-70b-versatile",
        messages: [
          { role: "system", content: "You are a competition and business-model diligence analyst for private-company investing." },
          { role: "user", content: prompt },
        ],
        temperature: 0.12,
        max_tokens: 5000,
      }),
      cache: "no-store",
    });

    if (!response.ok) throw new Error(`Groq competitor research failed with ${response.status}`);
    const data = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> };
    return NextResponse.json({
      report: data.choices?.[0]?.message?.content || "No competitor research returned.",
      sources: results.map((item) => ({ title: item.title, url: item.url })),
      peerCount: similarPeers.length,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Competitor research failed." }, { status: 500 });
  }
}
