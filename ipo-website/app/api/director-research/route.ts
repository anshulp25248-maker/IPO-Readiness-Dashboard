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
    const groqKey = envValue("GROQ_API_KEY");
    if (!groqKey) return NextResponse.json({ error: "GROQ_API_KEY is missing in .env." }, { status: 500 });

    const target = company?.cin || cin || company?.name || "";
    if (!target) return NextResponse.json({ error: "Provide a company or CIN." }, { status: 400 });

    const query = `${company?.name || ""} ${target} directors DIN MCA Zauba Tofler LinkedIn founder promoter directorships education`;
    const results = await tavilySearch(query);
    const prompt = `Prepare director diligence for this company/CIN. Use public feed results only. You may reference MCA, Zauba, Tofler, LinkedIn public snippets, news, and official pages if returned. Do not scrape private LinkedIn content. Do not hallucinate.

COMPANY DATA
${JSON.stringify(company || { cin }, null, 2)}

PUBLIC FEED
${sourceContext(results)}

Return a detailed markdown report with:
1. Director / promoter names found
2. DIN or identifiers if found
3. Current and past directorship signals
4. Education / professional qualification signals
5. Founder/operator credibility
6. Red flags and negative news
7. Verification gaps
8. Director score recommendation using rules: 3+ directorships = 10, 2 = 7, 1 = 4, professional degree only = 5, professional degree + multiple directorships = 10, insufficient data = Unverified.
9. Sources used`;

    const response = await fetch(groqUrl, {
      method: "POST",
      headers: { Authorization: `Bearer ${groqKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: envValue("GROQ_MODEL") || "llama-3.3-70b-versatile",
        messages: [
          { role: "system", content: "You are a cautious investment-banking KYC and promoter diligence analyst. Separate facts from inference." },
          { role: "user", content: prompt },
        ],
        temperature: 0.1,
        max_tokens: 5000,
      }),
      cache: "no-store",
    });

    if (!response.ok) throw new Error(`Groq director research failed with ${response.status}`);
    const data = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> };
    return NextResponse.json({
      report: data.choices?.[0]?.message?.content || "No director research returned.",
      sources: results.map((item) => ({ title: item.title, url: item.url })),
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Director research failed." }, { status: 500 });
  }
}
