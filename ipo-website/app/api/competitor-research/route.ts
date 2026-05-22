import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { generateAiText } from "../_lib/ai";
import { dashboardLaneEnvValue, sourceContext, sourceLinks, tavilySearchLane } from "../_lib/dashboard-research-lanes";
import { investmentBankingReportFormat } from "../_lib/report-format";

export const runtime = "nodejs";

function cleanProviderDetails(message: string) {
  return message
    .replace(/\bGroq\b/gi, "analysis service")
    .replace(/\bTavily\b/gi, "live feed")
    .replace(/\bGemini\b/gi, "analysis service")
    .replace(/\bOpenRouter\b/gi, "analysis service")
    .replace(/No AI provider configured\.[^|]*/gi, "Analysis service is not configured.");
}

export async function POST(request: Request) {
  try {
    const { company, peers } = (await request.json()) as { company?: Company; peers?: Company[] };
    if (!company) return NextResponse.json({ error: "No company supplied." }, { status: 400 });

    const similarPeers = (peers || [])
      .filter((peer) => peer.id !== company.id)
      .filter((peer) => peer.nicCode === company.nicCode || peer.activity === company.activity || peer.sector === company.sector)
      .slice(0, 25);

    const results = await tavilySearchLane("competitor-analysis", [
      `${company.name} ${company.activity} ${company.nicCode} competitors peers India private limited`,
      `${company.name} similar companies India competitors Zauba Tofler LinkedIn`,
      `${company.activity} ${company.sector} India listed unlisted competitors market share`,
    ]);

    const prompt = `Build a competitor and business-model diligence report from uploaded/parser data, the supplied peer universe, and the dedicated Competitor Analysis live-feed lane. Do not invent facts.

TARGET COMPANY
${JSON.stringify(company, null, 2)}

SIMILAR COMPANIES FROM UPLOADED LIST
${JSON.stringify(similarPeers, null, 2)}

COMPETITOR LIVE FEED
${sourceContext(results)}

${investmentBankingReportFormat}

Write the competitor report with these major section headings and detailed paragraphs under each: Business Model Interpretation; Uploaded Peer Universe; Public Competitor Landscape; Scarcity and Saturation Analysis; Differentiation and Moat; Red Flags and Commoditization Risk; Score Implication; Diligence Questions and Investment View.`;

    const ai = await generateAiText({
      task: "competitor",
      provider: "groq",
      apiKey: dashboardLaneEnvValue("competitor-analysis", "GROQ_API_KEY"),
      providerLabel: "Competitor Analysis",
      system: "You are a competition and business-model diligence analyst for private-company investing.",
      prompt,
      temperature: 0.12,
      maxTokens: 3000,
    });

    return NextResponse.json({
      report: ai.text || "No competitor research returned.",
      sources: sourceLinks(results),
      lane: "Competitor Analysis",
      peerCount: similarPeers.length,
      provider: ai.provider,
      model: ai.model,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? cleanProviderDetails(error.message) : "Competitor research failed." }, { status: 500 });
  }
}
