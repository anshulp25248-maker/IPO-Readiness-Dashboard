import { NextResponse } from "next/server";
import { Company } from "../../_data/companies";
import { envValue, generateAiText } from "../_lib/ai";
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
    const { company, cin } = (await request.json()) as { company?: Company; cin?: string };
    const target = company?.cin || cin || company?.name || "";
    if (!target) return NextResponse.json({ error: "Provide a company or CIN." }, { status: 400 });

    const prompt = `Prepare director diligence for this company/CIN from uploaded/parser data and the supplied company fields. Do not hallucinate. Mark unavailable public facts as Data Not Available.

COMPANY DATA
${JSON.stringify(company || { cin }, null, 2)}

${investmentBankingReportFormat}

Write the director report with these major section headings and detailed paragraphs under each: Director Identity and Source Confidence; DIN and Directorship Signals; Education and Operating Track Record; Promoter Credibility; Governance and Related-Party Concerns; Red Flags and Negative News; Verification Gaps; Director Score Recommendation; Investment View.`;

    const ai = await generateAiText({
      task: "director",
      provider: "groq",
      apiKey: envValue("CDR_DIRECTOR_PROFILE_GROQ_API_KEY") || envValue("GROQ_API_KEY_CDR_DIRECTOR_PROFILE") || envValue("GROQ_API_KEY"),
      providerLabel: "Director Profile",
      system: "You are a cautious investment-banking KYC and promoter diligence analyst. Separate facts from inference.",
      prompt,
      temperature: 0.1,
      maxTokens: 3000,
    });

    return NextResponse.json({
      report: ai.text || "No director research returned.",
      sources: [],
      provider: ai.provider,
      model: ai.model,
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? cleanProviderDetails(error.message) : "Director research failed." }, { status: 500 });
  }
}
