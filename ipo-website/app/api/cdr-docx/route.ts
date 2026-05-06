import {
  AlignmentType,
  Document as DocxDocument,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
} from "docx";
import { NextResponse } from "next/server";

export const runtime = "nodejs";

function safeFileName(value: string) {
  return value.replace(/[^a-z0-9._-]+/gi, "_").replace(/^_+|_+$/g, "") || "cdr-report";
}

function cleanLine(value: string) {
  return value
    .replace(/^#{1,6}\s*/, "")
    .replace(/^[-*]\s*/, "")
    .replace(/^\d+[\).]\s*/, "")
    .replace(/\*\*/g, "")
    .trim();
}

function isHeading(rawLine: string) {
  const line = cleanLine(rawLine);
  if (!line) return false;
  return (
    /^#{1,6}\s*/.test(rawLine) ||
    /^(section\s+)?\d+\s*[-.:]/i.test(rawLine.trim()) ||
    (line.endsWith(":") && line.length <= 90) ||
    (line === line.toUpperCase() && line.split(/\s+/).length <= 12)
  );
}

function textColor(line: string) {
  const lower = line.toLowerCase();
  if (/\bred flag\b|risk|litigation|penalty|default|negative|caution|discrepancy/.test(lower)) return "BE123C";
  if (/positive|strength|tailwind|opportunity|advantage|bull case/.test(lower)) return "065F46";
  return "111827";
}

function headingParagraph(text: string, level: typeof HeadingLevel.HEADING_1 | typeof HeadingLevel.HEADING_2) {
  return new Paragraph({
    heading: level,
    spacing: { before: 260, after: 160 },
    children: [
      new TextRun({
        text: cleanLine(text).toUpperCase(),
        bold: true,
        color: "091F4A",
        font: "Arial",
        size: level === HeadingLevel.HEADING_1 ? 30 : 25,
      }),
    ],
  });
}

function bodyParagraph(text: string) {
  const cleaned = cleanLine(text);
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 140, line: 276 },
    children: [
      new TextRun({
        text: cleaned,
        color: textColor(cleaned),
        font: "Arial",
        size: 21,
        bold: textColor(cleaned) !== "111827",
      }),
    ],
  });
}

function reportParagraphs(report: string) {
  return report
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => (isHeading(line) ? headingParagraph(line, HeadingLevel.HEADING_1) : bodyParagraph(line)));
}

export async function POST(request: Request) {
  try {
    const payload = await request.json();
    const company = payload?.company || {};
    const companyName = String(company?.name || "company");
    const score = payload?.score ?? "NA";
    const sources = Array.isArray(payload?.sources) ? payload.sources : [];

    const doc = new DocxDocument({
      sections: [
        {
          children: [
            new Paragraph({
              heading: HeadingLevel.TITLE,
              spacing: { after: 240 },
              children: [
                new TextRun({
                  text: "Scout Smarter Company Diligence Report",
                  bold: true,
                  color: "091F4A",
                  font: "Arial",
                  size: 38,
                }),
              ],
            }),
            bodyParagraph(`Company: ${company.name || "NA"}`),
            bodyParagraph(`CIN: ${company.cin || "NA"}`),
            bodyParagraph(`Score: ${score}/100`),
            bodyParagraph(`Paid-up Capital: ${company.paidUpCapital || "NA"}`),
            bodyParagraph(`Authorized Capital: ${company.authorizedCapital || "NA"}`),
            bodyParagraph(`Sector: ${company.sector || "NA"}`),
            bodyParagraph(`Location: ${company.city || "NA"}, ${company.state || "NA"}`),
            ...reportParagraphs(String(payload?.report || "")),
            ...(sources.length ? [headingParagraph("Source Feed", HeadingLevel.HEADING_1)] : []),
            ...sources.map((source: { title?: string; url?: string }) =>
              bodyParagraph(`${source.title || "Source"}: ${source.url || "NA"}`),
            ),
          ],
        },
      ],
    });

    const bytes = await Packer.toBuffer(doc);
    return new NextResponse(new Uint8Array(bytes), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${safeFileName(companyName)}-CDR.docx"`,
      },
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "DOCX generation failed." },
      { status: 500 },
    );
  }
}
