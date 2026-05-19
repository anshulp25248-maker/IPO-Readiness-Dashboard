"use client";

function cleanLine(line: string) {
  return line
    .replace(/^#+\s*/, "")
    .replace(/^[-*]\s*/, "")
    .replace(/^\d+\.\s*/, "")
    .replace(/\*\*/g, "")
    .trim();
}

function mergeParagraphLines(lines: string[]) {
  const merged: string[] = [];

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;
    const cleaned = cleanLine(line);
    const previous = merged[merged.length - 1];
    const startsNew =
      /^#{1,3}\s*/.test(line) ||
      /^(section\s+)?\d+\s*[.-]/i.test(line) ||
      /^\*\*[^*\n]{2,90}\*\*\s*$/.test(line) ||
      isSubheading(line) ||
      /^RED FLAG:/i.test(cleaned) ||
      /^POSITIVE:/i.test(cleaned);

    if (
      previous &&
      !startsNew &&
      previous.length < 360 &&
      !/[.!?)]$/.test(previous) &&
      !/:\s*$/.test(previous)
    ) {
      merged[merged.length - 1] = `${previous} ${cleaned}`;
    } else {
      merged.push(line);
    }
  }

  return merged;
}

function isSubheading(line: string) {
  const cleaned = cleanLine(line);
  return (
    /:$/.test(cleaned) ||
    /^(bull case|bear case|key monitorables|overall investment attractiveness|suggested due diligence|source|sources|recommendation|data not available|red flags?)/i.test(cleaned)
  );
}

export function ReportViewer({ report }: { report: string }) {
  const sections = report
    .split(/\n(?=(?:#{1,3}\s*)?(?:SECTION\s+)?\d+\s*[.-]|#{1,3}\s+|\*\*[^*\n]{2,90}\*\*\s*$)/im)
    .map((section) => section.trim())
    .filter(Boolean);

  if (!sections.length) {
    return null;
  }

  return (
    <div className="grid gap-6">
      {sections.map((section, index) => {
        const lines = mergeParagraphLines(section.split(/\r?\n/));
        const title = cleanLine(lines[0] || `Section ${index + 1}`);
        const body = lines.slice(1);

        return (
          <section
            key={`${title}-${index}`}
            className="rounded-xl border border-white/45 bg-white/65 p-4 shadow-[0_12px_32px_rgba(15,23,42,0.1)] backdrop-blur-xl"
          >
            <div className="flex items-start gap-3 border-b border-slate-950/10 pb-3">
              <span className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-slate-950 text-xs font-black text-white">
                {index + 1}
              </span>
              <h3 className="text-base font-black leading-6 text-slate-950">
                {title}
              </h3>
            </div>

            <div className="mt-4 space-y-3 text-sm font-medium leading-6 text-slate-900">
              {body.map((line, lineIndex) => {
                const cleaned = cleanLine(line);
                if (!cleaned) return null;

                const isRedFlag = /\bred flag\b|risk|unavailable|negative|caution|default|litigation|penalty|discrepancy/i.test(cleaned);
                const isPositive = /\bpositive\b|bull case|strength|tailwind|opportunity|attractive|advantage/i.test(cleaned);

                if (isSubheading(line)) {
                  return (
                    <h4
                      key={`${cleaned}-${lineIndex}`}
                      className="pt-1 text-sm font-black text-blue-950"
                    >
                      {cleaned}
                    </h4>
                  );
                }

                return (
                  <p
                    key={`${cleaned}-${lineIndex}`}
                    className={`rounded-lg px-3 py-2 text-justify leading-6 ${
                      isRedFlag
                        ? "border border-rose-200 bg-rose-50/85 font-semibold text-rose-950"
                        : isPositive
                          ? "border border-emerald-200 bg-emerald-50/80 font-semibold text-emerald-950"
                          : "bg-white/25 text-slate-950"
                    }`}
                  >
                    {cleaned}
                  </p>
                );
              })}
            </div>
          </section>
        );
      })}
    </div>
  );
}
