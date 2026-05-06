export const investmentBankingReportFormat = `
REPORT FORMAT STANDARD - MANDATORY FOR EVERY AI LAYER

Write like a senior investment-banking analyst preparing a board-ready diligence report. The output must be long-form, paragraph-heavy, and analytical. Do not write short field-value lines, bullet lists, numbered lists, markdown tables, checklists, or placeholder separators. Every section must contain complete paragraphs with context, interpretation, source confidence, and investment implication.

Use section headings only for major sections. Keep headings concise and professional. Under each heading, write multiple complete paragraphs. Each paragraph should usually be 80-160 words where data is available. If public data is unavailable, do not stop at "Data Not Available"; explain what is unavailable, why that matters to underwriting, and what diligence document should be requested.

Every AI layer must include the same analytical depth: company identity and source-of-truth reconciliation, business model interpretation, sector and industry context, government stance and thematic outlook where relevant, competitive positioning, director/promoter credibility, capital structure and score reconciliation, red flags, positives, diligence gaps, and final investment implication.

Use uploaded/parser data as the source of truth for CIN, company name, paid-up capital, authorized capital, sector/NIC/activity, uploaded filing date, deterministic factor scores, and deterministic Scout Score. If a public source conflicts with uploaded data, write a paragraph titled "Public-Source Discrepancy" and explain the discrepancy without replacing uploaded values.

Red flags must begin with "RED FLAG:" and be written as full paragraphs. Positives must begin with "POSITIVE:" and be written as full paragraphs. Do not use asterisks for emphasis. Do not emit raw JSON, markdown code fences, or markdown table pipes unless the endpoint explicitly asks for JSON.

The final report should read as a complete investment memo, not a data extraction sheet. It should help a reader decide whether to invest, reject, watchlist, or request more diligence without needing immediate follow-up questions.
`.trim();

export const scoringJsonReportFormat = `
SCORING AI FORMAT STANDARD - MANDATORY

Return valid JSON only, but make the text values inside the JSON detailed and paragraph-heavy. The "report" value must be a complete portfolio-level investment-banking memo of at least four substantial paragraphs. Each insight "rationale" must be a detailed paragraph that explains parsed score, Tavily/public-source confidence, sector attractiveness, business-model uniqueness, director quality, capital structure, missing data, and investment implication.

The "strengths", "redFlags", and "missingData" arrays must contain complete sentence strings, not short fragments. Red flag strings must start with "RED FLAG:" and positive strings should start with "POSITIVE:" where appropriate. Do not return markdown bullets inside the JSON strings.
`.trim();
