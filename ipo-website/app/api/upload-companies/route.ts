import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

function cleanCell(value: unknown) {
  return String(value ?? "").trim().slice(0, 240);
}

function normalizeHeader(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function uniqueHeaders(row: unknown[], rowIndex: number) {
  const seen = new Map<string, number>();
  return row.map((cell, cellIndex) => {
    const raw = cleanCell(cell) || `column_${cellIndex + 1}`;
    const normalized = normalizeHeader(raw) || `column_${cellIndex + 1}`;
    const count = seen.get(normalized) ?? 0;
    seen.set(normalized, count + 1);
    return count ? `${raw}_${count + 1}` : raw || `column_${rowIndex + 1}_${cellIndex + 1}`;
  });
}

function headerScore(row: unknown[]) {
  const normalized = row.map((cell) => normalizeHeader(cleanCell(cell))).filter(Boolean);
  const text = normalized.join(" ");
  let score = 0;
  if (/\b(cin|llpin|fcin)\b/.test(text)) score += 3;
  if (/company_name|llp_name|limited_liability_partnership_name|companyname/.test(text)) score += 3;
  if (/paidup_capital|paid_capital|paid_up_capital|obligation_of_contribution|total_obligation/.test(text)) score += 2;
  if (/authorized_capital|authorised_capital|activity_description|industrial_activity|activity_code|state|roc/.test(text)) score += 1;
  return score;
}

function looksLikeHeader(row: unknown[]) {
  return headerScore(row) >= 4;
}

function getNormalizedField(record: Record<string, string>, names: string[]) {
  const normalizedRecord = Object.entries(record).reduce<Record<string, string>>((current, [key, value]) => {
    current[normalizeHeader(key)] = value;
    return current;
  }, {});

  for (const name of names) {
    const value = normalizedRecord[normalizeHeader(name)];
    if (value) return value;
  }
  return "";
}

function enrichCanonicalFields(record: Record<string, string>): Record<string, string> {
  const companyName = getNormalizedField(record, [
    "company_name",
    "company name",
    "llp_name",
    "llp name",
    "limited liability partnership name",
    "limited_liability_partnership_name",
    "name",
  ]);
  const identifier = getNormalizedField(record, ["cin", "llpin", "fcin"]);
  const paidUpCapital = getNormalizedField(record, [
    "paidup capital",
    "paidup_capital",
    "paid capital",
    "paid_capital",
    "paid up capital",
    "paid_up_capital",
    "paid up share capital",
    "obligation of contribution(rs.)",
    "obligation_of_contribution_rs",
    "total obligation of contribution",
    "total_obligation_of_contribution",
  ]);
  const authorizedCapital = getNormalizedField(record, [
    "authorized capital",
    "authorized_capital",
    "authorised capital",
    "authorised_capital",
    "authorised share capital",
  ]);
  const state = getNormalizedField(record, ["state"]);
  const city = getNormalizedField(record, ["district", "city"]);
  const activity = getNormalizedField(record, [
    "activity description",
    "activity_description",
    "description",
    "industrial activity",
    "industrial_activity",
    "business activity",
  ]);
  const activityCode = getNormalizedField(record, ["activity code", "activity_code", "industrial activity", "industrial_activity", "nic code", "nic_code"]);
  const status = getNormalizedField(record, ["company status", "company_status", "status"]);
  const incorporationDate = getNormalizedField(record, ["date of registration", "date_of_registration", "date of incorporation", "date_of_incorporation", "founded", "date"]);

  return {
    ...record,
    ...(companyName ? { company_name: companyName } : {}),
    ...(identifier ? { cin: identifier } : {}),
    ...(paidUpCapital ? { paid_up_capital: paidUpCapital } : {}),
    ...(authorizedCapital ? { authorized_capital: authorizedCapital } : {}),
    ...(state ? { state } : {}),
    ...(city ? { city } : {}),
    ...(activity ? { activity, activity_description: activity, sector: activity } : {}),
    ...(activityCode ? { nic_code: activityCode } : {}),
    ...(status ? { company_status: status } : {}),
    ...(incorporationDate ? { incorporation_date: incorporationDate, last_filing_date: incorporationDate } : {}),
  };
}

function splitDelimitedLine(line: string, delimiter: "," | "\t") {
  const cells: string[] = [];
  let cell = "";
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const next = line[index + 1];
    if (char === '"' && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if (char === delimiter && !insideQuotes) {
      cells.push(cell.trim());
      cell = "";
    } else {
      cell += char;
    }
  }

  cells.push(cell.trim());
  return cells;
}

function parseDelimited(text: string, delimiter: "," | "\t") {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (lines.length < 2) return [];
  const headers = splitDelimitedLine(lines[0], delimiter).map((header) => header.trim() || "column");
  return lines.slice(1).map((line) => {
    const cells = splitDelimitedLine(line, delimiter);
    const row = headers.reduce<Record<string, string>>((current, header, index) => {
      current[header] = cleanCell(cells[index]);
      current._source = delimiter === "\t" ? "tsv" : "csv";
      return current;
    }, {});
    return enrichCanonicalFields(row);
  });
}

function parseJson(text: string) {
  const parsed = JSON.parse(text) as unknown;
  const rows = Array.isArray(parsed)
    ? parsed
    : Array.isArray((parsed as { rows?: unknown[] })?.rows)
      ? (parsed as { rows: unknown[] }).rows
      : [];

  return rows.map((row, index) => {
    if (!row || typeof row !== "object") {
      return { value: cleanCell(row), _source: "json" };
    }
    const record = Object.entries(row as Record<string, unknown>).reduce<Record<string, string>>((current, [key, value]) => {
      current[key] = cleanCell(value);
      current._source = `json_${index + 1}`;
      return current;
    }, {});
    return enrichCanonicalFields(record);
  });
}

function parseWorkbook(bytes: Buffer) {
  const workbook = XLSX.read(bytes, { type: "buffer", cellDates: false, raw: false });
  const rows: Array<Record<string, string>> = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });
    let headers: string[] = [];
    let section = "";

    matrix.forEach((row, rowIndex) => {
      const cleanedCells = row.map(cleanCell);
      const nonEmpty = cleanedCells.filter(Boolean);
      if (!nonEmpty.length) return;

      if (nonEmpty.length === 1 && !looksLikeHeader(cleanedCells)) {
        section = nonEmpty[0];
        return;
      }

      if (looksLikeHeader(cleanedCells)) {
        headers = uniqueHeaders(cleanedCells, rowIndex);
        return;
      }

      if (!headers.length) return;

      const record = headers.reduce<Record<string, string>>((current, header, index) => {
        current[header] = cleanCell(cleanedCells[index]);
        return current;
      }, {});
      const hasIdentity = getNormalizedField(record, ["cin", "llpin", "fcin", "company name", "llp name", "limited liability partnership name"]);
      if (!hasIdentity) return;

      const enriched = enrichCanonicalFields(record);
      enriched._source = section ? `sheet:${sheetName}:${section}` : `sheet:${sheetName}`;
      rows.push(enriched);
    });
  });

  return rows;
}

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "No file uploaded." }, { status: 400 });
    }

    const bytes = Buffer.from(await file.arrayBuffer());
    const extension = file.name.toLowerCase().split(".").pop() || "";
    const text = bytes.toString("utf8");
    let rows: Array<Record<string, string>> = [];

    if (["xlsx", "xlsm", "xls", "ods"].includes(extension)) {
      rows = parseWorkbook(bytes);
    } else if (extension === "csv") {
      rows = parseDelimited(text, ",");
    } else if (["tsv", "tab"].includes(extension)) {
      rows = parseDelimited(text, "\t");
    } else if (extension === "json") {
      rows = parseJson(text);
    } else {
      return NextResponse.json(
        { error: "Unsupported live upload type. Please use Excel, CSV, TSV, or JSON on the web app." },
        { status: 400 },
      );
    }

    return NextResponse.json({
      rows: rows.slice(0, 15000),
      rowCount: rows.length,
      fileName: file.name,
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "File parsing failed." },
      { status: 500 },
    );
  }
}
