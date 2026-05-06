import { NextResponse } from "next/server";
import * as XLSX from "xlsx";

export const runtime = "nodejs";

function cleanCell(value: unknown) {
  return String(value ?? "").trim().slice(0, 240);
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
    return headers.reduce<Record<string, string>>((row, header, index) => {
      row[header] = cleanCell(cells[index]);
      row._source = delimiter === "\t" ? "tsv" : "csv";
      return row;
    }, {});
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
    return Object.entries(row as Record<string, unknown>).reduce<Record<string, string>>((record, [key, value]) => {
      record[key] = cleanCell(value);
      record._source = `json_${index + 1}`;
      return record;
    }, {});
  });
}

function parseWorkbook(bytes: Buffer) {
  const workbook = XLSX.read(bytes, { type: "buffer", cellDates: false, raw: false });
  const rows: Array<Record<string, string>> = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const records = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: false });
    records.forEach((record) => {
      const cleaned = Object.entries(record).reduce<Record<string, string>>((row, [key, value]) => {
        row[key] = cleanCell(value);
        return row;
      }, {});
      cleaned._source = `sheet:${sheetName}`;
      rows.push(cleaned);
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
