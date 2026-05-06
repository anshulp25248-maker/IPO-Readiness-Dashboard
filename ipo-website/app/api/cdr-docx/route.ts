import { execFile } from "node:child_process";
import { mkdir, readFile, unlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { NextResponse } from "next/server";

export const runtime = "nodejs";

const execFileAsync = promisify(execFile);

function safeFileName(value: string) {
  return value.replace(/[^a-z0-9._-]+/gi, "_").replace(/^_+|_+$/g, "") || "cdr-report";
}

export async function POST(request: Request) {
  let inputPath = "";
  let outputPath = "";

  try {
    const payload = await request.json();
    const companyName = String(payload?.company?.name || "company");
    const tempDir = path.join(os.tmpdir(), "scout-smarter-cdr");
    await mkdir(tempDir, { recursive: true });

    const stamp = Date.now();
    inputPath = path.join(tempDir, `${stamp}.json`);
    outputPath = path.join(tempDir, `${stamp}-${safeFileName(companyName)}.docx`);
    await writeFile(inputPath, JSON.stringify(payload), "utf8");

    const scriptPath = path.join(process.cwd(), "scripts", "generate_cdr_docx.py");
    await execFileAsync("python", [scriptPath, inputPath, outputPath], {
      cwd: process.cwd(),
      timeout: 120000,
      maxBuffer: 4 * 1024 * 1024,
    });

    const bytes = await readFile(outputPath);
    return new NextResponse(bytes, {
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
  } finally {
    if (inputPath) await unlink(inputPath).catch(() => undefined);
    if (outputPath) await unlink(outputPath).catch(() => undefined);
  }
}
