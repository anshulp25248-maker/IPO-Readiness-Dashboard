import { execFile } from "node:child_process";
import { mkdir, unlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { NextResponse } from "next/server";

export const runtime = "nodejs";

const execFileAsync = promisify(execFile);

export async function POST(request: Request) {
  let savedPath = "";

  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "No file uploaded." }, { status: 400 });
    }

    const uploadDir = path.join(os.tmpdir(), "scout-smarter-uploads");
    await mkdir(uploadDir, { recursive: true });

    const safeName = file.name.replace(/[^a-z0-9._-]/gi, "_") || "company-file";
    savedPath = path.join(uploadDir, `${Date.now()}-${safeName}`);
    await writeFile(savedPath, Buffer.from(await file.arrayBuffer()));

    const scriptPath = path.join(process.cwd(), "scripts", "parse_company_file.py");
    const { stdout } = await execFileAsync("python", [scriptPath, savedPath], {
      cwd: process.cwd(),
      maxBuffer: 12 * 1024 * 1024,
      timeout: 120000,
    });

    const parsed = JSON.parse(stdout) as {
      rows?: Array<Record<string, string>>;
      rowCount?: number;
    };

    return NextResponse.json({
      rows: parsed.rows ?? [],
      rowCount: parsed.rowCount ?? parsed.rows?.length ?? 0,
      fileName: file.name,
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "File parsing failed." },
      { status: 500 },
    );
  } finally {
    if (savedPath) {
      await unlink(savedPath).catch(() => undefined);
    }
  }
}
