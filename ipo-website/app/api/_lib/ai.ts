import fs from "node:fs";
import path from "node:path";

type GenerateOptions = {
  task?: "scoring" | "company-search" | "cdr" | "competitor" | "director";
  system?: string;
  prompt: string;
  temperature?: number;
  maxTokens?: number;
  responseJson?: boolean;
};

export type AiResult = {
  text: string;
  provider: string;
  model: string;
};

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

export function envValue(key: string) {
  if (process.env[key]) return process.env[key] || "";
  for (const candidate of [
    path.join(process.cwd(), ".env.local"),
    path.join(process.cwd(), ".env"),
    path.join(process.cwd(), "..", ".env"),
  ]) {
    const value = readEnvFile(candidate)[key];
    if (value) return value;
  }
  return "";
}

const taskProviderOrder: Record<NonNullable<GenerateOptions["task"]>, string[]> = {
  scoring: ["gemini", "openrouter", "groq"],
  "company-search": ["openrouter", "gemini", "groq"],
  cdr: ["gemini", "openrouter", "groq"],
  competitor: ["groq", "openrouter", "gemini"],
  director: ["openrouter", "gemini", "groq"],
};

function envTaskKey(base: string, task?: GenerateOptions["task"]) {
  return task ? `${base}_${task.replace(/-/g, "_").toUpperCase()}` : base;
}

function providerOrder(task?: GenerateOptions["task"]) {
  const taskSpecific = envValue(envTaskKey("AI_PROVIDER_ORDER", task))
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);
  if (taskSpecific.length) return taskSpecific;

  const requested = envValue("AI_PROVIDER_ORDER")
    .split(",")
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);
  if (requested.length) return requested;

  return task ? taskProviderOrder[task] : ["gemini", "openrouter", "groq"];
}

function taskModel(base: string, fallback: string, task?: GenerateOptions["task"]) {
  return envValue(envTaskKey(base, task)) || envValue(base) || fallback;
}

function compactError(provider: string, response: Response, body: string) {
  const snippet = body.replace(/\s+/g, " ").slice(0, 240);
  return `${provider} failed with ${response.status}${snippet ? `: ${snippet}` : ""}`;
}

async function callGemini(options: GenerateOptions): Promise<AiResult> {
  const apiKey = envValue("GEMINI_API_KEY") || envValue("GOOGLE_API_KEY");
  if (!apiKey) throw new Error("Gemini key not configured");

  const model = taskModel("GEMINI_MODEL", "gemini-2.5-flash", options.task);
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [
          {
            role: "user",
            parts: [{ text: `${options.system ? `${options.system}\n\n` : ""}${options.prompt}` }],
          },
        ],
        generationConfig: {
          temperature: options.temperature ?? 0.1,
          maxOutputTokens: options.maxTokens ?? 2200,
          responseMimeType: options.responseJson ? "application/json" : "text/plain",
        },
      }),
      cache: "no-store",
    },
  );

  const data = await response.json().catch(async () => ({ raw: await response.text() }));
  if (!response.ok) throw new Error(compactError("Gemini", response, JSON.stringify(data)));

  const text = data?.candidates?.[0]?.content?.parts?.map((part: { text?: string }) => part.text || "").join("") || "";
  return { text, provider: "Gemini", model };
}

async function callGroq(options: GenerateOptions): Promise<AiResult> {
  const apiKey = envValue("GROQ_API_KEY");
  if (!apiKey) throw new Error("Groq key not configured");

  const preferred = taskModel("GROQ_MODEL", "llama-3.1-8b-instant", options.task);
  const models = preferred === "llama-3.1-8b-instant" ? [preferred] : [preferred, "llama-3.1-8b-instant"];
  let lastError = "";

  for (const model of models) {
    const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model,
        messages: [
          ...(options.system ? [{ role: "system", content: options.system }] : []),
          { role: "user", content: options.prompt },
        ],
        temperature: options.temperature ?? 0.1,
        max_tokens: options.maxTokens ?? 2200,
        response_format: options.responseJson ? { type: "json_object" } : undefined,
      }),
      cache: "no-store",
    });

    const body = await response.text();
    if (response.ok) {
      const data = JSON.parse(body) as { choices?: Array<{ message?: { content?: string } }> };
      return { text: data.choices?.[0]?.message?.content || "", provider: "Groq", model };
    }
    lastError = compactError("Groq", response, body);
    if (response.status !== 429) break;
  }

  throw new Error(lastError || "Groq failed");
}

async function callOpenRouter(options: GenerateOptions): Promise<AiResult> {
  const apiKey = envValue("OPENROUTER_API_KEY");
  if (!apiKey) throw new Error("OpenRouter key not configured");

  const model = taskModel("OPENROUTER_MODEL", "openrouter/free", options.task);
  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
      "HTTP-Referer": envValue("APP_PUBLIC_URL") || "https://scoutersmarter.vercel.app",
      "X-Title": "Scout Smarter",
    },
    body: JSON.stringify({
      model,
      messages: [
        ...(options.system ? [{ role: "system", content: options.system }] : []),
        { role: "user", content: options.prompt },
      ],
      temperature: options.temperature ?? 0.1,
      max_tokens: options.maxTokens ?? 2200,
    }),
    cache: "no-store",
  });

  const body = await response.text();
  if (!response.ok) throw new Error(compactError("OpenRouter", response, body));
  const data = JSON.parse(body) as { choices?: Array<{ message?: { content?: string } }> };
  return { text: data.choices?.[0]?.message?.content || "", provider: "OpenRouter", model };
}

export async function generateAiText(options: GenerateOptions): Promise<AiResult> {
  const configured = {
    gemini: Boolean(envValue("GEMINI_API_KEY") || envValue("GOOGLE_API_KEY")),
    groq: Boolean(envValue("GROQ_API_KEY")),
    openrouter: Boolean(envValue("OPENROUTER_API_KEY")),
  };
  const calls = {
    gemini: callGemini,
    groq: callGroq,
    openrouter: callOpenRouter,
  };
  const errors: string[] = [];

  for (const provider of providerOrder(options.task)) {
    if (!(provider in calls) || !configured[provider as keyof typeof configured]) continue;
    try {
      return await calls[provider as keyof typeof calls](options);
    } catch (error) {
      errors.push(error instanceof Error ? error.message : `${provider} failed`);
    }
  }

  if (!errors.length) {
    throw new Error("No AI provider configured. Add GEMINI_API_KEY, GROQ_API_KEY, or OPENROUTER_API_KEY in Vercel.");
  }
  throw new Error(errors.join(" | "));
}
