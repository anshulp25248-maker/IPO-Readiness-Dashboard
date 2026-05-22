import { envValue } from "./ai";

export type DashboardResearchLane = "other-companies" | "director-profile" | "competitor-analysis";

export type SearchResult = {
  title?: string;
  url?: string;
  content?: string;
};

const tavilyUrl = "https://api.tavily.com/search";

const laneConfig: Record<DashboardResearchLane, { envPrefix: string; fallbackPrefix: string; label: string }> = {
  "other-companies": {
    envPrefix: "DASHBOARD_OTHER_COMPANIES",
    fallbackPrefix: "CDR_OTHER_COMPANIES",
    label: "Other Companies",
  },
  "director-profile": {
    envPrefix: "DASHBOARD_DIRECTOR_PROFILE",
    fallbackPrefix: "CDR_DIRECTOR_PROFILE",
    label: "Director Profile",
  },
  "competitor-analysis": {
    envPrefix: "DASHBOARD_COMPETITOR_ANALYSIS",
    fallbackPrefix: "CDR_COMPETITOR_ANALYSIS",
    label: "Competitor Analysis",
  },
};

export function dashboardLaneLabel(lane: DashboardResearchLane) {
  return laneConfig[lane].label;
}

export function dashboardLaneEnvValue(lane: DashboardResearchLane, suffix: "GROQ_API_KEY" | "TAVILY_API_KEY") {
  const config = laneConfig[lane];
  const legacySuffix = suffix.replace("_API_KEY", "");

  return (
    envValue(`${config.envPrefix}_${suffix}`) ||
    envValue(`${legacySuffix}_API_KEY_${config.envPrefix}`) ||
    envValue(`${config.fallbackPrefix}_${suffix}`) ||
    envValue(`${legacySuffix}_API_KEY_${config.fallbackPrefix}`) ||
    envValue(suffix)
  );
}

async function tavilySearch(query: string, lane: DashboardResearchLane, maxResults: number) {
  const apiKey = dashboardLaneEnvValue(lane, "TAVILY_API_KEY");
  if (!apiKey) return [];

  const response = await fetch(tavilyUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: apiKey,
      query,
      search_depth: "advanced",
      include_answer: false,
      include_raw_content: false,
      max_results: maxResults,
    }),
    cache: "no-store",
    signal: AbortSignal.timeout(9000),
  });

  if (!response.ok) throw new Error(`Live feed search failed with ${response.status}`);
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

export async function tavilySearchLane(
  lane: DashboardResearchLane,
  queries: string[],
  options: { maxResultsPerQuery?: number; maxSources?: number } = {},
) {
  const maxResultsPerQuery = options.maxResultsPerQuery ?? 4;
  const maxSources = options.maxSources ?? 8;
  const batches = await Promise.all(
    queries
      .filter(Boolean)
      .slice(0, 4)
      .map((query) => tavilySearch(query, lane, maxResultsPerQuery).catch(() => [])),
  );

  const seen = new Set<string>();
  return batches
    .flat()
    .filter((item) => {
      const key = item.url || `${item.title}-${item.content}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, maxSources);
}

export function sourceContext(results: SearchResult[]) {
  if (!results.length) return "No live feed results returned for this dashboard lane. Use uploaded/parser data and mark unverifiable public facts as NA.";

  return results
    .map((item, index) =>
      [
        `${index + 1}. ${item.title || "Untitled"}`,
        `URL: ${item.url || "NA"}`,
        `Snippet: ${(item.content || "").slice(0, 650)}`,
      ].join("\n"),
    )
    .join("\n\n");
}

export function sourceLinks(results: SearchResult[]) {
  return results.map((item) => ({ title: item.title, url: item.url }));
}
