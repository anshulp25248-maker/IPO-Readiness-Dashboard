"use client";

import { createContext, useContext, useEffect, useMemo, useState } from "react";
import {
  Company,
  FactorSelection,
  calculateScore,
  companies as sampleCompanies,
  defaultFactorSelection,
  factorKeys,
  rankCompanies,
} from "../_data/companies";

export type AiCompanyInsight = {
  companyId: string;
  companyName: string;
  aiScore: number;
  recommendation: string;
  rationale: string;
  strengths: string[];
  redFlags: string[];
  missingData: string[];
};

type ScoutContextValue = {
  drawerOpen: boolean;
  setDrawerOpen: (open: boolean) => void;
  companies: Company[];
  rankedCompanies: Company[];
  topCompany: Company;
  factorSelection: FactorSelection;
  pendingFactorSelection: FactorSelection;
  includedFactorCount: number;
  uploadStatus: string;
  scoringStatus: string;
  aiScoreReport: string;
  aiCompanyInsights: Record<string, AiCompanyInsight>;
  scoreCompany: (company: Company) => number;
  toggleFactor: (key: keyof FactorSelection) => void;
  runScoring: () => Promise<void>;
  handleUpload: (file: File) => Promise<void>;
  replaceCompanies: (companyList: Company[], message: string) => void;
  resetCompanies: () => void;
};

const ScoutContext = createContext<ScoutContextValue | null>(null);

const storageKey = "scout-smarter-state-v3";

type StoredScoutState = {
  factorSelection?: FactorSelection;
  pendingFactorSelection?: FactorSelection;
};

type UploadParseResponse = {
  rows?: Record<string, string>[];
  rowCount?: number;
  fileName?: string;
  error?: string;
};

type ScoringAnalysisResponse = {
  report?: string;
  insights?: AiCompanyInsight[];
  error?: string;
};

function readStoredState(): StoredScoutState {
  if (typeof window === "undefined") {
    return {};
  }
  const saved = window.localStorage.getItem(storageKey);
  if (!saved) {
    return {};
  }
  try {
    return JSON.parse(saved) as StoredScoutState;
  } catch {
    window.localStorage.removeItem(storageKey);
    return {};
  }
}

function slug(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "company";
}

function normalizeHeader(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function getField(row: Record<string, string>, names: string[], fallback = "") {
  for (const name of names) {
    const normalizedName = normalizeHeader(name);
    const value = row[normalizedName] ?? row[name] ?? row[name.toUpperCase()];
    if (value && value.trim()) {
      return value.trim();
    }
  }
  return fallback;
}

function normalizeRow(row: Record<string, string>) {
  return Object.entries(row).reduce<Record<string, string>>((normalized, [key, value]) => {
    normalized[normalizeHeader(key)] = String(value ?? "").trim();
    return normalized;
  }, {});
}

const paidUpCapitalFloor = 1_000_000;

function parseMoney(value: string) {
  const normalized = String(value || "").toLowerCase().replace(/,/g, "");
  const match = normalized.match(/([0-9]+(?:\.[0-9]+)?)/);
  const amount = Number(match?.[1] ?? 0);
  if (!Number.isFinite(amount)) return 0;
  if (/\b(crore|crores|cr)\b/.test(normalized)) return amount * 10_000_000;
  if (/\b(lakh|lakhs|lac|lacs|l)\b/.test(normalized)) return amount * 100_000;
  if (/\b(thousand|k)\b/.test(normalized)) return amount * 1_000;
  return amount;
}

function formatMoney(value: number) {
  if (!value || !Number.isFinite(value)) {
    return "Data Not Available";
  }
  if (value >= 10_000_000) {
    return `Rs ${Math.round((value / 10_000_000) * 10) / 10} Cr`;
  }
  if (value >= 100_000) {
    return `Rs ${Math.round((value / 100_000) * 10) / 10} L`;
  }
  return `Rs ${value.toLocaleString("en-IN")}`;
}

function splitCsvLine(line: string) {
  const cells: string[] = [];
  let cell = "";
  let insideQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];
    if (char === '"' && nextChar === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if (char === "," && !insideQuotes) {
      cells.push(cell.trim());
      cell = "";
    } else {
      cell += char;
    }
  }
  cells.push(cell.trim());
  return cells;
}

function parseCsv(text: string) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (lines.length < 2) {
    return [];
  }
  const headers = splitCsvLine(lines[0]).map(normalizeHeader);
  return lines.slice(1).map((line) => {
    const cells = splitCsvLine(line);
    return headers.reduce<Record<string, string>>((row, header, index) => {
      row[header] = cells[index] ?? "";
      return row;
    }, {});
  });
}

const sectorScore: Record<string, number> = {
  technology: 9.4,
  healthcare: 9.2,
  "renewable energy": 9,
  defence: 10,
  "financial services": 8.4,
  manufacturing: 7.2,
  logistics: 6.8,
  infrastructure: 6.4,
  agriculture: 6.2,
  consumer: 5.8,
  education: 5.6,
};

const tierOneCities = ["mumbai", "delhi", "new delhi", "bengaluru", "bangalore", "hyderabad", "chennai", "pune", "ahmedabad", "kolkata", "gurgaon", "gurugram", "noida"];
const tierTwoCities = ["surat", "jaipur", "lucknow", "kanpur", "nagpur", "indore", "bhopal", "patna", "vadodara", "coimbatore", "kochi", "chandigarh"];

const highConvictionSectorTerms = [
  "semiconductor",
  "defence",
  "defense",
  "pharma",
  "medical device",
  "renewable",
  "ev",
  "battery",
  "artificial intelligence",
  "automation",
  "robotics",
];

const stableGrowthSectorTerms = [
  "healthcare",
  "financial",
  "saas",
  "software",
  "manufacturing",
  "logistics",
  "infrastructure",
  "consumer",
];

function scoreSectorStrength(sector: string, nicCode: string, activity: string) {
  const text = `${sector} ${nicCode} ${activity}`.toLowerCase();
  if (highConvictionSectorTerms.some((term) => text.includes(term))) return 10;
  if (/^(21|26|30|325|351|620|631)/.test(nicCode)) return 9;
  if (stableGrowthSectorTerms.some((term) => text.includes(term))) return 8;
  return sectorScore[sector.toLowerCase()] ?? 5;
}

function scoreGeography(city: string, state: string) {
  const text = `${city} ${state}`.toLowerCase();
  if (tierOneCities.some((item) => text.includes(item))) return 10;
  if (tierTwoCities.some((item) => text.includes(item))) return 7;
  return 4;
}

function scoreRatio(paid: number, authorized: number) {
  if (!paid || !authorized) return 0;
  const ratio = paid / authorized;
  if (ratio >= 1) return 10;
  if (ratio > 0.8) return 9;
  if (ratio >= 0.5) return 8;
  if (ratio >= 0.2) return 6;
  return 1;
}

function scoreFiling(value: string) {
  const lower = String(value || "").toLowerCase();
  if (!lower || lower === "na" || lower === "n/a") return { score: 0, reject: false, label: "NA" };

  const relative = lower.match(/(\d+)\s*(month|months|year|years)\s*ago/);
  let months = 0;
  if (relative) {
    months = Number(relative[1]) * (relative[2].startsWith("year") ? 12 : 1);
  } else {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return { score: 0, reject: false, label: "NA" };
    months = (new Date().getFullYear() - date.getFullYear()) * 12 + new Date().getMonth() - date.getMonth();
  }

  if (months <= 6) return { score: 10, reject: false, label: "within 6 months" };
  if (months <= 12) return { score: 7, reject: false, label: "6-12 months" };
  if (months <= 24) return { score: 4, reject: false, label: "12-24 months" };
  return { score: 0, reject: true, label: ">24 months" };
}

function scoreDirector(directorships: number, education: string) {
  const hasProfessionalDegree = /\b(b\.?tech|m\.?tech|mba|ms|ca|cfa|iit|iim|isb)\b/i.test(education);
  if (!directorships && (!education || education === "NA")) return 5;
  if (hasProfessionalDegree && directorships >= 2) return 10;
  if (directorships >= 3) return 10;
  if (directorships === 2) return 7;
  if (hasProfessionalDegree) return 5;
  if (directorships === 1) return 4;
  return 2;
}

function isActiveCompanyStatus(value: string) {
  const normalized = String(value || "").toLowerCase();
  if (!normalized) return true;
  if (/\b(inactive|strike|struck|closed|liquidation|liquidated|dissolved|dormant|amalgamated)\b/.test(normalized)) {
    return false;
  }
  return normalized.includes("active");
}

function mapRowsToCompanies(rows: Record<string, string>[]) {
  const normalizedRows = rows.map(normalizeRow);
  const paidValues = normalizedRows
    .map((row) => parseMoney(getField(row, ["paid_up_capital", "paidup capital", "paid up capital", "paid up share capital", "obligation_of_contribution_rs", "obligation of contribution(rs.)", "contribution"])))
    .filter((value) => value >= paidUpCapitalFloor);
  const maxPaid = Math.max(...paidValues, 1);
  const scarcity = new Map<string, number>();

  normalizedRows.forEach((row) => {
    const key = `${getField(row, ["nic_code", "nic", "activity_code"], "NA")}::${getField(row, ["activity", "activity_description", "business_activity"], "General")}`.toLowerCase();
    scarcity.set(key, (scarcity.get(key) ?? 0) + 1);
  });

  return normalizedRows.map<Company>((row, index) => {
    const name = getField(row, ["company_name", "company name", "llp_name", "llp name", "name"], `Uploaded Company ${index + 1}`);
    const sector = getField(row, ["sector", "activity_description", "description", "industrial_activity"], "Other");
    const city = getField(row, ["city", "district"], getField(row, ["state"], "NA"));
    const state = getField(row, ["state"], "NA");
    const paid = parseMoney(getField(row, ["paid_up_capital", "paidup capital", "paid up capital", "paid up share capital", "obligation_of_contribution_rs", "obligation of contribution(rs.)", "contribution"]));
    const authorized = parseMoney(getField(row, ["authorized_capital", "authorised_capital", "authorised capital", "authorized capital", "authorised share capital"], String(paid))) || paid;
    const nicCode = getField(row, ["nic_code", "nic", "activity_code", "industrial_activity"], "NA");
    const activity = getField(row, ["activity", "activity_description", "business_activity", "description"], "General business activity");
    const scarcityKey = `${nicCode}::${activity}`.toLowerCase();
    const scarcityCount = scarcity.get(scarcityKey) ?? 1;
    const directorships = Number(getField(row, ["director_directorships", "directorships"], "0")) || 0;
    const education = getField(row, ["director_education", "education"], "NA");
    const activeStatus = isActiveCompanyStatus(getField(row, ["status", "company_status", "company status"], "Active"));
    const filing = scoreFiling(getField(row, ["last_filing_date", "latest filing date", "last_filing"], ""));
    const rejectReasons = [
      !activeStatus ? "Company status is not active" : "",
      paid > 0 && paid < paidUpCapitalFloor ? "Paid-up capital is below Rs 10 lakh" : "",
      filing.reject ? "Latest filing is older than 24 months" : "",
    ].filter(Boolean);
    const status: Company["status"] = rejectReasons.length ? "Rejected" : "Active";

    return {
      id: `${slug(name)}-${index}`,
      name,
      cin: getField(row, ["cin", "llpin", "fcin"], `UPLOADED-${index + 1}`),
      sector,
      city,
      state,
      status,
      rejectionReason: rejectReasons.join("; ") || undefined,
      paidUpCapital: formatMoney(paid),
      authorizedCapital: formatMoney(authorized),
      paidUpCapitalValue: paid,
      authorizedCapitalValue: authorized,
      nicCode,
      activity,
      lastFiling: getField(row, ["last_filing_date", "latest filing date", "last_filing"], "NA"),
      director: {
        name: getField(row, ["director", "director_name", "directors"], "NA"),
        role: getField(row, ["director_role", "role"], "Director"),
        education,
        directorships,
        credibility: directorships >= 2 ? "Multiple directorship signal found" : "Limited public directorship signal",
      },
      factors: {
        paidUpCapital: paid <= 0 ? 0 : Math.min(10, (Math.log(Math.max(paid, 1)) / Math.log(Math.max(maxPaid, 1))) * 10),
        sector: scoreSectorStrength(sector, nicCode, activity),
        geography: scoreGeography(city, state),
        businessModel: scarcityCount <= 1 ? 10 : scarcityCount <= 10 ? 10 / scarcityCount : scarcityCount <= 100 ? 100 / scarcityCount : 1,
        directorProfile: scoreDirector(directorships, education),
        capitalRatio: scoreRatio(paid, authorized),
        filingCompliance: filing.score,
      },
      competitors: getField(row, ["competitors"], "NA")
        .split(/[|;]/)
        .map((item) => item.trim())
        .filter(Boolean)
        .slice(0, 3),
    };
  });
}

export function ScoutProvider({ children }: { children: React.ReactNode }) {
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [companies, setCompanies] = useState<Company[]>(sampleCompanies);
  const [factorSelection, setFactorSelection] = useState<FactorSelection>(() => ({
    ...defaultFactorSelection,
    ...readStoredState().factorSelection,
  }));
  const [pendingFactorSelection, setPendingFactorSelection] = useState<FactorSelection>(() => ({
    ...defaultFactorSelection,
    ...(readStoredState().pendingFactorSelection ?? readStoredState().factorSelection),
  }));
  const [uploadStatus, setUploadStatus] = useState("Sample company universe loaded.");
  const [scoringStatus, setScoringStatus] = useState("Scoring uses all seven factors with equal weight.");
  const [aiScoreReport, setAiScoreReport] = useState("");
  const [aiCompanyInsights, setAiCompanyInsights] = useState<Record<string, AiCompanyInsight>>({});

  useEffect(() => {
    try {
      window.localStorage.setItem(
        storageKey,
        JSON.stringify({ factorSelection, pendingFactorSelection }),
      );
      window.localStorage.removeItem("scout-smarter-state");
      window.localStorage.removeItem("scout-smarter-state-v2");
    } catch {
      window.localStorage.removeItem(storageKey);
    }
  }, [factorSelection, pendingFactorSelection]);

  const ranked = useMemo(() => rankCompanies(companies, factorSelection), [companies, factorSelection]);
  const includedFactorCount = factorKeys.filter((key) => factorSelection[key]).length;
  const pendingFactorCount = factorKeys.filter((key) => pendingFactorSelection[key]).length;

  async function handleUpload(file: File) {
    setUploadStatus(`Reading ${file.name} with pandas parser...`);

    let rows: Record<string, string>[] = [];
    try {
      const formData = new FormData();
      formData.append("file", file);
      const response = await fetch("/api/upload-companies", {
        method: "POST",
        body: formData,
      });
      const data = (await response.json()) as UploadParseResponse;
      if (!response.ok) {
        throw new Error(data.error || "File parsing failed.");
      }
      rows = data.rows || [];
    } catch (error) {
      if (!file.name.toLowerCase().endsWith(".csv")) {
        setUploadStatus(error instanceof Error ? error.message : "File parsing failed.");
        return;
      }
      rows = parseCsv(await file.text());
    }

    const parsedCompanies = mapRowsToCompanies(rows);
    if (!parsedCompanies.length) {
      setUploadStatus("No valid company records found in the uploaded file.");
      return;
    }
    setCompanies(parsedCompanies);
    setUploadStatus(`${parsedCompanies.length} companies uploaded. Ready to run scoring.`);
    setScoringStatus("Uploaded MCA data is ready. Choose factors, then press Run Scoring.");
  }

  function toggleFactor(key: keyof FactorSelection) {
    setPendingFactorSelection((current) => ({ ...current, [key]: !current[key] }));
  }

  async function runScoring() {
    setFactorSelection(pendingFactorSelection);
    setAiScoreReport("");
    setAiCompanyInsights({});
    setScoringStatus(
      `${pendingFactorCount} factor${pendingFactorCount === 1 ? "" : "s"} included. Running AI scoring layer...`,
    );

    const selectedRanked = rankCompanies(companies, pendingFactorSelection).slice(0, 10);
    const scores = selectedRanked.reduce<Record<string, number>>((values, company) => {
      values[company.id] = calculateScore(company, pendingFactorSelection);
      return values;
    }, {});

    if (!selectedRanked.length) {
      setScoringStatus("No eligible companies found after filters.");
      return;
    }

    try {
      const response = await fetch("/api/scoring-analysis", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ companies: selectedRanked, scores }),
      });
      const data = (await response.json()) as ScoringAnalysisResponse;
      if (!response.ok) throw new Error(data.error || "AI scoring layer failed.");
      const insightMap = (data.insights || []).reduce<Record<string, AiCompanyInsight>>((map, insight) => {
        if (insight.companyId) map[insight.companyId] = insight;
        return map;
      }, {});
      setAiScoreReport(data.report || "AI scoring layer completed.");
      setAiCompanyInsights(insightMap);
      setScoringStatus("Scoring complete. AI investment-readiness layer generated for the top companies.");
    } catch (error) {
      setScoringStatus(
        `${pendingFactorCount} factor${pendingFactorCount === 1 ? "" : "s"} included. AI layer unavailable: ${
          error instanceof Error ? error.message : "unknown error"
        }`,
      );
    }
  }

  function replaceCompanies(companyList: Company[], message: string) {
    setCompanies(companyList);
    setUploadStatus(message);
    setAiScoreReport("");
    setAiCompanyInsights({});
    setScoringStatus("AI feed loaded. Press Run Scoring after choosing factors.");
  }

  function resetCompanies() {
    setCompanies(sampleCompanies);
    setFactorSelection(defaultFactorSelection);
    setPendingFactorSelection(defaultFactorSelection);
    setUploadStatus("Sample company universe loaded.");
    setScoringStatus("Scoring uses all seven factors with equal weight.");
    setAiScoreReport("");
    setAiCompanyInsights({});
  }

  const value: ScoutContextValue = {
    drawerOpen,
    setDrawerOpen,
    companies,
    rankedCompanies: ranked,
    topCompany: ranked[0] ?? sampleCompanies[0],
    factorSelection,
    pendingFactorSelection,
    includedFactorCount,
    uploadStatus,
    scoringStatus,
    aiScoreReport,
    aiCompanyInsights,
    scoreCompany: (company) => calculateScore(company, factorSelection),
    toggleFactor,
    runScoring,
    handleUpload,
    replaceCompanies,
    resetCompanies,
  };

  return <ScoutContext.Provider value={value}>{children}</ScoutContext.Provider>;
}

export function useScout() {
  const context = useContext(ScoutContext);
  if (!context) {
    throw new Error("useScout must be used inside ScoutProvider");
  }
  return context;
}
