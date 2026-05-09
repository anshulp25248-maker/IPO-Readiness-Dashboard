import {
  Company,
  FactorKey,
  FactorWeights,
  IpoReadinessBand,
  ParserRejectionSummary,
  calculateWeightedScore,
  defaultFactorWeights,
  factorKeys,
} from "../_data/companies";

export const flagMessages: Record<string, string> = {
  "RF-01": "Paid-up capital below Rs 10 Lakh. Below BSE SME minimum threshold. Not eligible for SME IPO in current capital structure.",
  "RF-02": "Director directorship overload detected. Verify DIN independently on MCA portal before proceeding.",
  "RF-03": "No confirmed MCA filing in 24 plus months or annual forms are unavailable. Verify current regulatory standing on MCA portal directly.",
  "RF-04": "Paid-up capital exceeds authorised capital. This is either a data error or an illegal capital structure. Immediate disqualification recommended.",
  "RF-05": "Registered NIC code does not match stated business description. Verify actual operations and correct NIC code before proceeding.",
  "RF-06": "Company is less than 3 years old. Not currently eligible for SME IPO under the operating-history requirement.",
  "RF-07": "NIC code not present in uploaded data. AI has estimated sector from company name and description. Verify before relying on sector score.",
  "RF-08": "Company status verified as non-active on a public source. 30-point score penalty applied. Recommend rejection unless status can be formally disputed.",
  "RF-09": "Company not found on Zauba Corp, Tofler, or MCA portal. Status unverifiable from public sources. Exercise caution and verify directly with company.",
  "YF-01": "Authorised capital nearly fully utilised. Fresh issue at IPO will require authorised capital increase. Confirm promoter plan.",
  "YF-02": "Capital severely underdeployed relative to authorised limit. Verify company is operational and not a dormant registration.",
  "YF-03": "Gaps detected in historical filing record. Request explanation from promoter before proceeding.",
  "YF-04": "Only one directorship held by primary director. Verify independence plan and board expansion roadmap required for IPO compliance.",
  "YF-05": "Director education and operating background could not be confirmed. Request CV and references directly.",
  "YF-06": "Director holds 6 to 10 directorships. Verify each company type for shell company affiliation before proceeding.",
  "YF-07": "Primary director background does not align with company sector. Assess domain expertise gap and promoter plan.",
  "YF-08": "Company meets minimum 3-year IPO eligibility but has limited operating track record. Weight recent financials heavily in due diligence.",
  "YF-09": "Company registered in mining sector. Verify specific SEBI and sector regulator requirements.",
  "YF-10": "Company in retail or food service sector without clear chain or franchise scale signal. Verify number of locations and revenue concentration.",
};

export const bandMessages: Record<IpoReadinessBand, string> = {
  "IPO Ready": "Strong candidate for Pre-IPO round consideration. Proceed to full CDR generation and financial due diligence. All screening criteria met.",
  "Near Ready": "Promising screening profile. Address flagged items before committing to Pre-IPO round. Recommend CDR generation with focus on flagged areas.",
  "Development Stage": "Potential exists but meaningful gaps are present in current profile. Consider for watchlist or debt syndication mandate instead of equity.",
  "Not Recommended": "Insufficient profile for current SME IPO mandate. Structural or compliance gaps require resolution before investment consideration.",
};

const approvedCities = [
  "mumbai", "pune", "nagpur", "nashik", "aurangabad", "thane", "navi mumbai", "solapur", "kolhapur", "amravati", "sangli",
  "new delhi", "noida", "greater noida", "gurugram", "faridabad", "ghaziabad",
  "bengaluru", "mysuru", "hubli", "dharwad", "mangaluru", "belagavi", "davangere",
  "chennai", "coimbatore", "madurai", "tirupur", "salem", "tiruchirappalli", "vellore", "erode", "tirunelveli", "hosur", "ambattur",
  "ahmedabad", "surat", "vadodara", "rajkot", "gandhinagar", "bhavnagar", "jamnagar", "anand", "morbi", "mehsana", "vapi", "bharuch",
  "hyderabad", "warangal", "nizamabad", "karimnagar", "khammam",
  "visakhapatnam", "vijayawada", "guntur", "tirupati", "kakinada", "rajahmundry", "nellore",
  "jaipur", "jodhpur", "udaipur", "kota", "ajmer", "bikaner", "alwar", "bhilwara", "sikar",
  "lucknow", "kanpur", "agra", "varanasi", "allahabad", "meerut", "mathura", "moradabad", "bareilly", "aligarh", "gorakhpur",
  "ludhiana", "amritsar", "jalandhar", "patiala", "bathinda", "mohali", "chandigarh",
  "panipat", "ambala", "yamunanagar", "rohtak", "hisar", "karnal", "sonipat",
  "kolkata", "durgapur", "asansol", "siliguri", "howrah", "kharagpur", "haldia",
  "indore", "bhopal", "jabalpur", "gwalior", "ujjain", "ratlam", "dewas", "pithampur",
  "kochi", "thiruvananthapuram", "kozhikode", "thrissur", "kollam", "kannur", "palakkad", "alappuzha",
  "bhubaneswar", "cuttack", "rourkela", "sambalpur", "berhampur",
  "ranchi", "jamshedpur", "dhanbad", "bokaro",
  "raipur", "bhilai", "durg", "bilaspur", "korba",
  "guwahati", "silchar", "dibrugarh",
  "patna", "gaya", "bhagalpur", "muzaffarpur",
  "shimla", "baddi", "solan", "paonta sahib",
  "dehradun", "haridwar", "roorkee", "rudrapur", "kashipur", "haldwani", "rishikesh",
  "panaji", "margao", "vasco da gama", "mapusa",
  "jammu", "srinagar", "agartala", "shillong", "silvassa",
];

const cityAliases: Record<string, string> = {
  bangaluru: "bengaluru",
  bangalore: "bengaluru",
  bombay: "mumbai",
  calcutta: "kolkata",
  madras: "chennai",
  gurgaon: "gurugram",
};

const metroCities = new Set(["mumbai", "delhi", "new delhi", "bengaluru", "chennai", "hyderabad", "kolkata", "ahmedabad", "pune"]);

function text(value: unknown) {
  return String(value ?? "").trim();
}

function lower(value: unknown) {
  return text(value).toLowerCase();
}

function nicNumber(value: string) {
  const digits = value.replace(/\D/g, "");
  if (!digits) return null;
  return Number(digits.slice(0, 5));
}

function inRange(code: number | null, start: number, end: number) {
  return code !== null && code >= start && code <= end;
}

function companyText(company: Company) {
  return lower(
    [
      company.name,
      company.sector,
      company.activity,
      company.nicCode,
      company.city,
      company.state,
      company.director?.role,
      company.director?.credibility,
    ].join(" "),
  );
}

function isCommunityServiceCompany(company: Company) {
  const code = nicNumber(company.nicCode);
  const body = companyText(company);

  if (inRange(code, 87000, 88999) || inRange(code, 94000, 94999) || inRange(code, 97000, 99000)) return true;

  return /\b(community service|community.*social services|social service|social services|social work|charitable|charity|ngo|non government|non-government|non profit|non-profit|not for profit|not-for-profit|section 8|section-8|foundation|trust|society|welfare association|welfare society|religious|temple|mosque|church|gurudwara|ashram|club|membership organisation|membership organization|trade association|chamber of commerce|resident welfare|rwa)\b/.test(
    body,
  );
}

function isGovernmentCompany(company: Company) {
  const code = nicNumber(company.nicCode);
  const body = companyText(company);

  if (inRange(code, 84100, 84399)) return true;

  return /\b(government|govt|govt\.|public sector|psu|state owned|state-owned|central government|state government|government of|govt of|municipal|municipality|panchayat|ministry|department|development authority|industrial development authority|urban development authority|state electricity|state power|state road|state transport|state industrial|state infrastructure|electricity board|water board|housing board|transport corporation|road transport|smart city mission|cantonment|zilla parishad|gram panchayat)\b/.test(
    body,
  );
}

function rejectCompany(company: Company, rejectionReason: string) {
  return {
    ...company,
    status: "Rejected",
    rejectionReason,
    redFlags: company.redFlags ?? [],
    yellowFlags: company.yellowFlags ?? [],
    aiScoringError: `Parser rejected: ${rejectionReason}`,
  } satisfies Company;
}

function monthsSince(value: string) {
  const raw = lower(value);
  const relative = raw.match(/(\d+)\s*(month|months|year|years)\s*ago/);
  if (relative) return Number(relative[1]) * (relative[2].startsWith("year") ? 12 : 1);
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  const now = new Date();
  return Math.max(0, (now.getFullYear() - date.getFullYear()) * 12 + now.getMonth() - date.getMonth());
}

function yearsSince(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  const now = new Date();
  let years = now.getFullYear() - date.getFullYear();
  if (now.getMonth() < date.getMonth() || (now.getMonth() === date.getMonth() && now.getDate() < date.getDate())) years -= 1;
  return years;
}

function levenshtein(a: string, b: string) {
  const matrix = Array.from({ length: a.length + 1 }, (_, row) => [row]);
  for (let col = 1; col <= b.length; col += 1) matrix[0][col] = col;
  for (let row = 1; row <= a.length; row += 1) {
    for (let col = 1; col <= b.length; col += 1) {
      const cost = a[row - 1] === b[col - 1] ? 0 : 1;
      matrix[row][col] = Math.min(matrix[row - 1][col] + 1, matrix[row][col - 1] + 1, matrix[row - 1][col - 1] + cost);
    }
  }
  return matrix[a.length][b.length];
}

export function normalizeApprovedCity(city: string, state = "", address = "") {
  const cityText = lower(city);
  const stateText = lower(state);
  const addressText = lower(address);
  const joined = `${cityText} ${stateText} ${addressText}`;
  if (/(village|vill\.|gram|tehsil|taluka|rural|post office|p\.o\.|dist\.)/.test(joined)) return null;
  if (/(lakshadweep|andaman and nicobar|daman and diu|dadra and nagar haveli)/.test(joined) && !joined.includes("silvassa")) return null;

  const aliased = cityAliases[cityText] ?? cityText;
  if (approvedCities.includes(aliased)) return aliased;
  const embedded = approvedCities.find((item) => joined.includes(item));
  if (embedded) return embedded;

  let best = "";
  let bestScore = 0;
  for (const item of approvedCities) {
    const distance = levenshtein(aliased, item);
    const score = 1 - distance / Math.max(aliased.length, item.length, 1);
    if (score > bestScore) {
      best = item;
      bestScore = score;
    }
  }
  return bestScore >= 0.8 ? best : null;
}

export function applyParserFilters(companies: Company[]) {
  const summary: ParserRejectionSummary = {
    totalUploaded: companies.length,
    rejectedCapital: 0,
    rejectedGeography: 0,
    rejectedNic: 0,
    rejectedCommunityService: 0,
    rejectedGovernment: 0,
    rejectedTotal: 0,
    passingToAi: 0,
  };

  const passing: Company[] = [];
  const rejected: Company[] = [];

  companies.forEach((company) => {
    const paid = company.paidUpCapitalValue ?? 0;
    if (paid < 500_000) {
      summary.rejectedCapital += 1;
      rejected.push(rejectCompany(company, "Paid-up capital below Rs 5 Lakh"));
      return;
    }

    if (isCommunityServiceCompany(company)) {
      summary.rejectedCommunityService += 1;
      rejected.push(rejectCompany(company, "Community service, NGO, trust, society, welfare, or non-profit profile"));
      return;
    }

    if (isGovernmentCompany(company)) {
      summary.rejectedGovernment += 1;
      rejected.push(rejectCompany(company, "Government, public-sector, municipal, authority, or state-owned profile"));
      return;
    }

    passing.push({
      ...company,
      status: "Active",
    });
  });

  summary.rejectedTotal = rejected.length;
  summary.passingToAi = passing.length;
  return { passing, rejected, summary };
}

function clusterScore(nicCode: string, city: string) {
  const code = nicNumber(nicCode);
  const normalizedCity = normalizeApprovedCity(city) ?? lower(city);
  const inList = (cities: string[]) => cities.includes(normalizedCity);
  const result = (score: number, reasoning: string, match = score >= 8) => ({ score, clusterMatch: match, reasoning });

  const clusters = [
    { start: 62000, end: 62999, sector: "IT and Software", prime: ["bengaluru", "hyderabad", "pune", "chennai", "noida", "gurugram"], good: ["mumbai", "kolkata", "kochi"] },
    { start: 40000, end: 40299, sector: "Renewable Energy", prime: ["ahmedabad", "rajkot", "surat", "jaipur", "chennai", "coimbatore"], good: ["hyderabad", "pune", "bengaluru", "bhubaneswar"] },
    { start: 33110, end: 33190, sector: "Medical Devices", prime: ["bengaluru", "mumbai", "ahmedabad", "faridabad", "roorkee"], good: ["chennai", "new delhi", "pune"] },
    { start: 21000, end: 21020, sector: "Pharmaceuticals", prime: ["ahmedabad", "hyderabad", "mumbai", "baddi", "sikkim", "haridwar"], good: ["pune", "bengaluru", "chennai", "new delhi"] },
    { start: 26000, end: 26800, sector: "Electronics and Components", prime: ["bengaluru", "noida", "chennai", "hyderabad", "pune", "ambala", "roorkee"], good: [] },
    { start: 13100, end: 13990, sector: "Textiles", prime: ["surat", "tirupur", "ludhiana", "ahmedabad", "coimbatore", "mumbai", "panipat", "bhilwara", "erode"], good: [] },
    { start: 29000, end: 29309, sector: "Machinery and Equipment", prime: ["pune", "coimbatore", "rajkot", "chennai", "ludhiana", "faridabad", "aurangabad", "nashik"], good: [] },
    { start: 10000, end: 10890, sector: "Food Processing", prime: ["new delhi", "ludhiana", "mumbai", "pune", "ahmedabad", "hyderabad", "chennai", "kolkata", "jalandhar", "amritsar"], good: [] },
    { start: 64000, end: 66990, sector: "Financial Services", prime: ["mumbai", "new delhi", "bengaluru", "chennai", "ahmedabad", "hyderabad", "kolkata"], good: [] },
    { start: 45000, end: 45309, sector: "Auto Components", prime: ["pune", "chennai", "gurugram", "faridabad", "coimbatore", "aurangabad", "rajkot", "ludhiana"], good: [] },
  ];

  const cluster = clusters.find((item) => inRange(code, item.start, item.end));
  if (cluster) {
    if (inList(cluster.prime)) return result(10, `City is ${city}. NIC code maps to ${cluster.sector}. ${city} is a prime sector cluster.`);
    if (inList(cluster.good)) return result(8, `City is ${city}. NIC code maps to ${cluster.sector}. ${city} is a good sector cluster.`);
    return result(6, `City is ${city}. NIC code maps to ${cluster.sector}. The city is approved but not a prime cluster.`, false);
  }

  if (metroCities.has(normalizedCity)) return result(8, `${city} is a metro or major approved business city for an unlisted NIC cluster.`);
  if (approvedCities.includes(normalizedCity)) return result(6, `${city} is in the approved city list but does not map to a named sector cluster.`, false);
  return result(5, `${city} is not a named metro cluster but passed parser geography checks.`, false);
}

function paidUpScore(company: Company) {
  const paid = company.paidUpCapitalValue ?? 0;
  let score = 0;
  if (paid > 250_000_000) score = 10;
  else if (paid >= 150_000_000) score = 9;
  else if (paid >= 100_000_000) score = 8;
  else if (paid >= 50_000_000) score = 7;
  else if (paid >= 30_000_000) score = 5;
  else if (paid >= 10_000_000) score = 3;
  else if (paid >= 1_000_000) score = 1;
  const code = nicNumber(company.nicCode);
  if ((inRange(code, 62000, 63999) || inRange(code, 64000, 66999) || /professional services/i.test(company.activity)) && paid > 50_000_000) {
    score = Math.min(10, score + 1);
  }
  return { score, reasoning: `${company.paidUpCapital} falls in the prescribed paid-up capital tier for a score of ${score}/10.` };
}

function sectorScore(company: Company) {
  const code = nicNumber(company.nicCode);
  const textBody = lower(`${company.sector} ${company.activity} ${company.name}`);
  let policy = 2;
  if (/(pharma|medical device|telecom|white goods|specialty steel|food|textile|solar|battery|automobile|auto component|drone|semiconductor|electronics)/.test(textBody)) policy = 4;
  else if (/(defence|defense|aerospace|railway|space|green hydrogen|renewable|port|logistics)/.test(textBody)) policy = 3;
  else if (/(liquor|tobacco|gambling)/.test(textBody)) policy = 1;
  else if (/(coal|thermal|legacy plastic)/.test(textBody)) policy = 0;

  let appetite = 1;
  if (inRange(code, 62000, 63999) || inRange(code, 21000, 33199) || /(technology|health|renewable|defence|ev|fintech|specialty chemical)/.test(textBody)) appetite = 3;
  else if (/(food|textile|auto|logistics|fmcg)/.test(textBody)) appetite = 2;
  else if (/(mining|agriculture|commodity)/.test(textBody)) appetite = 0;

  let growth = 1;
  if (/(ai|technology|renewable|defence|ev|healthtech|fintech|semiconductor|battery)/.test(textBody)) growth = 3;
  else if (/(pharma|specialty chemical|auto|fmcg|logistics)/.test(textBody)) growth = 2;
  else if (/(retail|print|commodity|coal)/.test(textBody)) growth = 0;

  return {
    score: Math.max(0, Math.min(10, policy + appetite + growth)),
    reasoning: `Policy tailwind ${policy}/4, SME IPO appetite ${appetite}/3, and revenue growth potential ${growth}/3 based on NIC/activity sector mapping.`,
  };
}

function businessModelScore(company: Company) {
  const body = lower(`${company.name} ${company.activity} ${company.sector}`);
  const consistency = company.nicCode && company.nicCode !== "NA" ? 3 : 1;
  let scalability = 1;
  if (/(saas|platform|marketplace|ip|software|analytics|ai)/.test(body)) scalability = 3;
  else if (/(b2b|branded|specialty|manufacturing|device|product|process)/.test(body)) scalability = 2;
  else if (/(commodity|single client|project)/.test(body)) scalability = 0;
  const visibility = /(subscription|contract|retainer|annuity|long-term)/.test(body) ? 2 : /(repeat|regular|distribution|manufacturing)/.test(body) ? 1 : 0;
  const narrative = /(ai|machine learning|ev|renewable|clean tech|healthtech|medtech|fintech|defence|aerospace|semiconductor|electronics|space|green hydrogen|specialty chemical)/.test(body) ? 2 : 1;
  return {
    score: Math.max(0, Math.min(10, consistency + scalability + visibility + narrative)),
    reasoning: `NIC consistency ${consistency}/3, scalability ${scalability}/3, revenue visibility ${visibility}/2, and IPO narrative strength ${narrative}/2.`,
  };
}

function directorScore(company: Company) {
  const director = company.director;
  const education = lower(director.education);
  const directorships = Number(director.directorships || 0);
  const redFlags: string[] = [];
  const yellowFlags: string[] = [];
  const educationScore = /(iit|iim|isb|aiims|nlu|bits|lse|wharton|mit|stanford|oxford)/.test(education) ? 2 : /(ca|cfa|mba|b\.?tech|m\.?tech|engineering|management)/.test(education) ? 1 : 0;
  if (!educationScore) yellowFlags.push("YF-05");
  const experienceScore = director.credibility && !/limited|unavailable|na/i.test(director.credibility) ? 2 : 0;
  if (!experienceScore) yellowFlags.push("YF-05");
  let directorshipScore = 2;
  if (directorships >= 2 && directorships <= 5) directorshipScore = 3;
  else if (directorships === 1) {
    directorshipScore = 2;
    yellowFlags.push("YF-04");
  } else if (directorships >= 6 && directorships <= 10) {
    directorshipScore = 1;
    yellowFlags.push("YF-06");
  } else if (directorships > 10) {
    directorshipScore = 0;
    redFlags.push("RF-02");
  }
  const sectorAlignment = 1;
  return {
    score: Math.max(0, Math.min(10, educationScore + experienceScore + directorshipScore + sectorAlignment)),
    redFlags,
    yellowFlags,
    reasoning: `Education ${educationScore}/2, operating experience ${experienceScore}/3, directorship count ${directorshipScore}/3, and sector alignment ${sectorAlignment}/2.`,
  };
}

function filingScore(company: Company) {
  const months = monthsSince(company.lastFiling);
  const redFlags: string[] = [];
  const yellowFlags: string[] = [];
  let recency = 0;
  if (months === null) recency = 1;
  else if (months <= 6) recency = 4;
  else if (months <= 12) recency = 3;
  else if (months <= 18) recency = 2;
  else if (months <= 24) recency = 1;
  else {
    recency = 0;
    redFlags.push("RF-03");
  }
  const formText = lower(`${company.activity} ${company.lastFiling}`);
  const hasAoc = /aoc-?4/.test(formText);
  const hasMgt = /mgt-?7/.test(formText);
  const forms = hasAoc && hasMgt ? 3 : hasAoc || hasMgt ? 1 : months !== null && months <= 24 ? 1 : 0;
  if (!forms) redFlags.push("RF-03");
  const consistency = months !== null && months <= 18 ? 3 : months !== null && months <= 24 ? 1 : 1;
  if (consistency <= 1) yellowFlags.push("YF-03");
  return {
    score: Math.max(0, Math.min(10, recency + forms + consistency)),
    redFlags,
    yellowFlags,
    reasoning: `Last filing recency scored ${recency}/4, form completeness ${forms}/3, and filing consistency ${consistency}/3 from uploaded MCA fields.`,
  };
}

function ratioScore(company: Company) {
  const paid = company.paidUpCapitalValue ?? 0;
  const authorized = company.authorizedCapitalValue ?? 0;
  const ratio = authorized ? (paid / authorized) * 100 : 0;
  const redFlags: string[] = [];
  const yellowFlags: string[] = [];
  let score = 0;
  if (ratio >= 40 && ratio <= 75) score = 10;
  else if (ratio > 75 && ratio <= 90) score = 7;
  else if (ratio >= 10 && ratio < 40) score = 5;
  else if (ratio > 90 && ratio <= 100) {
    score = 4;
    yellowFlags.push("YF-01");
  } else if (ratio > 0 && ratio < 10) {
    score = 2;
    yellowFlags.push("YF-02");
  } else if (ratio > 100) {
    score = 0;
    redFlags.push("RF-04");
  }
  return {
    score,
    ratio,
    redFlags,
    yellowFlags,
    reasoning: `Paid-up to authorised capital utilisation is ${Math.round(ratio * 10) / 10}%, giving ${score}/10 under the SME IPO headroom rule.`,
  };
}

export function assignReadinessBand(score: number, redFlags: string[], yellowFlags: string[]): IpoReadinessBand {
  if (redFlags.includes("RF-04") || redFlags.includes("RF-06") || redFlags.includes("RF-08")) return "Not Recommended";
  if (redFlags.length >= 2 || score < 55) return "Not Recommended";
  if (redFlags.length === 1 || (score >= 55 && score <= 69)) return "Development Stage";
  if ((score >= 70 && score <= 84) || (score >= 85 && yellowFlags.length === 1)) return "Near Ready";
  if (score >= 85 && redFlags.length === 0) return "IPO Ready";
  return "Development Stage";
}

export function scoreCompanyDeterministically(company: Company, weights: FactorWeights = defaultFactorWeights) {
  const redFlags = new Set(company.redFlags ?? []);
  const yellowFlags = new Set(company.yellowFlags ?? []);
  const reasoning: Partial<Record<FactorKey, string>> = {};

  const sector = sectorScore(company);
  const business = businessModelScore(company);
  const paid = paidUpScore(company);
  const director = directorScore(company);
  const filing = filingScore(company);
  const ratio = ratioScore(company);
  const geography = clusterScore(company.nicCode, company.city);

  director.redFlags.forEach((flag) => redFlags.add(flag));
  director.yellowFlags.forEach((flag) => yellowFlags.add(flag));
  filing.redFlags.forEach((flag) => redFlags.add(flag));
  filing.yellowFlags.forEach((flag) => yellowFlags.add(flag));
  ratio.redFlags.forEach((flag) => redFlags.add(flag));
  ratio.yellowFlags.forEach((flag) => yellowFlags.add(flag));

  if ((company.paidUpCapitalValue ?? 0) < 1_000_000) redFlags.add("RF-01");
  const age = company.incorporationDate ? yearsSince(company.incorporationDate) : null;
  if (age !== null && age < 3) redFlags.add("RF-06");
  else if (age !== null && age < 5) yellowFlags.add("YF-08");
  if (!company.nicCode || company.nicCode === "NA") redFlags.add("RF-07");

  const factors: Record<FactorKey, number> = {
    sector: sector.score,
    businessModel: business.score,
    paidUpCapital: paid.score,
    directorProfile: director.score,
    filingCompliance: filing.score,
    capitalRatio: ratio.score,
    geography: geography.score,
  };
  reasoning.sector = sector.reasoning;
  reasoning.businessModel = business.reasoning;
  reasoning.paidUpCapital = paid.reasoning;
  reasoning.directorProfile = director.reasoning;
  reasoning.filingCompliance = filing.reasoning;
  reasoning.capitalRatio = ratio.reasoning;
  reasoning.geography = geography.reasoning;

  const compositeScore = calculateWeightedScore({ ...company, factors }, weights);
  const adjustedScore = Math.max(0, compositeScore - (redFlags.has("RF-08") ? 30 : 0));
  const band = assignReadinessBand(adjustedScore, [...redFlags], [...yellowFlags]);

  const status: Company["status"] =
    company.status === "Rejected" ? "Rejected" : redFlags.has("RF-08") ? "Non-Active" : company.status;

  return {
    ...company,
    status,
    factors,
    factorReasoning: reasoning,
    compositeScore,
    adjustedScore,
    redFlags: [...redFlags],
    yellowFlags: [...yellowFlags],
    ipoReadinessBand: band,
    ipoReadinessMessage: bandMessages[band],
  };
}

export function normalizeWeights(input: Partial<FactorWeights> | undefined): FactorWeights {
  const weights = { ...defaultFactorWeights, ...(input ?? {}) };
  factorKeys.forEach((key) => {
    weights[key] = Math.max(0, Math.min(30, Number(weights[key] ?? 0)));
  });
  const total = factorKeys.reduce((sum, key) => sum + weights[key], 0);
  if (Math.abs(total - 100) < 0.001) return weights;
  return defaultFactorWeights;
}
