export type FactorKey =
  | "paidUpCapital"
  | "sector"
  | "geography"
  | "businessModel"
  | "directorProfile"
  | "capitalRatio"
  | "filingCompliance";

export type Company = {
  id: string;
  name: string;
  cin: string;
  sector: string;
  city: string;
  state: string;
  status: "Active" | "Rejected" | "Non-Active" | "Unverified" | "Scoring Failed";
  rejectionReason?: string;
  paidUpCapital: string;
  authorizedCapital: string;
  paidUpCapitalValue?: number;
  authorizedCapitalValue?: number;
  nicCode: string;
  activity: string;
  lastFiling: string;
  incorporationDate?: string;
  director: {
    name: string;
    role: string;
    education: string;
    directorships: number;
    din?: string;
    credibility: string;
  };
  factors: Record<FactorKey, number>;
  factorReasoning?: Partial<Record<FactorKey, string>>;
  compositeScore?: number | null;
  adjustedScore?: number | null;
  ipoReadinessBand?: IpoReadinessBand;
  ipoReadinessMessage?: string;
  redFlags?: string[];
  yellowFlags?: string[];
  statusVerification?: {
    source: string;
    statusFound: string;
    verifiedActive: boolean;
    rf08Applied: boolean;
    rf09Applied: boolean;
    checkedAt?: string;
  };
  aiScoringError?: string;
  competitors: string[];
};

export const factorLabels: Record<FactorKey, string> = {
  paidUpCapital: "Paid-up Capital",
  sector: "Sector Strength",
  geography: "Geography",
  businessModel: "Business Model",
  directorProfile: "Director Profile",
  capitalRatio: "Auth / Paid-up Ratio",
  filingCompliance: "Filing Compliance",
};

export const factorKeys = Object.keys(factorLabels) as FactorKey[];

export type FactorSelection = Record<FactorKey, boolean>;
export type FactorWeights = Record<FactorKey, number>;
export type IpoReadinessBand = "IPO Ready" | "Near Ready" | "Development Stage" | "Not Recommended";

export type ParserRejectionSummary = {
  totalUploaded: number;
  rejectedCapital: number;
  rejectedGeography: number;
  rejectedNic: number;
  rejectedTotal: number;
  passingToAi: number;
};

export const defaultFactorSelection = factorKeys.reduce((selection, key) => {
  selection[key] = true;
  return selection;
}, {} as FactorSelection);

export const defaultFactorWeights: FactorWeights = {
  sector: 20,
  businessModel: 20,
  paidUpCapital: 18,
  directorProfile: 16,
  filingCompliance: 12,
  capitalRatio: 8,
  geography: 6,
};

export const companies: Company[] = [
  {
    id: "aether-grid",
    name: "Aether Grid Systems Private Limited",
    cin: "U72900MH2018PTC314920",
    sector: "Technology",
    city: "Mumbai",
    state: "Maharashtra",
    status: "Active",
    paidUpCapital: "Rs 18.4 Cr",
    authorizedCapital: "Rs 25.0 Cr",
    paidUpCapitalValue: 184000000,
    authorizedCapitalValue: 250000000,
    nicCode: "62013",
    activity: "AI-led grid analytics and industrial automation platform",
    lastFiling: "4 months ago",
    director: {
      name: "Arjun Mehta",
      role: "Founder Director",
      education: "MBA, ISB | B.Tech, IIT Bombay",
      directorships: 4,
      credibility: "Prior enterprise SaaS operator with multiple active directorships",
    },
    factors: {
      paidUpCapital: 9.8,
      sector: 9.4,
      geography: 10,
      businessModel: 9.2,
      directorProfile: 10,
      capitalRatio: 8,
      filingCompliance: 10,
    },
    competitors: ["LogicLadder", "Gramener", "Flutura"],
  },
  {
    id: "mednova",
    name: "Mednova Devices Private Limited",
    cin: "U33110KA2017PTC102419",
    sector: "Healthcare",
    city: "Bengaluru",
    state: "Karnataka",
    status: "Active",
    paidUpCapital: "Rs 12.7 Cr",
    authorizedCapital: "Rs 16.0 Cr",
    paidUpCapitalValue: 127000000,
    authorizedCapitalValue: 160000000,
    nicCode: "32502",
    activity: "Medical device manufacturing for diagnostic and clinical use",
    lastFiling: "7 months ago",
    director: {
      name: "Dr. Kavya Rao",
      role: "Managing Director",
      education: "MS Biomedical Engineering",
      directorships: 3,
      credibility: "Healthcare founder with specialist education and repeat company exposure",
    },
    factors: {
      paidUpCapital: 9.3,
      sector: 9.2,
      geography: 10,
      businessModel: 8.4,
      directorProfile: 10,
      capitalRatio: 8,
      filingCompliance: 7,
    },
    competitors: ["Trivitron", "Transasia", "Molbio"],
  },
  {
    id: "surya-cell",
    name: "Surya Cell Energy Private Limited",
    cin: "U40106GJ2019PTC108412",
    sector: "Renewable Energy",
    city: "Ahmedabad",
    state: "Gujarat",
    status: "Active",
    paidUpCapital: "Rs 10.1 Cr",
    authorizedCapital: "Rs 10.1 Cr",
    paidUpCapitalValue: 101000000,
    authorizedCapitalValue: 101000000,
    nicCode: "35105",
    activity: "Solar module assembly and distributed renewable energy systems",
    lastFiling: "5 months ago",
    director: {
      name: "Rishabh Shah",
      role: "Promoter Director",
      education: "B.Tech Electrical Engineering",
      directorships: 2,
      credibility: "Energy sector operator with relevant technical background",
    },
    factors: {
      paidUpCapital: 9,
      sector: 9,
      geography: 10,
      businessModel: 7.6,
      directorProfile: 7,
      capitalRatio: 10,
      filingCompliance: 10,
    },
    competitors: ["Waaree", "Vikram Solar", "Goldi Solar"],
  },
  {
    id: "finpulse",
    name: "Finpulse Credit Analytics Private Limited",
    cin: "U67190DL2020PTC371190",
    sector: "Financial Services",
    city: "New Delhi",
    state: "Delhi",
    status: "Active",
    paidUpCapital: "Rs 8.9 Cr",
    authorizedCapital: "Rs 14.0 Cr",
    paidUpCapitalValue: 89000000,
    authorizedCapitalValue: 140000000,
    nicCode: "66190",
    activity: "Credit analytics and digital lending infrastructure",
    lastFiling: "9 months ago",
    director: {
      name: "Naina Kapoor",
      role: "Whole-time Director",
      education: "MBA Finance",
      directorships: 2,
      credibility: "Financial-services operator with two active directorships",
    },
    factors: {
      paidUpCapital: 8.7,
      sector: 8.4,
      geography: 10,
      businessModel: 7.9,
      directorProfile: 7,
      capitalRatio: 8,
      filingCompliance: 7,
    },
    competitors: ["Perfios", "FinBox", "CreditVidya"],
  },
  {
    id: "bharat-robotics",
    name: "Bharat Robotics Components Private Limited",
    cin: "U29309TN2016PTC111280",
    sector: "Manufacturing",
    city: "Chennai",
    state: "Tamil Nadu",
    status: "Active",
    paidUpCapital: "Rs 7.2 Cr",
    authorizedCapital: "Rs 12.0 Cr",
    paidUpCapitalValue: 72000000,
    authorizedCapitalValue: 120000000,
    nicCode: "28199",
    activity: "Precision components for robotics and factory automation",
    lastFiling: "11 months ago",
    director: {
      name: "Vikram Narayanan",
      role: "Director",
      education: "B.Tech Mechanical Engineering",
      directorships: 1,
      credibility: "Relevant technical background; limited public operating trail",
    },
    factors: {
      paidUpCapital: 8.3,
      sector: 7.2,
      geography: 10,
      businessModel: 8.1,
      directorProfile: 5,
      capitalRatio: 8,
      filingCompliance: 7,
    },
    competitors: ["Systemantics", "DiFACTO", "Gridbots"],
  },
];

function isFactorSelection(value: FactorSelection | FactorWeights): value is FactorSelection {
  return Object.values(value).every((item) => typeof item === "boolean");
}

export function weightsFromSelection(selection: FactorSelection): FactorWeights {
  const included = factorKeys.filter((key) => selection[key]);
  if (!included.length) {
    return factorKeys.reduce((weights, key) => {
      weights[key] = 0;
      return weights;
    }, {} as FactorWeights);
  }
  const equalWeight = 100 / included.length;
  return factorKeys.reduce((weights, key) => {
    weights[key] = selection[key] ? equalWeight : 0;
    return weights;
  }, {} as FactorWeights);
}

export function validateWeights(weights: FactorWeights) {
  const values = factorKeys.map((key) => Number(weights[key] ?? 0));
  const total = values.reduce((sum, weight) => sum + weight, 0);
  return values.every((weight) => weight >= 0 && weight <= 30) && Math.abs(total - 100) < 0.001;
}

export function calculateWeightedScore(company: Company, weights: FactorWeights = defaultFactorWeights) {
  if (company.status === "Rejected" || !validateWeights(weights)) {
    return 0;
  }
  const weighted = factorKeys.reduce((sum, key) => {
    return sum + (Number(company.factors[key] ?? 0) * Number(weights[key] ?? 0)) / 100;
  }, 0);
  return Math.max(0, Math.min(100, Math.round(weighted * 10)));
}

export function calculateScore(
  company: Company,
  weightsOrSelection: FactorWeights | FactorSelection = defaultFactorWeights,
) {
  if (company.status === "Rejected") {
    return 0;
  }
  if (typeof company.adjustedScore === "number") {
    return Math.max(0, Math.min(100, Math.round(company.adjustedScore)));
  }
  if (typeof company.compositeScore === "number") {
    return Math.max(0, Math.min(100, Math.round(company.compositeScore)));
  }
  const weights = isFactorSelection(weightsOrSelection) ? weightsFromSelection(weightsOrSelection) : weightsOrSelection;
  return calculateWeightedScore(company, weights);
}

export function rankCompanies(
  companyList: Company[],
  weightsOrSelection: FactorWeights | FactorSelection = defaultFactorWeights,
) {
  return companyList
    .filter((company) => company.status !== "Rejected")
    .sort((a, b) => {
      const scoreDelta = calculateScore(b, weightsOrSelection) - calculateScore(a, weightsOrSelection);
      if (scoreDelta) return scoreDelta;
      return (a.redFlags?.length ?? 0) - (b.redFlags?.length ?? 0);
    });
}

export const rankedCompanies = rankCompanies(companies);

export const topCompany = rankedCompanies[0];
