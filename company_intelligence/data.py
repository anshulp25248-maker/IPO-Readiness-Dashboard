from __future__ import annotations

import io
import math
import re
from typing import Any

import pandas as pd


FACTOR_DEFINITIONS = [
    {
        "key": "paid_up_capital",
        "label": "Paid Up Capital",
        "default": 10,
        "help": "Rewards companies that already show stronger capitalization in the MCA dataset.",
    },
    {
        "key": "geography",
        "label": "Geography",
        "default": 8,
        "help": "Rewards companies based in stronger commercial clusters, startup corridors, and execution-friendly regions.",
    },
    {
        "key": "directors_credibility",
        "label": "Directors Credibility",
        "default": 9,
        "help": "Proxy score built from public-data quality, corporate email quality, entity structure, and credibility signals available from MCA data.",
    },
    {
        "key": "management_quality",
        "label": "Management Quality",
        "default": 9,
        "help": "Proxy score built from capital efficiency, entity structure, and operational readiness signals.",
    },
    {
        "key": "compliances",
        "label": "Compliances",
        "default": 10,
        "help": "Rewards active status plus completeness of core MCA compliance fields.",
    },
    {
        "key": "esg_governance",
        "label": "ESG Governance",
        "default": 7,
        "help": "Sector-sensitive governance proxy based on activity profile and legal-data hygiene.",
    },
    {
        "key": "unique_business_model",
        "label": "Unique Business Model",
        "default": 8,
        "help": "Higher for differentiated sectors, more specific activity descriptions, and less generic business models.",
    },
    {
        "key": "active_status",
        "label": "ACTIVE Status",
        "default": 10,
        "help": "Rewards companies that are operationally active according to MCA records.",
    },
    {
        "key": "reporting",
        "label": "Reporting",
        "default": 7,
        "help": "Measures how reportable the company is based on structured legal, address, and contact information.",
    },
    {
        "key": "data_availability",
        "label": "DATA Availability and Accessibility",
        "default": 8,
        "help": "Rewards companies with better contactability, better legal completeness, and stronger public-data readiness.",
    },
]

FREE_EMAIL_DOMAINS = {
    "gmail.com",
    "yahoo.com",
    "yahoo.in",
    "hotmail.com",
    "outlook.com",
    "live.com",
    "aol.com",
    "icloud.com",
    "rediffmail.com",
    "protonmail.com",
}

STATE_PRIORITY = {
    "Maharashtra": 9.6,
    "Karnataka": 9.5,
    "Delhi": 9.4,
    "Telangana": 9.2,
    "Tamil Nadu": 9.0,
    "Gujarat": 8.8,
    "Haryana": 8.6,
    "Uttar Pradesh": 7.8,
    "West Bengal": 7.7,
    "Rajasthan": 7.5,
    "Kerala": 7.5,
    "Punjab": 7.2,
    "Odisha": 6.8,
    "Madhya Pradesh": 6.8,
    "Andhra Pradesh": 6.7,
    "Uttarakhand": 6.6,
    "Goa": 7.6,
    "Bihar": 5.8,
    "Jharkhand": 5.8,
    "Chhattisgarh": 5.9,
    "Assam": 5.8,
}

CITY_PRIORITY = {
    "mumbai": 1.0,
    "pune": 0.95,
    "bangalore": 1.0,
    "bengaluru": 1.0,
    "gurgaon": 0.95,
    "gurugram": 0.95,
    "hyderabad": 0.95,
    "noida": 0.92,
    "delhi": 0.9,
    "new delhi": 0.9,
    "chennai": 0.9,
    "ahmedabad": 0.86,
    "surat": 0.8,
    "kolkata": 0.78,
    "jaipur": 0.74,
    "kochi": 0.72,
}

SECTOR_KEYWORDS = {
    "Technology & Digital": ["software", "saas", "cloud", "technology", "digital", "data", "analytics", "platform", "it services"],
    "Financial Services": ["fintech", "nbfc", "finance", "lending", "capital", "broking", "insurance", "wealth"],
    "Healthcare & Pharma": ["health", "pharma", "diagnostic", "biotech", "medical", "clinical"],
    "Consumer & Retail": ["trading", "retail", "consumer", "fashion", "apparel", "fmcg", "ecommerce"],
    "Industrial & Manufacturing": ["manufacturing", "machinery", "equipment", "industrial", "engineering", "fabrication"],
    "Construction & Real Estate": ["construction", "real estate", "infrastructure", "civil", "builders", "contract"],
    "Renewable Energy & Mobility": ["renewable", "solar", "battery", "wind", "ev", "clean energy", "mobility"],
    "Logistics & Mobility": ["logistics", "transport", "shipping", "warehouse", "supply chain"],
    "Agriculture & Food": ["agri", "food", "farm", "dairy", "beverage", "crop"],
    "Education & Services": ["education", "training", "consulting", "services", "community", "social"],
}

ESG_BIAS = {
    "Renewable Energy & Mobility": 8.9,
    "Healthcare & Pharma": 7.7,
    "Technology & Digital": 7.3,
    "Education & Services": 7.2,
    "Agriculture & Food": 6.8,
    "Financial Services": 6.5,
    "Logistics & Mobility": 6.0,
    "Industrial & Manufacturing": 5.8,
    "Consumer & Retail": 5.7,
    "Construction & Real Estate": 4.8,
}

GENERIC_BUSINESS_TERMS = {"trading", "business services", "services", "construction"}


def dataframe_to_download_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def format_currency(value: Any) -> str:
    amount = pd.to_numeric(value, errors="coerce")
    if pd.isna(amount):
        return "Not available"
    return f"INR {float(amount):,.0f}"


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def slugify_header(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", normalize_text(name).lower()).strip("_")


def detect_header_row(raw: pd.DataFrame) -> int:
    expected_sets = [
        {"cin", "company_name", "state"},
        {"llpin", "limited_liability_partnership_name", "state"},
        {"company_name", "activity_description", "email"},
    ]
    best_index = 0
    best_score = -1
    for index in range(min(12, len(raw))):
        values = {slugify_header(value) for value in raw.iloc[index].tolist()}
        score = max(len(values & expected) for expected in expected_sets)
        if score > best_score:
            best_score = score
            best_index = index
    return best_index


def normalize_columns(columns: list[str]) -> dict[str, str]:
    aliases = {
        "cin": "cin",
        "llpin": "cin",
        "company_name": "company_name",
        "limited_liability_partnership_name": "company_name",
        "date_of_registration": "registration_date",
        "date_of_incorporation": "registration_date",
        "state": "state",
        "roc": "roc",
        "company_status": "status",
        "category": "category",
        "class": "company_class",
        "company_type": "company_type",
        "authorized_capital": "authorized_capital",
        "total_obligation_of_contribution": "authorized_capital",
        "paidup_capital": "paid_up_capital",
        "number_of_partners": "number_of_partners",
        "number_of_designated_partners": "number_of_designated_partners",
        "activity_code": "activity_code",
        "activity_description": "activity_description",
        "registered_office_address": "registered_address",
        "email": "email",
        "type_of_office": "company_type",
    }
    return {column: aliases.get(slugify_header(column), slugify_header(column) or "unknown") for column in columns}


def to_numeric(value: Any) -> float:
    parsed = pd.to_numeric(value, errors="coerce")
    return 0.0 if pd.isna(parsed) else float(parsed)


def infer_sector(activity_description: str, activity_code: str, company_name: str) -> str:
    haystack = " ".join([normalize_text(activity_description), normalize_text(activity_code), normalize_text(company_name)]).lower()
    for sector, keywords in SECTOR_KEYWORDS.items():
        if any(re.search(rf"\b{re.escape(keyword.lower())}\b", haystack) for keyword in keywords):
            return sector
    return "General Business Services" if haystack.strip() else "Unclassified"


def read_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=object)
    header_row = detect_header_row(raw)
    data = pd.read_excel(xls, sheet_name=sheet_name, header=header_row, dtype=object)
    data = data.rename(columns=normalize_columns(list(data.columns)))
    return data.loc[:, ~data.columns.duplicated()].copy()


def infer_entity_type(frame: pd.DataFrame, sheet_name: str) -> str:
    lowered = sheet_name.lower()
    if "llp" in lowered:
        return "LLP"
    if "foreign" in lowered:
        return "Foreign Company"
    if "number_of_partners" in frame.columns:
        return "LLP"
    return "Company"


def prepare_standard_frame(frame: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    entity_type = infer_entity_type(frame, sheet_name)
    defaults = {
        "cin": "",
        "company_name": "",
        "registration_date": "",
        "state": "",
        "roc": "",
        "status": "",
        "category": "",
        "company_class": "",
        "company_type": "",
        "authorized_capital": 0.0,
        "paid_up_capital": 0.0,
        "number_of_partners": 0.0,
        "number_of_designated_partners": 0.0,
        "activity_code": "",
        "activity_description": "",
        "registered_address": "",
        "email": "",
    }
    for column, default in defaults.items():
        if column not in frame.columns:
            frame[column] = default

    frame["cin"] = frame["cin"].map(normalize_text)
    frame["company_name"] = frame["company_name"].map(normalize_text)
    frame = frame[(frame["cin"] != "") & (frame["company_name"] != "")].copy()

    frame["registration_date"] = pd.to_datetime(frame["registration_date"], errors="coerce")
    frame["state"] = frame["state"].map(normalize_text)
    frame["roc"] = frame["roc"].map(normalize_text)
    frame["status"] = frame["status"].map(normalize_text).replace("", "Unknown")
    frame["category"] = frame["category"].map(normalize_text)
    frame["company_class"] = frame["company_class"].map(normalize_text)
    frame["company_type"] = frame["company_type"].map(normalize_text)
    frame["authorized_capital"] = frame["authorized_capital"].apply(to_numeric)
    frame["paid_up_capital"] = frame["paid_up_capital"].apply(to_numeric)
    frame["number_of_partners"] = frame["number_of_partners"].apply(to_numeric)
    frame["number_of_designated_partners"] = frame["number_of_designated_partners"].apply(to_numeric)
    frame["activity_code"] = frame["activity_code"].map(normalize_text)
    frame["activity_description"] = frame["activity_description"].map(normalize_text)
    frame["registered_address"] = frame["registered_address"].map(normalize_text)
    frame["email"] = frame["email"].map(normalize_text).str.lower()
    frame["entity_type"] = entity_type
    frame["raw_sheet"] = sheet_name
    frame["authorized_capital"] = frame["authorized_capital"].where(frame["authorized_capital"] > 0, frame["paid_up_capital"])
    if entity_type == "LLP":
        frame["paid_up_capital"] = frame["paid_up_capital"].where(frame["paid_up_capital"] > 0, frame["authorized_capital"])

    frame["sector"] = frame.apply(
        lambda row: infer_sector(row["activity_description"], row["activity_code"], row["company_name"]),
        axis=1,
    )
    frame["company_id"] = frame["entity_type"].astype(str) + "::" + frame["cin"].astype(str) + "::" + frame["company_name"].astype(str)
    return frame[
        [
            "company_id",
            "cin",
            "company_name",
            "registration_date",
            "state",
            "roc",
            "status",
            "category",
            "company_class",
            "company_type",
            "authorized_capital",
            "paid_up_capital",
            "number_of_partners",
            "number_of_designated_partners",
            "activity_code",
            "activity_description",
            "registered_address",
            "email",
            "sector",
            "entity_type",
            "raw_sheet",
        ]
    ]


def load_mca_workbook(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, list[str]]:
    suffix = file_name.lower().split(".")[-1]
    messages: list[str] = []
    if suffix == "csv":
        frame = pd.read_csv(io.BytesIO(file_bytes))
        frame = frame.rename(columns=normalize_columns(list(frame.columns)))
        prepared = prepare_standard_frame(frame, "CSV Upload")
        return prepared, [f"CSV upload parsed into {len(prepared):,} records."]

    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    frames: list[pd.DataFrame] = []
    for sheet_name in xls.sheet_names:
        prepared = prepare_standard_frame(read_sheet(xls, sheet_name), sheet_name)
        if not prepared.empty:
            frames.append(prepared)
            messages.append(f"{sheet_name}: {len(prepared):,} usable records")

    if not frames:
        return pd.DataFrame(), ["No company records could be parsed from the uploaded workbook."]

    combined = pd.concat(frames, ignore_index=True).drop_duplicates(subset=["company_id"]).reset_index(drop=True)
    return combined, messages


def robust_scale(series: pd.Series) -> pd.Series:
    clean = pd.to_numeric(series, errors="coerce").fillna(0.0)
    if clean.empty:
        return clean
    if clean.nunique() <= 1:
        return pd.Series([0.5 if clean.iloc[0] > 0 else 0.0] * len(clean), index=series.index)
    low = clean.quantile(0.1)
    high = clean.quantile(0.9)
    if high <= low:
        low = clean.min()
        high = clean.max()
    denominator = high - low if high != low else 1.0
    return ((clean - low) / denominator).clip(0, 1)


def email_domain(email: str) -> str:
    parts = normalize_text(email).lower().split("@")
    return parts[1] if len(parts) == 2 else ""


def field_completeness(row: pd.Series) -> float:
    checks = [
        bool(row["cin"]),
        bool(row["company_name"]),
        bool(row["registered_address"]),
        bool(row["email"]),
        bool(row["activity_description"]),
        bool(row["state"]),
        bool(row["roc"]),
        pd.notna(row["registration_date"]),
    ]
    return sum(checks) / len(checks)


def status_score(status: str) -> float:
    lowered = normalize_text(status).lower()
    if "active" in lowered:
        return 10.0
    if "dormant" in lowered:
        return 6.0
    return 4.0 if lowered else 3.0


def geography_score(state: str, address: str) -> float:
    score = STATE_PRIORITY.get(state, 6.0)
    lowered_address = normalize_text(address).lower()
    city_bonus = max((bonus for city, bonus in CITY_PRIORITY.items() if city in lowered_address), default=0.0)
    return min(10.0, score + city_bonus)


def sector_rarity_scores(df: pd.DataFrame) -> pd.Series:
    frequencies = df["sector"].value_counts(normalize=True)
    return df["sector"].map(lambda sector: 1 - frequencies.get(sector, 0)).fillna(0.0)


def specificity_score(text: str) -> float:
    lowered = normalize_text(text).lower()
    if not lowered:
        return 3.0
    bonus = 0.0
    if len(lowered.split()) >= 4:
        bonus += 1.2
    if lowered in GENERIC_BUSINESS_TERMS:
        bonus -= 1.8
    if any(token in lowered for token in ["analytics", "renewable", "biotech", "platform", "mobility", "automation"]):
        bonus += 1.8
    return max(2.5, min(10.0, 5.8 + bonus))


def status_score_series(series: pd.Series) -> pd.Series:
    return series.map(status_score)


def score_companies(df: pd.DataFrame, weights: dict[str, int]) -> pd.DataFrame:
    if df.empty:
        return df

    scored = df.copy()
    paid_up_log = scored["paid_up_capital"].apply(lambda value: math.log10(value + 1.0))
    capital_norm = robust_scale(paid_up_log)
    capital_utilization = (scored["paid_up_capital"] / scored["authorized_capital"].replace(0, 1)).clip(0, 1.5)
    completeness = scored.apply(field_completeness, axis=1)
    rarity = sector_rarity_scores(scored)
    corporate_domain = scored["email"].map(email_domain).map(lambda domain: domain not in FREE_EMAIL_DOMAINS and bool(domain))

    scored["paid_up_capital_score"] = (2 + capital_norm * 8).clip(0, 10)
    scored["geography_score"] = scored.apply(lambda row: geography_score(row["state"], row["registered_address"]), axis=1)
    scored["directors_credibility_score"] = (
        3.2
        + status_score_series(scored["status"]) * 0.18
        + completeness * 2.3
        + capital_norm * 1.6
        + corporate_domain.astype(float) * 1.4
        + (scored["number_of_designated_partners"] >= 2).astype(float) * 0.8
    ).clip(0, 10)
    scored["management_quality_score"] = (
        3.0
        + capital_utilization.clip(0, 1).astype(float) * 2.6
        + corporate_domain.astype(float) * 1.4
        + scored["entity_type"].eq("Company").astype(float) * 0.6
        + scored["company_type"].str.contains("private", case=False, na=False).astype(float) * 0.8
        + (scored["number_of_partners"] >= 2).astype(float) * 0.8
        + completeness * 1.2
    ).clip(0, 10)
    scored["compliances_score"] = (
        2.0
        + status_score_series(scored["status"]) * 0.45
        + completeness * 3.4
        + scored["roc"].ne("").astype(float) * 0.7
        + scored["category"].ne("").astype(float) * 0.6
    ).clip(0, 10)
    scored["esg_governance_score"] = (
        scored["sector"].map(lambda sector: ESG_BIAS.get(sector, 6.0))
        + completeness * 1.1
        + corporate_domain.astype(float) * 0.6
    ).clip(0, 10)
    scored["unique_business_model_score"] = (rarity * 3.2 + scored["activity_description"].map(specificity_score)).clip(0, 10)
    scored["active_status_score"] = scored["status"].map(status_score).clip(0, 10)
    scored["reporting_score"] = (
        2.6
        + completeness * 4.1
        + scored["email"].ne("").astype(float) * 1.0
        + scored["registered_address"].ne("").astype(float) * 1.0
        + scored["activity_description"].ne("").astype(float) * 1.3
    ).clip(0, 10)
    scored["data_availability_score"] = (
        2.8
        + completeness * 4.4
        + corporate_domain.astype(float) * 1.0
        + scored["cin"].ne("").astype(float) * 0.8
        + scored["email"].ne("").astype(float) * 1.0
    ).clip(0, 10)

    weighted_total = pd.Series(0.0, index=scored.index)
    total_weight = sum(max(0, int(value)) for value in weights.values())
    for factor in FACTOR_DEFINITIONS:
        factor_key = factor["key"]
        weighted_total += scored[f"{factor_key}_score"] * max(0, int(weights.get(factor_key, 0)))

    scored["score"] = 0.0 if total_weight <= 0 else ((weighted_total / total_weight) * 10).clip(0, 100)
    return scored.sort_values(["score", "paid_up_capital"], ascending=[False, False]).reset_index(drop=True)


def get_company_record(df: pd.DataFrame, company_id: str | None) -> pd.Series | None:
    if df.empty or not company_id:
        return None
    match = df[df["company_id"] == company_id]
    return None if match.empty else match.iloc[0]


def build_factor_table(record: pd.Series, weights: dict[str, int]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Factor": factor["label"],
                "Raw score (0-10)": round(float(record.get(f"{factor['key']}_score", 0.0)), 1),
                "Weight": weights.get(factor["key"], 0),
            }
            for factor in FACTOR_DEFINITIONS
        ]
    )
