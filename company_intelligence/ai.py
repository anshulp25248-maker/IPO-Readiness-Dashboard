from __future__ import annotations

import re
from html import unescape
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

import requests
from bs4 import BeautifulSoup


PROVIDER_CONFIGS = {
    "OpenAI": {
        "model": "gpt-4.1",
        "base_url": "https://api.openai.com/v1",
        "help": "Standard OpenAI API base URL.",
    },
    "Claude": {
        "model": "claude-3-5-sonnet-latest",
        "base_url": "https://api.anthropic.com/v1",
        "help": "Anthropic Messages API base URL.",
    },
    "Gemini": {
        "model": "gemini-2.5-pro",
        "base_url": "https://generativelanguage.googleapis.com/v1beta",
        "help": "Google Generative Language API base URL.",
    },
    "OpenRouter": {
        "model": "openai/gpt-4.1",
        "base_url": "https://openrouter.ai/api/v1",
        "help": "OpenRouter Chat Completions base URL.",
    },
    "OpenClaw / OpenAI-Compatible": {
        "model": "openclaw",
        "base_url": "",
        "help": "Enter the OpenAI-compatible endpoint for OpenClaw or any custom hosted model.",
    },
    "Ollama": {
        "model": "llama3.1",
        "base_url": "http://localhost:11434",
        "help": "Local Ollama base URL. API key is not required for default local setups.",
    },
}

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0 Safari/537.36"
)


def _join_content(content: Any) -> str:
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts: list[str] = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict) and item.get("text"):
                parts.append(str(item["text"]))
        return "\n".join(parts)
    return str(content)


def call_llm(prompt: str, ai_config: dict[str, Any], max_tokens: int = 3000) -> str:
    provider = ai_config["provider"]
    model = ai_config["model"]
    api_key = ai_config.get("api_key", "")
    base_url = ai_config["base_url"].rstrip("/")
    temperature = ai_config.get("temperature", 0.2)
    system_prompt = (
        "You are a senior investment banker, forensic business analyst, sector strategist, "
        "and IPO readiness advisor for GreenFlow Ventures. Use only the company dossier and "
        "public research context supplied in the prompt. When information is inferred rather than "
        "explicitly supported, say that it is an inference."
    )

    if provider in {"OpenAI", "OpenRouter", "OpenClaw / OpenAI-Compatible"}:
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        if provider == "OpenRouter":
            headers["HTTP-Referer"] = "https://greenflow.ventures"
            headers["X-Title"] = "Company Intelligence System"
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        response = requests.post(f"{base_url}/chat/completions", headers=headers, json=payload, timeout=180)
        response.raise_for_status()
        data = response.json()
        return _join_content(data["choices"][0]["message"]["content"])

    if provider == "Claude":
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        payload = {
            "model": model,
            "system": system_prompt,
            "max_tokens": max_tokens,
            "temperature": temperature,
            "messages": [{"role": "user", "content": prompt}],
        }
        response = requests.post(f"{base_url}/messages", headers=headers, json=payload, timeout=180)
        response.raise_for_status()
        data = response.json()
        return "\n".join(item["text"] for item in data.get("content", []) if item.get("type") == "text")

    if provider == "Gemini":
        payload = {
            "systemInstruction": {"parts": [{"text": system_prompt}]},
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": temperature, "maxOutputTokens": max_tokens},
        }
        response = requests.post(
            f"{base_url}/models/{model}:generateContent?key={api_key}",
            json=payload,
            timeout=180,
        )
        response.raise_for_status()
        data = response.json()
        candidates = data.get("candidates", [])
        if not candidates:
            raise ValueError("Gemini returned no candidates.")
        parts = candidates[0].get("content", {}).get("parts", [])
        return "\n".join(part.get("text", "") for part in parts if part.get("text"))

    if provider == "Ollama":
        payload = {
            "model": model,
            "stream": False,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
            "options": {"temperature": temperature},
        }
        response = requests.post(f"{base_url}/api/chat", json=payload, timeout=180)
        response.raise_for_status()
        data = response.json()
        return _join_content(data.get("message", {}).get("content", ""))

    raise ValueError(f"Unsupported provider: {provider}")


def clean_result_url(url: str) -> str:
    if "duckduckgo.com/l/?" not in url:
        return url
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    return unquote(query.get("uddg", [url])[0])


def search_web(query: str, max_results: int = 6) -> list[dict[str, str]]:
    response = requests.get(
        "https://html.duckduckgo.com/html/",
        params={"q": query},
        headers={"User-Agent": USER_AGENT},
        timeout=30,
    )
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    items: list[dict[str, str]] = []
    for result in soup.select(".result"):
        anchor = result.select_one(".result__title a")
        if not anchor:
            continue
        title = unescape(anchor.get_text(" ", strip=True))
        url = clean_result_url(anchor.get("href", ""))
        snippet_node = result.select_one(".result__snippet")
        snippet = unescape(snippet_node.get_text(" ", strip=True)) if snippet_node else ""
        if title and url:
            items.append({"title": title, "url": url, "snippet": snippet})
        if len(items) >= max_results:
            break
    return items


def unique_results(*result_lists: list[dict[str, str]]) -> list[dict[str, str]]:
    deduped: list[dict[str, str]] = []
    seen: set[str] = set()
    for result_list in result_lists:
        for item in result_list:
            url = item.get("url", "")
            if url and url not in seen:
                seen.add(url)
                deduped.append(item)
    return deduped


def infer_leadership_hints(results: list[dict[str, str]]) -> list[str]:
    patterns = [
        re.compile(r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})\s*[-|,].{0,20}(Founder|Co-Founder|Director|CEO|Managing Director)", re.I),
        re.compile(r"(Founder|Co-Founder|Director|CEO|Managing Director).{0,20}([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2})", re.I),
    ]
    names: list[str] = []
    for result in results:
        text = f"{result.get('title', '')} {result.get('snippet', '')}"
        for pattern in patterns:
            match = pattern.search(text)
            if not match:
                continue
            groups = [item for item in match.groups() if item]
            for item in groups:
                cleaned = re.sub(r"\s+", " ", item).strip(" -|,")
                if cleaned.lower() in {"founder", "co-founder", "director", "ceo", "managing director"}:
                    continue
                if cleaned and cleaned not in names:
                    names.append(cleaned)
    return names[:5]


def build_company_research_context(company: dict[str, Any], extra_context: str = "") -> tuple[str, dict[str, Any]]:
    name = company["company_name"]
    cin = company["cin"]
    sector = company.get("sector", "")
    activity = company.get("activity_description", "")

    company_results = search_web(f'"{name}" "{cin}" company profile India', max_results=5)
    leadership_results = search_web(f'"{name}" directors founder linkedin', max_results=5)
    sector_results = search_web(f'India "{sector}" sector market growth listed peers', max_results=4)
    results = unique_results(company_results, leadership_results, sector_results)
    leadership = infer_leadership_hints(results)

    lines = [
        f"COMPANY: {name}",
        f"CIN / LLPIN: {cin}",
        f"SECTOR: {sector}",
        f"ACTIVITY: {activity}",
        "",
        "PUBLIC WEB RESEARCH SNIPPETS",
    ]
    for index, item in enumerate(results[:12], start=1):
        lines.append(f"{index}. {item['title']}")
        lines.append(f"   URL: {item['url']}")
        if item["snippet"]:
            lines.append(f"   Note: {item['snippet']}")
    if leadership:
        lines.extend(["", "LEADERSHIP HINTS", *[f"- {name}" for name in leadership]])
    if extra_context.strip():
        lines.extend(["", "ADDITIONAL USER CONTEXT", extra_context.strip()])
    return "\n".join(lines), {"sources": results, "leadership_hints": leadership}


def build_sector_research_context(sector: str, activity_hint: str = "") -> tuple[str, dict[str, Any]]:
    market_results = search_web(f'India "{sector}" market size CAGR policy listed players', max_results=6)
    ipo_results = search_web(f'India "{sector}" SME IPO recent performance', max_results=4)
    policy_results = search_web(f'India "{sector}" policy incentive regulation', max_results=4)
    results = unique_results(market_results, ipo_results, policy_results)

    lines = [f"SECTOR: {sector}", f"ACTIVITY HINT: {activity_hint}", "", "PUBLIC WEB RESEARCH SNIPPETS"]
    for index, item in enumerate(results[:12], start=1):
        lines.append(f"{index}. {item['title']}")
        lines.append(f"   URL: {item['url']}")
        if item["snippet"]:
            lines.append(f"   Note: {item['snippet']}")
    return "\n".join(lines), {"sources": results}


def build_public_company_snapshot(
    company_name: str,
    cin: str,
    sector: str,
    activity_description: str,
) -> dict[str, Any]:
    company_results = search_web(f'"{company_name}" "{cin}" company India', max_results=4)
    leadership_results = search_web(f'"{company_name}" founder director linkedin', max_results=4)
    results = unique_results(company_results, leadership_results)
    leadership = infer_leadership_hints(results)

    summary_parts: list[str] = []
    if sector:
        summary_parts.append(f"Sector cue: {sector}.")
    if activity_description:
        summary_parts.append(f"Activity cue: {activity_description}.")
    if results:
        summary_parts.append("Public-web matches suggest the following visible references:")
        summary_parts.extend(item["title"] for item in results[:3])

    return {"leadership_hints": leadership, "summary": " ".join(summary_parts), "sources": results}
