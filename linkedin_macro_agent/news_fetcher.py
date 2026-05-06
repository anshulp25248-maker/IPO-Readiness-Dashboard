from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from config import Settings, get_logger


@dataclass(slots=True)
class NewsArticle:
    """Normalized article shape shared across providers."""

    title: str
    description: str
    url: str
    published_at: datetime
    source: str
    provider: str

    def combined_text(self) -> str:
        return f"{self.title} {self.description}".strip()

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "description": self.description,
            "url": self.url,
            "published_at": self.published_at.isoformat(),
            "source": self.source,
            "provider": self.provider,
        }


class NewsFetcher:
    """Fetches finance, macro, and geopolitical news from external feeds."""

    NEWSDATA_URL = "https://newsdata.io/api/1/latest"
    NEWSAPI_URL = "https://newsapi.org/v2/everything"
    GDELT_URL = "https://api.gdeltproject.org/api/v2/doc/doc"

    QUERY_BUCKETS = [
        '(inflation OR "interest rates" OR "central bank" OR GDP OR recession OR "monetary policy" OR Fed OR ECB OR RBI)',
        '(oil OR OPEC OR crude OR energy OR sanctions OR tariffs OR trade OR shipping OR "supply chain")',
        '(war OR conflict OR geopolitics OR election OR crisis OR China OR "Middle East" OR Russia OR Ukraine)',
    ]

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("news_fetcher")
        self.session = self._build_session()

    def fetch_latest_news(self) -> list[NewsArticle]:
        """Fetch news from the preferred provider and the backup provider."""
        providers = []
        preferred = self.settings.news_provider.lower()
        if preferred in {"newsdata", "newsapi", "gdelt"}:
            providers.append(preferred)
        for provider in ("newsdata", "newsapi", "gdelt"):
            if provider not in providers:
                providers.append(provider)

        collected: list[NewsArticle] = []
        for provider in providers:
            try:
                if provider == "newsdata":
                    collected.extend(self._fetch_from_newsdata())
                elif provider == "newsapi":
                    collected.extend(self._fetch_from_newsapi())
                elif provider == "gdelt":
                    collected.extend(self._fetch_from_gdelt())
            except requests.RequestException as exc:
                self.logger.warning("News provider %s failed: %s", provider, exc)

        deduped = self._deduplicate_articles(collected)
        deduped.sort(key=lambda article: article.published_at, reverse=True)
        return deduped[: self.settings.max_articles_to_fetch]

    def _build_session(self) -> requests.Session:
        """Create a retry-enabled HTTP session for provider calls."""
        session = requests.Session()
        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=1.0,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods={"GET", "POST"},
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        session.headers.update({"User-Agent": "linkedin-macro-agent/1.0"})
        return session

    def _fetch_from_newsdata(self) -> list[NewsArticle]:
        """Pull articles from NewsData.io using its latest-news endpoint."""
        if not self.settings.newsdata_api_key:
            self.logger.info("NEWSDATA_API_KEY missing. Skipping NewsData.io.")
            return []

        articles: list[NewsArticle] = []
        page_token = ""
        max_pages = 4
        timeframe = f"{max(1, min(self.settings.fetch_window_hours, 48))}h"

        for _ in range(max_pages):
            params = {
                "apikey": self.settings.newsdata_api_key,
                "language": self.settings.news_language,
                "category": self.settings.newsdata_categories,
                "timeframe": timeframe,
                "removeduplicate": "1",
            }
            if self.settings.newsdata_country:
                params["country"] = self.settings.newsdata_country
            if self.settings.newsdata_query:
                params["q"] = self.settings.newsdata_query
            if self.settings.newsdata_query_in_title:
                params["qInTitle"] = self.settings.newsdata_query_in_title
            if page_token:
                params["page"] = page_token

            response = self._request_newsdata(params)
            response.raise_for_status()
            payload = response.json()

            for item in payload.get("results", []):
                title = (item.get("title") or "").strip()
                url = (item.get("link") or "").strip()
                published_at = self._parse_datetime(item.get("pubDate", ""))
                if not title or not url or not published_at:
                    continue

                description = (
                    item.get("description")
                    or item.get("content")
                    or item.get("snippet")
                    or ""
                )
                source = item.get("source_name") or item.get("source_id") or "Unknown"

                articles.append(
                    NewsArticle(
                        title=title,
                        description=description.strip(),
                        url=url,
                        published_at=published_at,
                        source=source,
                        provider="newsdata",
                    )
                )

            page_token = (payload.get("nextPage") or "").strip()
            if not page_token or len(articles) >= self.settings.max_articles_to_fetch:
                break

        return articles[: self.settings.max_articles_to_fetch]

    def _request_newsdata(self, params: dict[str, Any]) -> requests.Response:
        """Request NewsData.io and retry once without custom query fields if they are invalid."""
        response = self.session.get(
            self.NEWSDATA_URL,
            params=params,
            timeout=self.settings.request_timeout_seconds,
        )

        if response.status_code != 422:
            return response

        has_custom_query = bool(params.get("q") or params.get("qInTitle"))
        if not has_custom_query:
            return response

        fallback_params = dict(params)
        fallback_params.pop("q", None)
        fallback_params.pop("qInTitle", None)

        self.logger.warning(
            "NewsData.io rejected the custom query filters with HTTP 422. Retrying without q/qInTitle."
        )
        return self.session.get(
            self.NEWSDATA_URL,
            params=fallback_params,
            timeout=self.settings.request_timeout_seconds,
        )

    def _fetch_from_newsapi(self) -> list[NewsArticle]:
        """Pull articles from NewsAPI if an API key is available."""
        if not self.settings.news_api_key:
            self.logger.info("NEWS_API_KEY missing. Skipping NewsAPI.")
            return []

        articles: list[NewsArticle] = []
        since = (self._utcnow() - timedelta(hours=self.settings.fetch_window_hours)).isoformat()
        page_size = max(10, min(100, self.settings.max_articles_to_fetch // len(self.QUERY_BUCKETS) + 5))

        for query in self.QUERY_BUCKETS:
            params = {
                "q": query,
                "language": self.settings.news_language,
                "sortBy": "publishedAt",
                "searchIn": "title,description",
                "pageSize": page_size,
                "from": since,
                "apiKey": self.settings.news_api_key,
            }
            response = self.session.get(
                self.NEWSAPI_URL,
                params=params,
                timeout=self.settings.request_timeout_seconds,
            )
            response.raise_for_status()
            payload = response.json()

            for item in payload.get("articles", []):
                title = (item.get("title") or "").strip()
                url = (item.get("url") or "").strip()
                published_at = self._parse_datetime(item.get("publishedAt", ""))
                if not title or not url or not published_at:
                    continue

                articles.append(
                    NewsArticle(
                        title=title,
                        description=(item.get("description") or "").strip(),
                        url=url,
                        published_at=published_at,
                        source=(item.get("source") or {}).get("name", "Unknown"),
                        provider="newsapi",
                    )
                )

        return articles

    def _fetch_from_gdelt(self) -> list[NewsArticle]:
        """Pull articles from the public GDELT DOC API as a backup source."""
        articles: list[NewsArticle] = []
        max_records = max(10, self.settings.max_articles_to_fetch // len(self.QUERY_BUCKETS) + 5)

        for query in self.QUERY_BUCKETS:
            params = {
                "query": query,
                "mode": "ArtList",
                "maxrecords": max_records,
                "sort": "datedesc",
                "format": "json",
            }
            response = self.session.get(
                self.GDELT_URL,
                params=params,
                timeout=self.settings.request_timeout_seconds,
            )
            response.raise_for_status()
            payload = response.json()

            for item in payload.get("articles", []):
                title = (item.get("title") or "").strip()
                url = (item.get("url") or "").strip()
                published_at = self._parse_datetime(item.get("seendate", ""))
                if not title or not url or not published_at:
                    continue

                articles.append(
                    NewsArticle(
                        title=title,
                        description=(item.get("excerpt") or item.get("description") or "").strip(),
                        url=url,
                        published_at=published_at,
                        source=(item.get("domain") or item.get("sourcecountry") or "Unknown"),
                        provider="gdelt",
                    )
                )

        return articles

    def _deduplicate_articles(self, articles: list[NewsArticle]) -> list[NewsArticle]:
        """Remove duplicates by URL and by near-identical titles."""
        seen_urls: set[str] = set()
        seen_titles: set[str] = set()
        deduped: list[NewsArticle] = []

        for article in articles:
            normalized_url = article.url.split("?")[0].strip().lower()
            normalized_title = " ".join(article.title.lower().split())
            if normalized_url in seen_urls or normalized_title in seen_titles:
                continue

            seen_urls.add(normalized_url)
            seen_titles.add(normalized_title)
            deduped.append(article)

        return deduped

    @staticmethod
    def _utcnow() -> datetime:
        return datetime.now(timezone.utc)

    @staticmethod
    def _parse_datetime(value: str) -> datetime | None:
        """Handle both ISO timestamps and GDELT's seendate format."""
        if not value:
            return None

        known_formats = (
            "%Y-%m-%dT%H:%M:%SZ",
            "%Y-%m-%dT%H:%M:%S.%fZ",
            "%Y-%m-%dT%H:%M:%S%z",
            "%Y%m%dT%H%M%SZ",
        )

        for fmt in known_formats:
            try:
                parsed = datetime.strptime(value, fmt)
                if parsed.tzinfo is None:
                    parsed = parsed.replace(tzinfo=timezone.utc)
                return parsed.astimezone(timezone.utc)
            except ValueError:
                continue

        try:
            parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            return parsed.astimezone(timezone.utc)
        except ValueError:
            return None
