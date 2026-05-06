from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from config import Settings, get_logger
from memory import MemoryStore
from news_fetcher import NewsArticle


@dataclass(slots=True)
class ScoredArticle:
    """Holds the article plus the scoring breakdown used for selection."""

    article: NewsArticle
    score: float
    keyword_score: float
    global_score: float
    india_score: float
    recency_score: float
    urgency_score: float
    is_major_event: bool
    matched_keywords: list[str]
    reason: str
    topic_signature: str

    def to_dict(self) -> dict[str, Any]:
        payload = self.article.to_dict()
        payload.update(
            {
                "score": self.score,
                "keyword_score": self.keyword_score,
                "global_score": self.global_score,
                "india_score": self.india_score,
                "recency_score": self.recency_score,
                "urgency_score": self.urgency_score,
                "is_major_event": self.is_major_event,
                "matched_keywords": self.matched_keywords,
                "reason": self.reason,
                "topic_signature": self.topic_signature,
            }
        )
        return payload


class MacroNewsAnalyzer:
    """Scores news the way a macro analyst would rank market-moving events."""

    KEYWORD_WEIGHTS = {
        "inflation": 14,
        "inflation spike": 18,
        "interest rate": 15,
        "rate hike": 18,
        "rate cut": 14,
        "fed": 12,
        "ecb": 10,
        "boj": 10,
        "rbi": 14,
        "gdp": 10,
        "recession": 15,
        "oil": 12,
        "crude": 12,
        "opec": 12,
        "war": 20,
        "conflict": 18,
        "tariff": 14,
        "sanctions": 14,
        "currency": 12,
        "rupee": 12,
        "inr": 12,
        "dollar": 10,
        "default": 20,
        "collapse": 20,
        "crisis": 18,
    }

    GLOBAL_RELEVANCE_TERMS = {
        "us",
        "united states",
        "china",
        "europe",
        "eurozone",
        "japan",
        "russia",
        "ukraine",
        "middle east",
        "opec",
        "fed",
        "ecb",
        "boj",
        "treasury",
        "commodities",
        "oil",
        "dollar",
        "trade",
        "shipping",
        "sanctions",
    }

    INDIA_RELEVANCE_TERMS = {
        "india",
        "indian",
        "rbi",
        "rupee",
        "inr",
        "sensex",
        "nifty",
        "imports",
        "current account",
        "fii",
        "fpi",
    }

    MAJOR_EVENT_KEYWORDS = {
        "war",
        "invasion",
        "attack",
        "missile",
        "emergency",
        "collapse",
        "default",
        "crisis",
        "inflation spike",
        "surprise rate hike",
        "sanctions",
        "oil shock",
    }

    STOPWORDS = {
        "the",
        "and",
        "for",
        "that",
        "with",
        "from",
        "into",
        "about",
        "after",
        "amid",
        "over",
        "under",
        "will",
        "this",
        "when",
        "what",
        "have",
        "been",
        "more",
        "than",
        "says",
        "your",
        "their",
        "global",
        "market",
        "markets",
        "economy",
        "economic",
        "finance",
        "financial",
        "news",
    }

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("analyzer")

    def score_articles(self, articles: list[NewsArticle]) -> list[ScoredArticle]:
        """Score and rank every article."""
        scored = [self._score_article(article) for article in articles]
        return sorted(scored, key=lambda item: item.score, reverse=True)

    def select_top_articles(
        self,
        articles: list[NewsArticle],
        memory: MemoryStore,
        limit: int | None = None,
    ) -> list[ScoredArticle]:
        """Select the top 1-2 impactful stories while avoiding duplicates."""
        limit = limit or self.settings.top_articles_to_keep
        scored = self.score_articles(articles)
        selected: list[ScoredArticle] = []

        for candidate in scored:
            if candidate.score < 20:
                continue

            if memory.was_topic_recent(candidate.topic_signature) and not candidate.is_major_event:
                continue

            if any(
                self._topic_similarity(candidate.topic_signature, existing.topic_signature) >= 0.55
                for existing in selected
            ):
                continue

            if selected and candidate.score < selected[0].score * 0.68:
                continue

            selected.append(candidate)
            if len(selected) >= limit:
                break

        return selected

    def should_post_now(
        self,
        selected_articles: list[ScoredArticle],
        memory: MemoryStore,
        force: bool = False,
    ) -> tuple[bool, str]:
        """Decide whether to publish now based on cadence and urgency."""
        if force:
            return True, "Forced from the command line."

        if not selected_articles:
            return False, "No distinct high-impact stories passed the selection filters."

        days_since_last_post = memory.days_since_last_post()
        if days_since_last_post is None:
            return True, "No previous successful post was found in memory."

        if days_since_last_post >= self.settings.normal_post_interval_days:
            return True, f"{days_since_last_post:.1f} days have passed since the last post."

        if any(item.is_major_event for item in selected_articles):
            return True, "Major-event override triggered because a fresh high-impact shock was detected."

        return (
            False,
            f"Only {days_since_last_post:.1f} days have passed since the last post, and no major-event override fired.",
        )

    def _score_article(self, article: NewsArticle) -> ScoredArticle:
        text = article.combined_text().lower()
        matched_keywords = [keyword for keyword in self.KEYWORD_WEIGHTS if keyword in text]
        keyword_score = min(sum(self.KEYWORD_WEIGHTS[keyword] for keyword in matched_keywords), 45)

        global_hits = [term for term in self.GLOBAL_RELEVANCE_TERMS if term in text]
        global_score = min(len(global_hits) * 3.5, 18)

        india_hits = [term for term in self.INDIA_RELEVANCE_TERMS if term in text]
        india_score = min(len(india_hits) * 4.0, 14)

        hours_old = max(0.0, self._hours_old(article.published_at))
        recency_score = max(0.0, min(25.0, 25.0 * math.exp(-hours_old / 30.0)))

        urgent_hits = [term for term in self.MAJOR_EVENT_KEYWORDS if term in text]
        urgency_score = min(len(urgent_hits) * 6.0, 20.0)

        total_score = round(keyword_score + global_score + india_score + recency_score + urgency_score, 2)
        is_major_event = (
            bool(urgent_hits)
            and hours_old <= self.settings.major_event_window_hours
            and total_score >= 55
        )

        topic_signature = self._extract_topic_signature(article.combined_text())
        reason = self._build_reason(
            matched_keywords=matched_keywords,
            global_hits=global_hits,
            india_hits=india_hits,
            hours_old=hours_old,
            is_major_event=is_major_event,
        )

        return ScoredArticle(
            article=article,
            score=total_score,
            keyword_score=round(keyword_score, 2),
            global_score=round(global_score, 2),
            india_score=round(india_score, 2),
            recency_score=round(recency_score, 2),
            urgency_score=round(urgency_score, 2),
            is_major_event=is_major_event,
            matched_keywords=matched_keywords,
            reason=reason,
            topic_signature=topic_signature,
        )

    def _build_reason(
        self,
        *,
        matched_keywords: list[str],
        global_hits: list[str],
        india_hits: list[str],
        hours_old: float,
        is_major_event: bool,
    ) -> str:
        reason_parts = []

        if matched_keywords:
            reason_parts.append(f"Macro keywords: {', '.join(matched_keywords[:5])}")
        if global_hits:
            reason_parts.append(f"Global relevance: {', '.join(global_hits[:4])}")
        if india_hits:
            reason_parts.append(f"India angle: {', '.join(india_hits[:4])}")
        reason_parts.append(f"Recency: {hours_old:.1f} hours old")
        if is_major_event:
            reason_parts.append("Major-event override eligible")

        return " | ".join(reason_parts)

    def _extract_topic_signature(self, text: str) -> str:
        """Create a simple topic fingerprint from the most informative words."""
        tokens = re.findall(r"[a-zA-Z]{4,}", text.lower())
        filtered: list[str] = []
        for token in tokens:
            if token in self.STOPWORDS or token in filtered:
                continue
            filtered.append(token)
            if len(filtered) >= 6:
                break
        return " ".join(filtered)

    @staticmethod
    def _hours_old(published_at: datetime) -> float:
        now = datetime.now(timezone.utc)
        return (now - published_at).total_seconds() / 3600

    @staticmethod
    def _topic_similarity(left_signature: str, right_signature: str) -> float:
        left = set(left_signature.split())
        right = set(right_signature.split())
        if not left or not right:
            return 0.0
        return len(left.intersection(right)) / len(left.union(right))
