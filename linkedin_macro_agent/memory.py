from __future__ import annotations

import hashlib
import json
from datetime import datetime, timedelta, timezone
from typing import Any

from config import Settings, get_logger


class MemoryStore:
    """Persists posting history so the agent can avoid repeating itself."""

    DEFAULT_STATE = {
        "last_posted_at": None,
        "recent_posts": [],
    }

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("memory")

    def load_state(self) -> dict[str, Any]:
        """Load memory from disk, or return a clean default state."""
        if not self.settings.state_file.exists():
            return {"last_posted_at": None, "recent_posts": []}

        try:
            return json.loads(self.settings.state_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self.logger.warning("State file was invalid JSON. Starting from a clean state.")
            return {"last_posted_at": None, "recent_posts": []}

    def save_state(self, state: dict[str, Any]) -> None:
        """Write memory back to disk in a readable format."""
        self.settings.state_file.write_text(
            json.dumps(state, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    def days_since_last_post(self) -> float | None:
        """Return the number of days since the last successful post."""
        state = self.load_state()
        last_posted_at = state.get("last_posted_at")
        if not last_posted_at:
            return None

        parsed = self._parse_datetime(last_posted_at)
        if not parsed:
            return None

        elapsed = self.utcnow() - parsed
        return elapsed.total_seconds() / 86_400

    def was_topic_recent(
        self,
        topic_signature: str,
        lookback_days: int = 21,
        similarity_threshold: float = 0.60,
    ) -> bool:
        """Check whether the current topic is too similar to a recent post."""
        candidate_tokens = self._signature_tokens(topic_signature)
        if not candidate_tokens:
            return False

        cutoff = self.utcnow() - timedelta(days=lookback_days)
        state = self.load_state()

        for entry in state.get("recent_posts", []):
            posted_at = self._parse_datetime(entry.get("posted_at", ""))
            if not posted_at or posted_at < cutoff:
                continue

            historic_tokens = self._signature_tokens(entry.get("topic_signature", ""))
            if self._jaccard_similarity(candidate_tokens, historic_tokens) >= similarity_threshold:
                return True

        return False

    def record_post(
        self,
        *,
        topic_signature: str,
        selected_articles: list[dict[str, Any]],
        generated_post_text: str,
        method: str,
        post_id: str = "",
        used_fallback_generator: bool = False,
    ) -> None:
        """Persist a successful post to memory so future runs can avoid duplicates."""
        state = self.load_state()
        posted_at = self.utcnow().isoformat()

        entry = {
            "posted_at": posted_at,
            "topic_signature": topic_signature,
            "article_titles": [item.get("title", "") for item in selected_articles],
            "article_urls": [item.get("url", "") for item in selected_articles],
            "method": method,
            "post_id": post_id,
            "used_fallback_generator": used_fallback_generator,
            "post_hash": self.hash_text(generated_post_text),
        }

        recent_posts = [entry, *state.get("recent_posts", [])]
        state["last_posted_at"] = posted_at
        state["recent_posts"] = recent_posts[:50]
        self.save_state(state)

    def record_run(
        self,
        *,
        status: str,
        reason: str,
        selected_articles: list[dict[str, Any]] | None = None,
        generated_post_text: str = "",
        posting_method: str = "none",
        post_id: str = "",
        used_fallback_generator: bool = False,
        error_message: str = "",
    ) -> None:
        """Append a JSONL record for every pipeline run."""
        record = {
            "timestamp": self.utcnow().isoformat(),
            "status": status,
            "reason": reason,
            "posting_method": posting_method,
            "post_id": post_id,
            "used_fallback_generator": used_fallback_generator,
            "generated_post_text": generated_post_text,
            "selected_articles": selected_articles or [],
            "error_message": error_message,
        }

        with self.settings.run_history_file.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")

    @staticmethod
    def utcnow() -> datetime:
        """Small helper so all timestamps stay consistent."""
        return datetime.now(timezone.utc)

    @staticmethod
    def hash_text(text: str) -> str:
        """Hash text for lightweight deduping and auditing."""
        return hashlib.sha256(text.encode("utf-8")).hexdigest()

    @staticmethod
    def _parse_datetime(value: str) -> datetime | None:
        """Parse ISO timestamps from the memory file."""
        if not value:
            return None
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            return None

    @staticmethod
    def _signature_tokens(signature: str) -> set[str]:
        return {token for token in signature.lower().split() if token}

    @staticmethod
    def _jaccard_similarity(left: set[str], right: set[str]) -> float:
        if not left or not right:
            return 0.0
        intersection = left.intersection(right)
        union = left.union(right)
        return len(intersection) / len(union)
