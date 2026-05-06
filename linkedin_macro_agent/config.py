from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from logging.handlers import RotatingFileHandler
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent


def _read_bool(name: str, default: bool = False) -> bool:
    """Read a boolean from environment variables."""
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _read_int(name: str, default: int) -> int:
    """Read an integer from environment variables with a safe fallback."""
    value = os.getenv(name)
    if value is None or not value.strip():
        return default
    try:
        return int(value)
    except ValueError:
        return default


@dataclass(slots=True)
class Settings:
    """Central settings object shared by every module."""

    ai_provider: str
    openai_api_key: str
    openai_model: str
    openai_reasoning_effort: str
    gemini_api_key: str
    gemini_model: str
    news_provider: str
    news_api_key: str
    newsdata_api_key: str
    newsdata_country: str
    newsdata_categories: str
    newsdata_query: str
    newsdata_query_in_title: str
    news_language: str
    fetch_window_hours: int
    max_articles_to_fetch: int
    top_articles_to_keep: int
    posting_mode: str
    linkedin_access_token: str
    linkedin_person_urn: str
    linkedin_person_id: str
    linkedin_client_id: str
    linkedin_client_secret: str
    linkedin_redirect_uri: str
    linkedin_email: str
    linkedin_password: str
    linkedin_headless: bool
    chrome_binary_path: str
    normal_post_interval_days: int
    major_event_window_hours: int
    check_interval_minutes: int
    request_timeout_seconds: int
    state_file: Path
    run_history_file: Path
    generated_posts_dir: Path
    log_file: Path

    @property
    def data_dir(self) -> Path:
        return self.state_file.parent

    @property
    def log_dir(self) -> Path:
        return self.log_file.parent


def load_settings() -> Settings:
    """Load environment variables from .env and return a Settings object."""
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        load_dotenv(env_path)
    else:
        load_dotenv()

    settings = Settings(
        ai_provider=os.getenv("AI_PROVIDER", "gemini").strip().lower(),
        openai_api_key=os.getenv("OPENAI_API_KEY", "").strip(),
        openai_model=os.getenv("OPENAI_MODEL", "gpt-5-mini").strip(),
        openai_reasoning_effort=os.getenv("OPENAI_REASONING_EFFORT", "medium").strip(),
        gemini_api_key=os.getenv("GEMINI_API_KEY", "").strip(),
        gemini_model=os.getenv("GEMINI_MODEL", "gemini-2.5-flash").strip(),
        news_provider=os.getenv("NEWS_PROVIDER", "newsdata").strip().lower(),
        news_api_key=os.getenv("NEWS_API_KEY", "").strip(),
        newsdata_api_key=os.getenv("NEWSDATA_API_KEY", "").strip(),
        newsdata_country=os.getenv("NEWSDATA_COUNTRY", "").strip(),
        newsdata_categories=os.getenv("NEWSDATA_CATEGORIES", "business,politics").strip(),
        newsdata_query=os.getenv("NEWSDATA_QUERY", "").strip(),
        newsdata_query_in_title=os.getenv("NEWSDATA_QUERY_IN_TITLE", "").strip(),
        news_language=os.getenv("NEWS_LANGUAGE", "en").strip(),
        fetch_window_hours=_read_int("FETCH_WINDOW_HOURS", 72),
        max_articles_to_fetch=_read_int("MAX_ARTICLES_TO_FETCH", 40),
        top_articles_to_keep=_read_int("TOP_ARTICLES_TO_KEEP", 2),
        posting_mode=os.getenv("POSTING_MODE", "auto").strip().lower(),
        linkedin_access_token=os.getenv("LINKEDIN_ACCESS_TOKEN", "").strip(),
        linkedin_person_urn=os.getenv("LINKEDIN_PERSON_URN", "").strip(),
        linkedin_person_id=os.getenv("LINKEDIN_PERSON_ID", "").strip(),
        linkedin_client_id=os.getenv("LINKEDIN_CLIENT_ID", "").strip(),
        linkedin_client_secret=os.getenv("LINKEDIN_CLIENT_SECRET", "").strip(),
        linkedin_redirect_uri=os.getenv("LINKEDIN_REDIRECT_URI", "").strip(),
        linkedin_email=os.getenv("LINKEDIN_EMAIL", "").strip(),
        linkedin_password=os.getenv("LINKEDIN_PASSWORD", "").strip(),
        linkedin_headless=_read_bool("LINKEDIN_HEADLESS", False),
        chrome_binary_path=os.getenv("CHROME_BINARY_PATH", "").strip(),
        normal_post_interval_days=_read_int("NORMAL_POST_INTERVAL_DAYS", 3),
        major_event_window_hours=_read_int("MAJOR_EVENT_WINDOW_HOURS", 24),
        check_interval_minutes=_read_int("CHECK_INTERVAL_MINUTES", 60),
        request_timeout_seconds=_read_int("REQUEST_TIMEOUT_SECONDS", 25),
        state_file=BASE_DIR / "data" / "state.json",
        run_history_file=BASE_DIR / "data" / "run_history.jsonl",
        generated_posts_dir=BASE_DIR / "data" / "generated_posts",
        log_file=BASE_DIR / "logs" / "agent.log",
    )
    ensure_runtime_directories(settings)
    return settings


def ensure_runtime_directories(settings: Settings) -> None:
    """Create runtime folders the first time the project runs."""
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    settings.generated_posts_dir.mkdir(parents=True, exist_ok=True)
    settings.log_dir.mkdir(parents=True, exist_ok=True)


def setup_logging(settings: Settings) -> logging.Logger:
    """Create a reusable logger for the whole application."""
    logger = logging.getLogger("linkedin_macro_agent")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = RotatingFileHandler(
        settings.log_file,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.propagate = False
    return logger


def get_logger(name: str) -> logging.Logger:
    """Return a namespaced child logger."""
    return logging.getLogger(f"linkedin_macro_agent.{name}")
