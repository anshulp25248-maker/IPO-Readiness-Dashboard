from __future__ import annotations

import time

import schedule

from config import load_settings, setup_logging
from main import run_pipeline


def _scheduled_job() -> None:
    """Run one scheduler cycle and keep the process alive even if one run fails."""
    settings = load_settings()
    logger = setup_logging(settings)

    try:
        result = run_pipeline(force=False, dry_run=False, skip_posting=False, preferred_posting_mode=None)
        logger.info("Scheduler cycle finished with status=%s reason=%s", result.get("status"), result.get("reason"))
    except Exception as exc:
        logger.exception("Scheduler cycle failed: %s", exc)


def main() -> None:
    settings = load_settings()
    logger = setup_logging(settings)

    logger.info(
        "Starting scheduler. The agent will check for posting opportunities every %s minutes.",
        settings.check_interval_minutes,
    )

    schedule.every(settings.check_interval_minutes).minutes.do(_scheduled_job)
    _scheduled_job()

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    main()
