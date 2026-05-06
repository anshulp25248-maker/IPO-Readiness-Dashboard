from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from analyzer import MacroNewsAnalyzer
from config import load_settings, setup_logging
from content_generator import MacroContentGenerator
from linkedin_poster import LinkedInAPIPoster, LinkedInPostingManager, PostingResult
from memory import MemoryStore
from news_fetcher import NewsFetcher


def run_pipeline(
    *,
    force: bool = False,
    dry_run: bool = False,
    skip_posting: bool = False,
    preferred_posting_mode: str | None = None,
) -> dict[str, Any]:
    """Run one end-to-end cycle: fetch, score, generate, and optionally post."""
    settings = load_settings()
    logger = setup_logging(settings)

    memory = MemoryStore(settings)
    fetcher = NewsFetcher(settings)
    analyzer = MacroNewsAnalyzer(settings)
    generator = MacroContentGenerator(settings)
    posting_manager = LinkedInPostingManager(settings)

    logger.info("Starting pipeline run. force=%s dry_run=%s skip_posting=%s", force, dry_run, skip_posting)

    try:
        fetched_articles = fetcher.fetch_latest_news()
        if not fetched_articles:
            reason = "No news articles were returned from the configured providers."
            memory.record_run(status="skipped", reason=reason)
            return {"status": "skipped", "reason": reason, "selected_articles": []}

        selected_articles = analyzer.select_top_articles(fetched_articles, memory)
        decision_to_post, decision_reason = analyzer.should_post_now(selected_articles, memory, force=force)

        if not selected_articles:
            memory.record_run(status="skipped", reason=decision_reason, selected_articles=[])
            return {
                "status": "skipped",
                "reason": decision_reason,
                "selected_articles": [],
            }

        should_generate_preview = decision_to_post or dry_run or skip_posting or force
        if not should_generate_preview:
            selected_payload = [item.to_dict() for item in selected_articles]
            memory.record_run(status="skipped", reason=decision_reason, selected_articles=selected_payload)
            return {
                "status": "skipped",
                "reason": decision_reason,
                "selected_articles": selected_payload,
            }

        generated_post = generator.generate_post(selected_articles)
        generated_post_path = _save_generated_post(
            output_dir=settings.generated_posts_dir,
            topic_signature=generated_post.topic_signature or "macro-update",
            post_text=generated_post.post_text,
        )

        selected_payload = [item.to_dict() for item in selected_articles]

        posting_result = PostingResult(
            success=True,
            method="dry-run" if (dry_run or skip_posting or not decision_to_post) else "none",
            message="Post generation completed without publishing.",
        )
        status = "generated"

        if decision_to_post and not dry_run and not skip_posting:
            posting_result = posting_manager.post(
                generated_post.post_text,
                preferred_mode=preferred_posting_mode,
                dry_run=False,
            )
            status = "posted" if posting_result.success else "error"

        memory.record_run(
            status=status,
            reason=decision_reason,
            selected_articles=selected_payload,
            generated_post_text=generated_post.post_text,
            posting_method=posting_result.method,
            post_id=posting_result.post_id,
            used_fallback_generator=generated_post.used_fallback,
            error_message="" if posting_result.success else posting_result.message,
        )

        if posting_result.success and decision_to_post and not dry_run and not skip_posting:
            memory.record_post(
                topic_signature=generated_post.topic_signature,
                selected_articles=selected_payload,
                generated_post_text=generated_post.post_text,
                method=posting_result.method,
                post_id=posting_result.post_id,
                used_fallback_generator=generated_post.used_fallback,
            )

        return {
            "status": status,
            "reason": decision_reason,
            "selected_articles": selected_payload,
            "generated_post_path": str(generated_post_path),
            "post_text": generated_post.post_text,
            "word_count": generated_post.word_count,
            "posting_result": {
                "success": posting_result.success,
                "method": posting_result.method,
                "message": posting_result.message,
                "post_id": posting_result.post_id,
            },
            "model_used": generated_post.model_used,
            "used_fallback_generator": generated_post.used_fallback,
        }
    except Exception as exc:
        logger.exception("Pipeline failed: %s", exc)
        memory.record_run(status="error", reason="Unhandled exception during pipeline run.", error_message=str(exc))
        raise


def _save_generated_post(output_dir: Path, topic_signature: str, post_text: str) -> Path:
    """Save every generated post so the user can audit what was produced."""
    safe_slug = "-".join(topic_signature.split())[:60] or "macro-update"
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_path = output_dir / f"{timestamp}-{safe_slug}.txt"
    output_path.write_text(post_text, encoding="utf-8")
    return output_path


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="AI macro analyst + LinkedIn automation engine.",
    )
    parser.add_argument("--force", action="store_true", help="Ignore the normal 3-day cadence and post immediately.")
    parser.add_argument("--dry-run", action="store_true", help="Generate and print the post without sending it to LinkedIn.")
    parser.add_argument("--skip-posting", action="store_true", help="Generate the post and save it locally without publishing.")
    parser.add_argument(
        "--post-method",
        choices=["auto", "api", "selenium"],
        default=None,
        help="Choose how the post should be published.",
    )
    parser.add_argument(
        "--print-linkedin-auth-url",
        action="store_true",
        help="Print the LinkedIn OAuth URL for the API setup flow.",
    )
    parser.add_argument(
        "--exchange-linkedin-code",
        help="Exchange a LinkedIn OAuth authorization code for an access token.",
    )
    parser.add_argument(
        "--fetch-linkedin-profile",
        action="store_true",
        help="Fetch the authenticated LinkedIn profile to discover the person id / URN.",
    )
    return parser


def _format_cli_output(result: dict[str, Any]) -> str:
    lines = [
        f"Status: {result.get('status', 'unknown')}",
        f"Reason: {result.get('reason', 'n/a')}",
    ]

    posting_result = result.get("posting_result")
    if posting_result:
        lines.append(f"Posting method: {posting_result.get('method', 'n/a')}")
        lines.append(f"Posting message: {posting_result.get('message', 'n/a')}")

    if result.get("generated_post_path"):
        lines.append(f"Saved post: {result['generated_post_path']}")

    if result.get("post_text"):
        lines.append("")
        lines.append("Generated LinkedIn Post")
        lines.append(result["post_text"])

    return "\n".join(lines)


def main() -> None:
    args = _build_parser().parse_args()
    settings = load_settings()
    setup_logging(settings)
    api_poster = LinkedInAPIPoster(settings)

    if args.print_linkedin_auth_url:
        print(api_poster.build_authorization_url())
        return

    if args.exchange_linkedin_code:
        token_payload = api_poster.exchange_code_for_access_token(args.exchange_linkedin_code)
        print(json.dumps(token_payload, indent=2, ensure_ascii=False))
        return

    if args.fetch_linkedin_profile:
        profile = api_poster.fetch_member_profile()
        print(json.dumps(profile, indent=2, ensure_ascii=False))
        return

    result = run_pipeline(
        force=args.force,
        dry_run=args.dry_run,
        skip_posting=args.skip_posting,
        preferred_posting_mode=args.post_method,
    )
    print(_format_cli_output(result))


if __name__ == "__main__":
    main()
