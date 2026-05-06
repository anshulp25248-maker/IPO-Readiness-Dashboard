from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Sequence

try:
    from google import genai
except ImportError:  # pragma: no cover - depends on optional local install state
    genai = None

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - depends on optional local install state
    OpenAI = None

from analyzer import ScoredArticle
from config import Settings, get_logger


@dataclass(slots=True)
class GeneratedPost:
    """Normalized output from the content generation step."""

    post_text: str
    word_count: int
    hashtags: list[str]
    topic_signature: str
    model_used: str
    used_fallback: bool
    structured_payload: dict[str, Any]


class MacroContentGenerator:
    """Generates analyst-style LinkedIn content with Gemini or OpenAI, plus a safe fallback."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("content_generator")
        self.gemini_client = self._build_gemini_client()
        self.openai_client = self._build_openai_client()

    def generate_post(self, selected_articles: Sequence[ScoredArticle]) -> GeneratedPost:
        """Generate a formatted LinkedIn post from the top selected stories."""
        if not selected_articles:
            raise ValueError("At least one selected article is required to generate a post.")

        for provider in self._provider_order():
            try:
                if provider == "gemini" and self.gemini_client:
                    return self._generate_with_gemini(selected_articles)
                if provider == "openai" and self.openai_client:
                    return self._generate_with_openai(selected_articles)
            except Exception as exc:
                self.logger.warning("%s generation failed. Trying the next option if available: %s", provider.title(), exc)

        return self._generate_fallback(selected_articles)

    def _build_gemini_client(self):
        """Create the Gemini client if the SDK and API key are available."""
        if genai is None or not self.settings.gemini_api_key:
            return None
        return genai.Client(api_key=self.settings.gemini_api_key)

    def _build_openai_client(self):
        """Create the OpenAI client if the SDK and API key are available."""
        if OpenAI is None or not self.settings.openai_api_key:
            return None
        return OpenAI(api_key=self.settings.openai_api_key)

    def _provider_order(self) -> list[str]:
        """Decide which AI provider to try first."""
        provider = self.settings.ai_provider.lower()
        if provider == "gemini":
            return ["gemini"]
        if provider == "openai":
            return ["openai"]
        if provider == "auto":
            return ["gemini", "openai"]
        return ["gemini", "openai"]

    def _generate_with_gemini(self, selected_articles: Sequence[ScoredArticle]) -> GeneratedPost:
        """Use the official Gemini SDK with structured JSON output."""
        response = self.gemini_client.models.generate_content(
            model=self.settings.gemini_model,
            contents=self._compose_generation_prompt(selected_articles),
            config={
                "response_mime_type": "application/json",
                "response_json_schema": self._response_schema(),
            },
        )
        payload = json.loads(self._extract_gemini_text(response))
        post = self._build_post_from_payload(payload, self.settings.gemini_model)

        if not self._passes_constraints(post):
            payload = self._repair_with_gemini(payload, selected_articles, post.word_count)
            post = self._build_post_from_payload(payload, self.settings.gemini_model)

        return post

    def _repair_with_gemini(
        self,
        payload: dict[str, Any],
        selected_articles: Sequence[ScoredArticle],
        current_word_count: int,
    ) -> dict[str, Any]:
        """Ask Gemini for one repair pass when the first output misses constraints."""
        response = self.gemini_client.models.generate_content(
            model=self.settings.gemini_model,
            contents=self._compose_repair_prompt(payload, selected_articles, current_word_count),
            config={
                "response_mime_type": "application/json",
                "response_json_schema": self._response_schema(),
            },
        )
        return json.loads(self._extract_gemini_text(response))

    def _generate_with_openai(self, selected_articles: Sequence[ScoredArticle]) -> GeneratedPost:
        """Use the OpenAI Responses API to generate structured post content."""
        request_payload: dict[str, Any] = {
            "model": self.settings.openai_model,
            "instructions": self._developer_instructions(),
            "input": self._build_user_prompt(selected_articles),
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "linkedin_macro_post",
                    "schema": self._response_schema(),
                    "strict": True,
                }
            },
            "max_output_tokens": 1_200,
        }

        if self.settings.openai_model.startswith("gpt-5"):
            request_payload["reasoning"] = {"effort": self.settings.openai_reasoning_effort}

        response = self.openai_client.responses.create(**request_payload)
        payload = json.loads(self._extract_openai_text(response))
        post = self._build_post_from_payload(payload, self.settings.openai_model)

        if not self._passes_constraints(post):
            payload = self._repair_with_openai(payload, selected_articles, post.word_count)
            post = self._build_post_from_payload(payload, self.settings.openai_model)

        return post

    def _repair_with_openai(
        self,
        payload: dict[str, Any],
        selected_articles: Sequence[ScoredArticle],
        current_word_count: int,
    ) -> dict[str, Any]:
        """Ask OpenAI for one repair pass if the first output misses constraints."""
        request_payload: dict[str, Any] = {
            "model": self.settings.openai_model,
            "instructions": self._developer_instructions(),
            "input": self._compose_repair_prompt(payload, selected_articles, current_word_count),
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "linkedin_macro_post_repair",
                    "schema": self._response_schema(),
                    "strict": True,
                }
            },
            "max_output_tokens": 1_200,
        }

        if self.settings.openai_model.startswith("gpt-5"):
            request_payload["reasoning"] = {"effort": self.settings.openai_reasoning_effort}

        response = self.openai_client.responses.create(**request_payload)
        return json.loads(self._extract_openai_text(response))

    def _build_post_from_payload(self, payload: dict[str, Any], model_used: str) -> GeneratedPost:
        hashtags = self._normalize_hashtags(payload.get("hashtags", []))
        topic_signature = (payload.get("topic_signature") or "").strip()

        post_text = "\n".join(
            [
                "\U0001F6A8 Hook",
                payload.get("hook", "").strip(),
                "",
                payload.get("main_explanation", "").strip(),
                "",
                "\U0001F4CA What this really means:",
                payload.get("what_this_really_means", "").strip(),
                "",
                "\U0001F4AC Closing question",
                payload.get("closing_question", "").strip(),
                "",
                " ".join(hashtags),
            ]
        ).strip()

        return GeneratedPost(
            post_text=post_text,
            word_count=len(post_text.split()),
            hashtags=hashtags,
            topic_signature=topic_signature,
            model_used=model_used,
            used_fallback=False,
            structured_payload=payload,
        )

    def _generate_fallback(self, selected_articles: Sequence[ScoredArticle]) -> GeneratedPost:
        """Create a usable deterministic post if all AI providers fail."""
        lead = selected_articles[0]
        transmission_channel = self._guess_transmission_channel(lead)
        root_cause = self._guess_root_cause(lead)
        hook = (
            "This is not just another macro headline.\n"
            "It is a signal that the growth, inflation, and policy trade-off is shifting again."
        )

        second_story_line = ""
        if len(selected_articles) > 1:
            second_story_line = (
                f" The second high-scoring story also points in the same direction: "
                f"{selected_articles[1].article.title}."
            )

        main_explanation = (
            f"Root cause: {root_cause}. The primary transmission channel is {transmission_channel} "
            "\u2192 that matters because shocks rarely stay contained in one asset class. "
            "If the move raises input costs or tightens liquidity, inflation stays sticky while growth expectations soften. "
            "That is exactly the setup that keeps central banks cautious even when investors want quick easing. "
            "For markets, equities face earnings-reset risk, bonds can reprice around the rate path, commodities react first, "
            "and volatility usually rises before the macro data fully catches up."
            f"{second_story_line}"
        )

        what_this_really_means = (
            "The second-order effects are where the real signal sits \u2192 demand destruction, tougher policy trade-offs, "
            "and tighter financial conditions. For India, that can mean INR pressure, imported inflation risk, "
            "and a more delicate RBI balancing act between defending credibility and protecting growth."
        )

        closing_question = (
            "If this shock stays in the system for another quarter, will policymakers stabilize markets, "
            "or become the next source of volatility?"
        )

        payload = {
            "hook": hook,
            "main_explanation": main_explanation,
            "what_this_really_means": what_this_really_means,
            "closing_question": closing_question,
            "hashtags": self._pick_fallback_hashtags(lead),
            "topic_signature": lead.topic_signature,
        }

        post = self._build_post_from_payload(payload, "template-fallback")
        post.used_fallback = True
        return post

    @staticmethod
    def _response_schema() -> dict[str, Any]:
        """Return the common JSON schema shared across providers."""
        return {
            "type": "object",
            "properties": {
                "hook": {
                    "type": "string",
                    "description": "Exactly two punchy lines that hook a finance audience.",
                },
                "main_explanation": {
                    "type": "string",
                    "description": (
                        "Short paragraphs that cover root cause, transmission channel, macro impact, "
                        "market impact, currency impact, second-order effects, and India angle."
                    ),
                },
                "what_this_really_means": {
                    "type": "string",
                    "description": "A concise synthesis section starting from implications, not summary.",
                },
                "closing_question": {
                    "type": "string",
                    "description": "One thoughtful closing question that ends with a question mark.",
                },
                "hashtags": {
                    "type": "array",
                    "items": {"type": "string"},
                    "minItems": 3,
                    "maxItems": 5,
                },
                "topic_signature": {
                    "type": "string",
                    "description": "A short fingerprint of the core topic for memory storage.",
                },
            },
            "required": [
                "hook",
                "main_explanation",
                "what_this_really_means",
                "closing_question",
                "hashtags",
                "topic_signature",
            ],
            "additionalProperties": False,
        }

    def _compose_generation_prompt(self, selected_articles: Sequence[ScoredArticle]) -> str:
        """Combine style instructions and article context for Gemini or repair calls."""
        return "\n\n".join(
            [
                self._developer_instructions(),
                self._build_user_prompt(selected_articles),
            ]
        )

    def _compose_repair_prompt(
        self,
        payload: dict[str, Any],
        selected_articles: Sequence[ScoredArticle],
        current_word_count: int,
    ) -> str:
        """Create a single repair prompt that both Gemini and OpenAI can use."""
        return (
            f"{self._developer_instructions()}\n\n"
            "Your previous LinkedIn post missed one or more constraints.\n"
            f"Current word count: {current_word_count}\n"
            "Rewrite the post so the final formatted result lands between 150 and 250 words, "
            "keeps a strong hook, preserves the macro logic, ends with a question, and uses 3 to 5 hashtags.\n\n"
            "Original structured output:\n"
            f"{json.dumps(payload, ensure_ascii=False, indent=2)}\n\n"
            "Source article context:\n"
            f"{self._build_user_prompt(selected_articles)}"
        )

    def _developer_instructions(self) -> str:
        """System-style instructions that shape the final writing style."""
        return (
            "You are a senior macro strategist writing LinkedIn posts for finance professionals.\n"
            "Every post must think like a macro analyst, not a news summarizer.\n"
            "Mandatory reasoning structure:\n"
            "1. Root cause\n"
            "2. Primary transmission channel\n"
            "3. Macro impact: inflation, growth, rates\n"
            "4. Market impact: equities, bonds, commodities, volatility\n"
            "5. Currency impact: USD and emerging markets\n"
            "6. Second-order effects\n"
            "7. India angle: INR and RBI\n\n"
            "Writing rules:\n"
            "- The first two lines must be a strong hook.\n"
            "- Use arrows like \u2192 to show cause and effect.\n"
            "- Use short paragraphs.\n"
            "- Sound like an investment banking or macro analyst teaching an MBA student.\n"
            "- No fluff, no generic motivation, no emojis except the section labels already added later.\n"
            "- Keep the final formatted result between 150 and 250 words.\n"
            "- End with one thought-provoking question.\n"
            "- Return 3 to 5 relevant hashtags.\n"
            "- Do not invent numerical data or pretend certainty when the source context does not justify it."
        )

    def _build_user_prompt(self, selected_articles: Sequence[ScoredArticle]) -> str:
        """Serialize article context and score reasoning into the user prompt."""
        context_payload = [
            {
                "title": item.article.title,
                "description": item.article.description,
                "url": item.article.url,
                "published_at": item.article.published_at.isoformat(),
                "source": item.article.source,
                "score": item.score,
                "matched_keywords": item.matched_keywords,
                "score_reason": item.reason,
                "is_major_event": item.is_major_event,
                "topic_signature": item.topic_signature,
            }
            for item in selected_articles
        ]

        return (
            "Use the following macro-relevant stories to craft one high-quality LinkedIn post.\n"
            "Treat them as signals to interpret, not headlines to repeat.\n\n"
            f"{json.dumps(context_payload, indent=2, ensure_ascii=False)}"
        )

    @staticmethod
    def _extract_gemini_text(response: Any) -> str:
        """Pull JSON text from the Gemini response object."""
        text = getattr(response, "text", "")
        if text:
            return text
        raise ValueError("Could not extract text from the Gemini response.")

    @staticmethod
    def _extract_openai_text(response: Any) -> str:
        """Pull text from the OpenAI response object in a defensive way."""
        output_text = getattr(response, "output_text", None)
        if output_text:
            return output_text

        output = getattr(response, "output", [])
        fragments: list[str] = []
        for item in output:
            for content in getattr(item, "content", []):
                text = getattr(content, "text", "")
                if text:
                    fragments.append(text)

        if fragments:
            return "\n".join(fragments)

        raise ValueError("Could not extract text from the OpenAI response.")

    @staticmethod
    def _normalize_hashtags(raw_hashtags: list[str]) -> list[str]:
        """Ensure hashtags are clean, unique, and always start with #."""
        normalized: list[str] = []
        seen: set[str] = set()
        for hashtag in raw_hashtags:
            cleaned = hashtag.strip().replace(" ", "")
            if not cleaned:
                continue
            if not cleaned.startswith("#"):
                cleaned = f"#{cleaned.lstrip('#')}"
            cleaned = cleaned.replace("##", "#")
            lower = cleaned.lower()
            if lower in seen:
                continue
            seen.add(lower)
            normalized.append(cleaned)
            if len(normalized) >= 5:
                break

        while len(normalized) < 3:
            normalized.append(f"#MacroStrategy{len(normalized) + 1}")

        return normalized[:5]

    @staticmethod
    def _passes_constraints(post: GeneratedPost) -> bool:
        return 150 <= post.word_count <= 250 and 3 <= len(post.hashtags) <= 5 and "?" in post.post_text

    @staticmethod
    def _guess_transmission_channel(article: ScoredArticle) -> str:
        text = article.article.combined_text().lower()
        if any(term in text for term in ("oil", "crude", "opec", "shipping", "sanctions")):
            return "energy prices and trade costs"
        if any(term in text for term in ("fed", "ecb", "rate hike", "interest rate", "bond")):
            return "global liquidity and discount rates"
        if any(term in text for term in ("war", "conflict", "attack", "missile")):
            return "risk sentiment, supply disruption, and safe-haven flows"
        return "confidence, capital flows, and pricing power"

    @staticmethod
    def _guess_root_cause(article: ScoredArticle) -> str:
        text = article.article.combined_text().lower()
        if "inflation" in text:
            return "a price shock that risks forcing tighter policy for longer"
        if any(term in text for term in ("war", "conflict", "attack")):
            return "a geopolitical shock that can spill into trade, energy, and capital markets"
        if any(term in text for term in ("rate hike", "interest rate", "fed", "rbi", "ecb")):
            return "a policy shift that changes global financing conditions"
        return "a macro shock with cross-asset implications"

    @staticmethod
    def _pick_fallback_hashtags(article: ScoredArticle) -> list[str]:
        text = article.article.combined_text().lower()
        hashtags = ["#GlobalMacro", "#Finance", "#Geopolitics"]
        if "inflation" in text:
            hashtags.append("#Inflation")
        elif any(term in text for term in ("rate hike", "interest rate", "fed", "rbi", "ecb")):
            hashtags.append("#CentralBanks")
        elif any(term in text for term in ("war", "conflict", "sanctions")):
            hashtags.append("#RiskManagement")
        else:
            hashtags.append("#Markets")
        return hashtags[:5]
