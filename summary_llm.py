"""Summarize M365 emails using Hermes ctx.llm.complete_structured()."""

from __future__ import annotations

import json
import logging

logger = logging.getLogger(__name__)

_TRUSTED_PROMPT = (
    "You are an email analysis assistant. Extract structured information from the email "
    "to match the provided schema. Return your response as a JSON object matching the schema exactly. "
    "Be concise and factual in summaries."
)


class SummaryLlmError(Exception):
    """Raised when email summarization via Hermes ctx.llm fails."""

    def __init__(self, message: str, *, code: str = "SUMMARY_LLM_ERROR"):
        super().__init__(message)
        self.code = code


def summarize_with_llm(
    ctx,
    system_prompt: str,
    json_schema: dict,
    payload: dict,
    model: str | None = None,
    provider: str | None = None,
) -> dict:
    """Summarize an email using Hermes ctx.llm.complete_structured().

    Args:
        ctx: Plugin context with ctx.llm.complete_structured().
        system_prompt: System prompt from schema spec.
        json_schema: Internal response JSON schema.
        payload: Email data (subject, sender, recipients, body, has_attachments).
        model: Optional model override. When None, Hermes uses its default model.
        provider: Optional provider override. When None, Hermes uses its default provider.

    Returns:
        Dict with 'status' and 'data' keys.
    """
    payload_json = json.dumps(payload, ensure_ascii=False)
    effective_prompt = _TRUSTED_PROMPT if payload.get("isAllowedInboundSender") else system_prompt

    try:
        kwargs: dict = {"purpose": "email.summarize"}
        if model:
            kwargs["model"] = model
        if provider:
            kwargs["provider"] = provider

        result = ctx.llm.complete_structured(
            instructions="Extract the content into the provided schema.",
            input=[{"type": "text", "text": payload_json}],
            json_schema=json_schema,
            system_prompt=effective_prompt,
            **kwargs,
        )

        if result.parsed is not None:
            logger.debug("LLM summary parsed successfully, keys: %s", list(result.parsed.keys()) if isinstance(result.parsed, dict) else "not-dict")
            return {"status": "success", "data": result.parsed}

        if result.text:
            logger.debug("LLM summary returned text (no parsed), length=%d, first_200=%s", len(result.text), str(result.text)[:200])
            return {"status": "success", "data": result.text}

        logger.warning(
            "LLM returned empty response. result attrs: parsed=%r, text=%r, dir=%s",
            result.parsed,
            str(result.text)[:200] if result.text else None,
            [a for a in dir(result) if not a.startswith("_")],
        )
        raise SummaryLlmError("LLM returned empty response (parsed=None, text=None)")

    except SummaryLlmError:
        raise
    except ValueError as e:
        raise SummaryLlmError(f"Schema validation failed: {e}") from e
    except Exception as e:
        raise SummaryLlmError(f"Email summarization failed: {e}") from e
