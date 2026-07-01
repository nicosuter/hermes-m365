"""Summarize M365 emails using Hermes ctx.llm.acomplete_structured()."""

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


async def summarize_with_llm(
    ctx,
    system_prompt: str,
    json_schema: dict,
    payload: dict,
    model: str | None = None,
    provider: str | None = None,
    timeout: float | None = None,
) -> dict:
    """Summarize an email using Hermes ctx.llm.acomplete_structured().

    Uses the async variant (acomplete_structured) to avoid blocking the event
    loop — this function is called from async tool handlers (get_summary).
    The sync complete_structured uses a blocking HTTP client that can corrupt
    the async httpx connection pool, causing spurious ConnectionError.

    Args:
        ctx: Plugin context with ctx.llm.acomplete_structured().
        system_prompt: System prompt from schema spec.
        json_schema: Internal response JSON schema.
        payload: Email data (subject, sender, recipients, body, has_attachments).
        model: Optional model override. When None, Hermes uses its default model.
        provider: Optional provider override. When None, Hermes uses its default provider.
        timeout: Timeout in seconds for the LLM call. Defaults to 120s.
            The Hermes auxiliary_client default is 30s, which is too short
            for email summarization of large bodies.

    Returns:
        Dict with 'status' and 'data' keys.
    """
    payload_json = json.dumps(payload, ensure_ascii=False)
    effective_prompt = _TRUSTED_PROMPT if payload.get("isAllowedInboundSender") else system_prompt

    call_kwargs: dict = {"purpose": "email.summarize"}
    if model:
        call_kwargs["model"] = model
    if provider:
        call_kwargs["provider"] = provider
    if timeout is not None:
        call_kwargs["timeout"] = timeout

    def _try_parse(text: str) -> dict | None:
        try:
            parsed = json.loads(text)
            return parsed if isinstance(parsed, dict) else None
        except (json.JSONDecodeError, ValueError):
            return None

    async def _make_call(schema: dict | None, use_json_mode: bool) -> dict:
        result = await ctx.llm.acomplete_structured(
            instructions="Extract the content into the provided schema. Return your response as a JSON object matching the schema exactly.",
            input=[{"type": "text", "text": payload_json}],
            json_schema=schema,
            json_mode=use_json_mode,
            system_prompt=effective_prompt,
            **call_kwargs,
        )
        if result.parsed is not None:
            logger.debug("LLM summary parsed successfully, keys: %s", list(result.parsed.keys()) if isinstance(result.parsed, dict) else "not-dict")
            return {"status": "success", "data": result.parsed}
        if result.text:
            parsed_text = _try_parse(result.text)
            if parsed_text is not None:
                logger.debug("LLM summary parsed from text, keys: %s", list(parsed_text.keys()))
                return {"status": "success", "data": parsed_text}
        return {"status": "fail", "data": None}

    async def _wrap_call(schema: dict | None, use_json_mode: bool) -> dict:
        try:
            return await _make_call(schema, use_json_mode)
        except SummaryLlmError:
            raise
        except ValueError as e:
            raise SummaryLlmError(f"SchemaValidationError: {e}") from e
        except Exception as e:
            raise SummaryLlmError(f"Email summarization failed: {e}") from e

    result = await _wrap_call(json_schema, False)
    if result["status"] == "success":
        return result

    raise SummaryLlmError(
        "LLM returned non-JSON response for structured output. "
        "Model may not support structured output."
    )
