"""Tests for Hermes ctx.llm summary wrapper (summary_llm.py).

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnusedCallResult=false
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any

import pytest

from summary_llm import SummaryLlmError, _TRUSTED_PROMPT, summarize_with_llm


def _internal_schema() -> dict[str, object]:
    from summary_schema import build_internal_response_schema
    user_schema = {
        "type": "object",
        "properties": {"topic": {"type": "string"}},
        "required": ["topic"],
        "additionalProperties": False,
    }
    return build_internal_response_schema(user_schema)


_PAYLOAD: dict[str, object] = {"subject": "Meeting notes", "body": "Discussed Q3."}


# ---------------------------------------------------------------------------
# Error class
# ---------------------------------------------------------------------------

def test_error_class_attributes():
    err = SummaryLlmError("oops", code="SUMMARY_REFUSED")
    assert err.code == "SUMMARY_REFUSED"
    assert str(err) == "oops"
    assert isinstance(err, Exception)


def test_error_default_code():
    err = SummaryLlmError("something broke")
    assert err.code == "SUMMARY_LLM_ERROR"


# ---------------------------------------------------------------------------
# Success cases
# ---------------------------------------------------------------------------

async def test_success_with_parsed_result():
    """ctx.llm.acomplete_structured returns parsed result."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "Q3"}})

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="You are an email summarizer.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
    )
    assert result["status"] == "success"
    assert result["data"]["result"]["topic"] == "Q3"


async def test_success_with_text_fallback():
    """When parsed is None but text is valid JSON, parse it."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed=None, text='{"status": "ok", "result": {"topic": "Q3"}}')

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="You are an email summarizer.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
    )
    assert result["status"] == "success"
    assert result["data"]["result"]["topic"] == "Q3"


async def test_success_with_model_override():
    """Model param is passed through to ctx.llm.acomplete_structured()."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="Summarize.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
        model="gpt-4o-mini",
    )
    assert result["status"] == "success"
    assert mock_ctx._last_model == "gpt-4o-mini"


async def test_success_without_model():
    """When model is None, it is not passed to ctx.llm."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="Summarize.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
    )
    assert result["status"] == "success"
    assert mock_ctx._last_model is None


async def test_success_with_provider_override():
    """Provider param is passed through to ctx.llm.acomplete_structured()."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="Summarize.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
        provider="openai",
    )
    assert result["status"] == "success"
    assert mock_ctx._last_provider == "openai"


async def test_success_without_provider():
    """When provider is None, it is not passed to ctx.llm."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    result = await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="Summarize.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
    )
    assert result["status"] == "success"
    assert mock_ctx._last_provider is None


# ---------------------------------------------------------------------------
# Error cases
# ---------------------------------------------------------------------------

async def test_empty_response_raises():
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed=None, text=None)

    with pytest.raises(SummaryLlmError) as exc:
        await summarize_with_llm(
            ctx=mock_ctx,
            system_prompt="Summarize.",
            json_schema=_internal_schema(),
            payload=_PAYLOAD,
        )
    assert exc.value.code == "SUMMARY_LLM_ERROR"
    assert "non-JSON" in str(exc.value)
    assert mock_ctx.llm._call_count == 1


async def test_invalid_text_response_raises_without_fallback():
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed=None, text="not json")

    with pytest.raises(SummaryLlmError) as exc:
        await summarize_with_llm(
            ctx=mock_ctx,
            system_prompt="Summarize.",
            json_schema=_internal_schema(),
            payload=_PAYLOAD,
        )
    assert exc.value.code == "SUMMARY_LLM_ERROR"
    assert "non-JSON" in str(exc.value)
    assert mock_ctx.llm._call_count == 1


async def test_value_error_raises():
    """ValueError from ctx.llm raises SummaryLlmError."""
    mock_ctx = _FakeCtx()
    mock_ctx._exc = ValueError("schema validation failed")

    with pytest.raises(SummaryLlmError) as exc:
        await summarize_with_llm(
            ctx=mock_ctx,
            system_prompt="Summarize.",
            json_schema=_internal_schema(),
            payload=_PAYLOAD,
        )
    assert exc.value.code == "SUMMARY_LLM_ERROR"
    assert "SchemaValidationError" in str(exc.value)


async def test_generic_exception_raises():
    """Generic exception from ctx.llm raises SummaryLlmError."""
    mock_ctx = _FakeCtx()
    mock_ctx._exc = ConnectionError("connection refused")

    with pytest.raises(SummaryLlmError) as exc:
        await summarize_with_llm(
            ctx=mock_ctx,
            system_prompt="Summarize.",
            json_schema=_internal_schema(),
            payload=_PAYLOAD,
        )
    assert exc.value.code == "SUMMARY_LLM_ERROR"
    assert "Email summarization failed" in str(exc.value)


async def test_timeout_raises():
    """TimeoutError from ctx.llm raises SummaryLlmError."""
    mock_ctx = _FakeCtx()
    mock_ctx._exc = TimeoutError("timed out")

    with pytest.raises(SummaryLlmError) as exc:
        await summarize_with_llm(
            ctx=mock_ctx,
            system_prompt="Summarize.",
            json_schema=_internal_schema(),
            payload=_PAYLOAD,
        )
    assert exc.value.code == "SUMMARY_LLM_ERROR"


async def test_instructions_and_payload_passed():
    """Instructions and payload are correctly formatted."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="You are a summarizer.",
        json_schema=_internal_schema(),
        payload=_PAYLOAD,
    )

    assert "Extract the content into the provided schema" in mock_ctx._last_instructions
    input_blocks = mock_ctx._last_input
    assert isinstance(input_blocks, list) and len(input_blocks) == 1
    assert input_blocks[0]["type"] == "text"
    assert "Meeting notes" in input_blocks[0]["text"]
    assert mock_ctx._last_system_prompt == "You are a summarizer."
    assert mock_ctx._last_purpose == "email.summarize"


async def test_trusted_sender_uses_trusted_prompt():
    """Trusted senders get _TRUSTED_PROMPT instead of schema prompt."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    payload = {**_PAYLOAD, "isAllowedInboundSender": True}
    await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="You are a data extraction parser.",
        json_schema=_internal_schema(),
        payload=payload,
    )

    assert mock_ctx._last_system_prompt == _TRUSTED_PROMPT


async def test_untrusted_sender_uses_schema_prompt():
    """Untrusted senders get the schema system prompt (anti-injection language)."""
    mock_ctx = _FakeCtx()
    mock_ctx._result = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {"topic": "test"}})

    payload = {**_PAYLOAD, "isAllowedInboundSender": False}
    await summarize_with_llm(
        ctx=mock_ctx,
        system_prompt="You are a data extraction parser.",
        json_schema=_internal_schema(),
        payload=payload,
    )

    assert mock_ctx._last_system_prompt == "You are a data extraction parser."


# ---------------------------------------------------------------------------
# Mock infrastructure
# ---------------------------------------------------------------------------

@dataclass
class _FakeStructuredResult:
    parsed: Any = None
    text: str | None = None
    content_type: str = "json"


class _FakeLlm:
    def __init__(self, ctx: _FakeCtx) -> None:
        self._ctx = ctx
        self._call_count = 0

    async def acomplete_structured(
        self,
        instructions: str,
        input: Any,
        json_schema: dict | None = None,
        system_prompt: str | None = None,
        model: str | None = None,
        provider: str | None = None,
        purpose: str | None = None,
        json_mode: bool = False,
        **kwargs: Any,
    ) -> _FakeStructuredResult:
        self._ctx._last_instructions = instructions
        self._ctx._last_input = input
        self._last_json_schema = json_schema
        self._ctx._last_system_prompt = system_prompt
        self._ctx._last_model = model
        self._ctx._last_provider = provider
        self._ctx._last_purpose = purpose
        self._ctx._last_json_mode = json_mode
        self._call_count += 1

        if self._ctx._exc:
            raise self._ctx._exc
        if isinstance(self._ctx._result, list):
            idx = min(self._call_count - 1, len(self._ctx._result) - 1)
            return self._ctx._result[idx]
        return self._ctx._result


class _FakeCtx:
    def __init__(self) -> None:
        self._result: Any = _FakeStructuredResult(parsed={"status": "ok", "reason": None, "result": {}})
        self._exc: Exception | None = None
        self._last_instructions: str = ""
        self._last_input: Any = ""
        self._last_system_prompt: str | None = None
        self._last_model: str | None = None
        self._last_provider: str | None = None
        self._last_purpose: str | None = None
        self._last_json_mode: bool = False
        self.llm = _FakeLlm(self)
