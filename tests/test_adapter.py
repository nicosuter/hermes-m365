# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

"""Tests for M365 Email adapter lifecycle, polling, send, and state."""

import asyncio
import json
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, cast

import httpx
import pytest
import respx

from config import MailConfig
from graph import GRAPH_BASE_URL, GraphClient
from mail_tools import _encode_email_id
from state import PollState


# ── Helpers ──────────────────────────────────────────────────────────────

def _mock_token():
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "tok", "expires_in": 3600})
    )


def _mail_url(path: str) -> str:
    return f"{GRAPH_BASE_URL}/users/user%40example.org/{path}"


def _make_message(msg_id: str, sender: str, name: str = "", subject: str = "Test",
                  body: str = "Hello", received: str = "2026-06-17T12:00:00Z",
                  conversation_id: str = "conv-1", internet_message_id: str = "<msg@example.com>") -> dict:
    return {
        "id": msg_id,
        "subject": subject,
        "from": {"emailAddress": {"address": sender, "name": name}},
        "receivedDateTime": received,
        "body": {"contentType": "HTML", "content": body},
        "conversationId": conversation_id,
        "internetMessageId": internet_message_id,
        "hasAttachments": False,
    }


@pytest.fixture
def env(monkeypatch):
    monkeypatch.setenv("M365_MAIL_CLIENT_ID", "client-id")
    monkeypatch.setenv("M365_MAIL_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("M365_MAIL_TENANT_ID", "tenant-id")
    monkeypatch.setenv("M365_MAILBOX_USER", "user@example.org")
    monkeypatch.setenv("EMAIL_ALLOWED_USERS", "trusted@example.com")
    monkeypatch.setenv("M365_POLL_INTERVAL_SECONDS", "0.1")
    return monkeypatch


@pytest.fixture
def state_path(tmp_path: Path) -> Path:
    return tmp_path / ".runtime" / "poll-state.json"


# ── State Tests ────────────────────────────────────────────────────────────

def test_state_load_empty(tmp_path: Path):
    p = tmp_path / "state.json"
    s = PollState.load(p)
    assert s.watermark == ""
    assert not s.is_processed("msg-1")


def test_state_save_load_roundtrip(state_path: Path):
    s = PollState()
    s.add("msg-1", "2026-06-17T12:00:00Z")
    s.add("msg-2", "2026-06-17T13:00:00Z")
    s.save(state_path)

    s2 = PollState.load(state_path)
    assert s2.is_processed("msg-1")
    assert s2.is_processed("msg-2")
    assert not s2.is_processed("msg-3")
    assert s2.watermark == "2026-06-17T13:00:00Z"


def test_state_max_ids():
    s = PollState()
    for i in range(600):
        s.add(f"msg-{i}", f"2026-06-17T{i:04d}:00:00Z")
    assert len(s._processed_ids) <= 500
    assert not s.is_processed("msg-0")
    assert s.is_processed("msg-599")


# ── Adapter Lifecycle Tests ───────────────────────────────────────────────

@pytest.mark.asyncio
@pytest.mark.xfail(reason="send() is intentionally a stub; agent must use explicit email tools (send_email, reply_email, etc.)", strict=True)
@respx.mock
async def test_send_defaults_subject_when_no_metadata(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    await adapter.connect()

    send_route = respx.post(_mail_url("sendMail")).mock(
        return_value=httpx.Response(202, json={})
    )

    await adapter.send("m365:test@example.com", "plain body")

    payload = json.loads(send_route.calls.last.request.read().decode())
    assert payload["message"]["subject"] == "Message from Hermes"

    await adapter.disconnect()


@pytest.mark.asyncio
@respx.mock
async def test_send_returns_false_when_not_connected(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    result = await adapter.send("m365:test@example.com", "body")
    assert result.success is False


# ── get_chat_info Tests ───────────────────────────────────────────────────

def test_get_chat_info_with_email_prefix():
    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    info = adapter.get_chat_info("m365:USER@Example.COM")
    assert info["chat_id"] == "m365:user@example.com"
    assert info["chat_name"] == "user@example.com"
    assert info["chat_type"] == "dm"


def test_get_chat_info_without_email_prefix():
    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    info = adapter.get_chat_info("someone@example.com")
    assert info["chat_id"] == "m365:someone@example.com"
    assert info["chat_name"] == "someone@example.com"
    assert info["chat_type"] == "dm"


# ── Polling does NOT mark mail read ────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_polling_does_not_mark_mail_read(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker_calls: list[dict] = []
    async def tracker(event: dict) -> None:
        tracker_calls.append(event)
    adapter.handle_message = tracker  # type: ignore[method-assign]

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    msg = _make_message("msg-1", "trusted@example.com", "Trusted", received=future_ts)
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=25")).mock(
        return_value=httpx.Response(200, json={"value": [msg]})
    )

    mark_read_route = respx.patch(_mail_url("messages/msg-1")).mock(
        return_value=httpx.Response(200, json={})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert len(tracker_calls) == 1
    assert not mark_read_route.called


# ── Polling: Watermark Bug Fix (newest-first order) ────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_watermark_does_not_skip_newer_messages_in_same_poll(env, state_path):
    """BUG 1 fix: When messages arrive newest-first, all should be processed.

    Before the fix, processing msg-new (12:00) would advance the watermark,
    causing msg-older (11:00) to be incorrectly skipped as 'older than watermark'.
    """
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker_calls: list[dict] = []
    async def tracker(event: dict) -> None:
        tracker_calls.append(event)
    adapter.handle_message = tracker  # type: ignore[method-assign]

    # Messages in newest-first order (as Graph returns them)
    msg_newer = _make_message("msg-newer", "trusted@example.com", "Trusted",
                               subject="Newer", received="2026-06-17T12:00:00Z")
    msg_older = _make_message("msg-older", "trusted@example.com", "Trusted",
                               subject="Older", received="2026-06-17T11:00:00Z")

    # Set watermark to before both messages
    pre_state = PollState()
    pre_state.watermark = "2026-06-17T10:00:00Z"
    pre_state.save(state_path)

    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=25")).mock(
        return_value=httpx.Response(200, json={"value": [msg_newer, msg_older]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker_calls) == 2
    ids = {call["message_id"] for call in tracker_calls}
    assert ids == {"msg-newer", "msg-older"}


# ── Connection Callbacks ────────────────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_connection_callbacks_return_true_after_connect(env, state_path):
    """BUG 3 fix: check_connected() tracks runtime state; is_connected_fn() checks config."""
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter, check_connected, is_connected_fn

    adapter = M365EmailAdapter()
    assert check_connected() is False
    # With env vars set via fixture, is_connected_fn should see configuration
    assert is_connected_fn() is True

    await adapter.connect()
    assert check_connected() is True
    assert is_connected_fn() is True

    await adapter.disconnect()
    assert check_connected() is False
    # is_connected_fn reflects config, not runtime connection state
    assert is_connected_fn() is True


def test_check_requirements_returns_true_when_env_set(env):
    """check_requirements() must succeed when env is configured, without connect()."""
    _ = env
    from adapter import check_requirements

    assert check_requirements() is True


def test_check_requirements_returns_false_when_env_missing(monkeypatch):
    """check_requirements() must fail when required env vars are absent."""
    import os
    monkeypatch.delenv("M365_MAIL_CLIENT_ID", raising=False)
    monkeypatch.delenv("M365_MAIL_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("M365_MAIL_TENANT_ID", raising=False)
    monkeypatch.delenv("M365_MAILBOX_USER", raising=False)
    monkeypatch.delenv("EMAIL_ALLOWED_USERS", raising=False)

    from adapter import check_requirements

    assert check_requirements() is False


# ── Email Tool Confirmation Gate Tests (Parameterized) ─────────────────────

# Each tool has the same 3-phase pattern: returns token, sends direct when disabled, confirm completes.
_EMAIL_TOOL_PARAMS = [
    pytest.param(
        "send_email",
        {"to": "test@example.com", "subject": "hello", "body": "world"},
        "sendMail",
        id="send_email",
    ),
    pytest.param(
        "reply_email",
        {"email_id": _encode_email_id("msg1"), "body": "thanks"},
        "messages/msg1/reply",
        id="reply_email",
    ),
    pytest.param(
        "reply_all",
        {"email_id": _encode_email_id("msg1"), "body": "reply all"},
        "messages/msg1/replyAll",
        id="reply_all",
    ),
    pytest.param(
        "forward_email",
        {"email_id": _encode_email_id("msg1"), "to": "fwd@example.com", "body": "see below"},
        "messages/msg1/forward",
        id="forward_email",
    ),
]


@pytest.mark.asyncio
@pytest.mark.parametrize("tool_name,args,endpoint", _EMAIL_TOOL_PARAMS)
async def test_email_wrapper_returns_token_when_confirmation_enabled(env, tool_name, args, endpoint):
    """All email wrappers return confirmation token when confirmation is enabled."""
    _ = env
    from adapter import send_email_wrapper, reply_email_wrapper, reply_all_wrapper, forward_email_wrapper

    wrapper = {
        "send_email": send_email_wrapper,
        "reply_email": reply_email_wrapper,
        "reply_all": reply_all_wrapper,
        "forward_email": forward_email_wrapper,
    }[tool_name]

    result = await wrapper(**args)

    assert result["warning"] == "PROMPT_INJECTION_CHECK"
    assert "confirmation_token" in result
    assert isinstance(result["confirmation_token"], str)


@pytest.mark.asyncio
@respx.mock
@pytest.mark.parametrize("tool_name,args,endpoint", _EMAIL_TOOL_PARAMS)
async def test_email_wrapper_sends_directly_when_disabled(env, tool_name, args, endpoint):
    """All email wrappers send directly when DISABLE_SEND_CONFIRM=true."""
    env.setenv("DISABLE_SEND_CONFIRM", "true")
    _mock_token()
    _ = respx.post(_mail_url(endpoint)).mock(return_value=httpx.Response(202, json={}))

    from adapter import send_email_wrapper, reply_email_wrapper, reply_all_wrapper, forward_email_wrapper

    wrapper = {
        "send_email": send_email_wrapper,
        "reply_email": reply_email_wrapper,
        "reply_all": reply_all_wrapper,
        "forward_email": forward_email_wrapper,
    }[tool_name]

    result = await wrapper(**args)

    assert result["success"] is True
    assert result["statusCode"] == 202


@pytest.mark.asyncio
@respx.mock
@pytest.mark.parametrize("tool_name,args,endpoint", _EMAIL_TOOL_PARAMS)
async def test_confirm_email_wrapper_sends_stored_message(env, tool_name, args, endpoint):
    """All confirm wrappers complete pending sends and remove from _ongoing_sends."""
    _ = env
    from adapter import (
        send_email_wrapper, confirm_send_email_wrapper,
        reply_email_wrapper, confirm_reply_email_wrapper,
        reply_all_wrapper, confirm_reply_all_wrapper,
        forward_email_wrapper, confirm_forward_email_wrapper,
        _ongoing_sends,
    )

    wrapper = {
        "send_email": send_email_wrapper,
        "reply_email": reply_email_wrapper,
        "reply_all": reply_all_wrapper,
        "forward_email": forward_email_wrapper,
    }[tool_name]

    confirm_wrapper = {
        "send_email": confirm_send_email_wrapper,
        "reply_email": confirm_reply_email_wrapper,
        "reply_all": confirm_reply_all_wrapper,
        "forward_email": confirm_forward_email_wrapper,
    }[tool_name]

    # Use msg2 for confirm tests to avoid endpoint collision with param args
    confirm_args = {k: v for k, v in args.items()}
    if "email_id" in confirm_args:
        confirm_args["email_id"] = _encode_email_id("msg2")
    confirm_endpoint = endpoint.replace("msg1", "msg2")

    stored = await wrapper(**confirm_args)
    token = cast(str, stored["confirmation_token"])

    _mock_token()
    route = respx.post(_mail_url(confirm_endpoint)).mock(return_value=httpx.Response(202, json={}))

    result = await confirm_wrapper(confirmation_token=token)

    assert result["success"] is True
    assert result["statusCode"] == 202
    assert route.called
    assert token not in _ongoing_sends


# ── Invalid / Expired Token Tests ──────────────────────────────────────────

@pytest.mark.asyncio
async def test_confirm_send_email_fails_for_invalid_token(env):
    _ = env
    from adapter import confirm_send_email_wrapper

    result = await confirm_send_email_wrapper(confirmation_token="invalid-token")
    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
async def test_confirm_reply_email_fails_for_invalid_token(env):
    _ = env
    from adapter import confirm_reply_email_wrapper

    result = await confirm_reply_email_wrapper(confirmation_token="invalid-token")
    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
async def test_confirm_send_email_fails_for_expired_token(env, monkeypatch):
    _ = env
    from adapter import _ongoing_sends, send_email_wrapper, confirm_send_email_wrapper

    stored = await send_email_wrapper(to="test@example.com", subject="hello", body="world")
    token = cast(str, stored["confirmation_token"])

    monkeypatch.setitem(_ongoing_sends, token, {k: v for k, v in _ongoing_sends[token].items()})
    _ongoing_sends[token]["expires"] = (datetime.now(timezone.utc) - timedelta(hours=1)).isoformat()

    result = await confirm_send_email_wrapper(confirmation_token=token)
    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
async def test_confirm_forward_email_fails_for_expired_token(env, monkeypatch):
    _ = env
    from adapter import _ongoing_sends, forward_email_wrapper, confirm_forward_email_wrapper

    stored = await forward_email_wrapper(email_id=_encode_email_id("msg1"), to="fwd@example.com", body="see below")
    token = cast(str, stored["confirmation_token"])

    monkeypatch.setitem(_ongoing_sends, token, {k: v for k, v in _ongoing_sends[token].items()})
    _ongoing_sends[token]["expires"] = (datetime.now(timezone.utc) - timedelta(hours=1)).isoformat()

    result = await confirm_forward_email_wrapper(confirmation_token=token)
    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
@respx.mock
async def test_polling_respects_poll_top_client_side(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))
    env.setenv("M365_POLL_TOP", "3")

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker_calls: list[dict] = []
    async def tracker(event: dict) -> None:
        tracker_calls.append(event)
    adapter.handle_message = tracker  # type: ignore[method-assign]

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    messages = [
        _make_message(f"msg-{i}", "trusted@example.com", "Trusted",
                      subject=f"Email {i}", received=future_ts)
        for i in range(5)
    ]
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=3")).mock(
        return_value=httpx.Response(200, json={"value": messages})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker_calls) == 3
