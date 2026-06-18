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


class FakeHandleMessage:
    """Tracks calls to handle_message."""

    def __init__(self) -> None:
        self.calls: list[dict] = []

    async def __call__(self, event: dict) -> None:
        self.calls.append(event)


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
@respx.mock
async def test_connect_marks_connected_and_creates_poll_task(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    assert not adapter.is_connected
    assert adapter._poll_task is None

    result = await adapter.connect()
    assert result is True
    assert adapter.is_connected
    assert adapter._poll_task is not None
    assert isinstance(adapter._poll_task, asyncio.Task)

    await adapter.disconnect()


@pytest.mark.asyncio
@respx.mock
async def test_disconnect_cancels_poll_task_and_marks_disconnected(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    await adapter.connect()
    assert adapter.is_connected

    await adapter.disconnect()
    assert not adapter.is_connected
    assert adapter._poll_task is None


@pytest.mark.asyncio
@respx.mock
async def test_connect_returns_false_on_bad_config(monkeypatch):
    for key in ("M365_MAIL_CLIENT_ID", "M365_MAIL_CLIENT_SECRET", "M365_MAIL_TENANT_ID"):
        monkeypatch.delenv(key, raising=False)

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    result = await adapter.connect()
    assert result is False
    assert not adapter.is_connected


# ── Polling: First Run Watermark ──────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_first_run_records_watermark_and_does_not_emit_old_messages(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    old_msg = _make_message("msg-old", "trusted@example.com", "Trusted", received="2026-06-16T10:00:00Z")
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [old_msg]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 0

    assert state_path.exists()
    saved = json.loads(state_path.read_text())
    assert saved["watermark"] != ""


@pytest.mark.asyncio
@respx.mock
async def test_new_message_after_watermark_calls_handle_message(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    new_msg = _make_message(
        "msg-new", "trusted@example.com", "Trusted",
        subject="Important", body="<p>Urgent</p>",
        received=future_ts,
        conversation_id="conv-42",
        internet_message_id="<urgent@example.com>",
    )
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [new_msg]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 1
    event = tracker.calls[0]
    assert event["message_id"] == "msg-new"
    assert event["graph_message_id"] == "msg-new"
    assert event["subject"] == "Important"
    assert event["conversation_id"] == "conv-42"
    assert event["internetMessageId"] == "<urgent@example.com>"
    assert event["message_type"] == "email"
    assert event["source"]["chat_id"] == "m365:trusted@example.com"
    assert event["source"]["chat_name"] == "Trusted"
    assert event["source"]["chat_type"] == "dm"
    assert event["source"]["user_id"] == "trusted@example.com"
    assert event["source"]["user_name"] == "Trusted"
    assert "Urgent" in event["text"]


# ── Polling: Unallowed Sender Dropped ─────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_unallowed_sender_dropped(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))
    env.setenv("EMAIL_ALLOWED_USERS", "trusted@example.com")

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    evil_msg = _make_message("msg-evil", "evil@example.com", "Evil",
                             subject="Hack", body="<script>alert(1)</script>",
                             received=future_ts)
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [evil_msg]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 0


@pytest.mark.asyncio
@respx.mock
async def test_empty_allowed_users_drops_all(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))
    env.setenv("EMAIL_ALLOWED_USERS", "")

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    msg = _make_message("msg-1", "anyone@example.com", "Anyone", received=future_ts)
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [msg]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 0


# ── Polling: Duplicate Suppression ────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_duplicate_message_id_skipped_after_state_persistence(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    pre_state = PollState()
    pre_state.add("msg-1", "2026-06-17T12:00:00Z")
    pre_state.save(state_path)

    adapter = M365EmailAdapter()
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    msg = _make_message("msg-1", "trusted@example.com", "Trusted", received="2026-06-17T12:00:00Z")
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [msg]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 0


# ── Send Tests ────────────────────────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_send_posts_graph_sendmail(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    await adapter.connect()

    send_route = respx.post(_mail_url("sendMail")).mock(
        return_value=httpx.Response(202, json={})
    )

    result = await adapter.send(
        "m365:test@example.com",
        "hello",
        metadata={"subject": "Hi"},
    )

    assert result is True
    assert send_route.called
    payload = json.loads(send_route.calls.last.request.read().decode())
    assert payload["message"]["subject"] == "Hi"
    assert payload["message"]["body"]["content"] == "hello"
    assert payload["message"]["toRecipients"][0]["emailAddress"]["address"] == "test@example.com"

    await adapter.disconnect()


@pytest.mark.asyncio
@respx.mock
async def test_send_uses_thread_subject_fallback(env, state_path):
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter

    adapter = M365EmailAdapter()
    await adapter.connect()

    send_route = respx.post(_mail_url("sendMail")).mock(
        return_value=httpx.Response(202, json={})
    )

    await adapter.send(
        "m365:test@example.com",
        "reply body",
        metadata={"thread_subject": "Original Subject"},
    )

    payload = json.loads(send_route.calls.last.request.read().decode())
    assert payload["message"]["subject"] == "Re: Original Subject"

    await adapter.disconnect()


@pytest.mark.asyncio
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
    assert result is False


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
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    future_ts = (datetime.now(timezone.utc) + timedelta(seconds=10)).isoformat()
    msg = _make_message("msg-1", "trusted@example.com", "Trusted", received=future_ts)
    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [msg]})
    )

    mark_read_route = respx.patch(_mail_url("messages/msg-1")).mock(
        return_value=httpx.Response(200, json={})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert len(tracker.calls) == 1
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
    tracker = FakeHandleMessage()
    adapter.set_handle_message(tracker)

    # Messages in newest-first order (as Graph returns them)
    msg_newer = _make_message("msg-newer", "trusted@example.com", "Trusted",
                               subject="Newer", received="2026-06-17T12:00:00Z")
    msg_older = _make_message("msg-older", "trusted@example.com", "Trusted",
                               subject="Older", received="2026-06-17T11:00:00Z")

    # Set watermark to before both messages
    pre_state = PollState()
    pre_state.watermark = "2026-06-17T10:00:00Z"
    pre_state.save(state_path)

    inbox_route = respx.get(_mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=50")).mock(
        return_value=httpx.Response(200, json={"value": [msg_newer, msg_older]})
    )

    await adapter.connect()
    await asyncio.sleep(0.3)
    await adapter.disconnect()

    assert inbox_route.called
    assert len(tracker.calls) == 2
    ids = {call["message_id"] for call in tracker.calls}
    assert ids == {"msg-newer", "msg-older"}


# ── Connection Callbacks ────────────────────────────────────────────────────

@pytest.mark.asyncio
@respx.mock
async def test_connection_callbacks_return_true_after_connect(env, state_path):
    """BUG 3 fix: check_connected() and is_connected_fn() must reflect actual state."""
    _mock_token()
    env.setenv("M365_EMAIL_STATE_PATH", str(state_path))

    from adapter import M365EmailAdapter, check_connected, is_connected_fn

    adapter = M365EmailAdapter()
    assert check_connected() is False
    assert is_connected_fn() is False

    await adapter.connect()
    assert check_connected() is True
    assert is_connected_fn() is True

    await adapter.disconnect()
    assert check_connected() is False
    assert is_connected_fn() is False


# ── Send Email Confirmation Gate Tests ──────────────────────────────────────

@pytest.mark.asyncio
async def test_send_email_wrapper_returns_token_when_confirmation_enabled(env):
    _ = env
    from adapter import send_email_wrapper

    result = await send_email_wrapper(
        to="test@example.com", subject="hello", body="world"
    )

    assert result["warning"] == "PROMPT_INJECTION_CHECK"
    assert "confirmation_token" in result
    assert isinstance(result["confirmation_token"], str)


@pytest.mark.asyncio
@respx.mock
async def test_send_email_wrapper_sends_directly_when_disabled(env):
    env.setenv("DISABLE_SEND_CONFIRM", "true")
    _mock_token()
    _ = respx.post(_mail_url("sendMail")).mock(return_value=httpx.Response(202, json={}))

    from adapter import send_email_wrapper

    result = await send_email_wrapper(
        to="test@example.com", subject="hello", body="world"
    )

    assert result["success"] is True
    assert result["statusCode"] == 202


@pytest.mark.asyncio
@respx.mock
async def test_confirm_send_email_wrapper_sends_stored_message(env):
    _ = env
    from adapter import send_email_wrapper, confirm_send_email_wrapper, _ongoing_sends

    stored = await send_email_wrapper(
        to="test@example.com", subject="hello", body="world"
    )
    token = cast(str, stored["confirmation_token"])

    _mock_token()
    send_route = respx.post(_mail_url("sendMail")).mock(return_value=httpx.Response(202, json={}))

    result = await confirm_send_email_wrapper(confirmation_token=token)

    assert result["success"] is True
    assert result["statusCode"] == 202
    assert send_route.called
    assert token not in _ongoing_sends


@pytest.mark.asyncio
async def test_confirm_send_email_wrapper_fails_for_invalid_token(env):
    _ = env
    from adapter import confirm_send_email_wrapper

    result = await confirm_send_email_wrapper(confirmation_token="invalid-token")

    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
async def test_confirm_send_email_wrapper_fails_for_expired_token(env, monkeypatch):
    _ = env
    from adapter import _ongoing_sends, send_email_wrapper, confirm_send_email_wrapper
    from datetime import datetime, timezone

    stored = await send_email_wrapper(
        to="test@example.com", subject="hello", body="world"
    )
    token = cast(str, stored["confirmation_token"])

    monkeypatch.setitem(
        _ongoing_sends, token,
        {k: v for k, v in _ongoing_sends[token].items()}
    )
    _ongoing_sends[token]["expires"] = (
        datetime.now(timezone.utc) - timedelta(hours=1)
    ).isoformat()

    result = await confirm_send_email_wrapper(confirmation_token=token)

    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


# ── Reply Email Confirmation Gate Tests ─────────────────────────────────────

@pytest.mark.asyncio
async def test_reply_email_wrapper_returns_token_when_confirmation_enabled(env):
    _ = env
    from adapter import reply_email_wrapper

    result = await reply_email_wrapper(email_id="msg1", body="thanks")

    assert result["warning"] == "PROMPT_INJECTION_CHECK"
    assert "confirmation_token" in result
    assert isinstance(result["confirmation_token"], str)


@pytest.mark.asyncio
@respx.mock
async def test_reply_email_wrapper_sends_directly_when_disabled(env):
    env.setenv("DISABLE_SEND_CONFIRM", "true")
    _mock_token()
    _ = respx.post(_mail_url("messages/msg1/reply")).mock(return_value=httpx.Response(202, json={}))

    from adapter import reply_email_wrapper

    result = await reply_email_wrapper(email_id="msg1", body="thanks")

    assert result["success"] is True
    assert result["statusCode"] == 202


@pytest.mark.asyncio
@respx.mock
async def test_confirm_reply_email_wrapper_sends_stored_message(env):
    _ = env
    from adapter import reply_email_wrapper, confirm_reply_email_wrapper, _ongoing_sends

    stored = await reply_email_wrapper(email_id="msg2", body="thanks")
    token = cast(str, stored["confirmation_token"])

    _mock_token()
    reply_route = respx.post(_mail_url("messages/msg2/reply")).mock(return_value=httpx.Response(202, json={}))

    result = await confirm_reply_email_wrapper(confirmation_token=token)

    assert result["success"] is True
    assert result["statusCode"] == 202
    assert reply_route.called
    assert token not in _ongoing_sends


# ── Reply All Confirmation Gate Tests ─────────────────────────────────────

@pytest.mark.asyncio
async def test_reply_all_wrapper_returns_token_when_confirmation_enabled(env):
    _ = env
    from adapter import reply_all_wrapper

    result = await reply_all_wrapper(email_id="msg1", body="reply all")

    assert result["warning"] == "PROMPT_INJECTION_CHECK"
    assert "confirmation_token" in result
    assert isinstance(result["confirmation_token"], str)


@pytest.mark.asyncio
@respx.mock
async def test_reply_all_wrapper_sends_directly_when_disabled(env):
    env.setenv("DISABLE_SEND_CONFIRM", "true")
    _mock_token()
    _ = respx.post(_mail_url("messages/msg1/replyAll")).mock(return_value=httpx.Response(202, json={}))

    from adapter import reply_all_wrapper

    result = await reply_all_wrapper(email_id="msg1", body="reply all")

    assert result["success"] is True
    assert result["statusCode"] == 202


@pytest.mark.asyncio
@respx.mock
async def test_confirm_reply_all_wrapper_sends_stored_message(env):
    _ = env
    from adapter import reply_all_wrapper, confirm_reply_all_wrapper, _ongoing_sends

    stored = await reply_all_wrapper(email_id="msg2", body="reply all")
    token = cast(str, stored["confirmation_token"])

    _mock_token()
    reply_route = respx.post(_mail_url("messages/msg2/replyAll")).mock(return_value=httpx.Response(202, json={}))

    result = await confirm_reply_all_wrapper(confirmation_token=token)

    assert result["success"] is True
    assert result["statusCode"] == 202
    assert reply_route.called
    assert token not in _ongoing_sends


# ── Forward Email Confirmation Gate Tests ─────────────────────────────────

@pytest.mark.asyncio
async def test_forward_email_wrapper_returns_token_when_confirmation_enabled(env):
    _ = env
    from adapter import forward_email_wrapper

    result = await forward_email_wrapper(email_id="msg1", to="fwd@example.com", body="see below")

    assert result["warning"] == "PROMPT_INJECTION_CHECK"
    assert "confirmation_token" in result
    assert isinstance(result["confirmation_token"], str)


@pytest.mark.asyncio
@respx.mock
async def test_forward_email_wrapper_sends_directly_when_disabled(env):
    env.setenv("DISABLE_SEND_CONFIRM", "true")
    _mock_token()
    _ = respx.post(_mail_url("messages/msg1/forward")).mock(return_value=httpx.Response(202, json={}))

    from adapter import forward_email_wrapper

    result = await forward_email_wrapper(email_id="msg1", to="fwd@example.com", body="see below")

    assert result["success"] is True
    assert result["statusCode"] == 202


@pytest.mark.asyncio
@respx.mock
async def test_confirm_forward_email_wrapper_sends_stored_message(env):
    _ = env
    from adapter import forward_email_wrapper, confirm_forward_email_wrapper, _ongoing_sends

    stored = await forward_email_wrapper(email_id="msg2", to="fwd@example.com", body="see below")
    token = cast(str, stored["confirmation_token"])

    _mock_token()
    forward_route = respx.post(_mail_url("messages/msg2/forward")).mock(return_value=httpx.Response(202, json={}))

    result = await confirm_forward_email_wrapper(confirmation_token=token)

    assert result["success"] is True
    assert result["statusCode"] == 202
    assert forward_route.called
    assert token not in _ongoing_sends


# ── General Confirm Tests ─────────────────────────────────────────────────

@pytest.mark.asyncio
async def test_confirm_reply_email_wrapper_fails_for_invalid_token(env):
    _ = env
    from adapter import confirm_reply_email_wrapper

    result = await confirm_reply_email_wrapper(confirmation_token="invalid-token")

    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"


@pytest.mark.asyncio
async def test_confirm_forward_email_wrapper_fails_for_expired_token(env, monkeypatch):
    _ = env
    from adapter import _ongoing_sends, forward_email_wrapper, confirm_forward_email_wrapper
    from datetime import datetime, timezone

    stored = await forward_email_wrapper(email_id="msg1", to="fwd@example.com", body="see below")
    token = cast(str, stored["confirmation_token"])

    monkeypatch.setitem(
        _ongoing_sends, token,
        {k: v for k, v in _ongoing_sends[token].items()}
    )
    _ongoing_sends[token]["expires"] = (
        datetime.now(timezone.utc) - timedelta(hours=1)
    ).isoformat()

    result = await confirm_forward_email_wrapper(confirmation_token=token)

    assert result["error"] == "INVALID_OR_EXPIRED_TOKEN"
