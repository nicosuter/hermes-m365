"""Offline regression tests for sender policy and attachment safety."""

# pyright: reportMissingImports=false, reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnknownVariableType=false, reportUnusedCallResult=false

import base64
import importlib
from pathlib import Path
from typing import Protocol, cast

import httpx
import pytest
import respx

from config import MailConfig
from graph import GRAPH_BASE_URL, GraphClient
from mail_tools import get_attachment, get_email, list_mail


class _MailToolsModule(Protocol):
    def build_saved_path(self, attachment_id: str, filename: str) -> Path: ...


@pytest.fixture
def config() -> MailConfig:
    return MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"trusted@example.com"},
        attachment_max_bytes=64,
    )


def mock_token() -> None:
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )


def mail_url(path: str) -> str:
    return f"{GRAPH_BASE_URL}/users/user%40example.org/{path}"


def message_payload(message_id: str, sender: str, *, has_attachments: bool = True) -> dict[str, object]:
    return {
        "id": message_id,
        "subject": "Policy matrix",
        "from": {"emailAddress": {"address": sender, "name": sender.split("@")[0].title()}},
        "toRecipients": [{"emailAddress": {"address": "user@example.org", "name": "Assistant"}}],
        "receivedDateTime": "2026-06-17T10:00:00Z",
        "hasAttachments": has_attachments,
        "body": {"contentType": "html", "content": "<p>Body</p>"},
    }


def attachment_payload(attachment_id: str, name: str, content: bytes) -> dict[str, object]:
    return {
        "id": attachment_id,
        "name": name,
        "contentType": "application/octet-stream",
        "size": len(content),
        "isInline": False,
        "contentBytes": base64.b64encode(content).decode(),
    }


def reload_mail_tools_with_home(monkeypatch: pytest.MonkeyPatch, home: Path) -> _MailToolsModule:
    monkeypatch.setenv("HOME", str(home))
    import attachments
    import mail_tools

    importlib.reload(attachments)
    return cast(_MailToolsModule, cast(object, importlib.reload(mail_tools)))


@pytest.mark.asyncio
@respx.mock
async def test_policy_allowed_inbound_sender_can_read_and_fetch_attachment(
    config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    mail_tools = reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    respx.get(mail_url("messages/allowed-message")).mock(
        return_value=httpx.Response(200, json=message_payload("allowed-message", "trusted@example.com"))
    )
    respx.get(mail_url("messages/allowed-message/attachments")).mock(
        return_value=httpx.Response(200, json={"value": [attachment_payload("att-allowed", "résumé.pdf", b"allowed-bytes")]})
    )
    respx.get(mail_url("messages/allowed-message/attachments/att-allowed")).mock(
        return_value=httpx.Response(200, json=attachment_payload("att-allowed", "résumé.pdf", b"allowed-bytes"))
    )

    async with GraphClient(config) as client:
        email = await get_email(config=config, client=client, email_id="allowed-message")
        attachment = await get_attachment(config=config, client=client, email_id="allowed-message", attachment_id="att-allowed")

    assert email["isAllowedInboundSender"] is True
    assert "warning" not in email
    assert email["attachments"] == [
        {
            "attachmentId": "att-allowed",
            "name": "résumé.pdf",
            "contentType": "application/octet-stream",
            "size": len(b"allowed-bytes"),
            "isInline": False,
        }
    ]
    assert attachment["filename"] == "r_sum_.pdf"
    assert Path(cast(str, attachment["savedPath"])).read_bytes() == b"allowed-bytes"
    assert Path(cast(str, attachment["savedPath"])) == mail_tools.build_saved_path("att-allowed", "r_sum_.pdf")


@pytest.mark.asyncio
@respx.mock
async def test_policy_dropped_sender_manual_list_and_read_return_email_with_warning(config: MailConfig):
    mock_token()
    inbox_route = respx.get(
        mail_url("mailFolders/inbox/messages?$orderby=receivedDateTime+desc&$top=25")
    ).mock(return_value=httpx.Response(200, json={"value": [message_payload("dropped-message", "stranger@example.com")]}))
    respx.get(mail_url("messages/dropped-message")).mock(
        return_value=httpx.Response(200, json=message_payload("dropped-message", "stranger@example.com"))
    )
    respx.get(mail_url("messages/dropped-message/attachments")).mock(
        return_value=httpx.Response(200, json={"value": [attachment_payload("att-dropped", "secret.pdf", b"must-not-leak")]})
    )

    async with GraphClient(config) as client:
        listed = await list_mail(config=config, client=client, unreadOnly=False)
        email = await get_email(config=config, client=client, email_id="dropped-message")

    assert inbox_route.called
    emails: list[dict[str, object]] = cast(list[dict[str, object]], listed["emails"])
    assert emails[0]["from"] == {"name": "Stranger", "address": "stranger@example.com"}
    assert email["isAllowedInboundSender"] is False
    assert email["warning"] == "UNTRUSTED_SENDER_NOT_IN_EMAIL_ALLOWED_USERS"
    assert email["attachments"] == [
        {
            "attachmentId": "att-dropped",
            "name": "secret.pdf",
            "contentType": "application/octet-stream",
            "size": len(b"must-not-leak"),
            "isInline": False,
        }
    ]
    assert "contentBytes" not in str(email)
    assert "must-not-leak" not in str(email)


@pytest.mark.asyncio
@respx.mock
async def test_policy_dropped_sender_attachment_is_no_op_and_never_fetches_bytes(
    config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    respx.get(mail_url("messages/dropped-message")).mock(
        return_value=httpx.Response(200, json=message_payload("dropped-message", "stranger@example.com"))
    )
    attachment_route = respx.get(mail_url("messages/dropped-message/attachments/att-dropped")).mock(
        return_value=httpx.Response(200, json=attachment_payload("att-dropped", "secret.pdf", b"must-not-leak"))
    )

    async with GraphClient(config) as client:
        result = await get_attachment(config=config, client=client, email_id="dropped-message", attachment_id="att-dropped")

    assert result == {
        "error": "ATTACHMENT_BLOCKED_UNTRUSTED_SENDER",
        "message": "Attachment blocked: sender is not in EMAIL_ALLOWED_USERS.",
        "sender": "stranger@example.com",
        "emailId": "dropped-message",
        "attachmentId": "att-dropped",
    }
    assert not attachment_route.called
    assert not (tmp_path / ".hermes").exists()
    assert "must-not-leak" not in str(result)


@pytest.mark.asyncio
@respx.mock
async def test_duplicate_filenames_save_to_distinct_deterministic_paths(
    config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    respx.get(mail_url("messages/allowed-message")).mock(
        return_value=httpx.Response(200, json=message_payload("allowed-message", "trusted@example.com"))
    )
    respx.get(mail_url("messages/allowed-message/attachments/attachment-one")).mock(
        return_value=httpx.Response(200, json=attachment_payload("attachment-one", "invoice.pdf", b"one"))
    )
    respx.get(mail_url("messages/allowed-message/attachments/attachment-two")).mock(
        return_value=httpx.Response(200, json=attachment_payload("attachment-two", "invoice.pdf", b"two"))
    )

    async with GraphClient(config) as client:
        first = await get_attachment(config=config, client=client, email_id="allowed-message", attachment_id="attachment-one")
        second = await get_attachment(config=config, client=client, email_id="allowed-message", attachment_id="attachment-two")

    first_path = Path(cast(str, first["savedPath"]))
    second_path = Path(cast(str, second["savedPath"]))
    assert first_path != second_path
    assert first_path.name == "attachment-o-invoice.pdf"
    assert second_path.name == "attachment-t-invoice.pdf"
    assert first_path.read_bytes() == b"one"
    assert second_path.read_bytes() == b"two"


@pytest.mark.asyncio
@respx.mock
async def test_message_deleted_before_attachment_fetch_raises_404_without_writing(
    config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    respx.get(mail_url("messages/deleted-message")).mock(
        return_value=httpx.Response(200, json=message_payload("deleted-message", "trusted@example.com"))
    )
    respx.get(mail_url("messages/deleted-message/attachments/att-missing")).mock(
        return_value=httpx.Response(404, json={"error": {"code": "ErrorItemNotFound", "message": "Attachment not found"}})
    )

    async with GraphClient(config) as client:
        with pytest.raises(httpx.HTTPStatusError) as exc_info:
            await get_attachment(config=config, client=client, email_id="deleted-message", attachment_id="att-missing")

    assert exc_info.value.response.status_code == 404
    assert not (tmp_path / ".hermes").exists()
