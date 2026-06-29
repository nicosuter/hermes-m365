# pyright: reportMissingImports=false, reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnknownVariableType=false, reportUnusedCallResult=false, reportArgumentType=false, reportIndexIssue=false

import base64
import importlib
from pathlib import Path
from typing import Protocol, cast

import httpx
import pytest
import respx

from config import MailConfig
from graph import GRAPH_BASE_URL, GraphClient
from mail_tools import _build_odata_filter, _get_message_metadata, _trusted_sender_address, get_attachment, get_email, get_summary, list_mail, normalize_summary_response, send_email


@pytest.fixture
def config() -> MailConfig:
    return MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"trusted@example.com"},
        attachment_max_bytes=16,
    )


def mock_token() -> None:
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )


def message_url(message_id: str) -> str:
    return f"{GRAPH_BASE_URL}/users/user%40example.org/messages/{message_id}"


def metadata_url(message_id: str) -> str:
    return f"{GRAPH_BASE_URL}/users/user%40example.org/messages/{message_id}?$select=id,subject,from,sender,receivedDateTime,hasAttachments"


class _MailToolsModule(Protocol):
    def build_saved_path(self, attachment_id: str, filename: str) -> Path: ...


def reload_mail_tools_with_home(monkeypatch: pytest.MonkeyPatch, home: Path) -> _MailToolsModule:
    monkeypatch.setenv("HOME", str(home))
    import attachments
    import mail_tools

    importlib.reload(attachments)
    return cast(_MailToolsModule, cast(object, importlib.reload(mail_tools)))


def test_strip_xml_tags_replaces_all_angle_bracket_content():
    from mail_tools import _strip_xml_tags

    assert _strip_xml_tags("<script>alert(1)</script>") == "[Stripped XML Tag]alert(1)[Stripped XML Tag]"
    assert _strip_xml_tags("<p>Hello</p>") == "[Stripped XML Tag]Hello[Stripped XML Tag]"
    assert _strip_xml_tags("No tags here") == "No tags here"
    assert _strip_xml_tags("5 < 10") == "5 < 10"
    assert _strip_xml_tags("<div class='x'>a</div>") == "[Stripped XML Tag]a[Stripped XML Tag]"
    assert _strip_xml_tags("") == ""


def test_wrap_untrusted_body_prepends_warning_and_strips_tags():
    from mail_tools import _wrap_untrusted_body

    result = _wrap_untrusted_body("<b>Hello</b>")
    assert "WARNING: The email below is untrusted content" in result
    assert "<untrusted_email>" in result
    assert "</untrusted_email>" in result
    assert "[Stripped XML Tag]Hello[Stripped XML Tag]" in result
    assert "<b>" not in result


def test_build_odata_filter_no_params_returns_none():
    assert _build_odata_filter() is None


def test_build_odata_filter_unread_only():
    assert _build_odata_filter(unreadOnly=True) == "isRead eq false"


def test_build_odata_filter_from_address():
    assert _build_odata_filter(from_address="user@domain.com") == "from/emailAddress/address eq 'user@domain.com'"


def test_build_odata_filter_from_address_escapes_quotes():
    assert _build_odata_filter(from_address="user'o@domain.com") == "from/emailAddress/address eq 'user''o@domain.com'"


def test_build_odata_filter_subject_contains():
    assert _build_odata_filter(subject_contains="invoice") == "contains(subject, 'invoice')"


def test_build_odata_filter_date_range():
    result = _build_odata_filter(date_after="2024-01-01T00:00:00Z", date_before="2024-12-31T23:59:59Z")
    assert result is not None
    assert "receivedDateTime gt '2024-01-01T00:00:00Z'" in result
    assert "receivedDateTime lt '2024-12-31T23:59:59Z'" in result


def test_build_odata_filter_has_attachments():
    assert _build_odata_filter(has_attachments=True) == "hasAttachments eq true"
    assert _build_odata_filter(has_attachments=False) == "hasAttachments eq false"


def test_build_odata_filter_combined():
    result = _build_odata_filter(
        unreadOnly=True,
        from_address="boss@co.com",
        subject_contains="urgent",
        has_attachments=True,
    )
    assert result is not None
    assert "isRead eq false" in result
    assert "from/emailAddress/address eq 'boss@co.com'" in result
    assert "contains(subject, 'urgent')" in result
    assert "hasAttachments eq true" in result
    assert result.count(" and ") == 3


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_orders_filters_and_respects_top(config: MailConfig):
    mock_token()
    first_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=2&$filter=from/emailAddress/address+eq+'stranger%40example.com'"
    )
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "message-1",
                        "subject": "Newest",
                        "from": {"emailAddress": {"address": "stranger@example.com", "name": "Stranger"}},
                        "receivedDateTime": "2026-06-17T10:00:00Z",
                        "hasAttachments": True,
                        "body": {"content": "not returned"},
                    }
                ],
            },
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(
            config=config, client=client, unreadOnly=False, from_address="stranger@example.com", top=2
        )

    assert first_route.called
    assert result == {
        "emails": [
            {
                "id": "message-1",
                "subject": "Newest",
                "from": {"name": "Stranger", "address": "stranger@example.com"},
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": True,
            },
        ],
    }


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_unread_only_adds_isRead_filter(config: MailConfig):
    mock_token()
    first_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=25&$filter=isRead+eq+false"
    )
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "message-unread",
                        "subject": "Unread Email",
                        "from": {"emailAddress": {"address": "sender@example.com", "name": "Sender"}},
                        "receivedDateTime": "2026-06-17T12:00:00Z",
                        "hasAttachments": False,
                        "isRead": False,
                    }
                ]
            },
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(config=config, client=client, unreadOnly=True)

    assert first_route.called
    assert result == {
        "emails": [
            {
                "id": "message-unread",
                "subject": "Unread Email",
                "from": {"name": "Sender", "address": "sender@example.com"},
                "receivedDateTime": "2026-06-17T12:00:00Z",
                "hasAttachments": False,
            }
        ]
    }


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_unread_only_combines_with_from_filter(config: MailConfig):
    mock_token()
    first_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=5&$filter=isRead+eq+false+and+from/emailAddress/address+eq+'boss%40example.com'"
    )
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={"value": []},
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(
            config=config, client=client, unreadOnly=True, top=5, from_address="boss@example.com"
        )

    assert first_route.called
    assert result == {"emails": []}


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_multiple_structured_filters(config: MailConfig):
    """Structured filters combine correctly: unreadOnly + from + subjectContains + dateAfter + hasAttachments."""
    mock_token()
    first_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=10&$filter=isRead+eq+false+and+from/emailAddress/address+eq+'boss%40example.com'+and+contains(subject,+%27invoice%27)+and+receivedDateTime+gt+%272024-01-01T00%3A00%3A00Z%27+and+hasAttachments+eq+true"
    )
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={"value": []},
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(
            config=config, client=client, unreadOnly=True, top=10,
            from_address="boss@example.com",
            subject_contains="invoice",
            date_after="2024-01-01T00:00:00Z",
            has_attachments=True,
        )

    assert first_route.called
    assert result == {"emails": []}


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_no_filters_omits_filter_param(config: MailConfig):
    """When no filters are set (unreadOnly=False, all others None), $filter is omitted."""
    mock_token()
    first_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=25"
    )
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={"value": []},
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(
            config=config, client=client, unreadOnly=False
        )

    assert first_route.called
    assert result == {"emails": []}


@pytest.mark.asyncio
@respx.mock
async def test_get_email_blocks_untrusted_sender_with_metadata_only(config: MailConfig):
    """Untrusted sender gets hard-block envelope with NO body/attachments fetched."""
    mock_token()
    metadata_route = respx.get(metadata_url("message-untrusted")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-untrusted",
                "subject": "Phishing Attempt",
                "from": {"emailAddress": {"address": "stranger@example.com", "name": "Stranger"}},
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": True,
            },
        )
    )
    # Full message endpoint must NOT be called
    full_message_route = respx.get(message_url("message-untrusted")).mock(
        return_value=httpx.Response(200, json={"body": {"content": "secret"}})
    )
    # Attachment endpoint must NOT be called
    attachment_route = respx.get(f"{message_url('message-untrusted')}/attachments").mock(
        return_value=httpx.Response(200, json={"value": []})
    )

    async with GraphClient(config) as client:
        result = await get_email(config=config, client=client, email_id="message-untrusted")

    assert metadata_route.called
    assert not full_message_route.called
    assert not attachment_route.called
    assert result == {
        "error": "EMAIL_BODY_BLOCKED_UNTRUSTED_SENDER",
        "message": 'Email body blocked: sender is not in EMAIL_ALLOWED_USERS. Use get_summary(email_id, schema_name="general") for a schema-constrained summary.',
        "emailId": "message-untrusted",
        "sender": "stranger@example.com",
        "isAllowedInboundSender": False,
    }
    # Must NOT contain body, attachments, or attachmentsById
    assert "body" not in result
    assert "attachments" not in result
    assert "attachmentsById" not in result
    assert "warning" not in result


@pytest.mark.asyncio
@respx.mock
async def test_get_email_trusted_sender_gets_full_body_and_attachments(config: MailConfig):
    """Trusted sender gets full response: body, attachments, attachmentsById, isAllowedInboundSender=True."""
    mock_token()
    respx.get(metadata_url("message-trusted")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-trusted",
                "subject": "Quarterly Report",
                "from": {"emailAddress": {"address": "trusted@example.com", "name": "Trusted"}},
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": True,
            },
        )
    )
    respx.get(message_url("message-trusted")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-trusted",
                "subject": "Quarterly Report",
                "from": {"emailAddress": {"address": "trusted@example.com", "name": "Trusted"}},
                "toRecipients": [{"emailAddress": {"address": "user@example.org", "name": "Me"}}],
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": True,
                "body": {"contentType": "text", "content": "Here is the report."},
            },
        )
    )
    respx.get(f"{message_url('message-trusted')}/attachments").mock(
        return_value=httpx.Response(
            200,
            json={"value": [{"id": "att-1", "name": "report.pdf", "contentType": "application/pdf", "size": 100, "isInline": False}]},
        )
    )

    async with GraphClient(config) as client:
        result = await get_email(config=config, client=client, email_id="message-trusted")

    assert result["isAllowedInboundSender"] is True
    assert "warning" not in result
    assert "error" not in result
    assert result["body"] == "Here is the report."
    assert result["subject"] == "Quarterly Report"
    assert result["from"] == {"name": "Trusted", "address": "trusted@example.com"}
    assert result["attachments"] == [{"attachmentId": "att-1", "name": "report.pdf", "contentType": "application/pdf", "size": 100, "isInline": False}]
    assert result["attachmentsById"] == {"att-1": cast(dict[str, object], result["attachments"][0])}


@pytest.mark.asyncio
@respx.mock
async def test_get_email_trusted_sender_no_attachments(config: MailConfig):
    """Trusted sender with no attachments skips attachment request entirely."""
    mock_token()
    respx.get(metadata_url("message-2")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-2",
                "subject": "No attachments",
                "from": {"emailAddress": {"address": "TRUSTED@example.com", "name": "Trusted"}},
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": False,
            },
        )
    )
    respx.get(message_url("message-2")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-2",
                "subject": "No attachments",
                "from": {"emailAddress": {"address": "TRUSTED@example.com", "name": "Trusted"}},
                "receivedDateTime": "2026-06-17T10:00:00Z",
                "hasAttachments": False,
                "body": {"content": "Plain body"},
            },
        )
    )
    # Attachment endpoint must NOT be called when hasAttachments=False
    attachment_route = respx.get(f"{message_url('message-2')}/attachments").mock(
        return_value=httpx.Response(200, json={"value": []})
    )

    async with GraphClient(config) as client:
        result = await get_email(config=config, client=client, email_id="message-2")

    assert result["isAllowedInboundSender"] is True
    assert "warning" not in result
    assert result["attachments"] == []
    assert result["attachmentsById"] == {}
    assert not attachment_route.called


@pytest.mark.asyncio
@respx.mock
async def test_get_attachment_blocks_untrusted_sender_without_fetching_or_writing(config: MailConfig, tmp_path: Path):
    reload_mail_tools_with_home(pytest.MonkeyPatch(), tmp_path)
    mock_token()
    respx.get(message_url("message-1")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "message-1",
                "from": {"emailAddress": {"address": "stranger@example.com"}},
            },
        )
    )
    attachment_route = respx.get(f"{message_url('message-1')}/attachments/attachment-1").mock(
        return_value=httpx.Response(200, json={"contentBytes": "bmV2ZXItZmV0Y2hlZA=="})
    )

    async with GraphClient(config) as client:
        result = await get_attachment(config=config, client=client, email_id="message-1", attachment_id="attachment-1")

    assert result == {
        "error": "ATTACHMENT_BLOCKED_UNTRUSTED_SENDER",
        "message": "Attachment blocked: sender is not in EMAIL_ALLOWED_USERS.",
        "sender": "stranger@example.com",
        "emailId": "message-1",
        "attachmentId": "attachment-1",
    }
    assert not attachment_route.called
    assert not (tmp_path / ".hermes").exists()


@pytest.mark.asyncio
@respx.mock
async def test_get_attachment_fetches_allowed_sender_enforces_size_and_saves_file(
    config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    mail_tools = reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    content = b"hello"
    respx.get(message_url("message-2")).mock(
        return_value=httpx.Response(
            200,
            json={"id": "message-2", "from": {"emailAddress": {"address": "trusted@example.com"}}},
        )
    )
    respx.get(f"{message_url('message-2')}/attachments/attachment-2").mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "attachment-2",
                "name": "../invoice.pdf",
                "contentType": "application/pdf",
                "size": len(content),
                "isInline": False,
                "contentBytes": base64.b64encode(content).decode(),
            },
        )
    )

    async with GraphClient(config) as client:
        result = await get_attachment(config=config, client=client, email_id="message-2", attachment_id="attachment-2")

    saved_path = mail_tools.build_saved_path("attachment-2", "../invoice.pdf")
    assert result == {
        "filename": "invoice.pdf",
        "mimeType": "application/pdf",
        "size": len(content),
        "savedPath": str(saved_path),
        "isInline": False,
    }
    assert saved_path.read_bytes() == content


@pytest.mark.asyncio
@respx.mock
async def test_get_attachment_rejects_oversized_attachment_before_writing(config: MailConfig, monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    reload_mail_tools_with_home(monkeypatch, tmp_path)
    mock_token()
    respx.get(message_url("message-3")).mock(
        return_value=httpx.Response(
            200,
            json={"id": "message-3", "from": {"emailAddress": {"address": "trusted@example.com"}}},
        )
    )
    respx.get(f"{message_url('message-3')}/attachments/attachment-3").mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "attachment-3",
                "name": "large.bin",
                "contentType": "application/octet-stream",
                "size": 17,
                "contentBytes": base64.b64encode(b"0123456789abcdefg").decode(),
            },
        )
    )

    async with GraphClient(config) as client:
        with pytest.raises(ValueError, match="exceeding maximum allowed size"):
            await get_attachment(config=config, client=client, email_id="message-3", attachment_id="attachment-3")

    assert not (tmp_path / ".hermes").exists()


@pytest.mark.asyncio
@respx.mock
async def test_send_email_posts_text_message_and_optional_reply_to(config: MailConfig):
    mock_token()
    send_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/sendMail").mock(
        return_value=httpx.Response(202, json={})
    )

    async with GraphClient(config) as client:
        result = await send_email(
            config=config,
            client=client,
            to=["person@example.com", "Other <other@example.com>"],
            subject="Hello",
            body="Plain text body",
            reply_to="reply@example.com",
        )

    assert result == {"success": True, "statusCode": 202}
    payload = send_route.calls.last.request.read().decode()
    assert payload == (
        '{"message":{"subject":"Hello","body":{"contentType":"text","content":"Plain text body"},'
        '"toRecipients":[{"emailAddress":{"address":"person@example.com"}},'
        '{"emailAddress":{"address":"Other <other@example.com>"}}],"replyTo":[{"emailAddress":{"address":"reply@example.com"}}]},'
        '"saveToSentItems":true}'
    )


@pytest.mark.asyncio
@respx.mock
async def test_send_email_posts_html_message_when_content_type_html(config: MailConfig):
    mock_token()
    send_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/sendMail").mock(
        return_value=httpx.Response(202, json={})
    )

    async with GraphClient(config) as client:
        result = await send_email(
            config=config,
            client=client,
            to="person@example.com",
            subject="HTML Email",
            body="<h1>Hello</h1><p>This is HTML</p>",
            content_type="html",
        )

    assert result == {"success": True, "statusCode": 202}
    payload = send_route.calls.last.request.read().decode()
    assert '"contentType":"html"' in payload
    assert '"content":"<h1>Hello</h1><p>This is HTML</p>"' in payload


@pytest.mark.asyncio
@respx.mock
async def test_send_email_defaults_to_text_content_type(config: MailConfig):
    mock_token()
    send_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/sendMail").mock(
        return_value=httpx.Response(202, json={})
    )

    async with GraphClient(config) as client:
        result = await send_email(
            config=config,
            client=client,
            to="person@example.com",
            subject="Default CT",
            body="Plain text",
        )

    assert result == {"success": True, "statusCode": 202}
    payload = send_route.calls.last.request.read().decode()
    assert '"contentType":"text"' in payload


@pytest.mark.asyncio
@respx.mock
async def test_mark_read_sets_isRead_true(config: MailConfig):
    from mail_tools import mark_read

    mock_token()
    patch_route = respx.patch(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await mark_read(config=config, client=client, email_id="message-1")

    assert result == {"success": True, "emailId": "message-1", "isRead": True}
    assert patch_route.called
    payload = patch_route.calls.last.request.read().decode()
    assert payload == '{"isRead":true}'


@pytest.mark.asyncio
@respx.mock
async def test_mark_unread_sets_isRead_false(config: MailConfig):
    from mail_tools import mark_unread

    mock_token()
    patch_route = respx.patch(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-2").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await mark_unread(config=config, client=client, email_id="message-2")

    assert result == {"success": True, "emailId": "message-2", "isRead": False}
    assert patch_route.called
    payload = patch_route.calls.last.request.read().decode()
    assert payload == '{"isRead":false}'


@pytest.mark.asyncio
@respx.mock
async def test_reply_email_posts_to_reply_endpoint(config: MailConfig):
    from mail_tools import reply_email

    mock_token()
    reply_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/reply").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await reply_email(config=config, client=client, email_id="message-1", body="Replied!")

    assert result == {"success": True, "statusCode": 200}
    assert reply_route.called
    payload = reply_route.calls.last.request.read().decode()
    assert payload == '{"message":{"body":{"contentType":"text","content":"Replied!"}}}'


@pytest.mark.asyncio
@respx.mock
async def test_reply_email_posts_html_when_content_type_html(config: MailConfig):
    from mail_tools import reply_email

    mock_token()
    reply_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/reply").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await reply_email(
            config=config, client=client, email_id="message-1", body="<h1>HTML Reply</h1>", content_type="html"
        )

    assert result == {"success": True, "statusCode": 200}
    assert reply_route.called
    payload = reply_route.calls.last.request.read().decode()
    assert '"contentType":"html"' in payload
    assert '"content":"<h1>HTML Reply</h1>"' in payload


@pytest.mark.asyncio
@respx.mock
async def test_reply_all_posts_to_replyAll_endpoint(config: MailConfig):
    from mail_tools import reply_all

    mock_token()
    reply_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/replyAll").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await reply_all(config=config, client=client, email_id="message-1", body="Replied to all!")

    assert result == {"success": True, "statusCode": 200}
    assert reply_route.called
    payload = reply_route.calls.last.request.read().decode()
    assert payload == '{"message":{"body":{"contentType":"text","content":"Replied to all!"}}}'


@pytest.mark.asyncio
@respx.mock
async def test_reply_all_posts_html_when_content_type_html(config: MailConfig):
    from mail_tools import reply_all

    mock_token()
    reply_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/replyAll").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await reply_all(
            config=config, client=client, email_id="message-1", body="<p>HTML reply all</p>", content_type="html"
        )

    assert result == {"success": True, "statusCode": 200}
    assert reply_route.called
    payload = reply_route.calls.last.request.read().decode()
    assert '"contentType":"html"' in payload
    assert '"content":"<p>HTML reply all</p>"' in payload


@pytest.mark.asyncio
@respx.mock
async def test_forward_email_posts_to_forward_endpoint_with_recipients(config: MailConfig):
    from mail_tools import forward_email

    mock_token()
    forward_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/forward").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await forward_email(
            config=config, client=client, email_id="message-1", to="new@example.com", body="FW: forwarded"
        )

    assert result == {"success": True, "statusCode": 200}
    assert forward_route.called
    payload = forward_route.calls.last.request.read().decode()
    assert payload == '{"message":{"body":{"contentType":"text","content":"FW: forwarded"},"toRecipients":[{"emailAddress":{"address":"new@example.com"}}]}}'


@pytest.mark.asyncio
@respx.mock
async def test_forward_email_posts_html_when_content_type_html(config: MailConfig):
    from mail_tools import forward_email

    mock_token()
    forward_route = respx.post(f"{GRAPH_BASE_URL}/users/user%40example.org/messages/message-1/forward").mock(
        return_value=httpx.Response(200, json={})
    )

    async with GraphClient(config) as client:
        result = await forward_email(
            config=config, client=client, email_id="message-1", to="new@example.com", body="<p>HTML forward</p>", content_type="html"
        )

    assert result == {"success": True, "statusCode": 200}
    assert forward_route.called
    payload = forward_route.calls.last.request.read().decode()
    assert '"contentType":"html"' in payload
    assert '"content":"<p>HTML forward</p>"' in payload


@pytest.mark.asyncio
@respx.mock
async def test_list_mail_truncates_client_side_when_api_returns_more(config: MailConfig):
    mock_token()
    messages = [
        {
            "id": f"msg-{i}",
            "subject": f"Email {i}",
            "from": {"emailAddress": {"address": "a@example.com", "name": "A"}},
            "receivedDateTime": f"2026-06-18T{i:02d}:00:00Z",
            "hasAttachments": False,
        }
        for i in range(1, 6)
    ]
    expected_url = (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime+desc&$top=2&$filter=isRead+eq+false"
    )
    route = respx.get(expected_url).mock(
        return_value=httpx.Response(200, json={"value": messages})
    )

    async with GraphClient(config) as client:
        result = await list_mail(config=config, client=client, unreadOnly=True, top=2)

    assert route.called
    emails: list[dict[str, object]] = cast(list[dict[str, object]], result["emails"])
    assert len(emails) == 2
    assert emails[0]["subject"] == "Email 1"
    assert emails[1]["subject"] == "Email 2"


def test_trusted_sender_address_precedence_from_over_sender():
    """Sender precedence: from.emailAddress.address takes priority over sender.emailAddress.address."""
    msg_both = {
        "from": {"emailAddress": {"address": "From@Example.com", "name": "From User"}},
        "sender": {"emailAddress": {"address": "Sender@Example.com", "name": "Sender User"}},
    }
    assert _trusted_sender_address(msg_both) == "from@example.com"

    msg_only_sender = {
        "from": {"emailAddress": {"name": "No address field"}},
        "sender": {"emailAddress": {"address": "SenderOnly@Example.com"}},
    }
    assert _trusted_sender_address(msg_only_sender) == "senderonly@example.com"

    msg_only_from = {
        "from": {"emailAddress": {"address": " FromOnly@Example.com "}},
        "sender": None,
    }
    assert _trusted_sender_address(msg_only_from) == "fromonly@example.com"


def test_trusted_sender_address_missing_or_malformed_returns_none():
    """Missing or malformed from + sender returns None → denies access."""
    assert _trusted_sender_address({}) is None
    assert _trusted_sender_address({"from": None}) is None
    assert _trusted_sender_address({"from": {"emailAddress": None}}) is None
    assert _trusted_sender_address({"from": {"emailAddress": {"address": ""}}}) is None
    assert _trusted_sender_address({"from": {"emailAddress": {"address": "   "}}}) is None
    assert _trusted_sender_address({"replyTo": {"emailAddress": {"address": "hacker@example.com"}}}) is None
    assert _trusted_sender_address({"from": "not a dict"}) is None


def test_trusted_sender_address_ignores_reply_to():
    """replyTo must never be used for trust determination."""
    msg = {
        "replyTo": [{"emailAddress": {"address": "admin@example.com"}}],
        "from": {"emailAddress": {"name": "No address"}},
    }
    assert _trusted_sender_address(msg) is None


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_returns_llm_error_when_ctx_is_none(config: MailConfig):
    """get_summary returns SUMMARY_LLM_ERROR when ctx is None."""
    from mail_tools import get_summary

    mock_token()
    respx.get(message_url("msg-1")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-1",
                "subject": "Test",
                "from": {"emailAddress": {"address": "trusted@example.com", "name": "Trusted"}},
                "receivedDateTime": "2026-06-18T10:00:00Z",
                "hasAttachments": False,
                "body": {"content": "Hello"},
            },
        )
    )

    async with GraphClient(config) as client:
        result = await get_summary(config=config, client=client, ctx=None, email_id="msg-1")

    assert result.get("error") == "SUMMARY_LLM_ERROR"


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_returns_schema_error_for_unknown_schema_name(tmp_path, monkeypatch: pytest.MonkeyPatch):
    """get_summary returns SUMMARY_SCHEMA_ERROR for an unknown schema name."""
    import mail_tools as mt_mod
    import summary_schema as ss_mod

    good_cfg = MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"anyone@example.com"},
    )

    mock_token()
    respx.get(message_url("msg-schema")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-schema",
                "subject": "Test Subject",
                "from": {"emailAddress": {"address": "anyone@example.com", "name": "Sender"}},
                "receivedDateTime": "2026-06-18T10:00:00Z",
                "hasAttachments": False,
                "body": {"content": "Hello world"},
            },
        )
    )

    monkeypatch.setattr(mt_mod, "load_summary_schema", lambda name: (_ for _ in ()).throw(
        ss_mod.SummarySchemaError(f"Schema file not found: {name}")
    ))

    async with GraphClient(good_cfg) as client:
        result = await get_summary(config=good_cfg, client=client, ctx=None, email_id="msg-schema", schema_name="nonexistent")

    assert result.get("error") == "SUMMARY_SCHEMA_ERROR"
    assert "nonexistent" in str(result.get("message", ""))


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_works_for_untrusted_sender(tmp_path, monkeypatch: pytest.MonkeyPatch):
    """get_summary works for untrusted senders and sets isAllowedInboundSender=False."""
    import mail_tools as mt_mod
    import summary_schema as ss_mod

    good_cfg = MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"trusted@example.com"},
    )

    fake_spec = ss_mod.SummarySchemaSpec(
        name="general",
        description="General summary",
        system_prompt="Summarize this email.",
        json_schema={"type": "object", "properties": {"summary": {"type": "string"}}, "required": ["summary"], "additionalProperties": False},
    )

    captured_payloads: list[dict] = []

    def fake_summarize(*, ctx, system_prompt, json_schema, payload, model=None, provider=None):
        captured_payloads.append(payload)
        assert isinstance(payload, dict)
        assert payload["email_id"] == "msg-untrusted-summary"
        assert payload["isAllowedInboundSender"] is False
        assert payload["body"] == "Suspicious content here."
        assert payload["hasAttachments"] is False
        return {"status": "ok", "reason": None, "result": {"summary": "Suspicious email detected."}}

    mock_token()
    respx.get(message_url("msg-untrusted-summary")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-untrusted-summary",
                "subject": "Urgent Action Required",
                "from": {"emailAddress": {"address": "stranger@example.com", "name": "Stranger"}},
                "toRecipients": [{"emailAddress": {"address": "user@example.org", "name": "Me"}}],
                "receivedDateTime": "2026-06-18T10:00:00Z",
                "hasAttachments": False,
                "body": {"contentType": "text", "content": "Suspicious content here."},
            },
        )
    )

    # Attachment endpoint must NOT be called
    attachment_route = respx.get(f"{message_url('msg-untrusted-summary')}/attachments").mock(
        return_value=httpx.Response(200, json={"value": []})
    )

    monkeypatch.setattr(mt_mod, "load_summary_schema", lambda name: fake_spec)
    monkeypatch.setattr(mt_mod, "summarize_with_llm", fake_summarize)

    async with GraphClient(good_cfg) as client:
        result = await get_summary(config=good_cfg, client=client, ctx=None, email_id="msg-untrusted-summary", schema_name="general")

    assert not attachment_route.called  # attachments were NOT fetched
    assert result == {"schemaName": "general", "emailId": "msg-untrusted-summary", "summary": {"summary": "Suspicious email detected."}}


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_works_for_trusted_sender(tmp_path, monkeypatch: pytest.MonkeyPatch):
    """get_summary works for trusted senders and sets isAllowedInboundSender=True."""
    import mail_tools as mt_mod
    import summary_schema as ss_mod

    good_cfg = MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"boss@example.com"},
    )

    fake_spec = ss_mod.SummarySchemaSpec(
        name="general",
        description="General summary",
        system_prompt="Summarize this email.",
        json_schema={"type": "object", "properties": {"summary": {"type": "string"}}, "required": ["summary"], "additionalProperties": False},
    )

    captured_payloads: list[dict] = []

    def fake_summarize(*, ctx, system_prompt, json_schema, payload, model=None, provider=None):
        captured_payloads.append(payload)
        return {"status": "ok", "reason": None, "result": {"summary": "Quarterly results are positive."}}

    mock_token()
    respx.get(message_url("msg-trusted-summary")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-trusted-summary",
                "subject": "Q2 Results",
                "from": {"emailAddress": {"address": "boss@example.com", "name": "Boss"}},
                "sender": {"emailAddress": {"address": "boss@example.com", "name": "Boss"}},
                "toRecipients": [{"emailAddress": {"address": "user@example.org", "name": "Me"}}],
                "receivedDateTime": "2026-06-18T12:00:00Z",
                "hasAttachments": True,
                "body": {"contentType": "html", "content": "<p>Results look <b>great</b>.</p>"},
            },
        )
    )

    monkeypatch.setattr(mt_mod, "load_summary_schema", lambda name: fake_spec)
    monkeypatch.setattr(mt_mod, "summarize_with_llm", fake_summarize)

    async with GraphClient(good_cfg) as client:
        result = await get_summary(config=good_cfg, client=client, ctx=None, email_id="msg-trusted-summary")

    assert len(captured_payloads) == 1
    payload = captured_payloads[0]
    assert payload["email_id"] == "msg-trusted-summary"
    assert payload["subject"] == "Q2 Results"
    assert payload["from"] == {"name": "Boss", "address": "boss@example.com"}
    assert payload["to"] == [{"name": "Me", "address": "user@example.org"}]
    assert payload["receivedDateTime"] == "2026-06-18T12:00:00Z"
    assert payload["isAllowedInboundSender"] is True
    assert payload["hasAttachments"] is True
    # Body was sanitized (HTML stripped)
    assert "Results look" in payload["body"]
    assert "<b>" not in payload["body"]

    assert result == {"schemaName": "general", "emailId": "msg-trusted-summary", "summary": {"summary": "Quarterly results are positive."}}


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_propagates_llm_error_code(tmp_path, monkeypatch: pytest.MonkeyPatch):
    """get_summary propagates SummaryLlmError code and message in error envelope."""
    import mail_tools as mt_mod
    import summary_llm as sl_mod
    import summary_schema as ss_mod

    good_cfg = MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={},
    )

    fake_spec = ss_mod.SummarySchemaSpec(
        name="general",
        description="General summary",
        system_prompt="Summarize.",
        json_schema={"type": "object", "properties": {"summary": {"type": "string"}}, "required": ["summary"], "additionalProperties": False},
    )

    def fake_refusal(*, ctx, system_prompt, json_schema, payload, model=None, provider=None):
        raise sl_mod.SummaryLlmError("Model refused to generate summary", code="SUMMARY_REFUSED")

    mock_token()
    respx.get(message_url("msg-refused")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-refused",
                "subject": "Refused",
                "from": {"emailAddress": {"address": "x@example.com"}},
                "receivedDateTime": "2026-06-18T10:00:00Z",
                "hasAttachments": False,
                "body": {"content": "test"},
            },
        )
    )

    monkeypatch.setattr(mt_mod, "load_summary_schema", lambda name: fake_spec)
    monkeypatch.setattr(mt_mod, "summarize_with_llm", fake_refusal)

    async with GraphClient(good_cfg) as client:
        result = await get_summary(config=good_cfg, client=client, ctx=None, email_id="msg-refused")

    assert result == {"error": "SUMMARY_REFUSED", "message": "Model refused to generate summary"}


# ---------------------------------------------------------------------------
# normalize_summary_response tests
# ---------------------------------------------------------------------------


def test_normalize_summary_response_ok_returns_clean_envelope():
    """Status ok with dict result returns clean envelope without internal keys."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-42",
        internal={"status": "ok", "reason": None, "result": {"summary": "Good."}},
    )
    assert result == {
        "schemaName": "general",
        "emailId": "msg-42",
        "summary": {"summary": "Good."},
    }
    # Must NOT contain internal keys
    assert "status" not in result
    assert "reason" not in result


def test_normalize_summary_response_ok_with_nested_result():
    """Status ok with nested dict result passes through intact."""
    result = normalize_summary_response(
        schema_name="detailed",
        email_id="msg-99",
        internal={
            "status": "ok",
            "reason": None,
            "result": {
                "subject": "Q2 Results",
                "sentiment": "positive",
                "actionItems": ["Follow up on budget"],
            },
        },
    )
    assert result["schemaName"] == "detailed"
    assert result["emailId"] == "msg-99"
    assert result["summary"]["subject"] == "Q2 Results"
    assert result["summary"]["actionItems"] == ["Follow up on budget"]


def test_normalize_summary_response_wrong_type_returns_error():
    """Status wrong_type returns WRONG_TYPE error envelope."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-10",
        internal={"status": "wrong_type", "reason": "Content is a calendar invite.", "result": None},
    )
    assert result == {
        "error": "WRONG_TYPE",
        "message": "Email content does not match the requested summary schema.",
        "schemaName": "general",
        "emailId": "msg-10",
        "reason": "Content is a calendar invite.",
    }


def test_normalize_summary_response_wrong_type_missing_reason_returns_invalid():
    """Status wrong_type with missing/null reason returns SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-11",
        internal={"status": "wrong_type", "reason": None, "result": None},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "reason is missing or invalid" in str(result.get("message", ""))


def test_normalize_summary_response_ok_null_result_falls_through():
    """Status ok with null result falls through to SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-bad",
        internal={"status": "ok", "reason": None, "result": None},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "missing or invalid" in str(result.get("message", ""))
    assert result["schemaName"] == "general"
    assert result["emailId"] == "msg-bad"


def test_normalize_summary_response_ok_list_result_falls_through():
    """Status ok with list result falls through to SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-bad-2",
        internal={"status": "ok", "reason": None, "result": [{"key": "val"}]},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "missing or invalid" in str(result.get("message", ""))


def test_normalize_summary_response_unknown_status_falls_through():
    """Unknown status falls through to SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-err",
        internal={"status": "error", "reason": "timeout", "result": None},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "status='error'" in str(result.get("message", ""))
    assert result["schemaName"] == "general"
    assert result["emailId"] == "msg-err"


def test_normalize_summary_response_empty_dict_falls_through():
    """Empty internal dict falls through to SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="detailed",
        email_id="msg-empty",
        internal={},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "status=''" in str(result.get("message", ""))
    assert result["schemaName"] == "detailed"
    assert result["emailId"] == "msg-empty"


def test_normalize_summary_response_no_alternate_schemas_in_wrong_type():
    """Wrong_type error must NOT suggest alternate schema names."""
    result = normalize_summary_response(
        schema_name="invoice",
        email_id="msg-inv",
        internal={"status": "wrong_type", "reason": "Not an invoice.", "result": None},
    )
    assert result["error"] == "WRONG_TYPE"
    msg_text = str(result.get("message", ""))
    assert "alternate" not in msg_text.lower()
    assert "suggest" not in msg_text.lower()
    # Only contains: error, message, schemaName, emailId, reason
    assert set(result.keys()) == {"error", "message", "schemaName", "emailId", "reason"}


@pytest.mark.asyncio
@respx.mock
async def test_get_summary_uses_normalize_summary_response(tmp_path, monkeypatch: pytest.MonkeyPatch):
    """get_summary normalizes the internal wrapper before returning."""
    import mail_tools as mt_mod
    import summary_schema as ss_mod

    good_cfg = MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"trusted@example.com"},
    )

    fake_spec = ss_mod.SummarySchemaSpec(
        name="general",
        description="General summary",
        system_prompt="Summarize this email.",
        json_schema={"type": "object", "properties": {"summary": {"type": "string"}}, "required": ["summary"], "additionalProperties": False},
    )

    def fake_summarize(*, ctx, system_prompt, json_schema, payload, model=None, provider=None):
        return {"status": "wrong_type", "reason": "Calendar event, not email.", "result": None}

    mock_token()
    respx.get(message_url("msg-wrong-type")).mock(
        return_value=httpx.Response(
            200,
            json={
                "id": "msg-wrong-type",
                "subject": "Team Lunch",
                "from": {"emailAddress": {"address": "calendar@microsoft.com"}},
                "receivedDateTime": "2026-06-18T10:00:00Z",
                "hasAttachments": False,
                "body": {"content": "Meeting at noon"},
            },
        )
    )

    monkeypatch.setattr(mt_mod, "load_summary_schema", lambda name: fake_spec)
    monkeypatch.setattr(mt_mod, "summarize_with_llm", fake_summarize)

    async with GraphClient(good_cfg) as client:
        result = await get_summary(config=good_cfg, client=client, ctx=None, email_id="msg-wrong-type")

    assert result == {
        "error": "WRONG_TYPE",
        "message": "Email content does not match the requested summary schema.",
        "schemaName": "general",
        "emailId": "msg-wrong-type",
        "reason": "Calendar event, not email.",
    }
    # Internal keys must NOT leak
    assert "status" not in result
    assert "result" not in result


def test_normalize_summary_response_unwraps_llm_wrapper():
    """normalize_summary_response unwraps summarize_with_llm wrapper."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-42",
        internal={"status": "success", "data": {"status": "ok", "reason": None, "result": {"summary": "Good."}}},
    )
    assert result == {
        "schemaName": "general",
        "emailId": "msg-42",
        "summary": {"summary": "Good."},
    }


def test_normalize_summary_response_unwraps_llm_wrapper_wrong_type():
    """Unwrap LLM wrapper with wrong_type status."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-43",
        internal={"status": "success", "data": {"status": "wrong_type", "reason": "Calendar invite.", "result": None}},
    )
    assert result["error"] == "WRONG_TYPE"
    assert result["reason"] == "Calendar invite."


def test_normalize_summary_response_non_dict_data_fails():
    """LLM wrapper with non-dict data returns SUMMARY_INVALID_RESPONSE."""
    result = normalize_summary_response(
        schema_name="general",
        email_id="msg-44",
        internal={"status": "success", "data": "raw text response"},
    )
    assert result["error"] == "SUMMARY_INVALID_RESPONSE"
    assert "non-JSON" in str(result.get("message", ""))
