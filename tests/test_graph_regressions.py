"""Offline regression coverage for Graph edge cases."""

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

import httpx
import pytest
import respx

from config import MailConfig
from graph import GRAPH_BASE_URL, GraphClient
from mail_tools import list_mail


@pytest.fixture
def config() -> MailConfig:
    return MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users={"trusted@example.com"},
    )


def mock_token() -> None:
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )


def inbox_url(top: int = 50) -> str:
    return (
        f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages"
        f"?$orderby=receivedDateTime+desc&$top={top}"
    )


@pytest.mark.asyncio
@respx.mock
async def test_graph_403_raises_clear_http_error_without_retry(config: MailConfig):
    mock_token()
    messages_url = f"{GRAPH_BASE_URL}/users/user%40example.org/messages"
    graph_route = respx.get(messages_url).mock(
        return_value=httpx.Response(403, json={"error": {"code": "ErrorAccessDenied", "message": "Access denied"}})
    )

    async with GraphClient(config) as client:
        with pytest.raises(httpx.HTTPStatusError) as exc_info:
            await client.get(messages_url)

    assert exc_info.value.response.status_code == 403
    assert "Forbidden" in str(exc_info.value)
    assert graph_route.call_count == 1


@pytest.mark.asyncio
@respx.mock
async def test_graph_429_raises_clear_http_error_without_retry(config: MailConfig):
    mock_token()
    messages_url = f"{GRAPH_BASE_URL}/users/user%40example.org/messages"
    graph_route = respx.get(messages_url).mock(
        return_value=httpx.Response(429, headers={"Retry-After": "120"}, json={"error": {"code": "TooManyRequests"}})
    )

    async with GraphClient(config) as client:
        with pytest.raises(httpx.HTTPStatusError) as exc_info:
            await client.get(messages_url)

    assert exc_info.value.response.status_code == 429
    assert "Too Many Requests" in str(exc_info.value)
    assert graph_route.call_count == 1


@pytest.mark.asyncio
@respx.mock
async def test_empty_inbox_returns_empty_list(config: MailConfig):
    mock_token()
    route = respx.get(inbox_url()).mock(return_value=httpx.Response(200, json={"value": []}))

    async with GraphClient(config) as client:
        result = await list_mail(config=config, client=client, unreadOnly=False)

    assert route.called
    assert result == []


@pytest.mark.asyncio
@respx.mock
async def test_empty_page_still_follows_odata_next_link(config: MailConfig):
    mock_token()
    second_url = f"{GRAPH_BASE_URL}/users/user%40example.org/mailFolders/inbox/messages?page=2"
    first_route = respx.get(inbox_url(top=1)).mock(
        return_value=httpx.Response(200, json={"value": [], "@odata.nextLink": second_url})
    )
    second_route = respx.get(second_url).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [
                    {
                        "id": "message-2",
                        "subject": "Arrived later",
                        "from": {"emailAddress": {"address": "trusted@example.com", "name": "Trusted"}},
                        "receivedDateTime": "2026-06-17T10:00:00Z",
                        "hasAttachments": False,
                    }
                ]
            },
        )
    )

    async with GraphClient(config) as client:
        result = await list_mail(config=config, client=client, unreadOnly=False, top=1)

    assert first_route.called
    assert second_route.called
    assert result == [
        {
            "id": "message-2",
            "subject": "Arrived later",
            "from": {"name": "Trusted", "address": "trusted@example.com"},
            "receivedDateTime": "2026-06-17T10:00:00Z",
            "hasAttachments": False,
        }
    ]
