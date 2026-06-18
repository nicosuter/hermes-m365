"""Tests for Microsoft Graph authentication and URL handling."""

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

import httpx
import pytest
import respx

from config import MailConfig
from graph import GRAPH_BASE_URL, GraphClient


@pytest.fixture
def config():
    return MailConfig(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        mailbox_user="user@example.org",
        allowed_users=set(),
    )


@pytest.mark.asyncio
@respx.mock
async def test_token_request_uses_client_credentials(config):
    token_route = respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )

    async with GraphClient(config) as client:
        token = await client.acquire_token()

    assert token == "token-1"
    request = token_route.calls.last.request
    body = dict(item.split("=") for item in request.content.decode().split("&"))
    assert body["grant_type"] == "client_credentials"
    assert body["client_id"] == "client-id"
    assert body["client_secret"] == "client-secret"
    assert body["scope"] == "https%3A%2F%2Fgraph.microsoft.com%2F.default"


def test_mail_url_builder_uses_configured_user_not_me(config):
    client = GraphClient(config)

    url = client.mail_url("mailFolders/inbox/messages")

    assert url == "https://graph.microsoft.com/v1.0/users/user%40example.org/mailFolders/inbox/messages"
    assert "/users/user%40example.org/" in url
    assert "graph.microsoft.com/v1.0/me" not in url
    assert "graph.microsoft.com/v1.0/users/me" not in url


def test_mail_url_builder_rejects_me_paths(config):
    client = GraphClient(config)

    with pytest.raises(ValueError, match="/users/"):
        client.mail_url("/me/messages")


@pytest.mark.asyncio
@respx.mock
async def test_get_refreshes_token_once_on_unauthorized(config):
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        side_effect=[
            httpx.Response(200, json={"access_token": "expired-token", "expires_in": 3600}),
            httpx.Response(200, json={"access_token": "fresh-token", "expires_in": 3600}),
        ]
    )
    messages_url = f"{GRAPH_BASE_URL}/users/user%40example.org/messages"
    graph_route = respx.get(messages_url).mock(
        side_effect=[
            httpx.Response(401, json={"error": "expired"}),
            httpx.Response(200, json={"value": [{"id": "message-1"}]}),
        ]
    )

    async with GraphClient(config) as client:
        response = await client.get(messages_url)

    assert response.json() == {"value": [{"id": "message-1"}]}
    assert graph_route.calls[0].request.headers["authorization"] == "Bearer expired-token"
    assert graph_route.calls[1].request.headers["authorization"] == "Bearer fresh-token"


@pytest.mark.asyncio
@respx.mock
async def test_post_sends_auth_header_and_json(config):
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )
    send_url = f"{GRAPH_BASE_URL}/users/user%40example.org/sendMail"
    send_route = respx.post(send_url).mock(return_value=httpx.Response(202, json={}))

    async with GraphClient(config) as client:
        response = await client.post(send_url, json={"message": {"subject": "Hello"}})

    assert response.status_code == 202
    request = send_route.calls.last.request
    assert request.headers["authorization"] == "Bearer token-1"
    assert request.read() == b'{"message":{"subject":"Hello"}}'


@pytest.mark.asyncio
@respx.mock
async def test_pagination_follows_odata_next_link(config):
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        return_value=httpx.Response(200, json={"access_token": "token-1", "expires_in": 3600})
    )
    first_url = f"{GRAPH_BASE_URL}/users/user%40example.org/messages"
    second_url = f"{GRAPH_BASE_URL}/users/user%40example.org/messages/page-2"
    first_route = respx.get(first_url).mock(
        return_value=httpx.Response(
            200,
            json={
                "value": [{"id": "message-1"}],
                "@odata.nextLink": second_url,
            },
        )
    )
    second_route = respx.get(second_url).mock(
        return_value=httpx.Response(200, json={"value": [{"id": "message-2"}]})
    )

    async with GraphClient(config) as client:
        items = [item async for item in client.paginate(first_url)]

    assert items == [{"id": "message-1"}, {"id": "message-2"}]
    assert first_route.called
    assert second_route.called


@pytest.mark.asyncio
@respx.mock
async def test_post_refreshes_token_once_on_unauthorized(config):
    """BUG 2 fix: post() must refresh on 401, just like get()."""
    respx.post("https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token").mock(
        side_effect=[
            httpx.Response(200, json={"access_token": "expired-token", "expires_in": 3600}),
            httpx.Response(200, json={"access_token": "fresh-token", "expires_in": 3600}),
        ]
    )
    send_url = f"{GRAPH_BASE_URL}/users/user%40example.org/sendMail"
    graph_route = respx.post(send_url).mock(
        side_effect=[
            httpx.Response(401, json={"error": "expired"}),
            httpx.Response(202, json={}),
        ]
    )

    async with GraphClient(config) as client:
        response = await client.post(send_url, json={"message": {"subject": "Hello"}})

    assert response.status_code == 202
    assert graph_route.calls[0].request.headers["authorization"] == "Bearer expired-token"
    assert graph_route.calls[1].request.headers["authorization"] == "Bearer fresh-token"
