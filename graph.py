"""Microsoft Graph client for app-only M365 mail access."""

from __future__ import annotations

import time
from collections.abc import AsyncIterator
from typing import cast
from urllib.parse import quote

import httpx

from config import MailConfig


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_EXPIRY_SKEW_SECONDS = 60


class GraphClient:
    """Small async Microsoft Graph client scoped to one configured mailbox user."""

    def __init__(self, config: MailConfig, *, http_client: httpx.AsyncClient | None = None) -> None:
        self.config: MailConfig = config
        self._client: httpx.AsyncClient = http_client or httpx.AsyncClient()
        self._owns_client: bool = http_client is None
        self._access_token: str | None = None
        self._token_expires_at: float = 0.0

    async def __aenter__(self) -> "GraphClient":
        return self

    async def __aexit__(self, exc_type: object, exc: object, traceback: object) -> None:
        await self.aclose()

    async def aclose(self) -> None:
        """Close the underlying HTTP client when this instance created it."""
        if self._owns_client:
            await self._client.aclose()

    async def acquire_token(self, *, force_refresh: bool = False) -> str:
        """Acquire an app-only Microsoft identity token with client credentials."""
        now = time.time()
        if not force_refresh and self._access_token and now < self._token_expires_at:
            return self._access_token

        response = await self._client.post(
            self.token_url,
            data={
                "grant_type": "client_credentials",
                "client_id": self.config.client_id,
                "client_secret": self.config.client_secret,
                "scope": GRAPH_SCOPE,
            },
        )
        _ = response.raise_for_status()
        payload = cast(dict[str, object], response.json())
        token = str(payload["access_token"])
        raw_expires_in = payload.get("expires_in", 3600)
        expires_in = int(raw_expires_in) if isinstance(raw_expires_in, int | str) else 3600
        self._access_token = token
        self._token_expires_at = now + max(0, expires_in - TOKEN_EXPIRY_SKEW_SECONDS)
        return token

    @property
    def token_url(self) -> str:
        """Microsoft identity platform client-credentials endpoint for this tenant."""
        tenant = quote(self.config.tenant_id, safe="")
        return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

    def mail_url(self, path: str) -> str:
        """Build a Graph mail URL under the configured mailbox user's route."""
        normalized_path = path.strip()
        if not normalized_path:
            raise ValueError("Graph mail path must not be empty")
        if normalized_path.startswith(("http://", "https://")):
            raise ValueError("Graph mail path must be relative to the configured /users/ mailbox route")

        normalized_path = normalized_path.lstrip("/")
        if normalized_path == "me" or normalized_path.startswith("me/"):
            raise ValueError("Graph mail paths must use the configured /users/ mailbox route")
        if normalized_path.startswith("users/"):
            raise ValueError("Pass a mailbox-relative path; GraphClient adds the configured /users/ route")

        mailbox_user = quote(self.config.mailbox_user, safe="")
        return f"{GRAPH_BASE_URL}/users/{mailbox_user}/{normalized_path}"

    async def get(self, url: str) -> httpx.Response:
        """GET a Graph URL with bearer auth, refreshing once on 401."""
        response = await self._request("GET", url)
        if response.status_code == 401:
            response = await self._request("GET", url, force_refresh=True)
        _ = response.raise_for_status()
        return response

    async def post(self, url: str, json: dict[str, object]) -> httpx.Response:
        """POST JSON to a Graph URL with bearer auth, refreshing once on 401."""
        response = await self._request("POST", url, json=json)
        if response.status_code == 401:
            response = await self._request("POST", url, json=json, force_refresh=True)
        _ = response.raise_for_status()
        return response

    async def patch(self, url: str, json: dict[str, object]) -> httpx.Response:
        """PATCH JSON to a Graph URL with bearer auth, refreshing once on 401."""
        response = await self._request("PATCH", url, json=json)
        if response.status_code == 401:
            response = await self._request("PATCH", url, json=json, force_refresh=True)
        _ = response.raise_for_status()
        return response

    async def paginate(self, url: str) -> AsyncIterator[dict[str, object]]:
        """Yield all items from a Graph collection following @odata.nextLink."""
        next_url: str | None = url
        while next_url:
            response = await self.get(next_url)
            payload = cast(dict[str, object], response.json())
            raw_value = payload.get("value", [])
            value: list[object] = cast(list[object], raw_value) if isinstance(raw_value, list) else []
            for item in value:
                if isinstance(item, dict):
                    yield cast(dict[str, object], item)
            next_link = payload.get("@odata.nextLink")
            next_url = next_link if isinstance(next_link, str) else None

    async def _request(
        self,
        method: str,
        url: str,
        *,
        force_refresh: bool = False,
        json: dict[str, object] | None = None,
    ) -> httpx.Response:
        token = await self.acquire_token(force_refresh=force_refresh)
        headers = {
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.body-content-type="text"',
        }
        return await self._client.request(method, url, headers=headers, json=json)
