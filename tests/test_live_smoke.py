# pyright: reportMissingImports=false, reportMissingParameterType=false, reportUnknownArgumentType=false

import os

import pytest

from m365_email_hermes.config import MailConfig
from m365_email_hermes.graph import GraphClient
from m365_email_hermes.mail_tools import list_mail


REQUIRED_LIVE_ENV = (
    "M365_MAIL_CLIENT_ID",
    "M365_MAIL_CLIENT_SECRET",
    "M365_MAIL_TENANT_ID",
)


def live_tests_enabled() -> bool:
    return os.environ.get("M365_EMAIL_LIVE_TESTS", "").lower() == "true" and all(
        os.environ.get(key) for key in REQUIRED_LIVE_ENV
    )


pytestmark = [
    pytest.mark.skipif(
        not live_tests_enabled(),
        reason="set M365_EMAIL_LIVE_TESTS=true and M365 mail env vars to run live Graph smoke tests",
    ),
]


@pytest.mark.asyncio
async def test_live_list_mail_smoke_returns_collection() -> None:
    config = MailConfig.from_env(load_dotenv=False)

    async with GraphClient(config) as client:
        messages = await list_mail(config=config, client=client, top=1)

    assert isinstance(messages, list)
