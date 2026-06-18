# pyright: reportMissingImports=false
"""Configuration for the M365 Email Hermes plugin."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path


DEFAULT_ATTACHMENT_MAX_BYTES = 10_485_760
DEFAULT_POLL_TOP = 25
DEFAULT_EMAIL_STATE_PATH = Path(".runtime/poll-state.json")
REQUIRED_ENV_VARS = (
    "M365_MAIL_CLIENT_ID",
    "M365_MAIL_CLIENT_SECRET",
    "M365_MAIL_TENANT_ID",
    "M365_MAILBOX_USER",
    "EMAIL_ALLOWED_USERS",
)


class MailConfigError(RuntimeError):
    """Raised when M365 email configuration is invalid or incomplete."""


@dataclass(frozen=True)
class MailConfig:
    """Runtime configuration for Microsoft Graph mail access."""

    client_id: str
    client_secret: str
    tenant_id: str
    mailbox_user: str
    allowed_users: set[str] = field(default_factory=set)
    attachment_max_bytes: int = DEFAULT_ATTACHMENT_MAX_BYTES
    email_state_path: Path = DEFAULT_EMAIL_STATE_PATH
    poll_top: int = DEFAULT_POLL_TOP

    @classmethod
    def from_env(cls, *, load_dotenv: bool = True, project_root: Path | None = None) -> "MailConfig":
        """Load configuration from environment, optionally reading project-root .env first."""
        if load_dotenv:
            load_project_dotenv(project_root)

        env = os.environ
        missing = [key for key in REQUIRED_ENV_VARS if env.get(key) is None]
        if missing:
            joined = ", ".join(missing)
            raise MailConfigError(f"Missing required environment variables for M365 email: {joined}")

        return cls(
            client_id=env["M365_MAIL_CLIENT_ID"],
            client_secret=env["M365_MAIL_CLIENT_SECRET"],
            tenant_id=env["M365_MAIL_TENANT_ID"],
            mailbox_user=env["M365_MAILBOX_USER"],
            allowed_users=parse_email_allowed_users(env.get("EMAIL_ALLOWED_USERS")),
            attachment_max_bytes=parse_positive_int(
                env,
                "M365_ATTACHMENT_MAX_BYTES",
                DEFAULT_ATTACHMENT_MAX_BYTES,
            ),
            email_state_path=Path(env.get("M365_EMAIL_STATE_PATH", str(DEFAULT_EMAIL_STATE_PATH))),
            poll_top=parse_positive_int(
                env,
                "M365_POLL_TOP",
                DEFAULT_POLL_TOP,
            ),
        )


def project_root_from_module() -> Path:
    """Return the plugin project root (flat layout: module sits at root)."""
    return Path(__file__).resolve().parent


def load_project_dotenv(project_root: Path | None = None) -> None:
    """Load .env from the project root without overriding existing process env."""
    from dotenv import load_dotenv  # lazy: upstream usually handles env
    root = project_root or project_root_from_module()
    _ = load_dotenv(root / ".env", override=False)


def parse_email_allowed_users(raw_value: str | None) -> set[str]:
    """Parse comma-separated allowed sender addresses into lowercase normalized form."""
    if not raw_value:
        return set()
    return {address.strip().lower() for address in raw_value.split(",") if address.strip()}


def is_allowed_sender(address: str | None, allowed_users: set[str]) -> bool:
    """Return True when address appears in the configured allowlist."""
    if not address:
        return False
    return address.strip().lower() in allowed_users


def parse_positive_int(env: Mapping[str, str], key: str, default: int) -> int:
    """Read a positive integer env var with a clear configuration error."""
    raw_value = env.get(key)
    if raw_value is None or raw_value == "":
        return default
    try:
        value = int(raw_value)
    except ValueError as exc:
        raise MailConfigError(f"{key} must be a positive integer") from exc
    if value <= 0:
        raise MailConfigError(f"{key} must be a positive integer")
    return value
