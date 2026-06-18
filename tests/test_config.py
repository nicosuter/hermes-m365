"""Tests for M365 email configuration."""

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

from pathlib import Path

import pytest

from m365_email_hermes.config import MailConfig, MailConfigError, is_allowed_sender, parse_email_allowed_users


REQUIRED_ENV = (
    "M365_MAIL_CLIENT_ID",
    "M365_MAIL_CLIENT_SECRET",
    "M365_MAIL_TENANT_ID",
    "M365_MAILBOX_USER",
)


@pytest.fixture(autouse=True)
def clean_mail_env(monkeypatch):
    """Remove mail-related env vars so each test controls its config."""
    for key in (*REQUIRED_ENV, "M365_MAILBOX_USER", "EMAIL_ALLOWED_USERS", "M365_ATTACHMENT_MAX_BYTES", "M365_EMAIL_STATE_PATH"):
        monkeypatch.delenv(key, raising=False)


def set_required_env(monkeypatch):
    monkeypatch.setenv("M365_MAIL_CLIENT_ID", "client-id")
    monkeypatch.setenv("M365_MAIL_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("M365_MAIL_TENANT_ID", "tenant-id")
    monkeypatch.setenv("M365_MAILBOX_USER", "user@example.org")
    monkeypatch.setenv("EMAIL_ALLOWED_USERS", "trusted@example.com")


def test_required_env_vars_missing_raises_clear_error():
    with pytest.raises(MailConfigError) as exc_info:
        MailConfig.from_env(load_dotenv=False)

    message = str(exc_info.value)
    assert "Missing required environment variables" in message
    for key in REQUIRED_ENV:
        assert key in message


def test_allowed_users_matches_case_insensitively(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("EMAIL_ALLOWED_USERS", "Trusted@Example.com, friend@example.com")

    config = MailConfig.from_env(load_dotenv=False)

    assert parse_email_allowed_users("Trusted@Example.com, friend@example.com") == {
        "trusted@example.com",
        "friend@example.com",
    }
    assert is_allowed_sender("trusted@example.com", config.allowed_users)
    assert is_allowed_sender("TRUSTED@EXAMPLE.COM", config.allowed_users)
    assert is_allowed_sender("friend@example.com", config.allowed_users)
    assert not is_allowed_sender("stranger@example.com", config.allowed_users)


def test_missing_allowed_users_raises_error(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.delenv("EMAIL_ALLOWED_USERS", raising=False)

    with pytest.raises(MailConfigError) as exc_info:
        MailConfig.from_env(load_dotenv=False)

    assert "EMAIL_ALLOWED_USERS" in str(exc_info.value)


def test_mailbox_user_is_required_env_var(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("M365_MAILBOX_USER", "user@example.org")

    config = MailConfig.from_env(load_dotenv=False)

    assert config.mailbox_user == "user@example.org"


def test_optional_defaults(monkeypatch):
    set_required_env(monkeypatch)

    config = MailConfig.from_env(load_dotenv=False)

    assert config.attachment_max_bytes == 10_485_760
    assert config.email_state_path == Path(".runtime/poll-state.json")
