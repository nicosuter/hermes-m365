"""Tests for M365 email configuration."""

# pyright: reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnusedCallResult=false

from pathlib import Path

import pytest

from config import (
    MailConfig,
    MailConfigError,
    get_summary_model,
    get_summary_provider,
    is_allowed_sender,
    parse_email_allowed_users,
)


REQUIRED_ENV = (
    "M365_MAIL_CLIENT_ID",
    "M365_MAIL_CLIENT_SECRET",
    "M365_MAIL_TENANT_ID",
    "M365_MAILBOX_USER",
)


SUMMARY_ENV_VARS = (
    "M365_SUMMARY_MODEL",
)

TIMEOUT_ENV_VARS = (
    "M365_REQUEST_TIMEOUT",
)


@pytest.fixture(autouse=True)
def clean_mail_env(monkeypatch):
    """Remove mail-related env vars so each test controls its config."""
    for key in (*REQUIRED_ENV, "M365_MAILBOX_USER", "EMAIL_ALLOWED_USERS", "M365_ATTACHMENT_MAX_BYTES", "M365_EMAIL_STATE_PATH", *SUMMARY_ENV_VARS, *TIMEOUT_ENV_VARS):
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


def test_summary_model_populates_correctly(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("M365_SUMMARY_MODEL", "gpt-4o-mini")

    config = MailConfig.from_env(load_dotenv=False)

    assert get_summary_model(config) == "gpt-4o-mini"


def test_summary_model_defaults_to_none(monkeypatch):
    set_required_env(monkeypatch)

    config = MailConfig.from_env(load_dotenv=False)

    assert get_summary_model(config) is None


def test_summary_model_missing_does_not_break_from_env(monkeypatch):
    set_required_env(monkeypatch)

    config = MailConfig.from_env(load_dotenv=False)

    assert get_summary_model(config) is None


def test_summary_provider_populates_correctly(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("M365_SUMMARY_PROVIDER", "openai")

    config = MailConfig.from_env(load_dotenv=False)

    assert get_summary_provider(config) == "openai"


def test_summary_provider_defaults_to_none(monkeypatch):
    set_required_env(monkeypatch)

    config = MailConfig.from_env(load_dotenv=False)

    assert get_summary_provider(config) is None


def test_request_timeout_defaults_to_30(monkeypatch):
    set_required_env(monkeypatch)

    config = MailConfig.from_env(load_dotenv=False)

    assert config.request_timeout == 30.0


def test_request_timeout_populates_from_env(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("M365_REQUEST_TIMEOUT", "45.5")

    config = MailConfig.from_env(load_dotenv=False)

    assert config.request_timeout == 45.5


def test_request_timeout_invalid_value_raises_error(monkeypatch):
    set_required_env(monkeypatch)
    monkeypatch.setenv("M365_REQUEST_TIMEOUT", "not-a-number")

    with pytest.raises(MailConfigError) as exc_info:
        MailConfig.from_env(load_dotenv=False)

    assert "M365_REQUEST_TIMEOUT" in str(exc_info.value)
    assert "positive number" in str(exc_info.value)
