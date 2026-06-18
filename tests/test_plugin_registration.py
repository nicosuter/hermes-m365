"""Tests for plugin registration (plugin.yaml + adapter.py) and Hermes contract compliance."""

from __future__ import annotations

import os
from pathlib import Path

import pytest


from typing import Callable


# ── Fake context ───────────────────────────────────────────────────────────

class FakeContext:
    """Records register_platform() and register_tool() calls."""

    def __init__(self) -> None:
        self.platform_kwargs: dict[str, object] | None = None
        self.tools: list[tuple[str, object, str]] = []

    def register_platform(self, **kwargs: object) -> None:
        self.platform_kwargs = kwargs

    def register_tool(
        self,
        name: str,
        toolset: str,
        schema: dict,
        handler: Callable,
        **kwargs: object,
    ) -> None:
        self.tools.append((name, handler, schema.get("description", "")))


# ── Plugin YAML tests ──────────────────────────────────────────────────────

def _load_plugin_yaml() -> str:
    plugin_root = Path(__file__).resolve().parents[1]
    return (plugin_root / "plugin.yaml").read_text()


class TestPluginYaml:
    def test_kind_is_platform(self) -> None:
        content = _load_plugin_yaml()
        assert "kind: platform" in content

    def test_name_is_m365_email(self) -> None:
        content = _load_plugin_yaml()
        assert "name: m365_email" in content

    def test_surfaces_email_allowed_users(self) -> None:
        content = _load_plugin_yaml()
        assert "EMAIL_ALLOWED_USERS" in content

    def test_does_not_contain_trusted_senders(self) -> None:
        content = _load_plugin_yaml()
        assert "TRUSTED_SENDERS" not in content

    def test_requires_env_keys(self) -> None:
        content = _load_plugin_yaml()
        assert "M365_MAIL_CLIENT_ID" in content
        assert "M365_MAIL_CLIENT_SECRET" in content
        assert "M365_MAIL_TENANT_ID" in content


# ── Adapter registration tests ─────────────────────────────────────────────

class TestAdapterRegistration:
    def test_register_sets_platform_name(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        assert ctx.platform_kwargs is not None
        assert ctx.platform_kwargs["name"] == "m365_email"

    def test_registers_all_thirteen_tools(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        tool_names = {name for name, _, _ in ctx.tools}
        expected = {
            "list_mail", "get_email", "get_attachment", "send_email",
            "reply_email", "reply_all", "forward_email",
            "mark_read", "mark_unread",
            "confirm_send_email", "confirm_reply_email", "confirm_reply_all", "confirm_forward_email",
        }
        assert tool_names == expected

    def test_tool_descriptions_present(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        for _, _, desc in ctx.tools:
            assert desc and len(desc) > 0

    def test_required_env_in_platform(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        assert ctx.platform_kwargs is not None
        required = ctx.platform_kwargs["required_env"]
        assert isinstance(required, list)
        names = [r["name"] if isinstance(r, dict) else r for r in required]
        assert "M365_MAIL_CLIENT_ID" in names
        assert "M365_MAIL_CLIENT_SECRET" in names
        assert "M365_MAIL_TENANT_ID" in names


# ── Validation tests ───────────────────────────────────────────────────────

class TestValidateConfig:
    def test_returns_false_when_client_secret_missing(self) -> None:
        from adapter import validate_config

        env_patch = {
            "M365_MAIL_CLIENT_ID": "ok",
            "M365_MAIL_TENANT_ID": "ok",
        }
        with patch_environ(env_patch, clear_required=True):
            assert validate_config() is False

    def test_returns_true_when_all_required_present(self) -> None:
        from adapter import validate_config

        env_patch = {
            "M365_MAIL_CLIENT_ID": "id",
            "M365_MAIL_CLIENT_SECRET": "secret",
            "M365_MAIL_TENANT_ID": "tenant",
            "M365_MAILBOX_USER": "user@example.org",
            "EMAIL_ALLOWED_USERS": "trusted@example.com",
        }
        with patch_environ(env_patch, clear_required=False):
            assert validate_config() is True


# ── Hermes Contract Tests ──────────────────────────────────────────────────

class TestHermesContract:
    def test_adapter_extends_base_platform_adapter(self) -> None:
        from adapter import M365EmailAdapter
        from gateway.platforms.base import BasePlatformAdapter

        assert issubclass(M365EmailAdapter, BasePlatformAdapter)

    def test_adapter_implements_connect_disconnect_send(self) -> None:
        from adapter import M365EmailAdapter

        assert hasattr(M365EmailAdapter, "connect")
        assert hasattr(M365EmailAdapter, "disconnect")
        assert hasattr(M365EmailAdapter, "send")
        assert hasattr(M365EmailAdapter, "get_chat_info")

    def test_register_uses_singular_allowed_users_env(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        assert ctx.platform_kwargs is not None
        assert "allowed_users_env" in ctx.platform_kwargs
        assert ctx.platform_kwargs["allowed_users_env"] == "EMAIL_ALLOWED_USERS"
        assert "allowed_users_envs" not in ctx.platform_kwargs

    def test_register_includes_emoji_and_allow_update_command(self) -> None:
        from adapter import register

        ctx = FakeContext()
        register(ctx)
        assert ctx.platform_kwargs is not None
        assert ctx.platform_kwargs.get("emoji") == "📧"
        assert ctx.platform_kwargs.get("allow_update_command") is True

    def test_env_enablement_returns_none_when_not_configured(self) -> None:
        from adapter import env_enablement

        assert env_enablement() is None

    def test_env_enablement_returns_dict_when_configured(self, monkeypatch) -> None:
        from adapter import env_enablement

        for key in (
            "M365_MAIL_CLIENT_ID",
            "M365_MAIL_CLIENT_SECRET",
            "M365_MAIL_TENANT_ID",
            "M365_MAILBOX_USER",
            "EMAIL_ALLOWED_USERS",
        ):
            monkeypatch.setenv(key, "ok")

        result = env_enablement()
        assert isinstance(result, dict)

    def test_adapter_has_base_class_lifecycle_methods(self) -> None:
        from adapter import M365EmailAdapter
        from gateway.platforms.base import BasePlatformAdapter

        adapter = M365EmailAdapter()
        assert hasattr(adapter, "_mark_connected")
        assert hasattr(adapter, "_mark_disconnected")
        assert hasattr(adapter, "handle_message")

    def test_send_returns_send_result(self) -> None:
        from adapter import M365EmailAdapter
        from gateway.platforms.base import SendResult

        adapter = M365EmailAdapter()
        import asyncio
        result = asyncio.run(adapter.send("m365:test@example.com", "hello"))
        assert isinstance(result, SendResult)
        assert hasattr(result, "success")


# ── Helpers ────────────────────────────────────────────────────────────────

_REQUIRED_KEYS = {"M365_MAIL_CLIENT_ID", "M365_MAIL_CLIENT_SECRET", "M365_MAIL_TENANT_ID", "M365_MAILBOX_USER", "EMAIL_ALLOWED_USERS"}


class patch_environ:
    """Temporarily set/clear env vars for testing."""

    def __init__(self, values: dict[str, str], *, clear_required: bool = False) -> None:
        self._values = values
        self._clear_required = clear_required
        self._saved: dict[str, str | None] = {}

    def __enter__(self) -> "patch_environ":
        # Save originals before modifying anything
        if self._clear_required:
            for key in _REQUIRED_KEYS:
                self._saved[key] = os.environ.pop(key, None)
        for k, v in self._values.items():
            prev = os.environ.get(k)
            if k not in self._saved:
                self._saved[k] = prev
            os.environ[k] = v
        return self

    def __exit__(self, *_args: object) -> None:
        for k, original in self._saved.items():
            if original is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = original
