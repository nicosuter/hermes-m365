import sys
from unittest.mock import MagicMock

_gateway = MagicMock()
_gateway.platforms = MagicMock()
_gateway.platforms.base = MagicMock()
_gateway.config = MagicMock()


class MockBasePlatformAdapter:
    def __init__(self, config, platform):
        self._config = config
        self._platform = platform
        self._connected = False

    def _mark_connected(self):
        self._connected = True

    def _mark_disconnected(self):
        self._connected = False

    @property
    def is_connected(self):
        return self._connected

    def build_source(self, *, chat_id, chat_name, chat_type, user_id, user_name):
        return {
            "platform": "m365_email",
            "chat_id": chat_id,
            "chat_name": chat_name,
            "chat_type": chat_type,
            "user_id": user_id,
            "user_name": user_name,
        }

    def _set_fatal_error(self, code, message, retryable=False):
        pass

    async def handle_message(self, event):
        pass


class MockMessageEvent(dict):
    def __init__(self, text=None, message_type=None, source=None, message_id=None, **kwargs):
        super().__init__()
        self["text"] = text
        self["message_type"] = message_type
        self["source"] = source
        self["message_id"] = message_id
        self.update(kwargs)

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError:
            raise AttributeError(name)

    def __setattr__(self, name, value):
        self[name] = value


class MockSendResult:
    def __init__(self, success=False, message_id=None):
        self.success = success
        self.message_id = message_id


class MockMessageType:
    TEXT = "text"
    EMAIL = "email"


_gateway.platforms.base.BasePlatformAdapter = MockBasePlatformAdapter
_gateway.platforms.base.MessageEvent = MockMessageEvent
_gateway.platforms.base.MessageType = MockMessageType
_gateway.platforms.base.SendResult = MockSendResult
_gateway.config.Platform = lambda name: type("Platform", (), {"value": name})()

# Inject into sys.modules BEFORE any imports
sys.modules["gateway"] = _gateway
sys.modules["gateway.platforms"] = _gateway.platforms
sys.modules["gateway.platforms.base"] = _gateway.platforms.base
sys.modules["gateway.config"] = _gateway.config

# Also remove stale cached __init__ module if it exists
if "tests" in sys.modules:
    del sys.modules["tests"]
if "tests.__init__" in sys.modules:
    del sys.modules["tests.__init__"]

import importlib

import pytest


@pytest.fixture
def env(monkeypatch):
    monkeypatch.setenv("M365_MAIL_CLIENT_ID", "client-id")
    monkeypatch.setenv("M365_MAIL_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("M365_MAIL_TENANT_ID", "tenant-id")
    monkeypatch.setenv("M365_MAILBOX_USER", "user@example.org")
    monkeypatch.setenv("EMAIL_ALLOWED_USERS", "trusted@example.com")
    monkeypatch.setenv("M365_EMAIL_STATE_PATH", "/tmp/test-state.json")


@pytest.fixture
def state_path(tmp_path, monkeypatch):
    path = tmp_path / "test-state.json"
    monkeypatch.setenv("M365_EMAIL_STATE_PATH", str(path))
    return path
