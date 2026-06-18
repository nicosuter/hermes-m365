"""Test package init — mocks gateway modules before adapter imports."""

import sys
from unittest.mock import MagicMock

# Build mock gateway module tree
_gateway = MagicMock()
_gateway.platforms = MagicMock()
_gateway.platforms.base = MagicMock()
_gateway.config = MagicMock()


class MockBasePlatformAdapter:
    """Minimal BasePlatformAdapter stand-in for plugin tests."""

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

    async def handle_message(self, event):
        pass


class MockMessageEvent(dict):
    """Dict-like MessageEvent that supports both key and attribute access."""

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

sys.modules["gateway"] = _gateway
sys.modules["gateway.platforms"] = _gateway.platforms
sys.modules["gateway.platforms.base"] = _gateway.platforms.base
sys.modules["gateway.config"] = _gateway.config
