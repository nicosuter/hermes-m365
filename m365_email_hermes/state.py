"""Polling state persistence for M365 Email adapter."""

from __future__ import annotations

import json
from pathlib import Path
from typing import cast


_MAX_PROCESSED_IDS = 500


class PollState:
    """Tracks processed message IDs and the latest watermark timestamp."""

    def __init__(self) -> None:
        self._processed_ids: list[str] = []
        self._watermark: str = ""

    @property
    def watermark(self) -> str:
        """Latest receivedDateTime watermark (ISO 8601)."""
        return self._watermark

    @watermark.setter
    def watermark(self, value: str) -> None:
        self._watermark = value

    def is_processed(self, message_id: str) -> bool:
        """Return True if message_id has already been processed."""
        return message_id in self._processed_ids

    def add(self, message_id: str, timestamp: str) -> None:
        """Record a newly processed message ID and update watermark."""
        self._processed_ids.append(message_id)
        # Keep only last _MAX_PROCESSED_IDS
        if len(self._processed_ids) > _MAX_PROCESSED_IDS:
            self._processed_ids = self._processed_ids[-_MAX_PROCESSED_IDS:]
        # Update watermark to latest timestamp
        if timestamp > self._watermark:
            self._watermark = timestamp

    def save(self, path: Path) -> None:
        """Persist state to JSON file."""
        path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "processed_ids": self._processed_ids,
            "watermark": self._watermark,
        }
        path.write_text(json.dumps(data, indent=2))

    @classmethod
    def load(cls, path: Path) -> "PollState":
        """Load state from JSON file. Returns empty state if file does not exist."""
        instance = cls()
        if not path.exists():
            return instance
        try:
            data = cast(dict[str, object], json.loads(path.read_text()))
            raw_ids = data.get("processed_ids", [])
            if isinstance(raw_ids, list):
                instance._processed_ids = [str(item) for item in raw_ids if isinstance(item, str)]
            raw_watermark = data.get("watermark", "")
            if isinstance(raw_watermark, str):
                instance._watermark = raw_watermark
        except (json.JSONDecodeError, OSError):
            pass
        return instance
