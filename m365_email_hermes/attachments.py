# pyright: reportMissingImports=false, reportUnknownVariableType=false

from __future__ import annotations

import re
from pathlib import Path, PurePosixPath, PureWindowsPath

from m365_email_hermes.config import DEFAULT_ATTACHMENT_MAX_BYTES, is_allowed_sender


HERMES_EMAIL_ATTACHMENT_DIR = Path.home() / ".hermes" / "inbox" / "email"
CONTROL_CHAR_RE = re.compile(r"[\x00-\x1f\x7f]")
UNSAFE_FILENAME_CHARS_RE = re.compile(r"[^A-Za-z0-9._ -]+")


def sanitize_filename(filename: str) -> str:
    if not filename or CONTROL_CHAR_RE.search(filename):
        raise ValueError("Attachment filename must not be empty or contain control characters")

    basename = PureWindowsPath(PurePosixPath(filename).name).name.strip().strip(".")
    safe_name = UNSAFE_FILENAME_CHARS_RE.sub("_", basename).strip()
    if not safe_name:
        raise ValueError("Attachment filename must contain safe characters")
    return safe_name[:255]


def build_saved_path(attachment_id: str, filename: str) -> Path:
    if not attachment_id:
        raise ValueError("Attachment id must not be empty")
    safe_prefix = UNSAFE_FILENAME_CHARS_RE.sub("_", attachment_id[:12]).strip() or "attachment"
    return HERMES_EMAIL_ATTACHMENT_DIR / f"{safe_prefix}-{sanitize_filename(filename)}"


def check_attachment_sender(address: str | None, allowed_users: set[str]) -> tuple[bool, str | None]:
    if is_allowed_sender(address, allowed_users):
        return True, None
    return False, "Attachment access denied: sender is not in EMAIL_ALLOWED_USERS"


def enforce_attachment_size(size_bytes: int, max_bytes: int = DEFAULT_ATTACHMENT_MAX_BYTES) -> None:
    if size_bytes < 0:
        raise ValueError("Attachment size must not be negative")
    if size_bytes > max_bytes:
        raise ValueError(f"Attachment is {size_bytes} bytes, exceeding maximum allowed size of {max_bytes} bytes")


def build_inline_attachment_body(body: str, inline_attachments: list[dict[str, object]]) -> str:
    from m365_email_hermes.sanitize import insert_inline_attachment_markers

    return insert_inline_attachment_markers(body, inline_attachments)
