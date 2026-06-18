# pyright: reportMissingImports=false, reportMissingParameterType=false, reportUnknownArgumentType=false, reportUnknownMemberType=false, reportUnknownParameterType=false, reportUnknownVariableType=false, reportUnusedCallResult=false

import importlib
from pathlib import Path
from typing import Protocol
from typing import cast

import pytest

from m365_email_hermes.config import MailConfig, parse_email_allowed_users


class _AttachmentsModule(Protocol):
    HERMES_EMAIL_ATTACHMENT_DIR: Path

    def sanitize_filename(self, filename: str) -> str: ...

    def build_saved_path(self, attachment_id: str, filename: str) -> Path: ...

    def check_attachment_sender(self, address: str | None, allowed_users: set[str]) -> tuple[bool, str | None]: ...

    def enforce_attachment_size(self, size_bytes: int, max_bytes: int = ...) -> None: ...

    def build_inline_attachment_body(self, body: str, inline_attachments: list[dict[str, object]]) -> str: ...


def reload_attachments_with_home(monkeypatch: pytest.MonkeyPatch, home: Path) -> _AttachmentsModule:
    monkeypatch.setenv("HOME", str(home))
    import m365_email_hermes.attachments as attachments

    return cast(_AttachmentsModule, cast(object, importlib.reload(attachments)))


def test_attachment_directory_is_fixed_under_home(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    assert attachments.HERMES_EMAIL_ATTACHMENT_DIR == tmp_path / ".hermes" / "inbox" / "email"


def test_filename_path_traversal_saves_as_safe_basename_under_fixed_dir(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    path = attachments.build_saved_path("att-regular-1", "../../secret.txt")

    assert path == tmp_path / ".hermes" / "inbox" / "email" / "att-regular--secret.txt"
    assert path.parent == attachments.HERMES_EMAIL_ATTACHMENT_DIR


@pytest.mark.parametrize(
    ("filename", "expected"),
    [
        ("invoice.pdf", "invoice.pdf"),
        ("weird:file?.txt", "weird_file_.txt"),
        ("subdir/report.csv", "report.csv"),
        (r"C:\temp\scan.png", "scan.png"),
    ],
)
def test_sanitize_filename_returns_safe_basename(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, filename: str, expected: str):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    assert attachments.sanitize_filename(filename) == expected


@pytest.mark.parametrize("filename", ["", "\x00evil.txt", "bad\nname.txt", "...."])
def test_sanitize_filename_rejects_empty_control_or_dot_only_names(monkeypatch: pytest.MonkeyPatch, tmp_path: Path, filename: str):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    with pytest.raises(ValueError):
        attachments.sanitize_filename(filename)


def test_build_saved_path_uses_first_twelve_attachment_id_chars(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    path = attachments.build_saved_path("att-regular-1", "invoice.pdf")

    assert path == attachments.HERMES_EMAIL_ATTACHMENT_DIR / "att-regular--invoice.pdf"


def test_build_saved_path_is_deterministic_and_handles_filename_collisions(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    first = attachments.build_saved_path("attachment-one-abc", "invoice.pdf")
    first_again = attachments.build_saved_path("attachment-one-abc", "invoice.pdf")
    second = attachments.build_saved_path("attachment-two-xyz", "invoice.pdf")

    assert first == first_again
    assert first != second
    assert first.name == "attachment-o-invoice.pdf"
    assert second.name == "attachment-t-invoice.pdf"


def test_sender_gating_uses_email_allowed_users(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)
    allowed_users = parse_email_allowed_users("trusted@example.com")

    assert attachments.check_attachment_sender("TRUSTED@example.com", allowed_users) == (True, None)
    allowed, error = attachments.check_attachment_sender("stranger@example.com", allowed_users)
    assert not allowed
    assert error == "Attachment access denied: sender is not in EMAIL_ALLOWED_USERS"


def test_trusted_senders_env_var_is_not_read(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    reload_attachments_with_home(monkeypatch, tmp_path)
    monkeypatch.setenv("M365_MAIL_CLIENT_ID", "client-id")
    monkeypatch.setenv("M365_MAIL_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("M365_MAIL_TENANT_ID", "tenant-id")
    monkeypatch.setenv("M365_MAILBOX_USER", "user@example.org")
    monkeypatch.setenv("EMAIL_ALLOWED_USERS", "allowed@example.com")
    monkeypatch.setenv("TRUSTED_SENDERS", "trusted@example.com")

    config = MailConfig.from_env(load_dotenv=False)

    assert config.allowed_users == {"allowed@example.com"}
    assert "trusted@example.com" not in config.allowed_users


def test_default_max_size_allows_limit_and_rejects_excess(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    attachments.enforce_attachment_size(10_485_760)
    with pytest.raises(ValueError, match="exceeding maximum allowed size"):
        attachments.enforce_attachment_size(10_485_761)


def test_configured_max_size_is_enforced(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    attachments.enforce_attachment_size(5, max_bytes=5)
    with pytest.raises(ValueError, match="6 bytes"):
        attachments.enforce_attachment_size(6, max_bytes=5)


def test_inline_marker_helper_delegates_cid_replacement(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    attachments = reload_attachments_with_home(monkeypatch, tmp_path)

    text = attachments.build_inline_attachment_body(
        "Logo cid:logo123",
        [{"contentId": "logo123", "name": "logo.png", "contentType": "image/png"}],
    )

    assert text == 'Logo [Inline attachment called "logo.png" (image/png), use get_attachment to fetch]'
