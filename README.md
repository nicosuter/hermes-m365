# M365 Email Hermes Plugin

This plugin provides the M365 Email platform adapter and tools for the Hermes gateway, enabling asynchronous email polling and interactive email management.

## Purpose
Integrates Microsoft 365 (M365) email as a platform and toolset within the Hermes ecosystem.

## Installation

Add as a git submodule to your Hermes plugins directory:

```bash
git submodule add https://github.com/nicosuter/hermes-m365 ~/.hermes/hermes-agent/plugins/platforms/m365_email/
```

Then enable in `config.yaml`:

```yaml
platforms:
  m365_email:
    enabled: true
```

## Environment Variables

### Required
| Variable | Description |
| :--- | :--- |
| `M365_MAIL_CLIENT_ID` | Microsoft Entra (Azure AD) Application Client ID |
| `M365_MAIL_CLIENT_SECRET` | Microsoft Entra (Azure AD) Application Client Secret |
| `M365_MAIL_TENANT_ID` | Microsoft Entra (Azure AD) Tenant ID |
| `M365_MAILBOX_USER` | The primary email address to monitor. |

### Optional
| Variable | Description | Default |
| :--- | :--- | :--- |
| `EMAIL_ALLOWED_USERS` | Comma-separated list of email addresses allowed to trigger inbound events. | (None) |
| `M365_ATTACHMENT_MAX_BYTES` | Maximum size for attachment downloads in bytes. | `10485760` (10MB) |
| `M365_EMAIL_STATE_PATH` | Path to the JSON file storing polling state (watermark/processed IDs). | `.runtime/poll-state.json` |
| `M365_POLL_INTERVAL_SECONDS` | Frequency of inbox polling in seconds. | `30` |

## Tool Contracts

- `list_mail(top=50, filter=None, unreadOnly=False)`: List recent emails from the inbox. Returns all senders (subject to inbound drop logic). `unreadOnly` filters to unread messages only.
- `get_email(email_id)`: Retrieve email content. Returns sanitized text and attachment metadata.
- `get_attachment(email_id, attachment_id)`: Download an attachment.
- `send_email(to, subject, body, reply_to=None)`: Send a plain text email. `reply_to` sets the Reply-To email address header.

## Security & Safety

### Inbound Filtering (`EMAIL_ALLOWED_USERS`)
The `EMAIL_ALLOWED_USERS` environment variable defines the single allowlist for inbound interaction:
- **Inbound Drop**: Emails from senders NOT in this list are silently dropped during polling (no MessageEvent, no chat creation, no notification).
- **`get_email`**: Will return a warning if the sender is not in the allowlist.
- **`get_attachment`**: Is strictly gated; requests for attachments from unallowed senders will be rejected with error `ATTACHMENT_BLOCKED_UNTRUSTED_SENDER`.

### Attachment Safety
- `get_email` only provides metadata (filenames, sizes) and sanitized text.
- `get_attachment` downloads files to `~/.hermes/inbox/email/`.
- Downloads are gated by the `EMAIL_ALLOWED_USERS` allowlist.

## Configuration & Behavior
- **Graph Permissions**: Requires `Mail.Read` and `Mail.Send`.
- **Polling**:
  - Uses a watermark to prevent processing historical emails on the first run.
  - Does **not** mark emails as read in M365.
  - Uses a local state file to track processed message IDs.

## Testing
Run the test suite using:
```bash
uv run pytest -q
```

> **Note**: Detailed operational instructions for agents will be documented in `AGENTS.md`.
