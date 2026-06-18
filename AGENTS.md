# AGENTS.md — M365 Email Hermes Plugin

## What This Is

A **Hermes gateway platform plugin** (NOT an MCP server, NOT FastMCP). It treats an M365 mailbox as a messaging platform and registers explicit mail tools.

## Project Layout — Non-Standard

`adapter.py` and `plugin.yaml` **must stay at the project root** for Hermes plugin discovery. The `m365_email_hermes/` package contains all supporting modules. `__init__.py` does a `sys.path` hack to import root-level `adapter.py`.

```
m365-email-hermes-plugin/
├── plugin.yaml              # Hermes manifest — MUST be at root
├── adapter.py               # register(ctx) + M365EmailAdapter — MUST be at root
├── m365_email_hermes/       # package modules
│   ├── config.py            # MailConfig, EMAIL_ALLOWED_USERS parsing
│   ├── graph.py             # GraphClient (token, pagination, /users/{mailbox} routing)
│   ├── mail_tools.py        # list_mail, get_email, get_attachment, send_email, reply, forward, mark read/unread
│   ├── sanitize.py          # HTML→text, hidden content removal
│   ├── attachments.py       # filename safety, sender gating, deterministic paths
│   └── state.py             # PollState (watermark, processed IDs, _MAX_PROCESSED_IDS=500)
└── tests/                   # 98 tests + 1 skipped live smoke (no conftest.py)
```

## Tool Registry

`register(ctx)` wires **10 tools** into Hermes:

- `list_mail(top=50, filter=None, unreadOnly=False)`
- `get_email(email_id)` — returns sanitized text + attachment metadata only (no bytes)
- `get_attachment(email_id, attachment_id)` — downloads to `~/.hermes/inbox/email/`
- `send_email(to, subject, body, reply_to=None)` — requires confirmation token unless `DISABLE_SEND_CONFIRM=true`
- `reply_email(email_id, body)`
- `reply_all(email_id, body)`
- `forward_email(email_id, to, body)`
- `mark_read(email_id)` / `mark_unread(email_id)`
- `confirm_send_email(confirmation_token)` — completes a pending `send_email`

## Critical Constraints

- **Single allowlist:** `EMAIL_ALLOWED_USERS` controls inbound DROP, `get_email` warnings, AND `get_attachment` gating. No `TRUSTED_SENDERS` anywhere.
- **No `/me` endpoints:** All Graph calls route through `/users/{M365_MAILBOX_USER}/...`. `GraphClient.mail_url()` rejects absolute/user-routed paths.
- **Attachment path hardcoded:** `~/.hermes/inbox/email/`. No caller-controlled output paths.
- **Polling watermark:** Snapshot `starting_watermark = state.watermark` BEFORE the message loop. Never update during batch processing (fixes newest-first skip bug).
- **`get_email` returns metadata only:** No attachment bytes. Inline markers replace `cid:` references in sanitized body.

## Commands

```bash
uv run pytest -q                    # full suite (98 pass, 1 skip)
uv run pytest -q -k live            # live smoke (skipped by default)
uv run pytest tests/test_adapter.py -q   # adapter + polling only
grep -R '"/me"' . --include='*.py'  # must produce NO output
```

No separate lint/typecheck targets exist in `pyproject.toml`. Some source files contain inline `# pyright:` suppression comments.

## Env Vars

| Variable | Required at Runtime | Default / Notes |
|---|---|---|
| `M365_MAIL_CLIENT_ID` | yes | — |
| `M365_MAIL_CLIENT_SECRET` | yes | — |
| `M365_MAIL_TENANT_ID` | yes | — |
| `M365_MAILBOX_USER` | **yes** | `MailConfig.from_env()` raises if missing (plugin.yaml lists it optional, but code requires it) |
| `EMAIL_ALLOWED_USERS` | **yes** | Set to `""` for deny-all. Parsed as lowercase comma-separated set. |
| `M365_ATTACHMENT_MAX_BYTES` | no | `10485760` (10MB) |
| `M365_EMAIL_STATE_PATH` | no | `.runtime/poll-state.json` |
| `M365_POLL_INTERVAL_SECONDS` | no | `30` |
| `DISABLE_SEND_CONFIRM` | no | Set to `true` to bypass send confirmation token flow |

Env is loaded from a `.env` at the **project root** via `load_dotenv(override=False)` inside `MailConfig.from_env()`.

## Testing

- `pytest-asyncio` mode is `auto`.
- All fast tests use `respx`/`httpx` mocking — no live credentials needed.
- No `conftest.py`; each test file is self-contained.
- Live smoke (`test_live_smoke.py`) skips unless `M365_EMAIL_LIVE_TESTS=true` AND the three `M365_MAIL_*` env vars are set.

## Installation into Hermes

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

Then restart: `hermes gateway restart`
