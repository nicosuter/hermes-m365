# AGENTS.md — M365 Email Hermes Plugin

## What This Is

A **Hermes gateway platform plugin** (NOT an MCP server, NOT FastMCP). It treats an M365 mailbox as a messaging platform and registers explicit mail tools.

## Project Layout — Non-Standard

`adapter.py` and `plugin.yaml` **must stay at the project root** for Hermes plugin discovery. All other modules are sibling Python files (flat layout, no package).

```
m365-email-hermes-plugin/
├── plugin.yaml              # Hermes manifest — MUST be at root
├── adapter.py               # register(ctx) + M365EmailAdapter — MUST be at root
├── config.py                # MailConfig, EMAIL_ALLOWED_USERS parsing
├── graph.py                 # GraphClient (token, pagination, /users/{mailbox} routing)
├── mail_tools.py            # list_mail, get_email, get_attachment, send_email, reply, forward, mark read/unread
├── sanitize.py              # HTML→text, hidden content removal
├── attachments.py           # filename safety, sender gating, deterministic paths
├── state.py                 # PollState (watermark, processed IDs, _MAX_PROCESSED_IDS=500)
├── __init__.py              # module marker for Hermes plugin discovery
└── tests/                   # fast tests via respx/httpx, plus 1 live smoke
```

## Tool Registry

`register(ctx)` wires tools into Hermes:

- `list_mail(top=50, filter=None, unreadOnly=False)`
- `get_email(email_id)` — returns sanitized text + attachment metadata only (no bytes)
- `get_attachment(email_id, attachment_id)` — downloads to `~/.hermes/inbox/email/`
- `send_email(to, subject, body, reply_to=None)` — requires `confirm_send_email(token)` unless `DISABLE_SEND_CONFIRM=true`
- `reply_email(email_id, body)` — requires `confirm_reply_email(token)` unless `DISABLE_SEND_CONFIRM=true`
- `reply_all(email_id, body)` — requires `confirm_reply_all(token)` unless `DISABLE_SEND_CONFIRM=true`
- `forward_email(email_id, to, body)` — requires `confirm_forward_email(token)` unless `DISABLE_SEND_CONFIRM=true`
- `mark_read(email_id)` / `mark_unread(email_id)`
- `confirm_send_email(confirmation_token)` — completes a pending `send_email`
- `confirm_reply_email(confirmation_token)` — completes a pending `reply_email`
- `confirm_reply_all(confirmation_token)` — completes a pending `reply_all`
- `confirm_forward_email(confirmation_token)` — completes a pending `forward_email`

## Critical Constraints

- **Single allowlist:** `EMAIL_ALLOWED_USERS` controls inbound DROP, `get_email` warnings, AND `get_attachment` gating. No `TRUSTED_SENDERS` anywhere.
- **No `/me` endpoints:** All Graph calls route through `/users/{M365_MAILBOX_USER}/...`. `GraphClient.mail_url()` rejects absolute/user-routed paths.
- **Attachment path hardcoded:** `~/.hermes/inbox/email/`. No caller-controlled output paths.
- **Polling watermark:** Snapshot `starting_watermark = state.watermark` BEFORE the message loop. Never update during batch processing (fixes newest-first skip bug).
- **`get_email` returns metadata only:** No attachment bytes. Inline markers replace `cid:` references in sanitized body.

## Commands

```bash
uv run pytest -q                    # full suite
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

## Troubleshooting & Learnings

### Tool Registration

- **No `tools` shadowing:** Our module is `mail_tools.py`. Upstream Hermes has a `tools` package at `~/.hermes/hermes-agent/tools/`. We import FROM upstream (`from tools.lazy_deps import ensure`), so naming is fine.
- **Registration flow:** `__init__.py` → `adapter.register(ctx)` → `ctx.register_tool(toolset="m365_email", name=..., handler=...)` → upstream `tools.registry.ToolRegistry`.
- **`handler=` is required (BLOCKING):** Every `ctx.register_tool()` call MUST include the `handler=` keyword argument. Missing it causes `PluginContext.register_tool() missing 1 required positional argument: 'handler'` and the entire plugin fails to load — zero tools get registered. This has broken the plugin in the past.
