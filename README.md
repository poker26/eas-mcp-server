# Exchange MCP

MCP server that exposes an on-premise Exchange mailbox (EWS) to MCP
clients such as n8n or Claude Code. Folders, mail, calendar, contacts
and send-mail via the usual `exchange_*` tool family.

Per-folder timestamp cursor + Message-ID LRU deduplicate
`exchange_get_new_emails`, so repeated polls don't replay the same
messages.

## Layout

```
Dockerfile
docker-compose.yml
requirements.txt
.env.example
exchange_mcp/
  config.py           # pydantic-settings
  auth.py             # X-API-Key middleware
  state.py            # AtomicJSONState + per-folder cursor + Message-ID LRU
  backends/
    base.py           # MailBackend protocol + DTO
    ews.py            # exchangelib-based driver
  router.py           # Service wrapper over EWS
  health.py           # GET /health
  mcp_server.py       # FastMCP registration
  main.py             # FastAPI + FastMCP mount
  tools/
    __init__.py       # ALL_TOOLS
    folders.py
    mail.py
    calendar.py
    contacts.py
    attachments.py
```

## Quick start

```bash
cp .env.example .env
# fill in EXCHANGE_USER, EXCHANGE_PASSWORD, MCP_API_KEY
docker compose up -d
curl http://127.0.0.1:8903/health
```

Endpoints:

- `GET /health` — unauthenticated liveness probe (HEAD to `/EWS/Exchange.asmx`)
- `POST /mcp` — MCP transport (requires `X-API-Key` or `Authorization: Bearer ...`)

## Status

v0.1 skeleton. Working:

- `/health` with EWS reachability check (anonymous HEAD — no auth spam)
- `exchange_list_folders`
- `exchange_get_new_emails` (per-folder cursor + Message-ID dedup)
- `exchange_get_emails` (non-incremental, last month)
- `exchange_send_email`

Stubs / TODO:

- `exchange_get_calendar` / `exchange_get_new_events`
- `exchange_get_contacts`
- `exchange_search_emails`
- `exchange_get_attachment`
- Unit tests

## Design notes

- State is tracked by **timestamp + Message-ID**, not by native
  SyncKey / SyncState. This survives account moves and works with any
  backend that returns `datetime_received` and `InternetMessageId`.
- `healthcheck()` is a plain anonymous `HEAD` to the EWS endpoint — a
  401 with `WWW-Authenticate` is the expected "alive" answer. Real auth
  is validated on the first tool call, not on every probe.
- `router.py` is a single-backend wrapper today. The `MailBackend`
  Protocol is kept as a hook so a second backend (IMAP, Graph, etc.)
  can be added without touching `tools/`.
- The EAS channel (ActiveSync/WBXML) originally co-habited this server
  as a failover; it now lives in its own repo (`eas-mcp-server`) and
  the two run as independent MCPs.
