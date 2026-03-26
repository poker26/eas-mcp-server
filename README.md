# Exchange ActiveSync MCP Server

MCP server that provides access to Exchange email, calendar, and contacts
via the ActiveSync protocol. Designed for on-premise Exchange servers
that expose EAS with Basic authentication.

## Architecture

```
Claude / n8n  -->  MCP Server (HTTP)  -->  Exchange Server (EAS)
                   localhost:8900          mail.inplatlabs.ru:443
```

## Tools

| Tool | Description |
|------|-------------|
| `exchange_list_folders` | List all mailbox folders with IDs and types |
| `exchange_get_emails` | Fetch emails from Inbox or any folder |
| `exchange_get_calendar` | Fetch calendar events |
| `exchange_get_contacts` | Fetch contacts from address book |
| `exchange_search_emails` | Search emails by subject, sender, content |

## Calendar Sync Contract

- `exchange_get_calendar` provides a date-range snapshot for calendar UI/read flows.
- `exchange_get_new_events` is incremental and should be used for polling/event-driven updates.
- Warm-up requirement: call `exchange_get_new_events` until `is_initial=false` before relying on incremental results.
- Write operations (`exchange_create_event` / REST `POST /api/event`) preserve incremental SyncKey state and must not reset the baseline.

Recommended app flow:

1. Fetch snapshot: `exchange_get_calendar` (or REST `/api/calendar`) for the visible range.
2. Warm-up incremental stream: `exchange_get_new_events` until `is_initial=false`.
3. Poll `exchange_get_new_events` periodically and merge deltas into local state.

## Quick Start

### Local (without Docker)

```bash
pip install -r requirements.txt

export EAS_USERNAME="OFFICE\oleg.pokrovskiy"
export EAS_PASSWORD="your_password"

# stdio mode (for local MCP clients)
python server.py

# HTTP mode (for remote access)
python server.py --http --port=8000
```

### Docker

```bash
cp .env.example .env
# Edit .env with your credentials

docker compose up -d
```

Server will be available at `http://localhost:8900/mcp`.

### With Traefik

Uncomment the labels section in `docker-compose.yml` and adjust
the hostname to your domain.

## MCP Client Configuration

### Claude Desktop (claude_desktop_config.json)

```json
{
  "mcpServers": {
    "exchange": {
      "url": "http://localhost:8900/mcp"
    }
  }
}
```

### n8n

Use the MCP Client node with HTTP transport pointing to
`http://eas-mcp-server:8000/mcp` (if in same Docker network).

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `EAS_HOST` | Exchange server hostname | `mail.inplatlabs.ru` |
| `EAS_USERNAME` | Login (DOMAIN\user or email) | required |
| `EAS_PASSWORD` | Password | required |
| `EAS_DEVICE_ID` | Device ID for EAS whitelist | `EAS0LEGCLIENT0001` |
| `EAS_PROTOCOL` | EAS protocol version | `14.1` |

## Security Notes

- Credentials are stored in `.env` file - keep it secure
- The server disables TLS verification for Exchange (common with on-prem)
- Bind to `127.0.0.1` in production, use Traefik/nginx for external access
- Consider adding authentication to the MCP HTTP endpoint
