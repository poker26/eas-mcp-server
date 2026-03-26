# EAS MCP Server — n8n User Guide

Reference for all MCP tools available through the n8n MCP Client node.

## Interactive API Docs

Before configuring n8n, you can explore all available endpoints in the browser:

- **Swagger UI**: `http://<SERVER_IP>:8900/docs`
- **ReDoc**: `http://<SERVER_IP>:8900/redoc`

You can test any endpoint directly from Swagger UI by clicking "Authorize" and entering your API key.

## Connection Setup

In n8n MCP Client node:
- **Endpoint URL**: `http://<SERVER_IP>:8900/mcp`
- **Transport**: HTTP Streaming
- **Authentication**: Header Auth
  - **Header Name**: `X-API-Key`
  - **Header Value**: your API key from `.env`

---

## Tool: exchange_list_folders

List all mailbox folders with their IDs and types.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folder_type` | integer | No | all | Filter by type. Common values: `2` = Inbox, `3` = Drafts, `4` = Deleted, `5` = Sent, `7` = Tasks, `8` = Calendar, `9` = Contacts, `12` = User Mail |

**Example — list all folders:**
```json
{}
```

**Example — only calendar folders:**
```json
{
  "folder_type": 8
}
```

**Response contains:**
- `id` — folder ServerId (use this in other tools as `folder_id`)
- `name` — display name
- `type` — numeric type
- `type_name` — human-readable type

---

## Tool: exchange_get_emails

Fetch emails from a mailbox folder.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folder_id` | string | No | Inbox | Folder ServerId from `exchange_list_folders` |
| `max_items` | integer | No | 25 | Number of emails to return (1–100) |
| `include_body` | boolean | No | false | Include full email body text |

**Example — last 10 emails from Inbox:**
```json
{
  "max_items": 10
}
```

**Example — emails from Sent folder with body:**
```json
{
  "folder_id": "14",
  "max_items": 5,
  "include_body": true
}
```

**Response contains for each email:**
- `subject` — email subject
- `from` — sender
- `to` — recipients
- `cc` — CC recipients (if any)
- `date` — received date in ISO format
- `read` — true/false
- `importance` — low / normal / high
- `body` — full text (only if `include_body` = true)
- `preview` — short preview text

---

## Tool: exchange_search_emails

Search emails by subject, sender, or content.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `query` | string | **Yes** | — | Search text (case-insensitive, matches subject, from, to, body) |
| `folder_id` | string | No | Inbox | Folder to search |
| `max_results` | integer | No | 20 | Maximum results (1–50) |

**Example — find invoices:**
```json
{
  "query": "invoice"
}
```

**Example — search in Sent folder:**
```json
{
  "query": "quarterly report",
  "folder_id": "14",
  "max_results": 5
}
```

---

## Tool: exchange_get_calendar

Fetch calendar events.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folder_id` | string | No | Calendar | Calendar folder ServerId |
| `max_items` | integer | No | 50 | Number of events to return (1–200) |

**Example — get upcoming events:**
```json
{
  "max_items": 20
}
```

**Response contains for each event:**
- `subject` — event title
- `start` — start time
- `end` — end time
- `location` — location
- `organizer_name` / `organizer_email` — organizer
- `attendees` — list of participants
- `all_day` — 0 or 1
- `busy_status` — 0=Free, 1=Tentative, 2=Busy, 3=OOF
- `reminder` — minutes before event

---

## Tool: exchange_get_contacts

Fetch contacts from address book.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folder_id` | string | No | Contacts | Contacts folder ServerId |
| `max_items` | integer | No | 100 | Number of contacts to return (1–500) |

**Example:**
```json
{
  "max_items": 50
}
```

**Response contains for each contact:**
- `FileAs` — display name
- `FirstName`, `LastName`, `MiddleName`
- `Email1Address`, `Email2Address`, `Email3Address`
- `BusinessPhoneNumber`, `MobilePhoneNumber`, `HomePhoneNumber`
- `CompanyName`, `Department`, `JobTitle`

---

## Tool: exchange_send_email

Send an email.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `to` | string | **Yes** | — | Recipient email. Multiple: comma-separated |
| `subject` | string | **Yes** | — | Email subject |
| `body` | string | No | empty | Email body text |
| `cc` | string | No | empty | CC recipients, comma-separated |
| `content_type` | string | No | `plain` | `plain` for text, `html` for HTML body |

**Example — simple email:**
```json
{
  "to": "alice@company.com",
  "subject": "Meeting tomorrow",
  "body": "Hi Alice, let's meet at 10am. Best, Oleg"
}
```

**Example — HTML email with CC:**
```json
{
  "to": "alice@company.com, bob@company.com",
  "subject": "Q1 Report",
  "body": "<h1>Q1 Report</h1><p>Please find the summary below...</p>",
  "cc": "manager@company.com",
  "content_type": "html"
}
```

**Response:**
- `status: "sent"` — success
- `status: "error_..."` — failure with error code

---

## Tool: exchange_create_event

Create a calendar event.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `subject` | string | **Yes** | — | Event title |
| `start_time` | string | **Yes** | — | Start time in ISO format: `2026-03-25T10:00:00Z` |
| `end_time` | string | **Yes** | — | End time in ISO format: `2026-03-25T11:00:00Z` |
| `location` | string | No | empty | Event location |
| `body` | string | No | empty | Event description |
| `attendees` | string | No | empty | Comma-separated emails: `alice@co.com,bob@co.com` |
| `all_day` | boolean | No | false | All-day event |
| `reminder` | integer | No | 15 | Reminder in minutes before event |

**Important:** Times are in UTC. Moscow time = UTC + 3 hours.
To create event at 13:00 Moscow time, use `10:00:00Z`.

**Example — simple meeting:**
```json
{
  "subject": "Team standup",
  "start_time": "2026-03-25T07:00:00Z",
  "end_time": "2026-03-25T07:30:00Z",
  "location": "Conference Room A"
}
```

**Example — meeting with attendees:**
```json
{
  "subject": "Project review",
  "start_time": "2026-03-26T10:00:00Z",
  "end_time": "2026-03-26T11:00:00Z",
  "location": "Zoom",
  "body": "Quarterly project review meeting",
  "attendees": "alice@company.com, bob@company.com",
  "reminder": 30
}
```

**Example — all-day event:**
```json
{
  "subject": "Company holiday",
  "start_time": "2026-05-01T00:00:00Z",
  "end_time": "2026-05-02T00:00:00Z",
  "all_day": true,
  "reminder": 1440
}
```

**Response:**
- `status: "created"` — success
- `server_id` — Exchange internal ID of the created event
- `status: "error_..."` — failure

---

## Workflow Examples

### Daily email digest
```
Schedule (8:00 AM) → MCP (exchange_get_emails, max_items=30)
  → AI Agent: "Summarize these emails, highlight urgent"
  → Telegram: send digest
```

### Find contact and schedule meeting
```
Manual Trigger
  → MCP (exchange_get_contacts)
  → Code: find contact email by name
  → MCP (exchange_create_event, attendees=email)
  → MCP (exchange_send_email: notify about meeting)
```

### Search and forward
```
Manual Trigger
  → MCP (exchange_search_emails, query="invoice March")
  → Code: extract relevant emails
  → MCP (exchange_send_email: forward summary to manager)
```

---

## Tips

- **Folder IDs**: Run `exchange_list_folders` first to get ServerId values for other tools
- **Date format**: Always use UTC with `Z` suffix: `2026-03-25T10:00:00Z`
- **Moscow time**: Subtract 3 hours for UTC. 13:00 MSK = `10:00:00Z`
- **Attendees**: Use comma-separated emails, no angle brackets
- **Large mailboxes**: Use smaller `max_items` for faster responses
- **HTML emails**: Set `content_type` to `html` and use HTML tags in `body`

---

## Incremental Calendar Contract (important)

For reliable automations, treat calendar tools as two separate streams:

- **Snapshot stream**: `exchange_get_calendar` / REST `/api/calendar`
- **Delta stream**: `exchange_get_new_events` / REST `/api/new_events`

Warm-up procedure for `exchange_get_new_events`:

1. Call once (or more) until response has `is_initial: false`.
2. Only after that, use it as a true incremental feed.

Write operations (`exchange_create_event` / `/api/event`) are expected to preserve incremental state.  
After creating an event, the next incremental call should stay in steady state (`is_initial: false`) and return new items when available.
