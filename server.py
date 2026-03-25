"""
Exchange ActiveSync MCP Server
Provides tools for reading email, calendar, and contacts via EAS protocol.
"""

import json
import os
import logging
from contextlib import asynccontextmanager
from typing import Optional, List

from mcp.server.fastmcp import FastMCP, Context
from pydantic import BaseModel, Field, ConfigDict

from eas_client import EASClient, FOLDER_TYPES

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ============================================================
# Configuration from environment
# ============================================================
EAS_HOST = os.environ.get("EAS_HOST", "mail.inplatlabs.ru")
EAS_USERNAME = os.environ.get("EAS_USERNAME", "")
EAS_PASSWORD = os.environ.get("EAS_PASSWORD", "")
EAS_DEVICE_ID = os.environ.get("EAS_DEVICE_ID", "EAS0LEGCLIENT0001")
EAS_PROTOCOL = os.environ.get("EAS_PROTOCOL", "14.1")
EAS_EMAIL = os.environ.get("EAS_EMAIL", "")
STATE_FILE = os.environ.get("STATE_FILE", "/app/eas_state.json")
API_KEY = os.environ.get("API_KEY", "")
YANDEX_MAPS_KEY = os.environ.get("YANDEX_MAPS_KEY", "")


# ============================================================
# Lifespan: initialize EAS client
# ============================================================
@asynccontextmanager
async def app_lifespan(server):
    """Initialize EAS client and sync folder list on startup."""
    if not EAS_USERNAME or not EAS_PASSWORD:
        raise ValueError(
            "EAS_USERNAME and EAS_PASSWORD must be set. "
            "Use env vars or .env file."
        )

    client = EASClient(
        host=EAS_HOST,
        username=EAS_USERNAME,
        password=EAS_PASSWORD,
        device_id=EAS_DEVICE_ID,
        protocol_version=EAS_PROTOCOL,
        email_address=EAS_EMAIL,
        state_file=STATE_FILE,
    )

    logger.info("Connecting to Exchange: %s as %s", EAS_HOST, EAS_USERNAME)
    folders = client.folder_sync()
    logger.info("Found %d folders", len(folders))

    yield {"eas": client}

    client.close()


# ============================================================
# MCP Server
# ============================================================
_global_eas = None
mcp = FastMCP("exchange_eas_mcp", lifespan=app_lifespan)


def get_client(ctx=None) -> EASClient:
    if ctx:
        try:
            return ctx.request_context.lifespan_state["eas"]
        except Exception:
            pass
    # Fallback: create/reuse global client
    global _global_eas
    if "_global_eas" not in globals() or _global_eas is None:
        _global_eas = EASClient(
            host=EAS_HOST,
            username=EAS_USERNAME,
            password=EAS_PASSWORD,
            device_id=EAS_DEVICE_ID,
            protocol_version=EAS_PROTOCOL,
            email_address=EAS_EMAIL,
            state_file=STATE_FILE,
        )
        _global_eas.folder_sync()
    return _global_eas


# ============================================================
# Input Models
# ============================================================
class ListFoldersInput(BaseModel):
    """Input for listing Exchange folders."""
    model_config = ConfigDict(extra='forbid')
    folder_type: Optional[int] = Field(
        default=None,
        description="Filter by type: 2=Inbox, 3=Drafts, 4=Deleted, "
                    "5=Sent, 7=Tasks, 8=Calendar, 9=Contacts, 12=User Mail. "
                    "Omit to list all folders."
    )


class GetEmailsInput(BaseModel):
    """Input for fetching emails."""
    model_config = ConfigDict(extra='forbid')
    folder_id: Optional[str] = Field(
        default=None,
        description="Folder ServerId. Omit to use Inbox."
    )
    max_items: int = Field(
        default=25, ge=1, le=100,
        description="Maximum emails to return"
    )
    include_body: bool = Field(
        default=False,
        description="Include email body text (increases response size)"
    )


class GetCalendarInput(BaseModel):
    """Input for fetching calendar events."""
    model_config = ConfigDict(extra='forbid')
    folder_id: Optional[str] = Field(
        default=None,
        description="Calendar folder ServerId. Omit to use default Calendar."
    )
    max_items: int = Field(
        default=50, ge=1, le=200,
        description="Maximum events to return"
    )


class GetContactsInput(BaseModel):
    """Input for fetching contacts."""
    model_config = ConfigDict(extra='forbid')
    folder_id: Optional[str] = Field(
        default=None,
        description="Contacts folder ServerId. Omit to use default Contacts."
    )
    max_items: int = Field(
        default=100, ge=1, le=500,
        description="Maximum contacts to return"
    )


class SearchEmailInput(BaseModel):
    """Input for searching emails."""
    model_config = ConfigDict(extra='forbid')
    query: str = Field(
        ..., min_length=1, max_length=200,
        description="Search query (matches subject, from, body)"
    )
    folder_id: Optional[str] = Field(
        default=None,
        description="Folder to search. Omit for Inbox."
    )
    max_results: int = Field(
        default=20, ge=1, le=50,
        description="Maximum results"
    )


# ============================================================
# Tools
# ============================================================

@mcp.tool(
    name="exchange_list_folders",
    annotations={
        "title": "List Exchange Folders",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def exchange_list_folders(folder_type: int = None, ctx: Context = None) -> str:
    """List all Exchange mailbox folders with their IDs and types.

    Returns folder name, ServerId (needed for other tools), type, and parent.
    Use folder_type to filter (e.g., 2 for Inbox, 8 for Calendar).
    """
    client = get_client(ctx)

    if not client.folders:
        client.folder_sync()

    result = []
    for fid, f in sorted(client.folders.items(), key=lambda x: x[1].get("name", "")):
        ft = f.get("type", 0)
        if folder_type is not None and ft != folder_type:
            continue
        result.append({
            "id": fid,
            "name": f.get("name", ""),
            "type": ft,
            "type_name": FOLDER_TYPES.get(ft, f"Type {ft}"),
            "parent_id": f.get("parent", "0"),
        })

    return json.dumps({"folders": result, "count": len(result)}, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_get_emails",
    annotations={
        "title": "Get Emails from Exchange",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_get_emails(folder_id: str = None, max_items: int = 25, include_body: bool = False, ctx: Context = None) -> str:
    """Fetch emails from an Exchange mailbox folder.

    By default reads from Inbox. Returns subject, from, to, date, read status.
    Set include_body=true to get email body text (plain text).

    Args:
        params: GetEmailsInput with folder_id, max_items, include_body

    Returns:
        JSON with list of emails containing subject, from, to, date, read status
    """
    client = get_client(ctx)

    folder_id = folder_id
    if not folder_id:
        folder_id = client.find_folder(2)  # Inbox
        if not folder_id:
            return json.dumps({"error": "Inbox not found. Run exchange_list_folders first."})

    body_size = "51200" if include_body else "0"
    result = client.sync_folder(
        folder_id,
        window_size=max_items,
        body_type="1",
        body_size=body_size,
    )

    if not result.get("elements"):
        return json.dumps({
            "emails": [],
            "count": 0,
            "folder_id": folder_id,
            "status": result.get("status", "unknown"),
        }, ensure_ascii=False)

    emails = client.parse_emails(result["elements"])

    # Trim to max
    emails = emails[:max_items]

    # Clean up for output
    output = []
    for e in emails:
        item = {
            "subject": e.get("subject", "(no subject)"),
            "from": e.get("from", ""),
            "to": e.get("to", ""),
            "date": e.get("date", ""),
            "read": e.get("read", "0") == "1",
        }
        if e.get("cc"):
            item["cc"] = e["cc"]
        if e.get("importance"):
            item["importance"] = {"0": "low", "1": "normal", "2": "high"}.get(
                e["importance"], e["importance"])
        if include_body and e.get("body"):
            item["body"] = e["body"]
        if e.get("preview"):
            item["preview"] = e["preview"]
        if e.get("thread_topic"):
            item["thread_topic"] = e["thread_topic"]
        output.append(item)

    return json.dumps({
        "emails": output,
        "count": len(output),
        "folder_id": folder_id,
        "sync_key": result.get("sync_key"),
    }, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_get_calendar",
    annotations={
        "title": "Get Calendar Events from Exchange",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_get_calendar(
    folder_id: str = None,
    max_items: int = 500,
    date_from: str = "",
    date_to: str = "",
    ctx: Context = None,
) -> str:
    """Fetch calendar events from Exchange with optional date filtering.

    Args:
        folder_id: Calendar folder ServerId (default: primary Calendar)
        max_items: Maximum events to scan (default 500)
        date_from: Start date filter YYYY-MM-DD (inclusive). Omit for no lower bound.
        date_to: End date filter YYYY-MM-DD (inclusive). Omit for no upper bound.

    Returns:
        JSON with list of calendar events, filtered and sorted by start time.

    Examples:
        Today's events: date_from="2026-03-25", date_to="2026-03-25"
        This week: date_from="2026-03-24", date_to="2026-03-30"
        All events: omit both dates
    """
    client = get_client(ctx)

    fid = folder_id or client.find_folder(8)
    if not fid:
        return json.dumps({"error": "Calendar folder not found."})

    result = client.sync_folder(fid, window_size=max_items, body_type="1", body_size="4096")

    if not result.get("elements"):
        return json.dumps({"events": [], "count": 0, "date_from": date_from, "date_to": date_to}, ensure_ascii=False)

    events = client.parse_calendar(result["elements"])

    # Filter by date range, expanding recurring events
    if date_from or date_to:
        df = date_from.replace("-", "") if date_from else "00000000"
        dt = date_to.replace("-", "") if date_to else "99999999"
        events = client.expand_recurring(events, df, dt)
    else:
        events.sort(key=lambda e: e.get("start", ""))

    return json.dumps({
        "events": events,
        "count": len(events),
        "date_from": date_from,
        "date_to": date_to,
    }, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_get_contacts",
    annotations={
        "title": "Get Contacts from Exchange",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_get_contacts(folder_id: str = None, max_items: int = 100, ctx: Context = None) -> str:
    """Fetch contacts from Exchange address book.

    By default reads from the primary Contacts folder.
    Returns name, email, phone, company, job title.

    Args:
        params: GetContactsInput with folder_id, max_items

    Returns:
        JSON with list of contacts
    """
    client = get_client(ctx)

    folder_id = folder_id
    if not folder_id:
        folder_id = client.find_folder(9)  # Contacts
        if not folder_id:
            return json.dumps({"error": "Contacts folder not found."})

    result = client.sync_folder(
        folder_id,
        window_size=max_items,
        body_type="1",
        body_size="256",
    )

    if not result.get("elements"):
        return json.dumps({
            "contacts": [],
            "count": 0,
            "folder_id": folder_id,
            "status": result.get("status", "unknown"),
        }, ensure_ascii=False)

    contacts = client.parse_contacts(result["elements"])
    contacts = contacts[:max_items]

    return json.dumps({
        "contacts": contacts,
        "count": len(contacts),
        "folder_id": folder_id,
        "sync_key": result.get("sync_key"),
    }, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_search_emails",
    annotations={
        "title": "Search Emails in Exchange",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def exchange_search_emails(query: str = "", folder_id: str = None, max_results: int = 20, ctx: Context = None) -> str:
    """Search emails by subject, sender, or content.

    Fetches emails from the specified folder and filters locally by query.
    Query matches against subject, from, to, and preview fields (case-insensitive).

    Args:
        params: SearchEmailInput with query, folder_id, max_results

    Returns:
        JSON with matching emails
    """
    client = get_client(ctx)

    folder_id = folder_id
    if not folder_id:
        folder_id = client.find_folder(2)
        if not folder_id:
            return json.dumps({"error": "Inbox not found."})

    result = client.sync_folder(
        folder_id,
        window_size=200,
        body_type="1",
        body_size="1024",
    )

    if not result.get("elements"):
        return json.dumps({"results": [], "count": 0, "query": query})

    emails = client.parse_emails(result["elements"])

    q = query.lower()
    matches = []
    for e in emails:
        searchable = " ".join([
            e.get("subject", ""), e.get("from", ""),
            e.get("to", ""), e.get("preview", ""),
            e.get("body", ""),
        ]).lower()
        if q in searchable:
            matches.append({
                "subject": e.get("subject", ""),
                "from": e.get("from", ""),
                "to": e.get("to", ""),
                "date": e.get("date", ""),
                "read": e.get("read", "0") == "1",
                "preview": e.get("preview", "")[:200],
            })
            if len(matches) >= max_results:
                break

    return json.dumps({
        "results": matches,
        "count": len(matches),
        "query": query,
        "folder_id": folder_id,
    }, ensure_ascii=False, indent=2)


# ============================================================
# REST API with FastAPI (Swagger UI at /docs)
# ============================================================
from fastapi import FastAPI, Header, HTTPException, Query
from fastapi.responses import JSONResponse
from pydantic import BaseModel as PydanticBase, Field as PydField
from typing import Optional, List

api = FastAPI(
    title="Exchange ActiveSync API",
    description="REST API for accessing Microsoft Exchange via ActiveSync protocol. "
                "Provides email, calendar, contacts access and search.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

_rest_eas = None

def _rest_client():
    global _rest_eas
    if _rest_eas is None:
        _rest_eas = EASClient(
            host=EAS_HOST,
            username=EAS_USERNAME,
            password=EAS_PASSWORD,
            device_id=EAS_DEVICE_ID,
            protocol_version=EAS_PROTOCOL,
            email_address=EAS_EMAIL,
        )
        _rest_eas.folder_sync()
    return _rest_eas


def _verify_key(x_api_key: str = Header(default=None), authorization: str = Header(default=None)):
    if not API_KEY:
        return
    token = None
    if x_api_key:
        token = x_api_key
    elif authorization and authorization.startswith("Bearer "):
        token = authorization[7:]
    if token != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key. Use X-API-Key header.")


# --- Models for Swagger ---
class SendEmailRequest(PydanticBase):
    to: str = PydField(..., description="Recipient email. Comma-separated for multiple", example="alice@company.com")
    subject: str = PydField(..., description="Email subject", example="Meeting tomorrow")
    body: str = PydField("", description="Email body text")
    cc: str = PydField("", description="CC recipients, comma-separated")
    content_type: str = PydField("plain", description="'plain' or 'html'")

class CreateEventRequest(PydanticBase):
    subject: str = PydField(..., description="Event title", example="Team standup")
    start_time: str = PydField(..., description="Start time ISO format UTC", example="2026-03-25T10:00:00Z")
    end_time: str = PydField(..., description="End time ISO format UTC", example="2026-03-25T11:00:00Z")
    location: str = PydField("", description="Event location")
    body: str = PydField("", description="Event description")
    attendees: str = PydField("", description="Comma-separated emails", example="alice@co.com,bob@co.com")
    all_day: bool = PydField(False, description="All-day event")
    reminder: int = PydField(15, description="Reminder in minutes before event")


# --- Endpoints ---
@api.get("/api/health", tags=["System"], summary="Health check")
async def api_health():
    """Check server status. No authentication required."""
    return {"status": "ok", "host": EAS_HOST, "device_id": EAS_DEVICE_ID}


@api.get("/api/folders", tags=["Mailbox"], summary="List folders")
async def api_folders(
    type: Optional[int] = Query(None, description="Filter by folder type: 2=Inbox, 3=Drafts, 4=Deleted, 5=Sent, 7=Tasks, 8=Calendar, 9=Contacts"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """List all mailbox folders with IDs and types."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    if not c.folders:
        c.folder_sync()
    result = []
    for fid, f in sorted(c.folders.items(), key=lambda x: x[1].get("name", "")):
        t = f.get("type", 0)
        if type is not None and t != type:
            continue
        result.append({"id": fid, "name": f.get("name", ""), "type": t, "type_name": FOLDER_TYPES.get(t, f"Type {t}")})
    return {"folders": result, "count": len(result)}


@api.get("/api/emails", tags=["Email"], summary="Get emails")
async def api_emails(
    folder_id: Optional[str] = Query(None, description="Folder ServerId (default: Inbox)"),
    max: int = Query(25, ge=1, le=100, description="Maximum emails to return"),
    body: bool = Query(False, description="Include email body text"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Fetch emails from a mailbox folder. Default is Inbox."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(2)
    body_size = "51200" if body else "0"
    r = c.sync_folder(fid, window_size=max, body_type="1", body_size=body_size)
    emails = c.parse_emails(r.get("elements", []))[:max]
    output = []
    for e in emails:
        item = {"subject": e.get("subject", ""), "from": e.get("from", ""), "to": e.get("to", ""), "date": e.get("date", ""), "read": e.get("read", "0") == "1"}
        if e.get("cc"): item["cc"] = e["cc"]
        if body and e.get("body"): item["body"] = e["body"]
        if e.get("preview"): item["preview"] = e["preview"]
        output.append(item)
    return {"emails": output, "count": len(output)}


@api.get("/api/search", tags=["Email"], summary="Search emails")
async def api_search(
    q: str = Query(..., min_length=1, description="Search query (matches subject, from, to, body)"),
    folder_id: Optional[str] = Query(None, description="Folder to search (default: Inbox)"),
    max: int = Query(20, ge=1, le=50, description="Maximum results"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Search emails by subject, sender, or content. Case-insensitive."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(2)
    r = c.sync_folder(fid, window_size=200, body_type="1", body_size="1024")
    emails = c.parse_emails(r.get("elements", []))
    ql = q.lower()
    matches = []
    for e in emails:
        text = " ".join([e.get("subject",""), e.get("from",""), e.get("to",""), e.get("preview",""), e.get("body","")]).lower()
        if ql in text:
            matches.append({"subject": e.get("subject",""), "from": e.get("from",""), "date": e.get("date",""), "read": e.get("read","0") == "1"})
            if len(matches) >= max:
                break
    return {"results": matches, "count": len(matches), "query": q}


@api.get("/api/calendar", tags=["Calendar"], summary="Get calendar events")
async def api_calendar(
    folder_id: Optional[str] = Query(None, description="Calendar folder ServerId"),
    max: int = Query(500, ge=1, le=1000, description="Maximum events to scan"),
    date_from: Optional[str] = Query(None, description="Start date YYYY-MM-DD (inclusive)"),
    date_to: Optional[str] = Query(None, description="End date YYYY-MM-DD (inclusive)"),
    date: Optional[str] = Query(None, description="Shortcut: single day YYYY-MM-DD (sets both from and to)"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Fetch calendar events with date filtering.
    
    Use date for a single day, or date_from/date_to for a range.
    Examples: ?date=2026-03-25 or ?date_from=2026-03-24&date_to=2026-03-30
    """
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(8)
    r = c.sync_folder(fid, window_size=max, body_type="1", body_size="4096")
    events = c.parse_calendar(r.get("elements", []))

    # Handle date shortcut
    df = date_from
    dt = date_to
    if date:
        df = date
        dt = date

    if df or dt:
        df_s = (df or "").replace("-", "") or "00000000"
        dt_s = (dt or "").replace("-", "") or "99999999"
        events = c.expand_recurring(events, df_s, dt_s)

    events.sort(key=lambda e: e.get("start", ""))
    return {"events": events, "count": len(events), "date_from": df, "date_to": dt}



@api.get("/api/contacts", tags=["Contacts"], summary="Get contacts")
async def api_contacts(
    folder_id: Optional[str] = Query(None, description="Contacts folder ServerId"),
    max: int = Query(100, ge=1, le=500, description="Maximum contacts"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Fetch contacts from address book. Default is primary Contacts folder."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(9)
    r = c.sync_folder(fid, window_size=max, body_type="1", body_size="256")
    contacts = c.parse_contacts(r.get("elements", []))[:max]
    return {"contacts": contacts, "count": len(contacts)}


@api.post("/api/send", tags=["Email"], summary="Send email")
async def api_send_email(
    data: SendEmailRequest,
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Send an email via Exchange. Supports plain text and HTML."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    result = c.send_email(to=data.to, subject=data.subject, body=data.body, cc=data.cc, content_type=data.content_type)
    return result


@api.post("/api/event", tags=["Calendar"], summary="Create calendar event")
async def api_create_event(
    data: CreateEventRequest,
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Create a calendar event. Times must be in UTC (Z suffix). Moscow = UTC+3."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    attendees = None
    if data.attendees:
        attendees = [{"email": e.strip(), "name": e.strip()} for e in data.attendees.split(",") if e.strip()]
    result = c.create_event(
        subject=data.subject, start_time=data.start_time, end_time=data.end_time,
        location=data.location, body=data.body, attendees=attendees,
        all_day=data.all_day, reminder=data.reminder,
    )
    return result


@api.get("/api/new_emails", tags=["Email"], summary="Get new emails (incremental)")
async def api_new_emails(
    folder_id: Optional[str] = Query(None, description="Folder ServerId (default: Inbox)"),
    max: int = Query(50, ge=1, le=200, description="Maximum emails"),
    body: bool = Query(True, description="Include email body"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Get only NEW emails since last check. First call returns all current emails."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(2)
    body_size = "51200" if body else "0"
    result = c.sync_incremental(fid, window_size=max, body_type="1", body_size=body_size)
    emails = c.parse_emails(result.get("elements", []))
    output = []
    for e in emails:
        item = {"subject": e.get("subject",""), "from": e.get("from",""), "to": e.get("to",""), "date": e.get("date",""), "read": e.get("read","0") == "1"}
        if e.get("cc"): item["cc"] = e["cc"]
        if body and e.get("body"): item["body"] = e["body"]
        if e.get("preview"): item["preview"] = e["preview"]
        output.append(item)
    return {"emails": output, "count": len(output), "is_initial": result.get("is_initial", False)}


@api.get("/api/new_events", tags=["Calendar"], summary="Get new/changed events (incremental)")
async def api_new_events(
    folder_id: Optional[str] = Query(None, description="Calendar folder ServerId"),
    max: int = Query(50, ge=1, le=200, description="Maximum events"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Get only NEW or CHANGED calendar events since last check."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(8)
    result = c.sync_incremental(fid, window_size=max, body_type="1", body_size="4096")
    events = c.parse_calendar(result.get("elements", []))
    return {"events": events, "count": len(events), "is_initial": result.get("is_initial", False)}


@api.get("/api/attachment", tags=["Email"], summary="Download attachment")
async def api_get_attachment(
    file_reference: str = Query(..., description="FileReference from email attachment metadata"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Download an email attachment by FileReference. Returns base64-encoded data."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    result = c.get_attachment(file_reference)
    return result




# ============================================================
# Entry point
# ============================================================


@mcp.tool(
    name="exchange_get_attachment",
    annotations={
        "title": "Download Email Attachment",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def exchange_get_attachment(
    file_reference: str = "",
    ctx: Context = None,
) -> str:
    """Download an email attachment by its FileReference.

    First use exchange_get_emails with include_body=true to get attachment
    file_references, then use this tool to download.

    Args:
        file_reference: The FileReference string from email attachment metadata

    Returns:
        JSON with base64-encoded attachment data and content type
    """
    if not file_reference:
        return json.dumps({"error": "file_reference is required"})

    client = get_client(ctx)
    result = client.get_attachment(file_reference)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_get_new_emails",
    annotations={
        "title": "Get New Emails (Incremental)",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_get_new_emails(
    folder_id: str = None,
    max_items: int = 50,
    include_body: bool = True,
    ctx: Context = None,
) -> str:
    """Get only NEW emails since last check. Uses stored SyncKey.

    First call returns all current emails and stores the sync position.
    Subsequent calls return only new emails that arrived since last check.

    Args:
        folder_id: Folder ServerId (default: Inbox)
        max_items: Maximum emails to return
        include_body: Include email body text (default: true for new emails)

    Returns:
        JSON with new emails, count, and is_initial flag
    """
    client = get_client(ctx)
    fid = folder_id or client.find_folder(2)
    if not fid:
        return json.dumps({"error": "Inbox not found"})

    body_size = "51200" if include_body else "0"
    result = client.sync_incremental(
        fid, window_size=max_items,
        body_type="1", body_size=body_size,
    )

    emails = client.parse_emails(result.get("elements", []))

    output = []
    for e in emails:
        item = {
            "subject": e.get("subject", ""),
            "from": e.get("from", ""),
            "to": e.get("to", ""),
            "date": e.get("date", ""),
            "read": e.get("read", "0") == "1",
        }
        if e.get("cc"): item["cc"] = e["cc"]
        if include_body and e.get("body"): item["body"] = e["body"]
        if e.get("preview"): item["preview"] = e["preview"]
        output.append(item)

    return json.dumps({
        "emails": output,
        "count": len(output),
        "is_initial": result.get("is_initial", False),
        "status": result.get("status"),
        "folder_id": fid,
    }, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_get_new_events",
    annotations={
        "title": "Get New/Changed Calendar Events (Incremental)",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_get_new_events(
    folder_id: str = None,
    max_items: int = 50,
    ctx: Context = None,
) -> str:
    """Get only NEW or CHANGED calendar events since last check.

    First call returns all current events and stores sync position.
    Subsequent calls return only new/changed/deleted events.

    Args:
        folder_id: Calendar folder ServerId (default: Calendar)
        max_items: Maximum events to return

    Returns:
        JSON with new events, count, and is_initial flag
    """
    client = get_client(ctx)
    fid = folder_id or client.find_folder(8)
    if not fid:
        return json.dumps({"error": "Calendar not found"})

    result = client.sync_incremental(
        fid, window_size=max_items,
        body_type="1", body_size="4096",
    )

    events = client.parse_calendar(result.get("elements", []))

    return json.dumps({
        "events": events,
        "count": len(events),
        "is_initial": result.get("is_initial", False),
        "status": result.get("status"),
        "folder_id": fid,
    }, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_send_email",
    annotations={
        "title": "Send Email via Exchange",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def exchange_send_email(
    to: str = "",
    subject: str = "",
    body: str = "",
    cc: str = "",
    content_type: str = "plain",
    ctx: Context = None,
) -> str:
    """Send an email through Exchange.

    Args:
        to: Recipient email address (comma-separated for multiple)
        subject: Email subject line
        body: Email body text
        cc: CC recipients (optional, comma-separated)
        content_type: 'plain' for text or 'html' for HTML body

    Returns:
        JSON with send status
    """
    if not to or not subject:
        return json.dumps({"error": "Both 'to' and 'subject' are required"})

    client = get_client(ctx)
    result = client.send_email(
        to=to, subject=subject, body=body,
        cc=cc, content_type=content_type,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool(
    name="exchange_create_event",
    annotations={
        "title": "Create Calendar Event in Exchange",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    }
)
async def exchange_create_event(
    subject: str = "",
    start_time: str = "",
    end_time: str = "",
    location: str = "",
    body: str = "",
    attendees: str = "",
    all_day: bool = False,
    reminder: int = 15,
    ctx: Context = None,
) -> str:
    """Create a calendar event in Exchange.

    Args:
        subject: Event title
        start_time: Start in ISO format, e.g. '2026-03-25T10:00:00.000Z'
        end_time: End in ISO format, e.g. '2026-03-25T11:00:00.000Z'
        location: Event location (optional)
        body: Event description (optional)
        attendees: Comma-separated emails, e.g. 'alice@co.com,bob@co.com' (optional)
        all_day: Whether this is an all-day event (default false)
        reminder: Reminder in minutes before event (default 15)

    Returns:
        JSON with creation status and server ID
    """
    if not subject or not start_time or not end_time:
        return json.dumps({"error": "subject, start_time, and end_time are required"})

    att_list = None
    if attendees:
        att_list = [{"email": e.strip(), "name": e.strip()} for e in attendees.split(",") if e.strip()]

    client = get_client(ctx)
    result = client.create_event(
        subject=subject,
        start_time=start_time,
        end_time=end_time,
        location=location,
        body=body,
        attendees=att_list,
        all_day=all_day,
        reminder=reminder,
    )
    return json.dumps(result, ensure_ascii=False, indent=2)


class RewriteHostMiddleware:
    def __init__(self, app):
        self.app = app
    async def __call__(self, scope, receive, send):
        if scope["type"] in ("http", "websocket"):
            scope["headers"] = [
                (b"host", b"localhost:8000") if k == b"host" else (k, v)
                for k, v in scope.get("headers", [])
            ]
            scope["server"] = ("localhost", 8000)
        await self.app(scope, receive, send)


if __name__ == "__main__":
    import sys
    if "--http" in sys.argv:
        import uvicorn
        port = 8000
        for arg in sys.argv[1:]:
            if arg.startswith("--port="):
                port = int(arg.split("=")[1])
        logger.info("Starting MCP server on HTTP 0.0.0.0:%d", port)
        mcp_asgi = mcp.streamable_http_app()
        mcp_wrapped = RewriteHostMiddleware(mcp_asgi)

        class CombinedApp:
            def __init__(self, mcp_app, rest_app):
                self.mcp = mcp_app
                self.rest = rest_app
            async def __call__(self, scope, receive, send):
                if scope["type"] == "lifespan":
                    # Forward lifespan to MCP app (it needs startup/shutdown)
                    await self.mcp(scope, receive, send)
                    return
                path = scope.get("path", "")
                if path.startswith("/mcp"):
                    await self.mcp(scope, receive, send)
                else:
                    await self.rest(scope, receive, send)

        app = CombinedApp(mcp_wrapped, api)
        uvicorn.run(app, host="0.0.0.0", port=port)
    else:
        logger.info("Starting MCP server on stdio")
        mcp.run()
