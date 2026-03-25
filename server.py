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
GOOGLE_MAPS_KEY = os.environ.get("GOOGLE_MAPS_KEY", "")
HOME_ADDRESS = os.environ.get("HOME_ADDRESS", "")
DEFAULT_CITY = os.environ.get("DEFAULT_CITY", "Москва")
BUFFER_MINUTES = int(os.environ.get("BUFFER_MINUTES", "10"))
CALDAV_URL = os.environ.get("CALDAV_URL", "")
CALDAV_USERNAME = os.environ.get("CALDAV_USERNAME", "")
CALDAV_PASSWORD = os.environ.get("CALDAV_PASSWORD", "")


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
async def exchange_get_calendar(folder_id: str = None, max_items: int = 50, ctx: Context = None) -> str:
    """Fetch calendar events from Exchange.

    By default reads from the primary Calendar folder.
    Returns event subject, start/end times, location, organizer, attendees.

    Args:
        params: GetCalendarInput with folder_id, max_items

    Returns:
        JSON with list of calendar events
    """
    client = get_client(ctx)

    folder_id = folder_id
    if not folder_id:
        folder_id = client.find_folder(8)  # Calendar
        if not folder_id:
            return json.dumps({"error": "Calendar folder not found."})

    result = client.sync_folder(
        folder_id,
        window_size=max_items,
        body_type="1",
        body_size="1024",
    )

    if not result.get("elements"):
        return json.dumps({
            "events": [],
            "count": 0,
            "folder_id": folder_id,
            "status": result.get("status", "unknown"),
        }, ensure_ascii=False)

    events = client.parse_calendar(result["elements"])
    events = events[:max_items]

    return json.dumps({
        "events": events,
        "count": len(events),
        "folder_id": folder_id,
        "sync_key": result.get("sync_key"),
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
    max: int = Query(50, ge=1, le=200, description="Maximum events"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Fetch calendar events. Default is primary Calendar folder."""
    _verify_key(x_api_key, authorization)
    c = _rest_client()
    fid = folder_id or c.find_folder(8)
    r = c.sync_folder(fid, window_size=max, body_type="1", body_size="1024")
    events = c.parse_calendar(r.get("elements", []))[:max]
    return {"events": events, "count": len(events)}


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
# Travel Plan (reads from Yandex Calendar via CalDAV)
# ============================================================
import time as _time
import re as _re
import httpx as _httpx

_ONLINE_PATTERNS = [
    _re.compile(r'https?://[\w.-]*zoom\.us', _re.I),
    _re.compile(r'https?://telemost\.yandex\.(ru|com)', _re.I),
    _re.compile(r'https?://meet\.google\.com', _re.I),
    _re.compile(r'https?://teams\.microsoft\.com', _re.I),
]

def _is_online(location: str) -> bool:
    for p in _ONLINE_PATTERNS:
        if p.search(location):
            return True
    return False

def _geocode(address: str, api_key: str, city: str = "Москва") -> dict:
    try:
        resp = _httpx.get(
            "https://geocode-maps.yandex.ru/1.x/",
            params={"apikey": api_key, "geocode": f"{city}, {address}", "format": "json", "results": "1"},
            timeout=10,
        )
        data = resp.json()
        members = data.get("response", {}).get("GeoObjectCollection", {}).get("featureMember", [])
        if not members:
            return {"error": f"Not found: {address}"}
        geo = members[0]["GeoObject"]
        pos = geo["Point"]["pos"].split()
        found = geo.get("metaDataProperty", {}).get("GeocoderMetaData", {}).get("text", address)
        return {"lat": float(pos[1]), "lon": float(pos[0]), "found_address": found}
    except Exception as e:
        return {"error": str(e)}

def _route_google(lat1, lon1, lat2, lon2, mode, departure_time, api_key) -> dict:
    try:
        travel_mode = "DRIVE" if mode == "driving" else "TRANSIT"
        body = {
            "origin": {"location": {"latLng": {"latitude": lat1, "longitude": lon1}}},
            "destination": {"location": {"latLng": {"latitude": lat2, "longitude": lon2}}},
            "travelMode": travel_mode,
        }
        if travel_mode == "DRIVE":
            body["routingPreference"] = "TRAFFIC_AWARE"
        if departure_time:
            from datetime import datetime as _dt, timezone as _tz
            dt = _dt.fromtimestamp(departure_time, tz=_tz.utc)
            body["departureTime"] = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        resp = _httpx.post(
            "https://routes.googleapis.com/directions/v2:computeRoutes",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": api_key,
                "X-Goog-FieldMask": "routes.duration,routes.distanceMeters,routes.legs.steps.transitDetails",
            },
            json=body, timeout=15,
        )
        data = resp.json()
        if "error" in data:
            return {"error": data["error"].get("message", str(data["error"]))}
        routes = data.get("routes", [])
        if not routes:
            return {"error": "No route found"}
        route = routes[0]
        dur_str = route.get("duration", "0s")
        dur_sec = int(dur_str.replace("s", ""))
        dist_m = route.get("distanceMeters", 0)
        h, m = dur_sec // 3600, (dur_sec % 3600) // 60
        dur_text = f"{h} ч {m} мин" if h > 0 else f"{m} мин"
        dist_text = f"{dist_m/1000:.1f} км" if dist_m >= 1000 else f"{dist_m} м"
        transit_info = []
        for leg in route.get("legs", []):
            for step in leg.get("steps", []):
                td = step.get("transitDetails")
                if td:
                    line = td.get("transitLine", {})
                    vehicle = line.get("vehicle", {}).get("type", "")
                    name = line.get("nameShort") or line.get("name", "")
                    transit_info.append(f"{vehicle} {name}".strip())
        result = {"duration_sec": dur_sec, "duration_text": dur_text, "distance_text": dist_text}
        if transit_info:
            result["transit_details"] = transit_info
        return result
    except Exception as e:
        return {"error": str(e)}


def _fetch_caldav_events(date_str: str) -> list:
    """Fetch events for a specific date from Yandex Calendar via CalDAV REPORT.
    date_str: YYYYMMDD format
    Returns list of dicts with subject, start, end, location.
    """
    from datetime import datetime as _dt, timedelta as _td
    
    year, month, day = int(date_str[:4]), int(date_str[4:6]), int(date_str[6:8])
    start_utc = f"{date_str}T000000Z"
    # Next day
    next_day = _dt(year, month, day) + _td(days=1)
    end_utc = next_day.strftime("%Y%m%dT000000Z")

    xml_body = f"""<?xml version="1.0" encoding="utf-8"?>
<C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav">
  <D:prop>
    <D:getetag/>
    <C:calendar-data/>
  </D:prop>
  <C:filter>
    <C:comp-filter name="VCALENDAR">
      <C:comp-filter name="VEVENT">
        <C:time-range start="{start_utc}" end="{end_utc}"/>
      </C:comp-filter>
    </C:comp-filter>
  </C:filter>
</C:calendar-query>"""

    try:
        resp = _httpx.request(
            "REPORT",
            CALDAV_URL,
            headers={
                "Content-Type": "application/xml; charset=utf-8",
                "Depth": "1",
            },
            content=xml_body.encode("utf-8"),
            auth=(CALDAV_USERNAME, CALDAV_PASSWORD),
            timeout=30,
        )
        
        if resp.status_code not in (200, 207):
            return [{"error": f"CalDAV HTTP {resp.status_code}: {resp.text[:200]}"}]
        
        # Parse multistatus XML response
        import xml.etree.ElementTree as ET
        root = ET.fromstring(resp.text)
        
        ns = {
            "D": "DAV:",
            "C": "urn:ietf:params:xml:ns:caldav",
        }
        
        events = []
        for response_el in root.findall(".//D:response", ns):
            caldata_el = response_el.find(".//C:calendar-data", ns)
            if caldata_el is None or not caldata_el.text:
                continue
            
            ical = caldata_el.text
            # Parse iCalendar manually (no external lib needed)
            ev = _parse_ical_vevent(ical)
            if ev:
                events.append(ev)
        
        return events
        
    except Exception as e:
        return [{"error": str(e)}]


def _parse_ical_vevent(ical_text: str) -> dict:
    """Parse a VCALENDAR string and extract VEVENT fields."""
    lines = ical_text.replace("\r\n ", "").replace("\r\n\t", "").split("\n")
    # Handle actual line continuations
    unfolded = []
    for line in ical_text.splitlines():
        if line.startswith((" ", "\t")):
            if unfolded:
                unfolded[-1] += line[1:]
        else:
            unfolded.append(line)
    
    in_vevent = False
    ev = {}
    for line in unfolded:
        if line.strip() == "BEGIN:VEVENT":
            in_vevent = True
            ev = {}
            continue
        if line.strip() == "END:VEVENT":
            in_vevent = False
            continue
        if not in_vevent:
            continue
        
        # Parse KEY;params:VALUE or KEY:VALUE
        if ":" not in line:
            continue
        key_part, _, value = line.partition(":")
        # Strip parameters (e.g., DTSTART;TZID=Europe/Moscow)
        key = key_part.split(";")[0].strip().upper()
        value = value.strip()
        
        if key == "SUMMARY":
            ev["subject"] = value
        elif key == "DTSTART":
            ev["start"] = value
        elif key == "DTEND":
            ev["end"] = value
        elif key == "LOCATION":
            ev["location"] = value.replace("\\,", ",").replace("\\n", " ")
        elif key == "DESCRIPTION":
            ev["body"] = value.replace("\\n", "\n")
        elif key == "UID":
            ev["uid"] = value
        elif key == "ORGANIZER":
            ev["organizer"] = value
    
    return ev if ev.get("subject") or ev.get("start") else None


def _parse_ical_dt(dt_str: str):
    """Parse iCalendar datetime to Python datetime.
    Handles: 20260325T140000Z, 20260325T170000, 20260325
    """
    from datetime import datetime as _dt, timezone as _tz, timedelta as _td
    if not dt_str:
        return None
    dt_str = dt_str.strip()
    try:
        if dt_str.endswith("Z"):
            return _dt.strptime(dt_str, "%Y%m%dT%H%M%SZ").replace(tzinfo=_tz.utc)
        elif "T" in dt_str:
            # Assume Moscow time if no Z
            naive = _dt.strptime(dt_str, "%Y%m%dT%H%M%S")
            msk = _tz(_td(hours=3))
            return naive.replace(tzinfo=msk)
        else:
            # Date only (all-day event)
            return _dt.strptime(dt_str, "%Y%m%d").replace(tzinfo=_tz.utc)
    except:
        return None


@api.get("/api/travel_plan", tags=["Travel"], summary="Calculate travel plan for a day")
async def api_travel_plan(
    date: Optional[str] = Query(None, description="Date YYYY-MM-DD (default: today Moscow time)"),
    x_api_key: str = Header(default=None),
    authorization: str = Header(default=None),
):
    """Analyze Yandex Calendar events for a day and calculate travel times.
    
    Reads events from Yandex Calendar via CalDAV.
    Uses Yandex Geocoder for addresses, Google Routes API for driving (with traffic) and transit.
    Online meetings (Zoom/Telemost/Meet/Teams) are treated as home location.
    Events without location generate a warning.
    """
    _verify_key(x_api_key, authorization)

    if not YANDEX_MAPS_KEY:
        return JSONResponse({"error": "YANDEX_MAPS_KEY not set"}, status_code=500)
    if not GOOGLE_MAPS_KEY:
        return JSONResponse({"error": "GOOGLE_MAPS_KEY not set"}, status_code=500)
    if not HOME_ADDRESS:
        return JSONResponse({"error": "HOME_ADDRESS not set"}, status_code=500)
    if not CALDAV_URL:
        return JSONResponse({"error": "CALDAV_URL not set"}, status_code=500)

    from datetime import datetime as _dt, timezone as _tz, timedelta as _td

    now = _dt.now(_tz.utc)
    if date:
        target = date.replace("-", "")
    else:
        target = (now + _td(hours=3)).strftime("%Y%m%d")

    # Fetch events from Yandex Calendar
    raw_events = _fetch_caldav_events(target)
    
    if raw_events and "error" in raw_events[0]:
        return JSONResponse({"error": f"CalDAV error: {raw_events[0]['error']}"}, status_code=500)

    # Filter and sort
    day_events = [e for e in raw_events if e.get("subject") or e.get("start")]
    day_events.sort(key=lambda e: e.get("start", ""))

    if not day_events:
        return {"date": target, "events": 0, "segments": [], "warnings": [], "message": "No events for this day"}

    # Geocode home
    home_geo = _geocode(HOME_ADDRESS, YANDEX_MAPS_KEY, DEFAULT_CITY)
    if "error" in home_geo:
        return JSONResponse({"error": f"Cannot geocode home: {home_geo['error']}"}, status_code=500)

    # Process events
    processed = []
    for ev in day_events:
        loc = (ev.get("location") or "").strip()
        subject = ev.get("subject", "(no subject)")
        start_dt = _parse_ical_dt(ev.get("start", ""))
        end_dt = _parse_ical_dt(ev.get("end", ""))

        item = {
            "subject": subject,
            "start": ev.get("start", ""),
            "end": ev.get("end", ""),
            "start_dt": start_dt,
            "end_dt": end_dt,
            "location_raw": loc,
        }
        if not loc:
            item["type"] = "unknown"
            item["warning"] = f"Location not specified for '{subject}'"
        elif _is_online(loc):
            item["type"] = "online"
            item["lat"] = home_geo["lat"]
            item["lon"] = home_geo["lon"]
            item["resolved_address"] = "Home (online)"
        else:
            item["type"] = "offline"
            geo = _geocode(loc, YANDEX_MAPS_KEY, DEFAULT_CITY)
            if "error" in geo:
                item["type"] = "geocode_failed"
                item["warning"] = f"Cannot find '{loc}' for '{subject}'"
            else:
                item["lat"] = geo["lat"]
                item["lon"] = geo["lon"]
                item["resolved_address"] = geo["found_address"]
        processed.append(item)

    # Build travel segments: Home -> Event 1 -> Event 2 -> ...
    points = [{"type": "home", "lat": home_geo["lat"], "lon": home_geo["lon"],
               "resolved_address": home_geo["found_address"], "subject": "Home", "end_dt": None}]
    points.extend(processed)

    segments = []
    warnings = []

    for i in range(len(points) - 1):
        pf = points[i]
        pt = points[i + 1]
        seg = {"from": pf.get("subject", "?"), "to": pt.get("subject", "?"),
               "to_start": pt.get("start", ""), "to_location": pt.get("location_raw", "")}

        if pt.get("warning"):
            warnings.append(pt["warning"])
            seg["warning"] = pt["warning"]
            seg["routes"] = None
            segments.append(seg)
            continue

        if "lat" not in pf or "lat" not in pt:
            seg["warning"] = "Missing coordinates"
            seg["routes"] = None
            segments.append(seg)
            continue

        if abs(pf["lat"] - pt["lat"]) < 0.001 and abs(pf["lon"] - pt["lon"]) < 0.001:
            seg["routes"] = {"same_location": True}
            segments.append(seg)
            continue

        from_end = pf.get("end_dt")
        to_start = pt.get("start_dt")
        gap_min = (to_start - from_end).total_seconds() / 60 if from_end and to_start else None
        departure_ts = from_end.timestamp() if from_end else None

        driving = _route_google(pf["lat"], pf["lon"], pt["lat"], pt["lon"], "driving", departure_ts, GOOGLE_MAPS_KEY)
        driving_min = driving.get("duration_sec", 0) / 60 if "error" not in driving else None

        seg["gap_minutes"] = round(gap_min) if gap_min is not None else None
        seg["taxi"] = {
            "duration_min": round(driving_min) if driving_min else None,
            "duration_text": driving.get("duration_text"),
            "distance": driving.get("distance_text"),
            "error": driving.get("error"),
        }

        if gap_min is not None and driving_min is not None:
            available = gap_min - BUFFER_MINUTES
            ok = driving_min <= available
            seg["feasibility"] = {
                "available_minutes": round(available),
                "ok": ok,
            }
            if not ok:
                w = f"Not enough time: '{pf.get('subject','')}' -> '{pt.get('subject','')}': gap {round(gap_min)} min, taxi {round(driving_min)} min + {BUFFER_MINUTES} min buffer"
                warnings.append(w)
                seg["feasibility"]["warning"] = w

        segments.append(seg)

    return {
        "date": target,
        "events": len(day_events),
        "buffer_minutes": BUFFER_MINUTES,
        "home": home_geo["found_address"],
        "segments": segments,
        "warnings": warnings,
    }





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
