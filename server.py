"""
Exchange ActiveSync MCP Server
Provides tools for reading email, calendar, and contacts via EAS protocol.
"""

import json
import os
import logging
from contextlib import asynccontextmanager
from typing import Optional, List

from mcp.server.fastmcp import FastMCP
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


# ============================================================
# Lifespan: initialize EAS client
# ============================================================
@asynccontextmanager
async def app_lifespan():
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
    )

    logger.info("Connecting to Exchange: %s as %s", EAS_HOST, EAS_USERNAME)
    folders = client.folder_sync()
    logger.info("Found %d folders", len(folders))

    yield {"eas": client}

    client.close()


# ============================================================
# MCP Server
# ============================================================
mcp = FastMCP("exchange_eas_mcp", lifespan=app_lifespan)


def get_client(ctx) -> EASClient:
    return ctx.request_context.lifespan_state["eas"]


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
async def exchange_list_folders(params: ListFoldersInput, ctx=None) -> str:
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
        if params.folder_type is not None and ft != params.folder_type:
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
async def exchange_get_emails(params: GetEmailsInput, ctx=None) -> str:
    """Fetch emails from an Exchange mailbox folder.

    By default reads from Inbox. Returns subject, from, to, date, read status.
    Set include_body=true to get email body text (plain text).

    Args:
        params: GetEmailsInput with folder_id, max_items, include_body

    Returns:
        JSON with list of emails containing subject, from, to, date, read status
    """
    client = get_client(ctx)

    folder_id = params.folder_id
    if not folder_id:
        folder_id = client.find_folder(2)  # Inbox
        if not folder_id:
            return json.dumps({"error": "Inbox not found. Run exchange_list_folders first."})

    body_size = "51200" if params.include_body else "0"
    result = client.sync_folder(
        folder_id,
        window_size=params.max_items,
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
    emails = emails[:params.max_items]

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
        if params.include_body and e.get("body"):
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
async def exchange_get_calendar(params: GetCalendarInput, ctx=None) -> str:
    """Fetch calendar events from Exchange.

    By default reads from the primary Calendar folder.
    Returns event subject, start/end times, location, organizer, attendees.

    Args:
        params: GetCalendarInput with folder_id, max_items

    Returns:
        JSON with list of calendar events
    """
    client = get_client(ctx)

    folder_id = params.folder_id
    if not folder_id:
        folder_id = client.find_folder(8)  # Calendar
        if not folder_id:
            return json.dumps({"error": "Calendar folder not found."})

    result = client.sync_folder(
        folder_id,
        window_size=params.max_items,
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
    events = events[:params.max_items]

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
async def exchange_get_contacts(params: GetContactsInput, ctx=None) -> str:
    """Fetch contacts from Exchange address book.

    By default reads from the primary Contacts folder.
    Returns name, email, phone, company, job title.

    Args:
        params: GetContactsInput with folder_id, max_items

    Returns:
        JSON with list of contacts
    """
    client = get_client(ctx)

    folder_id = params.folder_id
    if not folder_id:
        folder_id = client.find_folder(9)  # Contacts
        if not folder_id:
            return json.dumps({"error": "Contacts folder not found."})

    result = client.sync_folder(
        folder_id,
        window_size=params.max_items,
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
    contacts = contacts[:params.max_items]

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
async def exchange_search_emails(params: SearchEmailInput, ctx=None) -> str:
    """Search emails by subject, sender, or content.

    Fetches emails from the specified folder and filters locally by query.
    Query matches against subject, from, to, and preview fields (case-insensitive).

    Args:
        params: SearchEmailInput with query, folder_id, max_results

    Returns:
        JSON with matching emails
    """
    client = get_client(ctx)

    folder_id = params.folder_id
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
        return json.dumps({"results": [], "count": 0, "query": params.query})

    emails = client.parse_emails(result["elements"])

    q = params.query.lower()
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
            if len(matches) >= params.max_results:
                break

    return json.dumps({
        "results": matches,
        "count": len(matches),
        "query": params.query,
        "folder_id": folder_id,
    }, ensure_ascii=False, indent=2)


# ============================================================
# Entry point
# ============================================================
if __name__ == "__main__":
    import sys
    transport = "stdio"
    port = 8000
    for arg in sys.argv[1:]:
        if arg == "--http":
            transport = "streamable_http"
        elif arg.startswith("--port="):
            port = int(arg.split("=")[1])

    if transport == "streamable_http":
        logger.info("Starting MCP server on HTTP port %d", port)
        mcp.run(transport="streamable_http", port=port)
    else:
        logger.info("Starting MCP server on stdio")
        mcp.run()
