"""FastMCP setup: register every tool from `exchange_mcp.tools`."""
from __future__ import annotations

import logging

from fastmcp import FastMCP

from .tools import ALL_TOOLS

logger = logging.getLogger(__name__)


def build_mcp() -> FastMCP:
    mcp = FastMCP(
        name="exchange-mcp",
        instructions=(
            "Exchange MCP server (EWS backend). Tools expose folders, "
            "mail, calendar, contacts and search. A per-folder timestamp "
            "cursor plus a Message-ID LRU dedup new-mail calls so clients "
            "don't see the same message twice on repeated polls."
        ),
    )
    for fn in ALL_TOOLS:
        mcp.tool(fn)
    logger.info("Registered %d tools with FastMCP", len(ALL_TOOLS))
    return mcp


mcp = build_mcp()
