"""Unauthenticated liveness endpoint."""
from __future__ import annotations

import logging

from fastapi import APIRouter

from . import __version__
from .clients import router as mail_router
from .config import settings

log = logging.getLogger(__name__)
router = APIRouter()


@router.get("/health", tags=["meta"])
async def health() -> dict:
    health_info = mail_router.health_snapshot()
    ews = health_info.get("ews", {})
    ews_ok = ews.get("ok", False)

    if ews_ok:
        status_msg = "ok"
    else:
        status_msg = "down: " + (ews.get("last_error") or "ews unreachable")

    return {
        "status": status_msg,
        "version": __version__,
        "exchange_host": settings.exchange_host,
        "channels": {"ews": ews},
        "state": mail_router.state.snapshot(),
    }
