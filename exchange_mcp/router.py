"""Thin service over the EWS backend.

State flow for `get_new_mail`:
    1. Compute query window: `since = cursor - SAFETY_MARGIN` (or None
       if no cursor yet). Safety margin bridges clock drift.
    2. Fetch items from EWS.
    3. Filter against Message-ID LRU → new items only.
    4. Advance cursor to max(received) of returned items.
    5. Record new Message-IDs in LRU.

Originally designed as a router that could fall back to an EAS channel
when VPN was down; EAS was split off into its own MCP server, so this
layer is now a single-backend wrapper. The Protocol and `_try` helper
are kept as hooks for future backends (IMAP, Graph, etc.) without
rewriting tools/.
"""
from __future__ import annotations

import logging
import os
import threading
import time
from datetime import datetime, timedelta, timezone
from typing import Optional

from .backends.base import BackendError, FolderInfo, MailBackend, MailItem
from .backends.ews import EWSBackend
from .config import settings
from .state import SharedState

logger = logging.getLogger(__name__)


_SAFETY_MARGIN = timedelta(minutes=5)
_HEALTH_TTL = 60.0  # seconds


class MailRouter:
    def __init__(self) -> None:
        self.ews = EWSBackend()
        self.state = SharedState(
            path=os.path.join(settings.state_dir, "router_state.json"),
        )
        self._health: dict[str, tuple[bool, float]] = {}
        self._health_lock = threading.RLock()

    # --- health ------------------------------------------------------
    def _healthy(self, backend: MailBackend) -> bool:
        with self._health_lock:
            entry = self._health.get(backend.name)
            if entry is not None:
                ok, ts = entry
                if time.monotonic() - ts < _HEALTH_TTL:
                    return ok
        ok = False
        try:
            ok = backend.healthcheck()
        except Exception as e:
            logger.warning("%s healthcheck raised: %s", backend.name, e)
        with self._health_lock:
            self._health[backend.name] = (ok, time.monotonic())
        return ok

    def _mark_unhealthy(self, backend: MailBackend) -> None:
        with self._health_lock:
            self._health[backend.name] = (False, time.monotonic())

    def health_snapshot(self) -> dict:
        ok = self._healthy(self.ews)
        last_err = getattr(self.ews, "last_error", lambda: None)()
        return {
            "ews": {"ok": ok, "last_error": last_err},
            "preferred": "ews",
        }

    # --- routing -----------------------------------------------------
    def _backend_order(self) -> list[MailBackend]:
        return [self.ews]

    def _try(self, op_name: str, fn):
        """Run `fn(backend)` on EWS. Kept as a wrapper so tools/ can stay
        backend-agnostic if we add a second channel later."""
        backend = self.ews
        try:
            return fn(backend), backend
        except BackendError as e:
            self._mark_unhealthy(backend)
            logger.warning("%s on %s: %s", op_name, backend.name, e)
            raise
        except Exception as e:
            self._mark_unhealthy(backend)
            logger.warning(
                "%s on %s raised %s: %s",
                op_name, backend.name, type(e).__name__, e,
            )
            raise BackendError(f"{op_name}: {type(e).__name__}: {e}") from e

    # --- operations --------------------------------------------------
    def list_folders(self) -> tuple[list[FolderInfo], str]:
        result, backend = self._try("list_folders", lambda b: b.list_folders())
        return result, backend.name

    def get_new_mail(
        self,
        folder_id: str,
        limit: int = 50,
        include_body: bool = True,
    ) -> tuple[list[MailItem], str, bool]:
        """Return (new_items, backend_used, is_initial).

        is_initial=True means no prior cursor existed; the caller may
        want to treat the first response specially (e.g. don't spam
        the user with a week of backlog).
        """
        cursor = self.state.get_cursor(folder_id)
        is_initial = cursor is None
        since = (cursor - _SAFETY_MARGIN) if cursor else None

        def _op(b: MailBackend) -> list[MailItem]:
            return b.get_items_since(folder_id, since, limit=limit,
                                     include_body=include_body)

        items, backend = self._try("get_new_mail", _op)

        new_items: list[MailItem] = []
        new_mids: list[str] = []
        max_received: Optional[datetime] = None
        for m in items:
            if m.message_id and self.state.contains(folder_id, m.message_id):
                continue
            new_items.append(m)
            if m.message_id:
                new_mids.append(m.message_id)
            if m.received and (max_received is None or m.received > max_received):
                max_received = m.received

        if max_received is not None:
            self.state.set_cursor(folder_id, max_received)
        if new_mids:
            self.state.mark_seen(folder_id, new_mids)

        return new_items, backend.name, is_initial
