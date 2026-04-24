"""Backend interface and DTOs for the EWS driver.

The driver normalizes exchangelib objects to `MailItem` / `FolderInfo`
so the router and MCP tools stay backend-agnostic. The Protocol is
kept as an extension point for future backends (IMAP, Graph, etc.).
InternetMessageId is used as the dedup key.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Protocol, Optional, runtime_checkable


@dataclass
class FolderInfo:
    id: str
    name: str
    # Role code: 2=Inbox, 3=Drafts, 4=Deleted, 5=Sent, 8=Calendar, 9=Contacts.
    # Lets the MCP layer pick folders by role instead of localized name.
    type: Optional[int] = None
    parent: Optional[str] = None


@dataclass
class MailItem:
    backend: str               # backend name, e.g. "ews"
    server_id: str             # backend-native id; not portable across backends
    message_id: str            # RFC 5322 Message-ID — portable, dedup key
    subject: str = ""
    sender: str = ""
    to: str = ""
    cc: str = ""
    received: Optional[datetime] = None  # UTC
    read: bool = False
    has_attachments: bool = False
    preview: str = ""
    body: str = ""
    body_is_html: bool = False

    def to_dict(self) -> dict:
        return {
            "backend": self.backend,
            "id": self.server_id,
            "message_id": self.message_id,
            "subject": self.subject,
            "from": self.sender,
            "to": self.to,
            "cc": self.cc,
            "date": self.received.isoformat() if self.received else "",
            "read": self.read,
            "has_attachments": self.has_attachments,
            "preview": self.preview,
            "body": self.body,
            "body_is_html": self.body_is_html,
        }


@dataclass
class CalendarItem:
    backend: str
    server_id: str
    uid: str                   # iCalendar UID — portable, dedup key
    subject: str = ""
    location: str = ""
    organizer: str = ""
    start: Optional[datetime] = None
    end: Optional[datetime] = None
    all_day: bool = False
    body: str = ""
    attendees: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "backend": self.backend,
            "id": self.server_id,
            "uid": self.uid,
            "subject": self.subject,
            "location": self.location,
            "organizer": self.organizer,
            "start": self.start.isoformat() if self.start else "",
            "end": self.end.isoformat() if self.end else "",
            "all_day": self.all_day,
            "body": self.body,
            "attendees": self.attendees,
        }


class BackendError(Exception):
    """Raised by a driver on an unrecoverable protocol error.

    The router marks the backend unhealthy on BackendError. Transient
    errors should be retried inside the driver before surfacing.
    """


@runtime_checkable
class MailBackend(Protocol):
    name: str                  # backend name, e.g. "ews"

    def healthcheck(self) -> bool: ...
    def list_folders(self) -> list[FolderInfo]: ...

    def get_items_since(
        self,
        folder_id: str,
        since: Optional[datetime],
        limit: int = 50,
        include_body: bool = True,
    ) -> list[MailItem]: ...

    def get_item(self, folder_id: str, server_id: str) -> Optional[MailItem]: ...

    def send_email(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: Optional[list[str]] = None,
        body_is_html: bool = False,
    ) -> None: ...
