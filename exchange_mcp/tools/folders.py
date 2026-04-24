"""Folder enumeration."""
from __future__ import annotations

from ..clients import router


def exchange_list_folders() -> dict:
    """List mailbox folders (Inbox, Sent, Calendar, etc.) with IDs and types.

    Returns a JSON-serializable dict with a "folders" key. Each folder
    has {id, name, type, parent}. `type` is a small integer role code
    (2=Inbox, 3=Drafts, 4=Deleted, 5=Sent, 8=Calendar, 9=Contacts, etc.)
    so downstream code can pick folders by semantic role instead of name.
    """
    folders, backend = router.list_folders()
    return {
        "backend": backend,
        "folders": [
            {"id": f.id, "name": f.name, "type": f.type, "parent": f.parent}
            for f in folders
        ],
        "count": len(folders),
    }


TOOLS = [exchange_list_folders]
