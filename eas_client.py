"""
EAS Client Library for Exchange ActiveSync
Handles WBXML encoding/decoding and EAS protocol commands.
"""

import io
import json
import logging
import uuid
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from typing import Optional

import httpx

logger = logging.getLogger(__name__)

# ============================================================
# WBXML Code Pages (verified against Exchange 2019)
# ============================================================

CP_AIRSYNC = {
    0x05: "Sync", 0x06: "Responses", 0x07: "Add",
    0x08: "Change", 0x09: "Delete", 0x0A: "Fetch",
    0x0B: "SyncKey", 0x0C: "ClientId", 0x0D: "ServerId",
    0x0E: "Status", 0x0F: "Collection", 0x10: "Class",
    0x11: "Version", 0x12: "CollectionId", 0x13: "GetChanges",
    0x14: "MoreAvailable", 0x15: "WindowSize", 0x16: "Commands",
    0x17: "Options", 0x18: "FilterType", 0x1B: "Conflict",
    0x1C: "Collections", 0x1D: "ApplicationData",
    0x1E: "DeletesAsMoves", 0x22: "MIMESupport",
}

CP_CONTACTS = {
    0x05: "Anniversary", 0x06: "AssistantName",
    0x07: "AssistantPhoneNumber", 0x08: "Birthday",
    0x09: "Body", 0x0A: "BodySize", 0x0B: "BodyTruncated",
    0x0C: "Business2PhoneNumber", 0x0D: "BusinessCity",
    0x0E: "BusinessCountry", 0x0F: "BusinessPostalCode",
    0x10: "BusinessState", 0x11: "BusinessStreet",
    0x12: "BusinessFaxNumber", 0x13: "BusinessPhoneNumber",
    0x14: "CarPhoneNumber", 0x15: "Categories",
    0x16: "Category", 0x17: "Children", 0x18: "Child",
    0x19: "CompanyName", 0x1A: "Department",
    0x1B: "Email1Address", 0x1C: "Email2Address",
    0x1D: "Email3Address", 0x1E: "FileAs",
    0x1F: "FirstName", 0x20: "Home2PhoneNumber",
    0x21: "HomeCity", 0x22: "HomeCountry",
    0x23: "HomePostalCode", 0x24: "HomeState",
    0x25: "HomeStreet", 0x26: "HomeFaxNumber",
    0x27: "HomePhoneNumber", 0x28: "JobTitle",
    0x29: "LastName", 0x2A: "MiddleName",
    0x2B: "MobilePhoneNumber", 0x2C: "OfficeLocation",
    0x33: "RadioPhoneNumber", 0x34: "Spouse",
    0x35: "Suffix", 0x36: "Title", 0x37: "WebPage",
}

CP_EMAIL = {
    0x05: "Attachment", 0x06: "Attachments", 0x07: "AttName",
    0x08: "AttSize", 0x0C: "Body", 0x0D: "BodySize",
    0x0E: "BodyTruncated", 0x0F: "DateReceived",
    0x10: "DisplayName", 0x11: "DisplayTo", 0x12: "Importance",
    0x13: "MessageClass", 0x14: "Subject", 0x15: "Read",
    0x16: "To", 0x17: "Cc", 0x18: "From", 0x19: "ReplyTo",
    0x1A: "AllDayEvent", 0x1D: "DtStamp", 0x1E: "EndTime",
    0x20: "BusyStatus", 0x21: "Location",
    0x23: "Organizer", 0x31: "StartTime",
    0x35: "ThreadTopic", 0x38: "MIMESize",
    0x39: "InternetCPID", 0x3C: "Flag", 0x3D: "FlagStatus",
}

CP_CALENDAR = {
    0x05: "TimeZone", 0x06: "AllDayEvent", 0x07: "Attendees",
    0x08: "Attendee", 0x09: "Attendee_Email",
    0x0A: "Attendee_Name", 0x0B: "Body", 0x0C: "BodyTruncated",
    0x0D: "BusyStatus", 0x0E: "Categories", 0x0F: "Category",
    0x11: "DtStamp", 0x12: "EndTime",
    0x13: "Exception", 0x14: "Exceptions",
    0x16: "Exception_StartTime", 0x17: "Location",
    0x18: "MeetingStatus", 0x19: "Organizer_Email",
    0x1A: "Organizer_Name", 0x1B: "Recurrence",
    0x24: "Reminder", 0x25: "Sensitivity", 0x26: "Subject",
    0x27: "StartTime", 0x28: "UID",
    0x29: "Attendee_Status", 0x2A: "Attendee_Type",
}

CP_FOLDER = {
    0x07: "DisplayName", 0x08: "ServerId", 0x09: "ParentId",
    0x0A: "Type", 0x0C: "Status", 0x0E: "Changes",
    0x0F: "Add", 0x10: "Delete", 0x11: "Update",
    0x12: "SyncKey", 0x16: "FolderSync", 0x17: "Count",
}

CP_AIRSYNCBASE = {
    0x05: "BodyPreference", 0x06: "Type", 0x07: "TruncationSize",
    0x0A: "Body", 0x0B: "Data", 0x0C: "EstimatedDataSize",
    0x0D: "Truncated", 0x0E: "Attachments", 0x0F: "Attachment",
    0x10: "DisplayName", 0x11: "FileReference",
    0x12: "Method", 0x15: "IsInline",
    0x16: "NativeBodyType", 0x17: "ContentType",
    0x19: "Preview",
}

ALL_PAGES = {
    0: CP_AIRSYNC, 1: CP_CONTACTS, 2: CP_EMAIL,
    4: CP_CALENDAR, 7: CP_FOLDER, 17: CP_AIRSYNCBASE,
}

FOLDER_TYPES = {
    1: "Generic", 2: "Inbox", 3: "Drafts", 4: "Deleted",
    5: "Sent", 6: "Outbox", 7: "Tasks", 8: "Calendar",
    9: "Contacts", 10: "Notes", 11: "Journal",
    12: "User Mail", 13: "User Calendar", 14: "User Contacts",
    15: "User Tasks", 17: "User Notes", 19: "Recipient Cache",
}


# ============================================================
# WBXML Encoder
# ============================================================
class WBXMLEncoder:
    def __init__(self):
        self.buf = bytearray(b'\x03\x01\x6a\x00')
        self.page = 0

    def switch(self, page):
        if page != self.page:
            self.buf.extend([0x00, page])
            self.page = page

    def tag_open(self, page, tag):
        self.switch(page)
        self.buf.append(tag | 0x40)

    def tag_empty(self, page, tag):
        self.switch(page)
        self.buf.append(tag)

    def end(self):
        self.buf.append(0x01)

    def string(self, s):
        self.buf.append(0x03)
        self.buf.extend(s.encode('utf-8'))
        self.buf.append(0x00)

    def tag_str(self, page, tag, value):
        self.tag_open(page, tag)
        self.string(value)
        self.end()

    def get(self):
        return bytes(self.buf)


# ============================================================
# WBXML Decoder
# ============================================================
class WBXMLDecoder:
    def __init__(self, data):
        self.data = data
        self.pos = 0
        self.page = 0

    def _rb(self):
        if self.pos >= len(self.data):
            return None
        b = self.data[self.pos]
        self.pos += 1
        return b

    def _rmb(self):
        r = 0
        while self.pos < len(self.data):
            b = self.data[self.pos]
            self.pos += 1
            r = (r << 7) | (b & 0x7F)
            if not (b & 0x80):
                return r
        return r

    def _rstr(self):
        c = []
        while self.pos < len(self.data):
            b = self.data[self.pos]
            self.pos += 1
            if b == 0:
                break
            c.append(b)
        return bytes(c).decode('utf-8', errors='replace')

    def decode(self):
        self._rb()
        self._rmb()
        self._rmb()
        stl = self._rmb()
        self.pos += stl
        result = []
        self._parse(result, 0)
        return result

    def _parse(self, result, depth):
        while self.pos < len(self.data):
            t = self._rb()
            if t is None or t == 0x01:
                return
            if t == 0x00:
                self.page = self._rb()
                continue
            if t == 0x03:
                s = self._rstr()
                if result and result[-1][2] is None:
                    result[-1] = (result[-1][0], result[-1][1], s)
                continue
            if t == 0xC3:
                ln = self._rmb()
                self.pos += ln
                continue
            hc = bool(t & 0x40)
            tid = t & 0x3F
            cp = ALL_PAGES.get(self.page, {})
            name = cp.get(tid, f"p{self.page}:0x{tid:02X}")
            result.append((depth, name, None))
            if hc:
                self._parse(result, depth + 1)


# ============================================================
# EAS Client
# ============================================================
class EASClient:
    def __init__(self, host: str, username: str, password: str,
                 device_id: str = "EAS0LEGCLIENT0001",
                 device_type: str = "EASClient",
                 protocol_version: str = "14.1",
                 email_address: str = ""):
        self.host = host
        self.url = f"https://{host}/Microsoft-Server-ActiveSync"
        self.username = username
        self.email_address = email_address or username
        self.device_id = device_id
        self.device_type = device_type
        self.protocol_version = protocol_version
        self.client = httpx.Client(
            auth=(username, password),
            verify=False,
            timeout=30.0,
        )
        self.folders: dict = {}
        self.sync_keys: dict = {}

    def close(self):
        self.client.close()

    def _post(self, cmd: str, wbxml: bytes) -> httpx.Response:
        return self.client.post(
            self.url,
            params={"Cmd": cmd, "User": self.username,
                    "DeviceId": self.device_id,
                    "DeviceType": self.device_type},
            headers={"MS-ASProtocolVersion": self.protocol_version,
                     "Content-Type": "application/vnd.ms-sync.wbxml"},
            content=wbxml,
        )

    def _decode(self, resp: httpx.Response) -> list:
        if resp.status_code == 200 and resp.content:
            return WBXMLDecoder(resp.content).decode()
        return []

    def _find(self, elements: list, tag_name: str) -> Optional[str]:
        for _, tag, val in elements:
            if tag == tag_name and val is not None:
                return val
        return None

    # --- FolderSync ---
    def folder_sync(self) -> dict:
        enc = WBXMLEncoder()
        enc.tag_open(7, 0x16)
        enc.tag_str(7, 0x12, "0")
        enc.end()

        resp = self._post("FolderSync", enc.get())
        elements = self._decode(resp)

        if self._find(elements, "Status") != "1":
            logger.error("FolderSync failed: %s", self._find(elements, "Status"))
            return {}

        folders = {}
        cur = {}
        for _, tag, value in elements:
            if tag == "DisplayName" and value:
                cur["name"] = value
            elif tag == "ServerId" and value:
                cur["id"] = value
            elif tag == "Type" and value:
                cur["type"] = int(value)
            elif tag == "ParentId" and value:
                cur["parent"] = value
            elif tag == "Add" and value is None:
                if cur.get("id"):
                    folders[cur["id"]] = cur
                cur = {}
        if cur.get("id"):
            folders[cur["id"]] = cur

        self.folders = folders
        return folders

    def find_folder(self, folder_type: int) -> Optional[str]:
        for fid, f in self.folders.items():
            if f.get("type") == folder_type:
                return fid
        return None

    # --- Sync ---
    def sync(self, collection_id: str, sync_key: str = "0",
             window_size: int = 50, body_type: str = "1",
             body_size: str = "51200") -> dict:
        enc = WBXMLEncoder()
        enc.tag_open(0, 0x05)   # Sync
        enc.tag_open(0, 0x1C)   # Collections
        enc.tag_open(0, 0x0F)   # Collection
        enc.tag_str(0, 0x0B, sync_key)
        enc.tag_str(0, 0x12, str(collection_id))

        if sync_key != "0":
            enc.tag_empty(0, 0x13)  # GetChanges
            enc.tag_str(0, 0x15, str(window_size))
            enc.tag_open(0, 0x17)   # Options
            enc.tag_open(17, 0x05)  # BodyPreference
            enc.tag_str(17, 0x06, body_type)
            enc.tag_str(17, 0x07, body_size)
            enc.end()  # BodyPreference
            enc.end()  # Options

        enc.end()  # Collection
        enc.end()  # Collections
        enc.end()  # Sync

        resp = self._post("Sync", enc.get())

        if resp.status_code != 200:
            return {"status": f"HTTP {resp.status_code}", "items": []}
        if not resp.content:
            return {"status": "no_changes", "items": [], "sync_key": sync_key}

        elements = self._decode(resp)
        status = self._find(elements, "Status")
        new_key = self._find(elements, "SyncKey")

        if new_key:
            self.sync_keys[collection_id] = new_key

        return {"status": status, "sync_key": new_key, "elements": elements}

    def sync_folder(self, collection_id: str, **kwargs) -> dict:
        r1 = self.sync(collection_id, "0")
        key = r1.get("sync_key")
        if not key:
            return r1
        return self.sync(collection_id, key, **kwargs)

    # --- Parsers ---
    def parse_emails(self, elements: list) -> list:
        emails = []
        cur = {}
        for _, tag, value in elements:
            if tag == "ServerId" and value:
                if cur.get("subject") or cur.get("from"):
                    emails.append(cur)
                cur = {"server_id": value}
                continue
            if value is not None:
                mapping = {
                    "Subject": "subject", "From": "from", "To": "to",
                    "Cc": "cc", "DateReceived": "date",
                    "DisplayTo": "display_to", "Importance": "importance",
                    "Read": "read", "MessageClass": "class",
                    "Data": "body", "Preview": "preview",
                    "EstimatedDataSize": "size", "ThreadTopic": "thread_topic",
                }
                if tag in mapping:
                    cur[mapping[tag]] = value
        if cur.get("subject") or cur.get("from"):
            emails.append(cur)
        return emails

    def parse_calendar(self, elements: list) -> list:
        events = []
        cur = {}
        for _, tag, value in elements:
            if tag == "ServerId" and value:
                if cur.get("subject") or cur.get("start"):
                    events.append(cur)
                cur = {"server_id": value}
                continue
            if value is None:
                continue
            mapping = {
                "Subject": "subject", "StartTime": "start",
                "EndTime": "end", "Location": "location",
                "Organizer_Name": "organizer_name",
                "Organizer_Email": "organizer_email",
                "AllDayEvent": "all_day", "BusyStatus": "busy_status",
                "Reminder": "reminder", "UID": "uid", "DtStamp": "stamp",
                "MeetingStatus": "meeting_status",
            }
            if tag == "Attendee_Name":
                cur.setdefault("attendees", []).append({"name": value})
            elif tag == "Attendee_Email":
                att = cur.get("attendees", [])
                if att:
                    att[-1]["email"] = value
            elif tag in mapping:
                cur[mapping[tag]] = value
        if cur.get("subject") or cur.get("start"):
            events.append(cur)
        return events

    def parse_contacts(self, elements: list) -> list:
        contacts = []
        cur = {}
        fields = {
            "FileAs", "FirstName", "LastName", "MiddleName",
            "CompanyName", "Department", "JobTitle",
            "Email1Address", "Email2Address", "Email3Address",
            "BusinessPhoneNumber", "MobilePhoneNumber",
            "HomePhoneNumber", "BusinessCity", "BusinessStreet",
        }
        for _, tag, value in elements:
            if tag == "ServerId" and value:
                if cur.get("FileAs") or cur.get("FirstName"):
                    contacts.append(cur)
                cur = {"server_id": value}
                continue
            if value is not None and tag in fields:
                cur[tag] = value
        if cur.get("FileAs") or cur.get("FirstName"):
            contacts.append(cur)
        return contacts

    # --- SendMail ---
    def send_email(self, to: str, subject: str, body: str,
                   cc: str = "", content_type: str = "plain") -> dict:
        """Send an email via EAS SendMail command.
        
        Args:
            to: Recipient email (comma-separated for multiple)
            subject: Email subject
            body: Email body text
            cc: CC recipients (optional)
            content_type: 'plain' or 'html'
        
        Returns:
            dict with status
        """
        # Build MIME message
        msg = MIMEText(body, content_type, "utf-8")
        msg["From"] = self.email_address
        msg["To"] = to
        msg["Subject"] = subject
        if cc:
            msg["Cc"] = cc
        msg["Date"] = datetime.utcnow().strftime("%a, %d %b %Y %H:%M:%S +0000")
        msg["Message-ID"] = f"<{uuid.uuid4()}@eas-mcp>"

        mime_data = msg.as_bytes()

        # SendMail uses raw MIME, not WBXML
        # But needs a small WBXML wrapper with ClientId
        client_id = str(uuid.uuid4())[:8]

        # Build SendMail WBXML
        enc = WBXMLEncoder()
        # Page 21: ComposeMail
        enc.switch(21)
        # <SendMail> tag 0x05 with content
        enc.buf.append(0x05 | 0x40)
        # <ClientId> tag 0x11
        enc.buf.append(0x11 | 0x40)
        enc.string(client_id)
        enc.end()
        # <SaveInSentItems/> tag 0x08
        enc.buf.append(0x08)
        # <Mime> tag 0x10 with content (opaque data)
        enc.buf.append(0x10 | 0x40)
        # Write opaque: token 0xC3 + mb_uint32 length + data
        enc.buf.append(0xC3)
        # Encode length as mb_uint32
        length = len(mime_data)
        length_bytes = []
        while True:
            length_bytes.insert(0, length & 0x7F)
            length >>= 7
            if length == 0:
                break
        for i in range(len(length_bytes) - 1):
            length_bytes[i] |= 0x80
        enc.buf.extend(length_bytes)
        enc.buf.extend(mime_data)
        enc.end()  # Mime
        enc.end()  # SendMail

        resp = self._post("SendMail", enc.get())

        if resp.status_code == 200:
            # Empty 200 = success for SendMail
            if not resp.content:
                return {"status": "sent", "to": to, "subject": subject}
            # Non-empty = error in WBXML
            elements = self._decode(resp)
            status = self._find(elements, "Status")
            return {"status": f"error_{status}", "elements": elements}
        else:
            return {"status": f"HTTP {resp.status_code}"}

    # --- Create Calendar Event ---
    def create_event(self, subject: str, start_time: str, end_time: str,
                     location: str = "", body: str = "",
                     attendees: list = None, all_day: bool = False,
                     reminder: int = 15) -> dict:
        """Create a calendar event via EAS Sync Add command.
        
        Args:
            subject: Event title
            start_time: ISO format e.g. '2026-03-25T10:00:00.000Z'
            end_time: ISO format e.g. '2026-03-25T11:00:00.000Z'
            location: Event location (optional)
            body: Event description (optional)
            attendees: List of dicts [{"email": "...", "name": "..."}] (optional)
            all_day: Whether it's an all-day event
            reminder: Reminder in minutes (default 15)
        
        Returns:
            dict with status and server_id
        """
        cal_id = self.find_folder(8)  # Calendar
        if not cal_id:
            return {"status": "error", "message": "Calendar folder not found"}

        # Get sync key first
        r1 = self.sync(cal_id, "0")
        sync_key = r1.get("sync_key")
        if not sync_key:
            return {"status": "error", "message": "Failed to get SyncKey"}

        client_id = str(uuid.uuid4())[:8]

        enc = WBXMLEncoder()
        # <Sync> page 0
        enc.tag_open(0, 0x05)
        # <Collections>
        enc.tag_open(0, 0x1C)
        # <Collection>
        enc.tag_open(0, 0x0F)
        # <SyncKey>
        enc.tag_str(0, 0x0B, sync_key)
        # <CollectionId>
        enc.tag_str(0, 0x12, str(cal_id))
        # <Commands>
        enc.tag_open(0, 0x16)
        # <Add>
        enc.tag_open(0, 0x07)
        # <ClientId>
        enc.tag_str(0, 0x0C, client_id)
        # <ApplicationData>
        enc.tag_open(0, 0x1D)

        # Calendar fields - page 4
        # <Subject>
        enc.tag_str(4, 0x26, subject)
        # <StartTime>
        enc.tag_str(4, 0x27, start_time)
        # <EndTime>
        enc.tag_str(4, 0x12, end_time)
        # <Location>
        if location:
            enc.tag_str(4, 0x17, location)
        # <AllDayEvent>
        enc.tag_str(4, 0x06, "1" if all_day else "0")
        # <BusyStatus> 2=Busy
        enc.tag_str(4, 0x0D, "2")
        # <MeetingStatus> 1=Meeting
        enc.tag_str(4, 0x18, "1" if attendees else "0")
        # <Reminder>
        enc.tag_str(4, 0x24, str(reminder))
        # <Sensitivity> 0=Normal
        enc.tag_str(4, 0x25, "0")
        # <DtStamp>
        enc.tag_str(4, 0x11, datetime.utcnow().strftime("%Y%m%dT%H%M%SZ"))
        # <UID>
        enc.tag_str(4, 0x28, str(uuid.uuid4()))

        # Body - AirSyncBase page 17
        if body:
            enc.tag_open(17, 0x0A)  # Body
            enc.tag_str(17, 0x06, "1")  # Type = plain text
            enc.tag_str(17, 0x0B, body)  # Data
            enc.end()  # Body

        # Attendees
        if attendees:
            enc.tag_open(4, 0x07)  # Attendees
            for att in attendees:
                enc.tag_open(4, 0x08)  # Attendee
                enc.tag_str(4, 0x09, att.get("email", ""))
                enc.tag_str(4, 0x0A, att.get("name", att.get("email", "")))
                enc.tag_str(4, 0x2A, "1")  # Type=Required
                enc.tag_str(4, 0x29, "0")  # Status=None
                enc.end()  # Attendee
            enc.end()  # Attendees

        enc.end()  # ApplicationData
        enc.end()  # Add
        enc.end()  # Commands
        enc.end()  # Collection
        enc.end()  # Collections
        enc.end()  # Sync

        resp = self._post("Sync", enc.get())

        if resp.status_code == 200 and resp.content:
            elements = self._decode(resp)
            status = self._find(elements, "Status")
            server_id = self._find(elements, "ServerId")
            return {
                "status": "created" if status == "1" else f"error_{status}",
                "server_id": server_id,
                "client_id": client_id,
            }
        elif resp.status_code == 200:
            return {"status": "created", "client_id": client_id}
        else:
            return {"status": f"HTTP {resp.status_code}"}
