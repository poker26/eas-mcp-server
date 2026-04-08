"""
Minimal EWS client for calendar attachment metadata fallback.

Used only when EAS does not expose calendar attachment fields.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import List, Optional
import xml.etree.ElementTree as ET

import httpx


EWS_NAMESPACES = {
    "s": "http://schemas.xmlsoap.org/soap/envelope/",
    "m": "http://schemas.microsoft.com/exchange/services/2006/messages",
    "t": "http://schemas.microsoft.com/exchange/services/2006/types",
}


@dataclass
class EWSAttachmentMetadata:
    name: str
    content_type: str
    size: int
    is_inline: bool
    attachment_id: str

    def to_dict(self) -> dict:
        return {
            "source": "ews",
            "name": self.name,
            "content_type": self.content_type,
            "size": self.size,
            "is_inline": self.is_inline,
            "attachment_id": self.attachment_id,
        }


class EWSClient:
    def __init__(self, url: str, username: str, password: str):
        self.url = url
        self.client = httpx.Client(
            auth=(username, password),
            verify=False,
            timeout=30.0,
        )

    def close(self) -> None:
        self.client.close()

    def _post_soap(self, body_xml: str) -> ET.Element:
        response = self.client.post(
            self.url,
            headers={"Content-Type": "text/xml; charset=utf-8"},
            content=body_xml.encode("utf-8"),
        )
        response.raise_for_status()
        return ET.fromstring(response.text)

    @staticmethod
    def _to_iso_datetime(raw_datetime: datetime) -> str:
        return raw_datetime.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    @staticmethod
    def _parse_eas_datetime(raw_value: str) -> Optional[datetime]:
        if not raw_value:
            return None
        normalized_value = raw_value.replace(".000", "")
        normalized_value = normalized_value.replace("Z", "+00:00")
        if "T" in normalized_value and "-" not in normalized_value:
            if len(normalized_value) >= 15:
                normalized_value = (
                    f"{normalized_value[0:4]}-{normalized_value[4:6]}-{normalized_value[6:8]}"
                    f"T{normalized_value[9:11]}:{normalized_value[11:13]}:{normalized_value[13:15]}+00:00"
                )
        try:
            parsed_value = datetime.fromisoformat(normalized_value)
        except ValueError:
            return None
        if parsed_value.tzinfo is None:
            return parsed_value.replace(tzinfo=timezone.utc)
        return parsed_value.astimezone(timezone.utc)

    def _find_calendar_item_id(
        self,
        uid: str,
        subject: str,
        start_time: str,
    ) -> Optional[tuple[str, str]]:
        parsed_start = self._parse_eas_datetime(start_time) or datetime.now(timezone.utc)
        search_start = parsed_start - timedelta(days=2)
        search_end = parsed_start + timedelta(days=2)

        soap_request = f"""<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="calendar:UID"/>
          <t:FieldURI FieldURI="calendar:Start"/>
          <t:FieldURI FieldURI="item:HasAttachments"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView StartDate="{self._to_iso_datetime(search_start)}"
                      EndDate="{self._to_iso_datetime(search_end)}"
                      MaxEntriesReturned="200"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:ParentFolderIds>
    </m:FindItem>
  </s:Body>
</s:Envelope>"""

        root = self._post_soap(soap_request)
        calendar_items = root.findall(".//t:CalendarItem", EWS_NAMESPACES)
        if not calendar_items:
            return None

        uid = (uid or "").strip()
        subject = (subject or "").strip()
        target_start = self._parse_eas_datetime(start_time)

        best_match_item = None
        for calendar_item in calendar_items:
            item_uid = (calendar_item.findtext("t:UID", default="", namespaces=EWS_NAMESPACES) or "").strip()
            item_subject = (calendar_item.findtext("t:Subject", default="", namespaces=EWS_NAMESPACES) or "").strip()
            item_start_text = (calendar_item.findtext("t:Start", default="", namespaces=EWS_NAMESPACES) or "").strip()
            item_start = self._parse_eas_datetime(item_start_text)

            is_uid_match = bool(uid) and item_uid == uid
            is_subject_match = bool(subject) and item_subject == subject
            is_start_match = bool(target_start and item_start and abs((target_start - item_start).total_seconds()) < 120)

            if is_uid_match or (is_subject_match and is_start_match) or (is_subject_match and not target_start):
                best_match_item = calendar_item
                break

        if best_match_item is None:
            return None

        item_id_node = best_match_item.find("t:ItemId", EWS_NAMESPACES)
        if item_id_node is None:
            return None

        item_id = item_id_node.attrib.get("Id", "")
        change_key = item_id_node.attrib.get("ChangeKey", "")
        if not item_id:
            return None
        return (item_id, change_key)

    def get_calendar_attachment_metadata(
        self,
        uid: str,
        subject: str,
        start_time: str,
    ) -> List[dict]:
        item_id_tuple = self._find_calendar_item_id(uid=uid, subject=subject, start_time=start_time)
        if not item_id_tuple:
            return []
        item_id, change_key = item_id_tuple

        change_key_attribute = f' ChangeKey="{change_key}"' if change_key else ""
        soap_request = f"""<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
            xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <s:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="{item_id}"{change_key_attribute}/>
      </m:ItemIds>
    </m:GetItem>
  </s:Body>
</s:Envelope>"""

        root = self._post_soap(soap_request)
        file_attachments = root.findall(".//t:FileAttachment", EWS_NAMESPACES)
        attachment_metadata: List[EWSAttachmentMetadata] = []
        for file_attachment in file_attachments:
            attachment_id_node = file_attachment.find("t:AttachmentId", EWS_NAMESPACES)
            attachment_id = attachment_id_node.attrib.get("Id", "") if attachment_id_node is not None else ""
            name = file_attachment.findtext("t:Name", default="", namespaces=EWS_NAMESPACES) or ""
            content_type = file_attachment.findtext("t:ContentType", default="", namespaces=EWS_NAMESPACES) or ""
            size_text = file_attachment.findtext("t:Size", default="0", namespaces=EWS_NAMESPACES) or "0"
            is_inline_text = file_attachment.findtext("t:IsInline", default="false", namespaces=EWS_NAMESPACES) or "false"
            try:
                size = int(size_text)
            except ValueError:
                size = 0
            is_inline = is_inline_text.strip().lower() == "true"
            attachment_metadata.append(
                EWSAttachmentMetadata(
                    name=name,
                    content_type=content_type,
                    size=size,
                    is_inline=is_inline,
                    attachment_id=attachment_id,
                )
            )

        return [attachment.to_dict() for attachment in attachment_metadata]
