"""
MinIO storage helper.

This module is intentionally isolated from SyncKey-sensitive EAS flows.
It introduces a lazy, optional storage layer for attachments/media.
"""

from dataclasses import dataclass
from datetime import timedelta
from typing import Optional

from minio import Minio
from minio.error import S3Error


@dataclass
class MinioStorageConfig:
    endpoint: str
    access_key: str
    secret_key: str
    bucket_media: str
    use_ssl: bool = True

    @property
    def is_enabled(self) -> bool:
        return bool(
            self.endpoint and
            self.access_key and
            self.secret_key and
            self.bucket_media
        )


class MinioStorage:
    def __init__(self, config: MinioStorageConfig):
        self.config = config
        self._client: Optional[Minio] = None

    def is_enabled(self) -> bool:
        return self.config.is_enabled

    def _get_client(self) -> Minio:
        if self._client is None:
            self._client = Minio(
                endpoint=self.config.endpoint,
                access_key=self.config.access_key,
                secret_key=self.config.secret_key,
                secure=self.config.use_ssl,
            )
        return self._client

    def ensure_bucket_exists(self) -> None:
        if not self.is_enabled():
            return
        client = self._get_client()
        if not client.bucket_exists(self.config.bucket_media):
            client.make_bucket(self.config.bucket_media)

    def upload_bytes(
        self,
        object_key: str,
        payload: bytes,
        content_type: str = "application/octet-stream",
    ) -> None:
        if not self.is_enabled():
            raise RuntimeError("MinIO storage is disabled")

        client = self._get_client()
        from io import BytesIO

        client.put_object(
            bucket_name=self.config.bucket_media,
            object_name=object_key,
            data=BytesIO(payload),
            length=len(payload),
            content_type=content_type,
        )

    def object_exists(self, object_key: str) -> bool:
        if not self.is_enabled():
            return False
        client = self._get_client()
        try:
            client.stat_object(
                bucket_name=self.config.bucket_media,
                object_name=object_key,
            )
            return True
        except S3Error as error:
            if error.code in ("NoSuchKey", "NoSuchObject", "NotFound"):
                return False
            raise

    def presigned_get_url(
        self,
        object_key: str,
        ttl_seconds: int = 86400,
    ) -> str:
        if not self.is_enabled():
            raise RuntimeError("MinIO storage is disabled")
        client = self._get_client()
        return client.presigned_get_object(
            bucket_name=self.config.bucket_media,
            object_name=object_key,
            expires=timedelta(seconds=ttl_seconds),
        )
