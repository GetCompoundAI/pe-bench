from datetime import datetime

from pydantic import BaseModel, field_serializer, field_validator


class DocumentCoreProperties(BaseModel):
    category: str | None = None
    content_status: str | None = None
    keywords: str | None = None
    last_modified_by: str | None = None
    last_printed: datetime | None = None
    revision: str | None = None
    version: str | None = None

    subject: str | None = None
    title: str | None = None
    creator: str | None = None
    description: str | None = None
    identifier: str | None = None
    language: str | None = None

    created: datetime | None = None
    modified: datetime | None = None

    @field_serializer("created", "modified", "last_printed")
    def serialize_datetime(self, v: datetime | None) -> str | None:
        if v is None:
            return None
        return v.isoformat()

    @field_validator("created", "modified", "last_printed", mode="before")
    @classmethod
    def validate_datetime(cls, v: datetime | str | None) -> datetime | None:
        if v is None:
            return None
        if isinstance(v, str):
            return datetime.fromisoformat(v)
        return v


class ExtendedProperties(BaseModel):
    template: str | None = None
    manager: str | None = None
    company: str | None = None

    application: str | None = None
    app_version: str | None = None
    doc_security: int | None = None
