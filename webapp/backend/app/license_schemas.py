"""Typed license payload and API response models."""

from datetime import date, datetime
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


class LicensePayload(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    product: str = Field(min_length=1, max_length=100)
    licensed_to: str = Field(min_length=1, max_length=300)
    license_type: str = Field(min_length=1, max_length=50)
    issued: date
    license_version: Literal[1]
    key_id: str | None = Field(default=None, min_length=1, max_length=100)

    @field_validator("product", "licensed_to", "license_type")
    @classmethod
    def values_must_be_trimmed(cls, value: str) -> str:
        if not value.strip() or value != value.strip():
            raise ValueError("Value must be non-empty and have no surrounding whitespace.")
        return value

    @field_validator("key_id")
    @classmethod
    def key_id_must_be_trimmed(cls, value: str | None) -> str:
        if value is None or not value.strip() or value != value.strip():
            raise ValueError("Key identifier must be non-empty and have no surrounding whitespace.")
        return value

    @field_validator("issued", mode="before")
    @classmethod
    def parse_iso_issue_date(cls, value: object) -> date:
        if not isinstance(value, str):
            raise ValueError("Issue date must be an ISO 8601 date.")
        try:
            parsed = date.fromisoformat(value)
        except ValueError as error:
            raise ValueError("Issue date must be an ISO 8601 date.") from error
        if parsed.isoformat() != value:
            raise ValueError("Issue date must be an ISO 8601 date.")
        return parsed


class LicenseRecord(BaseModel):
    payload: LicensePayload
    signature: str
    activated_at: datetime


class LicenseStatus(BaseModel):
    status: Literal["unlicensed", "licensed", "invalid"]
    payload: LicensePayload | None = None
    activated_at: datetime | None = None
    error: str | None = None
    trial_seconds_remaining: int | None = None
    signing_key_id: str | None = None


class LicenseKeyRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    key: str = Field(min_length=1, max_length=8192)
