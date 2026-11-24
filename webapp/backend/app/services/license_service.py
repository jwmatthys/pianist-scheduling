"""Offline Ed25519 license validation and activation."""

import base64
import binascii
import json
from datetime import datetime, timezone
from typing import Any

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from pydantic import ValidationError

from ..license_schemas import LicensePayload, LicenseStatus
from .license_repository import LicenseRepository, repository

PRODUCT_NAME = "Scheduling Maestro"
PUBLIC_KEY_B64 = "v3H85BcW068XJ0Q4ExFjoOxkF2GSB5W2K+wSnGmlKGs="
LICENSE_KEY_PREFIX = "SM1-"
MAX_LICENSE_FILE_BYTES = 64 * 1024


class LicenseServiceError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def canonical_payload_bytes(payload: LicensePayload) -> bytes:
    return json.dumps(
        payload.model_dump(mode="json"),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def encode_license_key(payload: LicensePayload, signature: str) -> str:
    envelope = json.dumps(
        {"payload": payload.model_dump(mode="json"), "signature": signature},
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    encoded = base64.urlsafe_b64encode(envelope).decode("ascii").rstrip("=")
    return f"{LICENSE_KEY_PREFIX}{encoded}"


def _reject_duplicate_json_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("Duplicate JSON fields are not allowed.")
        result[key] = value
    return result


class LicenseService:
    def __init__(
        self,
        license_repository: LicenseRepository = repository,
        public_key_bytes: bytes | None = None,
    ):
        self._repository = license_repository
        raw_key = public_key_bytes if public_key_bytes is not None else base64.b64decode(PUBLIC_KEY_B64, validate=True)
        if len(raw_key) != 32:
            raise RuntimeError("The embedded Scheduling Maestro license public key is invalid.")
        self._public_key = Ed25519PublicKey.from_public_bytes(raw_key)

    @property
    def max_license_file_bytes(self) -> int:
        return MAX_LICENSE_FILE_BYTES

    def initialize(self) -> None:
        self._repository.initialize()

    def get_license_status(self) -> LicenseStatus:
        record = self._repository.get_license_record()
        if record is None:
            return LicenseStatus(status="unlicensed")
        try:
            self._verify(record.payload, record.signature)
        except LicenseServiceError:
            return LicenseStatus(
                status="invalid",
                error="The stored license could not be verified.",
            )
        return LicenseStatus(
            status="licensed",
            payload=record.payload,
            activated_at=record.activated_at,
        )

    def activate_license_file(self, contents: bytes) -> LicenseStatus:
        if len(contents) > MAX_LICENSE_FILE_BYTES:
            raise LicenseServiceError("LICENSE_FILE_INVALID", "This license file is invalid.")
        envelope = self._parse_envelope(contents, "LICENSE_FILE_INVALID", "This license file is invalid.")
        return self._activate(envelope)

    def activate_license_key(self, key: str) -> LicenseStatus:
        normalized_key = key.strip()
        if not normalized_key.startswith(LICENSE_KEY_PREFIX):
            raise LicenseServiceError("LICENSE_KEY_INVALID", "License key could not be read.")
        encoded = normalized_key[len(LICENSE_KEY_PREFIX):]
        if not encoded or len(normalized_key) > 8192:
            raise LicenseServiceError("LICENSE_KEY_INVALID", "License key could not be read.")
        try:
            padded = encoded + "=" * (-len(encoded) % 4)
            contents = base64.b64decode(padded, altchars=b"-_", validate=True)
        except (binascii.Error, ValueError) as error:
            raise LicenseServiceError("LICENSE_KEY_INVALID", "License key could not be read.") from error
        envelope = self._parse_envelope(contents, "LICENSE_KEY_INVALID", "License key could not be read.")
        return self._activate(envelope)

    def clear_license(self) -> None:
        self._repository.clear_license()

    def _parse_envelope(self, contents: bytes, code: str, message: str) -> tuple[LicensePayload, str]:
        try:
            document = json.loads(
                contents.decode("utf-8"),
                object_pairs_hook=_reject_duplicate_json_keys,
            )
            if not isinstance(document, dict) or set(document) != {"payload", "signature"}:
                raise ValueError("Invalid license envelope.")
            if not isinstance(document["payload"], dict) or not isinstance(document["signature"], str):
                raise ValueError("Invalid license envelope.")
            payload = LicensePayload.model_validate(document["payload"])
            signature_bytes = base64.b64decode(document["signature"], validate=True)
            if len(signature_bytes) != 64:
                raise ValueError("Invalid signature length.")
            return payload, document["signature"]
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError, binascii.Error, ValidationError) as error:
            raise LicenseServiceError(code, message) from error

    def _activate(self, envelope: tuple[LicensePayload, str]) -> LicenseStatus:
        payload, signature = envelope
        self._verify(payload, signature)
        activated_at = datetime.now(timezone.utc).replace(microsecond=0)
        record = self._repository.save_license_record(payload, signature, activated_at)
        return LicenseStatus(
            status="licensed",
            payload=record.payload,
            activated_at=record.activated_at,
        )

    def _verify(self, payload: LicensePayload, signature: str) -> None:
        try:
            signature_bytes = base64.b64decode(signature, validate=True)
            self._public_key.verify(signature_bytes, canonical_payload_bytes(payload))
        except (InvalidSignature, ValueError, binascii.Error) as error:
            raise LicenseServiceError("INVALID_SIGNATURE", "Unable to verify this license.") from error
        if payload.product != PRODUCT_NAME:
            raise LicenseServiceError(
                "WRONG_PRODUCT",
                "This license was not issued for Scheduling Maestro.",
            )


service = LicenseService()
