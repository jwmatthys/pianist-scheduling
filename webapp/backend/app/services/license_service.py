"""Offline Ed25519 license validation and activation."""

import base64
import binascii
import json
import math
import time
from collections.abc import Mapping
from datetime import datetime, timezone
from threading import Lock
from typing import Any

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from pydantic import ValidationError

from ..license_keys import TRUSTED_PUBLIC_KEYS
from ..license_schemas import LicensePayload, LicenseStatus
from .license_repository import LicenseRepository, repository

PRODUCT_NAME = "Scheduling Maestro"
LICENSE_KEY_PREFIX = "SM1-"
MAX_LICENSE_FILE_BYTES = 64 * 1024
TRIAL_DURATION_MS = 30 * 60 * 1000
MAX_HEARTBEAT_GAP_MS = 10 * 1000


class LicenseServiceError(ValueError):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def canonical_payload_bytes(payload: LicensePayload) -> bytes:
    return json.dumps(
        payload.model_dump(mode="json", exclude_none=True),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")


def encode_license_key(payload: LicensePayload, signature: str) -> str:
    envelope = json.dumps(
        {"payload": payload.model_dump(mode="json", exclude_none=True), "signature": signature},
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
        trusted_public_keys: Mapping[str, bytes] | None = None,
    ):
        self._repository = license_repository
        self._trial_lock = Lock()
        self._last_trial_heartbeat: float | None = None
        if trusted_public_keys is None:
            try:
                trusted_public_keys = {
                    key_id: base64.b64decode(encoded_key, validate=True)
                    for key_id, encoded_key in TRUSTED_PUBLIC_KEYS.items()
                }
            except (ValueError, binascii.Error) as error:
                raise RuntimeError("The embedded Scheduling Maestro license trust store is invalid.") from error
        if (
            not trusted_public_keys
            or any(not key_id.strip() or key_id != key_id.strip() or len(key_id) > 100 or len(raw_key) != 32
                   for key_id, raw_key in trusted_public_keys.items())
            or len(set(trusted_public_keys.values())) != len(trusted_public_keys)
        ):
            raise RuntimeError("The Scheduling Maestro license trust store is invalid.")
        self._public_keys = {
            key_id: Ed25519PublicKey.from_public_bytes(raw_key)
            for key_id, raw_key in trusted_public_keys.items()
        }

    @property
    def max_license_file_bytes(self) -> int:
        return MAX_LICENSE_FILE_BYTES

    def initialize(self) -> None:
        self._repository.initialize()

    def get_license_status(self) -> LicenseStatus:
        record = self._repository.get_license_record()
        if record is None:
            return LicenseStatus(status="unlicensed", trial_seconds_remaining=self._trial_seconds_remaining())
        try:
            signing_key_id = self._verify(record.payload, record.signature)
        except LicenseServiceError:
            return LicenseStatus(
                status="invalid",
                error="The stored license could not be verified.",
                trial_seconds_remaining=self._trial_seconds_remaining(),
            )
        return LicenseStatus(
            status="licensed",
            payload=record.payload,
            activated_at=record.activated_at,
            signing_key_id=signing_key_id,
        )

    def _trial_seconds_remaining(self) -> int:
        return math.ceil(max(0, TRIAL_DURATION_MS - self._repository.get_trial_elapsed_ms()) / 1000)

    def trial_heartbeat(self) -> LicenseStatus:
        with self._trial_lock:
            now = time.monotonic()
            status = self.get_license_status()
            if status.status == "licensed":
                self._last_trial_heartbeat = None
                return status
            self._advance_trial(now)
            status.trial_seconds_remaining = self._trial_seconds_remaining()
            return status

    def _advance_trial(self, now: float) -> None:
        if self._last_trial_heartbeat is not None:
            elapsed_ms = min(MAX_HEARTBEAT_GAP_MS, max(0, round((now - self._last_trial_heartbeat) * 1000)))
            self._repository.add_trial_elapsed_ms(elapsed_ms)
        self._last_trial_heartbeat = now

    def activate_license_file(self, contents: bytes) -> LicenseStatus:
        return self._activate(self.parse_license_file(contents))

    def parse_license_file(self, contents: bytes) -> tuple[LicensePayload, str]:
        if len(contents) > MAX_LICENSE_FILE_BYTES:
            raise LicenseServiceError("LICENSE_FILE_INVALID", "This license file is invalid.")
        if contents.lstrip().startswith(LICENSE_KEY_PREFIX.encode("ascii")):
            try:
                key = contents.decode("utf-8")
            except UnicodeDecodeError as error:
                raise LicenseServiceError("LICENSE_KEY_INVALID", "License key could not be read.") from error
            return self.parse_license_key(key)
        return self._parse_envelope(contents, "LICENSE_FILE_INVALID", "This license file is invalid.")

    def activate_license_key(self, key: str) -> LicenseStatus:
        return self._activate(self.parse_license_key(key))

    def parse_license_key(self, key: str) -> tuple[LicensePayload, str]:
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
        return self._parse_envelope(contents, "LICENSE_KEY_INVALID", "License key could not be read.")

    def verify_license(self, payload: LicensePayload, signature: str) -> str:
        return self._verify(payload, signature)

    def clear_license(self) -> None:
        with self._trial_lock:
            self._repository.clear_license()
            self._last_trial_heartbeat = None

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
        signing_key_id = self._verify(payload, signature)
        activated_at = datetime.now(timezone.utc).replace(microsecond=0)
        with self._trial_lock:
            if self._last_trial_heartbeat is not None:
                self._advance_trial(time.monotonic())
            record = self._repository.save_license_record(payload, signature, activated_at)
            self._last_trial_heartbeat = None
        return LicenseStatus(
            status="licensed",
            payload=record.payload,
            activated_at=record.activated_at,
            signing_key_id=signing_key_id,
        )

    def _verify(self, payload: LicensePayload, signature: str) -> str:
        try:
            signature_bytes = base64.b64decode(signature, validate=True)
        except (ValueError, binascii.Error) as error:
            raise LicenseServiceError("INVALID_SIGNATURE", "Unable to verify this license.") from error
        canonical_payload = canonical_payload_bytes(payload)
        signing_key_id = None
        for key_id, public_key in self._public_keys.items():
            try:
                public_key.verify(signature_bytes, canonical_payload)
            except InvalidSignature:
                continue
            signing_key_id = key_id
            break
        if signing_key_id is None:
            raise LicenseServiceError("INVALID_SIGNATURE", "Unable to verify this license.")
        if payload.product != PRODUCT_NAME:
            raise LicenseServiceError(
                "WRONG_PRODUCT",
                "This license was not issued for Scheduling Maestro.",
            )
        return signing_key_id


service = LicenseService()
