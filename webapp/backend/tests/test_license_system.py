import base64
import asyncio
import json
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI
import httpx
from sqlalchemy import create_engine
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from app.license_schemas import LicensePayload
from app.database import APPLICATION_MIGRATIONS, migrate_application_database
from app.routers import license as license_router
from app.services.license_repository import LicenseRepository
from app.services.license_service import (
    LicenseService,
    LicenseServiceError,
    canonical_payload_bytes,
    encode_license_key,
)


class LicenseSystemTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.database_path = Path(self.temporary_directory.name) / "application.sqlite3"
        self.engine = create_engine(f"sqlite:///{self.database_path}")
        self.repository = LicenseRepository(self.engine)
        self.repository.initialize()
        self.private_key = Ed25519PrivateKey.generate()
        public_key = self.private_key.public_key().public_bytes(
            encoding=serialization.Encoding.Raw,
            format=serialization.PublicFormat.Raw,
        )
        self.service = LicenseService(self.repository, {"v1": public_key})
        self.payload = LicensePayload.model_validate({
            "product": "Scheduling Maestro",
            "licensed_to": "Synthetic University",
            "license_type": "PERPETUAL",
            "issued": "2026-11-23",
            "license_version": 1,
        })

    def tearDown(self):
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def signed_envelope(self, payload=None, signature=None):
        signed_payload = payload or self.payload
        signature_value = signature or base64.b64encode(
            self.private_key.sign(canonical_payload_bytes(signed_payload))
        ).decode("ascii")
        return {
            "payload": signed_payload.model_dump(mode="json", exclude_none=True),
            "signature": signature_value,
        }

    def file_contents(self, envelope=None):
        return json.dumps(envelope or self.signed_envelope()).encode("utf-8")

    def test_valid_license_file_activates_and_persists(self):
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), 2)
            self.assertEqual(
                connection.exec_driver_sql(
                    "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
                ).scalars().all(),
                ["license_activations", "trial_usage"],
            )

        activated = self.service.activate_license_file(self.file_contents())

        self.assertEqual(activated.status, "licensed")
        self.assertEqual(activated.payload, self.payload)
        self.assertIsNotNone(activated.activated_at)

        reopened_repository = LicenseRepository(self.engine)
        persisted = LicenseService(reopened_repository, {"v1": self.private_key.public_key().public_bytes(
            encoding=serialization.Encoding.Raw,
            format=serialization.PublicFormat.Raw,
        )}).get_license_status()
        self.assertEqual(persisted.status, "licensed")
        self.assertEqual(persisted.payload, self.payload)
        self.assertEqual(persisted.activated_at, activated.activated_at)
        self.assertEqual(persisted.signing_key_id, "v1")

    def test_invalid_signature_is_rejected(self):
        invalid_signature = base64.b64encode(b"\x00" * 64).decode("ascii")

        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(self.file_contents(self.signed_envelope(signature=invalid_signature)))

        self.assertEqual(raised.exception.code, "INVALID_SIGNATURE")
        self.assertEqual(self.service.get_license_status().status, "unlicensed")

    def test_malformed_license_file_is_rejected(self):
        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(b'{"payload":')

        self.assertEqual(raised.exception.code, "LICENSE_FILE_INVALID")

    def test_wrong_product_is_rejected_after_signature_verification(self):
        wrong_product = LicensePayload.model_validate({
            **self.payload.model_dump(mode="json", exclude_none=True),
            "product": "Another Product",
        })

        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(self.file_contents(self.signed_envelope(payload=wrong_product)))

        self.assertEqual(raised.exception.code, "WRONG_PRODUCT")

    def test_valid_license_key_uses_same_signed_payload(self):
        envelope = self.signed_envelope()
        key = encode_license_key(self.payload, envelope["signature"])

        activated = self.service.activate_license_key(key)

        self.assertEqual(activated.status, "licensed")
        self.assertEqual(activated.payload, self.payload)

    def test_api_replaces_existing_license_and_preserves_it_after_failed_activation(self):
        application = FastAPI()
        application.include_router(license_router.router)
        application.dependency_overrides[license_router.get_license_service] = lambda: self.service
        replacement = LicensePayload.model_validate({
            **self.payload.model_dump(mode="json", exclude_none=True),
            "licensed_to": "Replacement Conservatory",
            "issued": "2026-11-24",
        })
        replacement_envelope = self.signed_envelope(payload=replacement)
        replacement_key = encode_license_key(replacement, replacement_envelope["signature"])
        original_key = encode_license_key(self.payload, self.signed_envelope()["signature"])

        async def make_requests():
            transport = httpx.ASGITransport(app=application)
            async with httpx.AsyncClient(transport=transport, base_url="http://testserver") as client:
                self.assertEqual((await client.post("/api/license/activate-key", json={"key": original_key})).status_code, 200)
                response = await client.post(
                    "/api/license/activate-file",
                    files={"file": ("replacement.license", self.file_contents(replacement_envelope), "application/json")},
                )
                self.assertEqual(response.json()["payload"]["licensed_to"], "Replacement Conservatory")
                self.assertEqual((await client.get("/api/license/status")).json()["payload"]["issued"], "2026-11-24")
                self.assertEqual((await client.post("/api/license/activate-key", json={"key": "invalid"})).status_code, 400)
                self.assertEqual((await client.get("/api/license/status")).json()["payload"]["licensed_to"], "Replacement Conservatory")
                response = await client.post("/api/license/activate-key", json={"key": original_key})
                self.assertEqual(response.json()["payload"]["licensed_to"], "Synthetic University")
                response = await client.post("/api/license/activate-key", json={"key": replacement_key})
                self.assertEqual(response.json()["payload"]["licensed_to"], "Replacement Conservatory")

        asyncio.run(make_requests())
        reopened = LicenseService(LicenseRepository(self.engine), {
            "v1": self.private_key.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw),
        }).get_license_status()
        self.assertEqual(reopened.payload, replacement)

    def test_existing_unidentified_license_verifies_without_reactivation(self):
        legacy_payload = {
            "product": "Scheduling Maestro",
            "licensed_to": "Synthetic University",
            "license_type": "PERPETUAL",
            "issued": "2026-11-23",
            "license_version": 1,
        }
        canonical = json.dumps(legacy_payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
        signature = base64.b64encode(self.private_key.sign(canonical)).decode("ascii")
        original_json = json.dumps(legacy_payload)
        with self.engine.begin() as connection:
            connection.exec_driver_sql(
                "INSERT INTO license_activations VALUES (1, ?, ?, ?)",
                (original_json, signature, "2026-11-23T12:00:00Z"),
            )
        self.repository.add_trial_elapsed_ms(1800000)

        status = self.service.get_license_status()
        self.assertEqual(status.status, "licensed")
        self.assertIsNone(status.payload.key_id)
        self.assertEqual(status.signing_key_id, "v1")
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql(
                "SELECT payload_json FROM license_activations WHERE singleton = 1"
            ).scalar_one(), original_json)

        legacy_envelope = json.dumps({"payload": legacy_payload, "signature": signature}).encode("utf-8")
        legacy_key = "SM1-" + base64.urlsafe_b64encode(legacy_envelope).decode("ascii").rstrip("=")
        self.assertEqual(self.service.activate_license_file(legacy_envelope).signing_key_id, "v1")
        self.assertEqual(self.service.activate_license_key(legacy_key).signing_key_id, "v1")
        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(json.dumps({
                "payload": {**legacy_payload, "key_id": None}, "signature": signature,
            }).encode("utf-8"))
        self.assertEqual(raised.exception.code, "LICENSE_FILE_INVALID")

    def test_second_trusted_key_and_actual_signer_metadata(self):
        unrelated_key = Ed25519PrivateKey.generate().public_key().public_bytes(
            serialization.Encoding.Raw, serialization.PublicFormat.Raw,
        )
        actual_key = self.private_key.public_key().public_bytes(
            serialization.Encoding.Raw, serialization.PublicFormat.Raw,
        )
        service = LicenseService(self.repository, {"v1": unrelated_key, "v2": actual_key})
        payload = LicensePayload.model_validate({**self.payload.model_dump(mode="json", exclude_none=True), "key_id": "v1"})
        envelope = self.signed_envelope(payload=payload)

        result = service.activate_license_file(self.file_contents(envelope))
        self.assertEqual(result.status, "licensed")
        self.assertEqual(result.payload.key_id, "v1")
        self.assertEqual(result.signing_key_id, "v2")
        self.assertEqual(service.get_license_status().signing_key_id, "v2")

        envelope["payload"]["key_id"] = "v2"
        with self.assertRaises(LicenseServiceError) as raised:
            service.activate_license_file(self.file_contents(envelope))
        self.assertEqual(raised.exception.code, "INVALID_SIGNATURE")
        self.assertEqual(service.get_license_status().signing_key_id, "v2")

    def test_no_trusted_key_matches_and_empty_trust_store_fails(self):
        keys = {
            key_id: Ed25519PrivateKey.generate().public_key().public_bytes(
                serialization.Encoding.Raw, serialization.PublicFormat.Raw,
            )
            for key_id in ("v1", "v2")
        }
        service = LicenseService(self.repository, keys)
        with self.assertRaises(LicenseServiceError) as raised:
            service.activate_license_file(self.file_contents())
        self.assertEqual(raised.exception.code, "INVALID_SIGNATURE")
        with self.assertRaisesRegex(RuntimeError, "trust store"):
            LicenseService(self.repository, {})
        with self.assertRaisesRegex(RuntimeError, "trust store"):
            LicenseService(self.repository, {"v1": keys["v1"], "v2": keys["v1"]})

    def test_import_license_key_as_file_verifies_signature(self):
        envelope = self.signed_envelope()
        key = encode_license_key(self.payload, envelope["signature"])

        activated = self.service.activate_license_file(f"  {key}\n".encode())

        self.assertEqual(activated.status, "licensed")
        self.assertEqual(activated.payload, self.payload)

        tampered = encode_license_key(self.payload, base64.b64encode(b"\x00" * 64).decode())
        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(tampered.encode())
        self.assertEqual(raised.exception.code, "INVALID_SIGNATURE")
        self.assertEqual(self.service.get_license_status().payload, self.payload)

        with self.assertRaises(LicenseServiceError) as raised:
            self.service.activate_license_file(b"SM1-\xff")
        self.assertEqual(raised.exception.code, "LICENSE_KEY_INVALID")

    def test_clear_license_removes_persisted_activation(self):
        self.service.activate_license_file(self.file_contents())

        self.service.clear_license()

        self.assertEqual(self.service.get_license_status().status, "unlicensed")

    def test_trial_counts_usage_persists_and_stops_at_zero(self):
        with patch("app.services.license_service.time.monotonic", side_effect=[100, 105, 110, 115]):
            self.assertEqual(self.service.trial_heartbeat().trial_seconds_remaining, 1800)
            self.assertEqual(self.service.trial_heartbeat().trial_seconds_remaining, 1795)
            self.assertEqual(self.service.trial_heartbeat().trial_seconds_remaining, 1790)
            self.service.activate_license_file(self.file_contents())
        self.assertEqual(self.repository.get_trial_elapsed_ms(), 15000)
        self.assertEqual(self.service.trial_heartbeat().status, "licensed")
        self.assertEqual(self.repository.get_trial_elapsed_ms(), 15000)
        self.service.clear_license()
        self.assertEqual(LicenseService(self.repository).get_license_status().trial_seconds_remaining, 1785)
        self.repository.add_trial_elapsed_ms(1800000)
        self.assertEqual(self.service.get_license_status().trial_seconds_remaining, 0)
        self.assertEqual(self.service.activate_license_file(self.file_contents()).status, "licensed")

    def test_expired_trial_blocks_api_but_allows_activation_and_health(self):
        from app.main import app

        app.dependency_overrides[license_router.get_license_service] = lambda: self.service

        async def make_requests():
            transport = httpx.ASGITransport(app=app)
            async with httpx.AsyncClient(transport=transport, base_url="http://testserver") as client:
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 404)
                self.repository.add_trial_elapsed_ms(1800000)
                response = await client.get("/api/session", headers={"Origin": "http://localhost:5173"})
                self.assertEqual(response.status_code, 423)
                self.assertEqual(response.json()["detail"]["code"], "TRIAL_EXPIRED")
                self.assertEqual(response.headers["access-control-allow-origin"], "http://localhost:5173")
                self.assertEqual((await client.post("/api/not-a-route")).status_code, 423)
                self.assertEqual((await client.get("/api/health")).status_code, 200)
                self.assertEqual((await client.get("/api/license/status")).json()["trial_seconds_remaining"], 0)
                self.assertEqual((await client.post("/api/license/trial-heartbeat")).status_code, 200)
                self.assertEqual((await client.post("/api/license/activate-key", json={"key": "invalid"})).status_code, 400)
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 423)
                self.repository.save_license_record(
                    self.payload, base64.b64encode(b"\x00" * 64).decode(), datetime.now(timezone.utc),
                )
                self.assertEqual((await client.get("/api/license/status")).json()["status"], "invalid")
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 423)

                key = encode_license_key(self.payload, self.signed_envelope()["signature"])
                self.assertEqual((await client.post("/api/license/activate-key", json={"key": key})).status_code, 200)
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 404)
                self.service.clear_license()
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 423)
                response = await client.post(
                    "/api/license/activate-file",
                    files={"file": ("scheduling-maestro.license", self.file_contents(), "application/json")},
                )
                self.assertEqual(response.status_code, 200)
                self.assertEqual((await client.get("/api/not-a-route")).status_code, 404)

        try:
            with patch("app.main.license_service", self.service):
                asyncio.run(make_requests())
        finally:
            app.dependency_overrides.clear()

    def test_trial_does_not_count_an_unbounded_gap_between_heartbeats(self):
        with patch("app.services.license_service.time.monotonic", side_effect=[100, 200]):
            self.service.trial_heartbeat()
            self.assertEqual(self.service.trial_heartbeat().trial_seconds_remaining, 1790)

    def test_trial_reaches_zero_at_the_full_duration(self):
        self.repository.add_trial_elapsed_ms(1799999)
        self.assertEqual(self.service.get_license_status().trial_seconds_remaining, 1)
        self.repository.add_trial_elapsed_ms(1)
        self.assertEqual(self.service.get_license_status().trial_seconds_remaining, 0)

    def test_existing_application_database_migrates_without_losing_license(self):
        old_engine = create_engine(f"sqlite:///{Path(self.temporary_directory.name) / 'old.sqlite3'}")
        try:
            migrate_application_database(old_engine, APPLICATION_MIGRATIONS[:1])
            with old_engine.begin() as connection:
                connection.exec_driver_sql(
                    "INSERT INTO license_activations VALUES (1, ?, ?, ?)",
                    (json.dumps(self.payload.model_dump(mode="json", exclude_none=True)), self.signed_envelope()["signature"],
                     "2026-11-23T12:00:00Z"),
                )
            repository = LicenseRepository(old_engine)
            repository.initialize()
            self.assertEqual(repository.get_license_record().payload, self.payload)
            self.assertEqual(repository.get_trial_elapsed_ms(), 0)
        finally:
            old_engine.dispose()

    def test_api_status_and_activation_responses_are_typed(self):
        application = FastAPI()
        application.include_router(license_router.router)
        application.dependency_overrides[license_router.get_license_service] = lambda: self.service

        async def make_requests():
            transport = httpx.ASGITransport(app=application)
            async with httpx.AsyncClient(transport=transport, base_url="http://testserver") as client:
                status = await client.get("/api/license/status")
                self.assertEqual(status.status_code, 200)
                self.assertEqual(status.json(), {
                    "status": "unlicensed", "payload": None, "activated_at": None, "error": None,
                    "trial_seconds_remaining": 1800, "signing_key_id": None,
                })
                heartbeat = await client.post("/api/license/trial-heartbeat")
                self.assertEqual(heartbeat.json()["trial_seconds_remaining"], 1800)

                key_response = await client.post(
                    "/api/license/activate-key",
                    json={"key": encode_license_key(self.payload, self.signed_envelope()["signature"])},
                )
                self.assertEqual(key_response.status_code, 200)
                self.assertEqual(key_response.json()["status"], "licensed")
                self.assertEqual(key_response.json()["payload"]["licensed_to"], "Synthetic University")
                self.assertEqual(key_response.json()["signing_key_id"], "v1")
                self.assertIn("activated_at", key_response.json())

                file_response = await client.post(
                    "/api/license/activate-file",
                    files={"file": ("scheduling-maestro.license", self.file_contents(), "application/json")},
                )
                self.assertEqual(file_response.status_code, 200)
                self.assertEqual(file_response.json()["payload"]["license_type"], "PERPETUAL")

                key_file_response = await client.post(
                    "/api/license/activate-file",
                    files={"file": (
                        "renamed.license",
                        encode_license_key(self.payload, self.signed_envelope()["signature"]).encode(),
                        "text/plain",
                    )},
                )
                self.assertEqual(key_file_response.status_code, 200)
                self.assertEqual(key_file_response.json()["payload"]["licensed_to"], "Synthetic University")

                invalid_file = await client.post(
                    "/api/license/activate-file",
                    files={"file": ("broken.license", b"invalid json", "application/json")},
                )
                self.assertEqual(invalid_file.status_code, 400)
                self.assertEqual(invalid_file.json()["detail"]["code"], "LICENSE_FILE_INVALID")

        asyncio.run(make_requests())

        application.dependency_overrides.clear()


if __name__ == "__main__":
    unittest.main()
