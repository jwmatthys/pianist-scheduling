import base64
import asyncio
import json
import tempfile
import unittest
from pathlib import Path

from fastapi import FastAPI
import httpx
from sqlalchemy import create_engine
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from app.license_schemas import LicensePayload
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
        self.service = LicenseService(self.repository, public_key)
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
            "payload": signed_payload.model_dump(mode="json"),
            "signature": signature_value,
        }

    def file_contents(self, envelope=None):
        return json.dumps(envelope or self.signed_envelope()).encode("utf-8")

    def test_valid_license_file_activates_and_persists(self):
        with self.engine.connect() as connection:
            self.assertEqual(connection.exec_driver_sql("PRAGMA user_version").scalar_one(), 1)
            self.assertEqual(
                connection.exec_driver_sql(
                    "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
                ).scalars().all(),
                ["license_activations"],
            )

        activated = self.service.activate_license_file(self.file_contents())

        self.assertEqual(activated.status, "licensed")
        self.assertEqual(activated.payload, self.payload)
        self.assertIsNotNone(activated.activated_at)

        reopened_repository = LicenseRepository(self.engine)
        persisted = LicenseService(reopened_repository, self.private_key.public_key().public_bytes(
            encoding=serialization.Encoding.Raw,
            format=serialization.PublicFormat.Raw,
        )).get_license_status()
        self.assertEqual(persisted.status, "licensed")
        self.assertEqual(persisted.payload, self.payload)
        self.assertEqual(persisted.activated_at, activated.activated_at)

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
            **self.payload.model_dump(mode="json"),
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

    def test_clear_license_removes_persisted_activation(self):
        self.service.activate_license_file(self.file_contents())

        self.service.clear_license()

        self.assertEqual(self.service.get_license_status().status, "unlicensed")

    def test_api_status_and_activation_responses_are_typed(self):
        application = FastAPI()
        application.include_router(license_router.router)
        application.dependency_overrides[license_router.get_license_service] = lambda: self.service

        async def make_requests():
            transport = httpx.ASGITransport(app=application)
            async with httpx.AsyncClient(transport=transport, base_url="http://testserver") as client:
                status = await client.get("/api/license/status")
                self.assertEqual(status.status_code, 200)
                self.assertEqual(status.json(), {"status": "unlicensed", "payload": None, "activated_at": None, "error": None})

                key_response = await client.post(
                    "/api/license/activate-key",
                    json={"key": encode_license_key(self.payload, self.signed_envelope()["signature"])},
                )
                self.assertEqual(key_response.status_code, 200)
                self.assertEqual(key_response.json()["status"], "licensed")
                self.assertEqual(key_response.json()["payload"]["licensed_to"], "Synthetic University")
                self.assertIn("activated_at", key_response.json())

                file_response = await client.post(
                    "/api/license/activate-file",
                    files={"file": ("scheduling-maestro.license", self.file_contents(), "application/json")},
                )
                self.assertEqual(file_response.status_code, 200)
                self.assertEqual(file_response.json()["payload"]["license_type"], "PERPETUAL")

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
