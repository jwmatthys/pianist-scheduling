import base64
import importlib.util
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from sqlalchemy import create_engine

from app.license_schemas import LicensePayload
from app.services.license_repository import LicenseRepository
from app.services.license_service import LicenseService, canonical_payload_bytes, encode_license_key

REPOSITORY_ROOT = Path(__file__).resolve().parents[3]
spec = importlib.util.spec_from_file_location("inspect_license", REPOSITORY_ROOT / "tools" / "inspect_license.py")
inspector = importlib.util.module_from_spec(spec)
spec.loader.exec_module(inspector)


class LicenseInspectorTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary_directory.name)
        self.engine = create_engine(f"sqlite:///{self.directory / 'application.sqlite3'}")
        self.repository = LicenseRepository(self.engine)
        self.repository.initialize()
        self.private_key = Ed25519PrivateKey.generate()
        public_key = self.private_key.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
        self.service = LicenseService(self.repository, {"v1": public_key})
        self.service_patch = patch.object(inspector, "service", self.service)
        self.keys_patch = patch.object(inspector, "TRUSTED_PUBLIC_KEYS", {"v1": base64.b64encode(public_key).decode()})
        self.service_patch.start()
        self.keys_patch.start()
        self.payload = LicensePayload.model_validate({
            "product": "Scheduling Maestro",
            "licensed_to": "Synthetic University",
            "license_type": "PERPETUAL",
            "issued": "2026-11-23",
            "license_version": 1,
            "key_id": "v1",
        })

    def tearDown(self):
        self.keys_patch.stop()
        self.service_patch.stop()
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def envelope(self, payload=None, private_key=None):
        payload = payload or self.payload
        signature = base64.b64encode((private_key or self.private_key).sign(canonical_payload_bytes(payload))).decode()
        return payload, signature

    def inspect(self, args):
        output = StringIO()
        errors = StringIO()
        with redirect_stdout(output), redirect_stderr(errors):
            code = inspector.main(args)
        return code, output.getvalue(), errors.getvalue()

    def test_valid_file_and_key_show_verified_metadata_without_activation(self):
        payload, signature = self.envelope()
        path = self.directory / "synthetic.license"
        path.write_text(json.dumps({"payload": payload.model_dump(mode="json"), "signature": signature}), encoding="utf-8")
        key = encode_license_key(payload, signature)
        key_path = self.directory / "synthetic-key.txt"
        key_path.write_text(key, encoding="utf-8")

        for args in (["--file", str(path)], ["--file", str(key_path)], ["--key", key]):
            with self.subTest(method=args[0]):
                code, output, errors = self.inspect(args)
                self.assertEqual(code, 0)
                self.assertEqual(errors, "")
                for line in (
                    "Product: Scheduling Maestro",
                    "Licensed To: Synthetic University",
                    "License Type: PERPETUAL",
                    "Issued: 2026-11-23",
                    "License Version: 1",
                    "Signing Key ID: v1",
                    "Verification Result: VALID",
                ):
                    self.assertIn(line, output)
        self.assertIsNone(self.repository.get_license_record())
        self.assertEqual(self.repository.get_trial_elapsed_ms(), 0)

    def test_invalid_signature_and_unknown_signing_key_are_unverified(self):
        payload, _ = self.envelope()
        invalid = encode_license_key(payload, base64.b64encode(b"\x00" * 64).decode())
        code, output, _ = self.inspect(["--key", invalid])
        self.assertEqual(code, 1)
        self.assertIn("Verification Result: INVALID SIGNATURE", output)
        self.assertIn("Signing Key ID: v1 (unverified)", output)

        unknown_payload = payload.model_copy(update={"key_id": "v2"})
        unknown_payload, signature = self.envelope(unknown_payload, Ed25519PrivateKey.generate())
        code, output, _ = self.inspect(["--key", encode_license_key(unknown_payload, signature)])
        self.assertEqual(code, 1)
        self.assertIn("Verification Result: UNKNOWN SIGNING KEY", output)
        self.assertIn("may also be corrupted", output)

    def test_claimed_unknown_id_does_not_override_a_valid_signature(self):
        payload = self.payload.model_copy(update={"key_id": "v2"})
        payload, signature = self.envelope(payload)
        code, output, _ = self.inspect(["--key", encode_license_key(payload, signature)])
        self.assertEqual(code, 0)
        self.assertIn("Verification Result: VALID", output)
        self.assertIn("Signing Key ID: v1", output)
        self.assertIn("Declared Key ID: v2 (differs from verified signer)", output)

    def test_legacy_license_and_wrong_product_use_application_verification(self):
        legacy = self.payload.model_copy(update={"key_id": None})
        payload, signature = self.envelope(legacy)
        code, output, _ = self.inspect(["--key", encode_license_key(payload, signature)])
        self.assertEqual(code, 0)
        self.assertIn("Signing Key ID: v1", output)

        wrong_product = self.payload.model_copy(update={"product": "Another Product"})
        payload, signature = self.envelope(wrong_product)
        code, output, _ = self.inspect(["--key", encode_license_key(payload, signature)])
        self.assertEqual(code, 1)
        self.assertIn("Verification Result: WRONG PRODUCT", output)

    def test_malformed_file_key_and_unreadable_file(self):
        path = self.directory / "malformed.license"
        path.write_bytes(b'{"payload":')
        for args in (["--file", str(path)], ["--key", "SM1-invalid"], ["--key", "not-a-key"]):
            with self.subTest(method=args[0]):
                code, output, _ = self.inspect(args)
                self.assertEqual(code, 1)
                self.assertIn("Verification Result: CORRUPTED LICENSE", output)

        code, _, errors = self.inspect(["--file", str(self.directory / "missing.license")])
        self.assertEqual(code, 2)
        self.assertIn("Cannot read license file", errors)

    def test_unprintable_metadata_and_invalid_unicode_are_handled(self):
        payload = self.payload.model_copy(update={"licensed_to": "Synthetic\nUniversity"})
        invalid = encode_license_key(payload, base64.b64encode(b"\x00" * 64).decode())
        code, output, _ = self.inspect(["--key", invalid])
        self.assertEqual(code, 1)
        self.assertIn(r"Licensed To: Synthetic\u000aUniversity", output)
        self.assertIn("Verification Result: INVALID SIGNATURE", output)

        path = self.directory / "invalid-unicode.license"
        document = {
            "payload": {**self.payload.model_dump(mode="json"), "licensed_to": "\ud800"},
            "signature": base64.b64encode(b"\x00" * 64).decode(),
        }
        path.write_text(json.dumps(document), encoding="utf-8")
        code, output, _ = self.inspect(["--file", str(path)])
        self.assertEqual(code, 1)
        self.assertIn("Verification Result: CORRUPTED LICENSE", output)


if __name__ == "__main__":
    unittest.main()
