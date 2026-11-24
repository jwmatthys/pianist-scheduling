import base64
import importlib.util
import json
import os
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from datetime import date
from io import StringIO
from pathlib import Path
from unittest.mock import patch

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from cryptography.hazmat.primitives.asymmetric.rsa import generate_private_key
from sqlalchemy import create_engine

from app.services.license_repository import LicenseRepository
from app.services.license_service import LicenseService

REPOSITORY_ROOT = Path(__file__).resolve().parents[3]
spec = importlib.util.spec_from_file_location("generate_license", REPOSITORY_ROOT / "tools" / "generate_license.py")
generator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(generator)


class LicenseGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary_directory.name)
        private_key = Ed25519PrivateKey.generate()
        self.key_path = self.directory / "synthetic-private.pem"
        self.key_path.write_bytes(private_key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption(),
        ))
        self.public_key = private_key.public_key().public_bytes(
            serialization.Encoding.Raw,
            serialization.PublicFormat.Raw,
        )
        self.engine = create_engine(f"sqlite:///{self.directory / 'license.sqlite3'}")
        self.repository = LicenseRepository(self.engine)
        self.repository.initialize()
        self.service = LicenseService(self.repository, {"v1": self.public_key})
        self.public_patch = patch.object(
            generator, "TRUSTED_PUBLIC_KEYS", {"v1": base64.b64encode(self.public_key).decode("ascii")}
        )
        self.public_patch.start()

    def tearDown(self):
        self.public_patch.stop()
        self.engine.dispose()
        self.temporary_directory.cleanup()

    def test_cli_generates_equivalent_activatable_outputs(self):
        output = StringIO()
        with redirect_stdout(output):
            result = generator.main([
                "--licensed-to", "Synthetic University",
                "--private-key", str(self.key_path),
                "--output-dir", str(self.directory),
            ])
        self.assertEqual(result, 0)
        self.assertIn("Signing key ID: v1", output.getvalue())
        self.assertIn(base64.b64encode(self.public_key).decode("ascii"), output.getvalue())
        license_file = self.directory / "scheduling-maestro.license"
        key_file = self.directory / "license-key.txt"
        contents = license_file.read_bytes()
        envelope = json.loads(contents)
        self.assertEqual(envelope["payload"], {
            "product": "Scheduling Maestro",
            "licensed_to": "Synthetic University",
            "license_type": "PERPETUAL",
            "issued": date.today().isoformat(),
            "license_version": 1,
            "key_id": "v1",
        })
        key = key_file.read_text(encoding="utf-8")
        self.assertTrue(key.startswith("SM1-"))
        self.assertEqual(key, key.strip())
        file_status = self.service.activate_license_file(contents)
        self.service.clear_license()
        key_status = self.service.activate_license_key(key)
        self.service.clear_license()
        key_file_status = self.service.activate_license_file(key_file.read_bytes())
        self.assertEqual(file_status.status, "licensed")
        self.assertEqual(key_status.status, "licensed")
        self.assertEqual(file_status.payload, key_status.payload)
        self.assertEqual(file_status.payload, key_file_status.payload)
        self.assertEqual(file_status.signing_key_id, "v1")
        self.assertEqual(self.repository.get_license_record().signature, envelope["signature"])

    def test_default_output_is_current_directory(self):
        previous_directory = Path.cwd()
        try:
            os.chdir(self.directory)
            with redirect_stdout(StringIO()):
                self.assertEqual(generator.main(["--licensed-to", "Synthetic Customer", "--private-key", str(self.key_path)]), 0)
        finally:
            os.chdir(previous_directory)
        self.assertTrue((self.directory / "scheduling-maestro.license").is_file())
        self.assertTrue((self.directory / "license-key.txt").is_file())

    def test_selects_configured_rotated_key(self):
        private_key = Ed25519PrivateKey.generate()
        key_path = self.directory / "synthetic-v2.pem"
        key_path.write_bytes(private_key.private_bytes(
            serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption(),
        ))
        public_key = private_key.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
        with patch.object(generator, "TRUSTED_PUBLIC_KEYS", {"v2": base64.b64encode(public_key).decode("ascii")}):
            output = StringIO()
            with redirect_stdout(output):
                result = generator.main([
                    "--licensed-to", "Synthetic Customer", "--private-key", str(key_path),
                    "--key-id", "v2", "--output-dir", str(self.directory),
                ])
        self.assertEqual(result, 0)
        self.assertIn("Signing key ID: v2", output.getvalue())
        self.assertIn(base64.b64encode(public_key).decode("ascii"), output.getvalue())
        license_path = self.directory / "scheduling-maestro.license"
        key_file = self.directory / "license-key.txt"
        service = LicenseService(self.repository, {"v1": self.public_key, "v2": public_key})
        file_status = service.activate_license_file(license_path.read_bytes())
        key_status = service.activate_license_key(key_file.read_text(encoding="utf-8"))
        self.assertEqual(file_status.payload.key_id, "v2")
        self.assertEqual(file_status.signing_key_id, "v2")
        self.assertEqual(key_status.payload, file_status.payload)
        self.assertEqual(key_status.signing_key_id, "v2")

    def test_unconfigured_active_key_fails_before_writing(self):
        for key_id, error in (("v2", "not configured"), (" v1 ", "identifier")):
            with self.subTest(key_id=key_id):
                with self.assertRaisesRegex(generator.GenerationError, error):
                    generator.generate_license("Synthetic Customer", self.directory, self.key_path, key_id)
        self.assertFalse((self.directory / "scheduling-maestro.license").exists())
        self.assertFalse((self.directory / "license-key.txt").exists())

    def test_private_key_argument_is_required(self):
        error_output = StringIO()
        with redirect_stderr(error_output), self.assertRaises(SystemExit) as raised:
            generator.main(["--licensed-to", "Synthetic Customer", "--output-dir", str(self.directory)])
        self.assertEqual(raised.exception.code, 2)
        self.assertIn("--private-key", error_output.getvalue())
        self.assertFalse((self.directory / "scheduling-maestro.license").exists())

    def test_missing_invalid_or_mismatched_key_writes_nothing(self):
        for key_data, expected in (
            (None, "Cannot read private signing key"),
            (b"not a PEM", "not a readable"),
            (generate_private_key(public_exponent=65537, key_size=2048).private_bytes(
                serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption(),
            ), "must be an Ed25519 private key"),
            (Ed25519PrivateKey.generate().private_bytes(
                serialization.Encoding.PEM,
                serialization.PrivateFormat.PKCS8,
                serialization.NoEncryption(),
            ), "does not match"),
        ):
            with self.subTest(expected=expected):
                if key_data is None:
                    self.key_path.unlink(missing_ok=True)
                else:
                    self.key_path.write_bytes(key_data)
                with self.assertRaisesRegex(generator.GenerationError, expected):
                    generator.generate_license("Synthetic Customer", self.directory, self.key_path)
                self.assertFalse((self.directory / "scheduling-maestro.license").exists())
                self.assertFalse((self.directory / "license-key.txt").exists())

    def test_invalid_name_or_existing_output_does_not_overwrite(self):
        for name in ("", "  ", " Synthetic Customer ", "x" * 301):
            with self.subTest(name=name):
                with self.assertRaisesRegex(generator.GenerationError, "Licensed-to"):
                    generator.generate_license(name, self.directory, self.key_path)
        existing_key = self.directory / "license-key.txt"
        existing_key.write_text("existing output", encoding="utf-8")
        with self.assertRaisesRegex(generator.GenerationError, "without files named"):
            generator.generate_license("Synthetic Customer", self.directory, self.key_path)
        self.assertEqual(existing_key.read_text(encoding="utf-8"), "existing output")
        self.assertFalse((self.directory / "scheduling-maestro.license").exists())

    def test_cli_reports_unreadable_key_without_writing(self):
        error_output = StringIO()
        with patch.object(Path, "read_bytes", side_effect=PermissionError("Access denied")):
            with redirect_stderr(error_output), self.assertRaises(SystemExit) as raised:
                generator.main([
                    "--licensed-to", "Synthetic Customer", "--private-key", str(self.key_path),
                    "--output-dir", str(self.directory),
                ])
        self.assertEqual(raised.exception.code, 1)
        self.assertIn("Cannot read private signing key", error_output.getvalue())
        self.assertFalse((self.directory / "scheduling-maestro.license").exists())
        self.assertFalse((self.directory / "license-key.txt").exists())


if __name__ == "__main__":
    unittest.main()
