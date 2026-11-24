#!/usr/bin/env python3
"""Developer-only offline license generator; never included in the application."""

import argparse
import base64
import json
import sys
from datetime import date
from pathlib import Path

from cryptography.exceptions import UnsupportedAlgorithm
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from pydantic import ValidationError

REPOSITORY_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPOSITORY_ROOT / "webapp" / "backend"))

from app.license_keys import TRUSTED_PUBLIC_KEYS
from app.license_schemas import LicensePayload
from app.services.license_service import (
    PRODUCT_NAME,
    canonical_payload_bytes,
    encode_license_key,
)

ACTIVE_SIGNING_KEY_ID = "v1"


class GenerationError(Exception):
    pass


def load_private_key(path: Path, expected_public_key: str, key_id: str) -> Ed25519PrivateKey:
    try:
        pem = path.read_bytes()
    except OSError as error:
        raise GenerationError(
            f"Cannot read private signing key at {path}: {error.strerror or error}. "
            "Provide a readable private key or restore it from a secure backup."
        ) from error
    try:
        private_key = serialization.load_pem_private_key(pem, password=None)
    except (ValueError, TypeError, UnsupportedAlgorithm) as error:
        raise GenerationError(
            f"The signing key at {path} is not a readable, unencrypted private PEM. "
            "Restore the original Ed25519 key from a secure backup."
        ) from error
    if not isinstance(private_key, Ed25519PrivateKey):
        raise GenerationError(f"The signing key at {path} must be an Ed25519 private key.")
    public_key = private_key.public_key().public_bytes(
        encoding=serialization.Encoding.Raw,
        format=serialization.PublicFormat.Raw,
    )
    if base64.b64encode(public_key).decode("ascii") != expected_public_key:
        raise GenerationError(
            f"The private signing key does not match trusted public key '{key_id}'. "
            "Restore the matching private key before generating licenses."
        )
    return private_key


def generate_license(
    licensed_to: str, output_dir: Path, private_key_path: Path, key_id: str = ACTIVE_SIGNING_KEY_ID,
) -> tuple[Path, Path]:
    if (
        not key_id.strip()
        or key_id != key_id.strip()
        or len(key_id) > 100
    ):
        raise GenerationError("Signing key identifier must be non-empty, trimmed, and at most 100 characters.")
    public_key = TRUSTED_PUBLIC_KEYS.get(key_id)
    if public_key is None:
        raise GenerationError(
            f"Signing key '{key_id}' is not configured in the application trust store. "
            "Select a trusted key ID."
        )
    private_key = load_private_key(private_key_path.expanduser(), public_key, key_id)
    try:
        payload = LicensePayload.model_validate({
            "product": PRODUCT_NAME,
            "licensed_to": licensed_to,
            "license_type": "PERPETUAL",
            "issued": date.today().isoformat(),
            "license_version": 1,
            "key_id": key_id,
        })
    except ValidationError as error:
        raise GenerationError(
            "Licensed-to must be non-empty, at most 300 characters, and have no surrounding whitespace."
        ) from error

    try:
        canonical_payload = canonical_payload_bytes(payload)
    except UnicodeEncodeError as error:
        raise GenerationError("Licensed-to must contain valid Unicode text.") from error
    signature = base64.b64encode(private_key.sign(canonical_payload)).decode("ascii")
    license_json = json.dumps(
        {"payload": payload.model_dump(mode="json", exclude_none=True), "signature": signature},
        ensure_ascii=False,
        indent=2,
    ) + "\n"
    license_key = encode_license_key(payload, signature)
    license_path = output_dir / "scheduling-maestro.license"
    key_path = output_dir / "license-key.txt"
    created: list[Path] = []
    try:
        with license_path.open("x", encoding="utf-8", newline="\n") as license_file:
            created.append(license_path)
            with key_path.open("x", encoding="utf-8", newline="\n") as key_file:
                created.append(key_path)
                license_file.write(license_json)
                key_file.write(license_key)
    except OSError as error:
        cleanup_errors = []
        for path in reversed(created):
            try:
                path.unlink()
            except OSError as cleanup_error:
                cleanup_errors.append(f"{path}: {cleanup_error}")
        message = (
            f"Could not write license outputs in {output_dir}: {error}. "
            "Choose an existing writable directory without files named "
            "scheduling-maestro.license or license-key.txt."
        )
        if cleanup_errors:
            message += f" Remove incomplete output manually: {'; '.join(cleanup_errors)}."
        raise GenerationError(message) from error
    return license_path, key_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate an offline Scheduling Maestro license.")
    parser.add_argument("--licensed-to", required=True, help="Name of the licensed person or organization.")
    parser.add_argument(
        "--private-key", required=True, type=Path,
        help="Path to the Ed25519 private signing key; it is never auto-discovered.",
    )
    parser.add_argument(
        "--key-id", default=ACTIVE_SIGNING_KEY_ID,
        help=f"Trusted signing key identifier (default: {ACTIVE_SIGNING_KEY_ID}).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path.cwd(),
        help="Existing output directory (default: current directory). Existing files are not overwritten.",
    )
    args = parser.parse_args(argv)
    try:
        license_path, key_path = generate_license(args.licensed_to, args.output_dir, args.private_key, args.key_id)
    except GenerationError as error:
        parser.exit(1, f"Error: {error}\n")
    print(
        f"Signing key ID: {args.key_id}\n"
        f"Trusted public key (base64): {TRUSTED_PUBLIC_KEYS[args.key_id]}\n"
        f"License file: {license_path}\nLicense key: {key_path}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
