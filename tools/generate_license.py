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

from app.license_schemas import LicensePayload
from app.services.license_service import (
    PRODUCT_NAME,
    PUBLIC_KEY_B64,
    canonical_payload_bytes,
    encode_license_key,
)

PRIVATE_KEY_PATH = REPOSITORY_ROOT / ".local-license-keys" / "scheduling-maestro-ed25519-private.pem"


class GenerationError(Exception):
    pass


def load_private_key(path: Path) -> Ed25519PrivateKey:
    try:
        pem = path.read_bytes()
    except OSError as error:
        raise GenerationError(
            f"Cannot read private signing key at {path}: {error.strerror or error}. "
            "Restore the original key from a secure backup."
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
    if base64.b64encode(public_key).decode("ascii") != PUBLIC_KEY_B64:
        raise GenerationError(
            "The private signing key does not match the application's embedded public key. "
            "Restore the matching private key before generating licenses."
        )
    return private_key


def generate_license(licensed_to: str, output_dir: Path) -> tuple[Path, Path]:
    private_key = load_private_key(PRIVATE_KEY_PATH)
    try:
        payload = LicensePayload.model_validate({
            "product": PRODUCT_NAME,
            "licensed_to": licensed_to,
            "license_type": "PERPETUAL",
            "issued": date.today().isoformat(),
            "license_version": 1,
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
        {"payload": payload.model_dump(mode="json"), "signature": signature},
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
        "--output-dir",
        type=Path,
        default=Path.cwd(),
        help="Existing output directory (default: current directory). Existing files are not overwritten.",
    )
    args = parser.parse_args(argv)
    try:
        license_path, key_path = generate_license(args.licensed_to, args.output_dir)
    except GenerationError as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"License file: {license_path}\nLicense key: {key_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
