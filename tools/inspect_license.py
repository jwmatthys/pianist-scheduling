#!/usr/bin/env python3
"""Developer-only offline license inspection without activation."""

import argparse
import sys
from pathlib import Path

REPOSITORY_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPOSITORY_ROOT / "webapp" / "backend"))

from app.license_keys import TRUSTED_PUBLIC_KEYS
from app.services.license_service import LicenseServiceError, service


def _display(value: str) -> str:
    return "".join(character if character.isprintable() else f"\\u{ord(character):04x}" for character in value)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect an offline Scheduling Maestro license without activating it.")
    inputs = parser.add_mutually_exclusive_group(required=True)
    inputs.add_argument("--file", type=Path, help="Path to a license file or text file containing a license key.")
    inputs.add_argument("--key", help="Encoded license key.")
    args = parser.parse_args(argv)

    try:
        if args.file is not None:
            path = args.file.expanduser()
            try:
                with path.open("rb") as source:
                    contents = source.read(service.max_license_file_bytes + 1)
            except OSError as error:
                print(f"Error: Cannot read license file at {path}: {error.strerror or error}.", file=sys.stderr)
                return 2
            payload, signature = service.parse_license_file(contents)
        else:
            payload, signature = service.parse_license_key(args.key)
    except LicenseServiceError:
        print("Verification Result: CORRUPTED LICENSE")
        return 1

    try:
        signing_key_id = service.verify_license(payload, signature)
    except UnicodeEncodeError:
        print("Verification Result: CORRUPTED LICENSE")
        return 1
    except LicenseServiceError as error:
        if error.code == "INVALID_SIGNATURE":
            result = (
                "UNKNOWN SIGNING KEY"
                if payload.key_id is not None and payload.key_id not in TRUSTED_PUBLIC_KEYS
                else "INVALID SIGNATURE"
            )
        elif error.code == "WRONG_PRODUCT":
            result = "WRONG PRODUCT"
        else:
            raise
        signing_key_id = None
    else:
        result = "VALID"

    print(f"Product: {_display(payload.product)}")
    print(f"Licensed To: {_display(payload.licensed_to)}")
    print(f"License Type: {_display(payload.license_type)}")
    print(f"Issued: {payload.issued.isoformat()}")
    print(f"License Version: {payload.license_version}")
    if signing_key_id is None:
        print(f"Signing Key ID: {_display(payload.key_id or 'not specified')} (unverified)")
        print("License metadata is unverified.")
        if result == "UNKNOWN SIGNING KEY":
            print("The declared key is not trusted and no trusted key verified the signature; the license may also be corrupted.")
    else:
        print(f"Signing Key ID: {_display(signing_key_id)}")
        if payload.key_id is not None and payload.key_id != signing_key_id:
            print(f"Declared Key ID: {_display(payload.key_id)} (differs from verified signer)")
    print(f"Verification Result: {result}")
    return 0 if result == "VALID" else 1


if __name__ == "__main__":
    raise SystemExit(main())
