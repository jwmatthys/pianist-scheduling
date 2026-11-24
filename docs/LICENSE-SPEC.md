# Scheduling Maestro License and Activation System Specification

## Status

Phase 1 license infrastructure is implemented: offline Ed25519 verification, JSON license files, `SM1-` encoded license keys, local activation persistence, API endpoints, and a License section. Trial timing, expiration behavior, feature lockouts, and other Phase 2 behavior remain unimplemented.

## Purpose

Scheduling Maestro uses an offline perpetual-license system designed to:

- Provide a short evaluation period.
- Encourage purchase by honest users.
- Operate completely offline.
- Avoid accounts, passwords, subscriptions, and activation servers.
- Minimize customer support burden.

The system is not intended to provide enterprise-grade software protection. It is intended to discourage casual sharing while preserving a simple user experience.

---

# Goals

## Functional Goals

- Allow new users to evaluate the application for 30 minutes.
- Lock the application after the trial period expires.
- Allow activation using either:
  - a license file, or
  - a license key string.
- Unlock the application permanently.
- Operate entirely offline.
- Support Windows, macOS, and Linux.
- Identify the purchaser within the application.

## Non-Goals

The system does not attempt to:

- Prevent reverse engineering.
- Prevent determined piracy.
- Prevent customers from reinstalling on a new computer.
- Limit the number of computers on which a license may be installed.
- Support subscriptions.
- Support recurring billing.
- Support internet activation.
- Support online license validation.

---

# User Experience

## Trial Mode

Upon first launch:

```text
Scheduling Maestro

Trial Version

30 minutes remaining

[Import License File]
[Enter License Key]
```

The application functions normally during the trial period.

The trial timer measures cumulative usage time.

The remaining trial time should be visible from the Dashboard.

---

## Trial Expired

When cumulative usage reaches 30 minutes:

```text
Scheduling Maestro

Trial Expired

This copy requires activation.

[Import License File]
[Enter License Key]
[Quit]
```

All application features are locked.

Users may:

- activate the software
- quit the application

Users may not:

- import data
- edit data
- generate schedules
- view schedules
- modify settings

---

## Activated State

After successful activation:

```text
Scheduling Maestro

Licensed to:
University of Somewhere

License Type:
Perpetual
```

The application becomes permanently available.

No further activation is required.

---

# Activation Methods

Scheduling Maestro supports two activation methods.

## Method 1: License File

Preferred workflow.

File name:

```text
scheduling-maestro.license
```

User workflow:

```text
Dashboard
→ Import License File
→ Select scheduling-maestro.license
→ Activate
```

---

## Method 2: License Key

Alternative workflow.

User workflow:

```text
Dashboard
→ Enter License Key
→ Paste License Key
→ Activate
```

The license file and license key contain the same signed license payload.

---

# License Payload

Every license contains a signed payload.

Example:

```json
{
  "product": "Scheduling Maestro",
  "licensed_to": "University of Somewhere",
  "license_type": "PERPETUAL",
  "issued": "2026-11-23",
  "license_version": 1
}
```

---

## Fields

### product

Required.

Must equal:

```text
Scheduling Maestro
```

---

### licensed_to

Required.

Displayed throughout the application.

Examples:

```text
University of Somewhere

Dr. Jane Smith

Waukesha School District
```

---

### license_type

Required.

Current value:

```text
PERPETUAL
```

Future values may be added.

---

### issued

Required.

Date the license was created.

Used for informational display only.

Does not expire.

---

### license_version

Required.

Used to support future license format revisions.

Initial value:

```text
1
```

---

# Cryptographic Design

## Signing Algorithm

Use:

```text
Ed25519
```

---

## Private Key

The private key is used only by the license-generation tool.

The private key must never be distributed with the application.

---

## Public Key

The application ships with an embedded Ed25519 public key.

The application verifies all licenses locally.

---

## Verification

Activation succeeds only when:

- payload structure is valid
- product matches Scheduling Maestro
- signature verification succeeds

Activation fails otherwise.

---

# License File Format

File extension:

```text
.license
```

Example:

```json
{
  "payload": {
    "product": "Scheduling Maestro",
    "licensed_to": "University of Somewhere",
    "license_type": "PERPETUAL",
    "issued": "2026-11-23",
    "license_version": 1
  },
  "signature": "BASE64_SIGNATURE"
}
```

The file should remain human-readable.

---

# License Key Format

The license key is an encoded form of the same payload.

Example:

```text
SM-7B9K-3D2F-V8QL-T5MX-...
```

Human-friendly grouping is recommended.

The exact encoding format is implementation-defined.

The application must decode and verify the same underlying signed payload used in the license file.

### Phase 1 Encoding

The encoded key uses the `SM1-` prefix followed by unpadded URL-safe Base64 of the compact UTF-8 JSON envelope containing `payload` and `signature`. Signature generation signs the canonical payload JSON (UTF-8, sorted keys, compact separators); the key and license file carry the exact same payload and Base64 signature.

---

# Storage

## Trial Data

Stored locally.

Example fields:

```text
trial_seconds_used
first_run_timestamp
```

The trial period expires at:

```text
1800 seconds
```

(30 minutes)

---

## License Data

Stored locally after successful activation.

Stored information includes:

```text
payload
signature
activation_timestamp
```

---

# Dashboard Requirements

## Unlicensed

Display:

```text
Trial Version

XX minutes remaining
```

and activation controls.

---

## Licensed

Display:

```text
Licensed to:
<organization>

Perpetual License
```

The licensed entity should remain visible on the Dashboard at all times.

---

# Settings / About Screen

Provide a License section containing:

```text
Licensed To
License Type
Issue Date
Activation Date
```

and controls:

```text
Import License File
Enter License Key
```

---

# Offline Requirements

The licensing system must operate without:

- internet access
- cloud services
- activation servers
- user accounts
- email verification
- machine registration

Verification occurs entirely on the local machine.

---

# Cross-Platform Requirements

The licensing system must function identically on:

- Windows
- macOS
- Linux

No OS-specific identifiers may be used.

No machine fingerprinting may be used.

---

# Machine Independence

Licenses are not tied to a computer.

A valid license may be moved to another machine.

The same license may be re-imported if the user reinstalls the application.

---

# Error Handling

## Invalid License File

Show:

```text
This license file is invalid.
```

---

## Invalid Signature

Show:

```text
Unable to verify this license.
```

---

## Wrong Product

Show:

```text
This license was not issued for Scheduling Maestro.
```

---

## Corrupted License Key

Show:

```text
License key could not be read.
```

---

# Security Assumptions

This system is designed to prevent:

- accidental piracy
- casual sharing
- guessed license keys

This system is not designed to prevent:

- reverse engineering
- binary patching
- determined software piracy

This tradeoff is intentional.

---

## Generating Licenses

The generator is a developer-only tool, not part of application builds. From
the repository root, using the backend Python environment described in
[BUILDING.md](BUILDING.md), run:

```sh
python tools/generate_license.py \
  --licensed-to "University of Wisconsin-Whitewater"
```

This writes `scheduling-maestro.license` and `license-key.txt` in the current
directory. Pass `--output-dir /path/to/existing/directory` to select another
existing directory. The tool refuses to overwrite either output file. Both
files contain equivalent signed license data; share them with the customer
through an appropriate secure channel. Generated output filenames are
gitignored.

The tool reads the unencrypted Ed25519 private signing key from
`.local-license-keys/scheduling-maestro-ed25519-private.pem` in the repository
root, regardless of the current directory. It verifies that the key matches
the public key embedded in the application. The private key signs licenses;
the application uses only the public key for offline verification and never
reads the private key.

Back up the private key securely, encrypted and access-controlled, with at
least one separately stored offline copy. Test the ability to restore it.
Do not commit, distribute, or include it in builds; `.gitignore` is not a
security boundary and does not prevent other backup or synchronization
software from copying it. If the private key is lost, licenses cannot be
issued for the embedded public key; restore the matching backup. A replacement
key requires an application update and a deliberate compatibility plan for
previously issued licenses.

---

# Future Compatibility

Future versions may add:

- additional license types
- department licenses
- institutional licenses

without changing the verification architecture.

The introduction of future license types must not invalidate existing perpetual licenses.

---

# Success Criteria

A user can:

1. Install Scheduling Maestro.
2. Use it for 30 minutes.
3. Activate using either:
   - scheduling-maestro.license
   - a license key.
4. Continue using the application indefinitely.
5. Reinstall or move to another machine without contacting the developer.
6. Use all functionality without internet access.