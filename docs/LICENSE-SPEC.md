# Scheduling Maestro License and Activation System Specification

## Status

Offline Ed25519 verification, license activation, and a 30-minute cumulative-use trial with expiration lockout are implemented.

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

The countdown is displayed beside License on the Dashboard. It counts while
the application is open, including time spent in modules, and is stored in
the local application database rather than in a Scheduling Session. The
frontend sends a local heartbeat every five seconds; a gap contributes at most
ten seconds, so interrupted or suspended sessions do not count indefinitely.
When the counter reaches zero, it remains at zero and the application locks
until a valid license is activated. A valid activation hides the countdown and
pauses trial accounting.

---

## Trial Expired

When cumulative usage reaches 30 minutes, an activation window blocks the
application:

```text
Scheduling Maestro

Scheduling Maestro trial expired. Upload or enter your license key to activate.

[Import License File]
[Enter License Key]
```

All application features are locked.

Users may:

- activate the software
- quit the application using the window controls

Users may not:

- import data
- edit data
- generate schedules
- view schedules
- modify settings

The local API rejects scheduling requests with `TRIAL_EXPIRED` while the trial
is expired. Health, license status, trial heartbeat, and license activation
remain available. The lock preserves the current in-memory view so users can
continue after activation; requests already in progress are not canceled.

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

A different valid license can be activated at any time using either method.
It replaces the stored license and its displayed metadata. An unsuccessful
activation leaves the existing license unchanged.

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
  "license_version": 1,
  "key_id": "v1"
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

### key_id

The identifier of the signing key, included in and covered by the signed
payload for newly issued licenses. Older licenses omit this field and remain
valid. The verified signing-key identifier is also reported as
`signing_key_id` in license status metadata.

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

The application ships with an embedded set of trusted Ed25519 public keys.

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
    "license_version": 1,
    "key_id": "v1"
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
  --private-key ~/SchedulingMaestroKeys/scheduling-maestro-ed25519-private.pem \
  --licensed-to "University of Wisconsin-Whitewater"
```

This writes `scheduling-maestro.license` and `license-key.txt` in the current
directory. Pass `--output-dir /path/to/existing/directory` to select another
existing directory. The tool refuses to overwrite either output file. Both
files contain equivalent signed license data; share them with the customer
through an appropriate secure channel. Generated output filenames are
gitignored.

In the License dialog, Import License File accepts either the JSON
`scheduling-maestro.license` file or the text `license-key.txt` file. The
contents of `license-key.txt` can also be pasted into Enter License Key.
Renaming a file does not change its contents.

Store the unencrypted Ed25519 private signing key outside the repository, for
example in `~/SchedulingMaestroKeys/`, with access restricted to authorized
issuers. The generator never discovers a private key automatically:
`--private-key` is required. It defaults to key ID `v1`; pass `--key-id v2`
when issuing with a different configured trusted key. It checks that the
provided private key matches the selected trusted public key before writing
anything and prints both the key ID and the public key. The application uses
only trusted public keys for offline verification and never reads the private
key.

Back up the private key securely, encrypted and access-controlled, with at
least one separately stored offline copy. Test the ability to restore it.
Do not commit, distribute, or include it in builds; `.gitignore` is not a
security boundary and does not prevent other backup or synchronization
software from copying it. If the private key is lost, licenses cannot be
issued for the embedded public key; restore the matching backup. A replacement
key requires an application update and a deliberate compatibility plan for
previously issued licenses. See [Signing Key Rotation](#signing-key-rotation).

## Inspecting Licenses

The developer-only inspector uses the application's parsing and verification
logic without activating or changing the stored license:

```sh
python tools/inspect_license.py --file scheduling-maestro.license
python tools/inspect_license.py --key "SM1-..."
```

It prints the license metadata, verified signing key ID, and result. Malformed
input reports `CORRUPTED LICENSE`; a failed signature reports `INVALID
SIGNATURE`, or `UNKNOWN SIGNING KEY` when its declared key ID is not in the
trust store. An unknown ID on a failed signature is only a hint: the data
could also be corrupted. Metadata shown for a failed verification is
unverified. When a trusted key verifies the signature, its actual ID is
reported even if the declared ID differs. A license for another product
reports `WRONG PRODUCT`. Valid inspection exits with status 0; invalid
licenses and file-read errors exit nonzero. Passing a key on a command line
may expose it in shell history or process listings; use a local file instead
when appropriate.

## Signing Key Rotation

The application embeds the trusted public-key mapping in
[`webapp/backend/app/license_keys.py`](../webapp/backend/app/license_keys.py).
Each identifier uniquely identifies one Ed25519 public key. Verification
tries the trusted keys and succeeds if any validates the canonical signed
payload. The optional signed `key_id` is the issuer's identifier; the
`signing_key_id` returned in status metadata is determined by the key that
actually verified the signature, not by trusting that claim.

The original `v1` public key remains trusted. Legacy payloads without
`key_id` retain their original canonical bytes and existing stored
activations are reverified without modification or reactivation. The
`.license` envelope and `SM1-` encoding remain unchanged.

To introduce a future `v2` key:

1. Generate and securely back up a new private key outside application builds.
2. Add its public key under a unique `v2` identifier to the trust store, while
   retaining `v1`. Test both legacy and new signatures, then distribute the
   application update before issuing licenses with the new key.
3. Keep the `v2` private key outside the repository, and pass its path with
   `--private-key` and `--key-id v2` to the developer-only generator. The
   generator checks that the private key matches the selected trusted public
   key before writing either output.
4. Continue trusting `v1` so existing customers do not need to reactivate.

**Compatibility:** older application versions reject the new `key_id` field,
even on licenses still signed by `v1`. Distribute this application update
before issuing any newly generated license containing `key_id`. Previously
issued licenses continue to work in the updated application.

Keeping a compromised key trusted also keeps accepting signatures made with
it. Removing that key would invalidate its existing licenses, so any response
to compromise requires a separate migration decision; there is no online
revocation service.

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