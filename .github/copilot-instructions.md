# Accompanist Scheduler Development Instructions

## Architecture

- This is a Tauri 2 + React + TypeScript + Rust application.
- Keep UI/domain code independent of Tauri wherever practical.
- Platform-specific capabilities must go through the platform/service abstraction layer.
- Keep scheduling logic independent of Tauri, UI, persistence, and HTTP.
- Prefer a pure Rust scheduling-core crate if scheduling logic is implemented in Rust.
- Preserve the possibility of compiling the scheduling core to WebAssembly for a future browser version.
- Do not introduce a server dependency for functionality that can run locally.

## Privacy

This application handles student scheduling data.

The fundamental privacy principle is:

> The application provider should not need possession of student scheduling data in order to provide the scheduling service.

Therefore:

- Core scheduling must work offline.
- Student scheduling data must remain local by default.
- Do not add telemetry or analytics.
- Do not send student data to AI/LLM services.
- Do not send student data to remote services.
- Do not introduce cloud storage without explicit instruction.
- Use least-privilege Tauri permissions.
- Do not expose generic shell execution to the frontend.
- Avoid logging personally identifiable student information.
- Use only synthetic/anonymized student data in tests.

## Future Web Version

Treat these as independent decisions:

- Desktop vs Web
- Local vs Cloud

A future browser/SaaS version must be capable of running the scheduling engine and storing scheduling data locally in the browser.

Do not tightly couple:
- React to Tauri
- scheduling to Tauri
- scheduling to HTTP
- scheduling to a server
- project storage to a server
- licensing to student scheduling data

## Development Practices

- Preserve existing functionality unless explicitly changing it.
- Prefer incremental changes over repository-wide rewrites.
- Add regression tests before replacing working scheduling logic.
- Prefer boring, maintainable solutions over clever ones.
- Do not commit credentials, signing keys, certificates, passwords, tokens, or personally identifiable student data.
- Keep Tauri command handlers thin.
- Use structured types rather than unstructured JSON where practical.
- Return structured errors for recoverable failures.
- Do not use unwrap()/expect() for normal user-input or file-processing errors.

## Distribution

The application will eventually be commercially distributed on Windows and macOS.

Maintain compatibility with:
- Windows Microsoft Store distribution
- appropriately signed Windows direct distribution
- Apple Developer ID signing
- Apple notarization
- Prioritize Apple Silicon, but support Intel macOS where practical

Do not add actual production signing credentials to the repository.
