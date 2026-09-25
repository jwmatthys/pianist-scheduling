PROJECT GOAL

I want to migrate this application from Electron to Tauri 2 while preserving its existing user interface, scheduling behavior, data formats, and functionality.

This is an early-stage application, so now is the appropriate time to establish a clean long-term architecture.

The eventual product will be a commercial scheduling application for small university and music-school departments. It will process student scheduling information, so security, privacy, maintainability, cross-platform support, code signing, and easy distribution are priorities.

The product should initially work as a desktop application for Windows and macOS, but the architecture must preserve the ability to offer essentially the same product later as a browser-hosted SaaS/web application.

A major privacy goal is that BOTH the desktop application and a future web application should be capable of performing their core scheduling functions without sending student scheduling data to our servers.

The guiding principle is:

"The application provider does not need possession of student scheduling data in order to provide the scheduling service."

IMPORTANT: Before changing anything, inspect the entire existing repository, including:

- package.json
- Electron main/preload code
- frontend code
- Python/FastAPI/uvicorn backend
- scheduling/optimization code
- build scripts
- tests
- spreadsheet/file import/export logic
- persistence/data formats
- any native Node dependencies
- any Electron IPC
- any localhost HTTP APIs

Determine and document how the existing application currently works.

Do NOT simply perform a mechanical Electron-to-Tauri translation.

First create a migration plan based on the actual repository. Then execute the migration incrementally, testing after each major phase. Preserve existing working functionality unless a change is necessary for the new architecture.


==================================================
1. TARGET ARCHITECTURE
==================================================

Use:

- Tauri 2
- the existing frontend framework and UI wherever practical
- React if that is the existing frontend
- TypeScript rather than JavaScript where practical without creating an unnecessary frontend rewrite
- Rust for the Tauri native backend and native/system integration
- current stable Tauri 2 APIs and plugins
- npm as the JavaScript package manager unless this project already intentionally uses something else

The final installed desktop application must NOT require the user to install:

- Python
- Node.js
- Rust
- uvicorn
- a web server
- developer tools

The installed desktop application must be self-contained.

The long-term conceptual architecture should be:

                    Shared Application
                           |
             Shared application/domain layer
                           |
            Platform-independent interfaces
                   /                 \
                  /                   \
            Tauri adapter          Browser adapter
                |                       |
        Rust/native services        Browser APIs
                |                       |
            Local data              Local data


Core scheduling functionality should not depend unnecessarily on:

- Electron
- Tauri
- HTTP
- a remote server
- cloud storage
- a cloud account


==================================================
2. FUTURE WEB / SAAS COMPATIBILITY
==================================================

A major architectural requirement is that this application may later also be offered as a web/SaaS product.

Design the application so that the frontend, scheduling domain model, import/export formats, and as much application logic as practical can be reused in both:

1. the Tauri desktop application
2. a browser-hosted web/SaaS application

IMPORTANT PRIVACY REQUIREMENT:

The future SaaS/web version should be capable of operating in a local-data mode in which student scheduling data remains on the user's computer and is not uploaded to our servers.

In such a web deployment, our server may deliver:

- the web application
- static frontend assets
- application updates
- documentation
- licensing/account information if eventually needed

but the scheduling dataset itself should be capable of remaining entirely client-side.

The future web architecture should be capable of looking conceptually like:

Browser downloads application
          |
          v
      React UI
          |
          v
Shared application/domain layer
          |
          v
Local scheduling engine
          |
          +-- Imported spreadsheets processed locally
          +-- Projects stored locally
          +-- Scheduling performed locally
          +-- Exports generated locally

Student data should not have to pass through our server merely because the application is browser-hosted.

Do NOT build the frontend so ordinary application logic directly depends throughout the codebase on Tauri APIs.

Instead, create explicit abstractions/interfaces for platform-specific capabilities such as:

- opening/importing files
- saving/exporting files
- persistent project storage
- application preferences
- native dialogs
- platform information
- logging
- updates
- licensing if eventually required

Conceptually, the application might depend on interfaces/services such as:

- FileService
- ProjectStorage
- SettingsService
- PlatformService
- UpdateService

The Tauri desktop build can implement these using Tauri/Rust.

A future browser build can implement the same interfaces using browser APIs such as:

- File System Access APIs where appropriate
- file upload/download APIs
- IndexedDB
- browser-local persistent storage

UI components should not normally call Tauri invoke(), Rust commands, Tauri filesystem APIs, or browser persistence APIs directly.

They should communicate through the application's service/platform abstraction layer.


==================================================
3. LOCAL-FIRST / CLOUD-OPTIONAL ARCHITECTURE
==================================================

Treat these as SEPARATE architectural choices:

- Desktop vs. Web
- Local vs. Cloud

Do not assume:

Desktop = local forever

or:

Web/SaaS = server-stored student data

Core scheduling should be able to run locally in either environment.

Favor this model:

                    Shared Application
                           |
                 Local scheduling engine
                           |
              Local scheduling/project data
                    /              \
                   /                \
             Tauri/Desktop       Web/Browser


Optional future services may eventually exist above this architecture:

- licensing
- application updates
- optional cloud sync
- optional cloud backup
- optional multi-device access
- optional collaboration

But core scheduling must not depend on any of them.

If future collaboration or cloud synchronization is added, it should be an optional feature rather than a requirement for using the scheduler.

Do not make architectural decisions during this migration that unnecessarily couple:

- React to Tauri
- scheduling to Tauri
- scheduling to HTTP
- scheduling to a server
- project storage to a server
- student data to a cloud account
- licensing to student data

Licensing and payment systems, if eventually introduced, should be architecturally independent of student scheduling information.


==================================================
4. SHARED SCHEDULING ENGINE
==================================================

Evaluate carefully where the scheduling/optimization engine should ultimately live.

If the scheduling engine is rewritten entirely as Rust code behind Tauri-specific commands, a future browser version could require:

- a duplicate implementation
- WebAssembly
- or server-side scheduling

Because we specifically want the OPTION of keeping student data local in a future browser version, avoid architectural decisions that unnecessarily force scheduling data to a server.

If Rust is ultimately the best implementation language for the scheduling engine, structure the scheduling code as a pure Rust library/crate that contains NO Tauri dependencies.

Prefer conceptually:

                scheduling_core
                pure Rust library
                   /       \
                  /         \
          Tauri Desktop    Future WASM
              |                |
           Desktop           Browser

over:

Tauri command handlers
        |
scheduling logic embedded directly in handlers

Do NOT introduce WebAssembly now unless it provides an immediate practical benefit.

The current goal is architectural separation so that compiling the scheduling core to WebAssembly remains a realistic future option.

The domain/scheduling engine should ideally be:

- independently testable
- independent of UI
- independent of Tauri
- independent of HTTP
- independent of persistence
- reusable in future related scheduling products


==================================================
5. PYTHON / UVICORN MIGRATION
==================================================

The current prototype runs a Python backend using uvicorn.

Inspect exactly what the Python backend does before deciding how to replace it.

The preferred long-term desktop architecture is:

React frontend
       |
shared application layer
       |
typed platform/scheduling interface
       |
Rust
       |
local files/data

I would prefer to eliminate the localhost HTTP server in the production application.

However:

DO NOT rewrite working Python scheduling or optimization logic into Rust merely for architectural purity if doing so introduces substantial migration risk.

Follow this process:

1. Identify and document every API endpoint/function exposed by the existing Python/FastAPI backend.

2. Separate scheduling/domain logic from HTTP/FastAPI/uvicorn-specific logic.

3. Determine which functionality can safely move to Rust immediately.

4. Prefer native Rust implementations for straightforward:
   - filesystem operations
   - configuration
   - persistence
   - validation
   - native dialogs/system integration
   - application management

5. If the scheduling/optimization implementation is substantial and already reliable, preserve it temporarily as a bundled Tauri sidecar rather than rewriting it incorrectly.

6. If Python temporarily remains as a sidecar, package it so the end user does NOT need Python installed.

7. Do not expose a persistent unauthenticated localhost web service in the shipping application if it can reasonably be avoided.

8. If a Python sidecar is used, prefer a narrowly defined IPC/stdin/stdout/native invocation interface rather than maintaining the current application architecture as a general localhost HTTP service, if practical.

9. Ensure closing the application terminates all sidecar/helper processes.

10. Document any remaining Python dependency and establish a clear future migration path.

Eventually I would prefer production releases to have no Python runtime if the scheduling engine can be safely and correctly implemented in Rust.

Correctness is more important than eliminating Python immediately.


==================================================
6. PRIVACY AND SECURITY
==================================================

This application handles student scheduling information.

Adopt a local-first, least-privilege architecture.

The production desktop application should:

- process student scheduling data locally
- store project data locally
- not upload student information
- not require a cloud account
- not send student data to analytics services
- not send student data to telemetry services
- not send student data to AI or LLM services
- not send student data to advertising services
- avoid unnecessary network access
- avoid a network-accessible local server where practical
- request only the minimum Tauri permissions/capabilities necessary
- validate all arguments crossing the frontend/native boundary
- avoid arbitrary shell command execution
- never expose a generic "execute command" API to the frontend
- restrict filesystem access appropriately
- keep secrets/signing keys/certificates/passwords out of source control
- make it possible to truthfully tell customers that scheduling data remains on their device

Use Tauri 2's capability/permission model according to least-privilege principles.

Do not broadly enable native functionality just because it is convenient.

Do not add telemetry.

Do not add analytics.

Do not add network dependencies unless they provide a necessary and documented feature.

Do not introduce generative AI into the scheduling/data-processing path.

Use synthetic/anonymized data for repository test fixtures.


==================================================
7. ELECTRON REMOVAL
==================================================

Identify all Electron-specific functionality, including:

- main process
- preload scripts
- IPC handlers
- BrowserWindow configuration
- Electron dialogs
- Electron filesystem integration
- menus
- updater
- shell integration
- packaging
- build scripts
- Electron-specific security settings

For each item, document its replacement in the Tauri architecture.

Remove Electron dependencies and obsolete Electron configuration ONLY AFTER the equivalent Tauri functionality works.

Do not leave dead Electron code in the completed migration.


==================================================
8. PLATFORM ABSTRACTION
==================================================

Keep platform-specific functionality isolated.

Do not scatter code such as:

invoke(...)
Tauri-specific imports
filesystem calls
IndexedDB calls
browser storage calls

through ordinary React components.

Create a clear platform/service layer.

For example:

src/
    domain/
    scheduling/
    services/
    platform/
        interfaces/
        tauri/
        web/
    components/
    views/

This is only an example. Adapt the structure intelligently to the existing project.

The important requirement is:

UI/domain code
    |
platform-independent interface
    |
implementation
   / \
Tauri Browser

A future web version should be possible without rewriting the application's domain/UI architecture.


==================================================
9. FILE IMPORT / EXPORT
==================================================

Preserve all existing spreadsheet, CSV, and other file import/export behavior.

Use native Tauri file dialogs where appropriate in the desktop application.

Users should explicitly select files to import/open/save/export rather than granting unrestricted frontend filesystem access.

Keep:

- file selection
- parsing
- validation
- scheduling model conversion
- export generation

as separate concerns.

The parsing and validation layers must remain independently testable.

Preserve existing file formats where practical so data created with the current prototype remains compatible.


==================================================
10. APPLICATION DATA
==================================================

Create clear separation among:

1. imported source data
2. internal scheduling/domain models
3. application preferences
4. saved projects
5. generated/exported schedules

For desktop:

Use appropriate OS-specific application data/configuration locations rather than hard-coded paths.

Do not write application state beside the executable.

Never store user project data inside the installation directory.

For future web:

Preserve compatibility with client-side storage such as IndexedDB or similar browser-local storage.

Do not make saved project formats dependent upon an operating-system-specific storage implementation.


==================================================
11. RUST ARCHITECTURE
==================================================

Keep Rust modules small and purpose-specific.

A reasonable structure might resemble:

src-tauri/src/
    lib.rs
    commands/
    domain/
    scheduling/
    storage/
    import/
    export/
    validation/
    platform/

Do not force this exact structure if the actual project suggests something better.

Prefer an independent Rust workspace/crate for scheduling/domain logic if appropriate.

Tauri commands should be thin adapters calling ordinary Rust functions.

Do not put substantial scheduling/business logic directly in command handlers.

Use serde-compatible typed request/response structures between TypeScript and Rust where practical.

Avoid passing large amounts of unstructured/stringly-typed JSON when proper types would make the interface clearer and safer.


==================================================
12. ERROR HANDLING
==================================================

Do not use unwrap()/expect() for normal runtime situations involving:

- user input
- imported files
- missing files
- malformed spreadsheets
- failed exports
- validation errors
- recoverable scheduling failures

Return structured, understandable errors across platform boundaries.

Do not expose stack traces, filesystem internals, secrets, or sensitive information to ordinary users.

Application logs should be useful for debugging but must avoid personally identifiable student information wherever practical.

If possible, structure diagnostic reporting so a user can provide debugging information without disclosing student names or other scheduling records.


==================================================
13. DESKTOP USER EXPERIENCE
==================================================

Preserve the current visual interface unless a change is necessary.

The application should behave like a normal polished desktop application:

- native open/save dialogs
- sensible windows
- keyboard shortcuts where appropriate
- proper application name
- proper application icon
- clean startup/shutdown
- no orphan processes
- graceful errors
- predictable project saving
- appropriate unsaved-change warnings if relevant

Closing the application must cleanly terminate any temporary Python sidecar or helper process.


==================================================
14. BUILD EXPERIENCE
==================================================

I can currently build the Electron application with a single npm command.

Preserve similar simplicity.

Create clear npm scripts, ideally along the lines of:

npm run dev
npm run tauri:dev
npm run build
npm run tauri:build
npm test

Use existing project conventions where appropriate rather than introducing unnecessary duplicate scripts.

A clean checkout should have documented, reproducible prerequisites and build steps.

Do NOT require signing credentials for ordinary development/debug builds.


==================================================
15. DISTRIBUTION ARCHITECTURE
==================================================

Design the project now so production signing/distribution can be added later without restructuring the application.

Do NOT add real signing credentials now.

Where configuration will eventually need:

- publisher identity
- package identifier
- certificate identity
- product URLs
- signing configuration

use clear placeholders and document them.

Keep application/package identifiers stable once we approach public releases.


==================================================
16. WINDOWS DISTRIBUTION
==================================================

The eventual Windows version must be suitable for trusted commercial distribution and Microsoft Store submission.

Keep compatibility with current Tauri 2 and Microsoft-supported Windows distribution practices.

Do not prematurely lock the project into only one Windows packaging strategy.

Preserve the ability to choose between:

1. Microsoft Store distribution using an appropriate Store-compatible installer/package and Microsoft's current MSIX tooling if appropriate

2. signed MSI/EXE distribution through the Microsoft Store if appropriate

3. signed direct-download installer from our own website

For Microsoft Store builds, prepare the architecture for requirements such as:

- silent installation where required
- appropriate WebView2 packaging/install configuration
- package identity
- deterministic versioning
- proper icons/assets
- Store-compatible packaging

Keep Store-specific configuration separate from ordinary development/direct-distribution configuration where appropriate.

Do not purchase certificates or configure production signing yet.


==================================================
17. MACOS DISTRIBUTION
==================================================

Prepare the application for eventual direct distribution outside the Mac App Store.

Design for eventual use of:

- Apple Developer ID Application signing
- hardened runtime
- Apple notarization
- stapling/notarized distribution
- Tauri-supported macOS packaging such as DMG

Do not store:

- Apple IDs
- passwords
- App Store Connect credentials
- signing certificates
- private keys
- notarization credentials

in source control.

Signing identities and credentials should eventually be supplied through secure local environment configuration or CI secret storage.

We may eventually evaluate Mac App Store distribution, so avoid unnecessary architectural choices that would make future App Store packaging difficult.


==================================================
18. PLATFORM ARCHITECTURES
==================================================

Plan for:

- Windows x86_64 initially
- Windows ARM64 in the future if practical
- macOS Apple Silicon / arm64
- macOS Intel / x86_64 if reasonably practical

Do not assume helper executables work automatically across architectures.

If Python or another sidecar remains, explicitly account for platform/architecture-specific sidecar binaries.


==================================================
19. CI/CD PREPARATION
==================================================

Do not configure real production signing yet.

Prepare the repository so we can eventually create GitHub Actions release workflows.

Future goals:

- Windows release builds on Windows runners
- macOS release builds on macOS runners
- signing/notarization credentials stored only in GitHub Actions secrets or another secure secret store
- reproducible release artifacts
- automatic testing before releases
- no private keys committed to git
- no secrets embedded in application source

Document the expected future release process.


==================================================
20. DEPENDENCIES
==================================================

Minimize dependencies.

Before adding a Tauri plugin or Rust crate:

- determine whether it is actually necessary
- prefer official Tauri plugins for standard Tauri functionality
- prefer mature/well-maintained dependencies
- avoid obscure dependencies for trivial tasks
- avoid packages introducing unnecessary network behavior

Do not upgrade unrelated frontend dependencies merely because newer versions exist unless required for compatibility/security.

Remove obsolete Electron dependencies when the migration is complete.


==================================================
21. TESTING
==================================================

Before migration, identify the existing tests.

Preserve or add tests for scheduling/domain behavior.

At minimum verify:

- application launches
- existing frontend renders correctly
- current scheduling examples still work
- imports work
- exports work
- saved projects/state reopen correctly if that feature exists
- malformed imports fail gracefully
- app closes without orphan processes
- production build completes
- paths/storage work appropriately

WHEN MIGRATING SCHEDULING LOGIC FROM PYTHON TO RUST:

Create equivalence/regression tests BEFORE removing the Python implementation.

Use representative anonymized inputs and compare:

OLD Python implementation

versus

NEW Rust implementation

Verify equivalent:

- hard constraint behavior
- assignments
- conflicts
- workloads/hours
- optimization objectives
- edge cases

Where the optimizer is nondeterministic, compare correctness and objective/constraint behavior rather than requiring byte-for-byte identical output.

Do not delete the working Python implementation until the Rust replacement is demonstrated to be correct.


==================================================
22. DOCUMENTATION
==================================================

Create or update:

README.md

Also create:

docs/ARCHITECTURE.md
docs/BUILDING.md
docs/DISTRIBUTION.md
docs/PRIVACY-ARCHITECTURE.md
docs/MIGRATION-ELECTRON-TO-TAURI.md

ARCHITECTURE.md should explain:

- frontend architecture
- platform abstraction
- Rust backend
- scheduling engine
- desktop/web separation
- Python sidecar if one remains
- future WebAssembly possibility
- persistence model

BUILDING.md should explain:

- prerequisites
- development setup
- Tauri development
- production builds
- platform-specific requirements

DISTRIBUTION.md should explain:

- development builds
- Windows builds
- future Windows signing
- Microsoft Store preparation
- direct Windows distribution
- macOS builds
- future Apple Developer ID signing
- notarization
- credentials eventually needed
- placeholders that must be replaced before release

PRIVACY-ARCHITECTURE.md should clearly document:

- what student/user data exists
- where data is stored
- whether data leaves the device
- filesystem permissions
- network behavior
- logging policy
- sidecars/helper processes
- third-party components relevant to privacy
- future local-data web architecture
- optional future cloud functionality

MIGRATION-ELECTRON-TO-TAURI.md should document:

- original architecture
- migration decisions
- functionality replaced
- Electron APIs removed
- Python functionality migrated
- remaining technical debt


==================================================
23. MIGRATION PROCESS
==================================================

Do the migration incrementally rather than as one giant edit.

PHASE 1: ANALYSIS

- inspect repository
- document current architecture
- identify Electron APIs
- identify Python/uvicorn APIs
- identify scheduling dependencies
- identify persistence/import/export architecture
- identify tests
- write migration plan

Before making substantial changes, show me the migration plan and any important architectural decisions.


PHASE 2: TAURI SHELL

- scaffold Tauri 2 around existing frontend
- make existing UI launch under Tauri
- preserve Electron temporarily if useful
- establish basic build commands


PHASE 3: PLATFORM ABSTRACTION

- create service/platform interfaces
- isolate Tauri-specific functionality
- replace Electron main/preload/IPC functionality
- establish typed frontend/Rust communication
- migrate file dialogs/filesystem integration


PHASE 4: PYTHON BACKEND

- inventory FastAPI/uvicorn behavior
- separate domain logic from HTTP
- migrate straightforward functionality to Rust
- eliminate unnecessary localhost HTTP architecture
- safely preserve complex Python optimization as a bundled sidecar if necessary


PHASE 5: SCHEDULING ENGINE

- isolate scheduling domain model
- create regression/equivalence tests
- determine whether scheduling should remain temporarily Python or become pure Rust
- if migrating, put scheduling logic in a Tauri-independent Rust crate
- preserve future WebAssembly compatibility where practical


PHASE 6: FEATURE PARITY

- verify imports
- verify exports
- verify scheduling
- verify manual scheduling changes
- verify persistence
- verify error handling
- verify graceful shutdown
- verify privacy/network behavior


PHASE 7: ELECTRON REMOVAL

Only after Tauri reaches feature parity:

- remove Electron
- remove preload/main process
- remove Electron IPC
- remove obsolete build tooling
- remove obsolete dependencies
- remove dead configuration


PHASE 8: PRODUCTION PREPARATION

- establish clean Tauri bundle configuration
- establish stable application identifiers
- prepare Windows bundle configuration
- prepare macOS bundle configuration
- ensure release builds complete
- complete documentation


==================================================
24. IMPORTANT WORKING RULES
==================================================

Do not claim a phase is complete until it has actually been tested.

Do not silently remove features.

Do not rewrite the frontend unnecessarily.

Do not replace reliable scheduling logic merely because Rust is available.

Do not introduce a remote server dependency for functionality that can operate locally.

Do not introduce cloud storage for convenience.

Do not introduce WebAssembly yet merely to demonstrate that it works.

Do not over-engineer future functionality that we do not currently need.

Instead, maintain clean architectural boundaries so future capabilities remain possible.

When encountering an architectural choice with significant long-term consequences:

1. Explain the alternatives.
2. Recommend one.
3. Explain why.
4. Identify implications for:
   - desktop distribution
   - future browser deployment
   - privacy
   - maintenance
   - code signing
   - testing
5. Then make the change only if it is consistent with the stated goals.

Prefer boring, maintainable solutions over clever ones.

Never commit:

- private keys
- signing certificates
- tokens
- passwords
- developer credentials
- personally identifiable test/student data

Use synthetic/anonymized fixtures.


==================================================
25. FINAL ARCHITECTURAL PRINCIPLES
==================================================

Keep these principles in mind throughout the migration:

1. Core scheduling must work without Internet access.

2. The vendor should not need possession of student scheduling data in order to provide the scheduling service.

3. Desktop vs. web and local vs. cloud are independent choices.

4. A future web edition should be capable of processing and storing student scheduling data entirely on the user's device.

5. Cloud sync/collaboration, if ever implemented, should be optional.

6. Licensing must not require access to student scheduling data.

7. React should not be tightly coupled to Tauri.

8. Scheduling logic should not be tightly coupled to Tauri.

9. Scheduling logic should not depend on HTTP.

10. Scheduling should not require server-side computation.

11. If Rust becomes the long-term scheduling implementation, keep it sufficiently independent that future WebAssembly compilation remains possible.

12. Use least privilege.

13. Collect and transmit as little data as possible.

14. Development convenience must not silently create long-term privacy/security obligations.

15. Preserve correctness above architectural purity.


==================================================
26. COMPLETION REPORT
==================================================

At the conclusion of the migration work, provide me with:

1. A concise description of the final architecture.

2. A diagram of the important layers and data flow.

3. A list of Electron code/dependencies removed.

4. A list of Python functionality migrated to Rust.

5. Any Python functionality still remaining and why.

6. A description of the scheduling engine and whether it is suitable for future WebAssembly compilation.

7. A description of the platform abstraction and how a future browser implementation could use it.

8. Current development and production build commands.

9. Known Windows-specific limitations.

10. Known macOS-specific limitations.

11. Work still required for Windows Store distribution.

12. Work still required for Windows direct-download signing.

13. Work still required for Apple Developer ID signing/notarization.

14. Security/privacy issues I should review.

15. Any remaining network communications made by the application.

16. Any third-party services that could receive application/user/student information.

17. Manual testing I should perform before considering this migration complete.

18. Technical debt introduced during the migration.

19. Recommended next steps.

20. Confirmation that the production scheduling workflow can operate without Internet access and without transmitting student scheduling data off the user's device.
