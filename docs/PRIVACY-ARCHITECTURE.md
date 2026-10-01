# Privacy Architecture

Scheduling data remains on the user's device. The desktop application has no
telemetry, analytics, cloud storage, advertising, AI service, or external
network integration. React communicates with the local Python service only
over its ephemeral `127.0.0.1` listener; that listener is not exposed to the
LAN.

`.mpsession` archives are unencrypted and contain the session database, which
may include student educational information. Users should store and transfer
archives only under institutional policy. Export writes a user-selected
portable copy. Automatic replacement recovery archives are stored locally in
the application-data `recovery/` directory and limited to the latest three.

Tauri grants only file-open/file-save dialogs and read/write access to the
paths selected through those dialogs. Browser builds use local file selection
and download/save APIs. No session payload is sent to a remote service.