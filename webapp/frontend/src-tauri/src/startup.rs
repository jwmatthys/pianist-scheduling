//! Opt-in startup timing for the desktop shell.
//!
//! Active in debug builds, or in any build when `SCHEDULING_MAESTRO_STARTUP_PROFILE=1`
//! (`=0` forces it off). Only fixed phase labels and elapsed milliseconds are
//! recorded; never pass user or session data. To remove the instrumentation, delete
//! this module and the `startup::` calls.

use std::{
    sync::{Mutex, OnceLock},
    time::Instant,
};

pub const ENV_FLAG: &str = "SCHEDULING_MAESTRO_STARTUP_PROFILE";

/// Milestones the frontend may report through the `startup_mark` command.
pub const FRONTEND_MILESTONES: &[&str] = &[
    "React bootstrap begin",
    "Platform initialized",
    "API base resolution begin",
    "API base resolution complete",
    "Initial API calls begin",
    "Session load begin",
    "Session load complete",
    "Dashboard render begin",
    "Dashboard render complete",
    "First usable UI rendered",
    "Session load failed",
];

/// Milestones reported by the backend process as `STARTUP-EVENT <token>` stdout lines.
pub const BACKEND_EVENTS: &[(&str, &str)] = &[
    ("database-initialization-begin", "Database initialization begin"),
    ("database-initialization-failed", "Database initialization failed"),
];

struct Profile {
    start: Instant,
    enabled: bool,
    marks: Mutex<Vec<(String, f64)>>,
}

static PROFILE: OnceLock<Profile> = OnceLock::new();
static LISTENER: OnceLock<fn(&str)> = OnceLock::new();

fn resolve_enabled() -> bool {
    match std::env::var(ENV_FLAG).as_deref() {
        Ok("1") => true,
        Ok("0") => false,
        _ => cfg!(debug_assertions),
    }
}

/// Starts the monotonic clock; the first call defines "+0ms".
pub fn init() {
    PROFILE.get_or_init(|| Profile {
        start: Instant::now(),
        enabled: resolve_enabled(),
        marks: Mutex::new(Vec::new()),
    });
}

/// Whether timing logs are printed (and the backend asked to print its own).
pub fn enabled() -> bool {
    PROFILE.get().is_some_and(|profile| profile.enabled)
}

/// Registers the single consumer (the splash screen) of first-time milestones.
pub fn set_listener(listener: fn(&str)) {
    let _ = LISTENER.set(listener);
}

/// Records a milestone once (first occurrence wins) and notifies the listener.
/// Logging is opt-in, but recording and notification always happen so the splash
/// screen works in every build. Returns elapsed milliseconds for a new milestone.
pub fn mark(label: &str, detail: Option<&str>) -> Option<f64> {
    let profile = PROFILE.get()?;
    let elapsed = profile.start.elapsed().as_secs_f64() * 1000.0;
    {
        let mut marks = profile.marks.lock().ok()?;
        if marks.iter().any(|(existing, _)| existing == label) {
            return None;
        }
        marks.push((label.to_string(), elapsed));
    }
    if profile.enabled {
        match detail {
            Some(detail) => eprintln!("[STARTUP +{elapsed:.0}ms] {label} ({detail})"),
            None => eprintln!("[STARTUP +{elapsed:.0}ms] {label}"),
        }
    }
    if let Some(listener) = LISTENER.get() {
        listener(label);
    }
    Some(elapsed)
}

fn at(label: &str) -> Option<f64> {
    let profile = PROFILE.get()?;
    let marks = profile.marks.lock().ok()?;
    marks.iter().find(|(existing, _)| existing == label).map(|(_, value)| *value)
}

fn log_span(name: &str, begin: &str, end: &str) {
    if let (Some(begin), Some(end)) = (at(begin), at(end)) {
        eprintln!("[STARTUP]   {name}: {:.0} ms", end - begin);
    }
}

/// Prints measured values only; a span is omitted when either end was not recorded.
pub fn log_summary() {
    if !enabled() {
        return;
    }
    let Some(total) = at("First usable UI rendered") else {
        return;
    };
    eprintln!("[STARTUP] Total startup time: {total:.0} ms");
    log_span("Tauri startup (launch to backend spawn)", "App launch", "Backend spawn begin");
    log_span("Backend process spawn", "Backend spawn begin", "Backend spawn complete");
    log_span("Backend initialization (spawn complete to health ready)", "Backend spawn complete", "Backend health check success");
    log_span("Webview load and React bootstrap (endpoint published to React begin)", "API endpoint published", "React bootstrap begin");
    log_span("Frontend initialization (React begin to session request)", "React bootstrap begin", "Session load begin");
    log_span("Session load", "Session load begin", "Session load complete");
    log_span("Dashboard render (session loaded to first usable UI)", "Session load complete", "First usable UI rendered");
}
