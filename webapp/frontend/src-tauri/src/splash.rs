//! Native startup splash window, driven only by real startup milestones.
//!
//! `startup::mark` is the single event path: this module registers as its listener and
//! turns each first-time milestone into splash text. The splash page and image are served
//! from the binary through a custom protocol, so it needs no IPC capability and does not
//! depend on the frontend dev server or build output. No timers drive progress; the only
//! timer is a failure watchdog.

use std::{
    borrow::Cow,
    sync::{Mutex, OnceLock},
    thread,
    time::Duration,
};

use tauri::{
    http::{Request, Response},
    window::Color,
    AppHandle, Manager, Runtime, UriSchemeContext, Url, WebviewUrl, WebviewWindowBuilder,
};
#[cfg(target_os = "linux")]
use tauri::LogicalSize;

pub const SCHEME: &str = "splash";
const LABEL: &str = "splash";
const READY_MILESTONE: &str = "First usable UI rendered";
const INITIAL_MESSAGE: &str = "Starting Scheduling Maestro...";

// Placeholder artwork; replace the file to change the splash image.
static IMAGE: &[u8] = include_bytes!("../assets/splash/splash.png");
static PAGE: &str = include_str!("../assets/splash/splash.html");

/// Milestone -> text shown while it is the latest startup event.
const PROGRESS_MESSAGES: &[(&str, &str)] = &[
    ("App launch", INITIAL_MESSAGE),
    ("Backend spawn begin", "Launching backend..."),
    ("Backend health check begin", "Waiting for backend..."),
    ("Database initialization begin", "Initializing database..."),
    ("React bootstrap begin", "Loading application modules..."),
    ("Session load begin", "Loading session..."),
    ("Dashboard render begin", "Preparing dashboard..."),
];

/// Milestone -> concise failure text; the splash stays visible with this message.
const FAILURE_MESSAGES: &[(&str, &str)] = &[
    ("Database initialization failed", "Database initialization failed."),
    ("Session load failed", "Unable to load the session."),
];

pub const MSG_BACKEND_FAILED: &str = "Unable to start backend.";
pub const MSG_TIMED_OUT: &str = "Startup timed out.";
pub const MSG_WINDOW_FAILED: &str = "Unable to open the main window.";

struct State {
    message: &'static str,
    failed: bool,
    ready: bool,
}

static STATE: Mutex<State> = Mutex::new(State {
    message: INITIAL_MESSAGE,
    failed: false,
    ready: false,
});
static APP: OnceLock<AppHandle> = OnceLock::new();

fn lookup(table: &'static [(&'static str, &'static str)], label: &str) -> Option<&'static str> {
    table.iter().find(|(key, _)| *key == label).map(|(_, message)| *message)
}

/// Listener registered with `startup::set_listener`.
pub fn on_milestone(label: &str) {
    if label == READY_MILESTONE {
        finish();
    } else if let Some(message) = lookup(FAILURE_MESSAGES, label) {
        fail(message);
    } else if let Some(message) = lookup(PROGRESS_MESSAGES, label) {
        show(message);
    }
}

fn push(message: &str) {
    let Some(window) = APP.get().and_then(|app| app.get_webview_window(LABEL)) else {
        return;
    };
    let _ = window.eval(format!("window.__splash&&window.__splash.set({message:?})"));
}

fn show(message: &'static str) {
    {
        let Ok(mut state) = STATE.lock() else { return };
        if state.failed || state.ready {
            return;
        }
        state.message = message;
    }
    push(message);
}

/// Keeps the splash visible with a concise message; the first failure wins.
pub fn fail(message: &'static str) {
    {
        let Ok(mut state) = STATE.lock() else { return };
        if state.failed || state.ready {
            return;
        }
        state.failed = true;
        state.message = message;
    }
    eprintln!("[STARTUP] Startup failed: {message}");
    push(message);
    if let Some(window) = APP.get().and_then(|app| app.get_webview_window(LABEL)) {
        // Give the user a title bar and taskbar entry so the window can be closed.
        let _ = window.set_decorations(true);
        let _ = window.set_skip_taskbar(false);
    }
}

fn finish() {
    {
        let Ok(mut state) = STATE.lock() else { return };
        if state.failed || state.ready {
            return;
        }
        state.ready = true;
    }
    push("Ready");
    let Some(app) = APP.get() else { return };
    if let Some(main) = app.get_webview_window("main") {
        for result in [
            main.set_fullscreen(false),
            main.set_decorations(true),
            main.set_skip_taskbar(false),
            main.maximize(),
            main.show(),
            main.set_focus(),
        ] {
            if let Err(error) = result {
                eprintln!("[STARTUP] Could not prepare the main window: {error}");
            }
        }
    }
    if let Some(splash) = app.get_webview_window(LABEL) {
        // destroy() does not raise CloseRequested, which would exit the app.
        let _ = splash.destroy();
    }
}

/// Reports a timeout if startup has neither completed nor failed after `timeout`.
pub fn watch(timeout: Duration) {
    thread::spawn(move || {
        thread::sleep(timeout);
        let pending = STATE.lock().is_ok_and(|state| !state.ready && !state.failed);
        if pending {
            fail(MSG_TIMED_OUT);
        }
    });
}

/// Custom protocol handler serving the splash page and image.
pub fn protocol<R: Runtime>(_context: UriSchemeContext<'_, R>, request: Request<Vec<u8>>) -> Response<Cow<'static, [u8]>> {
    let (content_type, body): (&str, Cow<'static, [u8]>) = match request.uri().path() {
        "/splash.png" => ("image/png", Cow::Borrowed(IMAGE)),
        _ => ("text/html; charset=utf-8", Cow::Borrowed(PAGE.as_bytes())),
    };
    Response::builder()
        .header("Content-Type", content_type)
        .header("Cache-Control", "no-store")
        .body(body)
        .unwrap_or_else(|_| Response::new(Cow::Borrowed(&[][..])))
}

/// Window size: 16:9 to match the artwork, scaled to the primary monitor.
fn window_size(app: &AppHandle) -> (f64, f64) {
    const ASPECT: f64 = 16.0 / 9.0;
    let (mut width, _) = (960.0_f64, 540.0_f64);
    if let Ok(Some(monitor)) = app.primary_monitor() {
        let scale = monitor.scale_factor().max(1.0);
        let monitor_width = f64::from(monitor.size().width) / scale;
        let monitor_height = f64::from(monitor.size().height) / scale;
        width = (monitor_width * 0.5)
            .min(monitor_height * 0.8 * ASPECT)
            .clamp(480.0, 960.0);
    }
    (width, (width / ASPECT).round())
}

/// Creates the borderless splash window. Failure is non-fatal: startup still completes.
pub fn create(app: &AppHandle) -> tauri::Result<()> {
    let _ = APP.set(app.clone());
    let message = STATE.lock().map(|state| state.message).unwrap_or(INITIAL_MESSAGE);
    let url: Url = if cfg!(windows) { "http://splash.localhost/" } else { "splash://localhost/" }
        .parse()
        .map_err(tauri::Error::InvalidUrl)?;
    let (width, height) = window_size(app);
    let builder = WebviewWindowBuilder::new(app, LABEL, WebviewUrl::CustomProtocol(url))
        .title("Scheduling Maestro")
        .inner_size(width, height)
        .decorations(false)
        .resizable(false)
        .skip_taskbar(true)
        .always_on_top(true)
        .center()
        .background_color(Color(0, 0, 0, 255))
        .initialization_script(format!(
            "window.__splash={{m:{message:?},set:function(m){{this.m=m;if(this.apply)this.apply();}}}};"
        ));
    // GTK may grow an undecorated, non-resizable window after it is mapped, leaving an empty strip
    // below the page. Pinning min/max to the exact size and re-applying it keeps the size fixed.
    #[cfg(target_os = "linux")]
    let builder = builder.min_inner_size(width, height).max_inner_size(width, height);
    #[allow(unused_variables)]
    let window = builder.build()?;
    #[cfg(target_os = "linux")]
    {
        let _ = window.set_size(LogicalSize::new(width, height));
        let _ = window.center();
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_startup_milestone_has_a_user_facing_message() {
        assert_eq!(lookup(PROGRESS_MESSAGES, "App launch"), Some("Starting Scheduling Maestro..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "Backend spawn begin"), Some("Launching backend..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "Backend health check begin"), Some("Waiting for backend..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "Database initialization begin"), Some("Initializing database..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "React bootstrap begin"), Some("Loading application modules..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "Session load begin"), Some("Loading session..."));
        assert_eq!(lookup(PROGRESS_MESSAGES, "Dashboard render begin"), Some("Preparing dashboard..."));
    }

    #[test]
    fn unlisted_milestones_do_not_change_the_message() {
        assert_eq!(lookup(PROGRESS_MESSAGES, "Backend spawn complete"), None);
        assert_eq!(lookup(PROGRESS_MESSAGES, READY_MILESTONE), None);
    }

    #[test]
    fn failure_messages_are_concise_and_contain_no_technical_details() {
        for (_, message) in FAILURE_MESSAGES {
            assert!(message.len() < 40 && message.ends_with('.'));
        }
    }
}
