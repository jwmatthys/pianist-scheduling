use std::{
    error::Error,
    fs,
    io::{self, BufRead, BufReader, Write},
    net::{Ipv4Addr, TcpListener, TcpStream},
    path::PathBuf,
    process::{Child, Command, Stdio},
    sync::Mutex,
    thread,
    time::{Duration, Instant},
};

use tauri::{webview::PageLoadEvent, AppHandle, Manager, RunEvent, State, WebviewWindowBuilder, WindowEvent};

mod splash;
mod startup;

const SERVICE_STARTUP_TIMEOUT: Duration = Duration::from_secs(60);
// Failure watchdog for the whole startup (backend wait plus frontend load).
const STARTUP_WATCHDOG: Duration = Duration::from_secs(90);
const OFFSCREEN_POSITION: f64 = -32000.0;
const BACKEND_EVENT_PREFIX: &str = "STARTUP-EVENT ";
const BACKEND_EVENTS_ENV: &str = "SCHEDULING_MAESTRO_STARTUP_EVENTS";
#[cfg(windows)]
const CREATE_NO_WINDOW: u32 = 0x0800_0000;

struct BackendProcess {
    port: u16,
    child: Mutex<Child>,
}

impl BackendProcess {
    fn start(app: &tauri::AppHandle) -> Result<Self, Box<dyn Error>> {
        let listener = TcpListener::bind(("127.0.0.1", 0))?;
        let port = listener.local_addr()?.port();
        drop(listener);

        let mut command = if cfg!(debug_assertions) {
            let backend_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../backend");
            let venv_python = if cfg!(windows) {
                backend_dir.join(".venv/Scripts/python.exe")
            } else {
                backend_dir.join(".venv/bin/python")
            };
            let mut command = if venv_python.is_file() {
                Command::new(venv_python)
            } else {
                Command::new(if cfg!(windows) { "python" } else { "python3" })
            };
            command
                .arg("desktop_server.py")
                .current_dir(backend_dir);
            command
        } else {
            let executable_name = if cfg!(windows) {
                "pianist-scheduling-api.exe"
            } else {
                "pianist-scheduling-api"
            };
            let executable_dir = std::env::current_exe()?
                .parent()
                .ok_or_else(|| io::Error::other("Application executable has no parent directory"))?
                .to_path_buf();
            let executable = executable_dir.join(executable_name);
            if !executable.is_file() {
                return Err(io::Error::new(
                    io::ErrorKind::NotFound,
                    format!("Bundled scheduling service is missing: {}", executable.display()),
                )
                .into());
            }
            Command::new(executable)
        };

        let data_dir = app.path().app_data_dir()?;
        fs::create_dir_all(&data_dir)?;
        let database_path = data_dir.join("pianist_scheduling.db");
        let stderr = if cfg!(debug_assertions) || startup::enabled() {
            Stdio::inherit()
        } else {
            Stdio::null()
        };
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            command.process_group(0);
        }
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            command.creation_flags(CREATE_NO_WINDOW);
        }
        if startup::enabled() {
            command.env(startup::ENV_FLAG, "1");
        }
        startup::mark("Backend spawn begin", None);
        let mut child = command
            .arg("--port")
            .arg(port.to_string())
            .env("PIANIST_SCHEDULING_DB_PATH", database_path)
            .env(BACKEND_EVENTS_ENV, "1")
            .stdin(Stdio::null())
            .stdout(Stdio::piped())
            .stderr(stderr)
            .spawn()?;
        startup::mark("Backend spawn complete", None);
        let events_thread = child.stdout.take().map(|stdout| thread::spawn(move || forward_backend_events(stdout)));

        startup::mark("Backend health check begin", None);
        if let Err(error) = wait_for_service(&mut child, port) {
            terminate_process_tree(&mut child);
            let _ = child.wait();
            // The pipe closes with the process; join so a final failure event is not lost.
            if let Some(events_thread) = events_thread {
                let _ = events_thread.join();
            }
            return Err(error.into());
        }
        startup::mark("Backend health check success", None);

        Ok(Self {
            port,
            child: Mutex::new(child),
        })
    }

    fn stop(&self) {
        if let Ok(mut child) = self.child.lock() {
            if child.try_wait().ok().flatten().is_none() {
                terminate_process_tree(&mut child);
            }
            let _ = child.wait();
        }
    }
}

impl Drop for BackendProcess {
    fn drop(&mut self) {
        self.stop();
    }
}

#[cfg(unix)]
fn terminate_process_tree(child: &mut Child) {
    let process_group = child.id() as libc::pid_t;
    unsafe {
        libc::kill(-process_group, libc::SIGKILL);
    }
}

#[cfg(windows)]
fn terminate_process_tree(child: &mut Child) {
    let taskkill = std::env::var_os("SystemRoot")
        .map(PathBuf::from)
        .map(|root| root.join("System32/taskkill.exe"))
        .unwrap_or_else(|| PathBuf::from("taskkill.exe"));
    let pid = child.id().to_string();
    use std::os::windows::process::CommandExt;
    let terminated = Command::new(taskkill)
        .creation_flags(CREATE_NO_WINDOW)
        .args(["/PID", &pid, "/T", "/F"])
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .is_ok_and(|status| status.success());
    if !terminated {
        let _ = child.kill();
    }
}

#[cfg(not(any(unix, windows)))]
fn terminate_process_tree(child: &mut Child) {
    let _ = child.kill();
}

fn wait_for_service(child: &mut Child, port: u16) -> io::Result<()> {
    let deadline = Instant::now() + SERVICE_STARTUP_TIMEOUT;
    loop {
        if let Some(status) = child.try_wait()? {
            return Err(io::Error::other(format!(
                "Scheduling service exited before becoming ready ({status})"
            )));
        }
        if service_is_ready(port) {
            return Ok(());
        }
        if Instant::now() >= deadline {
            return Err(io::Error::new(
                io::ErrorKind::TimedOut,
                format!(
                    "Scheduling service did not become ready within {} seconds",
                    SERVICE_STARTUP_TIMEOUT.as_secs()
                ),
            ));
        }
        thread::sleep(Duration::from_millis(150));
    }
}

fn service_is_ready(port: u16) -> bool {
    let address = (Ipv4Addr::LOCALHOST, port);
    let Ok(mut stream) = TcpStream::connect_timeout(&address.into(), Duration::from_millis(250))
    else {
        return false;
    };
    let _ = stream.set_read_timeout(Some(Duration::from_millis(500)));
    if write!(
        stream,
        "GET /api/health HTTP/1.1\r\nHost: 127.0.0.1:{port}\r\nConnection: close\r\n\r\n"
    )
    .is_err()
    {
        return false;
    }

    let mut status_line = String::new();
    let Ok(read) = BufReader::new(stream).read_line(&mut status_line) else {
        return false;
    };
    read > 0 && (status_line.starts_with("HTTP/1.1 200 ") || status_line.starts_with("HTTP/1.0 200 "))
}

/// Reads fixed `STARTUP-EVENT <token>` lines from the backend's stdout until it closes.
/// Unknown tokens and any other output are ignored (and never logged).
fn forward_backend_events(stdout: impl io::Read) {
    for line in BufReader::new(stdout).lines().map_while(Result::ok) {
        let Some(token) = line.strip_prefix(BACKEND_EVENT_PREFIX) else {
            continue;
        };
        if let Some((_, label)) = startup::BACKEND_EVENTS.iter().find(|(known, _)| *known == token.trim()) {
            startup::mark(label, None);
        }
    }
}

#[tauri::command]
fn get_api_base(backend: State<'_, BackendProcess>) -> String {
    format!("http://127.0.0.1:{}", backend.port)
}

/// Receives frontend startup milestones; only allowlisted labels are accepted.
/// Async so window operations triggered by the splash never run on the main thread.
#[tauri::command]
async fn startup_mark(label: String) {
    if startup::FRONTEND_MILESTONES.contains(&label.as_str()) {
        startup::mark(&label, None);
        if label == "First usable UI rendered" {
            startup::log_summary();
        }
    }
}

/// Creates the main window from tauri.conf.json. WebView2 suspends pages in hidden windows,
/// so no startup milestone would ever fire; instead the window is created "visible" but
/// off-screen and out of the taskbar, and the splash moves it on-screen when ready.
fn create_main_window(app: &AppHandle) -> tauri::Result<()> {
    for config in app.config().app.windows.iter().filter(|window| window.label == "main") {
        WebviewWindowBuilder::from_config(app, config)?
            .visible(true)
            .decorations(true)
            .maximizable(true)
            .minimizable(true)
            .closable(true)
            .skip_taskbar(true)
            .position(OFFSCREEN_POSITION, OFFSCREEN_POSITION)
            .build()?;
    }
    Ok(())
}

fn start_backend_and_main_window(app: AppHandle) {
    match BackendProcess::start(&app) {
        Ok(backend) => {
            app.manage(backend);
            startup::mark("API endpoint published", None);
            if let Err(error) = create_main_window(&app) {
                eprintln!("[STARTUP] Main window creation failed: {error}");
                splash::fail(splash::MSG_WINDOW_FAILED);
            }
        }
        Err(error) => {
            eprintln!("[STARTUP] Backend failed to start: {error}");
            let timed_out = error
                .downcast_ref::<io::Error>()
                .is_some_and(|error| error.kind() == io::ErrorKind::TimedOut);
            // No-op when the backend already reported a more specific failure.
            splash::fail(if timed_out { splash::MSG_TIMED_OUT } else { splash::MSG_BACKEND_FAILED });
        }
    }
}

pub fn run() -> Result<(), Box<dyn Error>> {
    startup::init();
    startup::set_listener(splash::on_milestone);
    startup::mark("App launch", None);
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .register_uri_scheme_protocol(splash::SCHEME, splash::protocol)
        .setup(|app| {
            if let Err(error) = splash::create(app.handle()) {
                eprintln!("[STARTUP] Splash window creation failed: {error}");
            }
            startup::mark("Setup begin", None);
            splash::watch(STARTUP_WATCHDOG);
            // Off the setup thread so the event loop can paint the splash while the backend starts.
            let handle = app.handle().clone();
            thread::spawn(move || start_backend_and_main_window(handle));
            Ok(())
        })
        .on_page_load(|webview, payload| {
            if webview.label() != "main" {
                return;
            }
            match payload.event() {
                PageLoadEvent::Started => startup::mark("Webview page load started", None),
                PageLoadEvent::Finished => startup::mark("Webview page load finished", None),
            };
        })
        .invoke_handler(tauri::generate_handler![get_api_base, startup_mark])
        .build(tauri::generate_context!())?;
    startup::mark("Tauri application built", None);

    app.run(|app_handle, event| match event {
        RunEvent::Ready => {
            let detail = match app_handle
                .get_webview_window("splash")
                .and_then(|window| window.is_visible().ok())
            {
                Some(true) => "splash window visible",
                Some(false) => "splash window hidden",
                None => "splash window unavailable",
            };
            startup::mark("Event loop ready", Some(detail));
        }
        RunEvent::WindowEvent {
            event: WindowEvent::CloseRequested { .. },
            ..
        } => app_handle.exit(0),
        RunEvent::Exit => {
            if let Some(backend) = app_handle.try_state::<BackendProcess>() {
                backend.stop();
            }
        }
        _ => {}
    });
    Ok(())
}
