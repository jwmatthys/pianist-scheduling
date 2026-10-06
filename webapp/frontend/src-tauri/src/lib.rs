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

use tauri::{Manager, RunEvent, State, WindowEvent};

const SERVICE_STARTUP_TIMEOUT: Duration = Duration::from_secs(60);

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
        let stderr = if cfg!(debug_assertions) {
            Stdio::inherit()
        } else {
            Stdio::null()
        };
        #[cfg(unix)]
        {
            use std::os::unix::process::CommandExt;
            command.process_group(0);
        }
        let mut child = command
            .arg("--port")
            .arg(port.to_string())
            .env("PIANIST_SCHEDULING_DB_PATH", database_path)
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(stderr)
            .spawn()?;

        if let Err(error) = wait_for_service(&mut child, port) {
            terminate_process_tree(&mut child);
            let _ = child.wait();
            return Err(error.into());
        }

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
    let terminated = Command::new(taskkill)
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

#[tauri::command]
fn get_api_base(backend: State<'_, BackendProcess>) -> String {
    format!("http://127.0.0.1:{}", backend.port)
}

pub fn run() -> Result<(), Box<dyn Error>> {
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .setup(|app| {
            app.manage(BackendProcess::start(&app.handle())?);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![get_api_base])
        .build(tauri::generate_context!())?;

    app.run(|app_handle, event| match event {
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
