#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    if let Err(error) = music_program_scheduler_lib::run() {
        eprintln!("Unable to start Music Program Scheduler: {error}");
        std::process::exit(1);
    }
}
