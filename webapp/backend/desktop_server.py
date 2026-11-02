"""Entrypoint bundled with PyInstaller for the desktop application."""

import argparse

from app import startup_profile

startup_profile.mark("Backend process entry")

import uvicorn

startup_profile.mark("Server libraries imported")

from app.main import app

startup_profile.mark("Application import complete")


def main():
    parser = argparse.ArgumentParser(description="Pianist Scheduling desktop API")
    parser.add_argument("--port", type=int, required=True)
    args = parser.parse_args()
    startup_profile.mark("Arguments parsed")
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    main()