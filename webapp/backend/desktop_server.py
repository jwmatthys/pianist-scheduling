"""Entrypoint bundled with PyInstaller for the desktop application."""

import argparse

import uvicorn

from app.main import app


def main():
    parser = argparse.ArgumentParser(description="Pianist Scheduling desktop API")
    parser.add_argument("--port", type=int, required=True)
    args = parser.parse_args()
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    main()