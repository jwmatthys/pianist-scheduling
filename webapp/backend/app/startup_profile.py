"""Opt-in startup timing for the desktop backend.

Enabled only when SCHEDULING_MAESTRO_STARTUP_PROFILE=1 (Tauri sets it for the
sidecar while startup profiling is active). Disabled, mark() returns immediately.
Only fixed phase labels and elapsed milliseconds are recorded; never pass user
or session data as a label. Each label is recorded and printed once (first
occurrence). To remove the instrumentation, delete this module
and the startup_profile.mark(...) calls.
"""

import os
import sys
import time
from contextlib import nullcontext

ENV_FLAG = "SCHEDULING_MAESTRO_STARTUP_PROFILE"

_START = time.perf_counter()
_ENABLED = os.environ.get(ENV_FLAG) == "1"
_marks: dict[str, float] = {}
_blocks: list[tuple[str, float]] = []
# (module name, cumulative ms, self ms) for every module first loaded while the hook is installed.
_imports: list[tuple[str, float, float]] = []
_original_find_and_load = None

# Third-party packages reported by name whether or not startup imported them.
TRACKED_PACKAGES = ("fastapi", "starlette", "pydantic", "sqlalchemy", "uvicorn", "openpyxl", "pandas", "xlrd", "reportlab")


def enabled() -> bool:
    return _ENABLED


EVENTS_ENV_FLAG = "SCHEDULING_MAESTRO_STARTUP_EVENTS"


def emit_event(name: str) -> None:
    """Report a fixed startup milestone token to the desktop shell on stdout.

    The shell sets SCHEDULING_MAESTRO_STARTUP_EVENTS=1 and reads these lines to drive
    the splash screen; otherwise nothing is printed. Only fixed tokens are allowed.
    """
    if os.environ.get(EVENTS_ENV_FLAG) != "1" or sys.stdout is None:
        return
    try:
        print(f"STARTUP-EVENT {name}", flush=True)
    except OSError:
        pass


def mark(label: str) -> None:
    if not _ENABLED or label in _marks:
        return
    elapsed_ms = (time.perf_counter() - _START) * 1000
    _marks[label] = elapsed_ms
    print(f"[STARTUP backend +{elapsed_ms:.0f}ms] {label}", file=sys.stderr, flush=True)


class _Measure:
    def __init__(self, label: str):
        self.label = label

    def __enter__(self):
        self.begin = time.perf_counter()

    def __exit__(self, *exc):
        _blocks.append((self.label, (time.perf_counter() - self.begin) * 1000))
        return False


def measure(label: str):
    """Time a block (e.g. `with measure("Created engine"):`); a no-op when profiling is off."""
    return _Measure(label) if _ENABLED else nullcontext()


def _install_import_timing() -> None:
    """Wrap importlib's module loader so every first-time import records cumulative and self time."""
    global _original_find_and_load
    import importlib._bootstrap as bootstrap

    original = bootstrap._find_and_load
    child_totals: list[float] = []

    def timed_find_and_load(name, import_):
        child_totals.append(0.0)
        begin = time.perf_counter()
        try:
            return original(name, import_)
        finally:
            cumulative = (time.perf_counter() - begin) * 1000
            children = child_totals.pop()
            if child_totals:
                child_totals[-1] += cumulative
            _imports.append((name, cumulative, cumulative - children))

    _original_find_and_load = original
    bootstrap._find_and_load = timed_find_and_load


def _uninstall_import_timing() -> None:
    global _original_find_and_load
    if _original_find_and_load is not None:
        import importlib._bootstrap as bootstrap

        bootstrap._find_and_load = _original_find_and_load
        _original_find_and_load = None


if _ENABLED:
    _install_import_timing()


def span_ms(begin: str, end: str) -> float | None:
    if begin in _marks and end in _marks:
        return _marks[end] - _marks[begin]
    return None


def _emit(line: str) -> None:
    print(line, file=sys.stderr, flush=True)


def _log_import_report() -> None:
    if not _imports:
        return
    total_self = sum(item[2] for item in _imports)
    by_owner: dict[str, list[float]] = {}
    for name, _cumulative, self_ms in _imports:
        # Application modules are reported individually; everything else by top-level package.
        owner = name if name.startswith("app.") else name.split(".")[0]
        entry = by_owner.setdefault(owner, [0.0, 0])
        entry[0] += self_ms
        entry[1] += 1
    ranked = sorted(by_owner.items(), key=lambda item: item[1][0], reverse=True)

    _emit(f"[STARTUP backend] Import-time breakdown ({len(_imports)} modules loaded, {total_self:.0f} ms self time total)")
    _emit("[STARTUP backend] Top startup contributors (self time: module body, class construction, excluding nested imports)")
    for rank, (owner, (self_ms, count)) in enumerate(ranked[:20], start=1):
        _emit(f"[STARTUP backend]   {rank:>2}. {owner} ({count} modules) ... {self_ms:.0f} ms")

    cumulative_app = sorted(
        ((name, cum, own) for name, cum, own in _imports if name.startswith("app.")),
        key=lambda item: item[1],
        reverse=True,
    )
    _emit("[STARTUP backend] Application modules by cumulative import time (includes what each module imports)")
    for name, cumulative, own in cumulative_app[:20]:
        _emit(f"[STARTUP backend]   {name} ... {cumulative:.0f} ms cumulative, {own:.0f} ms self")

    _emit("[STARTUP backend] Tracked third-party packages (self time summed over the package)")
    for package in TRACKED_PACKAGES:
        entry = by_owner.get(package)
        if entry is None:
            _emit(f"[STARTUP backend]   {package}: not imported during startup")
        else:
            _emit(f"[STARTUP backend]   {package}: {entry[0]:.0f} ms ({entry[1]} modules)")


def log_summary() -> None:
    """Log measured spans between recorded marks; spans with a missing mark are omitted."""
    if not _ENABLED:
        return
    spans = (
        ("Backend module imports", "Backend process entry", "Application import complete"),
        ("Database initialization", "Database initialization begin", "Database initialization complete"),
        ("Migration check", "Migration check begin", "Migration check complete"),
        ("Backend total (process entry to FastAPI startup)", "Backend process entry", "FastAPI startup complete"),
    )
    for name, begin, end in spans:
        value = span_ms(begin, end)
        if value is not None:
            _emit(f"[STARTUP backend] {name}: {value:.0f} ms")
    for label, value in _blocks:
        _emit(f"[STARTUP backend] {label}: {value:.1f} ms")
    _uninstall_import_timing()
    _log_import_report()
