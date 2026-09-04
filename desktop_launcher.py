import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path


def _app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _find_free_port(start: int = 8501, end: int = 8599) -> int:
    for port in range(start, end + 1):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError("No free local port found for Streamlit.")


def _browser_candidates() -> list[Path]:
    candidates: list[Path] = []

    for env_name in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env_name)
        if not base:
            continue
        candidates.extend(
            [
                Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe",
                Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe",
            ]
        )

    return candidates


def _wait_for_streamlit(port: int, timeout: float = 20.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            if sock.connect_ex(("127.0.0.1", port)) == 0:
                return
        time.sleep(0.25)


def _open_browser(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    _wait_for_streamlit(port)

    for browser_path in _browser_candidates():
        if browser_path.exists():
            # A separate profile forces a dedicated browser process that can
            # be monitored instead of reusing a normal Edge/Chrome process.
            profile_dir = Path(tempfile.mkdtemp(prefix="webfleet-tools-browser-"))
            try:
                browser_process = subprocess.Popen(
                    [
                        str(browser_path),
                        f"--app={url}",
                        f"--user-data-dir={profile_dir}",
                        "--new-window",
                        "--guest",
                        "--no-first-run",
                        "--no-default-browser-check",
                        "--disable-background-mode",
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except OSError:
                shutil.rmtree(profile_dir, ignore_errors=True)
                continue

            # Closing the app window ends its browser process. Streamlit has no
            # browser-close callback, so terminate the packaged process here;
            # Windows will then release its lock on the executable.
            browser_process.wait()
            shutil.rmtree(profile_dir, ignore_errors=True)
            os._exit(0)

    webbrowser.open(url)


def main() -> None:
    root = _app_root()
    app_path = root / "app.py"
    if not app_path.exists():
        raise FileNotFoundError(f"Cannot find app.py at {app_path}")

    os.chdir(root)
    sys.path.insert(0, str(root))
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    port = _find_free_port()
    threading.Thread(target=_open_browser, args=(port,), daemon=True).start()

    from streamlit.web import cli as streamlit_cli

    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    streamlit_cli.main()


if __name__ == "__main__":
    main()
