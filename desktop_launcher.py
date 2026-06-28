import os
import socket
import sys
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


def _open_browser(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    time.sleep(2.5)
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
