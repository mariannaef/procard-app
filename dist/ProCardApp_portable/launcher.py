import os
import socket
import sys
import threading
import time
import urllib.request
import webbrowser

from streamlit.web import cli as stcli


APP_HOST = "127.0.0.1"
APP_PORT = 8501
APP_URL = f"http://{APP_HOST}:{APP_PORT}"


def _resource_path(relative_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(getattr(sys, "_MEIPASS"), relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def _url_is_ready(url: str) -> bool:
    try:
        with urllib.request.urlopen(url, timeout=2):
            return True
    except Exception:
        return False


def _port_is_in_use(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(1)
        return sock.connect_ex((host, port)) == 0


def _open_browser_when_ready(url: str, timeout_seconds: int = 60) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if _url_is_ready(url):
            webbrowser.open(url, new=2)
            return
        time.sleep(1)
    print(
        f"\nERROR: The app server did not become ready at {url} within {timeout_seconds} seconds.\n"
        "Possible causes:\n"
        "  - A Python import error in app.py or processor.py (check output above)\n"
        "  - Port 8501 is blocked by a firewall or antivirus\n"
        "  - The app crashed during startup\n",
        file=sys.stderr,
        flush=True,
    )


if __name__ == "__main__":
    app_script = _resource_path("app.py")

    if _url_is_ready(APP_URL):
        webbrowser.open(APP_URL, new=2)
        raise SystemExit(0)

    if _port_is_in_use(APP_HOST, APP_PORT):
        raise SystemExit(f"Port {APP_PORT} is already in use. Close the existing ProCard app or free the port, then try again.")

    threading.Thread(
        target=_open_browser_when_ready,
        args=(APP_URL,),
        daemon=True,
    ).start()
    sys.argv = [
        "streamlit",
        "run",
        app_script,
        f"--server.address={APP_HOST}",
        f"--server.port={APP_PORT}",
        "--server.headless=true",
        "--global.developmentMode=false",
    ]
    raise SystemExit(stcli.main())
