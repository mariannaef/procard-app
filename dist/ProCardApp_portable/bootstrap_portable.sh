#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_URL="http://127.0.0.1:8501"
VENV_DIR="$SCRIPT_DIR/.venv"
VENV_PYTHON="$VENV_DIR/bin/python"
REQUIREMENTS_PATH="$SCRIPT_DIR/requirements.txt"
LAUNCHER_PATH="$SCRIPT_DIR/launcher.py"
REQUIREMENTS_HASH_PATH="$VENV_DIR/.requirements.sha256"
PYTHON_DOWNLOAD_URL="https://www.python.org/downloads/macos/"

test_app_running() {
    curl --silent --fail --max-time 2 "$APP_URL" >/dev/null 2>&1
}

find_python_command() {
    local version_check='import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'

    if command -v python3 >/dev/null 2>&1; then
        if python3 -c "$version_check" >/dev/null 2>&1; then
            echo "python3"
            return 0
        fi
    fi

    if command -v python >/dev/null 2>&1; then
        if python -c "$version_check" >/dev/null 2>&1; then
            echo "python"
            return 0
        fi
    fi

    return 1
}

show_python_install_prompt() {
    echo
    echo "Python 3.10 or newer is required to run ProCard Reconciliation App."
    echo "Download the macOS installer from:"
    echo "$PYTHON_DOWNLOAD_URL"
    echo
    read -r -p "Open the Python download page now? (Y/N) " response
    if [[ "$response" =~ ^[Yy]([Ee][Ss])?$ ]]; then
        open "$PYTHON_DOWNLOAD_URL"
    fi
}

get_requirements_hash() {
    shasum -a 256 "$REQUIREMENTS_PATH" | awk '{print $1}'
}

if test_app_running; then
    open "$APP_URL"
    exit 0
fi

if [[ ! -f "$REQUIREMENTS_PATH" ]]; then
    echo "requirements.txt was not found in $SCRIPT_DIR" >&2
    exit 1
fi

if [[ ! -f "$LAUNCHER_PATH" ]]; then
    echo "launcher.py was not found in $SCRIPT_DIR" >&2
    exit 1
fi

if ! PYTHON_COMMAND="$(find_python_command)"; then
    show_python_install_prompt
    echo "Python 3.10 or newer is required. Install it, then run launch_app.command again." >&2
    exit 1
fi

if [[ ! -x "$VENV_PYTHON" ]]; then
    echo "Creating local Python environment..."
    "$PYTHON_COMMAND" -m venv "$VENV_DIR"
fi

EXPECTED_HASH="$(get_requirements_hash)"
INSTALLED_HASH=""
if [[ -f "$REQUIREMENTS_HASH_PATH" ]]; then
    INSTALLED_HASH="$(tr -d '[:space:]' < "$REQUIREMENTS_HASH_PATH")"
fi

if [[ "$INSTALLED_HASH" != "$EXPECTED_HASH" ]]; then
    echo "Installing app dependencies (first run may take a few minutes)..."
    "$VENV_PYTHON" -m pip install --upgrade pip
    "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS_PATH"
    printf "%s" "$EXPECTED_HASH" > "$REQUIREMENTS_HASH_PATH"
fi

echo "Starting ProCard Reconciliation App..."
exec "$VENV_PYTHON" "$LAUNCHER_PATH"