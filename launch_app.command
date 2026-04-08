#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
BOOTSTRAP_SCRIPT="$SCRIPT_DIR/bootstrap_portable.sh"

if [[ ! -f "$BOOTSTRAP_SCRIPT" ]]; then
    echo
    echo "ERROR: bootstrap_portable.sh not found."
    echo "Expected: $BOOTSTRAP_SCRIPT"
    echo "Keep the launcher files together after extracting the zip."
    echo
    read -r -p "Press Enter to close..." _
    exit 1
fi

chmod +x "$BOOTSTRAP_SCRIPT" "$0" 2>/dev/null || true
"$BOOTSTRAP_SCRIPT"
EXIT_CODE=$?

if [[ $EXIT_CODE -ne 0 ]]; then
    echo
    echo "ProCard App could not be started."
    read -r -p "Press Enter to close..." _
fi

exit $EXIT_CODE