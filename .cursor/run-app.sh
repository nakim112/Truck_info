#!/usr/bin/env bash
#
# Convenience launcher for the Truck Info GUI on the Cloud Agent VM.
# Ensures a virtual display is available, then runs the app from the venv.
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")/.."

DISPLAY_NUM="${TRUCK_INFO_DISPLAY:-99}"
VENV_DIR="${TRUCK_INFO_VENV:-$HOME/.venvs/truckinfo}"

if ! xdpyinfo -display ":${DISPLAY_NUM}" >/dev/null 2>&1; then
  bash .cursor/start.sh
fi

exec env DISPLAY=":${DISPLAY_NUM}" "${VENV_DIR}/bin/python" truck_info_ui_airport.py "$@"
