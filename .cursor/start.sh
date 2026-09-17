#!/usr/bin/env bash
#
# Cloud Agent start script: bring up a persistent virtual X display so the
# Tkinter GUI can run on the headless VM. Idempotent and safe across reboots.
set -euo pipefail

DISPLAY_NUM="${TRUCK_INFO_DISPLAY:-99}"

if xdpyinfo -display ":${DISPLAY_NUM}" >/dev/null 2>&1; then
  echo "[start] Xvfb display :${DISPLAY_NUM} already running."
else
  echo "[start] Launching Xvfb on :${DISPLAY_NUM}..."
  Xvfb ":${DISPLAY_NUM}" -screen 0 1280x800x24 -ac +extension GLX +render -noreset \
    >/tmp/xvfb.log 2>&1 &
  for _ in $(seq 1 20); do
    if xdpyinfo -display ":${DISPLAY_NUM}" >/dev/null 2>&1; then
      break
    fi
    sleep 0.5
  done
fi

if xdpyinfo -display ":${DISPLAY_NUM}" >/dev/null 2>&1; then
  echo "[start] Virtual display :${DISPLAY_NUM} is ready."
  echo "[start] Run the GUI with:"
  echo "[start]   DISPLAY=:${DISPLAY_NUM} .venv/bin/python truck_info_ui_airport.py"
  echo "[start] Or use the helper: DISPLAY=:${DISPLAY_NUM} bash .cursor/run-app.sh"
else
  echo "[start] WARNING: Xvfb failed to start; see /tmp/xvfb.log" >&2
fi
