#!/usr/bin/env bash
#
# Cloud Agent install script for the INC4 HUB Truck Info app.
#
# The app is a Windows-oriented Tkinter desktop GUI (built with PyInstaller for
# Windows). This script prepares it to RUN and be verified on the Linux Cloud
# Agent VM: it installs Tkinter, Korean fonts, and a headless X stack, then
# creates a virtualenv with the Linux-compatible subset of requirements.txt.
#
# The virtualenv lives OUTSIDE the git working tree (default:
# $HOME/.venvs/truckinfo) so it survives a fresh checkout when running from a
# prebuilt environment build that skips install on boot.
#
# Idempotent: safe to run repeatedly.
set -euo pipefail

# Always operate from the repository root regardless of invocation cwd.
cd "$(dirname "${BASH_SOURCE[0]}")/.."

VENV_DIR="${TRUCK_INFO_VENV:-$HOME/.venvs/truckinfo}"

echo "[install] Installing system packages (Tkinter, Korean fonts, headless X)..."
export DEBIAN_FRONTEND=noninteractive
sudo apt-get update -qq
sudo apt-get install -y --no-install-recommends \
  python3-tk \
  python3-venv \
  fonts-nanum \
  fonts-nanum-coding \
  fonts-noto-cjk \
  xvfb \
  x11-utils \
  imagemagick \
  scrot

echo "[install] Creating Python virtualenv at ${VENV_DIR}..."
mkdir -p "$(dirname "${VENV_DIR}")"
if [ ! -x "${VENV_DIR}/bin/python" ]; then
  python3 -m venv "${VENV_DIR}"
fi
# shellcheck disable=SC1091
. "${VENV_DIR}/bin/activate"
python -m pip install --upgrade pip

echo "[install] Installing Python dependencies (excluding Windows-only packages)..."
# pywin32 / pywin32-ctypes are Windows-only and cannot be installed on Linux.
# They are only used behind runtime guards for Excel COM automation on Windows,
# so the app runs fine without them here.
grep -viE '^(pywin32|pywin32-ctypes)([=<>!~ ]|$)' requirements.txt > /tmp/requirements-linux.txt
python -m pip install -r /tmp/requirements-linux.txt

echo "[install] Verifying key imports..."
python - <<'PY'
import tkinter, pandas, openpyxl, PIL, selenium
import config, utils, widgets, ui_palette, api_service, excel_service
print("[install] Imports OK: tkinter", tkinter.TkVersion,
      "| pandas", pandas.__version__)
PY

echo "[install] Done. Virtualenv: ${VENV_DIR}"
