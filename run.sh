#!/usr/bin/env bash
# SlideCraft Lite launcher (macOS / Linux)
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
    echo ">> First run: creating .venv"
    python3 -m venv .venv
fi
# shellcheck disable=SC1091
source .venv/bin/activate
if ! python -c "import flask" 2>/dev/null; then
    echo ">> Installing dependencies (one-time, ~30 s)"
    pip install --upgrade pip
    pip install -r requirements.txt
fi
if ! command -v libreoffice >/dev/null 2>&1 && ! command -v soffice >/dev/null 2>&1; then
    if [ ! -e "/Applications/LibreOffice.app/Contents/MacOS/soffice" ]; then
        echo "!! LibreOffice not found — PPTX conversion will use a low-fidelity fallback."
        echo "!! Install with:  brew install --cask libreoffice    (macOS)"
        echo "!!                sudo apt install libreoffice       (Debian/Ubuntu)"
    fi
fi
exec python app.py
