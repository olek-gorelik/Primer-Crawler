#!/bin/bash
# Double-clickable launcher for the Primer Crawler GUI
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

# Use python.org 3.12 (installed) for stable Tk
PYTHON_BIN="/Library/Frameworks/Python.framework/Versions/3.12/bin/python3"
export PYTHONPATH="/Library/Frameworks/Python.framework/Versions/3.12/lib/python3.12/site-packages:${PYTHONPATH}"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "python3.12 not found at $PYTHON_BIN"
  exit 1
fi

# Ensure requests is available
if ! "$PYTHON_BIN" - <<'PY'
import importlib.util, sys
sys.exit(0 if importlib.util.find_spec("requests") else 1)
PY
then
  "$PYTHON_BIN" -m pip install --quiet --upgrade requests || true
fi

exec "$PYTHON_BIN" primer_gui.py
