#!/bin/bash
# Double-click this file on macOS to launch the In-Kind IIF Generator GUI.
cd "$(dirname "$0")"
# Ensure openpyxl is available
python3 -c "import openpyxl" 2>/dev/null || pip3 install --user openpyxl
exec python3 "$(dirname "$0")/in_kind_iif_generator.py" --gui
