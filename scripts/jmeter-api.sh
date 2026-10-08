#!/usr/bin/env bash
# RHEL 9.x / Bash 5.x. Always pass arguments as an array (never use eval).
set -euo pipefail
umask 077

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
python_bin="${PYTHON_BIN:-python3}"
if ! command -v "$python_bin" >/dev/null 2>&1; then
    printf 'ERROR: Python 3.9+ is required. Install python3 or set PYTHON_BIN.\n' >&2
    exit 2
fi
exec "$python_bin" "$script_dir/../jmeter_api.py" "$@"
