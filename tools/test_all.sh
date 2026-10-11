#!/usr/bin/env bash
# Export the made-up test project, let the real Excel recalculate it, then check the result.
# Run from Git Bash:  bash tools/test_all.sh
# Works in a short temp folder because Excel 2013 refuses paths longer than 218 characters.
set -e
APP="$(cd "$(dirname "$0")/.." && pwd -W)"
OUT="${TEMP:-/tmp}/boq-test"
mkdir -p "$OUT"
node "$APP/tools/test_export.js" "$APP" "$OUT/out.xlsx"
for f in out out_lists; do
  echo "== $f.xlsx"
  powershell -NoProfile -ExecutionPolicy Bypass -File "$APP/tools/recalc.ps1" -Path "$(cygpath -w "$OUT/$f.xlsx")"
  python -P "$APP/tools/check_export.py" "$OUT/$f.xlsx" "$@"
done
echo "ALL OK — files in $(cygpath -w "$OUT")"
