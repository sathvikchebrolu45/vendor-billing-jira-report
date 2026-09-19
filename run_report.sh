#!/usr/bin/env bash
# One-command pipeline: reads the Excel timesheet + JIRA (via twg CLI) and
# writes a fresh CSV report into reports/ within this repo.
#
# Date range: static start of Aug 1 (current year) -> today (both defaults
# built into generate_vendor_jira_report.py). Override with env vars:
#   START_DATE=2026-08-01 END_DATE=2026-09-15 ./run_report.sh
#
# Usage:
#   ./run_report.sh
#   ./run_report.sh --project MID1 --excel "/path/to/Maersk MIDAS.xlsx"

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

REPORTS_DIR="$SCRIPT_DIR/reports"
mkdir -p "$REPORTS_DIR"

TIMESTAMP="$(date +%Y%m%d_%H%M%S)"
OUTPUT_FILE="$REPORTS_DIR/vendor_jira_report_${TIMESTAMP}.csv"

ARGS=(--output "$OUTPUT_FILE")

if [[ -n "${START_DATE:-}" ]]; then
  ARGS+=(--start "$START_DATE")
fi
if [[ -n "${END_DATE:-}" ]]; then
  ARGS+=(--end "$END_DATE")
fi

echo "Running report pipeline..."
python3 generate_vendor_jira_report.py "${ARGS[@]}" "$@"

echo ""
echo "Report written to: $OUTPUT_FILE"
