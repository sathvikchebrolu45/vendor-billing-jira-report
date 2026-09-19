#!/usr/bin/env bash
# One-command pipeline: reads the Excel timesheet + JIRA (via twg CLI),
# writes a fresh CSV report into reports/ within this repo, and commits +
# pushes it to the git repo automatically.
#
# Date range: static start of Aug 1 (current year) -> today (both defaults
# built into generate_vendor_jira_report.py). Override with env vars:
#   START_DATE=2026-08-01 END_DATE=2026-09-15 ./run_report.sh
#
# Set NO_GIT=1 to skip the commit/push step (report is still written locally).
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

if [[ "${NO_GIT:-0}" != "1" ]]; then
  git add "$OUTPUT_FILE"
  if git diff --cached --quiet; then
    echo "Nothing new to commit."
  else
    git commit -m "Add vendor/JIRA report: $(basename "$OUTPUT_FILE")" -q
    git push -q
    echo "Committed and pushed to git."
  fi
fi
