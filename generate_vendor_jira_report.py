#!/usr/bin/env python3
"""
generate_vendor_jira_report.py

One-shot script: reads the Maersk MIDAS timesheet (Excel), pulls JIRA work
items for a project/date window via the `twg` CLI, reconciles billed hours
against JIRA story delivery per vendor, and writes a single report (XLSX +
a printed Markdown table). No M365 Copilot agent required.

Prerequisites:
    - `twg` CLI installed and authenticated (run `twg doctor` to confirm).
    - openpyxl installed (pip install openpyxl).

Usage:
    python3 generate_vendor_jira_report.py \
        --excel "/Users/you/Downloads/Maersk MIDAS.xlsx" \
        --project MID1 --start 2026-08-01 --end 2026-09-30

    # Output defaults to vendor_jira_report_<YYYYMMDD_HHMMSS>.csv in the
    # current directory; override with --output.

Conversion constant: 1 story point = 8 hours (change STORY_POINT_HOURS below
if your team uses a different convention).
"""
import argparse
import csv
import json
import os
import shutil
import subprocess
import sys
from collections import defaultdict
from datetime import date, datetime

from openpyxl import load_workbook

STORY_POINT_HOURS = 8


# ---------------------------------------------------------------------------
# Excel timesheet parsing (Maersk MIDAS.xlsx, sheet "Plan")
# ---------------------------------------------------------------------------
def parse_billed_hours(excel_path: str, start: date, end: date) -> dict:
    """Return {vendor_name: {'partner': str, 'billed': float, 'by_month': {label: hrs}}}.

    Sums BLOCK_1 ("Billable" table, rows 7..PI-1) and BLOCK_2 (the second table
    marked "PI" in column A) together per vendor, for every weekly column whose
    row-4 start date falls within [start, end] inclusive.
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb["Plan"]

    pi_rows = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "PI"]
    if not pi_rows:
        sys.exit('ERROR: could not find the "PI" marker row in column A of sheet "Plan".')
    pi = pi_rows[0]

    # Locate week columns whose row-4 start date falls in [start, end], and
    # remember which calendar month each belongs to (for the Aug/Sep style
    # per-month breakdown in the final report).
    week_cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(4, c).value
        if hasattr(v, "year") and start <= v.date() <= end:
            week_cols.append((c, v.strftime("%b %Y")))
    if not week_cols:
        sys.exit(f"ERROR: no weekly columns found with a row-4 start date between {start} and {end}.")

    vendors = {}
    for row_range in (range(7, pi), range(pi + 2, ws.max_row + 1)):
        for r in row_range:
            name = ws.cell(r, 5).value
            if not name:
                break
            name = str(name).strip()
            partner = (ws.cell(r, 2).value or "").strip()
            d = vendors.setdefault(name, {"partner": partner, "billed": 0.0, "by_month": defaultdict(float)})
            if not d["partner"] and partner:
                d["partner"] = partner
            for c, month_label in week_cols:
                hrs = float(ws.cell(r, c).value or 0)
                d["billed"] += hrs
                d["by_month"][month_label] += hrs

    month_labels = sorted({label for _, label in week_cols}, key=lambda m: datetime.strptime(m, "%b %Y"))
    return vendors, month_labels


# ---------------------------------------------------------------------------
# JIRA export via twg CLI (same approach as fetch_jira_export.py)
# ---------------------------------------------------------------------------
def find_twg() -> str:
    path = shutil.which("twg")
    if path:
        return path
    default = os.path.expanduser("~/.local/bin/twg")
    if os.path.exists(default):
        return default
    sys.exit(
        "ERROR: twg CLI not found. Install it "
        "(curl -fsSL https://teamwork-graph.atlassian.com/cli/install | bash) "
        "or open a new terminal so PATH picks it up, then run `twg doctor`."
    )


def run_twg(twg_path: str, args: list[str]) -> str:
    result = subprocess.run([twg_path] + args, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"ERROR running: twg {' '.join(args)}", file=sys.stderr)
        print(result.stderr, file=sys.stderr)
        sys.exit(1)
    return result.stdout


def discover_story_points_field(twg_path: str) -> str | None:
    """Prefer an EXACT 'Story Points' field name match over similarly-named fields
    (e.g. 'Story point estimate', 'Story Points (custom)') which many JIRA
    instances also have."""
    out = run_twg(twg_path, ["api", "jira:/rest/api/2/field"])
    try:
        fields = json.loads(out)
    except json.JSONDecodeError:
        print("WARNING: could not parse field list; Story Points will be blank.", file=sys.stderr)
        return None
    exact_match, fallback_match = None, None
    for f in fields:
        name = (f.get("name") or "").strip().lower()
        if name == "story points":
            return f.get("id")
        if name in ("story point", "story point estimate") and fallback_match is None:
            fallback_match = f.get("id")
    return exact_match or fallback_match


def fetch_all_issues(twg_path: str, jql: str, fields: list[str], page_size: int = 100):
    """Paginate `twg jira workitem query`. --output-summary none forces the full
    JSON payload onto stdout - without it, twg truncates to a compact field
    preset when run non-interactively, silently dropping custom fields."""
    issues = []
    cursor = None
    while True:
        args = [
            "jira", "workitem", "query",
            "--jql", jql,
            "--fields", ",".join(fields),
            "--first", str(page_size),
            "--output", "json",
            "--output-summary", "none",
        ]
        if cursor:
            args += ["--after", cursor]
        out = run_twg(twg_path, args)
        data = json.loads(out)
        batch = data.get("data", {}).get("issues", [])
        issues.extend(batch)
        print(f"  fetched {len(issues)} JIRA issues so far...", file=sys.stderr)
        page_info = data.get("pageInfo", {})
        cursor = page_info.get("nextCursor")
        if not page_info.get("hasNextPage") or not batch or not cursor:
            break
    return issues


def fetch_jira_issues(project: str, start: date, end: date, story_points_field: str | None):
    twg_path = find_twg()

    if not story_points_field:
        print("Auto-detecting Story Points custom field id via twg...", file=sys.stderr)
        story_points_field = discover_story_points_field(twg_path)
        if not story_points_field:
            print("WARNING: could not auto-detect Story Points field; it will be blank.", file=sys.stderr)

    # No "assignee =" clause on purpose - pulls every work item in the project
    # for the period, for ALL vendors at once; grouped by vendor below.
    jql = (
        f'project = "{project}" AND '
        f'((status = Done AND resolutiondate >= "{start}" AND resolutiondate <= "{end}") '
        f'OR (status != Done AND created <= "{end}"))'
    )
    print(f"JQL: {jql}", file=sys.stderr)

    fields = ["key", "summary", "status", "assignee", "created", "resolutiondate", "issuetype"]
    if story_points_field:
        fields.append(story_points_field)

    raw_issues = fetch_all_issues(twg_path, jql, fields)

    issues = []
    for issue in raw_issues:
        assignee = issue.get("assignee") or {}
        status = issue.get("status") or {}
        issuetype = issue.get("issuetype") or {}
        issues.append({
            "key": issue.get("key"),
            "summary": issue.get("summary"),
            "status": status.get("name"),
            "assignee": (assignee.get("displayName") or "").strip(),
            "issuetype": issuetype.get("name"),
            "created": issue.get("created"),
            "resolutiondate": issue.get("resolutiondate"),
            "story_points": issue.get(story_points_field) if story_points_field else None,
        })
    return issues


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------
def reconcile(vendors: dict, issues: list, month_labels: list) -> list:
    by_assignee = defaultdict(list)
    for issue in issues:
        if issue["assignee"]:
            by_assignee[issue["assignee"].lower()].append(issue)

    rows = []
    for name in sorted(vendors, key=str.lower):
        v = vendors[name]
        matched = by_assignee.get(name.lower(), [])
        total_stories = len(matched)
        completed = sum(1 for i in matched if i["status"] == "Done")
        pct_completion = (completed / total_stories * 100) if total_stories else 0.0
        story_points = sum((i["story_points"] or 0) for i in matched)
        expected_hours = story_points * STORY_POINT_HOURS
        pct_hour_gap = ((v["billed"] - expected_hours) / expected_hours * 100) if expected_hours else None

        row = {
            "Vendor Name": name,
            "AWS/Partner": v["partner"],
        }
        for label in month_labels:
            row[f"Billed Hours ({label})"] = round(v["by_month"].get(label, 0.0), 1)
        row.update({
            "Billed Hours (Total)": round(v["billed"], 1),
            "Total Stories": total_stories,
            "Completed Stories": completed,
            "% Completion": round(pct_completion, 1),
            "Story Points": story_points,
            "Expected Hours": expected_hours,
            "% Hour Gap": round(pct_hour_gap, 1) if pct_hour_gap is not None else "N/A",
            "Notes": "" if total_stories else "No JIRA activity found",
        })
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
def write_report(rows: list, month_labels: list, output_path: str):
    columns = ["Vendor Name", "AWS/Partner"]
    columns += [f"Billed Hours ({label})" for label in month_labels]
    columns += ["Billed Hours (Total)", "Total Stories", "Completed Stories",
                "% Completion", "Story Points", "Expected Hours", "% Hour Gap", "Notes"]

    with open(output_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns)
        writer.writeheader()
        for row in rows:
            writer.writerow({c: row[c] for c in columns})
    return columns


def print_markdown_table(rows: list, columns: list):
    print("\n" + " | ".join(columns))
    print(" | ".join(["---"] * len(columns)))
    for row in rows:
        print(" | ".join(str(row[c]) for c in columns))


# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--excel", default=os.path.expanduser("~/Downloads/Maersk MIDAS.xlsx"),
                         help='Path to the timesheet Excel file (default: ~/Downloads/Maersk MIDAS.xlsx)')
    parser.add_argument("--project", default="MID1", help="JIRA project key (default: MID1)")
    parser.add_argument("--start", default=f"{date.today().year}-08-01",
                         help="Period start date, YYYY-MM-DD (default: Aug 1 of current year)")
    parser.add_argument("--end", default=date.today().isoformat(),
                         help="Period end date, YYYY-MM-DD (default: today)")
    parser.add_argument("--output", default=None,
                         help="Output CSV path (default: vendor_jira_report_<YYYYMMDD_HHMMSS>.csv)")
    parser.add_argument("--story-points-field", default=None,
                         help="Override the auto-detected Story Points custom field id (e.g. customfield_10006)")
    args = parser.parse_args()

    start = datetime.strptime(args.start, "%Y-%m-%d").date()
    end = datetime.strptime(args.end, "%Y-%m-%d").date()

    if not args.output:
        args.output = f"vendor_jira_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    if not os.path.exists(args.excel):
        sys.exit(f"ERROR: Excel file not found: {args.excel}")

    print(f"Reading billed hours from: {args.excel}", file=sys.stderr)
    vendors, month_labels = parse_billed_hours(args.excel, start, end)
    print(f"Found {len(vendors)} vendors; period columns: {month_labels}", file=sys.stderr)

    print(f"Fetching JIRA issues for project {args.project} ({start}..{end})...", file=sys.stderr)
    issues = fetch_jira_issues(args.project, start, end, args.story_points_field)
    print(f"Fetched {len(issues)} JIRA issues total.", file=sys.stderr)

    rows = reconcile(vendors, issues, month_labels)
    columns = write_report(rows, month_labels, args.output)

    print(f"\nDone. Wrote report to {args.output}", file=sys.stderr)
    print_markdown_table(rows, columns)


if __name__ == "__main__":
    main()
