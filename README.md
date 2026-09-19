# Vendor Billing vs JIRA Delivery Report

Reconciles vendor/partner **billed hours** (from the Maersk MIDAS timesheet Excel)
against **JIRA story delivery** (story count, % completion, story points) for a
given date window, and reports the hour gap between what was billed and what
the JIRA story points account for.

One script, run locally. No Copilot/Agent Builder setup required.

## What it computes, per vendor

| Column | Meaning |
|---|---|
| Billed Hours (per month + total) | Summed from the timesheet Excel for the given period |
| Total Stories | Count of JIRA work items assigned to that vendor in the period |
| Completed Stories / % Completion | Stories with status = Done, as a percentage |
| Story Points | Sum of the JIRA "Story Points" field across their matched stories |
| Expected Hours | `Story Points x 8` (1 story point = 8 hours; edit `STORY_POINT_HOURS` in the script to change this) |
| % Hour Gap | `(Billed Hours - Expected Hours) / Expected Hours x 100` |

## Prerequisites

1. **Python 3.10+** with dependencies installed:
   ```bash
   pip install -r requirements.txt
   ```
2. **`twg` CLI** (Atlassian Teamwork Graph CLI) installed and authenticated - this
   is what talks to JIRA, no manual API token needed:
   ```bash
   curl -fsSL --retry 2 https://teamwork-graph.atlassian.com/cli/install | bash
   twg doctor   # should print "Authenticated successfully"
   ```
3. The timesheet Excel file (e.g. "Maersk MIDAS.xlsx"), sheet named **"Plan"**,
   with:
   - Row 4 = weekly **start** dates, row 5 = weekly **end** dates.
   - Row 6 headers: column A = Business Title, B = AWS/Partner, C = Level,
     D = Work Stream, E = Name Surname.
   - A first vendor table starting at row 7, and a second table further down
     marked by a cell containing exactly `"PI"` in column A (both tables share
     the same weekly date columns and are summed together per vendor).

## Usage

### One-command pipeline (recommended)

```bash
./run_report.sh
```

This runs the full pipeline (Excel + JIRA + reconciliation) with the default
date range — **static start of August 1 of the current year, through today**
— and writes a timestamped CSV into `reports/` inside this repo (git-ignored,
since reports contain real employee data). Re-run it any time to get an
up-to-date report through "today".

Override the date range via env vars if needed:

```bash
START_DATE=2026-08-01 END_DATE=2026-09-15 ./run_report.sh
```

Any extra flags are passed straight through to the Python script, e.g.:

```bash
./run_report.sh --project MID1 --excel "/path/to/Maersk MIDAS.xlsx"
```

### Running the Python script directly

```bash
python3 generate_vendor_jira_report.py \
    --excel "/path/to/Maersk MIDAS.xlsx" \
    --project MID1 \
    --start 2026-08-01 \
    --end 2026-09-30
```

All flags are optional:

| Flag | Default |
|---|---|
| `--excel` | `~/Downloads/Maersk MIDAS.xlsx` |
| `--project` | `MID1` |
| `--start` | August 1 of the current year |
| `--end` | today |
| `--output` | `vendor_jira_report_<timestamp>.csv` |
| `--story-points-field` | auto-detected by field name (`customfield_XXXXX`) |

The script prints the report as a Markdown table to the terminal and writes
the same data to a timestamped CSV file you can open in Excel.

## How vendor names are matched

JIRA `assignee` display name is matched against the Excel `Name Surname`
column, case-insensitively and trimmed of whitespace. Vendors with no matching
JIRA issues in the period still appear in the report with zeros and a
"No JIRA activity found" note, rather than being dropped.

## Notes / known limitations

- If most JIRA stories don't have the "Story Points" field filled in, Expected
  Hours and % Hour Gap will show as `0` / `N/A` for those vendors - this is a
  JIRA data-entry gap, not a bug in the script. Check coverage with your team.
- If your JIRA instance has multiple similarly-named custom fields (e.g.
  "Story Points", "Story point estimate", "Story Points (custom)"), the script
  prefers an exact "Story Points" name match. Override with
  `--story-points-field customfield_XXXXX` if it picks the wrong one.
- Re-run the script whenever you want a refreshed report (e.g. monthly) with
  updated `--start`/`--end` values.
