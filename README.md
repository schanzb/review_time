# review_time

Payroll audit tool for a multi-site aquatics operation. Compares exported schedule data against Workday time-block audit history to surface:

- **Site transfers** — staff who worked at a location other than their home unit
- **Scheduled hours summary** — total scheduled hours per employee
- **Schedule vs. audit comparison** — unscheduled shifts, early arrivals, late departures, and missing time entries

A `schedule_audit_comparison.csv` is written to the project directory on each run.

## Setup

```bash
pip install openpyxl
```

## Usage

```bash
python3 review_time.py
```

Place the following files in the same directory as the script before running:

| File | Description |
|------|-------------|
| `home_units.csv` | Master employee list — columns: `Name, Email, Home Location, Position, EID` |
| `may-16.CSV`, `may-23.CSV`, … | Weekly schedule exports (any `.csv` except `home_units.csv` is treated as a schedule file) |
| `Time_Block_Audit_-_Cost_Center.xlsx` | Workday time-block audit export (optional — skips comparison if absent) |

## Local configuration

Create a `local_config.py` in the project directory to customize behavior. This file is gitignored and never committed.

```python
# local_config.py
# EIDs for staff who work as-needed with no scheduled shifts.
# These employees are excluded from the schedule comparison report.
UNSCHEDULED_STAFF = {
    12345,  # Example, Last
}
```

## Tuning

Two constants near the top of `review_time.py` control flagging behavior:

- `TOLERANCE_MINS` (default `10`) — grace window in minutes before an early/late deviation is flagged
- `MAX_MATCH_MINS` (default `240`) — if an audit block's start time is this far from every scheduled shift, it's treated as an extra block rather than a deviation
