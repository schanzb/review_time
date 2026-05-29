# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the script

```bash
python3 review_time.py
```

Requires `openpyxl`: `pip install openpyxl`

## What this does

A payroll/scheduling audit tool for a multi-site aquatics operation (Natick, Sharon, Nahanton). It reads exported schedule data and Workday time-block audit history, then produces three reports:

1. **Site transfers** — staff scheduled at a location other than their home unit
2. **Scheduled hours summary** — total hours per employee across all schedule files
3. **Schedule vs. audit comparison** — detects unscheduled shifts, early arrivals, and late departures

## Data sources

| File | Source | Role |
|------|--------|------|
| `home_units.csv` | Manually maintained | Master list of tracked employees with EID, home location, position |
| `may-16.CSV`, `may-23.CSV` (etc.) | Exported from scheduling system | One file per week; any `.csv` in the directory (except `home_units.csv`) is treated as a schedule file |
| `Time_Block_Audit_-_Cost_Center.xlsx` | Exported from Workday | Time-block audit history; optional — script skips Output 3 if absent |

## Key design decisions

**EID as the join key.** Employee numbers are parsed to `int` everywhere (stripping leading zeros). The schedule CSVs use `Employee Number`; the audit Excel uses `Employee ID` (column B); `home_units.csv` uses `EID`. All three resolve to the same integer.

**Audit history deduplication.** Workday exports every historical version of a time block. Rows are grouped by `(eid, date, in_time_str)` and the entry with the latest `Modified Moment` (column N) is kept as the canonical final state. Rows with a null `Modified Moment` are summary/clock-event rows and are skipped.

**Schedule time format vs. audit time format.** Schedule CSVs use 12-hour format (`12:00 PM`); the audit Excel uses 24-hour `HH:MM`. Both are normalized to minutes-since-midnight internally.

**Shift matching.** When an employee has multiple scheduled shifts on one day, each audit block is matched to the closest scheduled shift by start time.

**`SKIP_POSITIONS`** — non-shift position types (trainings, meetings, etc.) are excluded from both the schedule and site-transfer analysis. Add new ones here as the operation adds position types.

**`TOLERANCE_MINS`** (default `10`) — grace window before flagging early/late. Adjust at the top of the file.

**`LOCATION_MAP`** — keyword-to-location mapping used for site-transfer detection. "Memorial Beach" maps to Sharon. Add new sites here.
