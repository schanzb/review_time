import csv
import warnings
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

# Maps keywords in Position Name → canonical location matching home_units.csv
LOCATION_MAP = [
    ('memorial beach', 'Sharon'),
    ('natick',         'Natick'),
    ('nahanton',       'Nahanton'),
    ('sharon',         'Sharon'),
]

# Position categories that aren't regular work shifts — skip these
SKIP_POSITIONS = {
    'available senior staff',
    'general staff',
    'first aid/cpr training',
    'in person training',
    'motorboat training',
    'on the water training',
    'prep for season',
    'shift lead training',
    'shift supervisor meeting',
}

# Minutes of grace period before flagging early arrival / late departure
TOLERANCE_MINS = 10

# If an audit block's start time is this far from every scheduled shift,
# treat it as an extra block rather than a late/early deviation
MAX_MATCH_MINS = 240


def extract_location(position_name, shift_desc):
    for keyword, location in LOCATION_MAP:
        if keyword in position_name.lower():
            return location
    for keyword, location in LOCATION_MAP:
        if keyword in shift_desc.lower():
            return location
    return None


def load_home_units(path):
    home_units = {}
    with open(path, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            try:
                eid = int(row['EID'])
                home_units[eid] = {
                    'name': row['Name'],
                    'home': row['Home Location'],
                }
            except (ValueError, KeyError):
                pass
    return home_units


def parse_hhmm(s):
    """Parse 'HH:MM' (24-hour) string to minutes since midnight."""
    if not s:
        return None
    h, m = s.split(':')
    return int(h) * 60 + int(m)


def parse_12h(s):
    """Parse '12:00 PM' to minutes since midnight."""
    dt = datetime.strptime(s.strip(), '%I:%M %p')
    return dt.hour * 60 + dt.minute


def fmt_mins(mins):
    if mins is None:
        return '??:??'
    return f"{mins // 60:02d}:{mins % 60:02d}"


def load_schedule(schedule_files):
    """Returns {(eid, date_str): [(start_mins, end_mins)]}."""
    schedule = {}
    for path in schedule_files:
        with open(path, newline='') as f:
            for row in csv.DictReader(f):
                emp_num = row.get('Employee Number', '').strip().strip('"')
                if not emp_num:
                    continue
                try:
                    eid = int(emp_num)
                except ValueError:
                    continue

                position = row.get('Position Name', '').strip()
                if position.lower() in SKIP_POSITIONS:
                    continue

                date_raw = row.get('Date', '').strip()
                start_raw = row.get('Start Time', '').strip()
                end_raw = row.get('End Time', '').strip()
                if not date_raw or not start_raw or not end_raw:
                    continue

                try:
                    date_str = datetime.strptime(date_raw, '%m/%d/%Y').strftime('%Y-%m-%d')
                    start_mins = parse_12h(start_raw)
                    end_mins = parse_12h(end_raw)
                except ValueError:
                    continue

                schedule.setdefault((eid, date_str), []).append((start_mins, end_mins))
    return schedule


def load_audit_blocks(path):
    """
    Returns {(eid, date_str, in_time_str): (in_mins, out_mins, worker_str)}.

    Groups audit history rows by (eid, date, in_time) and keeps only the
    most-recently-modified entry per group, which reflects the final state.
    """
    with warnings.catch_warnings():
        warnings.simplefilter('ignore')
        wb = openpyxl.load_workbook(path)
    ws = wb.active

    # Step 1: per (eid, date, in_str), keep the row with the latest Modified Moment
    best = {}  # key → (modified_moment, in_mins, out_mins, worker)

    for row in ws.iter_rows(min_row=3, values_only=True):
        eid_raw    = row[1]   # Employee ID
        hist_date  = row[11]  # Historical Calculated Date
        modified   = row[13]  # Modified Moment
        in_str     = row[15]  # Time Block In Time: HH:mm
        out_str    = row[16]  # Time Block Out Time: HH:mm
        worker     = row[0]   # Worker display name

        if not eid_raw or not hist_date or not in_str or not out_str or not modified:
            continue

        try:
            eid = int(str(eid_raw).strip())
        except ValueError:
            continue

        date_str = str(hist_date)[:10]
        in_mins  = parse_hhmm(in_str)
        out_mins = parse_hhmm(out_str)
        if in_mins is None or out_mins is None:
            continue

        key = (eid, date_str, in_str)
        existing = best.get(key)
        if existing is None or modified > existing[0]:
            best[key] = (modified, in_mins, out_mins, str(worker))

    # Step 2: per (eid, date), remove stale blocks whose window overlaps a
    # more-recently-edited block.  This handles the case where an edit changes
    # the in_time, leaving the old in_time as a separate key in step 1.
    by_emp_date = defaultdict(list)
    for (eid, date_str, in_str), (modified, in_mins, out_mins, worker) in best.items():
        by_emp_date[(eid, date_str)].append((modified, in_str, in_mins, out_mins, worker))

    result = {}
    for (eid, date_str), entries in by_emp_date.items():
        entries.sort(key=lambda x: x[0], reverse=True)  # newest first
        kept = []  # list of (in_mins, out_mins) already accepted
        for modified, in_str, in_mins, out_mins, worker in entries:
            overlaps = any(
                max(in_mins, k_in) < min(out_mins, k_out)
                for k_in, k_out in kept
            )
            if not overlaps:
                kept.append((in_mins, out_mins))
                result[(eid, date_str, in_str)] = (in_mins, out_mins, worker)

    return result


def build_comparison_rows(schedule, audit_blocks, home_units):
    """
    Returns one row dict per (employee, scheduled-shift) pairing covering every
    scheduled shift and every audit time block.

    Flags produced:
      early in        arrived more than TOLERANCE_MINS before scheduled start
      late in         arrived more than TOLERANCE_MINS after scheduled start
      early out       left more than TOLERANCE_MINS before scheduled end
      late out        left more than TOLERANCE_MINS after scheduled end
      no time reported  scheduled shift has no matching audit block
      extra block     audit block on a date with no scheduled shift
      missed punch    audit block has an in time but no out time (or vice-versa)
    """
    rows = []

    # Collect all (eid, date) pairs from both sources
    sched_keys = set(schedule.keys())
    audit_by_emp_date = defaultdict(list)
    for (eid, date_str, _in_str), (in_mins, out_mins, worker) in audit_blocks.items():
        audit_by_emp_date[(eid, date_str)].append((in_mins, out_mins, worker))
    audit_keys = set(audit_by_emp_date.keys())
    all_keys = sched_keys | audit_keys

    for (eid, date_str) in sorted(all_keys):
        worker_name = (home_units.get(eid, {}).get('name')
                       or (audit_by_emp_date[(eid, date_str)][0][2]
                           if (eid, date_str) in audit_by_emp_date else f'EID {eid}'))
        sched_shifts = schedule.get((eid, date_str), [])
        actual_blocks = audit_by_emp_date.get((eid, date_str), [])

        if sched_shifts and not actual_blocks:
            # Scheduled but nothing reported
            for sched_start, sched_end in sched_shifts:
                rows.append({
                    'name':      worker_name,
                    'date':      date_str,
                    'sched':     f"{fmt_mins(sched_start)}–{fmt_mins(sched_end)}",
                    'reported':  '—',
                    'flags':     'no time reported',
                })
            continue

        if actual_blocks and not sched_shifts:
            # Reported time with no schedule on that date
            for in_mins, out_mins, _w in actual_blocks:
                flags = ['extra block']
                if in_mins is None or out_mins is None:
                    flags.append('missed punch')
                rows.append({
                    'name':      worker_name,
                    'date':      date_str,
                    'sched':     '—',
                    'reported':  f"{fmt_mins(in_mins)}–{fmt_mins(out_mins)}",
                    'flags':     ', '.join(flags),
                })
            continue

        # Both sides present — match each actual block to its closest scheduled shift
        matched_scheds = set()
        for in_mins, out_mins, _w in actual_blocks:
            if in_mins is None or out_mins is None:
                rows.append({
                    'name':      worker_name,
                    'date':      date_str,
                    'sched':     '—',
                    'reported':  f"{fmt_mins(in_mins)}–{fmt_mins(out_mins)}",
                    'flags':     'missed punch',
                })
                continue

            best_idx, (sched_start, sched_end) = min(
                enumerate(sched_shifts), key=lambda x: abs(x[1][0] - in_mins)
            )

            # Too far from any scheduled shift → treat as an extra block
            if abs(sched_start - in_mins) > MAX_MATCH_MINS:
                rows.append({
                    'name':      worker_name,
                    'date':      date_str,
                    'sched':     '—',
                    'reported':  f"{fmt_mins(in_mins)}–{fmt_mins(out_mins)}",
                    'flags':     'extra block',
                })
                continue

            matched_scheds.add(best_idx)

            parts = []
            early_in  = sched_start - in_mins   # positive → arrived early
            late_in   = in_mins - sched_start    # positive → arrived late
            early_out = sched_end - out_mins     # positive → left early
            late_out  = out_mins - sched_end     # positive → left late

            if early_in  > TOLERANCE_MINS: parts.append(f'early in ({early_in}m)')
            if late_in   > TOLERANCE_MINS: parts.append(f'late in ({late_in}m)')
            if early_out > TOLERANCE_MINS: parts.append(f'early out ({early_out}m)')
            if late_out  > TOLERANCE_MINS: parts.append(f'late out ({late_out}m)')

            rows.append({
                'name':      worker_name,
                'date':      date_str,
                'sched':     f"{fmt_mins(sched_start)}–{fmt_mins(sched_end)}",
                'reported':  f"{fmt_mins(in_mins)}–{fmt_mins(out_mins)}",
                'flags':     ', '.join(parts),
            })

        # Any scheduled shift that was never matched → no time reported
        for idx, (sched_start, sched_end) in enumerate(sched_shifts):
            if idx not in matched_scheds:
                rows.append({
                    'name':      worker_name,
                    'date':      date_str,
                    'sched':     f"{fmt_mins(sched_start)}–{fmt_mins(sched_end)}",
                    'reported':  '—',
                    'flags':     'no time reported',
                })

    return rows


def main():
    base_dir = Path(__file__).parent
    home_units = load_home_units(base_dir / 'home_units.csv')

    schedule_files = sorted(
        f for f in base_dir.iterdir()
        if f.suffix.lower() == '.csv' and f.name.lower() != 'home_units.csv'
    )

    flags = []
    staff = {}  # eid -> {name, first_name, total_hours}

    for schedule_file in schedule_files:
        with open(schedule_file, newline='') as f:
            for row in csv.DictReader(f):
                emp_num = row.get('Employee Number', '').strip().strip('"')
                if not emp_num:
                    continue
                try:
                    eid = int(emp_num)
                except ValueError:
                    continue
                if eid not in home_units:
                    continue

                position = row.get('Position Name', '').strip()
                if position.lower() in SKIP_POSITIONS:
                    continue

                name = row.get('Employee Name', '').strip() or home_units[eid]['name']
                first_name = row.get('Employee First Name', '').strip()
                if eid not in staff:
                    staff[eid] = {'name': name, 'first_name': first_name, 'total_hours': 0.0}
                try:
                    staff[eid]['total_hours'] += float(row.get('Duration', 0))
                except (ValueError, TypeError):
                    pass

                shift_desc = row.get('Shift Description', '').strip()
                location = extract_location(position, shift_desc)
                if location is None:
                    continue

                home = home_units[eid]['home']
                if location != home:
                    flags.append({
                        'name': name,
                        'first_name': first_name,
                        'date': row.get('Date', '').strip(),
                        'home': home,
                        'worked_at': location,
                        'file': schedule_file.name,
                    })

    # --- Output 1: site transfer flags ---
    if not flags:
        print('No site transfers found.')
    else:
        flags.sort(key=lambda x: (x['first_name'].lower(), x['name'].lower()))
        print(f"\n{'Employee':<28} {'Date':<12} {'Home Unit':<12} {'Worked At'}")
        print('-' * 65)
        for entry in flags:
            print(f"{entry['name']:<28} {entry['date']:<12} {entry['home']:<12} {entry['worked_at']}")
        print(f"\n{len(flags)} flag(s) across {len(schedule_files)} schedule file(s).")

    # --- Output 2: total hours for all scheduled staff ---
    staff_list = sorted(staff.values(), key=lambda x: (x['first_name'].lower(), x['name'].lower()))
    print(f"\n{'Employee':<28} {'Sched Hrs':>9}")
    print('-' * 38)
    for s in staff_list:
        print(f"{s['name']:<28} {s['total_hours']:>9.2f}")
    grand_total = sum(s['total_hours'] for s in staff_list)
    print(f"\n{len(staff_list)} staff member(s), {grand_total:.2f} scheduled hours.")

    # --- Output 3: schedule vs. audit comparison ---
    audit_path = base_dir / 'Time_Block_Audit_-_Cost_Center.xlsx'
    if not audit_path.exists():
        print('\n[No audit file found — skipping schedule comparison]')
        return

    schedule     = load_schedule(schedule_files)
    audit_blocks = load_audit_blocks(audit_path)
    rows         = build_comparison_rows(schedule, audit_blocks, home_units)

    flagged = [r for r in rows if r['flags']]
    clean   = [r for r in rows if not r['flags']]

    print(f'\n{"=" * 90}')
    print(f'SCHEDULE vs. AUDIT  (tolerance: {TOLERANCE_MINS} min)  —  '
          f'{len(flagged)} flagged, {len(clean)} clean')
    print(f'{"=" * 90}')
    print(f"{'Employee':<28} {'Date':<12} {'Scheduled':<13} {'Reported':<13} {'Flags'}")
    print('-' * 90)
    for r in rows:
        print(f"{r['name']:<28} {r['date']:<12} {r['sched']:<13} {r['reported']:<13} {r['flags']}")


if __name__ == '__main__':
    main()
