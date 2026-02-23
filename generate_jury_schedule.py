"""
Jury Day Schedule Generator

Reads:
  - sample_lesson_schedule.xlsx  (Lessons, Jury Information, Pianist availability)
  - timestamped assignments file  (Schedule - Assignments sheet)

Writes:
  - jury_schedule.xlsx

Scheduling rules:
  - Assigned pianists are inflexible (no substitutions).
  - Start times in Jury Information are the EARLIEST possible, not fixed.
  - Areas sharing a room are scheduled sequentially with no overlap.
  - Within an area, students are reordered greedily to avoid pianist conflicts,
    including conflicts caused by the same pianist being used in another area
    scheduled concurrently in a different room.
  - Gaps within an area (pianist unavailable for all remaining students) are
    minimized by jumping directly to the next free minute rather than inserting
    artificial waits.
  - Hourly breaks (10 min) and lunch breaks (30 min) are inserted when flagged.
  - 'Tentative' availability is treated as unavailable.
"""

import argparse
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')


# ── Helpers ───────────────────────────────────────────────────────────────────

def to_minutes(val):
    """Convert timedelta or time to integer minutes from midnight."""
    if isinstance(val, timedelta):
        return int(val.total_seconds() // 60)
    return val.hour * 60 + val.minute

def fmt(minutes):
    h, m = divmod(int(minutes), 60)
    return f"{h:02d}:{m:02d}"


# ── Data loading ──────────────────────────────────────────────────────────────

def load_pianist_unavailability(xl):
    """
    Returns {pianist_name: set_of_unavailable_minutes}.
    Each 30-min grid slot where availability != 'Available' marks all 30 minutes
    in that slot as unavailable. 'Tentative' is treated as unavailable.
    """
    result = {}
    for sheet in [s for s in xl.sheet_names if s.startswith('Pianist -')]:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        name = str(df.iloc[0, 1]).strip()
        unavail = set()
        for _, row in df.iloc[3:].iterrows():
            if pd.isna(row.iloc[0]):
                break
            val = row.iloc[7] if len(row) > 7 and not pd.isna(row.iloc[7]) else ''
            if str(val).strip() != 'Available':
                base = to_minutes(row.iloc[0])
                unavail.update(range(base, base + 30))
        result[name] = unavail
    return result

def load_assignments(filepath):
    """Returns {student_name: accompanist_name}."""
    df = pd.read_excel(filepath, sheet_name='Schedule - Assignments', header=None)
    df.columns = df.iloc[1].tolist()
    df = df.iloc[2:].reset_index(drop=True)
    result = {}
    for _, row in df.iterrows():
        s, a = row.get('Student'), row.get('Accompanist')
        if pd.notna(s) and pd.notna(a):
            result[str(s).strip()] = str(a).strip()
    return result


# ── Pianist booking ───────────────────────────────────────────────────────────

def is_free(pianist_unavail, pianist_booked, name, start, duration):
    """True if pianist is free for the entire [start, start+duration) window."""
    busy = pianist_unavail.get(name, set()) | pianist_booked.get(name, set())
    return not any((start + k) in busy for k in range(duration))

def book(pianist_booked, name, start, duration):
    pianist_booked.setdefault(name, set()).update(range(start, start + duration))

def next_free(pianist_unavail, pianist_booked, name, from_min, duration):
    """Earliest minute >= from_min where pianist is free for 'duration' minutes."""
    t = from_min
    while t < 23 * 60:
        if is_free(pianist_unavail, pianist_booked, name, t, duration):
            return t
        t += 1
    return None


# ── Area scheduling ───────────────────────────────────────────────────────────

def find_gapfree_start(students, earliest, slot_min, pianist_unavail, pianist_booked):
    """
    Find the earliest start time >= earliest from which the greedy algorithm
    can schedule all students with no internal waiting gaps.

    Tries 'earliest' plus every minute where any pianist transitions from
    busy→free, picking the first that results in a fully contiguous schedule.
    """
    if not students:
        return earliest

    # Candidate start times: earliest + every busy→free transition
    candidates = {earliest}
    for _, _, needs, pianist in students:
        if needs and pianist:
            busy = pianist_unavail.get(pianist, set()) | pianist_booked.get(pianist, set())
            for m in range(earliest, earliest + 10 * 60):
                if (m - 1) in busy and m not in busy:
                    candidates.add(m)

    for t_start in sorted(candidates):
        sim_booked = {k: set(v) for k, v in pianist_booked.items()}
        t = t_start
        remaining = list(students)
        ok = True
        while remaining:
            placed = False
            for i, (name, instr, needs, pianist) in enumerate(remaining):
                free = True
                if needs and pianist:
                    busy = pianist_unavail.get(pianist, set()) | sim_booked.get(pianist, set())
                    free = not any((t + k) in busy for k in range(slot_min))
                if free:
                    if needs and pianist:
                        sim_booked.setdefault(pianist, set()).update(range(t, t + slot_min))
                    remaining.pop(i)
                    t += slot_min
                    placed = True
                    break
            if not placed:
                ok = False
                break
        if ok:
            return t_start

    return earliest  # fallback


def schedule_area(students, earliest, slot_min, hourly_break, lunch_break,
                  pianist_unavail, pianist_booked, cross_area_booked):
    """
    Schedule students within one area.

    'students' is a list of (name, instrument, needs_pianist, pianist).
    'pianist_booked' is shared across all areas and updated in place.
    'cross_area_booked' is a snapshot of pianist_booked before this area started,
    representing only commitments from OTHER areas (used for sorting, not conflict checks).

    Strategy:
    - Find the actual start time (find_gapfree_start).
    - Pre-sort remaining students to form compact pianist blocks:
        1. Group students by pianist.
        2. Order groups by when that pianist last appears in cross_area_booked
           (i.e. when they finish their other-area work). Pianists who finish
           earlier get scheduled earlier in this area, so each pianist's slots
           are as contiguous as possible across the whole day.
        3. Within each group the original sheet order is preserved.
    - Greedily place students in that order. If the next student's pianist is
      temporarily unavailable (cross-area conflict), skip forward to the next
      student whose pianist IS free, then return to deferred students once the
      blocking conflict resolves.
    - Jump forward in time (not minute-by-minute) when no student can be placed.
    - Insert hourly (10-min) and lunch (30-min) breaks as configured.

    Returns (slots, end_minute).
    """
    if not students:
        return [], earliest

    actual_start = find_gapfree_start(students, earliest, slot_min,
                                      pianist_unavail, pianist_booked)

    # ── Pre-sort students into compact pianist blocks ─────────────────────────
    # For each pianist, find their last minute of cross-area commitment.
    # Groups are ordered so pianists who finish their other work earliest
    # are scheduled first — minimising idle gaps across the full day.
    def last_cross_area_minute(pianist):
        """Latest minute this pianist is committed in other areas (or earliest if none)."""
        busy = pianist_unavail.get(pianist, set()) | cross_area_booked.get(pianist, set())
        return max(busy) if busy else 0

    # Group students by pianist, preserving original within-group order
    groups = {}
    group_order = []
    for student in students:
        pianist = student[3]  # (name, instrument, needs_pianist, pianist)
        if pianist not in groups:
            groups[pianist] = []
            group_order.append(pianist)
        groups[pianist].append(student)

    # Sort groups: pianists with earlier last cross-area minute go first
    group_order.sort(key=lambda p: last_cross_area_minute(p) if p else -1)

    # Flatten back into a sorted student list
    remaining = []
    for pianist in group_order:
        remaining.extend(groups[pianist])
    # ─────────────────────────────────────────────────────────────────────────

    t = actual_start
    area_start = t
    slots = []
    juries_since_reset = 0
    juries_per_cycle = (60 - 10) // slot_min
    lunch_done = False

    while remaining:
        # Insert 10-min break after every full cycle of juries
        if hourly_break and juries_since_reset > 0 and juries_since_reset % juries_per_cycle == 0:
            slots.append({'type': 'break', 'time': t, 'label': '10-Minute Break'})
            t += 10
            juries_since_reset = 0

        # 30-min lunch break when we reach noon; reset jury count
        if lunch_break and not lunch_done and t >= 12 * 60:
            lunch_done = True
            slots.append({'type': 'lunch', 'time': t, 'label': 'Lunch Break (30 min)'})
            t += 30
            juries_since_reset = 0

        # Place the earliest student in the sorted list whose pianist is free now.
        # This preserves the block ordering while skipping temporarily blocked pianists.
        placed = False
        for i, (name, instrument, needs_pianist, pianist) in enumerate(remaining):
            free = is_free(pianist_unavail, pianist_booked, pianist, t, slot_min) \
                   if (needs_pianist and pianist) else True
            if free:
                if needs_pianist and pianist:
                    book(pianist_booked, pianist, t, slot_min)
                slots.append({
                    'type':         'student',
                    'time':         t,
                    'student':      name,
                    'instrument':   instrument,
                    'need_pianist': needs_pianist,
                    'pianist':      pianist if needs_pianist else '',
                    'duration':     slot_min,
                })
                remaining.pop(i)
                t += slot_min
                juries_since_reset += 1
                placed = True
                break

        if not placed:
            # Jump directly to the next minute any remaining student's pianist is free
            candidates = []
            for name, instrument, needs_pianist, pianist in remaining:
                if needs_pianist and pianist:
                    nf = next_free(pianist_unavail, pianist_booked, pianist, t + 1, slot_min)
                    if nf is not None:
                        candidates.append(nf)
                else:
                    candidates.append(t + 1)
            t = min(candidates) if candidates else t + 1

    return slots, t



# ── Full schedule ─────────────────────────────────────────────────────────────

def build_schedule(lessons, jury_info, pianist_unavail, assignments):
    jury_students = lessons[lessons['Jury'] == 1.0].copy()

    # Build per-area student lists with resolved pianist
    area_students = {}
    for _, row in jury_students.iterrows():
        name = str(row['Student Name']).strip()
        pianist = assignments.get(name, '')
        if not pianist:
            req = row.get('Required Pianist')
            if pd.notna(req) and str(req).strip() not in ('', 'nan'):
                pianist = str(req).strip()
        area_students.setdefault(row['Area'], []).append((
            name,
            str(row.get('Instrument', '')),
            row['Need Pianist'] == 1.0,
            pianist,
        ))

    # Parse jury info
    area_info = {}
    for _, row in jury_info.iterrows():
        area_info[row['Area']] = {
            'earliest':     to_minutes(row['Start Time']),
            'slot_min':     int(row['Jury Length']),
            'hourly_break': bool(row['Hourly Break']),
            'lunch_break':  bool(row['Lunch Break']),
            'location':     row['Location'],
        }

    # Group areas by room and sort by earliest start time
    rooms = {}
    for area, info in area_info.items():
        rooms.setdefault(info['location'], []).append(area)
    for room in rooms:
        rooms[room].sort(key=lambda a: area_info[a]['earliest'])

    # pianist_booked is shared across ALL areas (cross-area conflict tracking)
    pianist_booked = {}
    area_schedule = {}

    for room, area_list in rooms.items():
        room_cursor = None

        for area in area_list:
            info = area_info[area]
            earliest = info['earliest'] if room_cursor is None \
                       else max(info['earliest'], room_cursor)
            students = area_students.get(area, [])

            # Snapshot booked state before this area — used for cross-area urgency sorting
            cross_area_booked = {k: set(v) for k, v in pianist_booked.items()}
            slots, end = schedule_area(
                students, earliest,
                info['slot_min'], info['hourly_break'], info['lunch_break'],
                pianist_unavail, pianist_booked, cross_area_booked,
            )

            actual_start = slots[0]['time'] if slots else earliest
            area_schedule[area] = {
                'location':     room,
                'earliest':     info['earliest'],
                'actual_start': actual_start,
                'actual_end':   end,
                'slots':        slots,
                'slot_min':     info['slot_min'],
            }
            room_cursor = end

    return area_schedule


# ── Excel output ──────────────────────────────────────────────────────────────

def write_excel(area_schedule, output_path):
    wb = Workbook()
    wb.remove(wb.active)

    DARK_BLUE  = '1F4E79'
    MID_BLUE   = '2E75B6'
    LIGHT_BLUE = 'DDEEFF'
    YELLOW     = 'FFF2CC'
    AMBER      = 'FFE0A0'
    WHITE      = 'FFFFFF'

    thin = Side(style='thin', color='AAAAAA')
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    lft  = Alignment(horizontal='left',   vertical='center', wrap_text=True)

    def cell(ws, r, col, val, bold=False, size=10, fgcolor='000000',
             bg=None, align=None, border=None):
        c = ws.cell(row=r, column=col, value=val)
        c.font  = Font(name='Arial', bold=bold, size=size, color=fgcolor)
        if bg:     c.fill      = PatternFill('solid', start_color=bg)
        if align:  c.alignment = align
        if border: c.border    = border
        return c

    ordered = sorted(area_schedule.items(),
                     key=lambda x: (x[1]['location'], x[1]['actual_start']))

    # ── Summary sheet ─────────────────────────────────────────────────
    ws = wb.create_sheet('Summary')
    for col, w in zip('ABCDEF', [15, 15, 10, 10, 12, 30]):
        ws.column_dimensions[col].width = w
    ws.merge_cells('A1:F1')
    cell(ws,1,1,'JURY DAY SCHEDULE — SUMMARY',
         bold=True, size=14, fgcolor=WHITE, bg=DARK_BLUE, align=ctr)
    ws.row_dimensions[1].height = 30
    for i, h in enumerate(['Area','Location','Start','End','# Students','Notes'], 1):
        cell(ws,2,i,h, bold=True, size=10, fgcolor=WHITE, bg=MID_BLUE, align=ctr, border=bdr)

    prev_loc = None
    r = 3
    for area, sched in ordered:
        students = [s for s in sched['slots'] if s['type'] == 'student']
        # Compute internal gaps
        gaps, prev_end = [], None
        for slot in sched['slots']:
            if slot['type'] == 'break':
                prev_end = slot['time'] + 10; continue
            if slot['type'] == 'lunch':
                prev_end = slot['time'] + 30; continue
            if prev_end and slot['time'] > prev_end:
                gaps.append(slot['time'] - prev_end)
            prev_end = slot['time'] + slot['duration']

        bg = LIGHT_BLUE if sched['location'] != prev_loc else WHITE
        prev_loc = sched['location']

        notes = []
        delay = sched['actual_start'] - sched['earliest']
        if delay > 0:
            notes.append(f"Starts {delay}m after earliest ({fmt(sched['earliest'])})")
        if gaps:
            notes.append(f"{sum(gaps)}m wait gap(s) within area")

        cell(ws,r,1,area,                        bg=bg, align=lft, border=bdr)
        cell(ws,r,2,sched['location'],            bg=bg, align=ctr, border=bdr)
        cell(ws,r,3,fmt(sched['actual_start']),   bg=bg, align=ctr, border=bdr)
        cell(ws,r,4,fmt(sched['actual_end']),     bg=bg, align=ctr, border=bdr)
        cell(ws,r,5,len(students),                bg=bg, align=ctr, border=bdr)
        note_bg = AMBER if notes else bg
        cell(ws,r,6,'; '.join(notes), bg=note_bg, align=lft, border=bdr)
        r += 1

    # ── Per-area sheets ───────────────────────────────────────────────
    for area, sched in ordered:
        ws = wb.create_sheet(area)
        for col, w in zip('ABCDE', [10, 20, 16, 18, 14]):
            ws.column_dimensions[col].width = w
        ws.merge_cells('A1:E1')
        title = (f"{area} Jury  ·  {sched['location']}  ·  "
                 f"{fmt(sched['actual_start'])}–{fmt(sched['actual_end'])}"
                 f"  ({sched['slot_min']} min/student)")
        cell(ws,1,1,title, bold=True, size=12, fgcolor=WHITE, bg=DARK_BLUE, align=ctr)
        ws.row_dimensions[1].height = 26
        for i, h in enumerate(['Time','Student','Instrument','Pianist','Notes'], 1):
            cell(ws,2,i,h, bold=True, size=10, fgcolor=WHITE, bg=MID_BLUE, align=ctr, border=bdr)

        r, alt = 3, False
        prev_end = None
        for slot in sched['slots']:
            if slot['type'] == 'break':
                ws.merge_cells(f'A{r}:E{r}')
                cell(ws,r,1,f"— {slot['label']} ({fmt(slot['time'])}) —",
                     bold=True, bg=YELLOW, align=ctr, border=bdr)
                prev_end = slot['time'] + 10
                ws.row_dimensions[r].height = 18
                r += 1; alt = False; continue
            if slot['type'] == 'lunch':
                ws.merge_cells(f'A{r}:E{r}')
                cell(ws,r,1,f"— {slot['label']} ({fmt(slot['time'])}) —",
                     bold=True, bg=YELLOW, align=ctr, border=bdr)
                prev_end = slot['time'] + 30
                ws.row_dimensions[r].height = 18
                r += 1; alt = False; continue

            bg = LIGHT_BLUE if alt else WHITE
            alt = not alt
            note = ''
            if prev_end is not None and slot['time'] > prev_end:
                note = f"+{slot['time'] - prev_end}m wait"
                bg = AMBER

            cell(ws,r,1,fmt(slot['time']),  bg=bg, align=ctr, border=bdr)
            cell(ws,r,2,slot['student'],    bg=bg, align=lft, border=bdr)
            cell(ws,r,3,slot['instrument'], bg=bg, align=ctr, border=bdr)
            cell(ws,r,4,slot['pianist'],    bg=bg, align=ctr, border=bdr)
            cell(ws,r,5,note,               bg=bg, align=ctr, border=bdr)
            ws.row_dimensions[r].height = 18
            prev_end = slot['time'] + slot['duration']
            r += 1

    wb.save(output_path)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Generate jury day schedule.')
    parser.add_argument('--lessons',  required=True,
                        help='Lesson schedule .xlsx (Lessons, Jury Information, Pianist sheets)')
    parser.add_argument('--pianists', required=True,
                        help='Pianist assignments .xlsx (Schedule - Assignments sheet)')
    args = parser.parse_args()

    timestamp   = datetime.now().strftime('%Y%m%d_%H%M')
    output_dir  = os.path.dirname(args.lessons) or '.'
    output_file = os.path.join(output_dir, f'jury_schedule_{timestamp}.xlsx')

    xl              = pd.ExcelFile(args.lessons)
    lessons         = pd.read_excel(xl, sheet_name='Lessons')
    jury_info       = pd.read_excel(xl, sheet_name='Jury Information')
    pianist_unavail = load_pianist_unavailability(xl)
    assignments     = load_assignments(args.pianists)

    area_schedule = build_schedule(lessons, jury_info, pianist_unavail, assignments)
    write_excel(area_schedule, output_file)
    print(f"Saved: {output_file}\n")

    print(f"{'Area':<12} {'Location':<14} {'Earliest':>8} {'Start':>6} {'End':>6}  {'#':>3}  Notes")
    print('─' * 75)
    for area, sched in sorted(area_schedule.items(),
                               key=lambda x: (x[1]['location'], x[1]['actual_start'])):
        students = [s for s in sched['slots'] if s['type'] == 'student']
        gaps, prev_end = [], None
        for slot in sched['slots']:
            if slot['type'] == 'break':  prev_end = slot['time'] + 10;  continue
            if slot['type'] == 'lunch':  prev_end = slot['time'] + 30;  continue
            if prev_end and slot['time'] > prev_end:
                gaps.append(slot['time'] - prev_end)
            prev_end = slot['time'] + slot['duration']
        notes = []
        delay = sched['actual_start'] - sched['earliest']
        if delay: notes.append(f"delayed +{delay}m")
        if gaps:  notes.append(f"{sum(gaps)}m gap")
        print(f"{area:<12} {sched['location']:<14}"
              f" {fmt(sched['earliest']):>8} {fmt(sched['actual_start']):>6}"
              f" {fmt(sched['actual_end']):>6}  {len(students):>3}  {'  '.join(notes)}")

if __name__ == '__main__':
    main()
