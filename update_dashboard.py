#!/usr/bin/env python3
"""
Body Shop (RA) Andon Dashboard — Daily Auto-Update Script
Run at 17:00 Mon-Fri via Windows Task Scheduler
Reads HOP + DT Excel files, updates HTML DAYS_DATA, pushes to GitHub
"""

import openpyxl, warnings, datetime, json, re, shutil, subprocess, sys, os
warnings.filterwarnings('ignore')

# ── CONFIGURATION ──────────────────────────────────────────────────────────────
HOP_SRC = r"C:\Users\NYOUSIF\OneDrive - Volvo Cars\A Shop Production SI and Supervisors - Hop Line Downtime\HOP New Downtime Breakdown.xlsm"
DT_SRC  = r"C:\Users\NYOUSIF\OneDrive - Volvo Cars\A Shop Production SI and Supervisors - General\Downtime Tracker Logv6a.xlsm"
DASH    = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\body_shop_intelligence.html"
WORK    = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
LOG     = os.path.join(WORK, "update_log.txt")

# GitHub repo — set this to your repo path after cloning
GITHUB_REPO = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard"  # same folder if repo is here
GITHUB_ENABLED = True   # git configured — pushes to github.com/ninefyousif-sys/Andon-Dashboard

# Production calendar: WK12 D1 = Mon 16 Mar 2026 (reference anchor)
WK12_MON = datetime.date(2026, 3, 16)

def log(msg):
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

# ── DATE / WEEK HELPERS ────────────────────────────────────────────────────────
def date_to_wk_day(d):
    """Return (wk_str, day_int) for a given date using Volvo production calendar."""
    delta = (d - WK12_MON).days
    week_offset = delta // 7
    day_of_week = d.weekday()  # 0=Mon … 4=Fri, skip weekends
    if day_of_week > 4:
        return None, None  # weekend
    wk_num = 12 + week_offset
    day_num = day_of_week + 1  # 1=Mon … 5=Fri
    return f"WK{wk_num}", day_num

def last_n_working_days(n=7, reference=None):
    """Return list of last N working dates (Mon-Fri), most recent first."""
    ref = reference or datetime.date.today()
    result = []
    d = ref
    while len(result) < n:
        if d.weekday() < 5:
            result.append(d)
        d -= datetime.timedelta(days=1)
    return result

# ── EXCEL HELPERS ──────────────────────────────────────────────────────────────
def time_to_min(t):
    if t is None: return 0
    if isinstance(t, datetime.datetime): return t.hour*60+t.minute+t.second/60
    if isinstance(t, datetime.time): return t.hour*60+t.minute+t.second/60
    try:
        f = float(t)
        return f * 1440 if 0 < f < 2 else f
    except: return 0

def dur_min(s, e):
    sm, em = time_to_min(s), time_to_min(e)
    return round(em - sm, 1) if em > sm else 0

def t_fmt(t):
    if isinstance(t, (datetime.datetime, datetime.time)):
        return f"{t.hour:02d}:{t.minute:02d}"
    return ''

EXCLUDE_RESPONSIBILITY = {'shop flow', 'shop flow '}
EXCLUDE_ERROR_TYPES    = {'blocked', 'starved', 'blocked ', 'starved '}

def should_exclude(resp, err):
    r = str(resp).lower().strip() if resp else ''
    e = str(err).lower().strip()  if err  else ''
    if r in EXCLUDE_RESPONSIBILITY or 'shop flow' in r: return True
    if e in EXCLUDE_ERROR_TYPES or 'blocked' in e or 'starved' in e: return True
    return False

def hop_code(area):
    a = str(area)
    if '285' in a: return '285'
    if '286' in a: return '286'
    if '287' in a or 'tailgate' in a.lower(): return '287'
    if '198' in a: return '198'
    if '155' in a or 'roof' in a.lower(): return '155'
    if '197' in a: return '197'
    m = re.search(r'\b(\d{3})\b', a)
    return m.group(1) if m else None

def dt_code(area):
    a = str(area).lower()
    for pat, code in [
        ('rfc-236','236'), ('236 spa','236'), ('marr','138'),
        ('fwhlh-232','232'), ('232 spa','232'), ('fwhrh-233','233'), ('233 spa','233'),
        ('rwhlh-235','235'), ('rwhrh-237','237'),
        ('ff-234','234'), ('234 spa','234'),
        ('bso-258','258'), ('258 spa','258'),
        ('fs-231','231'), ('231 spa','231'),
        ('rf-136','236'), ('136 spa','236'),
        ('roof','155'), ('155','155')
    ]:
        if pat in a: return code
    m = re.search(r'\b(2\d{2})\b', str(area))
    return m.group(1) if m else None

HOP_COL = {'285':'#cc2838','286':'#d07018','287':'#d07018','198':'#a880e0',
           '155':'#30c0c5','197':'#5a7590'}
DT_COL  = {'236':'#0075be','138':'#0075be','232':'#0075be','233':'#0075be',
           '234':'#0075be','235':'#0075be','237':'#0075be','258':'#a880e0',
           '231':'#0075be','155':'#30c0c5'}

# ── MAIN READ FUNCTION ─────────────────────────────────────────────────────────
def read_day(hop_ws, dt_ws, wk_str, day_int):
    """Read PBI + Gantt for one WK/Day from both sheets, excluding B/S/SF."""
    pbi   = {}
    gantt = []

    for ws, code_fn, col_map, resp_idx, err_idx in [
        (hop_ws, hop_code, HOP_COL, 11, 12),
        (dt_ws,  dt_code,  DT_COL,  11, 12),
    ]:
        if ws is None:
            continue
        for i, row in enumerate(ws.rows):
            if i == 0: continue
            vals = [c.value for c in row]
            try: yr = int(float(str(vals[0])))
            except: continue
            if yr != 2026: continue
            if str(vals[1]).strip() != wk_str: continue
            try: d = int(float(str(vals[2])))
            except: continue
            if d != day_int: continue

            start, end = vals[3], vals[4]
            dur = dur_min(start, end)
            if dur <= 0: continue

            area = str(vals[7]) if vals[7] and str(vals[7]) != 'None' else ''
            if not area: continue

            resp = vals[resp_idx] if len(vals) > resp_idx else None
            err  = vals[err_idx]  if len(vals) > err_idx  else None
            if should_exclude(resp, err): continue

            code = code_fn(area)
            if code:
                pbi[code] = round(pbi.get(code, 0) + dur, 1)

            s, e = t_fmt(start), t_fmt(end)
            if s and dur > 0:
                col = col_map.get(code, '#5a7590')
                lbl_parts = []
                if resp: lbl_parts.append(str(resp))
                if err:  lbl_parts.append(str(err))
                lbl = ' · '.join(lbl_parts)[:50].replace("'", '').replace('"', '')
                gantt.append({
                    'line': area[:22].replace("'", ''),
                    's': s, 'e': e, 'sec': int(dur * 60),
                    'col': col, 'lbl': lbl
                })

    # Sort gantt by start time, keep only shift hours
    gantt = sorted([g for g in gantt if '06:00' <= g['s'] <= '15:00'], key=lambda x: x['s'])
    return pbi, gantt

# ── BUILD DAYS_DATA JAVASCRIPT ─────────────────────────────────────────────────
LABEL_MAP = {0:'Mon', 1:'Tue', 2:'Wed', 3:'Thu', 4:'Fri'}

# BOL/Empty hourly data — these are static per day from Snowflake
# In a full implementation, fetch from Snowflake. For now, keep last known values.
STATIC_PROD = {
    '2026-03-10': {'bol_h':[8,10,9,10,4,9,11,7],  'empty_h':[10,9,9,10,3,8,10,8],  'bol_tot':70, 'empty_tot':71},
    '2026-03-11': {'bol_h':[10,9,10,10,10,10,11,3],'empty_h':[9,9,9,11,10,8,10,8],  'bol_tot':76, 'empty_tot':77},
    '2026-03-12': {'bol_h':[10,8,9,10,11,9,10,6],  'empty_h':[10,10,6,10,12,9,5,11],'bol_tot':76, 'empty_tot':75},
    '2026-03-13': {'bol_h':[11,11,10,8,10,7,10,7], 'empty_h':[11,10,8,9,7,7,12,8],  'bol_tot':77, 'empty_tot':75},
    '2026-03-16': {'bol_h':[9,12,5,11,12,8,12,9],  'empty_h':[12,12,10,12,9,10,11,9],'bol_tot':98,'empty_tot':104,'overtime':True,'otNote':'Shift ran until 16:38'},
    '2026-03-17': {'bol_h':[7,12,6,12,11,10,11,7], 'empty_h':[12,12,11,11,13,10,10,7],'bol_tot':80,'empty_tot':89},
    '2026-03-18': {'bol_h':[8,13,12,11,10,7,13,12], 'empty_h':[10,10,12,11,11,6,13,13],'bol_tot':86,'empty_tot':86},
}

def pbi_to_js(pbi):
    items = sorted(pbi.items(), key=lambda x: -x[1])
    return '{' + ','.join(f"'{k}':{v}" for k, v in items) + '}'

def gantt_to_js(gantt):
    items = []
    for g in gantt[:25]:  # max 25 events to keep file size reasonable
        items.append(
            f"{{line:'{g['line']}',s:'{g['s']}',e:'{g['e']}',sec:{g['sec']},col:'{g['col']}',lbl:'{g['lbl']}'}}"
        )
    return '[' + ',\n      '.join(items) + ']'

def build_day_entry(date_str, wk_str, day_int, pbi, gantt, is_today=False):
    d = datetime.date.fromisoformat(date_str)
    day_name = LABEL_MAP[d.weekday()]
    label = f"{day_name} {d.day}-{d.strftime('%b')}"
    prod = STATIC_PROD.get(date_str, {
        'bol_h': [8]*8, 'empty_h': [8]*8, 'bol_tot': 80, 'empty_tot': 80
    })
    overtime = prod.get('overtime', False)
    ot_note  = prod.get('otNote', '')

    lines = [f"  '{date_str}':{{{repr(label).replace(chr(39),'')},"]
    lines.append(f"    label:'{label}',overtime:{'true' if overtime else 'false'}"
                 + (f",otNote:'{ot_note}'" if ot_note else '')
                 + (',isToday:true' if is_today else '') + ',')
    bol_h_js  = '[' + ','.join(str(x) for x in prod['bol_h'])  + ']'
    empty_h_js= '[' + ','.join(str(x) for x in prod['empty_h']) + ']'
    lines.append(f"    bol_h:{bol_h_js}, empty_h:{empty_h_js},")
    lines.append(f"    bol_tot:{prod['bol_tot']}, empty_tot:{prod['empty_tot']},")
    lines.append(f"    // {wk_str} D{day_int} · Equipment DT only · {sum(pbi.values()):.0f} min")
    lines.append(f"    pbi:{pbi_to_js(pbi)},")
    lines.append(f"    gantt:{gantt_to_js(gantt)}}},")
    return '\n'.join(lines)

# ── MAIN UPDATE FUNCTION ───────────────────────────────────────────────────────
def update():
    today = datetime.date.today()
    log(f"=== Daily update started ({today}) ===")

    # Copy fresh Excel files
    tmp_hop = os.path.join(WORK, '_hop_tmp.xlsm')
    tmp_dt  = os.path.join(WORK, '_dt_tmp.xlsm')
    for src, dst, name in [(HOP_SRC, tmp_hop, 'HOP'), (DT_SRC, tmp_dt, 'DT')]:
        if os.path.exists(src):
            shutil.copy2(src, dst)
            log(f"Copied {name} file ({os.path.getsize(dst)//1024} KB)")
        else:
            log(f"WARNING: {name} file not found at {src}")

    # Open workbooks
    hop_ws = dt_ws = None
    try:
        wb_hop = openpyxl.load_workbook(tmp_hop, data_only=True, read_only=True)
        hop_ws = wb_hop['New(DT LOG)']
        log("Opened HOP workbook")
    except Exception as e:
        log(f"ERROR opening HOP: {e}")

    try:
        wb_dt = openpyxl.load_workbook(tmp_dt, data_only=True, read_only=True)
        dt_ws = wb_dt['New(DT LOG)']
        log("Opened DT workbook")
    except Exception as e:
        log(f"ERROR opening DT: {e}")

    # Get last 7 working days
    working_days = last_n_working_days(7, today)
    log(f"Processing {len(working_days)} working days: {[str(d) for d in working_days]}")

    # Build DAYS_DATA entries
    entries = []
    for d in reversed(working_days):  # chronological order
        date_str = str(d)
        wk_str, day_int = date_to_wk_day(d)
        if wk_str is None:
            continue
        is_today = (d == today)
        pbi, gantt = read_day(hop_ws, dt_ws, wk_str, day_int)
        # Use cached static prod data for BOL/Empty (Snowflake not queried here)
        entry = build_day_entry(date_str, wk_str, day_int, pbi, gantt, is_today)
        log(f"  {date_str} ({wk_str} D{day_int}): {len(pbi)} stations, {len(gantt)} Gantt events, {sum(pbi.values()):.0f} min DT")
        entries.append(entry)

    # Build the full DAYS_DATA block
    new_days_data = "const DAYS_DATA = {\n" + '\n'.join(entries) + "\n};"

    # Patch the HTML file
    with open(DASH, 'r', encoding='utf-8') as f:
        html = f.read()

    # Replace DAYS_DATA block using regex
    pattern = r'const DAYS_DATA = \{.*?\};'
    new_html, n = re.subn(pattern, new_days_data, html, flags=re.DOTALL)
    if n != 1:
        log(f"ERROR: Expected 1 DAYS_DATA match, found {n}. HTML not updated.")
        return False

    # Update the "isToday" marker and date in title
    with open(DASH, 'w', encoding='utf-8') as f:
        f.write(new_html)
    log(f"HTML updated: {DASH}")

    # Git push if enabled
    if GITHUB_ENABLED:
        try:
            repo = GITHUB_REPO
            cmds = [
                ['git', '-C', repo, 'add', 'body_shop_intelligence.html'],
                ['git', '-C', repo, 'commit', '-m', f'Auto-update {today} 17:00'],
                ['git', '-C', repo, 'push', '--set-upstream', 'origin', 'main'],
            ]
            for cmd in cmds:
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                if result.returncode == 0:
                    log(f"Git: {' '.join(cmd[2:])} OK")
                else:
                    log(f"Git WARNING: {result.stderr.strip()}")
        except Exception as e:
            log(f"Git ERROR: {e}")
    else:
        log("GitHub push skipped (GITHUB_ENABLED=False). Set to True after git setup.")

    # Cleanup temp files
    for tmp in [tmp_hop, tmp_dt]:
        try: os.remove(tmp)
        except: pass

    log("=== Update complete ===\n")
    return True

if __name__ == '__main__':
    success = update()
    sys.exit(0 if success else 1)
