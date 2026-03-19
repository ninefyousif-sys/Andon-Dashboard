#!/usr/bin/env python3
"""
Body Shop (RA) Andon Dashboard — Daily Auto-Update Script
Run at 17:00 Mon-Fri via Windows Task Scheduler
Reads HOP + DT Excel files, updates HTML DAYS_DATA, pushes to GitHub
"""

import openpyxl, warnings, datetime, json, re, shutil, subprocess, sys, os
warnings.filterwarnings('ignore')

# ── Optional Snowflake import (skipped gracefully if not installed) ─────────────
try:
    import snowflake.connector
    SNOWFLAKE_AVAILABLE = True
except ImportError:
    SNOWFLAKE_AVAILABLE = False

# ── CONFIGURATION ──────────────────────────────────────────────────────────────
HOP_SRC = r"C:\Users\NYOUSIF\OneDrive - Volvo Cars\A Shop Production SI and Supervisors - General\Hop Line Downtime\HOP New Downtime Breakdown.xlsm"
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

    # Keep only shift hours, sort by duration DESC so critical events are never dropped
    # then re-sort by start time for display — this way the top-40 always includes
    # the longest stops (e.g. tool changer) even if they occur late in the shift
    shift_gantt = [g for g in gantt if '06:00' <= g['s'] <= '15:00']
    top_by_dur  = sorted(shift_gantt, key=lambda x: -x['sec'])[:40]   # top 40 by duration
    gantt       = sorted(top_by_dur, key=lambda x: x['s'])             # re-sort by time for display
    return pbi, gantt

# ── SNOWFLAKE: per-window production data ──────────────────────────────────────
# Connects to VCCH.PRODUCTION_TRACKING.BODY_TRACKING
# reg 13000 = BOL (entry), reg 19900 = Empty Skid (exit)
# Timestamps are stored in US/Eastern (UTC-4/-5); we convert to Brussels CET/CEST
# Cars traverse the shop in ~5h, so scan times appear 5h after production window start.
# Windows shift: W1 prod 06-07 CET → scans ~11-12 CET, W8 prod 13:40-14:40 → scans ~19:00

# Load Snowflake credentials from local file (never committed to GitHub)
_CREDS_FILE = os.path.join(WORK, 'snowflake_credentials.json')
def _load_snowflake_cfg():
    try:
        with open(_CREDS_FILE, encoding='utf-8') as f:
            c = json.load(f)
        cfg = {
            'account':   c['account'],
            'user':      c['user'],
            'role':      c.get('role', ''),
            'database':  c.get('database', 'VCCH'),
            'schema':    'PRODUCTION_TRACKING',
            'warehouse': c.get('warehouse', 'REPORTING'),
        }
        # Key-pair auth (preferred — never expires)
        if 'private_key_file' in c:
            from cryptography.hazmat.primitives.serialization import load_pem_private_key
            from cryptography.hazmat.backends import default_backend
            with open(c['private_key_file'], 'rb') as kf:
                pk = load_pem_private_key(kf.read(), password=None, backend=default_backend())
            from cryptography.hazmat.primitives.serialization import Encoding, PrivateFormat, NoEncryption
            cfg['private_key'] = pk.private_bytes(Encoding.DER, PrivateFormat.PKCS8, NoEncryption())
        else:
            cfg['password'] = c['password']   # fallback: token/password
        return cfg
    except Exception as e:
        log(f"  Snowflake creds load error: {e}")
        return None

SNOWFLAKE_CFG = _load_snowflake_cfg()  # None if file missing

# Hours in CET at which each production window's cars appear at Empty Skid scan point
# Empirically verified: W1(06-07) → scan hr 11, W2(07-08) → 12, ... W8(13:40-14:40) → 18-19
SCAN_HOURS_PER_WIN = [11, 12, 13, 14, 15, 16, 17, 18]  # one per window

def get_production_from_snowflake(date_str):
    """Query Snowflake for per-window BOL + Empty Skid counts."""
    if not SNOWFLAKE_AVAILABLE:
        log("  Snowflake connector not installed — pip install snowflake-connector-python")
        return None
    if not SNOWFLAKE_CFG:
        log(f"  Snowflake credentials file not found at {_CREDS_FILE}")
        return None
    try:
        conn = snowflake.connector.connect(**SNOWFLAKE_CFG)
        cursor = conn.cursor()
        sql = f"""
            SELECT "registrationPoint",
                   HOUR(CONVERT_TIMEZONE('Europe/Brussels', "timestampRegistrationPoint")) AS cet_hr,
                   COUNT(*) AS cars
            FROM   VCCH.PRODUCTION_TRACKING.BODY_TRACKING
            WHERE  DATE(CONVERT_TIMEZONE('Europe/Brussels', "timestampRegistrationPoint")) = '{date_str}'
              AND  "registrationPoint" IN ('13000', '19900')
              AND  HOUR(CONVERT_TIMEZONE('Europe/Brussels', "timestampRegistrationPoint")) BETWEEN 10 AND 20
            GROUP BY 1, 2
            ORDER BY 1, 2
        """
        cursor.execute(sql)
        rows = cursor.fetchall()
        conn.close()

        bol_h_map = {r[1]: r[2] for r in rows if str(r[0]) == '13000'}
        emp_h_map = {r[1]: r[2] for r in rows if str(r[0]) == '19900'}

        # Map scan hours to 8 windows; include hour-19 cars in last window
        bol_h = [bol_h_map.get(h, 0) for h in SCAN_HOURS_PER_WIN]
        emp_h = [emp_h_map.get(h, 0) for h in SCAN_HOURS_PER_WIN]
        bol_h[-1] += bol_h_map.get(19, 0)
        emp_h[-1] += emp_h_map.get(19, 0)

        bol_tot = sum(bol_h)
        emp_tot = sum(emp_h)

        if emp_tot == 0:
            log(f"  Snowflake returned 0 Empty Skid scans for {date_str} — using STATIC_PROD")
            return None

        log(f"  Snowflake OK: BOL={bol_tot}  Empty={emp_tot}  windows={bol_h}")
        return {'bol_h': bol_h, 'empty_h': emp_h, 'bol_tot': bol_tot, 'empty_tot': emp_tot}

    except Exception as e:
        log(f"  Snowflake error: {e}")
        return None

# ── BUILD DAYS_DATA JAVASCRIPT ─────────────────────────────────────────────────
LABEL_MAP = {0:'Mon', 1:'Tue', 2:'Wed', 3:'Thu', 4:'Fri'}

# Static production data (used as fallback when Snowflake is unavailable)
# Will be replaced by Snowflake data when SNOWFLAKE_AVAILABLE=True and creds are set
STATIC_PROD = {
    '2026-03-10': {'bol_h':[8,10,9,10,4,9,11,7],  'empty_h':[10,9,9,10,3,8,10,8],  'bol_tot':70, 'empty_tot':71},
    '2026-03-11': {'bol_h':[10,9,10,10,10,10,11,3],'empty_h':[9,9,9,11,10,8,10,8],  'bol_tot':76, 'empty_tot':77},
    '2026-03-12': {'bol_h':[10,8,9,10,11,9,10,6],  'empty_h':[10,10,6,10,12,9,5,11],'bol_tot':76, 'empty_tot':75},
    '2026-03-13': {'bol_h':[11,11,10,8,10,7,10,7], 'empty_h':[11,10,8,9,7,7,12,8],  'bol_tot':77, 'empty_tot':75},
    '2026-03-16': {'bol_h':[9,12,5,11,12,8,12,9],  'empty_h':[12,12,10,12,9,10,11,9],'bol_tot':98,'empty_tot':104,'overtime':True,'otNote':'Shift ran until 16:38'},
    '2026-03-17': {'bol_h':[7,12,6,12,11,10,11,7], 'empty_h':[12,12,11,11,13,10,10,7],'bol_tot':80,'empty_tot':89},
    '2026-03-18': {'bol_h':[8,13,12,11,10,7,13,12], 'empty_h':[10,10,12,11,11,6,13,13],'bol_tot':86,'empty_tot':86},
    '2026-03-19': {'bol_h':[9,8,10,10,4,10,8,8],   'empty_h':[10,11,12,10,5,10,13,10],'bol_tot':67,'empty_tot':81},
}

def pbi_to_js(pbi):
    items = sorted(pbi.items(), key=lambda x: -x[1])
    return '{' + ','.join(f"'{k}':{v}" for k, v in items) + '}'

def gantt_to_js(gantt):
    items = []
    for g in gantt[:40]:  # up to 40 events (sorted by duration DESC then time ASC)
        items.append(
            f"{{line:'{g['line']}',s:'{g['s']}',e:'{g['e']}',sec:{g['sec']},col:'{g['col']}',lbl:'{g['lbl']}'}}"
        )
    return '[' + ',\n      '.join(items) + ']'

def build_day_entry(date_str, wk_str, day_int, pbi, gantt, is_today=False):
    d = datetime.date.fromisoformat(date_str)
    day_name = LABEL_MAP[d.weekday()]
    label = f"{day_name} {d.day}-{d.strftime('%b')}"
    # Try Snowflake first for dates not in STATIC_PROD (gives real per-window OPR data)
    prod = STATIC_PROD.get(date_str)
    if not prod:
        prod = get_production_from_snowflake(date_str)
    if not prod:
        prod = {'bol_h': [8]*8, 'empty_h': [8]*8, 'bol_tot': 80, 'empty_tot': 80}
    overtime = prod.get('overtime', False)
    ot_note  = prod.get('otNote', '')

    ot_str    = f",otNote:'{ot_note}'" if ot_note else ''
    today_str = ',isToday:true' if is_today else ''
    lines = [f"  '{date_str}':{{label:'{label}',overtime:{'true' if overtime else 'false'}{ot_str}{today_str},"]
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
                ['git', '-C', repo, 'push', 'origin', 'main'],
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
