#!/usr/bin/env python3
"""
A-Shop Morning Meeting Dashboard — Daily Update Script
Run each morning at 05:50 Mon-Fri via Windows Task Scheduler.

Sources:
  Downtime  → same OneDrive Excel files used by the Andon Dashboard
  PPT data  → OneDrive / 2026 RTM / A Shop 26WxxDx.pptx
                (Safety, Part Quality, Bodies OOF, Scrap)
"""

import openpyxl, warnings, datetime, json, re, shutil, os, sys, subprocess
warnings.filterwarnings('ignore')

# ── CONFIG ─────────────────────────────────────────────────────────────────────
BASE_ONEDRIVE = r"C:\Users\NYOUSIF\OneDrive - Volvo Cars\A Shop Production SI and Supervisors - General"

HOP_SRC = os.path.join(BASE_ONEDRIVE, r"Hop Line Downtime\HOP New Downtime Breakdown.xlsm")
DT_SRC  = os.path.join(BASE_ONEDRIVE, r"Downtime Tracker Logv6a.xlsm")
PPT_DIR = os.path.join(BASE_ONEDRIVE, r"2026 RTM")           # daily PPT folder

DASH = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\morning_meeting_dashboard.html"
WORK = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
LOG  = os.path.join(WORK, "mm_update_log.txt")

WK12_MON = datetime.date(2026, 3, 16)  # production calendar anchor

# ── TARGETS (must match TARGETS in HTML) ───────────────────────────────────────
TARGETS = {
    'ashop_ftt':  {'tgt': 98.0,  'dir': 'ge'},
    'ashop_dpv':  {'tgt': 0.08,  'dir': 'le'},
    'wg_dpv':     {'tgt': 1.2,   'dir': 'le'},
    'wd6_ftt':    {'tgt': 100.0, 'dir': 'ge'},
    'wd6_dpv':    {'tgt': 0.12,  'dir': 'le'},
    'cal_ftt':    {'tgt': 100.0, 'dir': 'ge'},
    'final1_ftt': {'tgt': 100.0, 'dir': 'ge'},
    'final2_ftt': {'tgt': 100.0, 'dir': 'ge'},
}

# ── AREA MAPPING ──────────────────────────────────────────────────────────────
HOP_AREA = {'285':'upperbody','286':'upperbody','287':'hangon','197':'hangon','198':'hangon','155':'bodysides'}
DT_AREA  = {'231':'underbody','232':'underbody','233':'underbody','234':'underbody',
            '235':'underbody','236':'underbody','237':'underbody','138':'underbody',
            '258':'bodysides','155':'bodysides'}
PLANNED_KW = {'pm','planned','planned maintenance','preventive','scheduled','maintenance pm'}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def log(msg):
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG, 'a', encoding='utf-8') as f: f.write(line+'\n')

def time_to_min(t):
    if t is None: return 0
    if isinstance(t, (datetime.datetime, datetime.time)): return t.hour*60+t.minute+t.second/60
    try:
        f = float(t); return f*1440 if 0 < f < 2 else f
    except: return 0

def dur_min(s, e):
    sm, em = time_to_min(s), time_to_min(e)
    return round(em-sm, 1) if em > sm else 0

def t_fmt(t):
    if isinstance(t, (datetime.datetime, datetime.time)): return f"{t.hour:02d}:{t.minute:02d}"
    return ''

def is_planned(resp, err):
    for v in [resp, err]:
        if v and any(k in str(v).lower() for k in PLANNED_KW): return True
    return False

def should_exclude(resp, err):
    r = str(resp).lower().strip() if resp else ''
    e = str(err).lower().strip()  if err  else ''
    if 'shop flow' in r or r in {'shop flow','shop flow '}: return True
    if 'blocked' in e or 'starved' in e: return True
    return False

def hop_code(area):
    a = str(area)
    for pre, c in [('285','285'),('286','286'),('287','287'),('198','198'),('155','155'),('197','197')]:
        if pre in a: return c
    if 'tailgate' in a.lower(): return '287'
    if 'roof'     in a.lower(): return '155'
    m = re.search(r'\b(\d{3})\b', a)
    return m.group(1) if m else None

def dt_code(area):
    a = str(area).lower()
    for pat, c in [('232','232'),('233','233'),('234','234'),('235','235'),('236','236'),
                   ('237','237'),('231','231'),('138','138'),('258','258'),('155','155'),('roof','155')]:
        if pat in a: return c
    m = re.search(r'\b(2\d{2})\b', str(area))
    return m.group(1) if m else None

def date_to_wk_day(d):
    if d.weekday() > 4: return None, None
    delta = (d - WK12_MON).days
    return f"WK{12 + delta//7}", d.weekday()+1

def prev_working_day(ref=None):
    d = (ref or datetime.date.today()) - datetime.timedelta(days=1)
    while d.weekday() > 4: d -= datetime.timedelta(days=1)
    return d

# ── DOWNTIME READING ──────────────────────────────────────────────────────────
def read_area_dt(hop_ws, dt_ws, wk_str, day_int):
    areas = {k: {'total':0,'planned':0,'unplanned':0,'events':[]}
             for k in ['underbody','upperbody','hangon','bodysides']}

    for ws, code_fn, area_map in [(hop_ws, hop_code, HOP_AREA),(dt_ws, dt_code, DT_AREA)]:
        if ws is None: continue
        for i, row in enumerate(ws.rows):
            if i == 0: continue
            vals = [c.value for c in row]
            try:    yr = int(float(str(vals[0])))
            except: continue
            if yr != 2026: continue
            if str(vals[1]).strip() != wk_str: continue
            try:    d = int(float(str(vals[2])))
            except: continue
            if d != day_int: continue

            start, end = vals[3], vals[4]
            dur = dur_min(start, end)
            if dur <= 0: continue

            area_str = str(vals[7]) if vals[7] and str(vals[7]) != 'None' else ''
            if not area_str: continue

            resp = vals[11] if len(vals) > 11 else None
            err  = vals[12] if len(vals) > 12 else None
            if should_exclude(resp, err): continue

            code     = code_fn(area_str)
            area_key = area_map.get(code)
            if not area_key: continue

            planned_flag = is_planned(resp, err)
            dur_r = round(dur, 1)
            areas[area_key]['total']     = round(areas[area_key]['total']     + dur_r, 1)
            if planned_flag: areas[area_key]['planned']  = round(areas[area_key]['planned']  + dur_r, 1)
            else:            areas[area_key]['unplanned']= round(areas[area_key]['unplanned']+ dur_r, 1)

            if len(areas[area_key]['events']) < 10:  # top 10 events per area
                cause = ' · '.join(filter(None,[str(resp).strip() if resp else '',
                                                str(err).strip()  if err  else '']))[:60] or area_str[:30]
                areas[area_key]['events'].append({
                    'station': area_str[:20], 'start': t_fmt(start), 'end': t_fmt(end),
                    'dur_min': dur_r, 'cause': cause, 'planned': planned_flag
                })
    return areas

# ── PPT READING ───────────────────────────────────────────────────────────────
def find_ppt(wk_str, day_int):
    """Find A Shop 26W12D3.pptx style file in the 2026 RTM folder."""
    if not os.path.isdir(PPT_DIR):
        log(f"WARNING: PPT folder not found: {PPT_DIR}")
        return None
    wk_num = wk_str.replace('WK','')
    fname  = f"A Shop 26W{wk_num}D{day_int}.pptx"
    # Try exact name first, then case-insensitive scan
    direct = os.path.join(PPT_DIR, fname)
    if os.path.exists(direct): return direct
    for f in os.listdir(PPT_DIR):
        if f.lower() == fname.lower():
            return os.path.join(PPT_DIR, f)
    log(f"WARNING: PPT file not found: {fname}")
    return None

def read_ppt_markdown(ppt_path):
    """Run markitdown on PPT, return full markdown text."""
    try:
        result = subprocess.run(
            [sys.executable, '-m', 'markitdown', ppt_path],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            log(f"PPT read OK: {len(result.stdout)} chars")
            return result.stdout
        else:
            log(f"WARNING: markitdown returned {result.returncode}: {result.stderr[:200]}")
            return ''
    except Exception as e:
        log(f"WARNING: Cannot read PPT: {e}")
        return ''

def parse_md_table(text):
    """Parse first markdown table in text → list of dicts."""
    rows, headers, in_table = [], [], False
    for line in text.split('\n'):
        line = line.strip()
        if not (line.startswith('|') and line.endswith('|')):
            if in_table: break
            continue
        cells = [c.strip() for c in line[1:-1].split('|')]
        if not in_table:
            headers = cells; in_table = True; continue
        if all(re.match(r'^[-: ]+$', c) for c in cells): continue  # separator
        if any(c and c != 'None' and c != '' for c in cells):
            rows.append({headers[j] if j < len(headers) else f'c{j}': cells[j]
                         for j in range(len(cells))})
    return rows

def normalize_pq(rows):
    """Map PPT Part Quality table columns → dashboard-friendly keys."""
    COL_MAP = {
        'PROD.AREA': 'area', 'PROD. AREA': 'area',
        'EX90 /723N': 'model', 'EX90 / 723N': 'model', 'EX90/723N': 'model',
        'PART DESCRIPTION': 'part',
        'SUPPLIER NAME': 'supplier',
        'PART NUMBER': 'partno',
        'HOW BIG/HOW MANY?': 'qty', 'HOW BIG/HOW MANY': 'qty',
        'STATUS': 'status',
        'REPEATER': 'repeater',
        'Sort': 'sort', 'SORT': 'sort',
        'PROBLEM STATEMENT/DETAILS': 'problem', 'PROBLEM STATEMENT': 'problem',
        'HANDSHAKE/VIRA': 'handshake', 'HANDSHAKE / VIRA': 'handshake',
    }
    result = []
    for r in rows:
        mapped = {}
        for k, v in r.items():
            for col, key in COL_MAP.items():
                if col.lower() in k.lower():
                    mapped[key] = v; break
        if mapped.get('area') or mapped.get('part'):
            result.append(mapped)
    return result

def normalize_bof(rows):
    """Map PPT Bodies OOF table columns → dashboard-friendly keys."""
    COL_MAP = {
        'Automation': 'mode', 'AUTOMATION': 'mode',
        'Body/RFID': 'rfid', 'BODY/RFID': 'rfid',
        'TYPE': 'type',
        'Is it an Underbody': 'bodytype', 'Underbody or Complete': 'bodytype',
        'Location Staged': 'location', 'LOCATION': 'location',
        'Status': 'status', 'STATUS': 'status',
        'Reason': 'reason', 'REASON': 'reason',
        'Date Removed': 'removed',
        'Expected Repair': 'expected',
        'Dummy Order': 'dummy',
        'Responsible Champion': 'champion', 'Champion': 'champion',
    }
    result = []
    for r in rows:
        mapped = {}
        for k, v in r.items():
            for col, key in COL_MAP.items():
                if col.lower() in k.lower():
                    mapped[key] = v; break
        if mapped.get('rfid') or mapped.get('reason'):
            result.append(mapped)
    return result

def parse_ppt_data(md_text):
    """Split markdown by slide and extract structured data."""
    data = {
        'safety':       {'title': '', 'detail': '', 'meta': ''},
        'part_quality': [],
        'bodies_oof':   [],
        'scrap':        '$0',
        'scrap_note':   'Data from PPT — verify with production records',
    }
    if not md_text: return data

    # Split into slide chunks
    chunks = re.split(r'<!--\s*Slide number:\s*(\d+)\s*-->', md_text)

    for i in range(1, len(chunks), 2):
        slide_num = int(chunks[i])
        content   = chunks[i+1] if i+1 < len(chunks) else ''

        # Slide 3 — Safety
        if slide_num == 3:
            lines = [l.strip() for l in content.split('\n') if l.strip() and not l.startswith('!')]
            title  = next((l for l in lines if 'ALERT' in l.upper() or 'SAFETY' in l.upper()), 'SAFETY ALERT')
            detail = next((l for l in lines if len(l) > 30 and 'ALERT' not in l.upper()
                           and not l.startswith('#') and not re.match(r'\d{1,2}/\d+',l)), '')
            date   = next((l for l in lines if re.search(r'\d{1,2}/\d{1,2}/\d{4}', l)), '')
            data['safety'] = {'title': title.strip(), 'detail': detail.strip(), 'meta': date.strip()}

        # Slide 4 — Part Quality Issues
        elif slide_num == 4:
            rows = parse_md_table(content)
            if rows:
                data['part_quality'] = normalize_pq(rows)
                log(f"  Part Quality: {len(data['part_quality'])} rows extracted")

        # Slide 6 — Bodies Out of Flow
        elif slide_num == 6:
            rows = parse_md_table(content)
            if rows:
                data['bodies_oof'] = normalize_bof(rows)
                log(f"  Bodies OOF: {len(data['bodies_oof'])} rows extracted")

        # Slide 17 — Scrap
        elif slide_num == 17:
            m = re.search(r'\$\s*[\d,]+', content)
            data['scrap'] = m.group(0).replace(' ','') if m else '$0'
            data['scrap_note'] = ('$0 — supplier warranty covers all scrapped parts'
                                  if data['scrap'] == '$0' else
                                  f"{data['scrap']} scrap cost — review breakdown below")
            log(f"  Scrap: {data['scrap']}")

    return data

# ── BUILD MM_DATA JS ──────────────────────────────────────────────────────────
def build_mm_data(areas, ppt_data, report_date, wk_str, day_int):
    SHIFT = 540
    total    = round(sum(a['total']     for a in areas.values()), 1)
    planned  = round(sum(a['planned']   for a in areas.values()), 1)
    unplanned= round(sum(a['unplanned'] for a in areas.values()), 1)
    avail    = round((SHIFT - unplanned) / SHIFT * 100, 1)

    def area_js(key):
        a = areas[key]
        evts = ','.join(
            f"{{station:{json.dumps(e['station'])},dur:{e['dur_min']},"
            f"cause:{json.dumps(e['cause'])},planned:{'true' if e['planned'] else 'false'}}}"
            for e in a['events']
        )
        return f"{{total:{a['total']},planned:{a['planned']},unplanned:{a['unplanned']},events:[{evts}]}}"

    pq_js  = json.dumps(ppt_data['part_quality'],  ensure_ascii=False)
    bof_js = json.dumps(ppt_data['bodies_oof'],     ensure_ascii=False)
    saf_js = json.dumps(ppt_data['safety'],         ensure_ascii=False)

    return (
        f"const MM_DATA = {{\n"
        f"  date: {json.dumps(str(report_date))},\n"
        f"  wk: {json.dumps(wk_str)},\n"
        f"  day: {day_int},\n"
        f"  downtime: {{\n"
        f"    total_min: {total}, planned_min: {planned},\n"
        f"    unplanned_min: {unplanned}, availability: {avail},\n"
        f"    areas: {{\n"
        f"      underbody: {area_js('underbody')},\n"
        f"      upperbody: {area_js('upperbody')},\n"
        f"      hangon:    {area_js('hangon')},\n"
        f"      bodysides: {area_js('bodysides')}\n"
        f"    }}\n"
        f"  }},\n"
        f"  ppt: {{\n"
        f"    safety: {saf_js},\n"
        f"    part_quality: {pq_js},\n"
        f"    bodies_oof: {bof_js},\n"
        f"    scrap: {json.dumps(ppt_data['scrap'])},\n"
        f"    scrap_note: {json.dumps(ppt_data['scrap_note'])}\n"
        f"  }}\n"
        f"}};"
    )

# ── PATCH HTML ────────────────────────────────────────────────────────────────
def patch_html(new_mm_js):
    with open(DASH, 'r', encoding='utf-8') as f:
        html = f.read()
    new_html, n = re.subn(r'const MM_DATA = \{.*?\};', new_mm_js, html, flags=re.DOTALL)
    if n != 1:
        log(f"ERROR: Expected 1 MM_DATA block, got {n}")
        return False
    with open(DASH, 'w', encoding='utf-8') as f:
        f.write(new_html)
    log(f"HTML patched: {DASH}")
    return True

# ── MAIN ──────────────────────────────────────────────────────────────────────
def update():
    today       = datetime.date.today()
    report_date = prev_working_day(today)
    wk_str, day_int = date_to_wk_day(report_date)

    log(f"=== Morning Meeting update started ({today}) ===")
    log(f"Report date: {report_date}  ({wk_str} D{day_int})")

    if wk_str is None:
        log("ERROR: report date is a weekend"); return False

    # ── 1. Read downtime from Excel ──
    tmp_hop = os.path.join(WORK, '_mm_hop.xlsm')
    tmp_dt  = os.path.join(WORK, '_mm_dt.xlsm')
    for src, dst, name in [(HOP_SRC, tmp_hop,'HOP'),(DT_SRC, tmp_dt,'DT')]:
        if os.path.exists(src):
            shutil.copy2(src, dst)
            log(f"Copied {name}: {os.path.getsize(dst)//1024} KB")
        else:
            log(f"WARNING: {name} not found: {src}")

    hop_ws = dt_ws = None
    try:
        wb = openpyxl.load_workbook(tmp_hop, data_only=True, read_only=True)
        hop_ws = wb['New(DT LOG)']; log("Opened HOP workbook")
    except Exception as e: log(f"HOP error: {e}")
    try:
        wb = openpyxl.load_workbook(tmp_dt, data_only=True, read_only=True)
        dt_ws = wb['New(DT LOG)']; log("Opened DT workbook")
    except Exception as e: log(f"DT error: {e}")

    areas = read_area_dt(hop_ws, dt_ws, wk_str, day_int)
    for k, a in areas.items():
        log(f"  {k:12s}: total={a['total']} min  planned={a['planned']}  unplanned={a['unplanned']}")

    # ── 2. Read PPT ──
    ppt_path = find_ppt(wk_str, day_int)
    if ppt_path:
        log(f"Found PPT: {ppt_path}")
        md_text  = read_ppt_markdown(ppt_path)
        ppt_data = parse_ppt_data(md_text)
    else:
        log("WARNING: Using empty PPT data (file not found)")
        ppt_data = {'safety':{'title':'','detail':'','meta':''},
                    'part_quality':[], 'bodies_oof':[], 'scrap':'$0',
                    'scrap_note':'PPT file not found — update manually'}

    # ── 3. Build + patch ──
    mm_js = build_mm_data(areas, ppt_data, report_date, wk_str, day_int)
    success = patch_html(mm_js)

    for tmp in [tmp_hop, tmp_dt]:
        try: os.remove(tmp)
        except: pass

    log("=== Morning Meeting update complete ===\n")
    return success

if __name__ == '__main__':
    sys.exit(0 if update() else 1)
