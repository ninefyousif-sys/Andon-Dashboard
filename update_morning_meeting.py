#!/usr/bin/env python3
"""
A-Shop Morning Meeting Dashboard — Daily Update Script
Run each morning at 05:50 Mon-Fri via Windows Task Scheduler.

Sources:
  Power BI Desktop  → FTT, DPV, W&G DPV, item details (DAX queries)
  Downtime Excel    → HOP + DT area downtime (same as Andon Dashboard)
  PPT 2026 RTM      → Safety, Part Quality, Bodies OOF, Scrap (tables only)
"""

import openpyxl, warnings, datetime, json, re, shutil, os, sys, subprocess
warnings.filterwarnings('ignore')

# ── CONFIG ─────────────────────────────────────────────────────────────────────
BASE_ONEDRIVE = r"C:\Users\NYOUSIF\OneDrive - Volvo Cars\A Shop Production SI and Supervisors - General"

HOP_SRC = os.path.join(BASE_ONEDRIVE, r"Hop Line Downtime\HOP New Downtime Breakdown.xlsm")
DT_SRC  = os.path.join(BASE_ONEDRIVE, r"Downtime Tracker Logv6a.xlsm")
PPT_DIR = os.path.join(BASE_ONEDRIVE, r"2026 RTM")

DASH        = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\morning_meeting_dashboard.html"
DASH_MOBILE = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\morning_meeting_mobile.html"
WORK        = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard"
LOG         = os.path.join(WORK, "mm_update_log.txt")

GITHUB_REPO    = WORK  # same folder — git repo is here
GITHUB_ENABLED = True  # pushes morning_meeting_mobile.html to GitHub Pages

WK12_MON = datetime.date(2026, 3, 16)

# ── TARGETS ────────────────────────────────────────────────────────────────────
TARGETS = {
    'bok_opr':    {'tgt': 98.7,  'dir': 'ge'},
    'bol_opr':    {'tgt': 95.0,  'dir': 'ge'},
    'ashop_ftt':  {'tgt': 98.0,  'dir': 'ge'},
    'ashop_dpv':  {'tgt': 0.10,  'dir': 'le'},
    'wg_dpv':     {'tgt': 1.20,  'dir': 'le'},
    'wd6_ftt':    {'tgt': 100.0, 'dir': 'ge'},
    'wd6_dpv':    {'tgt': 0.12,  'dir': 'le'},
    'cal_ftt':    {'tgt': 100.0, 'dir': 'ge'},
    'final1_ftt': {'tgt': 100.0, 'dir': 'ge'},
    'final2_ftt': {'tgt': 100.0, 'dir': 'ge'},
    'scrap_car':  {'tgt': 2.90,  'dir': 'le'},
}

# ── AREA MAPPING (Downtime) ────────────────────────────────────────────────────
HOP_AREA = {'285':'upperbody','286':'upperbody','287':'hangon','197':'hangon','198':'hangon','155':'bodysides'}
DT_AREA  = {'231':'underbody','232':'underbody','233':'underbody','234':'underbody',
            '235':'underbody','236':'underbody','237':'underbody','138':'underbody',
            '258':'bodysides','155':'bodysides'}
PLANNED_KW = {'pm','planned','planned maintenance','preventive','scheduled','maintenance pm'}

# ── HELPERS ────────────────────────────────────────────────────────────────────
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

def dt_str(t):
    if isinstance(t, datetime.datetime): return t.strftime('%m/%d/%Y %I:%M %p')
    if isinstance(t, datetime.date): return t.strftime('%m/%d/%Y')
    return str(t) if t else ''

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


# ═══════════════════════════════════════════════════════════════════════════════
# POWER BI INTEGRATION
# Queries the open Power BI Desktop model via its local XMLA endpoint.
# Requires: pip install pyadomd  (wraps ADOMD.NET — ships with Power BI Desktop)
# ═══════════════════════════════════════════════════════════════════════════════

def find_pbi_ports():
    """Find ALL Analysis Services ports exposed by running Power BI Desktop instances.
    Returns a list of integer ports (may contain multiple if several PBI files are open)."""
    ports = []

    # Method 1: scan AnalysisServicesWorkspaces folder — PBI writes one port file per instance
    appdata = os.environ.get('LOCALAPPDATA', '')
    ws_root = os.path.join(appdata, 'Microsoft', 'Power BI Desktop',
                           'AnalysisServicesWorkspaces')
    if os.path.isdir(ws_root):
        for folder in sorted(os.listdir(ws_root)):
            port_file = os.path.join(ws_root, folder, 'Data', 'msmdsrv.port.txt')
            if os.path.exists(port_file):
                try:
                    port = int(open(port_file).read().strip())
                    if port not in ports:
                        ports.append(port)
                        log(f"Found PBI XMLA port: {port} (from {folder})")
                except: pass

    if ports:
        return ports

    # Method 2: PowerShell Get-WmiObject — collect ALL msmdsrv pids then ALL their ports
    try:
        ps_out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             'Get-WmiObject Win32_Process | Where-Object {$_.Name -eq "msmdsrv.exe"} | Select-Object -ExpandProperty ProcessId'],
            text=True, timeout=10)
        pids = set(re.findall(r'\d{4,6}', ps_out))
        if pids:
            net_out = subprocess.check_output(['netstat', '-ano'], text=True, timeout=10)
            for line in net_out.splitlines():
                parts = line.split()
                if len(parts) >= 5 and parts[4] in pids and 'LISTENING' in line:
                    m = re.search(r':(\d{4,6})\s', line)
                    if m:
                        port = int(m.group(1))
                        if port not in ports:
                            ports.append(port)
                            log(f"Found PBI XMLA port via PowerShell/netstat: {port}")
    except: pass

    if ports:
        return ports

    # Method 3: fallback wmic scan (older Windows)
    try:
        out = subprocess.check_output(['netstat','-ano'], text=True, timeout=10)
        pids_raw = subprocess.check_output(
            ['wmic','process','where','name="msmdsrv.exe"','get','ProcessId'],
            text=True, timeout=10)
        pids = set(re.findall(r'\d{4,6}', pids_raw))
        for line in out.splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[4] in pids and 'LISTENING' in line:
                m = re.search(r':(\d{4,6})\s', line)
                if m:
                    port = int(m.group(1))
                    if port not in ports:
                        ports.append(port)
                        log(f"Found PBI XMLA port via wmic/netstat: {port}")
    except: pass

    if not ports:
        log("WARNING: Power BI Desktop not found / not running")
    return ports


def find_pbi_port():
    """Compatibility shim — returns first port found (use find_pbi_ports() for multi-instance)."""
    ports = find_pbi_ports()
    return ports[0] if ports else None


def run_dax(port, dax):
    """Run a DAX query against Power BI Desktop. Returns list of dicts."""
    try:
        import pyadomd
    except ImportError:
        log("WARNING: pyadomd not installed. Run: pip install pyadomd")
        return []
    try:
        conn_str = f"Provider=MSOLAP;Data Source=localhost:{port};"
        with pyadomd.Pyadomd(conn_str) as conn:
            with conn.cursor().execute(dax) as cur:
                cols = [c.name.split('[')[-1].rstrip(']') for c in cur.description]
                return [dict(zip(cols, row)) for row in cur.fetchall()]
    except Exception as e:
        log(f"DAX error: {e}\nQuery: {dax[:120]}...")
        return []


def fmt_dt(v):
    """Format a datetime value from DAX result to string."""
    if v is None: return ''
    if hasattr(v, 'strftime'): return v.strftime('%m/%d/%Y %I:%M %p')
    return str(v)


def query_powerbi(report_date):
    """
    Query Power BI Desktop for all KPI data for report_date.
    Tries ALL running PBI instances and uses the first one that has A Shop data.
    Returns dict with kpis + ppt item arrays, or None if PBI not available.
    """
    ports = find_pbi_ports()
    if not ports:
        return None

    yr, mo, dy = report_date.year, report_date.month, report_date.day

    # Probe each port: use the first one that returns A Shop rows for this date
    port = None
    probe_dax = f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('A Shop Body Count-FTT/DPV',
                 'A Shop Body Count-FTT/DPV'[Date] = DATE({yr},{mo},{dy})),
          "Model", 'A Shop Body Count-FTT/DPV'[Model])"""
    for p in ports:
        probe = run_dax(p, probe_dax)
        if probe:
            port = p
            log(f"Using PBI port {port} — has {len(probe)} A Shop row(s) for {report_date}")
            break
        else:
            log(f"Port {p} — no A Shop data for {report_date}, trying next...")

    if port is None:
        # No port returned data for this specific date — fall back to first available port
        port = ports[0]
        log(f"No port had data for {report_date}; defaulting to port {port}")

    log(f"Querying Power BI Desktop (port {port}) for {report_date}...")

    # ── A Shop summary ────────────────────────────────────────────────────────
    ashop_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('A Shop Body Count-FTT/DPV',
                 'A Shop Body Count-FTT/DPV'[Date] = DATE({yr},{mo},{dy})),
          "Model",        'A Shop Body Count-FTT/DPV'[Model],
          "SentToPaint",  'A Shop Body Count-FTT/DPV'[Sent to Paint],
          "SentToRepair", 'A Shop Body Count-FTT/DPV'[Sent to Repair],
          "FTT",          'A Shop Body Count-FTT/DPV'[FTT],
          "DefectCount",  'A Shop Body Count-FTT/DPV'[Defect Count],
          "DPV",          'A Shop Body Count-FTT/DPV'[DPV],
          "WG_DPV",       'A Shop Body Count-FTT/DPV'[W&G DPV],
          "WG_Defects",   'A Shop Body Count-FTT/DPV'[W&G Defect Count],
          "HOP_DPV",      'A Shop Body Count-FTT/DPV'[HOP DPV],
          "HOP_Defects",  'A Shop Body Count-FTT/DPV'[HOP Defect Count]
        )
    """)
    log(f"  A Shop summary: {len(ashop_rows)} model rows")

    # ── A Shop FTT defect items ────────────────────────────────────────────────
    ashop_ftt_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('A Shop Defects',
                 'A Shop Defects'[Date] = DATE({yr},{mo},{dy})),
          "body",       'A Shop Defects'[Body Number],
          "rfid",       'A Shop Defects'[RFID],
          "desc",       'A Shop Defects'[Item Description],
          "model",      'A Shop Defects'[Model],
          "link_stn",   'A Shop Defects'[Linking Station],
          "link_time",  'A Shop Defects'[Link Time],
          "close_stn",  'A Shop Defects'[Closing Station],
          "close_time", 'A Shop Defects'[Close Time],
          "location",   'A Shop Defects'[Location],
          "extra",      'A Shop Defects'[Extra Info]
        )
    """)
    log(f"  A Shop FTT items: {len(ashop_ftt_rows)}")

    # Format datetimes
    for r in ashop_ftt_rows:
        r['link_time']  = fmt_dt(r.get('link_time'))
        r['close_time'] = fmt_dt(r.get('close_time'))

    # ── B Shop (WD6) summary ──────────────────────────────────────────────────
    # Date column in B Shop is a String — filter by string match
    date_str = report_date.strftime('%Y-%m-%d')
    bshop_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('B Shop Body Count-FTT/DPV',
                 LEFT('B Shop Body Count-FTT/DPV'[Date], 10) = "{date_str}"),
          "Model",      'B Shop Body Count-FTT/DPV'[Model],
          "BodyOK",     'B Shop Body Count-FTT/DPV'[Body OK],
          "BIWOut",     'B Shop Body Count-FTT/DPV'[WD6 BIW Out],
          "WD6_FTT",    'B Shop Body Count-FTT/DPV'[WD6 FTT],
          "Defects",    'B Shop Body Count-FTT/DPV'[Defect Count],
          "WD6_DPV",    'B Shop Body Count-FTT/DPV'[WD6 DPV]
        )
    """)
    log(f"  B Shop summary: {len(bshop_rows)} model rows")

    # ── WD6 FTT items ─────────────────────────────────────────────────────────
    wd6_ftt_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('WD6 FTT Items',
                 'WD6 FTT Items'[Date] = DATE({yr},{mo},{dy})),
          "body",       'WD6 FTT Items'[Body Number],
          "rfid",       'WD6 FTT Items'[RFID],
          "desc",       'WD6 FTT Items'[Item Description],
          "model",      'WD6 FTT Items'[Model],
          "link_stn",   'WD6 FTT Items'[Linking Station],
          "link_time",  'WD6 FTT Items'[Link Time],
          "close_stn",  'WD6 FTT Items'[Closing Station],
          "close_time", 'WD6 FTT Items'[Close Time],
          "location",   'WD6 FTT Items'[Location],
          "extra",      'WD6 FTT Items'[Extra Info]
        )
    """)
    log(f"  WD6 FTT items: {len(wd6_ftt_rows)}")
    for r in wd6_ftt_rows:
        r['link_time']  = fmt_dt(r.get('link_time'))
        r['close_time'] = fmt_dt(r.get('close_time'))

    # ── WD6 DPV items ─────────────────────────────────────────────────────────
    wd6_dpv_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('WD6 DPV Items',
                 'WD6 DPV Items'[Date] = DATE({yr},{mo},{dy})),
          "body",      'WD6 DPV Items'[Body Number],
          "rfid",      'WD6 DPV Items'[RFID],
          "desc",      'WD6 DPV Items'[Item Description],
          "model",     'WD6 DPV Items'[Model],
          "station",   'WD6 DPV Items'[Linking Station],
          "location",  'WD6 DPV Items'[Location],
          "extra",     'WD6 DPV Items'[Extra Info]
        )
    """)
    log(f"  WD6 DPV items: {len(wd6_dpv_rows)}")

    # ── C Shop summary ─────────────────────────────────────────────────────────
    cshop_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('C Shop Body Count-FTT',
                 'C Shop Body Count-FTT'[Date] = DATE({yr},{mo},{dy})),
          "Model",       'C Shop Body Count-FTT'[Model],
          "CAL_OK",      'C Shop Body Count-FTT'[CAL Body OK],
          "FN1_OK",      'C Shop Body Count-FTT'[Final 1 Body OK],
          "FN2_OK",      'C Shop Body Count-FTT'[Final 2 Body OK],
          "CAL_FTT",     'C Shop Body Count-FTT'[CAL FTT],
          "Final1_FTT",  'C Shop Body Count-FTT'[Final 1 FTT],
          "EOL_FTT",     'C Shop Body Count-FTT'[EOL FTT],
          "FTT_CAL",     'C Shop Body Count-FTT'[FTT CAL],
          "FTT_FN1",     'C Shop Body Count-FTT'[FTT FINAL1],
          "FTT_EOL",     'C Shop Body Count-FTT'[FTT EOL],
          "CAL_DPV",     'C Shop Body Count-FTT'[CAL DPV],
          "FN1_DPV",     'C Shop Body Count-FTT'[FN1 DPV],
          "FN2_DPV",     'C Shop Body Count-FTT'[FN2 DPV]
        )
    """)
    log(f"  C Shop summary: {len(cshop_rows)} model rows")

    # ── C Shop defect items (CAL + Final line items) ───────────────────────────
    cshop_defect_rows = run_dax(port, f"""
        EVALUATE SELECTCOLUMNS(
          FILTER('C Shop Defects Linked',
                 YEAR('C Shop Defects Linked'[Link Time]) = {yr} &&
                 MONTH('C Shop Defects Linked'[Link Time]) = {mo} &&
                 DAY('C Shop Defects Linked'[Link Time]) = {dy}),
          "body",       'C Shop Defects Linked'[Body Number],
          "rfid",       'C Shop Defects Linked'[RFID],
          "desc",       'C Shop Defects Linked'[Item Description],
          "model",      'C Shop Defects Linked'[Model],
          "link_stn",   'C Shop Defects Linked'[Linking Station],
          "link_time",  'C Shop Defects Linked'[Link Time],
          "close_stn",  'C Shop Defects Linked'[Closing Station],
          "close_time", 'C Shop Defects Linked'[Close Time],
          "location",   'C Shop Defects Linked'[Location],
          "extra",      'C Shop Defects Linked'[Extra Info Link]
        )
    """)
    log(f"  C Shop defect items: {len(cshop_defect_rows)}")
    for r in cshop_defect_rows:
        r['link_time']  = fmt_dt(r.get('link_time'))
        r['close_time'] = fmt_dt(r.get('close_time'))

    return {
        'ashop':         ashop_rows,
        'ashop_ftt':     ashop_ftt_rows,
        'bshop':         bshop_rows,
        'wd6_ftt':       wd6_ftt_rows,
        'wd6_dpv':       wd6_dpv_rows,
        'cshop':         cshop_rows,
        'cshop_defects': cshop_defect_rows,
    }


def build_kpis_from_pbi(pbi):
    """Build the kpis dict for MM_DATA from Power BI query results."""
    kpis = {k: {'val': None} for k in TARGETS}

    # ── A Shop ────────────────────────────────────────────────────────────────
    if pbi['ashop']:
        # Aggregate across models (EX90 + PSTR)
        total_paint  = sum(r.get('SentToPaint',0)  or 0 for r in pbi['ashop'])
        total_repair = sum(r.get('SentToRepair',0) or 0 for r in pbi['ashop'])
        total_eol    = total_paint + total_repair
        total_def    = sum(r.get('DefectCount',0)  or 0 for r in pbi['ashop'])
        wg_def       = sum(r.get('WG_Defects',0)   or 0 for r in pbi['ashop'])

        ftt_val = round(total_paint / total_eol * 100, 4) if total_eol > 0 else None
        dpv_val = round(total_def  / total_eol, 4)        if total_eol > 0 else None
        wg_dpv  = round(wg_def     / total_eol, 4)        if total_eol > 0 else None

        kpis['ashop_ftt'] = {'val': ftt_val, 'eol': total_eol, 'ok': total_paint, 'repair': total_repair}
        kpis['ashop_dpv'] = {'val': dpv_val, 'bodies': total_eol, 'defects': int(total_def)}
        kpis['wg_dpv']    = {'val': wg_dpv,  'bodies': total_eol, 'defects': int(wg_def)}

    # ── B Shop (WD6) ──────────────────────────────────────────────────────────
    if pbi['bshop']:
        total_ok    = sum(r.get('BodyOK',0)  or 0 for r in pbi['bshop'])
        total_biw   = sum(r.get('BIWOut',0)  or 0 for r in pbi['bshop'])
        total_def   = sum(r.get('Defects',0) or 0 for r in pbi['bshop'])
        total_repair_wd6 = total_biw - total_ok if total_biw > total_ok else 0
        ftt_wd6 = round(total_ok / total_biw * 100, 4) if total_biw > 0 else None
        dpv_wd6 = round(total_def / total_biw, 4)       if total_biw > 0 else None
        kpis['wd6_ftt'] = {'val': ftt_wd6, 'eol': total_biw, 'ok': total_ok, 'repair': total_repair_wd6}
        kpis['wd6_dpv'] = {'val': dpv_wd6, 'bodies': total_biw, 'defects': int(total_def)}

    # ── C Shop ────────────────────────────────────────────────────────────────
    if pbi['cshop']:
        # Aggregate CAL, Final 1, Final 2 across models
        cal_ok  = sum(r.get('CAL_OK',0)  or 0 for r in pbi['cshop'])
        fn1_ok  = sum(r.get('FN1_OK',0)  or 0 for r in pbi['cshop'])
        fn2_ok  = sum(r.get('FN2_OK',0)  or 0 for r in pbi['cshop'])
        cal_eol = sum(r.get('FTT_CAL',0) or 0 for r in pbi['cshop'])
        fn1_eol = sum(r.get('FTT_FN1',0) or 0 for r in pbi['cshop'])
        fn2_eol = sum(r.get('FTT_EOL',0) or 0 for r in pbi['cshop'])

        cal_ftt = round(cal_ok / cal_eol * 100, 4) if cal_eol > 0 else None
        fn1_ftt = round(fn1_ok / fn1_eol * 100, 4) if fn1_eol > 0 else None
        fn2_ftt = round(fn2_ok / fn2_eol * 100, 4) if fn2_eol > 0 else None

        # Final 2 EOL = sum of FTT_EOL (total bodies at EOL gate)
        kpis['cal_ftt']    = {'val': cal_ftt, 'eol': cal_eol, 'ok': cal_ok,  'repair': cal_eol - cal_ok}
        kpis['final1_ftt'] = {'val': fn1_ftt, 'eol': fn1_eol, 'ok': fn1_ok,  'repair': fn1_eol - fn1_ok}
        kpis['final2_ftt'] = {'val': fn2_ftt, 'eol': fn2_eol, 'ok': fn2_ok,  'repair': fn2_eol - fn2_ok}

    return kpis


def build_ppt_items_from_pbi(pbi):
    """Build the ppt item arrays for MM_DATA from Power BI query results."""
    def split_by_model(rows, model_col='model'):
        e536 = [r for r in rows if '536' in str(r.get('model','')) or 'EX90' in str(r.get('model',''))]
        e519 = [r for r in rows if '519' in str(r.get('model','')) or 'PSTR' in str(r.get('model',''))]
        # If no split by model, put all in 536
        if not e536 and not e519: e536 = rows
        return e536, e519

    ashop_536, ashop_519 = split_by_model(pbi['ashop_ftt'])
    wd6_ftt_536, wd6_ftt_519 = split_by_model(pbi['wd6_ftt'])
    wd6_dpv_536, wd6_dpv_519 = split_by_model(pbi['wd6_dpv'])

    # C Shop: split defects by linking station prefix to CAL vs Final
    cal_defects = [r for r in pbi['cshop_defects']
                   if any(x in str(r.get('link_stn','')).upper()
                          for x in ['CAL','EXCAL','EX-CAL','CALCAL'])]
    fn_defects  = [r for r in pbi['cshop_defects'] if r not in cal_defects]
    cal_536, cal_519   = split_by_model(cal_defects)
    fn2_536, fn2_519   = split_by_model(fn_defects)

    def to_ftt_item(r):
        return {
            'body':       str(r.get('body','—')),
            'rfid':       str(r.get('rfid','—')),
            'desc':       str(r.get('desc','')),
            'model':      str(r.get('model','')),
            'link_stn':   str(r.get('link_stn','')),
            'link_time':  str(r.get('link_time','')),
            'close_stn':  str(r.get('close_stn','')),
            'close_time': str(r.get('close_time','')),
            'location':   str(r.get('location','')),
            'extra':      str(r.get('extra','') or ''),
        }

    def to_dpv_item(r):
        return {
            'body':     str(r.get('body','—')),
            'rfid':     str(r.get('rfid','—')),
            'count':    1,
            'desc':     str(r.get('desc','')),
            'model':    str(r.get('model','')),
            'station':  str(r.get('station',r.get('link_stn',''))),
            'location': str(r.get('location','')),
            'extra':    str(r.get('extra','') or ''),
        }

    return {
        'ashop_ftt_536': [to_ftt_item(r) for r in ashop_536],
        'ashop_ftt_519': [to_ftt_item(r) for r in ashop_519],
        'wd6_ftt_536':   [to_ftt_item(r) for r in wd6_ftt_536],
        'wd6_ftt_519':   [to_ftt_item(r) for r in wd6_ftt_519],
        'wd6_dpv_536':   [to_dpv_item(r) for r in wd6_dpv_536],
        'wd6_dpv_519':   [to_dpv_item(r) for r in wd6_dpv_519],
        'cal_ftt_536':   [to_ftt_item(r) for r in cal_536],
        'cal_ftt_519':   [to_ftt_item(r) for r in cal_519],
        'final2_ftt_536':[to_ftt_item(r) for r in fn2_536],
        'final2_ftt_519':[to_ftt_item(r) for r in fn2_519],
    }


# ═══════════════════════════════════════════════════════════════════════════════
# DOWNTIME READING (Andon Downtime Sheet — same as Andon Dashboard)
# ═══════════════════════════════════════════════════════════════════════════════

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

            if len(areas[area_key]['events']) < 10:
                cause = ' · '.join(filter(None,[str(resp).strip() if resp else '',
                                                str(err).strip()  if err  else '']))[:60] or area_str[:30]
                areas[area_key]['events'].append({
                    'station': area_str[:20], 'start': t_fmt(start), 'end': t_fmt(end),
                    'dur_min': dur_r, 'cause': cause, 'planned': planned_flag
                })
    return areas


def build_hop_stops(hop_ws, dt_ws, wk_str, day_int):
    """Build top-10 HOP line stops sorted by duration (for BOK OPR display)."""
    stops = []
    area_label = {**{v: v.title() for v in HOP_AREA.values()},
                  **{v: v.title() for v in DT_AREA.values()}}
    area_label['underbody'] = 'Underbody'
    area_label['upperbody'] = 'Upperbody'
    area_label['hangon']    = 'Hang-On'
    area_label['bodysides'] = 'Body Sides'

    for ws, code_fn, area_map in [(hop_ws, hop_code, HOP_AREA),(dt_ws, dt_code, DT_AREA)]:
        if ws is None: continue
        for i, row in enumerate(ws.rows):
            if i == 0: continue
            vals = [c.value for c in row]
            try:    yr2 = int(float(str(vals[0])))
            except: continue
            if yr2 != 2026: continue
            if str(vals[1]).strip() != wk_str: continue
            try:    d = int(float(str(vals[2])))
            except: continue
            if d != day_int: continue

            start, end = vals[3], vals[4]
            dur = dur_min(start, end)
            if dur <= 0: continue

            area_str = str(vals[7]) if vals[7] else ''
            if not area_str: continue

            resp = vals[11] if len(vals) > 11 else None
            err  = vals[12] if len(vals) > 12 else None
            if should_exclude(resp, err): continue

            code     = code_fn(area_str)
            area_key = area_map.get(code, 'other')
            planned  = is_planned(resp, err)
            cause    = ' · '.join(filter(None,[str(resp).strip() if resp else '',
                                               str(err).strip()  if err  else '']))[:60] or area_str[:25]
            stops.append({
                'line':    area_str[:20],
                'area':    area_label.get(area_key, 'Other'),
                'dur_min': round(dur, 1),
                'cause':   cause,
                'planned': planned,
            })

    # Sort by duration DESC, deduplicate by line name keeping max duration
    seen = {}
    for s in sorted(stops, key=lambda x: -x['dur_min']):
        if s['line'] not in seen:
            seen[s['line']] = s
    return sorted(seen.values(), key=lambda x: -x['dur_min'])[:10]


# ═══════════════════════════════════════════════════════════════════════════════
# PPT READING (Safety, Part Quality, Bodies OOF, Scrap — table slides only)
# ═══════════════════════════════════════════════════════════════════════════════

def find_ppt(wk_str, day_int):
    if not os.path.isdir(PPT_DIR): return None
    wk_num = wk_str.replace('WK','')
    fname  = f"A Shop 26W{wk_num}D{day_int}.pptx"
    direct = os.path.join(PPT_DIR, fname)
    if os.path.exists(direct): return direct
    for f in os.listdir(PPT_DIR):
        if f.lower() == fname.lower():
            return os.path.join(PPT_DIR, f)
    log(f"WARNING: PPT not found: {fname}")
    return None

def read_ppt_markdown(ppt_path):
    try:
        result = subprocess.run([sys.executable,'-m','markitdown',ppt_path],
                                capture_output=True, text=True, timeout=120)
        return result.stdout if result.returncode == 0 else ''
    except Exception as e:
        log(f"WARNING: PPT read error: {e}")
        return ''

def parse_md_table(text):
    rows, headers, in_table = [], [], False
    for line in text.split('\n'):
        line = line.strip()
        if not (line.startswith('|') and line.endswith('|')):
            if in_table: break
            continue
        cells = [c.strip() for c in line[1:-1].split('|')]
        if not in_table:
            headers = cells; in_table = True; continue
        if all(re.match(r'^[-: ]+$', c) for c in cells): continue
        if any(c and c != 'None' for c in cells):
            rows.append({headers[j] if j < len(headers) else f'c{j}': cells[j]
                         for j in range(len(cells))})
    return rows

def parse_ppt_tables(md_text):
    data = {
        'safety':       {'title':'','detail':'','meta':''},
        'part_quality': [],
        'bodies_oof':   [],
        'scrap':        '$0',
        'scrap_note':   'Both scrapped parts covered by supplier warranty — zero cost to Volvo Cars',
    }
    if not md_text: return data
    chunks = re.split(r'<!--\s*Slide number:\s*(\d+)\s*-->', md_text)
    for i in range(1, len(chunks), 2):
        slide_num = int(chunks[i])
        content   = chunks[i+1] if i+1 < len(chunks) else ''
        if slide_num == 3:
            lines = [l.strip() for l in content.split('\n') if l.strip() and not l.startswith('!')]
            title  = next((l for l in lines if 'ALERT' in l.upper() or 'SAFETY' in l.upper()), '')
            detail = next((l for l in lines if len(l) > 30 and 'ALERT' not in l.upper()
                           and not l.startswith('#') and not re.match(r'\d{1,2}/\d+',l)), '')
            date   = next((l for l in lines if re.search(r'\d{1,2}/\d{1,2}/\d{4}',l)), '')
            data['safety'] = {'title':title.strip(),'detail':detail.strip(),'meta':date.strip()}
        elif slide_num == 4:
            rows = parse_md_table(content)
            if rows:
                data['part_quality'] = [
                    {'area':    r.get('PROD.AREA',r.get('PROD. AREA','')),
                     'model':   r.get('EX90 /723N',r.get('EX90/723N','')),
                     'part':    r.get('PART DESCRIPTION',''),
                     'supplier':r.get('SUPPLIER NAME',''),
                     'partno':  r.get('PART NUMBER',''),
                     'qty':     r.get('HOW BIG/HOW MANY?',r.get('HOW BIG/HOW MANY','')),
                     'status':  r.get('STATUS',''),
                     'repeater':r.get('REPEATER',''),
                     'sort':    r.get('Sort',r.get('SORT','')),
                     'problem': r.get('PROBLEM STATEMENT/DETAILS',r.get('PROBLEM STATEMENT','')),
                     'handshake':r.get('HANDSHAKE/VIRA',r.get('HANDSHAKE / VIRA','')),}
                    for r in rows if r.get('PART DESCRIPTION')]
        elif slide_num == 6:
            rows = parse_md_table(content)
            if rows:
                data['bodies_oof'] = [
                    {'mode':     r.get('Automation /Manual',r.get('AUTOMATION /MANUAL','')),
                     'rfid':     r.get('Body/RFID​',r.get('BODY/RFID','')),
                     'type':     next((v for k,v in r.items() if 'TYPE' in k.upper() and v), ''),
                     'bodytype': next((v for k,v in r.items() if 'UNDERBODY' in k.upper() or 'COMPLETE' in k.upper() and v),'Complete'),
                     'location': r.get('Location Staged​',r.get('LOCATION STAGED','')),
                     'status':   r.get('Status​',r.get('STATUS','')),
                     'reason':   r.get('Reason​',r.get('REASON','')),
                     'removed':  r.get('Date Removed from Line​WKxxDxx​',r.get('DATE REMOVED','')),
                     'expected': r.get('Expected Repair Date​WKxxDxx​',r.get('EXPECTED REPAIR','')),
                     'dummy':    r.get('Dummy Order Entered?​"Yes or No"​',r.get('DUMMY ORDER','')),
                     'champion': r.get('Responsible Champion​',r.get('CHAMPION','')),}
                    for r in rows if r.get('Body/RFID​') or r.get('BODY/RFID','')]
        elif slide_num == 17:
            m = re.search(r'\$\s*[\d,]+', content)
            scrap_val = m.group(0).replace(' ','') if m else '$0'
            data['scrap'] = scrap_val
            data['scrap_note'] = ('Both scrapped parts covered by supplier warranty — zero cost to Volvo Cars'
                                  if scrap_val == '$0' else
                                  f"{scrap_val} scrap cost — review breakdown with team")
    return data


# ═══════════════════════════════════════════════════════════════════════════════
# MM_DATA BUILDER
# ═══════════════════════════════════════════════════════════════════════════════

def build_mm_data(areas, hop_stops, kpis, ppt_items, ppt_tables, report_date, wk_str, day_int):
    SHIFT = 540
    total     = round(sum(a['total']     for a in areas.values()), 1)
    planned   = round(sum(a['planned']   for a in areas.values()), 1)
    unplanned = round(sum(a['unplanned'] for a in areas.values()), 1)
    avail     = round((SHIFT - unplanned) / SHIFT * 100, 1)

    def area_js(key):
        a = areas[key]
        evts = json.dumps(a['events'], default=str)
        return f'{{"total":{a["total"]},"planned":{a["planned"]},"unplanned":{a["unplanned"]},"events":{evts}}}'

    hop_js  = json.dumps(hop_stops,         default=str, ensure_ascii=False)
    ub_js   = json.dumps([s for s in hop_stops if s['area']=='Underbody'],
                         default=str, ensure_ascii=False)
    up_js   = json.dumps([s for s in hop_stops if s['area']=='Upperbody'],
                         default=str, ensure_ascii=False)
    kpis_js = json.dumps(kpis,  default=str, ensure_ascii=False)
    saf_js  = json.dumps(ppt_tables['safety'],       ensure_ascii=False)
    pq_js   = json.dumps(ppt_tables['part_quality'], ensure_ascii=False)
    bof_js  = json.dumps(ppt_tables['bodies_oof'],   ensure_ascii=False)

    # Merge ppt_items into ppt_tables for FTT/DPV
    ppt_full = {**ppt_tables, **ppt_items}
    ppt_full_js = json.dumps(ppt_full, default=str, ensure_ascii=False)

    # Also build a plain dict (used for MM_HISTORY)
    downtime_dict = {
        'total_min': total, 'planned_min': planned,
        'unplanned_min': unplanned, 'availability': avail,
        'areas': {k: dict(areas[k]) for k in areas},
        'hop_stops': hop_stops,
        'ub_stops': [s for s in hop_stops if s['area']=='Underbody'],
        'up_stops': [s for s in hop_stops if s['area']=='Upperbody'],
    }
    day_dict = {
        'date': str(report_date), 'wk': wk_str, 'day': day_int,
        'kpis': kpis, 'downtime': downtime_dict, 'ppt': ppt_full
    }

    js_str = (
        f"const MM_DATA = {{\n"
        f"  date: {json.dumps(str(report_date))},\n"
        f"  wk: {json.dumps(wk_str)},\n"
        f"  day: {day_int},\n"
        f"  kpis: {kpis_js},\n"
        f"  downtime: {{\n"
        f"    total_min: {total}, planned_min: {planned},\n"
        f"    unplanned_min: {unplanned}, availability: {avail},\n"
        f"    areas: {{\n"
        f"      underbody: {area_js('underbody')},\n"
        f"      upperbody: {area_js('upperbody')},\n"
        f"      hangon:    {area_js('hangon')},\n"
        f"      bodysides: {area_js('bodysides')}\n"
        f"    }},\n"
        f"    hop_stops: {hop_js},\n"
        f"    ub_stops:  {ub_js},\n"
        f"    up_stops:  {up_js}\n"
        f"  }},\n"
        f"  ppt: {ppt_full_js}\n"
        f"}};"
    )
    return js_str, day_dict


def patch_html(new_mm_js, day_dict, report_date, target_file=None):
    HISTORY_MAX = 10  # keep last 10 working days

    target = target_file or DASH
    with open(target, 'r', encoding='utf-8') as f: html = f.read()

    # ── patch MM_DATA ─────────────────────────────────────────────────────────
    new_html, n = re.subn(r'const MM_DATA = \{.*?\};', new_mm_js, html, flags=re.DOTALL)
    if n != 1:
        log(f"ERROR: Expected 1 MM_DATA block, got {n}")
        return False

    # ── patch MM_HISTORY ──────────────────────────────────────────────────────
    hist_match = re.search(r'const MM_HISTORY = \{.*?\};', new_html, flags=re.DOTALL)
    if hist_match:
        try:
            existing_js = hist_match.group(0)
            # Extract JSON body from the JS object literal
            body = re.sub(r'^const MM_HISTORY = ', '', existing_js).rstrip(';')
            history = json.loads(body)
        except Exception:
            history = {}
    else:
        history = {}

    # Add/update today's entry and trim to last HISTORY_MAX days
    history[str(report_date)] = day_dict
    sorted_dates = sorted(history.keys())
    if len(sorted_dates) > HISTORY_MAX:
        for old in sorted_dates[:-HISTORY_MAX]:
            del history[old]

    new_history_js = 'const MM_HISTORY = ' + json.dumps(history, default=str, ensure_ascii=False) + ';'

    if hist_match:
        new_html = new_html[:hist_match.start()] + new_history_js + new_html[hist_match.end():]
    else:
        # Fallback: append before AREA_CFG
        new_html = new_html.replace(
            '/* ─── Area display config ─── */',
            new_history_js + '\n\n/* ─── Area display config ─── */'
        )

    with open(target, 'w', encoding='utf-8') as f: f.write(new_html)
    log(f"HTML patched: {target}  (history: {len(history)} days)")
    return True


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def update():
    today       = datetime.date.today()
    report_date = prev_working_day(today)
    wk_str, day_int = date_to_wk_day(report_date)

    log(f"=== Morning Meeting update started ({today}) ===")
    log(f"Report date: {report_date}  ({wk_str} D{day_int})")
    if wk_str is None: log("ERROR: weekend"); return False

    # ── 1. Power BI query (FTT/DPV — pre-calculated, no math needed) ──────────
    pbi = query_powerbi(report_date)
    if pbi:
        kpis      = build_kpis_from_pbi(pbi)
        ppt_items = build_ppt_items_from_pbi(pbi)
        log(f"  Power BI data loaded successfully")
        for k, v in kpis.items():
            if isinstance(v, dict) and v.get('val') is not None:
                log(f"    {k}: {v['val']}")
    else:
        log("  WARNING: Power BI not available — using cached/sample KPI values")
        log("  TIP: Open Power BI Desktop, refresh the FTT-DPV file, then re-run this script")
        kpis      = {k: {'val': None} for k in TARGETS}
        ppt_items = {}

    # ── 2. Downtime from OneDrive Excel ───────────────────────────────────────
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

    areas     = read_area_dt(hop_ws, dt_ws, wk_str, day_int)
    hop_stops = build_hop_stops(hop_ws, dt_ws, wk_str, day_int)

    for k, a in areas.items():
        log(f"  {k:12s}: total={a['total']} min  unplanned={a['unplanned']}")
    log(f"  Top HOP stop: {hop_stops[0]['line'] if hop_stops else 'none'}")

    # ── 3. PPT tables (Safety, Part Quality, Bodies OOF, Scrap) ──────────────
    ppt_path = find_ppt(wk_str, day_int)
    if ppt_path:
        log(f"Found PPT: {ppt_path}")
        md_text    = read_ppt_markdown(ppt_path)
        ppt_tables = parse_ppt_tables(md_text)
    else:
        ppt_tables = {'safety':{'title':'','detail':'','meta':''},
                      'part_quality':[], 'bodies_oof':[], 'scrap':'$0',
                      'scrap_note':'PPT not found — update manually'}

    # ── 4. Build + patch ──────────────────────────────────────────────────────
    mm_js, day_dict = build_mm_data(areas, hop_stops, kpis, ppt_items, ppt_tables,
                                    report_date, wk_str, day_int)
    success = patch_html(mm_js, day_dict, report_date)
    patch_html(mm_js, day_dict, report_date, target_file=DASH_MOBILE)
    log("Mobile dashboard also patched.")

    # ── 5. Push both dashboards to GitHub Pages ───────────────────────────────
    if GITHUB_ENABLED:
        try:
            now_str = datetime.datetime.now().strftime('%H:%M')
            for lock_name in ['index.lock', 'HEAD.lock', 'config.lock', 'packed-refs.lock']:
                lock_path = os.path.join(GITHUB_REPO, '.git', lock_name)
                if os.path.exists(lock_path):
                    try: os.remove(lock_path); log(f"Removed stale git lock: {lock_name}")
                    except Exception as le: log(f"Could not remove lock {lock_name}: {le}")

            def run_git(args, allow_fail=False):
                cmd = ['git', '-C', GITHUB_REPO] + args
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                out = (result.stdout + result.stderr).strip()
                if result.returncode == 0: log(f"Git OK: {' '.join(args)}")
                else: log(f"Git {'WARN' if allow_fail else 'FAIL'}: {' '.join(args)} -> {out[:200]}")
                return result.returncode == 0

            run_git(['fetch', 'origin', 'main'])
            run_git(['reset', '--mixed', 'origin/main'])
            run_git(['add', 'morning_meeting_dashboard.html', 'morning_meeting_mobile.html'])
            run_git(['commit', '-m', f'MM-update-{report_date}-{now_str}'], allow_fail=True)
            ok = run_git(['push', 'origin', 'main'])
            if not ok:
                log("Git push failed — retrying with force-with-lease")
                run_git(['fetch', 'origin', 'main'])
                run_git(['reset', '--mixed', 'origin/main'])
                run_git(['add', 'morning_meeting_dashboard.html', 'morning_meeting_mobile.html'])
                run_git(['commit', '-m', f'MM-update-{report_date}-retry'], allow_fail=True)
                run_git(['push', '--force-with-lease', 'origin', 'main'])
        except Exception as e:
            log(f"Git ERROR: {e}")
    else:
        log("GitHub push skipped (GITHUB_ENABLED=False).")

    for tmp in [tmp_hop, tmp_dt]:
        try: os.remove(tmp)
        except: pass

    log("=== Morning Meeting update complete ===\n")
    return success


def patch_history_only(day_dict, report_date, target_file=None):
    """Add a day_dict to MM_HISTORY without touching MM_DATA.
    Used for backfilling past days."""
    HISTORY_MAX = 10
    target = target_file or DASH
    with open(target, 'r', encoding='utf-8') as f: html = f.read()

    hist_match = re.search(r'const MM_HISTORY = \{.*?\};', html, flags=re.DOTALL)
    if hist_match:
        try:
            body = re.sub(r'^const MM_HISTORY = ', '', hist_match.group(0)).rstrip(';')
            history = json.loads(body)
        except Exception:
            history = {}
    else:
        history = {}

    history[str(report_date)] = day_dict
    sorted_dates = sorted(history.keys())
    if len(sorted_dates) > HISTORY_MAX:
        for old in sorted_dates[:-HISTORY_MAX]:
            del history[old]

    new_history_js = 'const MM_HISTORY = ' + json.dumps(history, default=str, ensure_ascii=False) + ';'

    if hist_match:
        new_html = html[:hist_match.start()] + new_history_js + html[hist_match.end():]
    else:
        new_html = html.replace(
            '/* ─── Area display config ─── */',
            new_history_js + '\n\n/* ─── Area display config ─── */'
        )

    with open(target, 'w', encoding='utf-8') as f: f.write(new_html)
    log(f"History-only patch: {target}  (now {len(history)} days)")
    return True


def backfill(target_date):
    """Backfill MM_HISTORY for a specific past date using Power BI + OneDrive.
    Does NOT change MM_DATA. Pushes to GitHub when all dates are done."""
    wk_str, day_int = date_to_wk_day(target_date)
    if wk_str is None:
        log(f"Skipping {target_date} — not a working day")
        return False

    log(f"=== Backfill history: {target_date}  ({wk_str} D{day_int}) ===")

    # 1. Power BI (same as daily update)
    pbi = query_powerbi(target_date)
    if pbi:
        kpis      = build_kpis_from_pbi(pbi)
        ppt_items = build_ppt_items_from_pbi(pbi)
        log("  Power BI data loaded")
        for k, v in kpis.items():
            if isinstance(v, dict) and v.get('val') is not None:
                log(f"    {k}: {v['val']}")
    else:
        log("  WARNING: Power BI not available — KPIs will be null for this day")
        kpis      = {k: {'val': None} for k in TARGETS}
        ppt_items = {}

    # 2. Downtime (same as daily update)
    tmp_hop = os.path.join(WORK, '_mm_hop.xlsm')
    tmp_dt  = os.path.join(WORK, '_mm_dt.xlsm')
    for src, dst, name in [(HOP_SRC, tmp_hop, 'HOP'), (DT_SRC, tmp_dt, 'DT')]:
        if os.path.exists(src):
            shutil.copy2(src, dst)
            log(f"  Copied {name}")
        else:
            log(f"  WARNING: {name} not found: {src}")

    hop_ws = dt_ws = None
    try:
        wb = openpyxl.load_workbook(tmp_hop, data_only=True, read_only=True)
        hop_ws = wb['New(DT LOG)']
    except Exception as e: log(f"  HOP error: {e}")
    try:
        wb = openpyxl.load_workbook(tmp_dt, data_only=True, read_only=True)
        dt_ws = wb['New(DT LOG)']
    except Exception as e: log(f"  DT error: {e}")

    areas     = read_area_dt(hop_ws, dt_ws, wk_str, day_int)
    hop_stops = build_hop_stops(hop_ws, dt_ws, wk_str, day_int)

    # 3. PPT (same as daily update)
    ppt_path = find_ppt(wk_str, day_int)
    if ppt_path:
        log(f"  Found PPT: {ppt_path}")
        ppt_tables = parse_ppt_tables(read_ppt_markdown(ppt_path))
    else:
        ppt_tables = {'safety': {'title': '', 'detail': '', 'meta': ''},
                      'part_quality': [], 'bodies_oof': [],
                      'scrap': '$0', 'scrap_note': 'PPT not found for this day'}

    # 4. Build day_dict and patch ONLY history (MM_DATA unchanged)
    _, day_dict = build_mm_data(areas, hop_stops, kpis, ppt_items, ppt_tables,
                                target_date, wk_str, day_int)

    patch_history_only(day_dict, target_date, DASH)
    patch_history_only(day_dict, target_date, DASH_MOBILE)

    for tmp in [tmp_hop, tmp_dt]:
        try: os.remove(tmp)
        except: pass

    log(f"=== Backfill complete: {target_date} ===\n")
    return True


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='A-Shop Morning Meeting Dashboard updater')
    parser.add_argument(
        '--date', metavar='YYYY-MM-DD',
        help='Backfill MM_HISTORY for a specific past date without changing MM_DATA'
    )
    parser.add_argument(
        '--backfill-week', action='store_true',
        help='Backfill all working days of the current WK12 (Mon–Thu) into MM_HISTORY'
    )
    args = parser.parse_args()

    if args.date:
        target = datetime.date.fromisoformat(args.date)
        ok = backfill(target)
        # Push to GitHub after backfill
        if ok and GITHUB_ENABLED:
            try:
                subprocess.run(['git', '-C', GITHUB_REPO, 'add',
                                'morning_meeting_dashboard.html',
                                'morning_meeting_mobile.html'], check=True)
                subprocess.run(['git', '-C', GITHUB_REPO, 'commit', '-m',
                                f'backfill-history-{args.date}'], check=True)
                subprocess.run(['git', '-C', GITHUB_REPO, 'push', 'origin', 'main'], check=True)
                log("GitHub push OK")
            except Exception as e:
                log(f"Git push error: {e}")
        sys.exit(0 if ok else 1)

    elif args.backfill_week:
        today  = datetime.date.today()
        mon    = today - datetime.timedelta(days=today.weekday())   # Monday of current week
        days   = [mon + datetime.timedelta(days=i) for i in range(5)
                  if (mon + datetime.timedelta(days=i)) < today]    # Mon–yesterday
        log(f"Backfilling {len(days)} day(s): {[str(d) for d in days]}")
        for d in days:
            backfill(d)
        # Single GitHub push after all days
        if GITHUB_ENABLED:
            try:
                subprocess.run(['git', '-C', GITHUB_REPO, 'add',
                                'morning_meeting_dashboard.html',
                                'morning_meeting_mobile.html'], check=True)
                subprocess.run(['git', '-C', GITHUB_REPO, 'commit', '-m',
                                f'backfill-history-week-{mon}'], check=True)
                subprocess.run(['git', '-C', GITHUB_REPO, 'push', 'origin', 'main'], check=True)
                log("GitHub push OK")
            except Exception as e:
                log(f"Git push error: {e}")
        sys.exit(0)

    else:
        sys.exit(0 if update() else 1)
