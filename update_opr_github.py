#!/usr/bin/env python3
"""
update_opr_github.py — runs in GitHub Actions (cloud, no laptop needed)
Queries Snowflake for per-window BOL + Empty Skid data and updates the HTML.
Credentials come from GitHub Secrets (env vars), NOT stored in files.
Downtime data (HOP/DT files) is updated separately by update_dashboard.py
running on the laptop — this script only handles Snowflake OPR data.
"""

import os, re, datetime, json
import snowflake.connector

DASH = 'body_shop_intelligence.html'

# ── Snowflake connection from GitHub Secrets ─────────────────────────────────
# Uses private key if SNOWFLAKE_PRIVATE_KEY secret exists, else falls back to token
_pk_pem = os.environ.get('SNOWFLAKE_PRIVATE_KEY', '')
if _pk_pem:
    from cryptography.hazmat.primitives.serialization import load_pem_private_key, Encoding, PrivateFormat, NoEncryption
    from cryptography.hazmat.backends import default_backend
    _pk = load_pem_private_key(_pk_pem.encode(), password=None, backend=default_backend())
    _pk_der = _pk.private_bytes(Encoding.DER, PrivateFormat.PKCS8, NoEncryption())
    conn = snowflake.connector.connect(
        account     = os.environ['SNOWFLAKE_ACCOUNT'],
        user        = os.environ['SNOWFLAKE_USER'],
        private_key = _pk_der,
        role        = os.environ.get('SNOWFLAKE_ROLE', ''),
        warehouse   = os.environ.get('SNOWFLAKE_WAREHOUSE', 'REPORTING'),
        database    = 'VCCH',
        schema      = 'PRODUCTION_TRACKING',
    )
else:
    # Fallback: JWT token (works until June 2026)
    conn = snowflake.connector.connect(
        account   = os.environ['SNOWFLAKE_ACCOUNT'],
        user      = os.environ['SNOWFLAKE_USER'],
        password  = os.environ['SNOWFLAKE_TOKEN'],
        role      = os.environ.get('SNOWFLAKE_ROLE', ''),
        warehouse = os.environ.get('SNOWFLAKE_WAREHOUSE', 'REPORTING'),
        database  = 'VCCH',
        schema    = 'PRODUCTION_TRACKING',
    )

# ── Date helpers ──────────────────────────────────────────────────────────────
def last_7_working_days():
    today = datetime.date.today()
    days, d = [], today
    while len(days) < 7:
        if d.weekday() < 5:
            days.append(d)
        d -= datetime.timedelta(days=1)
    return days   # newest first

# Exact scan-time boundaries per production window (CET scan time = prod time + ~5h transit)
# W5 prod 10:40-11:40 → scan 15:40-16:40 (old hourly mapping missed this, causing wrong values)
WIN_SCAN_RANGES = [
    ('W1','11:00','12:00'), ('W2','12:00','13:00'),
    ('W3','13:10','14:10'), ('W4','14:10','15:00'),
    ('W5','15:40','16:40'), ('W6','16:40','17:30'),
    ('W7','17:40','18:40'), ('W8','18:40','19:40'),
]

def get_production(date_str):
    sql = f"""
        SELECT "registrationPoint",
               CASE
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '06:00:00' AND '06:59:59' THEN 1
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '07:00:00' AND '07:59:59' THEN 2
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '08:10:00' AND '09:09:59' THEN 3
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '09:10:00' AND '09:59:59' THEN 4
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '10:40:00' AND '11:39:59' THEN 5
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '11:40:00' AND '12:29:59' THEN 6
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '12:40:00' AND '13:39:59' THEN 7
                 WHEN TO_TIME(DATEADD('hour',-5,CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint"))) BETWEEN '13:40:00' AND '14:39:59' THEN 8
               END AS win,
               COUNT(*) AS cars
        FROM   VCCH.PRODUCTION_TRACKING.BODY_TRACKING
        WHERE  DATE(CONVERT_TIMEZONE('Europe/Brussels',"timestampRegistrationPoint")) = '{date_str}'
          AND  "registrationPoint" IN ('13000','19900')
        GROUP BY 1,2
        HAVING win IS NOT NULL
        ORDER BY 1,2
    """
    cursor = conn.cursor()
    cursor.execute(sql)
    rows = cursor.fetchall()

    bol = {r[1]: r[2] for r in rows if str(r[0]) == '13000' and r[1]}
    emp = {r[1]: r[2] for r in rows if str(r[0]) == '19900' and r[1]}

    bol_h = [bol.get(i+1, 0) for i in range(8)]
    emp_h = [emp.get(i+1, 0) for i in range(8)]

    bol_tot = sum(bol_h)
    emp_tot = sum(emp_h)
    print(f"  {date_str}: BOL={bol_tot} Empty={emp_tot}  windows={emp_h}")
    return bol_h, emp_h, bol_tot, emp_tot

# ── Patch HTML ────────────────────────────────────────────────────────────────
def patch_day(html, date_str, bol_h, emp_h, bol_tot, emp_tot):
    """Replace just the bol_h/empty_h/bol_tot/empty_tot for one day."""
    def arr(lst): return '[' + ','.join(str(x) for x in lst) + ']'

    # Match the existing bol_h line for this date and replace values
    pattern = (
        rf"('{re.escape(date_str)}':\{{[^}}]*?label:[^,]+,"
        r"[^}]*?)\n    bol_h:\[[^\]]+\], empty_h:\[[^\]]+\],\n"
        r"    bol_tot:\d+, empty_tot:\d+,"
    )
    replacement = (
        rf"\1\n    bol_h:{arr(bol_h)}, empty_h:{arr(emp_h)},\n"
        rf"    bol_tot:{bol_tot}, empty_tot:{emp_tot},"
    )
    new_html, n = re.subn(pattern, replacement, html, flags=re.DOTALL)
    if n == 1:
        print(f"  Patched {date_str}")
    else:
        print(f"  Could not patch {date_str} (pattern not found — may not exist yet)")
    return new_html

# ── Main ──────────────────────────────────────────────────────────────────────
print(f"=== GitHub Actions OPR update {datetime.datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')} ===")

with open(DASH, encoding='utf-8') as f:
    html = f.read()

# Overtime days: Snowflake standard-shift query misses OT cars → always keep known values
OVERTIME_DATES = {'2026-03-16'}  # add dates here when OT occurs

days = last_7_working_days()
for d in days:
    date_str = str(d)
    if date_str not in html:
        print(f"  {date_str} not in DAYS_DATA yet — skipping (laptop update will add it)")
        continue
    if date_str in OVERTIME_DATES:
        print(f"  {date_str} is overtime — keeping existing values (Snowflake misses OT cars)")
        continue
    try:
        bol_h, emp_h, bol_tot, emp_tot = get_production(date_str)
        if emp_tot > 0:
            html = patch_day(html, date_str, bol_h, emp_h, bol_tot, emp_tot)
        else:
            print(f"  {date_str}: no data yet (shift not started?)")
    except Exception as e:
        print(f"  {date_str} error: {e}")

conn.close()

with open(DASH, 'w', encoding='utf-8') as f:
    f.write(html)

print("=== Done ===")
