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

# Scan hours in CET for each of the 8 production windows
# W1(06-07) → scan hr 11, W2(07-08) → 12, ..., W8(13:40-14:40) → 18
SCAN_HOURS = [11, 12, 13, 14, 15, 16, 17, 18]

def get_production(date_str):
    sql = f"""
        SELECT "registrationPoint",
               HOUR(CONVERT_TIMEZONE('Europe/Brussels',
                    "timestampRegistrationPoint")) AS cet_hr,
               COUNT(*) AS cars
        FROM   VCCH.PRODUCTION_TRACKING.BODY_TRACKING
        WHERE  DATE(CONVERT_TIMEZONE('Europe/Brussels',
                    "timestampRegistrationPoint")) = '{date_str}'
          AND  "registrationPoint" IN ('13000', '19900')
          AND  HOUR(CONVERT_TIMEZONE('Europe/Brussels',
                    "timestampRegistrationPoint")) BETWEEN 10 AND 20
        GROUP BY 1, 2
        ORDER BY 1, 2
    """
    cursor = conn.cursor()
    cursor.execute(sql)
    rows = cursor.fetchall()

    bol = {r[1]: r[2] for r in rows if str(r[0]) == '13000'}
    emp = {r[1]: r[2] for r in rows if str(r[0]) == '19900'}

    bol_h = [bol.get(h, 0) for h in SCAN_HOURS]
    emp_h = [emp.get(h, 0) for h in SCAN_HOURS]
    # include any hour-19 late cars in the last window
    bol_h[-1] += bol.get(19, 0)
    emp_h[-1] += emp.get(19, 0)

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

days = last_7_working_days()
for d in days:
    date_str = str(d)
    if date_str not in html:
        print(f"  {date_str} not in DAYS_DATA yet — skipping (laptop update will add it)")
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
