"""Full audit of MM_HISTORY in dashboard HTML — KPIs + item counts for D1-D5."""
import json, re, os, sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

HTML = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\morning_meeting_dashboard.html"
DATES = ["2026-03-16","2026-03-17","2026-03-18","2026-03-19","2026-03-20"]
DAY_NAMES = {"2026-03-16":"D1","2026-03-17":"D2","2026-03-18":"D3",
             "2026-03-19":"D4","2026-03-20":"D5"}

KPI_KEYS = ['bok_opr','bol_opr','scrap_car','ashop_ftt','ashop_dpv',
            'wg_dpv','wd6_ftt','wd6_dpv','cal_ftt','final1_ftt','final2_ftt']

ITEM_KEYS = ['ashop_ftt_536','ashop_ftt_519','ashop_dpv_536','ashop_dpv_519',
             'wg_dpv_536','wg_dpv_519','wd6_ftt_536','wd6_ftt_519',
             'wd6_dpv_536','wd6_dpv_519','cal_ftt_536','cal_ftt_519',
             'final1_ftt_536','final1_ftt_519','final2_ftt_536','final2_ftt_519']

with open(HTML, encoding='utf-8', errors='replace') as f:
    content = f.read()

hist_match = re.search(r'const MM_HISTORY = \{.*?\};', content, flags=re.DOTALL)
body = re.sub(r'^const MM_HISTORY = ', '', hist_match.group(0)).rstrip(';')
history = json.loads(body)

print("=" * 70)
print("  WK12 DASHBOARD AUDIT — KPIs and Item Details")
print("=" * 70)

for date in DATES:
    day = history.get(date, {})
    label = DAY_NAMES[date]
    kpis = day.get('kpis', {})
    ppt  = day.get('ppt', {})
    print(f"\n{'─'*70}")
    print(f"  {label} ({date})")
    print(f"{'─'*70}")

    # KPIs
    print("  KPIs:")
    for k in KPI_KEYS:
        v = kpis.get(k, {})
        val = v.get('val') if isinstance(v, dict) else v
        status = "✓" if val is not None else "✗ NULL"
        print(f"    {k:20s} = {str(val):10s}  {status}")

    # Item counts
    print("  PPT Items:")
    for k in ITEM_KEYS:
        items = ppt.get(k) or []
        n = len(items)
        flag = "" if n == 0 else f"  [{n} items]"
        if n > 0:
            sample = items[:3]
            print(f"    {k:25s}: {n:3d} items  {sample}")
        else:
            print(f"    {k:25s}:   0 items")

print(f"\n{'='*70}")
print("  AUDIT COMPLETE")
print(f"{'='*70}")
