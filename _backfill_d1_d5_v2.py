"""Second full backfill D1-D5 with all WD6/Final1 fixes applied.

Changes from first run:
  1. WD6 - FTT keyword ('WD6 - FTT' with spaces) added to SLIDE_CONFIG
  2. WD6 FTT positional scan (between W&G DPV and B-DPV)
  3. Final1 FTT uses 5-slide window from B-DPV (handles D4 structure)
  4. OPR/FTT sanity checks (drops absurd OCR values like bok_opr=11585 → null)
  5. Includes D5 (2026-03-20) — skipped automatically if PPT not found
"""
import sys, subprocess, datetime, time, os

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPT   = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\update_morning_meeting.py"
PYTHON   = r"C:\Users\NYOUSIF\AppData\Local\Python\pythoncore-3.14-64\python.exe"
DATES    = ["2026-03-16", "2026-03-17", "2026-03-18", "2026-03-19", "2026-03-20"]
LOG_FILE = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\mm_update_log.txt"
OUT_LOG  = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\_backfill_d1_d5_v2_out.txt"

def ts():
    return datetime.datetime.now().strftime('%H:%M:%S')

def log_line_count():
    try:
        with open(LOG_FILE, encoding='utf-8', errors='replace') as f:
            return sum(1 for _ in f)
    except Exception:
        return 0

with open(OUT_LOG, 'w', encoding='utf-8', errors='replace') as lf:
    def wlog(msg):
        line = f"[{ts()}] {msg}"
        print(line, flush=True)
        lf.write(line + '\n')
        lf.flush()

    wlog("=== Backfill D1-D5 v2 (WD6 FTT + Final1 + sanity checks) started ===")
    wlog(f"    Dates: {DATES}")

    for date in DATES:
        wlog(f"\n--- Starting backfill for {date} ---")
        before_count = log_line_count()
        wlog(f"    mm_update_log.txt before: {before_count} lines")

        proc = subprocess.Popen(
            [PYTHON, SCRIPT, "--date", date],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        wlog(f"    Subprocess PID: {proc.pid}")

        start   = time.time()
        timeout = 1200  # 20 minutes per day
        while proc.poll() is None:
            elapsed = time.time() - start
            if elapsed > timeout:
                proc.kill()
                wlog(f"  TIMEOUT after {int(timeout)}s — killed PID {proc.pid}")
                break
            time.sleep(15)
            cur = log_line_count()
            wlog(f"    Elapsed: {int(elapsed)}s | log lines: {cur} (+{cur - before_count})")

        rc          = proc.returncode if proc.returncode is not None else -1
        after_count = log_line_count()
        wlog(f"--- Done: {date}  exit={rc}  new log lines: {after_count - before_count} ---")

    wlog("\n=== All D1-D5 backfills complete ===")

print("v2 backfill script finished.")
