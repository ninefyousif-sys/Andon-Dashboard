"""Run backfill for D2-D4 sequentially.
Uses creationflags=CREATE_NEW_PROCESS_GROUP to prevent git credential helpers
from inheriting our stdout pipe and causing deadlock.
Progress is visible in mm_update_log.txt.
"""
import sys, subprocess, datetime, time, os

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SCRIPT  = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\update_morning_meeting.py"
PYTHON  = r"C:\Users\NYOUSIF\AppData\Local\Python\pythoncore-3.14-64\python.exe"
DATES   = ["2026-03-17", "2026-03-18", "2026-03-19"]
LOG_FILE = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\mm_update_log.txt"
OUT_LOG  = r"C:\Users\NYOUSIF\Desktop\AShop_Dashboard\_backfill_d2_d4_out.txt"

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

    wlog("=== Backfill D2-D4 started ===")
    for date in DATES:
        wlog(f"\n--- Starting backfill for {date} ---")
        before_count = log_line_count()
        wlog(f"    mm_update_log.txt before: {before_count} lines")

        # Use subprocess.Popen + wait() instead of run() to avoid pipe-lock.
        # Don't capture stdout — let it flow to the console / be discarded.
        proc = subprocess.Popen(
            [PYTHON, SCRIPT, "--date", date],
            stdout=subprocess.DEVNULL,  # discard stdout (avoids git-helper pipe deadlock)
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,  # no handle inheritance
        )
        wlog(f"    Subprocess PID: {proc.pid}")

        # Poll every 15 seconds until process exits or 15 min timeout
        start = time.time()
        timeout = 900  # 15 minutes
        while proc.poll() is None:
            elapsed = time.time() - start
            if elapsed > timeout:
                proc.kill()
                wlog(f"  TIMEOUT after {timeout}s — killed PID {proc.pid}")
                break
            time.sleep(15)
            cur_count = log_line_count()
            wlog(f"    Elapsed: {int(elapsed)}s | mm_update_log: {cur_count} lines (+{cur_count - before_count})")

        rc = proc.returncode if proc.returncode is not None else -1
        after_count = log_line_count()
        wlog(f"--- Done: {date}  exit={rc}  new log lines: {after_count - before_count} ---")

    wlog("\n=== All D2-D4 backfills complete ===")

print("Done.")
