#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_webdata.py — cron อัปเดตข้อมูล Dashboard (จ-ศ 12:00)
Flow: ตรวจต้นทาง -> regenerate -> reconcile -> commit -> push -> ตรวจ production
กฎ:
- แก้ data_channels.json + shipto_data.json + sku_weekly_data.json
- ไม่แตะไฟล์ค้าง/ไฟล์อื่น
- ถ้าข้อมูลไม่เปลี่ยน -> ไม่ commit (exit 0 เงียบ)
- ถ้าผิดพลาด -> exit 1 (cron ส่ง error alert)
"""
import json, os, subprocess, sys, tempfile, shutil, datetime

WORKDIR = os.path.expanduser(
    r"~/OneDrive - TIA NGEE HIANG (CHAOSUA) CO.,LTD/Desktop/Chaosua_Ice/Web/LotusMakro-Dashboard")
SRC_XLSX = os.path.expanduser(
    r"~/OneDrive - TIA NGEE HIANG (CHAOSUA) CO.,LTD/Desktop/Chaosua_Ice/Data/Sales report/Sale Lotus Makro_Dashboard_Pivot.xlsx")
OUT_FILES = ["data_channels.json", "shipto_data.json", "sku_weekly_data.json"]
STATE_FILE = os.path.join(WORKDIR, ".webdata_state.json")
REMOTE = "origin"
BRANCH = "master"
PROD_URL = "https://moodlamstudio-cpu.github.io/chaosua-mt-dashboard/"

def sh(cmd, cwd=WORKDIR, timeout=600):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", timeout=timeout)
    return r.returncode, (r.stdout or "") + (r.stderr or "")

def main():
    dry = os.environ.get("DRY_RUN") == "1"
    print("=== เริ่มอัปเดต Dashboard" + (" (DRY RUN)" if dry else "") + " ===", flush=True)
    if not os.path.exists(WORKDIR):
        print("FAIL: ไม่พบโปรเจกต์", WORKDIR); sys.exit(1)
    if not os.path.exists(SRC_XLSX):
        print("FAIL: ไม่พบไฟล์ต้นทาง", SRC_XLSX); sys.exit(1)

    src_mtime = os.path.getmtime(SRC_XLSX)
    state = {}
    if os.path.exists(STATE_FILE):
        try: state = json.load(open(STATE_FILE, encoding="utf-8"))
        except Exception: state = {}
    if state.get("src_mtime") == src_mtime:
        print("ไม่มีข้อมูลใหม่จากต้นทาง (mtime เท่าเดิม) — ข้าม"); sys.exit(0)
    # ถ้าเพิ่ง commit ไปแล้ว (state ล่าสุด = commit ล่าสุด) และไฟล์ไม่เปลี่ยน -> ข้าม

    # 1. git status ก่อน (รายงานไฟล์ค้าง แต่ไม่แตะ)
    rc, out = sh(["git", "status", "--porcelain"])
    dirty = [l for l in out.splitlines() if l.strip() and not l.startswith("??")]
    untracked = [l[3:] for l in out.splitlines() if l.startswith("??")]
    if dirty:
        print("NOTE ไฟล์ค้างใน git (ไม่แตะ):", "; ".join(l[:60] for l in dirty[:10]))

    # 2. backup ไฟล์ที่จะ regenerate
    tmp = tempfile.mkdtemp(prefix="webdata_")
    for f in OUT_FILES:
        p = os.path.join(WORKDIR, f)
        if os.path.exists(p):
            shutil.copy2(p, os.path.join(tmp, f))

    # 3. รัน regenerate
    rc, out = sh([sys.executable, "regenerate_webdata.py"])
    if rc != 0:
        for f in OUT_FILES:  # rollback
            bp = os.path.join(tmp, f)
            if os.path.exists(bp):
                shutil.copy2(bp, os.path.join(WORKDIR, f))
        print("FAIL: regenerate_webdata.py error rc=", rc)
        print(out[-1500:]); sys.exit(1)
    print(out.strip()[-800:])

    # 4. reconcile: เปรียบเทียบกับ committed version
    rc, committed = sh(["git", "show", f"HEAD:{OUT_FILES[0]}"])
    if rc == 0:
        try:
            new = json.load(open(os.path.join(WORKDIR, OUT_FILES[0]), encoding="utf-8"))
            old = json.loads(committed)
            same = json.dumps(new, sort_keys=True) == json.dumps(old, sort_keys=True)
        except Exception:
            same = False
        if same:
            print("ข้อมูลเหมือน HEAD — ไม่ต้อง commit"); sys.exit(0)

    new = json.load(open(os.path.join(WORKDIR, OUT_FILES[0]), encoding="utf-8"))
    last_date = new.get("_lastSalesDate", "?")
    mt = new.get("MT", {})
    print(f"ข้อมูลใหม่: last_sales_date={last_date} MT total_mb={mt.get('total_mb')}")

    if dry:
        print("DRY RUN — หยุดก่อน commit (ไม่มีการแก้ git)"); sys.exit(0)

    # 5. commit เฉพาะ 2 ไฟล์
    rc, out = sh(["git", "add"] + OUT_FILES)
    if rc != 0: print("FAIL git add:", out[-500:]); sys.exit(1)
    msg = f"Refresh sales data through {last_date}"
    rc, out = sh(["git", "commit", "-m", msg])
    if rc != 0:
        print("NOTE git commit:", out[-500:])
        print("FAIL: commit ไม่สำเร็จ (อาจไม่มีอะไรเปลี่ยน)"); sys.exit(0 if "nothing to commit" in out.lower() else 1)

    # 6. pull --rebase (กัน conflict) แล้ว push
    rc, out = sh(["git", "pull", "--rebase", REMOTE, BRANCH], timeout=300)
    if rc != 0:
        print("WARN pull --rebase:", out[-500:])
    rc, out = sh(["git", "push", REMOTE, BRANCH], timeout=300)
    if rc != 0:
        print("FAIL push:", out[-500:]); sys.exit(1)

    # 7. commit hash + ตรวจ production
    rc, out = sh(["git", "rev-parse", "HEAD"])
    commit_hash = out.strip()[:7]
    try:
        import urllib.request
        r = urllib.request.urlopen(PROD_URL, timeout=60)
        prod_status = f"HTTP {r.status}"
    except Exception as e:
        prod_status = f"ERROR {e}"

    # 8. state file
    json.dump({"src_mtime": src_mtime, "last_date": last_date, "commit": commit_hash,
               "ran_at": datetime.datetime.now().isoformat()},
              open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

    print("")
    print("อัปเดตข้อมูลเว็บเรียบร้อยครับ")
    print(f"- ข้อมูลล่าสุดถึง: {last_date}")
    print(f"- ไฟล์ที่อัปเดต: {', '.join(OUT_FILES)}")
    print(f"- Commit: {commit_hash} — {msg}")
    print(f"- Production: {prod_status}")
    print("- อัปเดตรวม: ตาราง SKU รายสัปดาห์")

if __name__ == "__main__":
    main()
