"""
3M Service Life Software - Automation Tool v2.2

Workflow:
  Step 0  Homepage -> select China region -> click >
  Step 1a Settings -> click "Estimate gas/vapor cartridge service life"
  Step 1b Disclaimer -> click "I Accept"
  Step 2  Contaminants -> search by CAS -> check -> fill exposure/unit/breakthrough -> click >
  Step 3  Cartridge -> select respirator type -> search model -> click card -> click >
  Step 4  Conditions -> set humidity/temp/work intensity/pressure -> click >
  Step 5  Results -> read service life -> click "Generate PDF" -> save
"""

import asyncio
import csv
import sys
import os
import re
import json
import subprocess
import urllib.request
from pathlib import Path
from datetime import datetime

# ─── Dependency check ────────────────────────────────────────────────────────

def pip_install(package):
    mirrors = [
        "https://pypi.tuna.tsinghua.edu.cn/simple",
        "https://mirrors.aliyun.com/pypi/simple",
        "https://pypi.org/simple",
    ]
    for mirror in mirrors:
        print(f"  Trying mirror: {mirror}")
        ret = subprocess.call(
            [sys.executable, "-m", "pip", "install", package, "-i", mirror,
             "--trusted-host", mirror.split("/")[2], "-q"],
            env={**os.environ, "no_proxy": "*", "NO_PROXY": "*"},
        )
        if ret == 0:
            return
    raise RuntimeError(f"Failed to install {package} from all mirrors.")

def check_and_install_deps():
    try:
        import playwright  # noqa
    except ImportError:
        print("  Installing playwright (using CN mirror)...")
        pip_install("playwright")

# ─── Color output ─────────────────────────────────────────────────────────────

class C:
    RESET  = "\033[0m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    CYAN   = "\033[96m"
    BOLD   = "\033[1m"
    DIM    = "\033[2m"

def log(msg, level="info"):
    icons  = {"info": "  i ", "ok": "  v ", "warn": "  ! ", "error": "  x ", "step": "  > "}
    colors = {"info": C.CYAN, "ok": C.GREEN, "warn": C.YELLOW, "error": C.RED, "step": C.BOLD}
    print(f"{colors.get(level,'')}{icons.get(level,'    ')}{msg}{C.RESET}")

# ─── Lookup tables ────────────────────────────────────────────────────────────

# humidity: user writes <65 or >=65, mapped to site dropdown label
HUMIDITY_MAP = {
    "<65":  "<65%", "low":  "<65%",
    ">=65": "≥65%", "≥65":  "≥65%", "high": "≥65%",
}

# temperature: fixed dropdown options on site (Celsius)
TEMP_OPTIONS = [10, 15, 20, 25, 30, 35, 40]

# work intensity: user writes light/moderate/heavy, site uses Chinese labels
INTENSITY_MAP = {
    "light":    "轻度", "轻度": "轻度", "轻": "轻度",
    "moderate": "中度", "中度": "中度", "中": "中度",
    "heavy":    "重度", "重度": "重度", "重": "重度",
}

REQUIRED_FIELDS = ["cas", "cartridge", "humidity", "temp_c", "work_intensity"]

# ─── Unit conversion ─────────────────────────────────────────────────────────

def _get_mol_weight(cas: str) -> float:
    url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{cas}/property/MolecularWeight/JSON"
    try:
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read())
        return float(data["PropertyTable"]["Properties"][0]["MolecularWeight"])
    except Exception:
        raise RuntimeError(f"PubChem 查不到 CAS {cas} 的分子量，请手动换算为 ppm 后填写")

def preprocess_tasks(tasks: list) -> list:
    """Resolve exposure source (exposure vs OEL) and convert mg/m3 to ppm."""
    out = []
    for i, t in enumerate(tasks, 1):
        t = dict(t)

        # Determine exposure source
        if not t.get("exposure") and t.get("oel"):
            t["exposure"] = t["oel"]
            t["unit"] = t.get("oel-unit", "mg/m3")
            t["_exposure_source"] = "OEL"
            log(f"  [{i}] 使用 OEL 作为暴露浓度: {t['exposure']} {t['unit']}", "info")
        else:
            t["_exposure_source"] = "exposure"

        unit = t.get("unit", "ppm").lower().replace(" ", "").replace("³", "3")
        if unit in ("mg/m3",):
            cas = t["cas"]
            log(f"  [{i}] 查询分子量 CAS {cas} (PubChem)...", "info")
            mw = _get_mol_weight(cas)
            exp_ppm = round(float(t["exposure"]) * 24.45 / mw, 6)
            log(f"  [{i}] {t['exposure']} mg/m³ → {exp_ppm} ppm (MW={mw})", "ok")
            t["_original_exposure"] = t["exposure"]
            t["_original_unit"] = "mg/m³"
            t["_mol_weight"] = str(mw)
            t["exposure"] = str(exp_ppm)
            if t.get("breakthrough"):
                bt_ppm = round(float(t["breakthrough"]) * 24.45 / mw, 6)
                t["breakthrough"] = str(bt_ppm)
        else:
            t["_original_exposure"] = ""
            t["_original_unit"] = ""
            t["_mol_weight"] = ""

        # Auto-set breakthrough to 10% of exposure (ppm) if not provided
        if not t.get("breakthrough"):
            bt_auto = round(float(t["exposure"]) * 0.1, 6)
            t["breakthrough"] = str(bt_auto)
            log(f"  [{i}] 突破浓度自动设为 {bt_auto} ppm (= 10% × 暴露浓度)", "info")

        out.append(t)
    return out

# ─── CSV load & validate ──────────────────────────────────────────────────────

def load_tasks(csv_path: str) -> list:
    for enc in ("utf-8-sig", "gbk", "gb18030", "utf-8"):
        try:
            with open(csv_path, newline="", encoding=enc) as f:
                rows = list(csv.DictReader(f))
            return [{k.strip().lower(): v.strip() for k, v in row.items()} for row in rows]
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f"Cannot decode {csv_path}. Save it as UTF-8 or GBK in Excel.")

def validate_tasks(tasks: list) -> list:
    errors = []
    for i, t in enumerate(tasks, 1):
        for f in REQUIRED_FIELDS:
            if not t.get(f):
                errors.append(f"Row {i}: missing required field '{f}'")
        if not t.get("exposure") and not t.get("oel"):
            errors.append(f"Row {i}: must have either 'exposure' or 'oel'")
        hum = t.get("humidity", "").lower().replace(" ", "")
        if hum not in HUMIDITY_MAP:
            errors.append(f"Row {i}: humidity '{t.get('humidity')}' invalid, use <65 or >=65")
        try:
            int(t.get("temp_c", "x"))
        except ValueError:
            errors.append(f"Row {i}: temp_c must be a number")
    return errors

# ─── Cookie / privacy banner ──────────────────────────────────────────────────

async def _dismiss_cookie_banner(page):
    """Dismiss 3M cookie banner - button is 'Reject Non-Essential Cookies'."""
    await asyncio.sleep(1.5)
    for name in ["Reject Non-Essential Cookies", "Accept All Cookies", "保存更改"]:
        try:
            btn = page.get_by_role("button", name=name)
            if await btn.count() > 0:
                await btn.click()
                log(f"  Cookie banner dismissed: '{name}'", "ok")
                await asyncio.sleep(0.8)
                return
        except Exception:
            continue
    # Fallback: class-based
    try:
        btn = page.locator("button.cmp-save-btn")
        if await btn.count() > 0:
            await btn.first.click()
            await asyncio.sleep(0.8)
    except Exception:
        pass

# ─── Popup handler ────────────────────────────────────────────────────────────

async def dismiss_popup(page):
    for _ in range(3):
        try:
            btn = page.locator("button:has-text('OK'), button:has-text('确定')")
            if await btn.count() > 0:
                await btn.first.click()
                await asyncio.sleep(0.4)
                return True
        except Exception:
            pass
    return False

# ─── Click next button ────────────────────────────────────────────────────────

async def click_next(page):
    for sel in [
        "button.btn_primary",
        "button.btn-next",
        "button[aria-label='Next']",
    ]:
        try:
            btn = page.locator(sel)
            if await btn.count() > 0:
                await btn.last.click()
                await page.wait_for_load_state("load")
                return
        except Exception:
            pass

    btns = page.locator("button")
    cnt = await btns.count()
    for i in range(cnt - 1, max(cnt - 5, -1), -1):
        btn = btns.nth(i)
        txt = (await btn.inner_text()).strip()
        if txt in (">", "›", "下一步", "Next", "继续"):
            await btn.click()
            await page.wait_for_load_state("load")
            return

# ─── Single task ──────────────────────────────────────────────────────────────

async def run_task(page, task: dict, output_dir: Path, idx: int, total: int) -> dict:
    cas           = task["cas"]
    exposure      = task["exposure"]
    cartridge     = task["cartridge"]          # e.g. 6001CN
    humidity_key  = task["humidity"].lower().replace(" ", "")
    temp_raw      = int(task["temp_c"])
    intensity_key = task.get("work_intensity", "light").lower()
    pressure      = task.get("pressure_atm", "1")
    breakthrough  = task.get("breakthrough", "")
    resp_type     = task.get("respirator_type", "可更换式")
    chem_name     = task.get("name", cas)

    humidity_label  = HUMIDITY_MAP.get(humidity_key, "<65%")
    temp_val        = min(TEMP_OPTIONS, key=lambda x: abs(x - temp_raw))
    intensity_label = INTENSITY_MAP.get(intensity_key, "轻度")

    result = {"name": chem_name, "cas": cas, "cartridge": cartridge,
              "status": "failed", "service_life": "",
              "exposure_source": task.get("_exposure_source", "exposure"),
              "exposure_ppm": exposure,
              "original_exposure": task.get("_original_exposure", ""),
              "original_unit": task.get("_original_unit", ""),
              "mol_weight": task.get("_mol_weight", ""),
              "pdf_path": "", "error": ""}
    op = "初始化"

    log(f"[{idx}/{total}] {chem_name} ({cas}) x {cartridge}", "step")

    try:
        # ── Step 0: navigate, select China, dismiss cookie, click next ─────────
        log("  Step 0 select region...", "info")
        op = "加载 3M SLS 首页"
        await page.goto("https://sls.3m.com/", wait_until="load", timeout=30000)
        await page.get_by_text("China (中国大陆) - 简体中文").click()
        await _dismiss_cookie_banner(page)   # cookie appears AFTER clicking China
        await asyncio.sleep(0.3)
        await page.locator("button").nth(2).click()
        await page.wait_for_load_state("load")

        # ── Step 1a: select function ──────────────────────────────────────────
        log("  Step 1a select function...", "info")
        await page.get_by_text("估算气体/蒸气滤毒盒的使用寿命").wait_for(state="visible", timeout=10000)
        await page.get_by_text("估算气体/蒸气滤毒盒的使用寿命").click()
        await page.wait_for_load_state("load")

        # ── Step 1b: accept disclaimer ────────────────────────────────────────
        log("  Step 1b accept disclaimer...", "info")
        await page.get_by_role("button", name="我接受").wait_for(state="visible", timeout=10000)
        await page.get_by_role("button", name="我接受").click()
        await page.wait_for_load_state("load")

        # ── Step 2: search chemical ───────────────────────────────────────────
        log(f"  Step 2 search CAS {cas}...", "info")
        op = "进入化学品搜索页"
        search = page.get_by_role("searchbox", name="通过名称或CAS#查找")
        await search.wait_for(state="visible", timeout=10000)
        await search.fill(cas)
        await page.locator(".inputSearch-label-icon > svg").click()
        await asyncio.sleep(3)
        await _dismiss_cookie_banner(page)  # banner sometimes reappears mid-session

        # verify correct CAS appears in results, then click its checkbox
        op = f"搜索化学品 CAS# {cas}"
        await page.locator(f"text=CAS# {cas}").wait_for(state="visible", timeout=30000)
        await page.locator(".inputBox-label-icon > svg").first.click()
        await asyncio.sleep(0.5)

        log(f"  Step 2 exposure={exposure} ppm, breakthrough={breakthrough or 'none'}...", "info")
        op = "填写暴露浓度"
        exp_box = page.get_by_role("textbox", name=re.compile(r'暴露')).first
        await exp_box.wait_for(state="visible", timeout=8000)
        await exp_box.fill(str(exposure))

        if breakthrough:
            bt = page.locator("input[name='breakthroughValue']")
            if await bt.count() > 0:
                await bt.fill(str(breakthrough))

        # next (empty-text button nth(2))
        await page.get_by_role("button").filter(has_text=re.compile(r"^$")).nth(2).click()
        await page.wait_for_load_state("load")

        # ── Step 3: respirator type + cartridge ──────────────────────────────
        log(f"  Step 3 select cartridge {cartridge}...", "info")
        await asyncio.sleep(1)

        # select respirator type, then click 完成
        op = f"选择呼吸器类型 '{resp_type}'"
        await page.get_by_text(resp_type).click()
        await page.get_by_role("button", name="完成").click()
        await asyncio.sleep(0.5)

        # search for cartridge model
        cart_search = page.get_by_role("searchbox", name="通过名称查找")
        await cart_search.wait_for(state="visible", timeout=8000)
        await cart_search.click()
        await cart_search.fill(cartridge)
        await page.locator(".inputSearch-label-icon > svg").click()
        await asyncio.sleep(3)

        # click the cartridge button (partial name match on model number)
        op = f"搜索滤盒型号 '{cartridge}'"
        cart_btn = page.get_by_role("button", name=re.compile(cartridge, re.IGNORECASE))
        await cart_btn.first.wait_for(state="visible", timeout=30000)
        await cart_btn.first.click()
        await asyncio.sleep(0.5)

        await page.get_by_role("button", name="选择并继续").click()
        await page.wait_for_load_state("load")

        # ── Step 4: conditions ────────────────────────────────────────────────
        log(f"  Step 4 temp={temp_val}C, intensity={intensity_label}, ATM={pressure}...", "info")

        # temperature: match by visible text containing the number, skip blank first option
        op = "加载环境设置页"
        temp_sel = page.get_by_label("温度")
        await temp_sel.wait_for(state="visible", timeout=10000)
        opts = await temp_sel.evaluate(
            "el => [...el.options].map(o => ({v: o.value, t: o.text.trim()}))"
        )
        log(f"  Temp options: {[o['t'] for o in opts]}", "info")
        op = f"设置温度 '{temp_val}°C'"
        t_opt = next((o for o in opts if re.search(r'\b' + str(temp_val) + r'\b', o["t"])), None)
        if t_opt:
            await temp_sel.select_option(value=t_opt["v"])
        else:
            available = [o["t"] for o in opts if o["t"]]
            raise Exception(f"No options were found matching temp '{temp_val}°C'. Available: {available}")

        # work intensity: match by visible label text
        int_sel = page.get_by_label("工作强度")
        opts = await int_sel.evaluate(
            "el => [...el.options].map(o => ({v: o.value, t: o.text.trim()}))"
        )
        log(f"  Intensity options: {[o['t'] for o in opts]}", "info")
        op = f"设置工作强度 '{task.get('work_intensity')}'"
        i_opt = next((o for o in opts if intensity_label in o["t"]), None)
        if i_opt:
            await int_sel.select_option(value=i_opt["v"])
        else:
            available = [o["t"] for o in opts if o["t"]]
            raise Exception(f"No options were found matching intensity '{task.get('work_intensity')}'. Available: {available}")

        # atmospheric pressure
        atm = page.get_by_role("spinbutton", name="大气压")
        await atm.wait_for(state="visible", timeout=10000)
        await atm.fill(str(pressure))

        # next (recorded as button nth(4))
        await page.get_by_role("button").nth(4).click()
        await page.wait_for_load_state("load")

        # ── Step 5: results & PDF ─────────────────────────────────────────────
        log("  Step 5 generate and download PDF...", "info")
        op = "加载查阅结果页"
        await page.wait_for_selector("text=使用寿命估算", timeout=20000)

        body = await page.inner_text("body")
        service_life = ""
        for pat in [r'使用寿命估算[：:\s]*(.+?(?:分钟|小时))', r'(\d+[\d.]*)\s*(?:分钟|小时)']:
            m = re.search(pat, body)
            if m:
                service_life = m.group(1).strip()
                break
        log(f"  Service life: {service_life or 'see PDF'}", "ok")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe = re.sub(r'[\\/:*?"<>|]', '_', chem_name)
        pdf_path = output_dir / f"{safe}_{cartridge}_{timestamp}.pdf"

        await page.get_by_role("button", name="生成PDF文档").click()
        await asyncio.sleep(2)

        async with page.expect_download(timeout=30000) as dl_info:
            await page.get_by_role("button", name="下载PDF文件").click()
        dl = await dl_info.value
        await dl.save_as(str(pdf_path))
        log(f"  Saved: {pdf_path.name}", "ok")

        try:
            await page.get_by_role("button", name="完成").first.click()
            await asyncio.sleep(0.5)
        except Exception:
            pass

        result.update({"status": "ok", "service_life": service_life, "pdf_path": str(pdf_path)})

    except Exception as e:
        err_str = str(e)
        is_timeout = "TimeoutError" in type(e).__name__ or "Timeout" in err_str

        if "No options were found" in err_str or "no option" in err_str.lower():
            msg = f"下拉选项不存在：{op}（请核对 CSV 中填写的值）"
        elif is_timeout and "CAS#" in op:
            msg = f"未找到化学品：CAS# {cas} 在网站中无搜索结果"
        elif is_timeout and "滤盒" in op:
            msg = f"未找到滤盒：型号 {cartridge!r} 在搜索结果中不存在"
        elif is_timeout:
            msg = f"网络超时：{op} 加载失败，请检查网络后重试"
        else:
            msg = f"{op} 失败：{err_str}"

        result["error"] = msg
        log(f"  Failed: {msg}", "error")
        try:
            shot = output_dir / f"ERROR_{re.sub(r'[\\/:*?<>|]','_',chem_name)}_{datetime.now().strftime('%H%M%S')}.png"
            await page.screenshot(path=str(shot), full_page=True)
            log(f"  Screenshot saved: {shot.name}", "warn")
        except Exception:
            pass

    return result

# ─── Batch runner ─────────────────────────────────────────────────────────────

async def run_all(tasks: list, output_dir: Path, headless: bool) -> list:
    from playwright.async_api import async_playwright

    results = []
    async with async_playwright() as p:
        browser = None
        used_channel = None
        for channel in ["chrome", "msedge"]:
            try:
                browser = await p.chromium.launch(
                    channel=channel,
                    headless=headless,
                    args=["--no-sandbox", "--disable-dev-shm-usage"],
                )
                used_channel = channel
                log(f"Using system browser: {channel}", "ok")
                break
            except Exception:
                continue
        if browser is None:
            log("Chrome and Edge not found. Please install Google Chrome or Microsoft Edge.", "error")
            log("Download Chrome: https://www.google.com/chrome/", "info")
            raise RuntimeError("No usable browser found.")
        ctx = await browser.new_context(
            locale="zh-CN",
            viewport={"width": 1280, "height": 900},
            accept_downloads=True,
        )
        # Pre-inject OneTrust consent cookies to suppress the cookie banner
        await ctx.add_cookies([
            {"name": "OptanonAlertBoxClosed",
             "value": "2024-01-01T00:00:00.000Z",
             "domain": ".3m.com", "path": "/"},
            {"name": "OptanonConsent",
             "value": "isGpcEnabled=0&interactionCount=1&isAnonUser=1&groups=C0001%3A1%2CC0002%3A1%2CC0003%3A1%2CC0004%3A1",
             "domain": ".3m.com", "path": "/"},
        ])
        page = await ctx.new_page()

        for i, task in enumerate(tasks, 1):
            res = await run_task(page, task, output_dir, i, len(tasks))
            if res["status"] != "ok":
                log(f"  Retrying [{i}/{len(tasks)}] {res['name']}...", "warn")
                await asyncio.sleep(3)
                res = await run_task(page, task, output_dir, i, len(tasks))
            results.append(res)
            if i < len(tasks):
                await asyncio.sleep(2)

        await browser.close()
    return results

def save_summary(results: list, output_dir: Path) -> Path:
    path = output_dir / f"summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["name", "cas", "cartridge", "status", "service_life",
                                           "exposure_source", "exposure_ppm",
                                           "original_exposure", "original_unit",
                                           "mol_weight", "pdf_path", "error"],
                           extrasaction="ignore")
        w.writeheader()
        w.writerows(results)
    return path

# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    if sys.platform == "win32":
        os.system("color")

    print(f"""
{C.BOLD}{C.CYAN}╔══════════════════════════════════════════╗
║   3M Cartridge Service Life Tool v2.2   ║
║   Playwright headless - runs in bg      ║
╚══════════════════════════════════════════╝{C.RESET}
""")

    log("Checking dependencies...", "info")
    check_and_install_deps()
    log("Dependencies ready", "ok")

    script_dir = Path(__file__).parent

    if len(sys.argv) > 1:
        csv_path = Path(sys.argv[1])
    else:
        csv_path = script_dir / "tasks.csv"
        if not csv_path.exists():
            log("tasks.csv not found. Drag the file here and press Enter.", "warn")
            dragged = input("  Path > ").strip().strip('"')
            if not dragged:
                sys.exit(0)
            csv_path = Path(dragged)

    if not csv_path.exists():
        log(f"File not found: {csv_path}", "error")
        input("Press Enter to exit...")
        sys.exit(1)

    tasks = load_tasks(str(csv_path))
    log(f"Loaded {len(tasks)} task(s)", "ok")

    errors = validate_tasks(tasks)
    fatal = [e for e in errors if "missing required" in e]
    if fatal:
        log("Errors found, please fix and re-run:", "error")
        for e in fatal:
            print(f"  {C.RED}- {e}{C.RESET}")
        input("Press Enter to exit...")
        sys.exit(1)
    for e in errors:
        log(e, "warn")

    log("Preprocessing tasks (unit conversion if needed)...", "info")
    try:
        tasks = preprocess_tasks(tasks)
    except RuntimeError as e:
        log(str(e), "error")
        input("Press Enter to exit...")
        sys.exit(1)
    log("Preprocessing done", "ok")

    output_dir = script_dir / "output"
    output_dir.mkdir(exist_ok=True)
    log(f"Output folder: {output_dir}", "info")

    print(f"\n{C.DIM}  Default: silent background mode (screen not occupied).{C.RESET}")
    print(f"{C.DIM}  Type v to enable visual mode (browser window visible, good for first-run debug).{C.RESET}")
    choice = input("  > Enter=silent, v=visual: ").strip().lower()
    headless = (choice != "v")
    log("Silent background mode" if headless else "Visual mode (browser window will appear)", "ok")
    print()

    start = datetime.now()
    results = asyncio.run(run_all(tasks, output_dir, headless))
    elapsed = (datetime.now() - start).seconds

    summary = save_summary(results, output_dir)
    ok  = sum(1 for r in results if r["status"] == "ok")
    bad = len(results) - ok

    print(f"""
{C.BOLD}{'─'*44}
  All done! Elapsed: {elapsed}s
{'─'*44}{C.RESET}
  {C.GREEN}v OK:     {ok}{C.RESET}   {C.RED}x Failed: {bad}{C.RESET}
  Folder:  {output_dir}
  Summary: {summary.name}
{'─'*44}
""")

    if bad:
        log("Failed tasks (check ERROR_*.png or run in visual mode):", "warn")
        for r in results:
            if r["status"] == "failed":
                print(f"  {C.RED}- {r['name']} / {r['cartridge']}: {r['error']}{C.RESET}")

    input("\n  Press Enter to exit...")

if __name__ == "__main__":
    main()
