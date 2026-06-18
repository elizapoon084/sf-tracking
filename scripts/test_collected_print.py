# 獨立測試：Collected 待攬收頁打印運單 PDF
# 用法（PowerShell）：
#   python scripts/test_collected_print.py SF1234567890,SF0987654321
import sys, os, time, base64
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
from playwright.sync_api import sync_playwright

CHROME_PROFILE = r"C:\ChromeAutomation"
COLLECTED_URL  = ("https://camp.sf-express.com/Collected"
                  "?tabName=pending_collection&isCollect=false")
SAVE_DIR       = r"C:\Users\user\Desktop\順丰E順递\data\orders"

# 從命令列讀運單號，沒有的話用假號測試
waybill_str = sys.argv[1] if len(sys.argv) > 1 else "SF0000000000000"
print(f"測試運單號：{waybill_str}")

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE,
        channel="chrome",
        headless=False,
        args=["--start-maximized"],
        no_viewport=True,
    )
    page = ctx.new_page()

    # ── 1. 開 Collected 頁 ──────────────────────────────────────────
    print("1. 開啟 Collected 頁...")
    page.goto(COLLECTED_URL, wait_until="domcontentloaded")
    time.sleep(5)

    # 截圖（初始狀態）
    ss1 = os.path.join(SAVE_DIR, "_test_collected_01_init.png")
    page.screenshot(path=ss1)
    print(f"   截圖已儲存：{ss1}")

    # ── 2. 找 textarea 並填入運單號 ─────────────────────────────────
    print("2. 找 textarea...")
    ta = None
    for sel in [
        "textarea.waybill-textarea",
        "textarea.waybill-textarea-width",
        "textarea[placeholder*='手動輸入']",
        "textarea[placeholder*='單號']",
    ]:
        try:
            loc = page.locator(sel)
            if loc.first.is_visible(timeout=3000):
                ta = loc.first
                print(f"   ✅ 找到 textarea：{sel}")
                break
        except Exception:
            pass
    if ta is None:
        print("   ❌ 找不到 textarea，截圖後退出")
        page.screenshot(path=os.path.join(SAVE_DIR, "_test_collected_02_no_textarea.png"))
        ctx.close(); sys.exit(1)

    # 逐個運單號 fill + Enter → 轉成 chip，再按搜尋 icon
    waybills_list = [w.strip() for w in waybill_str.split(",") if w.strip()]
    for w in waybills_list:
        ta.click()
        ta.fill(w)
        time.sleep(0.5)
        ta.press("Enter")
        time.sleep(1)
    print(f"   已輸入 {len(waybills_list)} 個運單號 chip")
    time.sleep(2)

    # ── 3. 找搜尋按鈕並點擊 ─────────────────────────────────────────
    print("3. 找搜尋 icon/按鈕...")
    searched = False
    for sel, desc in [
        ("div.func-icon-container img",          "func-icon-container img"),
        (".func-icon-container img:first-child",  "func-icon-container img:first-child"),
        ("[class*='func-icon'] img",              "[class*=func-icon] img"),
        (".el-icon-search",                       "el-icon-search"),
        ("button.search-btn",                     "button.search-btn"),
    ]:
        try:
            loc = page.locator(sel)
            if loc.first.is_visible(timeout=2000):
                loc.first.click()
                searched = True
                print(f"   ✅ 點擊搜尋：{desc}")
                break
        except Exception:
            pass
    if not searched:
        print("   ⚠️  找不到搜尋按鈕，改用 Enter")
        ta.press("Enter")

    time.sleep(6)

    # 截圖（搜尋後）
    ss2 = os.path.join(SAVE_DIR, "_test_collected_02_after_search.png")
    page.screenshot(path=ss2)
    print(f"   截圖：{ss2}")

    # ── 4. 等表格行並報告行數 ────────────────────────────────────────
    print("4. 等表格...")
    try:
        page.locator(".el-table__body tr").first.wait_for(state="visible", timeout=15000)
        row_count = page.locator(".el-table__body tr").count()
        print(f"   ✅ 表格已出現，行數：{row_count}")
    except Exception as e:
        print(f"   ❌ 等表格超時：{e}")
        ctx.close(); sys.exit(1)
    time.sleep(2)

    # ── 5. 全選 header checkbox ──────────────────────────────────────
    print("5. 全選 checkbox...")
    cb_found = False
    # 先試 CSS selectors
    for sel, desc in [
        ("th.el-table-column--selection .el-checkbox__inner",          "th.el-table-column--selection inner"),
        (".el-table__header-wrapper th .el-checkbox__inner",           "header-wrapper th inner"),
        (".el-table__fixed-header-wrapper .el-checkbox__inner",        "fixed-header-wrapper inner"),
        (".el-table-column--selection .el-checkbox__inner",            "el-table-column--selection inner"),
        (".el-table__header .el-checkbox__inner",                      "el-table__header inner"),
        ("thead .el-checkbox__inner",                                  "thead inner"),
        (".el-table__header input[type=checkbox]",                     "header input checkbox"),
    ]:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=2000):
                loc.click()
                cb_found = True
                print(f"   ✅ 點擊 header checkbox（CSS）：{desc}")
                break
        except Exception:
            pass
    # CSS 全失敗 → 用 JS 直接點
    if not cb_found:
        try:
            clicked = page.evaluate("""
                () => {
                    const th = document.querySelector(
                        'th.el-table-column--selection, th.is-center.el-table-column--selection'
                    );
                    if (th) {
                        const inner = th.querySelector('.el-checkbox__inner, input[type=checkbox]');
                        if (inner) { inner.click(); return 'clicked:' + (inner.className || inner.type); }
                    }
                    // fallback: 點第一個 th 裡的 checkbox
                    const cb = document.querySelector('thead .el-checkbox__inner');
                    if (cb) { cb.click(); return 'thead-fallback'; }
                    return 'not-found';
                }
            """)
            print(f"   JS 全選結果：{clicked}")
            cb_found = (clicked != "not-found")
        except Exception as e:
            print(f"   JS 全選失敗：{e}")
    if not cb_found:
        print("   ❌ 全選徹底失敗")
    time.sleep(2)

    # 報告選中行數
    checked = page.locator(".el-table__body .el-checkbox.is-checked").count()
    print(f"   選中行數：{checked}")

    # 截圖（全選後）
    ss3 = os.path.join(SAVE_DIR, "_test_collected_03_after_select.png")
    page.screenshot(path=ss3)
    print(f"   截圖：{ss3}")

    # ── 6. 批量打印 ──────────────────────────────────────────────────
    print("6. 點批量打印...")
    btn_found = False
    for sel, desc in [
        ("span:has-text('批量打印')",                                  "span 批量打印"),
        ("div.tag.btn-option span:has-text('批量打印')",               "tag btn-option span"),
        ("button:has-text('批量打印')",                                "button 批量打印"),
        ("[class*='btn'] span:has-text('批量打印')",                   "[class*=btn] span"),
    ]:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=3000):
                loc.click()
                btn_found = True
                print(f"   ✅ 點擊批量打印：{desc}")
                break
        except Exception:
            pass
    if not btn_found:
        print("   ❌ 找不到批量打印按鈕")
        ctx.close(); sys.exit(1)
    time.sleep(4)

    # 截圖（打印設置 dialog）
    ss4 = os.path.join(SAVE_DIR, "_test_collected_04_print_dialog.png")
    page.screenshot(path=ss4)
    print(f"   截圖：{ss4}")

    # ── 7. 等 PrintContent dialog 並找列印面單 ───────────────────────
    print("7. 等打印設置 dialog...")
    dialog_found = False
    for sel, desc in [
        (".PrintContent",         ".PrintContent"),
        (".sf-modal.PrintContent","sf-modal.PrintContent"),
        (".sf-modal",             ".sf-modal"),
        ("[class*='PrintContent']","[class*=PrintContent]"),
    ]:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=12000)
            dialog_found = True
            print(f"   ✅ dialog 出現：{desc}")
            break
        except Exception:
            pass
    if not dialog_found:
        print("   ❌ dialog 未出現，檢查截圖")
        ctx.close(); sys.exit(1)
    time.sleep(2)

    # ── 8. 點列印面單 → 攔截 popup blob ─────────────────────────────
    print("8. 點列印面單...")
    try:
        with page.expect_popup(timeout=30000) as popup_info:
            for sel, desc in [
                (".print_btn_block span.print",  "print_btn_block span.print"),
                (".print_btn_block .print",      "print_btn_block .print"),
                (".print_btn_block button",      "print_btn_block button"),
                ("span.print",                   "span.print"),
            ]:
                try:
                    loc = page.locator(sel).first
                    if loc.is_visible(timeout=3000):
                        loc.click()
                        print(f"   ✅ 點擊列印面單：{desc}")
                        break
                except Exception:
                    pass
            time.sleep(3)

        popup = popup_info.value
        blob_url = popup.url
        print(f"   popup URL：{blob_url[:80]}")
        time.sleep(3)

        if blob_url.startswith("blob:"):
            safe = blob_url.replace("'", "")
            b64 = page.evaluate(f"""
                async () => {{
                    const r = await fetch('{safe}');
                    const ab = await r.arrayBuffer();
                    const u8 = new Uint8Array(ab);
                    const chunks = [];
                    for (let i = 0; i < u8.length; i += 8192) {{
                        chunks.push(String.fromCharCode(
                            ...u8.subarray(i, Math.min(i+8192, u8.length))
                        ));
                    }}
                    return btoa(chunks.join(''));
                }}
            """)
            pdf_bytes = base64.b64decode(b64)
            out = os.path.join(SAVE_DIR, "_test_waybill_output.pdf")
            with open(out, "wb") as f:
                f.write(pdf_bytes)
            print(f"   ✅ PDF 已儲存（{len(pdf_bytes)} bytes）：{out}")
        else:
            print(f"   ⚠️  非 blob URL：{blob_url}")
        try: popup.close()
        except Exception: pass

    except Exception as e:
        print(f"   ❌ 列印面單失敗：{e}")

    print("\n完成。請查看 data/orders/ 的截圖了解每步驟狀態。")
    input("按 Enter 關閉瀏覽器...")
    ctx.close()
