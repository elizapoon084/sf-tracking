# 上傳 SF Excel + 打印運單 PDF（不需 POS）
import sys, os, re, json, time, base64
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
from playwright.sync_api import sync_playwright
from collections import defaultdict

CHROME_PROFILE  = r"C:\ChromeAutomation"
CAMP_URL        = "https://camp.sf-express.com/MonthCard"
CAMP_BATCH_URL  = "https://camp.sf-express.com/web/portal/jkd-hongkong/batchorder"
COLLECTED_URL   = "https://camp.sf-express.com/Collected?tabName=pending_collection&isCollect=false"
ORDERS_DIR      = r"C:\Users\user\Desktop\順丰E順递\data\orders"
MONTHLY_ACCOUNT = "8526937071"
TODAY           = "20260619"

EXCEL_PATH = r"C:\Users\user\Desktop\馬凱川_彭秋曲_20260619.xlsx"

# 客人 → 存檔目錄
CUSTOMERS = {
    "馬凱川": os.path.join(ORDERS_DIR, f"馬凱川_{TODAY}"),
    "彭秋曲": os.path.join(ORDERS_DIR, f"彭秋曲_{TODAY}"),
}
# Excel 行對應的客人（按上傳順序）
ORDER_CUSTOMERS = ["馬凱川", "彭秋曲"]

for d in CUSTOMERS.values():
    os.makedirs(d, exist_ok=True)


def _dismiss(pg):
    for _ in range(5):
        try:
            if pg.locator(".el-message-box__wrapper").is_visible(timeout=1500):
                btn = pg.locator(".el-message-box__btns button.el-button--primary")
                if btn.is_visible(timeout=1000):
                    btn.click(); time.sleep(0.8); continue
        except Exception:
            pass
        break


def _fetch_blob(pg, blob_url):
    safe = blob_url.replace("'", "")
    b64 = pg.evaluate(f"""
        async () => {{
            const r = await fetch('{safe}');
            const ab = await r.arrayBuffer();
            const u8 = new Uint8Array(ab);
            const ch = [];
            for (let i = 0; i < u8.length; i += 8192)
                ch.push(String.fromCharCode(...u8.subarray(i, Math.min(i+8192, u8.length))));
            return btoa(ch.join(''));
        }}
    """)
    return base64.b64decode(b64)


with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        CHROME_PROFILE, channel="chrome", headless=False,
        args=["--start-maximized"], no_viewport=True,
    )
    page = ctx.new_page()

    # ══ 階段一：上傳 Excel ══════════════════════════════════════════════
    print("\n" + "="*55)
    print("  階段一：上傳 SF 批量 Excel")
    print("="*55)

    page.goto(CAMP_URL, wait_until="domcontentloaded", timeout=40000)
    time.sleep(4)
    _dismiss(page)

    # 選月結帳號
    confirmed = page.evaluate("""
        () => {
            const btns = Array.from(document.querySelectorAll('button'));
            const b = btns.find(b => b.offsetParent !== null && b.textContent.trim() === '確認');
            if (b) { b.click(); return true; } return false;
        }
    """)
    if confirmed:
        print("  ✅ 已選擇月結帳號")
        time.sleep(4)

    # 去批量頁
    page.goto(CAMP_BATCH_URL, wait_until="domcontentloaded", timeout=40000)
    time.sleep(8)
    _dismiss(page)

    has_input = page.evaluate("() => !!document.querySelector('input[type=\"file\"]')")
    if not has_input:
        print("  等頁面完全渲染...")
        time.sleep(10)

    # 上傳
    for sel in ["input[type=file]", ".el-upload input[type=file]", "input[accept]"]:
        try:
            page.locator(sel).first.set_input_files(EXCEL_PATH)
            print(f"  ✅ 已上傳 Excel（{sel}）")
            break
        except Exception:
            pass
    time.sleep(6)
    _dismiss(page)

    # 等校驗
    print("  等待校驗...")
    for _ in range(30):
        if page.locator("text=導入失敗").first.is_visible(timeout=1000):
            break
        time.sleep(1)
    time.sleep(2)

    body = page.inner_text("body")
    ok_n  = int(m.group(1)) if (m := re.search(r"校驗成功\s+(\d+)", body)) else 0
    fail_n = int(m.group(1)) if (m := re.search(r"導入失敗\s+(\d+)", body)) else 0
    print(f"  校驗：成功 {ok_n} 行，失敗 {fail_n} 行")
    if fail_n > 0:
        page.screenshot(path=os.path.join(ORDERS_DIR, f"_validate_fail_{TODAY}.png"))
        raise RuntimeError(f"校驗失敗 {fail_n} 行")

    # 同意 + 提交
    page.locator("label", has_text="我同意").click()
    time.sleep(1)
    print("  提交訂單...")
    with page.expect_response(
        lambda r: "createBatchOrder" in r.url and r.status == 200, timeout=30000
    ) as resp_info:
        page.locator("button:has-text('提交訂單')").click()

    data = resp_info.value.json()
    time.sleep(3)

    # 解析運單號
    waybill_map = {}  # customer_name → [waybill, ...]
    success_list = data.get("successRecordList", [])
    for idx, rec in enumerate(success_list):
        sf_no = ""
        pj_str = rec.get("printJson", "")
        if pj_str:
            try:
                pj = json.loads(pj_str)
                sf_no = pj.get("masterWaybillNo", "") or (
                    pj.get("subWaybillNoList", [{}])[0].get("waybillNo", "")
                )
            except Exception:
                pass
        if not sf_no:
            for v in rec.values():
                if re.match(r"SF\d{10,}", str(v or "")):
                    sf_no = str(v); break
        cust = ORDER_CUSTOMERS[idx] if idx < len(ORDER_CUSTOMERS) else f"客人{idx}"
        if sf_no:
            waybill_map.setdefault(cust, []).append(sf_no)
        print(f"  [{idx}] {cust} → {sf_no or '?'}")

    # 若 API 解析失敗，從頁面取
    if not waybill_map:
        all_sf = list(dict.fromkeys(re.findall(r"SF\d{13,}", page.inner_text("body"))))
        for i, sf in enumerate(all_sf):
            cust = ORDER_CUSTOMERS[i] if i < len(ORDER_CUSTOMERS) else f"客人{i}"
            waybill_map.setdefault(cust, []).append(sf)
        print(f"  頁面掃描到運單：{all_sf}")

    print(f"\n  運單號：{waybill_map}")

    # ══ 階段二：Collected 頁批量打印 ═══════════════════════════════════
    print("\n" + "="*55)
    print("  階段二：查件服務打印運單 PDF")
    print("="*55)

    for cust_name, waybills in waybill_map.items():
        print(f"\n  客人: {cust_name}  運單: {', '.join(waybills)}")
        dest = os.path.join(CUSTOMERS[cust_name], f"{cust_name}_{TODAY}_運單.pdf")
        pdf_saved = False

        for attempt in range(3):
            try:
                page.goto(COLLECTED_URL, wait_until="domcontentloaded")
                time.sleep(5)
                _dismiss(page)

                # 逐個運單號輸入 → chip
                ta = page.locator("textarea.waybill-textarea, textarea[placeholder*='手動輸入']").first
                ta.wait_for(state="visible", timeout=12000)
                for w in waybills:
                    ta.click(); ta.fill(w); time.sleep(0.5)
                    ta.press("Enter"); time.sleep(1)
                time.sleep(2)

                # 點搜尋 icon
                for sel in ["div.func-icon-container img", ".func-icon-container img"]:
                    try:
                        page.locator(sel).first.click(timeout=2000); break
                    except Exception:
                        pass
                time.sleep(6)

                # 等表格
                page.locator(".el-table__body tr").first.wait_for(state="visible", timeout=20000)
                time.sleep(2)

                # 全選（fixed-header-wrapper）
                hdr = page.locator(".el-table__fixed-header-wrapper .el-checkbox__inner").first
                hdr.wait_for(state="visible", timeout=8000)
                hdr.click(); time.sleep(1)

                # 批量打印
                page.locator("span:has-text('批量打印'), button:has-text('批量打印')").first.click()
                time.sleep(3)

                # 等 PrintContent dialog
                page.locator(".PrintContent").first.wait_for(state="visible", timeout=20000)
                time.sleep(2)

                # 列印面單 → blob popup
                with page.expect_popup(timeout=30000) as popup_info:
                    page.locator(".print_btn_block span.print, .print_btn_block .print").first.click()
                    time.sleep(2)

                popup = popup_info.value
                blob_url = popup.url
                time.sleep(3)

                if blob_url.startswith("blob:"):
                    pdf_bytes = _fetch_blob(page, blob_url)
                    try: popup.close()
                    except Exception: pass
                    if len(pdf_bytes) > 1000:
                        with open(dest, "wb") as f:
                            f.write(pdf_bytes)
                        print(f"  ✅ 已儲存：{dest}")
                        pdf_saved = True
                        break
                    else:
                        print(f"  ⚠️  PDF 太小，重試...")
                else:
                    print(f"  ⚠️  非 blob URL: {blob_url[:60]}")
                    try: popup.close()
                    except Exception: pass

            except Exception as e:
                print(f"  ⚠️  失敗（{attempt+1}/3）：{e}")
                time.sleep(8)

        if not pdf_saved:
            print(f"  ❌ {cust_name} 運單 PDF 未能下載")
        time.sleep(3)

    print("\n" + "="*55)
    print("  完成！PDF 已存至各客人資料夾")
    print("="*55)
    ctx.close()
