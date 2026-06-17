# -*- coding: utf-8 -*-
"""
fix_missing_pos_files.py
掃描所有 order folder，找出缺少 _收貨明細 或 _清關 的訂單，
自動返去 POS 網站補下載，直到 4 個檔案齊全。
"""
import os, sys, time, re
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)
from datetime import datetime
from playwright.sync_api import sync_playwright

ORDERS_DIR = r"C:\Users\user\Desktop\順丰E順递\data\orders"
POS_URL    = "https://online-store-99126206.web.app/"
POS_PASS   = "0000"
today      = datetime.now().strftime("%Y%m%d")

# ── 1. 掃描所有 folder，找缺失的 ─────────────────────────────────────────────
def scan_missing():
    missing = []
    for folder in sorted(os.listdir(ORDERS_DIR)):
        folder_path = os.path.join(ORDERS_DIR, folder)
        if not os.path.isdir(folder_path):
            continue
        files = os.listdir(folder_path)
        has_receipt = any("_收貨明細" in f for f in files)
        has_customs = any("_清關"    in f for f in files)
        if has_receipt and has_customs:
            continue  # 齊全，跳過

        # 從 folder 名稱解析 customer / date / pos_order_no
        # 格式：{customer}_{YYYYMMDD}_{ORD-xxxxxx}
        m = re.match(r'^(.+)_(\d{8})_(ORD-\w+)$', folder)
        if not m:
            print(f"  [跳過] 無法解析 folder 名：{folder}")
            continue
        customer, date_str, pos_order_no = m.group(1), m.group(2), m.group(3)

        need = []
        if not has_receipt: need.append("收貨明細")
        if not has_customs: need.append("清關")
        missing.append({
            "folder":        folder,
            "folder_path":   folder_path,
            "customer":      customer,
            "date_str":      date_str,
            "pos_order_no":  pos_order_no,
            "need":          need,
        })
        print(f"  [缺] {folder}  →  缺: {', '.join(need)}")
    return missing

# ── 2. 對單個訂單補下載 ────────────────────────────────────────────────────────
def fix_one(page, ctx, entry):
    folder_path  = entry["folder_path"]
    customer     = entry["customer"]
    pos_order_no = entry["pos_order_no"]
    date_str     = entry["date_str"]
    need         = entry["need"]
    file_base    = f"{customer}_{date_str}_{pos_order_no}"

    print(f"\n  補下載：{entry['folder']}  ({', '.join(need)})")

    try:
        page.goto(POS_URL, wait_until="domcontentloaded", timeout=25000)
        time.sleep(3)

        # 登入後台
        page.locator("button:has-text('后台管理')").first.click()
        time.sleep(0.8)
        page.locator("input[type='password']").first.fill(POS_PASS)
        page.keyboard.press("Enter")
        time.sleep(2)

        # 進「記錄」tab
        page.evaluate("""() => {
            for (const el of document.querySelectorAll('button')) {
                if (el.offsetParent === null) continue;
                const spans = [...el.querySelectorAll('span')];
                if (spans.some(s => s.textContent.trim() === '記錄')
                    || el.textContent.trim() === '記錄') {
                    el.click(); return;
                }
            }
        }""")
        time.sleep(2)

        # 先等 orders 從 Firebase 載入（至少有 1 個 a[download] 出現）
        try:
            page.wait_for_function(
                "() => document.querySelectorAll('a[download]').length > 0",
                timeout=15000)
            print("    訂單已載入")
        except Exception:
            print("    等待訂單載入超時，繼續嘗試")

        # 搜尋訂單（click + type 觸發 React onChange filter）
        search = page.locator("input[placeholder*='搜尋單號']").first
        search.wait_for(state="visible", timeout=8000)
        search.click()
        search.triple_click()     # 全選現有內容
        search.type(pos_order_no, delay=80)  # 逐字輸入確保 React filter 更新
        time.sleep(4)
        print(f"    搜尋 {pos_order_no} 完成")

        # ── 收貨明細（用 ORD 號精準定位，避免點錯）────────────────────────────
        if "收貨明細" in need:
            try:
                page.wait_for_function(f"""() => {{
                    for (const a of document.querySelectorAll('a[download]')) {{
                        const dl = a.getAttribute('download') || '';
                        if (dl.includes('{pos_order_no}') && dl.includes('明細')
                            && a.href && a.href.startsWith('blob:')) return true;
                    }}
                    return false;
                }}""", timeout=20000)
                print("    收貨明細 PDF 已就緒")
            except Exception:
                print("    等待收貨明細 PDF 超時，嘗試直接點擊")
            try:
                with page.expect_download(timeout=15000) as dl_info:
                    page.evaluate(f"""() => {{
                        for (const a of document.querySelectorAll('a[download]')) {{
                            if (a.offsetParent === null) continue;
                            const dl = a.getAttribute('download') || '';
                            if (dl.includes('{pos_order_no}') && dl.includes('明細')) {{
                                a.click(); return true;
                            }}
                        }}
                    }}""")
                dl  = dl_info.value
                ext = os.path.splitext(dl.suggested_filename)[1] or ".pdf"
                path = os.path.join(folder_path, f"{file_base}_收貨明細{ext}")
                dl.save_as(path)
                print(f"    收貨明細已儲存: {path}")
            except Exception as e:
                print(f"    收貨明細下載失敗: {e}")

        # ── 清關（用 ORD 號精準定位）────────────────────────────────────────
        if "清關" in need:
            try:
                page.wait_for_function(f"""() => {{
                    for (const a of document.querySelectorAll('a[download]')) {{
                        const dl = a.getAttribute('download') || '';
                        if (dl.includes('{pos_order_no}') && dl.includes('清關')
                            && a.href && a.href.startsWith('blob:')) return true;
                    }}
                    return false;
                }}""", timeout=20000)
                print("    清關 PDF 已就緒")
            except Exception:
                print("    等待清關 PDF 超時，嘗試直接點擊")
            try:
                with page.expect_download(timeout=15000) as dl_info:
                    page.evaluate(f"""() => {{
                        for (const a of document.querySelectorAll('a[download]')) {{
                            if (a.offsetParent === null) continue;
                            const dl = a.getAttribute('download') || '';
                            if (dl.includes('{pos_order_no}') && dl.includes('清關')) {{
                                a.click(); return true;
                            }}
                        }}
                    }}""")
                dl  = dl_info.value
                ext = os.path.splitext(dl.suggested_filename)[1] or ".pdf"
                path = os.path.join(folder_path, f"{file_base}_清關{ext}")
                dl.save_as(path)
                print(f"    ✅ 清關PDF已儲存: {path}")
            except Exception as e:
                print(f"    ❌ 清關PDF下載失敗: {e}")

    except Exception as e:
        print(f"    ❌ 補下載失敗: {e}")

# ── 3. 主流程 ────────────────────────────────────────────────────────────────
print("=" * 60)
print("  掃描所有訂單 folder，找出缺少 收貨明細 / 清關 的訂單")
print("=" * 60)
missing = scan_missing()

if not missing:
    print("\n  ✅ 所有訂單檔案齊全，無需補下載！")
else:
    print(f"\n  共 {len(missing)} 個訂單需要補下載，開啟瀏覽器...")
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        ctx     = browser.new_context()
        page    = ctx.new_page()
        ok = fail = 0
        for entry in missing:
            fix_one(page, ctx, entry)
            # 驗證結果
            files = os.listdir(entry["folder_path"])
            still_missing = []
            if "收貨明細" in entry["need"] and not any("_收貨明細" in f for f in files):
                still_missing.append("收貨明細")
            if "清關" in entry["need"] and not any("_清關" in f for f in files):
                still_missing.append("清關")
            if still_missing:
                print(f"  ❌ {entry['folder']} 仍缺: {', '.join(still_missing)}")
                fail += 1
            else:
                print(f"  ✅ {entry['folder']} 補齊完成")
                ok += 1
            time.sleep(1)
        browser.close()

    print("\n" + "=" * 60)
    print(f"  補下載完成：✅ {ok} 成功  ❌ {fail} 失敗")
    print("=" * 60)
