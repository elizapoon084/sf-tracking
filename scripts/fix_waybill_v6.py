# -*- coding: utf-8 -*-
"""
fix_waybill_v6.py  —  重新列印順丰運單（修正申報價值）
=========================================================
py5 跑完後直接跑此腳本。

流程：
  讀 last_session.json（py5 寫的）
  → 確認視窗
  → 開 Chrome → SF 運單列表頁
  → 逐一：在列表搜尋運單號 → 點入詳情 → 修改訂單
          → 直接保存 → 列印電子運單 → 覆蓋原 PDF
"""

import os
import sys
import json
import time
import base64
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright

# ─── 路徑設定（與 py5 完全一致）──────────────────────────────────────────────
CHROME_PROFILE  = r"C:\ChromeAutomation"
LOGS_DIR        = r"C:\Users\user\Desktop\順丰E順递\logs"
SESSION_FILE    = r"C:\Users\user\Desktop\順丰E順递\data\last_session.json"

SF_LIST_URL     = "https://hk.sf-express.com/hk/tc/waybill/list"
SF_DETAIL_BASE  = "https://hk.sf-express.com/hk/tc/waybill/appointment-detail"

today = date.today().strftime("%Y%m%d")


# ══════════════════════════════════════════════════════════════════════════════
# 讀 last_session.json
# ══════════════════════════════════════════════════════════════════════════════

def read_session() -> list:
    if not os.path.exists(SESSION_FILE):
        messagebox.showerror(
            "找不到 Session 檔",
            f"請先跑 py5，py5 跑完會自動產生：\n{SESSION_FILE}"
        )
        sys.exit(1)

    with open(SESSION_FILE, encoding="utf-8") as f:
        data = json.load(f)

    if not data:
        messagebox.showinfo("Session 是空的", "py5 上次沒有成功完成任何訂單。")
        sys.exit(0)

    print(f"  讀到 {len(data)} 張運單（來自上次 py5）")
    for d in data:
        print(f"    {d['customer']:10s}  {d['waybill']}")
    return data


# ══════════════════════════════════════════════════════════════════════════════
# 確認視窗
# ══════════════════════════════════════════════════════════════════════════════

def show_confirm_gui(session: list) -> bool:
    result = [False]

    root = tk.Tk()
    root.title("重新列印順丰運單 v6")
    root.geometry("680x360")

    tk.Label(
        root,
        text=f"py5 剛完成了 {len(session)} 張運單，現在重新列印覆蓋原 PDF",
        font=("", 12, "bold"), anchor="w",
    ).pack(fill="x", padx=12, pady=(12, 4))

    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True, padx=12, pady=4)

    cols = ("#", "客人名", "順丰運單號", "PDF 路徑")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
    tree.heading("#",          text="#")
    tree.heading("客人名",     text="客人名")
    tree.heading("順丰運單號", text="順丰運單號")
    tree.heading("PDF 路徑",   text="PDF 路徑")
    tree.column("#",          width=30,  anchor="center")
    tree.column("客人名",     width=90)
    tree.column("順丰運單號", width=160)
    tree.column("PDF 路徑",   width=340)
    tree.pack(fill="both", expand=True, side="left")

    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")

    for i, d in enumerate(session):
        tree.insert("", "end", values=(
            i + 1,
            d.get("customer", ""),
            d.get("waybill",  ""),
            d.get("pdf_path", ""),
        ))

    tk.Label(
        root,
        text="流程：列表頁找運單 → 修改訂單 → 直接保存 → 列印電子運單 → 覆蓋原 PDF",
        fg="#555", font=("", 9),
    ).pack(anchor="w", padx=14, pady=(2, 0))

    def on_confirm():
        result[0] = True
        root.destroy()

    def on_cancel():
        sys.exit(0)

    btn_row = tk.Frame(root)
    btn_row.pack(pady=(8, 14))
    tk.Button(btn_row, text="開始重新列印",
              command=on_confirm,
              bg="#27ae60", fg="white",
              padx=16, pady=6, font=("", 11, "bold")).pack(side="left", padx=8)
    tk.Button(btn_row, text="取消",
              command=on_cancel,
              bg="#e74c3c", fg="white",
              padx=16, pady=6, font=("", 11)).pack(side="left", padx=8)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()
    return result[0]


# ══════════════════════════════════════════════════════════════════════════════
# 工具
# ══════════════════════════════════════════════════════════════════════════════

def shot(page, label):
    os.makedirs(LOGS_DIR, exist_ok=True)
    p = os.path.join(LOGS_DIR, f"py6_{label}.png")
    try:
        page.screenshot(path=p, full_page=False)
    except Exception:
        pass


def dismiss_popups(page):
    """處理 Chrome「還原網頁」彈窗及其他系統彈窗。"""
    try:
        page.evaluate("""() => {
            const dismiss = ['不用了','關閉','No thanks','Dismiss','取消','否'];
            for (const el of document.querySelectorAll('button')) {
                const t = (el.textContent || '').trim();
                if (dismiss.some(d => t.includes(d))) {
                    el.click(); return;
                }
            }
        }""")
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    time.sleep(0.5)


# ══════════════════════════════════════════════════════════════════════════════
# 核心步驟 1：列表頁搜尋運單，點入詳情
# ══════════════════════════════════════════════════════════════════════════════

def go_to_waybill_detail(page, waybill: str):
    """
    去 SF 運單列表頁
    → 真正按 Ctrl+F 輸入運單號（與手動一樣）
    → 關閉搜尋欄 → 點擊找到的運單進入詳情
    """
    print(f"  1a. 前往 SF 運單列表頁")
    page.goto(SF_LIST_URL, wait_until="domcontentloaded", timeout=30000)
    time.sleep(3)
    dismiss_popups(page)
    shot(page, f"{waybill}_0_list")

    # ── 真正按 Ctrl+F，輸入運單號，再關閉搜尋欄 ────────────────────────────
    print(f"  1b. Ctrl+F 搜尋 {waybill}")
    page.keyboard.press("Control+f")
    time.sleep(0.8)
    page.keyboard.type(waybill, delay=80)   # 逐字輸入，模擬手動打字
    time.sleep(1.0)
    page.keyboard.press("Escape")           # 關閉搜尋欄
    time.sleep(0.8)

    # ── 點擊含運單號的元素（Ctrl+F 已令頁面滾到並標示該位置）───────────────
    print(f"  1c. 點擊運單號進入詳情")
    for attempt in range(8):
        try:
            loc = page.get_by_text(waybill, exact=False).first
            loc.click(timeout=3000)
            print(f"  ✅ 已點入運單（attempt {attempt+1}）")
            time.sleep(3)
            dismiss_popups(page)
            return
        except Exception:
            pass
        time.sleep(1.5)

    # ── Fallback：直接跳 URL ─────────────────────────────────────────────────
    print(f"  ⚠️  點擊失敗，直接前往詳情頁 URL")
    page.goto(f"{SF_DETAIL_BASE}/{waybill}",
              wait_until="domcontentloaded", timeout=30000)
    time.sleep(3)
    dismiss_popups(page)


# ══════════════════════════════════════════════════════════════════════════════
# 核心：重新列印一張運單
# ══════════════════════════════════════════════════════════════════════════════

def reprint_one_waybill(ctx, entry: dict) -> bool:
    waybill  = str(entry.get("waybill",  "")).strip()
    customer = str(entry.get("customer", "")).strip()
    pdf_path = str(entry.get("pdf_path", "")).strip()

    print(f"\n{'='*55}")
    print(f"  運單: {waybill}  客人: {customer}")
    print(f"  覆蓋: {pdf_path}")
    print(f"{'='*55}")

    page = ctx.new_page()
    try:
        # ── 1: 列表頁找運單並點入 ────────────────────────────────────────────
        go_to_waybill_detail(page, waybill)
        shot(page, f"{waybill}_1_detail")

        # ── 2: 點「修改訂單」 ────────────────────────────────────────────────
        print("  2. 點「修改訂單」")
        clicked = False
        for attempt in range(8):
            clicked = page.evaluate("""() => {
                for (const el of document.querySelectorAll(
                        'button,a,[role="button"],span,div')) {
                    if (el.offsetParent === null) continue;
                    const t = (el.textContent || '').trim();
                    if (t === '修改訂單' || t === '修改订单') {
                        el.click(); return true;
                    }
                }
                return false;
            }""")
            if clicked:
                print(f"     OK (attempt {attempt + 1})")
                break
            time.sleep(1.5)

        if not clicked:
            print("     找不到「修改訂單」，跳過")
            return False

        time.sleep(3)
        shot(page, f"{waybill}_2_modify")

        # ── 3: 直接點「保存」（不改任何東西）───────────────────────────────
        print("  3. 直接點「保存」（不改任何東西）")
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(0.8)

        saved = False
        for attempt in range(8):
            saved = page.evaluate("""() => {
                // 只看按鈕文字是否含「保存」，不靠 class（避免 cancelAndSave 被誤篩）
                const targets = ['保存','儲存'];
                for (const el of document.querySelectorAll('button,[role="button"]')) {
                    if (el.offsetParent === null) continue;
                    const t = (el.textContent || '').trim();
                    // 文字本身不能是取消類
                    if (t === '取消' || t === '取消寄件') continue;
                    if (targets.some(c => t === c || t.endsWith(c))) {
                        el.click(); return true;
                    }
                }
                return false;
            }""")
            if saved:
                print(f"     OK (attempt {attempt + 1})")
                break
            time.sleep(1.5)

        if not saved:
            print("     找不到「保存」，跳過")
            return False

        print("     等待成功頁面...")
        time.sleep(5)
        shot(page, f"{waybill}_3_success")

        # ── 4: 點「列印電子運單」→ 等 modal 出現 ──────────────────────────────
        print("  4. 點「列印電子運單」")
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(0.5)

        clicked_print = False
        for attempt in range(6):
            clicked_print = page.evaluate("""() => {
                const labels = ['列印電子運單','打印電子運單'];
                for (const el of document.querySelectorAll(
                        'button,a,[role="button"],span')) {
                    if (el.offsetParent === null) continue;
                    const t = (el.textContent || '').trim();
                    if (labels.some(l => t === l || t.includes(l))) {
                        el.click(); return true;
                    }
                }
                return false;
            }""")
            if clicked_print:
                print(f"     OK (attempt {attempt + 1})")
                break
            time.sleep(1.5)

        if not clicked_print:
            print("     找不到「列印電子運單」，跳過")
            return False

        # 等 modal 彈出（圖3：面單預覽 modal）
        print("     等待面單 modal 出現...")
        time.sleep(3)

        # ── 5: 在 modal 裡點紅色「列印面單」→ 新 tab 開啟 → CDP 儲存 PDF ────
        print("  5. 點「列印面單」並儲存 PDF")
        try:
            with page.context.expect_page(timeout=20000) as new_pg_info:
                page.evaluate("""() => {
                    // 在 modal 裡找紅色「列印面單」按鈕
                    const labels = ['列印面單','打印面單'];
                    for (const el of document.querySelectorAll(
                            'button,a,[role="button"]')) {
                        if (el.offsetParent === null) continue;
                        const t = (el.textContent || '').trim();
                        if (labels.some(l => t === l || t.includes(l))) {
                            el.click(); return true;
                        }
                    }
                    return false;
                }""")

            print_page = new_pg_info.value
            print("     新 tab 已開啟（列印預覽），等待載入...")
            print_page.wait_for_load_state("domcontentloaded", timeout=15000)
            time.sleep(3)  # 等圖4的列印預覽完全渲染

            # 用 CDP 直接儲存 PDF（繞過 Chrome 列印對話框）
            cdp = page.context.new_cdp_session(print_page)
            res = cdp.send("Page.printToPDF", {
                "printBackground":    True,
                "preferCSSPageSize":  True,
                "paperWidth":         8.27,
                "paperHeight":        11.69,
                "marginTop":          0,
                "marginBottom":       0,
                "marginLeft":         0,
                "marginRight":        0,
            })
            pdf_bytes = base64.b64decode(res["data"])
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            with open(pdf_path, "wb") as f:
                f.write(pdf_bytes)
            cdp.detach()
            print_page.close()
            print(f"  ✅ PDF 已覆蓋：{pdf_path}")

        except Exception as pe:
            print(f"     PDF 儲存失敗：{pe}")

        return True

    except Exception as e:
        print(f"\n  {customer} ({waybill}) 失敗：{e}")
        import traceback; traceback.print_exc()
        return False
    finally:
        try:
            page.close()
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    print("=" * 55)
    print("  fix_waybill_v6 — 重新列印順丰運單")
    print("=" * 55)

    # A. 讀 session
    session = read_session()

    # B. 確認視窗
    confirmed = show_confirm_gui(session)
    if not confirmed:
        sys.exit(0)

    # C. 關閉殘留 Chrome
    subprocess.run(
        ["powershell", "-Command",
         "Get-WmiObject Win32_Process"
         " | Where-Object { $_.CommandLine -like '*ChromeAutomation*' }"
         " | ForEach-Object { $_.Terminate() }"],
        capture_output=True,
    )
    time.sleep(1.5)
    for lf in ["lockfile", "SingletonLock", "SingletonSocket", "SingletonCookie"]:
        try:
            os.remove(os.path.join(CHROME_PROFILE, lf))
        except Exception:
            pass

    # D. 開 Chrome，逐一重新列印
    success_count = 0
    failed_count  = 0

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            CHROME_PROFILE,
            channel  = "chrome",
            headless = False,
            args     = ["--disable-blink-features=AutomationControlled"],
            slow_mo  = 150,
            viewport = {"width": 1280, "height": 900},
        )

        for entry in session:
            ok = reprint_one_waybill(ctx, entry)
            if ok:
                success_count += 1
            else:
                failed_count += 1
            time.sleep(1)

        print(f"\n{'='*55}")
        print(f"  完成：✅ {success_count} 成功  ❌ {failed_count} 失敗")
        print(f"{'='*55}")

        input("\n按 Enter 關閉瀏覽器…")
        ctx.close()
