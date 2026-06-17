# -*- coding: utf-8 -*-
"""
backfill_pos_docs.py — 為已有的訂單文件夾補下載 收貨明細 + 清關PDF
=========================================================
掃描 data/orders/ 所有文件夾，若缺少 收貨明細 或 清關 PDF，
回到 POS 銷售記錄下載並存入同一文件夾，最後 git push 更新 Streamlit。

執行：
  python "C:\Users\user\Desktop\順丰E順递\scripts\backfill_pos_docs.py"
"""
import os
import re
import sys
import time
import subprocess

sys.stdout.reconfigure(encoding="utf-8")

from playwright.sync_api import sync_playwright

CHROME_PROFILE = r"C:\ChromeAutomation"
ORDERS_DIR     = r"C:\Users\user\Desktop\順丰E順递\data\orders"
_REPO          = r"C:\Users\user\Desktop\順丰E順递"
POS_URL        = "https://online-store-99126206.web.app/"
POS_PASS       = "0000"


def _needs_backfill(folder_path: str) -> tuple[bool, bool]:
    """Return (missing_receipt, missing_customs) for a folder."""
    files = os.listdir(folder_path)
    has_receipt = any("收貨明細" in f for f in files)
    has_customs = any("清關" in f for f in files)
    return (not has_receipt), (not has_customs)


def _download_missing(page, pos_order_no: str, save_dir: str,
                       folder_name: str, want_receipt: bool, want_customs: bool):
    """Log into POS, search order, download whichever files are missing."""
    # Derive file_base from folder name (strip the ORD-xxx part already has customer+date)
    # folder_name e.g. "黄业伟_20260430_ORD-752035"
    file_base = folder_name   # use folder name directly as base

    try:
        page.goto(POS_URL, wait_until="domcontentloaded", timeout=20000)
        time.sleep(3)

        page.locator("button:has-text('后台管理')").first.click()
        time.sleep(0.8)
        page.locator("input[type='password']").first.fill(POS_PASS)
        page.keyboard.press("Enter")
        time.sleep(2)

        # 點「記錄」nav tab
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

        search = page.locator("input[placeholder*='搜尋單號']").first
        search.wait_for(state="visible", timeout=5000)
        search.click()
        search.fill(pos_order_no)
        time.sleep(2)
        print(f"  ✅ 搜尋 {pos_order_no}")

        if want_receipt:
            try:
                with page.expect_download(timeout=12000) as dl_info:
                    page.evaluate("""() => {
                        for (const a of document.querySelectorAll('a[download]')) {
                            if (a.offsetParent === null) continue;
                            if ((a.getAttribute('download') || '').includes('明細')) {
                                a.click(); return true;
                            }
                        }
                        for (const btn of document.querySelectorAll('button')) {
                            if (btn.offsetParent === null) continue;
                            if ((btn.textContent || '').trim() === '收貨明細') {
                                btn.click(); return true;
                            }
                        }
                    }""")
                dl = dl_info.value
                ext = os.path.splitext(dl.suggested_filename)[1] or ".pdf"
                path = os.path.join(save_dir, f"{file_base}_收貨明細{ext}")
                dl.save_as(path)
                print(f"  ✅ 收貨明細 → {os.path.basename(path)}")
            except Exception as e:
                print(f"  ⚠️  收貨明細下載失敗: {e}")

        if want_customs:
            try:
                with page.expect_download(timeout=12000) as dl_info:
                    page.evaluate("""() => {
                        for (const a of document.querySelectorAll('a[download]')) {
                            if (a.offsetParent === null) continue;
                            if ((a.getAttribute('download') || '').includes('清關')) {
                                a.click(); return true;
                            }
                        }
                        for (const btn of document.querySelectorAll('button')) {
                            if (btn.offsetParent === null) continue;
                            if ((btn.textContent || '').trim() === '清關PDF') {
                                btn.click(); return true;
                            }
                        }
                    }""")
                dl = dl_info.value
                ext = os.path.splitext(dl.suggested_filename)[1] or ".pdf"
                path = os.path.join(save_dir, f"{file_base}_清關{ext}")
                dl.save_as(path)
                print(f"  ✅ 清關PDF → {os.path.basename(path)}")
            except Exception as e:
                print(f"  ⚠️  清關PDF下載失敗: {e}")

    except Exception as e:
        print(f"  ❌ {pos_order_no} 失敗: {e}")


def main():
    # ── 掃描需要補檔的文件夾 ─────────────────────────────────────────────────
    ORD_RE = re.compile(r"(ORD-\d+)")
    todo = []
    for folder in sorted(os.listdir(ORDERS_DIR)):
        folder_path = os.path.join(ORDERS_DIR, folder)
        if not os.path.isdir(folder_path):
            continue
        m = ORD_RE.search(folder)
        if not m:
            continue
        pos_order_no = m.group(1)
        want_receipt, want_customs = _needs_backfill(folder_path)
        if want_receipt or want_customs:
            todo.append({
                "folder_name":  folder,
                "folder_path":  folder_path,
                "pos_order_no": pos_order_no,
                "want_receipt": want_receipt,
                "want_customs": want_customs,
            })

    if not todo:
        print("✅ 全部文件夾都已有收貨明細及清關PDF，無需補檔")
        return

    print(f"需要補檔的訂單：{len(todo)} 個")
    for t in todo:
        missing = []
        if t["want_receipt"]: missing.append("收貨明細")
        if t["want_customs"]: missing.append("清關PDF")
        print(f"  {t['folder_name']}  →  缺少：{', '.join(missing)}")

    print("\n開始補下載…")

    # ── 啟動 Chrome ──────────────────────────────────────────────────────────
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            CHROME_PROFILE, channel="chrome", headless=False,
            args=["--disable-blink-features=AutomationControlled"],
            slow_mo=150, viewport={"width": 1280, "height": 900},
        )

        ok_count = fail_count = 0
        for t in todo:
            print(f"\n{'='*55}")
            print(f"  {t['folder_name']}")
            print(f"{'='*55}")
            pos_page = ctx.new_page()
            try:
                _download_missing(
                    pos_page,
                    t["pos_order_no"],
                    t["folder_path"],
                    t["folder_name"],
                    t["want_receipt"],
                    t["want_customs"],
                )
                ok_count += 1
            except Exception as e:
                print(f"  ❌ 失敗: {e}")
                fail_count += 1
            finally:
                pos_page.close()
            time.sleep(1)

        input("\n按 Enter 關閉瀏覽器…")
        ctx.close()

    # ── Git push 更新 Streamlit ──────────────────────────────────────────────
    print("\n☁️  同步到 GitHub…")
    try:
        subprocess.run(["git", "-C", _REPO, "add", "data/orders"],
                       capture_output=True, check=True)
        subprocess.run(["git", "-C", _REPO, "commit", "-m",
                        f"backfill: 補收貨明細+清關PDF ({ok_count} 個)"],
                       capture_output=True, check=True)
        subprocess.run(["git", "-C", _REPO, "push", "origin", "main"],
                       capture_output=True, check=True)
        print("  ✅ 已推送到 GitHub，Streamlit 約 30 秒後更新")
    except Exception as e:
        print(f"  ⚠️  推送失敗: {e}")

    print(f"\n完成：✅ {ok_count} 成功  ❌ {fail_count} 失敗")


if __name__ == "__main__":
    main()
