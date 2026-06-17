# -*- coding: utf-8 -*-
"""
download_waybills_camp.py
=========================
從 camp.sf-express.com/ScanPrint 下載運單PDF，
並按 {人名}_{日期}_{運單號} 格式儲存到 data/orders/
"""
import os, sys, time, subprocess
from datetime import date
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")

CHROME_PROFILE = r"C:\ChromeAutomation"
ORDERS_DIR     = r"C:\Users\user\Desktop\順丰E順递\data\orders"
SCAN_URL       = "https://camp.sf-express.com/ScanPrint"

# ── 要下載的運單列表（運單號, 客戶名）────────────────────────────────────
ORDERS = [
    ("SF0215123734807", "麦华照"),
    ("SF0215123753217", "方启源"),
]
# ─────────────────────────────────────────────────────────────────────────────

TODAY = date.today().strftime("%Y%m%d")


def dismiss_dialogs(page, max_tries=5):
    for _ in range(max_tries):
        try:
            if page.locator(".el-message-box__wrapper").is_visible(timeout=1500):
                btn = page.locator(".el-message-box__btns button.el-button--primary")
                if btn.is_visible(timeout=1000):
                    btn.click()
                    time.sleep(0.8)
                    continue
        except Exception:
            pass
        break


def run():
    subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], capture_output=True)
    time.sleep(2)

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            CHROME_PROFILE, channel="chrome", headless=False,
            slow_mo=200,
            args=["--disable-blink-features=AutomationControlled", "--disable-infobars"],
            viewport={"width": 1280, "height": 800},
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(SCAN_URL, wait_until="domcontentloaded")
        time.sleep(4)
        dismiss_dialogs(page)

        print("選擇「下載到本地」...")
        page.locator(".el-select").first.click()
        time.sleep(1)
        page.locator("li.el-select-dropdown__item", has_text="下載到本地").click()
        time.sleep(1)
        dismiss_dialogs(page)

        scan_input = page.locator("input[placeholder='此處為掃描結果']")

        for waybill, customer in ORDERS:
            folder = os.path.join(ORDERS_DIR, f"{customer}_{TODAY}_{waybill}")
            os.makedirs(folder, exist_ok=True)
            dest_path = os.path.join(folder, f"{customer}_{TODAY}_{waybill}_運單.pdf")

            if os.path.exists(dest_path):
                print(f"已存在，跳過：{os.path.basename(dest_path)}")
                continue

            print(f"打印 {waybill} ({customer})...")
            dismiss_dialogs(page)
            scan_input.click()
            scan_input.fill(waybill)
            time.sleep(0.5)

            # 同時等待兩個回應：scanPrint JSON + OSS 二進制 PDF
            with page.expect_response(
                lambda r: "eos-scp-core" in r.url and ".pdf" in r.url,
                timeout=25000
            ) as pdf_resp_info:
                page.locator("button", has_text="打印").click()

            pdf_resp = pdf_resp_info.value
            pdf_bytes = pdf_resp.body()
            print(f"  捕捉到 PDF: {len(pdf_bytes)} bytes")

            if len(pdf_bytes) > 1000:
                with open(dest_path, "wb") as f:
                    f.write(pdf_bytes)
                print(f"  已儲存：{dest_path}")
            else:
                print(f"  PDF太小，可能有問題：{pdf_bytes[:100]}")

            time.sleep(1.5)

        ctx.close()
        print("\n全部完成！")
        for waybill, customer in ORDERS:
            dest = os.path.join(ORDERS_DIR, f"{customer}_{TODAY}_{waybill}",
                                f"{customer}_{TODAY}_{waybill}_運單.pdf")
            status = "✓" if os.path.exists(dest) else "✗ 缺失"
            print(f"  {status} {dest}")


if __name__ == "__main__":
    run()
