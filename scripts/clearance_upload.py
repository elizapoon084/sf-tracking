#!/usr/bin/env python3
# clearance_upload.py — SF Express HK 清關自動上傳（證照單證 + 清關單證）
# Usage: python clearance_upload.py
#   勾選「試跑模式」可預覽文件路徑而不真正提交

import os
import sys
import json
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ─── Config ───────────────────────────────────────────────────────────────────
CHROME_PROFILE = r"C:\ChromeAutomation"
SESSION_FILE   = r"C:\Users\user\Desktop\順丰E順递\data\last_session.json"
ID_FOLDER      = r"C:\Users\user\Desktop\順丰E順递\身份証"
CLEARANCE_URL  = "https://hk.sf-express.com/hk/tc/clearance?type=upload"
# Iframe hosted at sf-international.com (loaded after clicking the upload tab)
IFRAME_HOST    = "sf-international.com"

# ─── Simplified ↔ Traditional mapping (common surname/given-name chars) ───────
_S2T = {
    # 常見姓氏
    '叶':'葉','张':'張','陈':'陳','刘':'劉','黄':'黃','赵':'趙','吴':'吳',
    '郑':'鄭','孙':'孫','罗':'羅','冯':'馮','钱':'錢','苏':'蘇','杨':'楊',
    '邓':'鄧','蒋':'蔣','韩':'韓','谢':'謝','许':'許','邹':'鄒','顾':'顧',
    '萧':'蕭','龚':'龔','卢':'盧','叶':'葉','马':'馬','凤':'鳳',
    # 常見名字字
    '业':'業','泽':'澤','鸿':'鴻','兰':'蘭','丽':'麗','华':'華',
    '国':'國','东':'東','龙':'龍','长':'長','来':'來','远':'遠','财':'財',
    '贵':'貴','宝':'寶','军':'軍','强':'強','胜':'勝','飞':'飛','鹏':'鵬',
    '伟':'偉','绍':'紹','云':'雲','灿':'燦','锋':'鋒','杰':'傑',
    '晓':'曉','卫':'衛','恺':'愷','玮':'瑋','辉':'輝',
    '颖':'穎','莲':'蓮','涛':'濤','洁':'潔','维':'維','凯':'凱',
    '艺':'藝','联':'聯','绿':'綠','银':'銀','编':'編','铭':'銘',
    '秀':'秀','春':'春','梅':'梅','燕':'燕','芳':'芳','娜':'娜',
    '俊':'俊','峰':'峰','珊':'珊','婷':'婷','蕾':'蕾',
    # 補充缺少的字
    '蔼':'藹','荣':'榮','宁':'寧','义':'義','风':'風','时':'時',
    '丰':'豐','务':'務','带':'帶','层':'層','变':'變','话':'話',
    '进':'進','问':'問','间':'間','关':'關','开':'開','电':'電',
    '头':'頭','经':'經','给':'給','绕':'繞','总':'總','设':'設',
}
_T2S = {v: k for k, v in _S2T.items()}


def _normalize(name: str) -> str:
    """Convert all chars to traditional Chinese using the mapping table."""
    return ''.join(_S2T.get(c, c) for c in name)


def find_id_folder(customer: str):
    """Return path to the customer's ID card subfolder, or None.

    Tries exact match first, then simplified→traditional normalization,
    then a character-overlap fuzzy fallback (allows 1 mismatch for 3-char names).
    """
    cust_norm = _normalize(customer)
    best_path, best_score = None, 0

    for d in os.listdir(ID_FOLDER):
        full = os.path.join(ID_FOLDER, d)
        if not os.path.isdir(full):
            continue
        d_norm = _normalize(d)
        # Exact or normalized exact match
        if d == customer or d_norm == cust_norm:
            return full
        # Fuzzy: count matching normalized characters
        score = sum(1 for a, b in zip(d_norm, cust_norm) if a == b)
        if score > best_score:
            best_score = score
            best_path = full

    # Accept fuzzy match if at least (len-1) chars match
    if best_path and best_score >= max(1, len(customer) - 1):
        return best_path
    return None


def find_id_cards(customer: str):
    """Return (front_path, back_path).  Either may be None if not found."""
    folder = find_id_folder(customer)
    if not folder:
        return None, None

    front = back = None
    for f in os.listdir(folder):
        fl = f.lower()
        fp = os.path.join(folder, f)
        if '正面' in f or 'front' in fl:
            front = fp
        elif '背面' in f or '背囸' in f or 'back' in fl:
            back = fp

    return front, back


def derive_pdf_paths(entry: dict):
    """Derive 小票 and 清關 PDF paths from a last_session.json entry."""
    pdf_path = entry.get('pdf_path', '')
    waybill  = entry.get('waybill', '')
    if not pdf_path:
        return None, None

    folder   = os.path.dirname(pdf_path)
    basename = os.path.basename(pdf_path)          # e.g. 張三_20260501_ORD-1_SF123_運單.pdf

    suffix = f'_{waybill}_運單.pdf'
    if basename.endswith(suffix):
        base = basename[: -len(suffix)]            # → 張三_20260501_ORD-1
    else:
        base = basename.replace('_運單.pdf', '').replace(f'_{waybill}', '')

    xiaopiao  = os.path.join(folder, f'{base}.pdf')
    # 搵資料夾中任何含「明細+清關」的 PDF（檔名格式可能不一）
    guanzheng = None
    if os.path.isdir(folder):
        for fn in os.listdir(folder):
            if '明細+清關' in fn and fn.endswith('.pdf'):
                guanzheng = os.path.join(folder, fn)
                break
    if guanzheng is None:
        # 舊檔名 fallback
        guanzheng = os.path.join(folder, f'{base}_清關.pdf')
    # 新流程只有合一PDF，沒有獨立小票 → 用合一PDF代替
    if not os.path.exists(xiaopiao) and guanzheng and os.path.exists(str(guanzheng)):
        xiaopiao = guanzheng
    return xiaopiao, guanzheng


# ─── GUI ──────────────────────────────────────────────────────────────────────
class ClearanceApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SF Express 清關自動上傳")
        self.root.geometry("780x600")
        self.root.resizable(True, True)

        self.dry_run = tk.BooleanVar(value=False)
        self.entries: list = []
        self.checks: list  = []
        self._build_ui()
        self._load_sessions()

    # ── UI build ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── header ──
        top = tk.Frame(self.root, pady=8, padx=12)
        top.pack(fill='x')
        tk.Label(top, text="SF Express 清關自動上傳",
                 font=('Arial', 14, 'bold')).pack(side='left')
        tk.Checkbutton(top, text="🔍 試跑模式（不提交）", variable=self.dry_run,
                       font=('Arial', 11), fg='#666').pack(side='right')

        # ── scrollable entry list ──
        lf = tk.LabelFrame(self.root, text="選擇要處理的運單",
                           padx=8, pady=6, font=('Arial', 11, 'bold'))
        lf.pack(fill='both', expand=True, padx=10, pady=4)

        canvas = tk.Canvas(lf)
        sb = ttk.Scrollbar(lf, orient='vertical', command=canvas.yview)
        self.sf = tk.Frame(canvas)
        self.sf.bind('<Configure>',
                     lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self.sf, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # column headers
        hdr = tk.Frame(self.sf, bg='#e8e8e8')
        hdr.pack(fill='x', pady=(0, 3))
        for text, w in [('✓', 3), ('客戶', 10), ('運單號', 22),
                        ('身份証', 8), ('小票PDF', 8), ('清關PDF', 8)]:
            tk.Label(hdr, text=text, width=w, anchor='w', bg='#e8e8e8',
                     font=('Arial', 10, 'bold')).pack(side='left')

        # ── bottom buttons ──
        bf = tk.Frame(self.root, pady=8, padx=10)
        bf.pack(fill='x')
        tk.Button(bf, text='全選',   command=self._sel_all,   width=7).pack(side='left', padx=3)
        tk.Button(bf, text='全不選', command=self._desel_all, width=7).pack(side='left', padx=3)
        self.run_btn = tk.Button(bf, text='▶ 開始上傳', command=self._run,
                                 bg='#1976D2', fg='white',
                                 font=('Arial', 12, 'bold'), width=14)
        self.run_btn.pack(side='right', padx=4)

        # ── log area ──
        logf = tk.LabelFrame(self.root, text='執行日誌', padx=6, pady=4)
        logf.pack(fill='x', padx=10, pady=(0, 8))
        self.log_text = tk.Text(logf, height=7, font=('Consolas', 9), state='disabled')
        log_sb = ttk.Scrollbar(logf, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sb.set)
        log_sb.pack(side='right', fill='y')
        self.log_text.pack(fill='x')

    # ── load sessions ─────────────────────────────────────────────────────────
    def _load_sessions(self):
        if not os.path.exists(SESSION_FILE):
            self.log(f"❌ 找不到: {SESSION_FILE}")
            return
        with open(SESSION_FILE, encoding='utf-8') as f:
            sessions = json.load(f)
        if not sessions:
            self.log("⚠️  last_session.json 是空的")
            return

        self.entries = sessions
        for s in sessions:
            customer  = s.get('customer', '?')
            waybill   = s.get('waybill', '?')
            front, back = find_id_cards(customer)
            xp, gk     = derive_pdf_paths(s)

            id_ok = '✅' if (front and back
                              and os.path.exists(front)
                              and os.path.exists(back)) else '❌'
            xp_ok = '✅' if (xp and os.path.exists(xp)) else '❌'
            gk_ok = '✅' if (gk and os.path.exists(gk)) else '❌'

            row = tk.Frame(self.sf, pady=1)
            row.pack(fill='x')

            var = tk.BooleanVar(value=True)
            self.checks.append(var)
            tk.Checkbutton(row, variable=var, width=2).pack(side='left')
            tk.Label(row, text=customer, width=10, anchor='w',
                     font=('Arial', 10)).pack(side='left')
            tk.Label(row, text=waybill, width=22, anchor='w',
                     font=('Consolas', 9)).pack(side='left')
            for badge, w in [(id_ok, 8), (xp_ok, 8), (gk_ok, 8)]:
                tk.Label(row, text=badge, width=w, anchor='w').pack(side='left')

        self.log(f"✅ 載入 {len(sessions)} 筆訂單")

    # ── helpers ───────────────────────────────────────────────────────────────
    def log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n')
        self.log_text.see('end')
        self.log_text.configure(state='disabled')
        self.root.update_idletasks()

    def _sel_all(self):
        for v in self.checks: v.set(True)

    def _desel_all(self):
        for v in self.checks: v.set(False)

    # ── run ───────────────────────────────────────────────────────────────────
    def _run(self):
        selected = [self.entries[i] for i in range(len(self.entries))
                    if self.checks[i].get()]
        if not selected:
            messagebox.showwarning("未選擇", "請先勾選至少一筆訂單")
            return

        dry = self.dry_run.get()
        if not dry and not messagebox.askyesno(
                "確認上傳", f"即將上傳 {len(selected)} 筆清關資料，確認？"):
            return
        if dry:
            self.log("🔍 試跑模式 — 只顯示文件路徑，不會真正提交")

        self.run_btn.configure(state='disabled', text='執行中…')
        threading.Thread(target=self._worker, args=(selected, dry),
                         daemon=True).start()

    def _worker(self, entries, dry_run):
        try:
            run_clearance(entries, dry_run, log_fn=self.log)
            self.root.after(0, lambda: messagebox.showinfo("完成", "清關上傳全部完成！"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("錯誤", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.configure(
                state='normal', text='▶ 開始上傳'))

    def run(self):
        self.root.mainloop()


# ─── Playwright automation ────────────────────────────────────────────────────

def _get_iframe(page):
    """Return the sf-international iframe frame after the upload tab is clicked."""
    page.goto(CLEARANCE_URL, wait_until='domcontentloaded', timeout=30000)
    time.sleep(5)

    # 偵測是否被重定向到登入頁
    cur_url = page.url
    if 'login' in cur_url.lower() or 'sign' in cur_url.lower():
        raise Exception(
            f"SF 網站未登入！已跳轉到：{cur_url}\n"
            "請用 Chrome 開啟 C:\\ChromeAutomation 配置檔手動登入 SF HK 網站，再重跑程式。"
        )

    # 點擊「清關證件圖片上傳」Tab（不是 button，是 tab 標籤）
    _btn_labels = [
        '清關證件圖片上傳', '清关证件图片上传',
        '清關上傳', '清關單証上傳', '上傳', '证照',
    ]
    clicked = False
    for _lbl in _btn_labels:
        try:
            page.click(f'text={_lbl}', timeout=4000)
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        # 印出頁面所有可見文字幫助排查
        try:
            visible = page.evaluate("() => document.body.innerText.slice(0, 500)")
            print(f"  [debug] 頁面可見文字(前500): {visible}")
        except Exception:
            pass

    # 等 iframe 載入，最多 25 秒重試
    for _attempt in range(25):
        for fr in page.frames:
            if IFRAME_HOST in fr.url:
                return fr
        time.sleep(1)

    # 截圖儲存以便排查
    _ss_path = os.path.join(os.path.dirname(SESSION_FILE), 'clearance_debug.png')
    try:
        page.screenshot(path=_ss_path)
        print(f"  [debug] 截圖已儲存: {_ss_path}")
        print(f"  [debug] 當前 URL: {page.url}")
        print(f"  [debug] 所有 frames: {[f.url for f in page.frames]}")
    except Exception:
        pass

    raise Exception(
        f"找不到 {IFRAME_HOST} iframe（清關頁面未正確載入）\n"
        "最可能原因：SF HK 網站 session 過期，需要重新登入。\n"
        f"截圖已儲存到: {_ss_path}\n"
        "請打開 Chrome（使用 C:\\ChromeAutomation 配置），登入 https://hk.sf-express.com/ 後再試。"
    )


def _upload_via_chooser(page, frame, css_area: str, filepath: str, label: str, log_fn):
    """Click an Ant Design upload area (span[role=button]) and set files via file chooser."""
    log_fn(f"    上傳 {label}: {os.path.basename(filepath)}")
    with page.expect_file_chooser(timeout=8000) as fc:
        frame.locator(css_area).click()
    fc.value.set_files(filepath)


def _upload_customs_file(page, frame, filepath: str, label: str, log_fn):
    """Upload a file to the currently active 清關単証 sub-tab.

    Tries two strategies:
    1. Click visible '點擊上傳' link → file chooser
    2. Directly set files on visible input[type=file] (hidden-input pattern)
    """
    log_fn(f"    上傳 {label}: {os.path.basename(filepath)}")

    # Strategy 1: file chooser via visible '點擊上傳' link
    links = frame.locator('text=點擊上傳').all()
    for link in links:
        try:
            if link.is_visible():
                with page.expect_file_chooser(timeout=8000) as fc:
                    link.click()
                fc.value.set_files(filepath)
                log_fn(f"    (方法1 file-chooser 完成)")
                return
        except Exception:
            continue

    # Strategy 2: directly set files on a visible/accessible input[type=file]
    inputs = frame.locator('input[type="file"]').all()
    for inp in inputs:
        try:
            inp.set_input_files(filepath)
            log_fn(f"    (方法2 hidden-input 完成)")
            return
        except Exception:
            continue

    raise Exception(f"找不到可見的上傳按鈕（{label}）")


def _handle_confirm_modal(page, frame, log_fn):
    """Click confirm button in any post-submit modal.

    Handles two modal types:
    - 同意 confirmation: .ant-modal-confirm-btns (ant-btn-danger or last button)
    - 成功 notification: .ant-modal-footer (ant-btn-primary, e.g. 確認)
    """
    try:
        # Wait for any modal to appear
        frame.wait_for_selector('.ant-modal-wrap:not([style*="display: none"])', timeout=10000)
        time.sleep(0.5)
        # Try confirm-style modal first (danger button = 同意)
        try:
            frame.click('.ant-modal-confirm-btns .ant-btn-danger', timeout=3000)
            log_fn("    ✓ 彈窗：點擊「同意」")
        except Exception:
            try:
                frame.click('.ant-modal-confirm-btns button:last-child', timeout=3000)
                log_fn("    ✓ 彈窗：點擊最後按鈕")
            except Exception:
                # Success/notification modal (e.g. 清關單証上傳成功 → 確認)
                try:
                    frame.click('.ant-modal-footer .ant-btn-primary', timeout=3000)
                    log_fn("    ✓ 彈窗：點擊「確認」")
                except Exception:
                    # Last resort: click any visible button in the modal
                    frame.locator('.ant-modal button:visible').last.click(force=True)
                    log_fn("    ✓ 彈窗：點擊可見按鈕")
        # Wait for modal to close
        try:
            frame.wait_for_selector('.ant-modal-wrap', state='hidden', timeout=8000)
        except Exception:
            time.sleep(2)
        time.sleep(3)  # 按同意後等 3 秒才進下一步
    except Exception:
        pass   # no modal = nothing to do


def _do_id_tab(page, frame, front: str, back: str, log_fn):
    """証照單証 tab: upload front+back ID → agree → submit → handle modal."""
    log_fn("  📋 Step 1: 証照単証")

    # Upload front (身份証個人面)
    _upload_via_chooser(page, frame,
                        '.parent-content.active .firstUpload span[role="button"]',
                        front, '正面', log_fn)
    log_fn("    等待 OCR…")
    time.sleep(5)   # OCR takes a few seconds;姓名/号码 fields appear after this

    # Upload back (身份証國徽面)
    _upload_via_chooser(page, frame,
                        '.parent-content.active .lastUpload span[role="button"]',
                        back, '背面', log_fn)
    time.sleep(3)

    # Agreement checkbox
    chk = frame.locator('input[type="checkbox"]').first
    if not chk.is_checked():
        chk.evaluate('el => el.click()')
        log_fn("    ✓ 同意條款")
    time.sleep(0.5)

    # Submit — force=True bypasses Playwright's CSS/aria disabled check
    submit = frame.locator('button.search-btn').first
    submit.click(force=True)
    log_fn("    提交中…")
    time.sleep(1)
    _handle_confirm_modal(page, frame, log_fn)
    log_fn("    ✅ 証照單証完成")


def _do_customs_tab(page, frame, xiaopiao: str, guanzheng: str, log_fn):
    """清關単証 tab: upload 清關PDF → submit. Returns True on success."""
    log_fn("  📦 Step 2: 清關單証")

    # Ensure any lingering modal overlay is fully gone before clicking the tab
    try:
        frame.wait_for_selector('.ant-modal-wrap', state='hidden', timeout=8000)
    except Exception:
        try:
            frame.locator('.ant-modal-confirm-btns button').last.click(force=True)
            frame.wait_for_selector('.ant-modal-wrap', state='hidden', timeout=5000)
        except Exception:
            time.sleep(2)

    # Switch to 清關単証 tab (second .van-tab)
    frame.locator('.van-tab').nth(1).click()
    time.sleep(1)

    # Handle "未提交資料" switch-confirmation modal (click 確認 to proceed)
    try:
        frame.wait_for_selector('.ant-modal-confirm-btns', timeout=4000)
        frame.locator('.ant-modal-confirm-btns button').last.click()
        log_fn("    ✓ 確認切換頁面")
        time.sleep(2)
    except Exception:
        pass   # no confirmation needed

    # Log available sub-tabs for debugging
    try:
        tabs = [t.inner_text() for t in frame.locator('.van-tab, .ant-tabs-tab').all()]
        log_fn(f"    子標籤: {tabs}")
    except Exception:
        pass

    # --- 發票 sub-tab: upload combined PDF (receipt + packing + customs) ---
    try:
        frame.get_by_text('發票', exact=True).first.click()
        time.sleep(1)
        log_fn("    ✓ 已切換到「發票」子標籤")
    except Exception:
        log_fn("    (找不到「發票」子標籤，在預設欄上傳)")

    log_fn("    [發票] 上傳清關PDF")
    _upload_customs_file(page, frame, guanzheng, '清關PDF(含小票+明細)', log_fn)
    time.sleep(5)

    # Wait up to 60s for server to confirm uploads
    log_fn("    等待伺服器確認上傳…")
    # Use .liquidation selector to target 清關単証 panel specifically (not hidden 証照 panel)
    deadline = time.time() + 60
    while time.time() < deadline:
        enabled = frame.evaluate("""() => {
            const btn = document.querySelector('.liquidation button.search-btn')
                     || document.querySelector('.van-tab__pane button.search-btn:not([style*="display: none"])');
            return btn ? !btn.disabled : false;
        }""")
        if enabled:
            break
        time.sleep(2)

    # Use JavaScript click — bypasses disabled attribute, same as user manually clicking in browser
    log_fn("    點擊提交（JS click）…")
    frame.evaluate("""() => {
        const btn = document.querySelector('.liquidation button.search-btn')
                 || document.querySelector('.van-tab__pane button.search-btn');
        if (btn) btn.click();
    }""")
    log_fn("    提交中…")
    time.sleep(4)
    _handle_confirm_modal(page, frame, log_fn)
    log_fn("    ✅ 清關単証完成")
    return True


# ── Keep pagination helpers for possible future use ───────────────────────────
def _next_page(page) -> bool:
    """Click the pagination Next button. Returns False if already on last page.

    SF HK waybill list uses <span class="pagination_navBtn__SGt2O"> for ALL
    pagination controls.  The LAST span is the Next (>) arrow; when on the
    final page it gains pagination_disabled__eyBdH.
    """
    try:
        all_nav = page.locator('span[class*="pagination_navBtn"]').all()
        if not all_nav:
            return False
        next_btn = all_nav[-1]          # last span = Next arrow
        cls = next_btn.get_attribute('class') or ''
        if 'disabled' in cls:
            return False                # already on last page
        next_btn.click()
        time.sleep(2)
        return True
    except Exception:
        pass
    return False


def _find_and_open_waybill(page, waybill: str, log_fn, max_pages=50) -> bool:
    """
    Paginate through /waybill/list searching for the waybill number
    (equivalent to pressing Ctrl+F on each page).
    When found, click the clearance icon button in that row.
    Returns True if successfully clicked.
    """
    for page_num in range(1, max_pages + 1):
        log_fn(f"    第 {page_num} 頁掃描中…")
        time.sleep(1.5)

        if waybill in page.content():
            log_fn(f"    ✅ 第 {page_num} 頁找到 {waybill}")

            # Locate the row that contains this waybill text
            # The row could be a <tr>, <li>, or a <div> block
            row_selectors = [
                f'tr:has-text("{waybill}")',
                f'li:has-text("{waybill}")',
                f'.waybill-item:has-text("{waybill}")',
                f'[class*="item"]:has-text("{waybill}")',
                f'[class*="row"]:has-text("{waybill}")',
            ]
            row = None
            for sel in row_selectors:
                try:
                    el = page.locator(sel).first
                    if el.count() > 0:
                        row = el
                        break
                except Exception:
                    continue

            if row is None:
                log_fn(f"    ⚠️  找到文字但定位不到行元素，嘗試直接點擊文字")
                # Fallback: click the waybill text itself
                try:
                    page.locator(f'text={waybill}').first.click()
                    time.sleep(2)
                    return True
                except Exception as e:
                    log_fn(f"    ❌ 點擊失敗: {e}")
                    return False

            # Log all icon buttons in the row so user can confirm which is correct
            try:
                btns = row.locator('button, a[role="button"], span[role="button"], i').all()
                log_fn(f"    該行有 {len(btns)} 個操作按鈕（點擊第 {CLEARANCE_BTN_INDEX+1} 個）")
            except Exception:
                pass

            # Click the designated clearance icon button
            try:
                row.locator('button, a[role="button"], span[role="button"]').nth(
                    CLEARANCE_BTN_INDEX).click()
                time.sleep(2)
                return True
            except Exception as e:
                log_fn(f"    ❌ 點擊清關按鈕失敗: {e}")
                return False

        # Not on this page — go to next
        if not _next_page(page):
            log_fn(f"    已是最後一頁（共掃 {page_num} 頁），找不到 {waybill}")
            return False

    log_fn(f"    超過 {max_pages} 頁仍找不到 {waybill}")
    return False


def _upload_id_tab(page, front: str, back: str, log_fn):
    """證照單證 tab: upload ID front/back → agree → 提交."""
    log_fn("  📋 Step 1: 證照單證")

    try:
        page.click('text=證照單證', timeout=5000)
        time.sleep(1)
    except Exception:
        log_fn("    (證照單證 已是預設 tab)")

    # Upload front ID
    log_fn(f"    上傳正面: {os.path.basename(front)}")
    try:
        page.locator('input[type="file"]').nth(0).set_input_files(front)
        time.sleep(1)
    except Exception as e:
        log_fn(f"    ⚠️  上傳正面失敗: {e}")

    # Upload back ID
    log_fn(f"    上傳背面: {os.path.basename(back)}")
    try:
        page.locator('input[type="file"]').nth(1).set_input_files(back)
    except Exception as e:
        log_fn(f"    ⚠️  上傳背面失敗: {e}")

    log_fn("    等待 3 秒…")
    time.sleep(3)

    # Check agreement checkbox
    try:
        for chk in page.locator('input[type="checkbox"]').all():
            if not chk.is_checked():
                chk.click()
                log_fn("    ✓ 已勾選同意條款")
                break
    except Exception:
        log_fn("    (找不到 checkbox，略過)")

    try:
        page.locator('button:has-text("提交")').first.click()
        log_fn("    等待 5 秒…")
        time.sleep(5)
        log_fn("    ✅ 證照單證 提交完成")
    except Exception as e:
        log_fn(f"    ❌ 提交失敗: {e}")


def _upload_customs_tab(page, xiaopiao: str, guanzheng: str, log_fn):
    """清關單證 tab: upload 小票+清關 PDF → 提交."""
    log_fn("  📦 Step 2: 清關單證")

    try:
        page.click('text=清關單證', timeout=8000)
        time.sleep(1)
    except Exception as e:
        log_fn(f"    ⚠️  找不到清關單證 tab: {e}")

    log_fn(f"    上傳小票 PDF: {os.path.basename(xiaopiao)}")
    try:
        page.locator('input[type="file"]').nth(0).set_input_files(xiaopiao)
        time.sleep(1)
    except Exception as e:
        log_fn(f"    ⚠️  上傳小票失敗: {e}")

    log_fn(f"    上傳清關 PDF: {os.path.basename(guanzheng)}")
    try:
        page.locator('input[type="file"]').nth(1).set_input_files(guanzheng)
        time.sleep(1)
    except Exception as e:
        log_fn(f"    ⚠️  上傳清關失敗: {e}")

    try:
        page.locator('button:has-text("提交")').first.click()
        log_fn("    等待 3 秒…")
        time.sleep(3)
        log_fn("    ✅ 清關單證 提交完成")
    except Exception as e:
        log_fn(f"    ❌ 提交失敗: {e}")


def _save_session_state(entries: list):
    """Write updated progress flags back to last_session.json."""
    with open(SESSION_FILE, 'w', encoding='utf-8') as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


def run_clearance(entries: list, dry_run=False, log_fn=print, force_redo=False):
    """Run clearance upload for all entries.

    Progress is saved to last_session.json after each step so re-runs skip
    already-completed steps (prevents duplicate submissions on SF website).
    Pass force_redo=True (or --force flag) to re-do all steps regardless.
    """
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=CHROME_PROFILE,
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
            no_viewport=True,
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        for idx, entry in enumerate(entries):
            customer  = entry.get('customer', '?')
            waybill   = entry.get('waybill', '')
            id_done   = entry.get('id_uploaded', False) and not force_redo
            cus_done  = entry.get('customs_uploaded', False) and not force_redo

            log_fn(f"\n[{idx+1}/{len(entries)}] {customer} — {waybill}")

            if id_done:
                log_fn("  [跳過] 証照單証 — 上次已完成")
            if cus_done:
                log_fn("  [跳過] 清關單証 — 上次已完成")
            if id_done and cus_done:
                log_fn("  ✅ 兩步均已完成，略過")
                continue

            front, back = find_id_cards(customer)
            xiaopiao, guanzheng = derive_pdf_paths(entry)

            # Validate files (only check what's actually needed)
            missing = []
            if not id_done:
                if not front or not os.path.exists(str(front)):
                    missing.append(f"身份証正面: {front}")
                if not back or not os.path.exists(str(back)):
                    missing.append(f"身份証背面: {back}")
            if not cus_done:
                if not guanzheng or not os.path.exists(str(guanzheng)):
                    missing.append(f"清關PDF: {guanzheng}")

            if missing:
                log_fn("  ⚠️  跳過（缺少文件）：")
                for m in missing: log_fn(f"    - {m}")
                continue

            if dry_run:
                if not id_done:
                    log_fn(f"  [試跑] 身份証正面: {front}")
                    log_fn(f"  [試跑] 身份証背面: {back}")
                if not cus_done:
                    log_fn(f"  [試跑] 清關PDF:   {guanzheng}")
                log_fn("  [試跑] 略過所有瀏覽器操作")
                continue

            try:
                # ── 導航到清關上傳頁，獲取 iframe ─────────────────────────
                frame = _get_iframe(page)

                # ── 填入運單號 ─────────────────────────────────────────────
                frame.locator('input.ant-input').first.fill(waybill)
                frame.locator('input.ant-input').first.press('Tab')
                time.sleep(1)
                log_fn(f"  運單號: {waybill}")

                # ── Step 1: 証照単証（若已完成則跳過）─────────────────────
                if not id_done:
                    _do_id_tab(page, frame, front, back, log_fn)
                    entry['id_uploaded'] = True
                    _save_session_state(entries)
                    log_fn("  [記錄] 証照單証完成 → last_session.json 已更新")

                # ── Step 2: 清關単証（若已完成則跳過）─────────────────────
                if not cus_done:
                    success = _do_customs_tab(page, frame, xiaopiao, guanzheng, log_fn)
                    if success:
                        entry['customs_uploaded'] = True
                        _save_session_state(entries)
                        log_fn("  [記錄] 清關單証完成 → last_session.json 已更新")

            except PWTimeout as e:
                log_fn(f"  ❌ 超時: {e}")
            except Exception as e:
                log_fn(f"  ❌ 錯誤: {e}")

        log_fn("\n🎉 全部完成！")
        ctx.close()


# ─── Entry point ──────────────────────────────────────────────────────────────
if __name__ == '__main__':
    _auto    = '--auto'   in sys.argv   # 無 GUI，直接跑
    _step2   = '--step2'  in sys.argv   # 只跑 Step 2（強制跳過 Step 1）
    _force   = '--force'  in sys.argv   # 忽略已完成標記，全部重跑

    if _auto or _step2:
        _mode = "Step2 only" if _step2 else "自動模式"
        print(f"[清關上傳] {_mode} — 讀取 {SESSION_FILE}")
        try:
            with open(SESSION_FILE, encoding='utf-8') as _f:
                _entries = json.load(_f)
        except Exception as e:
            print(f"[清關上傳] 讀取 last_session.json 失敗: {e}")
            sys.exit(1)
        if not _entries:
            print("[清關上傳] last_session.json 為空，略過")
            sys.exit(0)

        if _step2:
            # 強制 Step 1 標記為已完成，只跑 Step 2
            for _e in _entries:
                _e['id_uploaded'] = True

        print(f"[清關上傳] 共 {len(_entries)} 筆，開始上傳...")
        run_clearance(_entries, dry_run=False, log_fn=print, force_redo=_force)
    else:
        ClearanceApp().run()
