# -*- coding: utf-8 -*-
"""
順丰寄件自動化系統 — Main GUI
Tab1: 新寄件 (10-row batch input)  |  Tab2: 狀態追蹤
"""
import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk

from config import EXCEL_PATH, IMAGES_DIR
from excel_manager import ExcelManager
from logger import get_logger, toast_ok, toast_error
from order_parser import parse_order

log = get_logger(__name__)

_NUM_ROWS   = 10
_STEP_COUNT = 6   # steps per order

# ── ID photo finder ────────────────────────────────────────────────────────────

def _find_id_photos(id_dir: str) -> tuple[str, str]:
    """Return (front_path, back_path) from id_dir. Returns ('','') if missing."""
    if not os.path.isdir(id_dir):
        return "", ""
    front = back = ""
    for fname in os.listdir(id_dir):
        low = fname.lower()
        if not any(low.endswith(ext) for ext in (".jpg", ".jpeg", ".png")):
            continue
        if any(k in low for k in ("front", "正面", "正")):
            front = os.path.join(id_dir, fname)
        elif any(k in low for k in ("back", "背面", "背")):
            back = os.path.join(id_dir, fname)
    return front, back


def _resolve_id_dir(name: str, simplified: str) -> str:
    """Return the ID folder path, trying original name then simplified."""
    for n in (name, simplified):
        d = os.path.join(IMAGES_DIR, n)
        if os.path.isdir(d):
            return d
    return os.path.join(IMAGES_DIR, name)


# ── Main app ───────────────────────────────────────────────────────────────────

class AutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("順丰寄件自動化 v1.0")
        self.geometry("900x780")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")

        self.excel          = ExcelManager()
        self._parsed_orders: list[dict] = []   # validated orders ready to run
        self._row_vars:  list[tk.StringVar] = []
        self._row_status: list[tk.StringVar] = []

        self._build_ui()
        log.info("GUI started")

    # ── Build UI ──────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        tab1 = ttk.Frame(nb)
        tab2 = ttk.Frame(nb)
        nb.add(tab1, text="📦  新寄件")
        nb.add(tab2, text="📊  狀態追蹤")

        self._build_order_tab(tab1)
        self._build_status_tab(tab2)

    # ── Tab 1 ─────────────────────────────────────────────────────────────────

    def _build_order_tab(self, frame: ttk.Frame) -> None:
        # ── 10-row input grid ──────────────────────────────────────────────────
        input_lf = tk.LabelFrame(
            frame, text="WhatsApp 訂單 (每行一個客人，貼入後按解析)",
            padx=6, pady=6,
        )
        input_lf.pack(fill="x", padx=10, pady=(10, 4))

        for i in range(_NUM_ROWS):
            sv  = tk.StringVar()
            row_st = tk.StringVar(value="")
            self._row_vars.append(sv)
            self._row_status.append(row_st)

            row_frame = tk.Frame(input_lf)
            row_frame.pack(fill="x", pady=1)

            tk.Label(row_frame, text=f"{i+1:2d}.", width=3,
                     anchor="e", font=("Courier", 10)).pack(side="left")
            tk.Entry(row_frame, textvariable=sv,
                     font=("", 10), relief="solid", bd=1).pack(
                side="left", fill="x", expand=True, padx=(4, 4))
            tk.Label(row_frame, textvariable=row_st,
                     width=10, anchor="w", font=("", 9)).pack(side="left")

        btn_row = tk.Frame(input_lf)
        btn_row.pack(fill="x", pady=(6, 0))
        tk.Button(btn_row, text="🗑 清除全部", command=self._clear_all,
                  bg="#95a5a6", fg="white", padx=8, pady=3).pack(side="left")
        tk.Button(btn_row, text="🔍 解析全部訂單", command=self.on_parse_all,
                  bg="#2980b9", fg="white", padx=12, pady=4,
                  font=("", 11)).pack(side="right")

        # ── Parsed results table ───────────────────────────────────────────────
        result_lf = tk.LabelFrame(frame, text="已解析訂單", padx=6, pady=4)
        result_lf.pack(fill="both", expand=True, padx=10, pady=4)

        cols = ("#", "收件人", "電話", "貨品", "VIP總額", "身份證", "狀態")
        self.order_tree = ttk.Treeview(result_lf, columns=cols,
                                       show="headings", height=8)
        self.order_tree.heading("#", text="#")
        self.order_tree.column("#",      width=28,  anchor="center")
        self.order_tree.heading("收件人", text="收件人")
        self.order_tree.column("收件人",  width=90)
        self.order_tree.heading("電話",   text="電話")
        self.order_tree.column("電話",    width=115)
        self.order_tree.heading("貨品",   text="貨品")
        self.order_tree.column("貨品",    width=230)
        self.order_tree.heading("VIP總額", text="VIP總額")
        self.order_tree.column("VIP總額", width=75,  anchor="center")
        self.order_tree.heading("身份證",  text="身份證")
        self.order_tree.column("身份證",   width=60,  anchor="center")
        self.order_tree.heading("狀態",    text="狀態")
        self.order_tree.column("狀態",     width=90,  anchor="center")
        self.order_tree.pack(fill="both", expand=True)

        self.order_tree.tag_configure("ok",      foreground="#27ae60")
        self.order_tree.tag_configure("error",   foreground="#e74c3c")
        self.order_tree.tag_configure("running", foreground="#e67e22")
        self.order_tree.tag_configure("done",    foreground="#2980b9")

        # ── Progress ───────────────────────────────────────────────────────────
        prog_lf = tk.LabelFrame(frame, text="進度", padx=6, pady=4)
        prog_lf.pack(fill="x", padx=10, pady=4)

        self.progress_bar = ttk.Progressbar(prog_lf, length=700, maximum=100)
        self.progress_bar.pack(fill="x")
        self.status_label = tk.Label(prog_lf, text="就緒", anchor="w", fg="#555")
        self.status_label.pack(anchor="w", pady=2)

        # ── Run button ─────────────────────────────────────────────────────────
        self.run_btn = tk.Button(
            frame, text="🚀 開始批量寄件", command=self.on_run_all,
            bg="#27ae60", fg="white", font=("", 13, "bold"),
            padx=20, pady=10, state="disabled",
        )
        self.run_btn.pack(pady=(4, 10))

    # ── Tab 2 ─────────────────────────────────────────────────────────────────

    def _build_status_tab(self, frame: ttk.Frame) -> None:
        btn_frame = tk.Frame(frame)
        btn_frame.pack(fill="x", padx=10, pady=8)

        tk.Button(btn_frame, text="🔄 更新所有運單狀態",
                  command=self.on_refresh_status,
                  bg="#8e44ad", fg="white", padx=10, pady=6).pack(side="left", padx=4)
        tk.Button(btn_frame, text="📂 開啟 Excel",
                  command=self.on_open_excel,
                  bg="#2c3e50", fg="white", padx=10, pady=6).pack(side="left", padx=4)
        tk.Button(btn_frame, text="📦 更新產品資料庫",
                  command=self.on_update_products,
                  bg="#d35400", fg="white", padx=10, pady=6).pack(side="left", padx=4)

        cols = ("日期", "客人名", "運單號", "貨品摘要", "件數", "收件地址", "最新狀態", "更新時間", "備註")
        self.track_tree = ttk.Treeview(frame, columns=cols, show="headings", height=22)
        for c in cols:
            self.track_tree.heading(c, text=c)
        self.track_tree.column("日期",    width=90)
        self.track_tree.column("客人名",  width=80)
        self.track_tree.column("運單號",  width=135)
        self.track_tree.column("貨品摘要",width=200)
        self.track_tree.column("件數",    width=40,  anchor="center")
        self.track_tree.column("收件地址",width=180)
        self.track_tree.column("最新狀態",width=85)
        self.track_tree.column("更新時間",width=130)
        self.track_tree.column("備註",    width=100)
        self.track_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.track_tree.yview)
        self.track_tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")

        self._reload_treeview()

    # ── Event handlers ────────────────────────────────────────────────────────

    def _clear_all(self) -> None:
        for sv in self._row_vars:
            sv.set("")
        for st in self._row_status:
            st.set("")
        self._parsed_orders.clear()
        for item in self.order_tree.get_children():
            self.order_tree.delete(item)
        self.run_btn.config(state="disabled")
        self._set_status("就緒")

    def on_parse_all(self) -> None:
        self._parsed_orders.clear()
        for item in self.order_tree.get_children():
            self.order_tree.delete(item)

        ok_count = 0
        for i, sv in enumerate(self._row_vars):
            raw = sv.get().strip()
            if not raw:
                self._row_status[i].set("")
                continue
            try:
                order = parse_order(raw)

                # Check ID photos
                name       = order["name"]
                simplified = order.get("name_simplified", name)
                id_dir     = _resolve_id_dir(name, simplified)
                front, back = _find_id_photos(id_dir)
                order["id_front"] = front
                order["id_back"]  = back
                id_ok = "✅" if (front and back) else "❌缺"

                items_str = ", ".join(
                    f"{it.get('name', it['sku'])}×{it['qty']}" for it in order["items"])
                order["_row_idx"] = i
                order["_tree_id"] = None
                self._parsed_orders.append(order)

                iid = self.order_tree.insert(
                    "", "end",
                    values=(i + 1, name, order["phone"],
                            items_str, f"HKD {order['total']:.0f}",
                            id_ok, "待寄出"),
                    tags=("ok",),
                )
                order["_tree_id"] = iid
                self._row_status[i].set("✅")
                ok_count += 1

            except ValueError as e:
                self._row_status[i].set("❌")
                self.order_tree.insert(
                    "", "end",
                    values=(i + 1, "—", "—", str(e)[:60], "—", "—", "解析失敗"),
                    tags=("error",),
                )

        if ok_count:
            self.run_btn.config(state="normal")
            self._set_status(f"✅ 解析完成：{ok_count} 個訂單準備好，按「開始批量寄件」")
        else:
            self.run_btn.config(state="disabled")
            self._set_status("❌ 沒有成功解析的訂單")

    def on_run_all(self) -> None:
        missing = [o["name"] for o in self._parsed_orders
                   if not o.get("id_front") or not o.get("id_back")]
        if missing:
            log.warning("ID photos missing for: %s — proceeding anyway", missing)

        self.run_btn.config(state="disabled")
        self.progress_bar["value"] = 0
        total = len(self._parsed_orders)
        self._set_status(f"⏳ 開始批量寄件，共 {total} 個訂單…")

        threading.Thread(
            target=self._batch_worker,
            args=(list(self._parsed_orders),),
            daemon=True,
        ).start()

    def on_refresh_status(self) -> None:
        self._set_status("⏳ 正在更新運單狀態…")
        threading.Thread(target=self._refresh_status_worker, daemon=True).start()

    def on_open_excel(self) -> None:
        if os.path.exists(EXCEL_PATH):
            os.startfile(EXCEL_PATH)
        else:
            messagebox.showinfo("未找到", f"Excel 不存在：{EXCEL_PATH}")

    def on_update_products(self) -> None:
        self._set_status("⏳ 正在更新產品資料庫…")
        def _worker():
            try:
                from product_scraper import scrape_products
                result = scrape_products(force_refresh=True)
                self.after(0, lambda: self._set_status(
                    f"✅ 產品資料庫更新完成 ({len(result)} 件產品)"))
            except Exception as e:
                self.after(0, lambda: self._set_status(f"❌ 更新失敗: {e}"))
        threading.Thread(target=_worker, daemon=True).start()

    # ── Batch worker ──────────────────────────────────────────────────────────

    def _batch_worker(self, orders: list) -> None:
        total   = len(orders)
        success = 0
        failed  = 0

        for idx, order in enumerate(orders):
            name    = order["name"]
            tree_id = order.get("_tree_id")

            def _set_row(tid, status_text, tag):
                if tid:
                    vals = list(self.order_tree.item(tid, "values"))
                    vals[6] = status_text
                    self.order_tree.item(tid, values=vals, tags=(tag,))

            self.after(0, lambda tid=tree_id: _set_row(tid, "⏳進行中", "running"))
            self.after(0, lambda i=idx, n=name: self._set_status(
                f"⏳ [{i+1}/{total}] {n} — POS 下單中…"))

            try:
                # POS
                from pos_automation import run_pos_checkout
                pos_order_no, pdf_path = run_pos_checkout(order)

                # Excel placeholder row
                excel_row = self.excel.append_order(
                    order, pos_order_no=pos_order_no, pdf_path=pdf_path)

                # SF submission
                self.after(0, lambda i=idx, n=name: self._set_status(
                    f"⏳ [{i+1}/{total}] {n} — 順丰填單中…"))
                from sf_automation import run_sf_submission, SubmissionCancelledError
                try:
                    waybill = run_sf_submission(order, pdf_path)
                except SubmissionCancelledError:
                    self.after(0, lambda tid=tree_id: _set_row(tid, "⛔已取消", "error"))
                    failed += 1
                    continue

                self.excel.update_waybill(excel_row, waybill)
                toast_ok(f"{name} 完成，運單號：{waybill}")
                self.after(0, lambda tid=tree_id, w=waybill:
                           _set_row(tid, f"✅{w}", "done"))
                success += 1

            except Exception as e:
                err = str(e)[:80]
                log.exception("Batch order failed: %s", name)
                toast_error(f"{name} 失敗", err)
                self.after(0, lambda tid=tree_id: _set_row(tid, "❌失敗", "error"))
                failed += 1
            finally:
                from browser_utils import close_all
                close_all()

            # Update overall progress
            pct = int((idx + 1) / total * 100)
            self.after(0, lambda p=pct: self.progress_bar.config(value=p))

        self.after(200, self._reload_treeview)
        summary = f"✅ 批量完成：{success} 成功，{failed} 失敗（共 {total} 個）"
        self.after(0, lambda: self._set_status(summary))
        self.after(0, lambda: self.run_btn.config(state="normal"))
        log.info(summary)

    def _refresh_status_worker(self) -> None:
        try:
            from status_updater import update_all_statuses
            results = update_all_statuses(self.excel)
            self.after(0, self._reload_treeview)
            self.after(0, lambda: self._set_status(
                f"✅ 已更新 {len(results)} 個運單狀態"))
        except Exception as e:
            self.after(0, lambda: self._set_status(f"❌ 更新失敗: {e}"))

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _set_status(self, text: str) -> None:
        self.status_label.config(text=text)

    def _reload_treeview(self) -> None:
        for row in self.track_tree.get_children():
            self.track_tree.delete(row)
        from config import (COL_DATE, COL_NAME, COL_WAYBILL, COL_ITEMS,
                             COL_QTY, COL_ADDRESS, COL_STATUS,
                             COL_STATUS_TIME, COL_NOTES)
        for r in self.excel.get_recent_rows(n=100):
            self.track_tree.insert("", "end", values=(
                r[COL_DATE - 1]        or "",
                r[COL_NAME - 1]        or "",
                r[COL_WAYBILL - 1]     or "",
                r[COL_ITEMS - 1]       or "",
                r[COL_QTY - 1]         or "",
                r[COL_ADDRESS - 1]     or "",
                r[COL_STATUS - 1]      or "",
                r[COL_STATUS_TIME - 1] or "",
                r[COL_NOTES - 1]       or "",
            ))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()
