"""
ui.py — 訂單未交量報表產生器的圖形介面模組。

提供以下功能：
  show_startup_dialog()      — 主啟動視窗（路徑設定 + 執行 + 即時 Log）
  show_error(title, message) — 錯誤提示對話框
  show_info(title, message)  — 資訊提示對話框
  ask_continue_without_erp() — ERP 連線失敗時詢問是否繼續
"""

import os
import sys
import tkinter as tk
import logging
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

from credentials import COMPANIES, DEFAULT_COMPANY


# 全域變數：主視窗參考，供 show_error / show_info 等函式使用
_root = None


class TextHandler(logging.Handler):
    """將 logging 訊息即時導向 tkinter Text 元件的 Handler。"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append_text():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        self.text_widget.after(0, append_text)


# ---------------------------------------------------------------------------
# 主啟動視窗
# ---------------------------------------------------------------------------

def show_startup_dialog(default_company: str = "", default_excel: str = "",
                        default_plan_excel: str = "", default_output: str = "",
                        execute_callback=None):
    """
    開啟卡片式啟動視窗，包含路徑設定、執行按鈕及即時 Log 顯示。

    Parameters
    ----------
    default_company : str    上次選擇的公司代碼（從 paths.json 讀取）
    default_excel : str      訂單資訊管制表的預設路徑
    default_plan_excel : str 生產計畫表的預設路徑
    default_output : str     報表輸出的預設路徑
    execute_callback : callable  點擊「執行」後呼叫的回呼函式
    """
    result = {}

    global _root
    _root = TkinterDnD.Tk()
    root = _root
    root.title("訂單未交量報表產生器")
    root.configure(bg="#F4F6F7")

    # ── 設定視窗圖示 ──
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    # ── 視窗置中 ──
    w, h = 720, 680
    root.update_idletasks()
    x = (root.winfo_screenwidth() - w) // 2
    y = max(0, (root.winfo_screenheight() - h) // 2 - 20)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.minsize(620, 550)

    # ── 深色標題列 ──
    header = tk.Frame(root, bg="#34495E", height=50)
    header.pack(side="top", fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="訂單未交量報表產生器", bg="#34495E", fg="white",
             font=("Microsoft JhengHei", 16, "bold")).pack(expand=True)

    # ── 主內容區 ──
    main_frame = tk.Frame(root, bg="#F4F6F7", padx=20, pady=10)
    main_frame.pack(fill="both", expand=True)

    tk.Label(main_frame, text="路徑設定", bg="#F4F6F7", fg="#7F8C8D",
             font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w", pady=(0, 6))

    # ── 卡片元件：路徑選擇 ──
    def create_path_card(parent, tag_color, tag_text, title_text, var, browse_cmd):
        """建立一張路徑選擇卡片，支援拖拉檔案與瀏覽按鈕。"""
        card = tk.Frame(parent, bg="white", pady=6, padx=10)

        # 左側色條
        accent = tk.Frame(card, bg=tag_color, width=6)
        accent.pack(side="left", fill="y", padx=(0, 10))

        # 分類標籤
        tag_lbl = tk.Label(card, text=tag_text, bg=tag_color, fg="white",
                           font=("Microsoft JhengHei", 10, "bold"), width=8, pady=2)
        tag_lbl.pack(side="left", padx=(0, 15))

        # 文字區：標題 + 路徑顯示
        text_frame = tk.Frame(card, bg="white")
        text_frame.pack(side="left", fill="x", expand=True)

        title_lbl = tk.Label(text_frame, text=title_text, bg="white", fg="#2C3E50",
                             font=("Microsoft JhengHei", 11, "bold"), anchor="w")
        title_lbl.pack(side="top", fill="x")

        path_lbl = tk.Label(text_frame, textvariable=var, bg="white", fg="#95A5A6",
                            font=("Microsoft JhengHei", 9), anchor="w",
                            justify="left", wraplength=450)
        path_lbl.pack(side="top", fill="x", pady=(2, 0))

        # 瀏覽按鈕
        btn = tk.Button(card, text="瀏覽", bg="#E0E0E0", fg="#333333", relief="flat",
                        font=("Microsoft JhengHei", 10), padx=15, command=browse_cmd)
        btn.pack(side="right", padx=(10, 0))

        # 拖拉支援：卡片上所有元件都接受拖拉
        def on_drop(event):
            raw = event.data.strip()
            if raw.startswith("{") and raw.endswith("}"):
                raw = raw[1:-1]
            var.set(raw.replace("\\", "/"))

        for w in (card, text_frame, title_lbl, path_lbl, tag_lbl, accent):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", on_drop)

        return card

    # ── 卡片元件：下拉選單 ──
    def create_combo_card(parent, tag_color, tag_text, title_text, var, options):
        """建立一張下拉選單卡片（用於公司別選擇）。"""
        card = tk.Frame(parent, bg="white", pady=6, padx=10)
        accent = tk.Frame(card, bg=tag_color, width=6)
        accent.pack(side="left", fill="y", padx=(0, 10))
        tag_lbl = tk.Label(card, text=tag_text, bg=tag_color, fg="white",
                           font=("Microsoft JhengHei", 10, "bold"), width=8, pady=2)
        tag_lbl.pack(side="left", padx=(0, 15))

        text_frame = tk.Frame(card, bg="white")
        text_frame.pack(side="left", fill="x", expand=True)
        tk.Label(text_frame, text=title_text, bg="white", fg="#2C3E50",
                 font=("Microsoft JhengHei", 11, "bold"), anchor="w").pack(side="top", fill="x", pady=(0, 3))

        cb = ttk.Combobox(text_frame, textvariable=var, values=options,
                          state="readonly", width=40, font=("Microsoft JhengHei", 10))
        cb.pack(side="top", anchor="w")
        return card

    # ── 公司別選擇 ──
    company_options = [f"{k} - {v[0]}" for k, v in COMPANIES.items()]
    target_company = default_company if default_company else DEFAULT_COMPANY
    default_option = next((opt for opt in company_options if opt.startswith(target_company)),
                          company_options[0])
    company_var = tk.StringVar(value=default_option)
    create_combo_card(main_frame, "#9B59B6", "資料庫", "公司別（ERP 資料庫）",
                      company_var, company_options).pack(fill="x", pady=(0, 6))

    # ── 訂單資訊管制表路徑 ──
    excel_var = tk.StringVar(value=default_excel)
    def browse_excel():
        path = filedialog.askopenfilename(
            title="選擇 Excel 來源檔案",
            filetypes=[("Excel 檔案", "*.xlsx *.xls"), ("所有檔案", "*.*")],
            initialfile=excel_var.get())
        if path:
            excel_var.set(path)
    create_path_card(main_frame, "#2ECC71", "訂單資訊", "【訂單資訊管制表】路徑",
                     excel_var, browse_excel).pack(fill="x", pady=(0, 6))

    # ── 生產計劃表路徑 ──
    plan_var = tk.StringVar(value=default_plan_excel)
    def browse_plan():
        path = filedialog.askopenfilename(
            title="選擇生產計畫表",
            filetypes=[("Excel 活頁簿", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")],
            initialfile=plan_var.get())
        if path:
            plan_var.set(path)
    create_path_card(main_frame, "#F39C12", "生產計畫", "【生產計畫表】路徑",
                     plan_var, browse_plan).pack(fill="x", pady=(0, 6))

    # ── 輸出位置 ──
    output_var = tk.StringVar(value=default_output)
    def browse_output():
        initial = output_var.get()
        path = filedialog.asksaveasfilename(
            title="選擇報表儲存位置", defaultextension=".xlsx",
            filetypes=[("Excel 活頁簿", "*.xlsx"), ("所有檔案", "*.*")],
            initialdir=os.path.dirname(initial) if initial else "",
            initialfile=os.path.basename(initial) if initial else "")
        if path:
            output_var.set(path)
    create_path_card(main_frame, "#34495E", "輸出位置", "輸出 Excel 路徑",
                     output_var, browse_output).pack(fill="x", pady=(0, 10))

    # ── 操作列 ──
    tk.Frame(main_frame, bg="#E5E7E9", height=2).pack(fill="x", pady=(0, 6))

    action_frame = tk.Frame(main_frame, bg="#F4F6F7")
    action_frame.pack(fill="x")

    auto_open_var = tk.BooleanVar(value=True)
    tk.Checkbutton(action_frame, text="完成後自動開啟 Excel", variable=auto_open_var,
                   bg="#F4F6F7", font=("Microsoft JhengHei", 10),
                   activebackground="#F4F6F7").pack(side="left")

    # ── 執行記錄（即時 Log，置於最底部） ──
    tk.Label(main_frame, text="執行記錄", bg="#F4F6F7", fg="#7F8C8D",
             font=("Microsoft JhengHei", 10, "bold")).pack(anchor="w", pady=(6, 3))
    log_text = tk.Text(main_frame, bg="#1E1E1E", fg="#D4D4D4",
                       font=("Consolas", 9), height=8, state='disabled')
    log_text.pack(fill="both", expand=True, pady=(0, 6))

    # 將 logging 輸出導向 Text 元件
    text_handler = TextHandler(log_text)
    text_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logging.getLogger().addHandler(text_handler)

    def on_ok():
        """驗證必填欄位後執行回呼或返回結果。"""
        code = company_var.get().split(" - ")[0].strip()
        path = excel_var.get().strip()
        plan_path = plan_var.get().strip()
        out_path = output_var.get().strip()
        if not path:
            messagebox.showerror("錯誤", "請選擇來源 Excel 檔案路徑。")
            return
        if not out_path:
            messagebox.showerror("錯誤", "請選擇報表儲存路徑。")
            return
        if not execute_callback:
            result["company"] = code
            result["excel"] = path
            result["plan_excel"] = plan_path
            result["output"] = out_path
            result["auto_open"] = auto_open_var.get()
            root.destroy()
        else:
            # 執行期間禁用按鈕、清空 Log
            btn_exec.configure(state="disabled")
            log_text.configure(state='normal')
            log_text.delete(1.0, tk.END)
            log_text.configure(state='disabled')
            
            def run_task():
                try:
                    execute_callback(code, path, plan_path, out_path, auto_open_var.get())
                finally:
                    # 恢復按鈕狀態（確保在 Main Thread 執行）
                    root.after(0, lambda: btn_exec.configure(state="normal"))
                    
            import threading
            threading.Thread(target=run_task, daemon=True).start()

    # 執行按鈕（含 Hover 效果）
    btn_exec = tk.Button(action_frame, text="▶ 執行", bg="#34495E", fg="white",
                         font=("Microsoft JhengHei", 11, "bold"), relief="flat",
                         padx=25, pady=8, command=on_ok)
    btn_exec.pack(side="right")
    btn_exec.bind("<Enter>", lambda e: e.widget.configure(bg="#2C3E50")
                  if e.widget['state'] != 'disabled' else None)
    btn_exec.bind("<Leave>", lambda e: e.widget.configure(bg="#34495E")
                  if e.widget['state'] != 'disabled' else None)

    def on_close():
        """關閉視窗時移除 Log Handler 並清理全域參考。"""
        logging.getLogger().removeHandler(text_handler)
        global _root
        _root = None
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()

    # 無回呼模式：回傳結果（保留向下相容）
    if not execute_callback:
        if not result:
            sys.exit(0)
        return (result["company"], result["excel"], result["plan_excel"],
                result["output"], result["auto_open"])


# ---------------------------------------------------------------------------
# 簡易對話框
# ---------------------------------------------------------------------------

def show_error(title: str, message: str):
    """顯示錯誤提示對話框。若主視窗存在則附屬於主視窗。"""
    if _root:
        _root.after(0, lambda: messagebox.showerror(title, message, master=_root))
    else:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror(title, message, master=root)
        root.destroy()


def show_info(title: str, message: str):
    """顯示資訊提示對話框。若主視窗存在則附屬於主視窗。"""
    if _root:
        _root.after(0, lambda: messagebox.showinfo(title, message, master=_root))
    else:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(title, message, master=root)
        root.destroy()


def ask_continue_without_erp() -> bool:
    """ERP 連線失敗時詢問使用者是否以庫存量=0 繼續產生報表。"""
    if _root:
        import threading
        ans = [False]
        ev = threading.Event()
        def _ask():
            ans[0] = messagebox.askyesno(
                "ERP 連線失敗",
                "無法連線至 ERP 資料庫。\n\n"
                "是否以庫存量 = 0 繼續產生報表？\n"
                "（選「否」將結束程式）",
                master=_root)
            ev.set()
        _root.after(0, _ask)
        ev.wait()
        return ans[0]
    else:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        answer = messagebox.askyesno(
            "ERP 連線失敗",
            "無法連線至 ERP 資料庫。\n\n"
            "是否以庫存量 = 0 繼續產生報表？\n"
            "（選「否」將結束程式）",
            master=root)
        root.destroy()
        return answer
