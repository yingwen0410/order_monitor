"""
main.py — 訂單未交量報表產生器的主程式進入點。

執行流程：
  1. 載入 config.ini 設定檔
  2. 載入 paths.json 上次記錄的路徑
  3. 開啟啟動視窗（選擇公司別、來源 Excel、生產記錄表、輸出位置）
  4. 讀取來源 Excel（唯讀模式，不修改原檔）
  5. 正規化 DataFrame（欄位重命名、日期解析、清洗換行字元）
  6. 連線 ERP 查詢庫存量
  7. 將庫存量對照至 DataFrame
  8. 載入允備貨清單（取得最大安基量）
  9. 產出三頁式報表（過期 / 未過期 / 客戶總表）
 10. 完成提示並自動開啟報表（可選）
"""

import logging
import os
import sys
from datetime import date

import utils
import reader
import erp
import writer
import ui


def _load_allow_lookup(path: str) -> dict:
    """
    載入生產計畫表中「允備貨清單」工作表。
    建立 (客戶, 品號) → 最大安基量 的對照表，用於客戶總表。

    注意：來源 Excel 使用自訂數字格式 0_ "R" 讓數值後面顯示 R，
    但 pandas 讀取時只拿到純數字。因此改用 openpyxl 直接讀取，
    同時檢查 number_format 來還原 R 後綴。
    """
    import win32com.client

    if not path or not os.path.exists(path):
        logging.warning("未提供生產計畫表或檔案不存在 (%s)，最大安基量欄位將留空", path)
        return {}

    logging.info("載入允備貨清單：%s", path)
    try:
        # 使用 Excel COM 介面直接讀取「顯示文字」，
        # 這樣不管儲存格用什麼格式（"R"、"PCS"、"M"…），都能完整還原。
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)

        ws = None
        for s in wb.Sheets:
            if s.Name == "允備貨清單":
                ws = s
                break
        
        if ws is None:
            logging.warning("找不到「允備貨清單」工作表")
            wb.Close(False)
            excel.Quit()
            return {}

        lookup = {}
        # row 1 = 更新說明, row 2 = 標題列, row 3 以後才是資料
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
        for r in range(3, last_row + 1):
            customer_raw = ws.Cells(r, 1).Value
            item_no_raw  = ws.Cells(r, 2).Value
            if customer_raw is None or item_no_raw is None:
                continue

            customer = str(customer_raw).replace("\n", " ").strip()
            item_no  = str(item_no_raw).strip()

            # .Text 回傳的就是 Excel 畫面上顯示的文字，包含格式後綴
            display_text = ws.Cells(r, 7).Text
            if display_text is None or str(display_text).strip() == "":
                continue

            lookup[(customer, item_no)] = str(display_text).strip()

        wb.Close(False)
        excel.Quit()
        logging.info("允備貨清單：載入 %d 筆配對", len(lookup))
        return lookup
    except Exception as e:
        logging.warning("無法載入允備貨清單：%s", e)
        # 確保發生異常時也能關閉 Excel
        try:
            if 'wb' in locals() and wb is not None:
                wb.Close(False)
            if 'excel' in locals() and excel is not None:
                excel.Quit()
        except Exception:
            pass
        return {}


def main():
    # ── 初始化 Logging ──
    utils.setup_logging()

    # ── 1. 載入 config.ini ──
    try:
        cfg = utils.load_config()
    except FileNotFoundError as e:
        ui.show_error("設定檔錯誤", str(e))
        sys.exit(1)

    # 從 config.ini 讀取預設值（首次安裝時使用）
    default_excel = cfg.get("paths", "default_excel", fallback="")
    default_plan  = cfg.get("paths", "default_plan_excel", fallback="")
    source_sheet  = cfg.get("paths", "source_sheet",  fallback="待出貨-膜類")
    skip_zero     = cfg.getboolean("options", "skip_zero_undelivered", fallback=True)

    # ── 2. 載入 paths.json（覆蓋 config.ini 的預設值） ──
    saved_paths = utils.load_saved_paths()
    default_company = saved_paths.get("company", "")
    default_excel   = saved_paths.get("excel", default_excel)
    default_plan    = saved_paths.get("plan_excel", default_plan)

    # 輸出路徑：自動帶入今天日期的檔案名稱
    today = date.today()
    default_name = f"訂單狀態報表_{today:%Y%m%d}.xlsx"

    saved_output_dir = saved_paths.get("output_dir", "")
    if saved_output_dir and os.path.isdir(saved_output_dir):
        default_output = os.path.join(saved_output_dir, default_name).replace("\\", "/")
    else:
        default_output = os.path.join(os.path.expanduser("~"), "Desktop", default_name).replace("\\", "/")

    # ── 3. 定義執行回呼（使用者點擊「執行」後觸發） ──
    def execute_callback(company_code, excel_path, plan_path, output_path, auto_open):
        # 若使用者選擇的是資料夾，自動補上檔案名稱
        if os.path.isdir(output_path):
            output_path = os.path.join(output_path, default_name).replace("\\", "/")

        # 儲存本次選擇的路徑，下次開啟時自動帶入
        utils.save_paths({
            "company": company_code,
            "excel": excel_path,
            "plan_excel": plan_path,
            "output_dir": os.path.dirname(output_path)
        })

        logging.info("公司別：%s，來源：%s，計畫表：%s，輸出：%s，自動開啟：%s",
                      company_code, excel_path, plan_path, output_path, auto_open)

        # ── 4. 讀取來源 Excel ──
        try:
            raw_df = reader.read_source(excel_path, source_sheet)
        except FileNotFoundError:
            ui.show_error("找不到來源檔案",
                          f"無法存取以下路徑，請確認網路連線或重新選擇檔案：\n{excel_path}")
            return
        except PermissionError:
            ui.show_error("檔案被占用", "來源 Excel 檔案目前被其他程式占用，請稍後再試。")
            return
        except ValueError as e:
            ui.show_error("Excel 格式錯誤", str(e))
            return
        except Exception as e:
            logging.exception("讀取 Excel 失敗")
            ui.show_error("讀取 Excel 失敗", f"發生未知錯誤：\n{e}")
            return

        # ── 5. 正規化 DataFrame ──
        try:
            df = utils.normalize(raw_df)
        except Exception as e:
            logging.exception("資料正規化失敗")
            ui.show_error("資料處理錯誤", str(e))
            return

        if skip_zero:
            before = len(df)
            df = utils.filter_zero_undelivered(df)
            logging.info("略過未交量=0：%d → %d 列", before, len(df))

        if df.empty:
            ui.show_info("無資料", "過濾後沒有任何有效訂單資料，請確認來源檔案內容。")
            return

        # ── 6. 連線 ERP 查詢庫存 ──
        inventory: dict = {}
        try:
            inventory = erp.fetch_inventory(company_code)
        except Exception as e:
            logging.warning("ERP 連線失敗：%s", e)
            if not ui.ask_continue_without_erp():
                return
            logging.info("以庫存量=0 繼續執行")

        # ── 7. 將庫存量對照至品號 ──
        df[utils.COL_STOCK] = df[utils.COL_PART_NO].map(inventory)

        unmatched = df[df[utils.COL_STOCK].isna()][utils.COL_PART_NO].unique()
        if len(unmatched) > 0:
            logging.warning("以下 %d 個品號在 ERP 庫別 %s 中無庫存紀錄（顯示為 N/A）：%s",
                            len(unmatched), "11A1", list(unmatched))

        # ── 8. 載入允備貨清單並產出報表 ──
        allow_lookup = _load_allow_lookup(plan_path)
        try:
            writer.write_report(df, today, output_path, allow_lookup)
        except PermissionError:
            ui.show_error("輸出失敗",
                          f"無法儲存報表，請確認輸出檔案未被 Excel 開啟：\n{output_path}")
            return
        except Exception as e:
            logging.exception("寫出報表失敗")
            ui.show_error("報表產生失敗", f"發生未知錯誤：\n{e}")
            return

        # ── 9. 完成 ──
        ui.show_info("完成", f"報表已儲存至：\n{output_path}")
        logging.info("執行完成")

        if auto_open:
            try:
                os.startfile(output_path)
            except Exception as e:
                logging.warning("無法自動開啟報表：%s", e)

    # ── 開啟啟動視窗，等待使用者操作 ──
    ui.show_startup_dialog(default_company, default_excel, default_plan,
                           default_output, execute_callback)


if __name__ == "__main__":
    main()
