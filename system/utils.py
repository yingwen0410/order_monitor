"""
utils.py — 共用工具函式與常數。

包含：
  - config.ini / paths.json 的讀寫
  - logging 初始化
  - DataFrame 正規化（欄位重命名、日期解析、資料清洗）
  - 內部欄位名稱常數（COL_*）
"""

import os
import sys
import logging
import configparser
import json
from datetime import date

import pandas as pd


# ---------------------------------------------------------------------------
# 設定檔（config.ini）
# ---------------------------------------------------------------------------

def _base_dir() -> str:
    """取得程式所在的根目錄（打包後為 .exe 所在目錄，開發時為腳本目錄）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_config_path() -> str:
    """取得 config.ini 的完整路徑。"""
    return os.path.join(_base_dir(), "config.ini")


def load_config() -> configparser.ConfigParser:
    """載入 config.ini 並回傳 ConfigParser 物件。找不到檔案時拋出 FileNotFoundError。"""
    cfg = configparser.ConfigParser()
    path = get_config_path()
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到設定檔：{path}")
    cfg.read(path, encoding="utf-8")
    return cfg


# ---------------------------------------------------------------------------
# 路徑記憶檔（paths.json）— 記錄使用者上次選擇的路徑
# ---------------------------------------------------------------------------

def _local_data_dir() -> str:
    """取得本機資料夾路徑，用於儲存個人設定與 Log，避免 NAS 共用衝突。"""
    local_dir = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "OrderMonitor")
    os.makedirs(local_dir, exist_ok=True)
    return local_dir


def get_paths_json_path() -> str:
    """取得 paths.json 的完整路徑（移至本機資料夾）。"""
    return os.path.join(_local_data_dir(), "paths.json")


def load_saved_paths() -> dict:
    """讀取 paths.json，回傳路徑字典。檔案不存在或讀取失敗時回傳空字典。"""
    path = get_paths_json_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.warning("無法讀取 paths.json：%s", e)
    return {}


def save_paths(paths_dict: dict):
    """將路徑字典寫入 paths.json，供下次開啟時自動帶入。"""
    path = get_paths_json_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(paths_dict, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logging.warning("無法儲存 paths.json：%s", e)


# ---------------------------------------------------------------------------
# Logging 初始化
# ---------------------------------------------------------------------------

def setup_logging():
    """初始化 logging：同時輸出至本機 order_monitor.log 檔案與 stdout。"""
    log_path = os.path.join(_local_data_dir(), "order_monitor.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


# ---------------------------------------------------------------------------
# DataFrame 欄位常數（內部統一名稱）
# ---------------------------------------------------------------------------

COL_CUSTOMER     = "customer"         # 客戶名稱
COL_PRODUCT_NAME = "product_name"     # 品名規格
COL_PART_NO      = "part_no"          # 品號
COL_DELIVERY     = "delivery_date"    # 客戶交期
COL_UNDELIVERED  = "undelivered_qty"  # 未交量
COL_STOCK        = "stock_qty"        # 庫存量（來自 ERP）
COL_STATUS       = "status"           # 狀態（過期 / 未過期）
COL_SURPLUS      = "stock_surplus"    # 庫存差額 = 庫存量 - 未交量


# ---------------------------------------------------------------------------
# DataFrame 正規化
# ---------------------------------------------------------------------------

def normalize(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    將來源 Excel 的原始欄位名稱轉換為內部統一名稱，
    並進行日期解析、數值轉換、換行字元清洗。
    """
    # 原始 Excel 欄位 → 內部名稱 的對應表
    rename_map = {
        "Customer":    COL_CUSTOMER,
        "品名 \n規格": COL_PRODUCT_NAME,
        "品號":        COL_PART_NO,
        "客戶交期":    COL_DELIVERY,
        "未交量":      COL_UNDELIVERED,
    }

    # 僅保留存在的欄位
    available = {k: v for k, v in rename_map.items() if k in raw_df.columns}
    missing = set(rename_map.keys()) - set(available.keys())
    if missing:
        logging.warning("以下欄位在 Excel 中未找到，將嘗試繼續：%s", missing)

    df = raw_df[list(available.keys())].copy()
    df.rename(columns=available, inplace=True)

    # 清洗品名與客戶名稱中的換行字元
    for col in [COL_PRODUCT_NAME, COL_CUSTOMER]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(r"\n", " ", regex=True)
                .str.strip()
            )

    # 解析客戶交期為日期格式（無法辨識的值變成 NaT 並移除）
    if COL_DELIVERY in df.columns:
        df[COL_DELIVERY] = pd.to_datetime(df[COL_DELIVERY], errors="coerce")
        nat_count = df[COL_DELIVERY].isna().sum()
        if nat_count > 0:
            logging.warning("客戶交期欄位有 %d 筆無效日期，已跳過", nat_count)
        df.dropna(subset=[COL_DELIVERY], inplace=True)

    # 將未交量轉為數值（無法轉換的值填 0）
    if COL_UNDELIVERED in df.columns:
        df[COL_UNDELIVERED] = pd.to_numeric(df[COL_UNDELIVERED], errors="coerce").fillna(0)

    df.reset_index(drop=True, inplace=True)
    return df


def filter_zero_undelivered(df: pd.DataFrame) -> pd.DataFrame:
    """過濾掉未交量為 0 的列。"""
    return df[df[COL_UNDELIVERED] > 0].copy()


def add_status_column(df: pd.DataFrame, today: date) -> pd.DataFrame:
    """依據客戶交期與今日日期，新增「狀態」欄（過期 / 未過期）。"""
    df = df.copy()
    df[COL_STATUS] = df[COL_DELIVERY].apply(
        lambda d: "過期" if d.date() <= today else "未過期"
    )
    return df
