"""
erp.py — 連線 ERP（SQL Server）查詢庫存量。

透過 pyodbc 連接 dbo.INVMC 資料表，
依指定庫別查詢所有品號的現有庫存量。
"""

import logging

import pyodbc

from credentials import DB_CONFIG, COMPANIES


# 查詢 SQL：MC001=品號, MC002=庫別, MC007=庫存量
_QUERY = """
SELECT MC001, MC007
FROM dbo.INVMC
WHERE MC002 = ?
"""


def fetch_inventory(company_code: str) -> dict:
    """
    查詢指定公司的 ERP 庫存資料。

    Parameters
    ----------
    company_code : str
        公司代碼（如 SFT），用來查找對應的 SQL Server 資料庫名稱。

    Returns
    -------
    dict  { 品號 (str): 庫存量 (float) }
    """
    c         = DB_CONFIG
    host      = c["server"]
    port      = int(c.get("port", 1433))
    warehouse = c["warehouse"]                    # 庫別（如 11A1）
    db_name   = COMPANIES[company_code][1]        # (顯示名稱, DB名稱)[1]

    logging.info("連線 ERP（公司別：%s，資料庫：%s，庫別：%s）…",
                 company_code, db_name, warehouse)

    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={host},{port};"
        f"DATABASE={db_name};"
        f"UID={c['username']};"
        f"PWD={c['password']};"
        f"Encrypt=yes;"
        f"TrustServerCertificate=yes;"
    )
    conn = pyodbc.connect(conn_str, timeout=10)
    try:
        cursor = conn.cursor()
        cursor.execute(_QUERY, warehouse)
        rows = cursor.fetchall()
        # 品號去除前後空白，庫存量轉為浮點數（None 視為 0）
        inventory = {str(row[0]).strip(): float(row[1] or 0) for row in rows}
        logging.info("ERP 查詢完成，共 %d 筆品號庫存", len(inventory))
        return inventory
    finally:
        conn.close()
