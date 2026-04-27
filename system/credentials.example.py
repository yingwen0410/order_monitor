# credentials.example.py — ERP 資料庫連線資訊範例
# 請將此檔案複製並更名為 credentials.py，填寫實際的連線資訊

# ERP SQL Server 連線設定
DB_CONFIG = {
    "server":   "192.168.1.X",
    "username": "your_username",
    "password": "your_password",
    "driver":   "ODBC Driver 17 for SQL Server",
    "warehouse": "11A1",        # 預設查詢庫別
}

# 公司代碼 → (顯示名稱, SQL Server 資料庫名稱)
COMPANIES = {
    "SFT": ("範例公司 A", "DB_A"),
    "GTE": ("範例公司 B", "DB_B"),
}

# 啟動時的預設公司
DEFAULT_COMPANY = "SFT"
