# Order Monitor — 訂單未交量報表產生器

讀取 NAS 上的訂單管制表並查詢 ERP 庫存，自動產出四頁式 Excel 報表，供倉管人員掌握未交訂單狀態。

---

## 功能特色

- 以唯讀模式讀取 NAS 上的來源 Excel，不修改原始檔案
- 連線 ERP SQL Server 查詢即時庫存
- 依時間序列遞減分配庫存，跨客戶公平扣減
- 產出報表含四個工作表：過期未交 / 未過期未交 / 客戶總表 / 品號庫存
- 支援拖拉檔案選擇路徑，並記憶上次使用的設定

---

## 系統需求

- Windows 10 / 11
- 可存取 NAS 共用路徑與 ERP SQL Server
- [ODBC Driver 17 for SQL Server](https://learn.microsoft.com/zh-tw/sql/connect/odbc/download-odbc-driver-for-sql-server)

> Python 執行環境已隨附於 `system/python_portable/`，執行端不需另外安裝 Python。

---

## 快速開始

### 1. 取得專案

```bash
git clone <repo-url>
cd order-monitor
```

### 2. 建立連線設定

依照範本建立 `credentials.py`，填入實際的 ERP 連線資訊：

```bash
cp system/credentials.example.py system/credentials.py
# 編輯 system/credentials.py，填入伺服器位址、帳密、公司代碼對照表
```

### 3. 建立建置環境（一次性）

```powershell
cd system
python -m venv .venv-build
.\.venv-build\Scripts\Activate.ps1
pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 tkinterdnd2==0.4.3 pywin32==306 pyinstaller
cd ..
```

### 4. 調整預設路徑（選擇性）

編輯 `system/config.ini`，設定 NAS 預設路徑與來源工作表名稱。

### 5. 打包 EXE

```powershell
.\build.ps1
```

輸出位置：`dist/訂單未交量報表産生器/`，將此資料夾複製給倉管人員。

### 6. 執行

雙擊 `dist/訂單未交量報表産生器/訂單未交量報表産生器.exe`。

---

## 設定檔說明

| 檔案 | 用途 |
|---|---|
| `system/credentials.py` | ERP 伺服器位址、帳密、公司代碼對照表（不進版控） |
| `system/config.ini` | NAS 預設路徑、來源工作表名稱等非機密設定 |

---

## 專案結構

```
order-monitor/
├── build.ps1                  # 一鍵打包腳本
├── order-monitor.spec         # PyInstaller 打包設定
└── system/
    ├── main.py                # 主流程與應用程式進入點
    ├── ui.py                  # Tkinter GUI
    ├── reader.py              # 來源 Excel 讀取（openpyxl，唯讀）
    ├── writer.py              # 四頁式報表產生
    ├── erp.py                 # ERP 庫存查詢（pyodbc）
    ├── utils.py               # 共用工具、欄位常數、路徑記憶
    ├── credentials.py         # 連線帳密（不進版控）
    ├── credentials.example.py # 帳密範本
    └── config.ini             # 非機密設定
```

---

## 套件相依

```
pandas==2.2.2
openpyxl==3.1.2
pyodbc==5.2.0
tkinterdnd2==0.4.3
pywin32==306
```

---

## 備註

- `system/credentials.py` 與 `system/python_portable/` 已列入 `.gitignore`，不納入版本控管。
- 來源 Excel 一律以唯讀模式開啟，工具不會修改原始檔案。
- 庫存扣減依時間順序進行（過期優先，再依交期由舊到新），確保跨客戶分配結果一致。
