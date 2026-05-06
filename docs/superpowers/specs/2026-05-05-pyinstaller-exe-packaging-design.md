---
title: PyInstaller EXE 打包設計
date: 2026-05-05
status: approved
---

# PyInstaller EXE 打包設計

## 背景與目標

目前程式以 VBS + portable Python 的形式部署至倉管人員桌面，`credentials.py` 明文放在 `system/` 資料夾，帳密對任何可瀏覽資料夾的人員完全可見。

目標：
1. 將帳密編譯進 EXE 二進位檔，移除明文 `credentials.py`
2. 簡化部署：一個資料夾，雙擊 EXE 執行，不再需要 VBS
3. 維持可維護性：帳密或程式更新時有明確的重建流程

---

## 部署結構（before / after）

### 現在

```
order-monitor/                         ← NAS 或桌面
├── 訂單未交量報表產生器.vbs            ← 入口
└── system/
    ├── main.py
    ├── credentials.py                  ← 帳密明文可見
    ├── config.ini
    └── python_portable/               ← 龐大，需手動建置
```

### 打包後（倉管人員拿到的資料夾）

```
訂單未交量報表產生器/                   ← 整包複製到桌面
├── 訂單未交量報表產生器.exe            ← 雙擊執行（帳密已編譯入內）
├── config.ini                          ← 可手動調整 NAS 路徑
└── _internal/                          ← PyInstaller 支援檔（DLL、pyd）
    └── ...
```

- VBS 完全移除
- `credentials.py` 不存在於資料夾中（已編譯進 EXE）
- `paths.json` 寫至 `%LOCALAPPDATA%\OrderMonitor\`（原本已如此，不變）

---

## 建置流程

### 新增檔案

| 檔案 | 用途 |
|---|---|
| `order-monitor.spec` | PyInstaller 打包設定，處理相容性 hooks |
| `build.ps1` | 一鍵建置腳本（放專案根目錄）|

### 建置指令

```powershell
# 開發機上執行一次，之後每次更新都執行這一行
.\build.ps1
```

`build.ps1` 執行流程：
1. 確認 PyInstaller 已安裝（`pip install pyinstaller`）
2. 執行 `pyinstaller order-monitor.spec`
3. 將 `system/config.ini` 複製到 `dist/訂單未交量報表產生器/`
4. 輸出提示：「建置完成，輸出在 dist/訂單未交量報表產生器/」

### 更新流程（給下一位維護者）

| 變更類型 | 步驟 |
|---|---|
| 程式邏輯更新 | 修改 `system/*.py` → 執行 `build.ps1` → 覆蓋 NAS 資料夾 |
| 帳密更新 | 修改 `system/credentials.py` → 執行 `build.ps1` → 覆蓋 NAS 資料夾 |
| NAS 路徑更新 | 直接編輯 `dist/訂單未交量報表產生器/config.ini`（不需重新打包）|

---

## 相容性處理（.spec 設定）

以下三個套件需要特別處理：

### tkinterdnd2

PyInstaller 不會自動找到 tkdnd 的原生 tcl/tk extension。需在 `.spec` 的 `datas` 中手動指定 tkinterdnd2 套件資料夾路徑：

```python
import tkinterdnd2
datas = [(os.path.dirname(tkinterdnd2.__file__), 'tkinterdnd2')]
```

### pywin32（win32com）

COM 元件需要 `pywintypes` 和 `win32com` 的隱藏 import。在 `.spec` 中加入：

```python
hiddenimports = ['win32com', 'win32com.client', 'pywintypes', 'win32api']
```

### pyodbc

通常自動解決。若連線失敗，確認目標機器已安裝 `ODBC Driver 17 for SQL Server`（這是外部系統驅動，不能打包進 EXE）。

---

## 已就緒的程式碼

`utils.py` 的 `_base_dir()` 已有 `sys.frozen` 判斷，打包後會正確找到 `config.ini`：

```python
def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)  # EXE 所在目錄
    return os.path.dirname(os.path.abspath(__file__))
```

`paths.json` 已寫至 `%LOCALAPPDATA%\OrderMonitor\`，無需異動。

---

## 不在範圍內

- 帳密加密（DPAPI）：選擇 EXE 打包方案後，帳密已藏入二進位，不額外加密
- 自動更新機制：維護者手動覆蓋資料夾即可
- 程式碼簽章（Code Signing）：非必要，不在本次範圍
