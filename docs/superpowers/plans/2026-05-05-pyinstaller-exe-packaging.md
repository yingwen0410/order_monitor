# PyInstaller EXE 打包 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 將訂單未交量報表產生器打包成 Windows EXE（資料夾形式），帳密編譯進二進位檔，移除 VBS 入口點，部署只需複製一個資料夾。

**Architecture:** 使用 PyInstaller `--onedir` 模式，以 `order-monitor.spec` 處理 tkinterdnd2 / pywin32 / openpyxl 的相容性問題。開發者執行 `build.ps1` 一鍵產出 `dist/訂單未交量報表産生器/`，其中包含 EXE 和 config.ini；`credentials.py` 編譯入 EXE，不出現在輸出資料夾中。

**Tech Stack:** PyInstaller 6.x、Python 3.11、tkinterdnd2、pywin32、pyodbc、openpyxl、pandas

---

## File Map

| 動作 | 路徑 | 說明 |
|---|---|---|
| 新增 | `order-monitor.spec` | PyInstaller 打包設定 |
| 新增 | `build.ps1` | 一鍵建置腳本 |
| 修改 | `.gitignore` | 修正 dist/build 路徑、保留 spec 進版控 |
| 修改 | `README.md` | 移除 VBS 說明，加入 build 說明 |
| 修改 | `system/CLAUDE.md` | 更新部署方式 |
| 刪除 | `訂單未交量報表産生器.vbs` | EXE 取代後不再需要 |

---

## Task 1: 準備建置環境

**Files:**
- （無需建立檔案，確認環境可用）

- [ ] **Step 1: 確認本機有 Python 3.11（非 python_portable，是開發機的 Python）**

```powershell
python --version
# 預期：Python 3.11.x
```

若無，至 [python.org](https://www.python.org/downloads/) 下載安裝。

- [ ] **Step 2: 在 `system/` 內建立建置用 venv**

```powershell
cd system
python -m venv .venv-build
.\.venv-build\Scripts\Activate.ps1
```

- [ ] **Step 3: 安裝所有套件 + PyInstaller**

```powershell
pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 tkinterdnd2==0.4.3 pywin32==306 pyinstaller
```

- [ ] **Step 4: 確認 PyInstaller 可執行**

```powershell
pyinstaller --version
# 預期：6.x.x
```

- [ ] **Step 5: 回到專案根目錄**

```powershell
cd ..
```

> 注意：建置 venv 不進版控（`.gitignore` 已排除 `system/.venv/`）。如果 `.venv-build` 未被排除，執行完 Task 2 後加到 `.gitignore`。

---

## Task 2: 修正 .gitignore

**Files:**
- Modify: `.gitignore`

目前 `.gitignore` 有 `*.spec`（會把要進版控的 spec 排除）且 dist/build 路徑只覆蓋 `system/` 底下。需修正。

- [ ] **Step 1: 開啟 `.gitignore`，找到 Build and Dist 區塊**

目前內容：
```
# Build and Dist
system/build/
system/dist/
*.spec
```

- [ ] **Step 2: 改成以下內容（保留 spec 進版控，覆蓋根目錄 dist/build）**

```
# Build and Dist
build/
dist/
system/build/
system/dist/
system/.venv-build/
```

- [ ] **Step 3: 確認 `order-monitor.spec` 不在排除清單中（已移除 `*.spec`）**

```powershell
git check-ignore -v order-monitor.spec
# 預期：無輸出（表示不被排除）
```

- [ ] **Step 4: Commit**

```powershell
git add .gitignore
git commit -m "fix: update gitignore for pyinstaller build output"
```

---

## Task 3: 建立 order-monitor.spec

**Files:**
- Create: `order-monitor.spec`

- [ ] **Step 1: 確認 tkinterdnd2 路徑（在建置 venv 啟動狀態下）**

```powershell
system\.venv-build\Scripts\python.exe -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))"
# 預期：...system\.venv-build\Lib\site-packages\tkinterdnd2
```

- [ ] **Step 2: 建立 `order-monitor.spec`（內容如下，直接貼上）**

```python
# order-monitor.spec
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# tkinterdnd2 需手動指定資料夾，讓 tcl/tk DnD extension 一起打包
import tkinterdnd2 as _tkdnd
_tkdnd_path = os.path.dirname(_tkdnd.__file__)

a = Analysis(
    ['system/main.py'],
    pathex=['system'],          # 讓 import utils / reader / erp / writer / ui 能解析
    binaries=[],
    datas=[
        (_tkdnd_path, 'tkinterdnd2'),       # DnD extension
        *collect_data_files('openpyxl'),    # openpyxl 內建樣板檔
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.shell',
        'pywintypes',
        'win32api',
        'win32con',
        'win32gui',
        *collect_submodules('win32com'),
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='訂單未交量報表産生器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,              # 不顯示 console 視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='訂單未交量報表産生器',
)
```

- [ ] **Step 3: Commit**

```powershell
git add order-monitor.spec
git commit -m "feat: add pyinstaller spec with tkinterdnd2 and pywin32 hooks"
```

---

## Task 4: 建立 build.ps1

**Files:**
- Create: `build.ps1`

- [ ] **Step 1: 建立 `build.ps1`（內容如下）**

```powershell
$ErrorActionPreference = "Stop"

$DIST_NAME = "訂單未交量報表産生器"
$DIST_PATH = "dist\$DIST_NAME"
$VENV_PYTHON = "system\.venv-build\Scripts\python.exe"

# 確認 credentials.py 存在（帳密必須在打包前就位）
if (-not (Test-Path "system\credentials.py")) {
    Write-Error @"
找不到 system\credentials.py
請先依照 system\credentials.example.py 建立帳密檔案，再執行打包。
"@
    exit 1
}

# 確認建置 venv 存在
if (-not (Test-Path $VENV_PYTHON)) {
    Write-Error @"
找不到建置環境：$VENV_PYTHON
請先執行以下指令建立：
  cd system
  python -m venv .venv-build
  .\.venv-build\Scripts\Activate.ps1
  pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 tkinterdnd2==0.4.3 pywin32==306 pyinstaller
"@
    exit 1
}

Write-Host "=== 開始打包 ===" -ForegroundColor Cyan
& $VENV_PYTHON -m PyInstaller order-monitor.spec --clean --noconfirm

if (-not $?) {
    Write-Error "PyInstaller 打包失敗，請查看上方錯誤訊息。"
    exit 1
}

Write-Host "=== 複製 config.ini ===" -ForegroundColor Cyan
Copy-Item "system\config.ini" "$DIST_PATH\config.ini" -Force

Write-Host ""
Write-Host "=== 打包完成 ===" -ForegroundColor Green
Write-Host "輸出位置：$DIST_PATH"
Write-Host "請將此資料夾整包提供給倉管人員。"
```

- [ ] **Step 2: Commit**

```powershell
git add build.ps1
git commit -m "feat: add build.ps1 one-click build script"
```

---

## Task 5: 執行第一次建置並煙霧測試

**Files:**
- （不修改程式碼，只驗證建置結果）

- [ ] **Step 1: 執行建置腳本**

```powershell
.\build.ps1
```

預期最後看到：
```
=== 打包完成 ===
輸出位置：dist\訂單未交量報表産生器
```

- [ ] **Step 2: 確認輸出結構正確**

```powershell
Get-ChildItem "dist\訂單未交量報表産生器" -Name
```

預期出現：
- `訂單未交量報表産生器.exe`
- `config.ini`
- `_internal\`

**不應出現** `credentials.py`。

- [ ] **Step 3: 確認 credentials.py 不在 dist 中**

```powershell
Get-ChildItem "dist\" -Recurse -Filter "credentials.py"
# 預期：無任何輸出
```

- [ ] **Step 4: 雙擊 EXE 確認啟動畫面出現**

手動操作：雙擊 `dist\訂單未交量報表産生器\訂單未交量報表産生器.exe`

預期行為：
1. 出現深色啟動畫面（「訂單未交量報表産生器 / 正在啟動中，請稍候…」）
2. 啟動畫面消失後，主視窗 GUI 開啟
3. 公司別下拉選單有資料（來自 credentials.py 的 COMPANIES）

若啟動後立刻崩潰，前往 Task 6（排錯）。若正常開啟，跳到 Task 7。

---

## Task 6: 排錯——常見打包問題（視需要執行）

本 Task 僅在 Task 5 Step 4 失敗時執行。

### 6a: 找出錯誤訊息

- [ ] **Step 1: 改成 console 模式重新打包，看到實際錯誤**

開啟 `order-monitor.spec`，找到 `console=False`，暫時改成 `console=True`：

```python
exe = EXE(
    ...
    console=True,   # 暫時改成 True 看錯誤
    ...
)
```

- [ ] **Step 2: 重新建置並執行**

```powershell
.\build.ps1
.\dist\訂單未交量報表産生器\訂單未交量報表産生器.exe
```

記錄 console 視窗顯示的錯誤訊息。

### 6b: tkinterdnd2 無法載入（ModuleNotFoundError 或 TclError）

- [ ] **Step 1: 確認 tkinterdnd2 資料夾有被打包進去**

```powershell
Test-Path "dist\訂單未交量報表産生器\_internal\tkinterdnd2"
# 預期：True
```

若為 False，表示 spec 中的 `_tkdnd_path` 路徑不正確。重新確認 tkinterdnd2 安裝路徑：

```powershell
system\.venv-build\Scripts\python.exe -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))"
```

將輸出路徑直接硬碼到 spec 的 `datas` 中再試一次：

```python
datas=[
    (r'C:\確切路徑\tkinterdnd2', 'tkinterdnd2'),
    ...
]
```

### 6c: pywin32 錯誤（ImportError: DLL load failed）

- [ ] **Step 1: 執行 pywin32 post-install（如果尚未執行過）**

```powershell
system\.venv-build\Scripts\python.exe system\.venv-build\Scripts\pywin32_postinstall.py -install
```

- [ ] **Step 2: 重新建置**

```powershell
.\build.ps1
```

### 6d: 找不到 openpyxl 樣板（jinja2 / template 錯誤）

- [ ] **Step 1: 確認 openpyxl data files 有打包**

```powershell
Test-Path "dist\訂單未交量報表産生器\_internal\openpyxl"
# 預期：True
```

若 False，在 spec 的 `datas` 中加入明確路徑：

```python
import openpyxl
*collect_data_files('openpyxl'),
# 如果 collect_data_files 失效，改用：
(os.path.join(os.path.dirname(openpyxl.__file__), 'templates'), 'openpyxl/templates'),
```

### 6e: 排錯完成後恢復 console=False

- [ ] **Step 1: 將 spec 中 `console=True` 改回 `console=False`，重新建置**

```powershell
.\build.ps1
```

---

## Task 7: 移除 VBS，更新文件

**Files:**
- Delete: `訂單未交量報表産生器.vbs`
- Modify: `README.md`
- Modify: `system/CLAUDE.md`

- [ ] **Step 1: 刪除 VBS 檔案**

```powershell
Remove-Item "訂單未交量報表産生器.vbs"
```

- [ ] **Step 2: 更新 README.md — 快速開始區塊**

找到 `### 3. 建置 Python 執行環境` 區塊，整段替換為：

```markdown
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
```

- [ ] **Step 3: 更新 README.md — 專案結構區塊**

將 `訂單未交量報表産生器.vbs` 那行替換為：

```markdown
order-monitor/
├── build.ps1                  # 一鍵打包腳本
├── order-monitor.spec         # PyInstaller 打包設定
└── system/
    ├── main.py
    ├── ui.py
    ├── reader.py
    ├── writer.py
    ├── erp.py
    ├── utils.py
    ├── credentials.py         # 連線帳密（不進版控）
    ├── credentials.example.py
    ├── config.ini
    └── build_portable.ps1     # 已棄用（保留供參考）
```

- [ ] **Step 4: 更新 system/CLAUDE.md — 啟動與執行區塊**

找到：
```markdown
倉管端使用外層的 `訂單未交量報表産生器.vbs` 雙擊執行。
```

替換為：
```markdown
倉管端雙擊 `dist/訂單未交量報表産生器/訂單未交量報表産生器.exe` 執行。
維護者更新程式或帳密後，執行根目錄的 `build.ps1` 重新打包。
```

- [ ] **Step 5: Commit**

```powershell
git add -A
git commit -m "feat: replace VBS with PyInstaller EXE, add build.ps1 and spec"
```

---

## Task 8: 最終驗證

- [ ] **Step 1: 確認 git 狀態乾淨**

```powershell
git status
# 預期：nothing to commit, working tree clean
```

- [ ] **Step 2: 確認 dist/ 不在版控中**

```powershell
git check-ignore -v dist
# 預期：.gitignore:... dist
```

- [ ] **Step 3: 確認 order-monitor.spec 在版控中**

```powershell
git ls-files order-monitor.spec
# 預期：order-monitor.spec
```

- [ ] **Step 4: 模擬倉管人員操作**

1. 將 `dist\訂單未交量報表産生器\` 整個資料夾複製到桌面
2. 雙擊 `訂單未交量報表産生器.exe`
3. 確認啟動畫面 → 主視窗正常開啟
4. 確認公司別下拉選單有正確的公司資料
5. 確認資料夾內看不到 `credentials.py`

---

## 帳密更新流程（給下一位維護者）

```
1. 修改 system\credentials.py（改 server IP / 密碼 / 公司清單）
2. 執行 .\build.ps1
3. 將 dist\訂單未交量報表産生器\ 覆蓋到 NAS 讓使用者重新複製
```
