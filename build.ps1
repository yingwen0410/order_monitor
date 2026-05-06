$ErrorActionPreference = "Stop"

$DIST_NAME = "訂單未交量報表産生器"
$DIST_PATH = "dist\$DIST_NAME"
$VENV_PYTHON = "system\.venv-build\Scripts\python.exe"

# Verify credentials.py exists before building (passwords must be in place)
if (-not (Test-Path "system\credentials.py")) {
    Write-Error @"
找不到 system\credentials.py
請先依照 system\credentials.example.py 建立帳密檔案，再執行打包。
"@
    exit 1
}

# Verify build venv exists
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

if ($LASTEXITCODE -ne 0) {
    Write-Error "PyInstaller 打包失敗，請查看上方錯誤訊息。"
    exit 1
}

Write-Host "=== 複製 config.ini ===" -ForegroundColor Cyan
Copy-Item "system\config.ini" "$DIST_PATH\config.ini" -Force

Write-Host ""
Write-Host "=== 打包完成 ===" -ForegroundColor Green
Write-Host "輸出位置：$DIST_PATH"
Write-Host "請將此資料夾整包提供給倉管人員。"
