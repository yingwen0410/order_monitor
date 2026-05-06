$DIST_NAME = "dist\訂單未交量報表産生器"
$VENV_PYTHON = "system\.venv-build\Scripts\python.exe"

if (-not (Test-Path "system\credentials.py")) {
    Write-Host "ERROR: system\credentials.py not found." -ForegroundColor Red
    Write-Host "Please create it from system\credentials.example.py before building." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $VENV_PYTHON)) {
    Write-Host "ERROR: Build venv not found at $VENV_PYTHON" -ForegroundColor Red
    Write-Host "Run the following to set it up:" -ForegroundColor Yellow
    Write-Host "  cd system" -ForegroundColor Yellow
    Write-Host "  python -m venv .venv-build" -ForegroundColor Yellow
    Write-Host "  .\.venv-build\Scripts\Activate.ps1" -ForegroundColor Yellow
    Write-Host "  pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 tkinterdnd2==0.4.3 pywin32==306 pyinstaller" -ForegroundColor Yellow
    exit 1
}

Write-Host "=== Building EXE ===" -ForegroundColor Cyan
& $VENV_PYTHON -m PyInstaller order-monitor.spec --clean --noconfirm
$buildResult = $LASTEXITCODE

if ($buildResult -ne 0) {
    Write-Host "ERROR: PyInstaller failed (exit code $buildResult). See output above." -ForegroundColor Red
    exit 1
}

Write-Host "=== Copying config.ini ===" -ForegroundColor Cyan
Copy-Item "system\config.ini" "$DIST_NAME\config.ini" -Force

Write-Host ""
Write-Host "=== Build complete ===" -ForegroundColor Green
Write-Host "Output: $DIST_NAME"
Write-Host "Copy this folder to warehouse staff."
