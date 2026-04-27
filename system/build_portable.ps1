$ErrorActionPreference = "Stop"

$PYTHON_VER = "3.11.9"
$ZIP_URL = "https://www.python.org/ftp/python/$PYTHON_VER/python-$PYTHON_VER-embed-amd64.zip"
$PIP_URL = "https://bootstrap.pypa.io/get-pip.py"
$PORTABLE_DIR = "$PSScriptRoot\python_portable"

Write-Host "Creating python_portable directory..."
if (Test-Path $PORTABLE_DIR) {
    Remove-Item -Recurse -Force $PORTABLE_DIR
}
New-Item -ItemType Directory -Force -Path $PORTABLE_DIR | Out-Null

Write-Host "Downloading Python Embeddable $PYTHON_VER..."
Invoke-WebRequest -Uri $ZIP_URL -OutFile "$PORTABLE_DIR\python.zip"

Write-Host "Extracting Python..."
Expand-Archive -Path "$PORTABLE_DIR\python.zip" -DestinationPath $PORTABLE_DIR -Force
Remove-Item "$PORTABLE_DIR\python.zip"

Write-Host "Configuring _pth file to enable site-packages..."
$PTH_FILE = "$PORTABLE_DIR\python311._pth"
$pthContent = Get-Content $PTH_FILE
$pthContent = $pthContent -replace '#import site', 'import site'
Set-Content -Path $PTH_FILE -Value $pthContent

Write-Host "Downloading get-pip.py..."
Invoke-WebRequest -Uri $PIP_URL -OutFile "$PORTABLE_DIR\get-pip.py"

Write-Host "Installing pip..."
& "$PORTABLE_DIR\python.exe" "$PORTABLE_DIR\get-pip.py" --no-warn-script-location

Write-Host "Installing required packages..."
& "$PORTABLE_DIR\python.exe" -m pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 tkinterdnd2 pywin32 --no-warn-script-location

Remove-Item "$PORTABLE_DIR\get-pip.py"
Write-Host "Done! Portable Python is ready at .\python_portable"
