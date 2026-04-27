@echo off
cls
echo Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Downloading installer...
    curl -L -o "%TEMP%\python_setup.exe" "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
    echo Installing Python (please wait)...
    "%TEMP%\python_setup.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
    echo Done. Please close and re-open this window, then run again.
    pause
    exit /b
)

echo Python found. Installing packages...
python -m pip install pandas==2.2.2 openpyxl==3.1.2 pyodbc==5.2.0 -q
echo All packages installed.
echo.
echo IMPORTANT: ODBC Driver 17 for SQL Server must also be installed.
echo Download from Microsoft if not already present:
echo   https://aka.ms/downloadmsodbcsql
pause
