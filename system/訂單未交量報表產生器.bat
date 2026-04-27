@echo off
cls
title Order Report Generator
echo.
echo  ========================================
echo   Starting application, please wait...
echo  ========================================
echo.
pushd "%~dp0"
.\python_portable\python.exe main.py
popd
pause
