@echo off
cd /d "%~dp0"
echo.
echo ========================================
echo   通用 MSIX 打包工具 v1.0
echo ========================================
echo.
powershell -ExecutionPolicy Bypass -File make_msix.ps1
echo.
pause
