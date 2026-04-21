@echo off
cd /d "%~dp0"
echo.
echo ========================================
echo   MSIX 打包工具 — 管理员模式
echo ========================================
echo.
powershell -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File `\"%~dp0make_msix.ps1`\" -Config `\"%~dp0config_sublime.json`\"' -Verb RunAs"
echo.
pause
