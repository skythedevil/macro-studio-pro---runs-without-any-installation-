@echo off
title Macro Studio Pro Launcher
echo Starting Macro Studio Pro Suite...
powershell -ExecutionPolicy Bypass -File "%~dp0StudioProCode.ps1"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Macro Studio failed to start.
    echo Please check 'MacroStudio_Error.log' for details.
    pause
)
