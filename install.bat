@echo off
setlocal
title WPS AI Assistant - Install

echo.
echo  ===  WPS AI Assistant  Install  ===
echo.

echo [1/3] Starting AI service...
tasklist /fi "imagename eq wps_ai_agent.exe" 2>nul | find /i "wps_ai_agent.exe" >nul
if not errorlevel 1 (
    echo       already running, skipped
    goto :step2
)
start "" /B "%~dp0wps_ai_agent.exe"
timeout /t 4 /nobreak >nul
tasklist /fi "imagename eq wps_ai_agent.exe" 2>nul | find /i "wps_ai_agent.exe" >nul
if errorlevel 1 (
    echo.
    echo  ERROR: service failed to start.
    echo  Try right-clicking install.bat and choose "Run as administrator".
    pause
    exit /b 1
)
echo       done

:step2
echo [2/3] Setting auto-start on login...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "WPS_AI_Agent" /t REG_SZ /d "\"%~dp0wps_ai_agent.exe\"" /f >nul
echo       done

echo [3/3] Registering WPS plugin...
timeout /t 2 /nobreak >nul
echo       done

echo.
echo  =========================================
echo   Install complete!
echo  =========================================
echo.
echo   Next steps:
echo   1. Fully quit WPS Office
echo      (right-click WPS tray icon - Exit)
echo   2. Reopen WPS - click "AI Assistant" tab
echo   3. Click the gear icon, enter your API Key
echo.
echo   The service starts automatically on boot.
echo   No need to repeat this installation.
echo.
pause
