@echo off
setlocal

echo.
echo  === Uninstalling WPS AI Assistant ===
echo.

echo [1/3] Stopping service...
taskkill /f /im wps_ai_agent.exe 2>nul
echo       done

echo [2/3] Removing auto-start entry...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "WPS_AI_Agent" /f >nul 2>&1
echo       done

echo [3/3] Removing WPS plugin registration...
if exist "%APPDATA%\kingsoft\wps\jsaddons\publish.xml" (
    powershell -NoProfile -Command "(Get-Content '%APPDATA%\kingsoft\wps\jsaddons\publish.xml' -Raw) -replace '(?s)\s*<jspluginonline[^>]*name=""wps-ai-agent""[^/]*/>', '' | Set-Content '%APPDATA%\kingsoft\wps\jsaddons\publish.xml'"
)
if exist "%~dp0wps_agent.lock" del /f "%~dp0wps_agent.lock" >nul
echo       done

echo.
echo  Uninstall complete. Please restart WPS Office.
echo.
pause
