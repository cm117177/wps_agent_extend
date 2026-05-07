@echo off
setlocal

echo.
echo  ===  WPS AI Assistant Build  ===
echo.

echo [1/3] Installing Python dependencies...
pip install pyinstaller || goto :fail
pip install pymupdf || goto :fail
pip install fastapi || goto :fail
pip install uvicorn || goto :fail
pip install openai || goto :fail
pip install python-multipart || goto :fail

echo [2/3] Building executable (first run takes 3-5 minutes)...
pyinstaller ^
  --onefile ^
  --noconsole ^
  --name wps_ai_agent ^
  --hidden-import uvicorn.logging ^
  --hidden-import uvicorn.loops.auto ^
  --hidden-import uvicorn.protocols.http.auto ^
  --hidden-import uvicorn.protocols.websockets.auto ^
  --hidden-import uvicorn.lifespan.on ^
  --hidden-import uvicorn.lifespan.off ^
  --collect-all uvicorn ^
  --collect-all fastapi ^
  --collect-all starlette ^
  --collect-all pymupdf ^
  --exclude-module uvloop ^
  --exclude-module pytest ^
  main.py
if errorlevel 1 goto :fail

echo [3/3] Assembling distribution folder...
set OUT=dist\WPS_AI_Assistant
if exist "%OUT%" rd /s /q "%OUT%"
mkdir "%OUT%"
mkdir "%OUT%\frontend"
mkdir "%OUT%\wpsaddon"

move dist\wps_ai_agent.exe "%OUT%\" >nul
xcopy /E /I /Y frontend  "%OUT%\frontend"  >nul
xcopy /E /I /Y wpsaddon  "%OUT%\wpsaddon"  >nul
copy install.bat   "%OUT%\" >nul
copy uninstall.bat "%OUT%\" >nul
copy README.txt    "%OUT%\" >nul

:: Write a blank config so your own API key is NOT distributed
echo {"api_key":"","model":"deepseek-chat","base_url":"https://api.deepseek.com/v1","temperature":0.3,"max_tokens":2048} > "%OUT%\config.json"

echo.
echo  === Build complete! ===
echo  Distribution folder: dist\WPS_AI_Assistant\
echo  Zip that folder and send to friends.
echo.
pause
exit /b 0

:fail
echo.
echo  === Build FAILED. See error above. ===
echo.
pause
exit /b 1
