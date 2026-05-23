@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo   MIS Generator - one-click setup ^& launch
echo ============================================================
echo.

if not exist ".env.local" (
  echo ERROR: .env.local is missing. Please create it first.
  pause
  exit /b 1
)

findstr /B /C:"ANTHROPIC_API_KEY=sk-" .env.local >nul 2>&1
if errorlevel 1 (
  echo WARNING: No API key detected in .env.local.
  echo Open .env.local in Notepad, paste your key after ANTHROPIC_API_KEY=, save, then re-run this.
  echo.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Installing dependencies ^(first run takes 1-3 minutes^)...
  echo.
  call npm install
  if errorlevel 1 (
    echo.
    echo npm install failed. Check the messages above.
    pause
    exit /b 1
  )
  echo.
  echo Dependencies installed.
  echo.
)

echo Starting dev server at http://localhost:3000
echo Press Ctrl+C in this window to stop the server.
echo.

call npm run dev
pause
