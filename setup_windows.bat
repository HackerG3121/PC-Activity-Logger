@echo off
title PC Activity Logger — Windows Setup
color 1F
cls

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║   PC ACTIVITY LOGGER — Windows Setup                ║
echo  ║   Installs deps, sets sync folder, automates run    ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: ── Step 1: Check Python ──────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Download from: https://python.org/downloads
    echo  IMPORTANT: Tick "Add Python to PATH" during install.
    pause & exit /b 1
)
echo  [OK] Python found.

:: ── Step 2: Install packages ──────────────────────────────
echo.
echo  Installing required packages...
pip install psutil openpyxl pywin32 --quiet --upgrade
if errorlevel 1 (
    echo  [WARN] Some packages may not have installed cleanly.
) else (
    echo  [OK] Packages installed: psutil, openpyxl, pywin32
)

:: ── Step 3: Choose sync folder ────────────────────────────
echo.
echo  Where should the log file be saved?
echo  (This folder syncs to your phone for mobile access)
echo.
echo    1. OneDrive   (recommended — built into Windows 11)
echo    2. Google Drive
echo    3. Dropbox
echo    4. Desktop    (no mobile sync)
echo.
set /p CHOICE="  Enter 1-4: "

if "%CHOICE%"=="1" set FOLDER=OneDrive
if "%CHOICE%"=="2" set FOLDER=Google Drive
if "%CHOICE%"=="3" set FOLDER=Dropbox
if "%CHOICE%"=="4" set FOLDER=Desktop
if not defined FOLDER (
    echo  Invalid choice. Defaulting to Desktop.
    set FOLDER=Desktop
)

:: Update LOG_FILE path in script
set SCRIPT=%~dp0pc_logger.py
powershell -Command ^
  "(Get-Content '%SCRIPT%') ^
   -replace 'OneDrive.*PC_ActivityLog', '\"%FOLDER%\", \"PC_ActivityLog' ^
   | Set-Content '%SCRIPT%'"
echo  [OK] Log path → %USERPROFILE%\%FOLDER%\PC_ActivityLog.xlsx

:: ── Step 4: Schedule task every 15 minutes ────────────────
echo.
echo  Setting up automatic run every 15 minutes via Task Scheduler...

schtasks /query /tn "PC_ActivityLogger" >nul 2>&1
if not errorlevel 1 (
    schtasks /delete /tn "PC_ActivityLogger" /f >nul 2>&1
    echo  [INFO] Removed existing scheduled task.
)

:: Create task: runs at login AND repeats every 15 min, hidden
schtasks /create ^
  /tn "PC_ActivityLogger" ^
  /tr "python \"%SCRIPT%\"" ^
  /sc ONLOGON ^
  /ru %USERNAME% ^
  /rl HIGHEST ^
  /f >nul 2>&1

:: Add 15-minute repetition via XML patch
set XMLFILE=%TEMP%\pclogger_task.xml
schtasks /query /tn "PC_ActivityLogger" /xml > "%XMLFILE%" 2>nul
powershell -Command ^
  "(Get-Content '%XMLFILE%') ^
   -replace '<Repetition/>', ^
   '<Repetition><Interval>PT15M</Interval><Duration>P1D</Duration><StopAtDurationEnd>false</StopAtDurationEnd></Repetition>' ^
   | Set-Content '%XMLFILE%'"
schtasks /delete /tn "PC_ActivityLogger" /f >nul 2>&1
schtasks /create /tn "PC_ActivityLogger" /xml "%XMLFILE%" /f >nul 2>&1
del "%XMLFILE%" >nul 2>&1

if errorlevel 1 (
    echo  [WARN] Could not create scheduled task. Try running as Administrator.
    echo  Falling back to Startup folder method...
    set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
    (
        echo @echo off
        echo start /min python "%SCRIPT%"
    ) > "%STARTUP%\PC_Logger_Autostart.bat"
    echo  [OK] Added to Startup folder instead.
) else (
    echo  [OK] Task Scheduler: runs every 15 minutes, starts on login.
)

:: ── Step 5: Launch now ────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║   Setup complete! Starting logger now...            ║
echo  ║   Press Ctrl+C at any time to stop.                 ║
echo  ╚══════════════════════════════════════════════════════╝
echo.
python "%SCRIPT%"
pause
