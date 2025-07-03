@echo off
REM fxhxyz
REM MIT license
REM fxhxyz.vercel.app
REM github.com/fxhxyz4

chcp 65001 >nul

:: ================= Admin Check & Start =================
net session >nul 2>&1

if %errorlevel% neq 0 (
    color 0C
    echo [WARN] Run script as Admin.

    pause
    exit /b
)

:: ================== Configure options ==================
setlocal enabledelayedexpansion

:: Debug option
set "_debug=0"

:: Windows activation option
set "_winActivation=0"

:: MS Office activation option
set "_officeActivation=0"

:: kms_auto path
set "_kmsPath=%~dp0"

:: kms_auto file name
set "_kmsFile=kms_auto"

:: kms_auto log ext
set "_kmsLogExt=.log"

:: kms_auto ext
set "_kmsExt=.bat"

:: Logger option
set "_log=1"

:: Full log path
set "_logPath=%_kmsPath%%_kmsFile%%_kmsLogExt%"

if "%_debug%"=="1" (
    if not exist "%_kmsPath%%_kmsFile%%_kmsExt%" (
        echo [ERROR] Not found %_kmsFile%%_kmsExt% in %_kmsPath%
        if "%_log%"=="1" call :log "Not found %_kmsFile%%_kmsExt% in %_kmsPath%"

        exit /b
    )
)

:: ========================== Main Start ==========================
call :main
exit /b

:: ========================== Functions ==========================

:main
call :cmd_clear
if "%_log%"=="1" call :log_clear
call :showBanner

color 0A
call :showFooter

call :showMenu

if "%_log%"=="1" call :log "Script running"
pause
exit /b

:log
REM call :log "message"
set "_msg=%~1"
for /f "tokens=1-2 delims= " %%a in ("%date% %time%") do (
    echo [%%a %%b] !_msg!>> "%_logPath%"
)
exit /b

:log_clear
if exist "%_logPath%" del /q "%_logPath%"
exit /b

:cmd_clear
cls
exit /b

:showBanner
echo.
echo.

echo.
echo ███████╗██╗  ██╗██╗  ██╗██╗  ██╗██╗   ██╗███████╗
echo ██╔════╝╚██╗██╔╝██║  ██║╚██╗██╔╝╚██╗ ██╔╝╚══███╔╝
echo █████╗   ╚███╔╝ ███████║ ╚███╔╝  ╚████╔╝   ███╔╝ 
echo ██╔══╝   ██╔██╗ ██╔══██║ ██╔██╗   ╚██╔╝   ███╔╝  
echo ██║     ██╔╝ ██╗██║  ██║██╔╝ ██╗   ██║   ███████╗
echo ╚═╝     ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝   ╚══════╝

echo.
echo ======= Windows/Office activation script =======

echo.
echo.
exit /b

:showFooter
echo Mit license
echo github.com/fxhxyz4/kms
exit /b

:showMenu
echo.
echo.

echo Enable debug
echo Enable log

echo.

echo.
echo 1. Activate Windows

echo 2. Activate Microsoft Office 365
echo 3. Check activation status

echo 4. Delete activation
echo 5. Exit
