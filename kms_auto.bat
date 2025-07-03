@echo off
:: @R fxhxyz
:: @R Mit license
:: @R fxhxyz.vercel.app
:: @R github.com/fxhxyz4

chcp 65001 >nul

:: ================= Admin Check & Start ================= ::

net session >nul 2>&1

if %errorlevel% neq 0 (
    color 0C
    echo [WARN] Run script as Admin.

    pause
    exit /b
) else (
    color 08
    call :main
)

:: ================== Configure options ================== ::

:: enable delayed variable expansion
setlocal enabledelayedexpansion

:: Debug option
set _debug=0

:: Windows activation option
set _winActivation=0

:: MS Office activation option
set _officeActivation=0

:: kms_auto path
set _kmsPath=%~dp0

:: kms_auto file name
set _kmsFile=kms_auto

:: kms_auto log ext
set _kmsLogExt=.log

:: kms_auto ext
set _kmsExt=.bat

:: Logger option
set _log=1

:: Full log path
set _logPath=%_kmsPath%%_kmsFile%%_kmsLogExt%

if "%_debug%"=="1" (
    if not exist "%_kmsPath%%_kmsFile%%_kmsExt%" (
        echo [ERROR] Not found %_kmsFile%%_kmsExt% in %_kmsPath%
        if "%_log%"=="1" call :log "Not found %_kmsFile%%_kmsExt% in %_kmsPath%"
        exit /b
    )
)

:: ========================== Functions ========================== ::

:log
:: call :log "message"
set "_msg=%~1"
for /f "tokens=1-2 delims= " %%a in ("%date% %time%") do (
    echo [%%a %%b] !_msg!>> "%_logPath%"
)

:: Clear log file
:log_clear
if exist "%_logPath%" del /q "%_logPath%"

:: Cmd clear
:cmd_clear
cls

:: Show start banner
:showBanner
echo.
echo.
color 0B
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
color 0A
echo Mit license
echo github.com/fxhxyz4/kms
echo.
echo.
pause

:main
call :cmd_clear
if "%_log%"=="1" call :log_clear
call :showBanner

if "%_log%"=="1" call :log "Script running"

pause
