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
    echo [WARN] Запустите скрипт от имени администратора

    pause
    exit /b
)

:: ================== Configure options ==================
setlocal enabledelayedexpansion

:: Debug option
set "_debug=0"

:: Windows activation option
set "_winActivation=0"

:: Office activation option
set "_officeActivation=0"

:: KMS path and filenames
set "_kmsPath=%~dp0"

set "_kmsFile=kms_auto"

set "_kmsExt=.bat"

set "_kmsLogExt=.log"

:: Log option
set "_log=1"

set "_logPath=%_kmsPath%%_kmsFile%%_kmsLogExt%"

:: vlmcsd config
set "_vlmcsdURL=https://github.com/Wind4/vlmcsd/releases/latest/download/binaries.tar.gz"

set "_downloadPath=%TEMP%\vlmcsd_binaries.tar.gz"

set "_extractPath=%TEMP%\vlmcsd_bin"

set "_exeName=vlmcsd.exe"

set "_vlmcsdLocalPath=%_extractPath%\%_exeName%"

set "_vlmcsdPort=1688"

:: Debug
if "%_debug%"=="1" (
    if not exist "%_kmsPath%%_kmsFile%%_kmsExt%" (
        echo [ERROR] Не найден %_kmsFile%%_kmsExt% в %_kmsPath%
        if "%_log%"=="1" call :log "Не найден %_kmsFile%%_kmsExt%"

        exit /b
    )
)

:: Get local ip
for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %ComputerName% ^| findstr [') do set NetworkIP=%%a

:: ========================= vlmcsd update =========================
call :updateVLMCS

:: ========= RUN VLMCD =========
tasklist | find /i "%_exeName%" >nul

if %errorlevel% neq 0 (
    echo [INFO] Запуск локального KMS-сервера...
    start "" /min "%_vlmcsdLocalPath%" -P %_vlmcsdPort%

    timeout /t 2 >nul
)

:: ================== Start Script ==================
call :main

:: ========= MENU =========
:show_menu
echo.
echo ╔═══════════════════════════════════════╗
echo ║           WINDOWS / OFFICE            ║
echo ╠═══════════════════════════════════════╣
echo ║ 1. Activate Windows                   ║
echo ║ 2. Activate Microsoft Office          ║
echo ║ 3. Check activation status            ║
echo ║ 4. Delete kms                         ║
echo ║ 5. Exit                               ║
echo ╚═══════════════════════════════════════╝
echo.

set /p choice=Select function: 
if "%choice%"=="1" call :activateWindows & goto :show_menu
if "%choice%"=="2" call :activateOffice  & goto :show_menu
if "%choice%"=="3" call :checkStatus     & goto :show_menu
if "%choice%"=="4" call :resetKMS        & goto :show_menu
if "%choice%"=="5" exit /b

echo [ERROR] Неверный выбор. Повторите.
pause

@REM goto :show_menu

:: ========================== Functions ==========================

:main
call :cmd_clear
if "%_log%"=="1" call :log_clear

color 0A
call :show_banner
call :show_menu

:log
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

:show_banner
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
echo.
echo.
echo.
exit /b

:show_footer
echo MIT license
echo github.com/fxhxyz4/kms
exit /b

:updateVLMCS
echo [INFO] Проверка/обновление vlmcsd...

if exist "%_vlmcsdLocalPath%" del /q "%_vlmcsdLocalPath%"
if exist "%_extractPath%" rmdir /s /q "%_extractPath%"

powershell -Command "Invoke-WebRequest -Uri '%_vlmcsdURL%' -OutFile '%_downloadPath%'" 

if not exist "%_downloadPath%" (
    echo [ERROR] Не удалось скачать vlmcsd.
    exit /b
)

mkdir "%_extractPath%" >nul 2>&1

tar -xf "%_downloadPath%" -C "%_extractPath%"

set "_sourceVlmcsd=%_extractPath%\binaries\Windows\intel\vlmcsd-Windows-x64.exe"

if not exist "!_sourceVlmcsd!" (
    echo [ERROR] vlmcsd-Windows-x64.exe не найден после распаковки.
    exit /b
)

copy /y "!_sourceVlmcsd!" "%_vlmcsdLocalPath%" >nul

echo [INFO] vlmcsd обновлён.
exit /b

:activateWindows
echo [INFO] Активация Windows...

:: Определяем корректный cscript (64‑бит если доступен)
if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)

:: Устанавливаем KMS‑ключ (GVLK)
"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX

:: Очищаем предыдущие настройки KMS и задаём локальный хост
"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /ckms
timeout /t 1 >nul
"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /skms localhost:1688

:: Запускаем активацию
"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /ato
if %errorlevel% neq 0 (
    echo [ERROR] Активация не удалась. Код: %errorlevel%
) else (
    echo [INFO] Активация завершена успешно!
)

"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /xpr

echo.
pause
exit /b

:resetKMS
echo [INFO] Удаление ключа и сброс настроек KMS...
if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)

"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /upk
timeout /t 1 >nul
"%_cscript%" //nologo "%windir%\System32\slmgr.vbs" /ckms

echo [INFO] KMS сброшен.
pause
exit /b

:checkStatus
echo.
echo [INFO] Статус Windows:
if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)
"%_cscript%" //nologo "%windir%\system32\slmgr.vbs" /dli
echo.

echo [INFO] Статус Office:
set "officeScript="
for /f "delims=" %%A in ('dir /b /s ospp.vbs 2^>nul') do (
    set "officeScript=%%A"
    goto :officeCheck
)
echo [WARN] ospp.vbs для Office не найден.
pause
exit /b

:officeCheck
echo --- !officeScript! ---
cscript //nologo "!officeScript!" /dstatus
echo.
pause
exit /b

:activateOffice
echo [INFO] Поиск Office...
set "officeScript="

for /f "delims=" %%A in ('dir /b /s ospp.vbs 2^>nul') do (
    set "officeScript=%%A"
    goto :officeFound
)

echo [ERROR] ospp.vbs не найден.
pause
exit /b

:officeFound
echo [INFO] Найден: !officeScript!

:: Устанавливаем адрес локального KMS
cscript //nologo "!officeScript!" /sethst:127.0.0.1

:: Запускаем активацию Office
cscript //nologo "!officeScript!" /act

echo.
pause
exit /b