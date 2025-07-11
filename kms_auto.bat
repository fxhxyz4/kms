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

set "_vlmcsdPort=1688"
set "_kmsPath=%~dp0"
set "_kmsFile=kms_docker_auto"
set "_kmsExt=.bat"
set "_kmsLogExt=.log"
set "_log=1"
set "_logPath=%_kmsPath%%_kmsFile%%_kmsLogExt%"

for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %ComputerName% ^| findstr [') do set NetworkIP=%%a

call :checkInstallDocker
call :run_kms_docker
call :main

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
goto :show_menu

:main
call :cmd_clear
if "%_log%"=="1" call :log_clear
color 0A
call :show_banner
call :show_menu
exit /b

:checkInstallDocker
docker --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Docker не найден — пытаюсь установить через winget...
    winget install --accept-package-agreements --accept-source-agreements -e Docker.DockerDesktop
    timeout /t 10 >nul
    docker --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Docker всё ещё не найден. Возможно, требуется перезагрузка системы.
        echo Установите Docker вручную: https://www.docker.com/products/docker-desktop
        pause
        exit /b
    )
)
echo [INFO] Docker установлен и готов к использованию.
exit /b

:run_kms_docker
echo [INFO] Проверка запущенного Docker-контейнера KMS...
docker ps --filter "name=vlmcsd_kms" --format "{{.Names}}" | findstr vlmcsd_kms >nul
if %errorlevel% equ 0 (
    echo [INFO] Контейнер vlmcsd_kms уже запущен.
) else (
    echo [INFO] Контейнер vlmcsd_kms не найден. Запускаю...
    docker image inspect mikolatero/vlmcsd >nul 2>&1
    if %errorlevel% neq 0 (
        echo [INFO] Образ mikolatero/vlmcsd не найден. Скачиваю...
        docker pull mikolatero/vlmcsd
    )
    docker run -d --rm --name vlmcsd_kms -p 1688:1688 mikolatero/vlmcsd
    if %errorlevel% neq 0 (
        echo [ERROR] Не удалось запустить контейнер KMS.
        pause
        exit /b
    )
    echo [INFO] Контейнер KMS запущен на порту %_vlmcsdPort%.
)
set "WSL_IP=127.0.0.1"
echo [INFO] Используем IP для активации: %WSL_IP%
exit /b

:activateWindows
echo [INFO] Активация Windows...
if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)
:: KMS client key для Win10 Pro
set KMS_KEY=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /upk
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ipk %KMS_KEY%
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ckms
timeout /t 1 >nul
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /skms 127.0.0.1:%_vlmcsdPort%
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ato
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /xpr
pause
exit /b

:resetKMS
echo [INFO] Удаление ключа и сброс настроек KMS...
if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /upk
timeout /t 1 >nul
%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ckms
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
%_cscript% //nologo "%windir%\system32\slmgr.vbs" /dli
echo.
echo [INFO] Статус Office:
set "officeScript="
for /f "delims=" %%A in ('dir /b /s ospp.vbs 2^>nul') do (
    set "officeScript=%%A"
    goto :officeCheck
)
echo [WARN] ospp.vbs не найден.
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
echo [INFO] Найден: !officeScript!"
cscript //nologo "!officeScript!" /sethst:%WSL_IP%
cscript //nologo "!officeScript!" /act
echo.
pause
exit /b

:cmd_clear
cls
exit /b

:log
set "_msg=%~1"
for /f "tokens=1-2 delims= " %%a in ("%date% %time%") do (
    echo [%%a %%b] !_msg!>> "%_logPath%"
)
exit /b

:log_clear
if exist "%_logPath%" del /q "%_logPath%"
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
exit /b

:show_footer
echo MIT license
echo github.com/fxhxyz4/kms
exit /b