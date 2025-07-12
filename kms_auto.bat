@echo off
REM fxhxyz
REM MIT license
REM fxhxyz.vercel.app
REM github.com/fxhxyz4

chcp 65001 >nul

:: ==================== Admin Check ====================
net session >nul 2>&1

if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] Run with admin!

    pause
    exit /b
)

:: ================== Configure options ==================
setlocal enabledelayedexpansion

:: KMS server
set "_vlmcsdPort=1688"

set "_kmsPath=%~dp0"

set "_kmsFile=kms_docker_auto"

set "_kmsExt=.bat"

set "_kmsLogExt=.log"

set "_log=1"

set "_logPath=%_kmsPath%%_kmsFile%%_kmsLogExt%"

:: Get ip
for /f "delims=[] tokens=2" %%a in ('ping -4 -n 1 %ComputerName% ^| findstr [') do set NetworkIP=%%a

:: ===================== Program start =====================
call :main

:show_menu
echo.
echo ╔═══════════════════════════════════════════════╗
echo ║                WINDOWS / OFFICE               ║
echo ╠═══════════════════════════════════════════════╣
echo ║ 1. Activate Windows                           ║
echo ║ 2. Activate Microsoft Office                  ║
echo ║ 3. Check activation status                    ║
echo ║ 4. Delete kms                                 ║
echo ║ 5. Exit                                       ║
echo ╚═══════════════════════════════════════════════╝
echo.

set /p choice=Select function: 

if "%choice%"=="1" call :activateWindows & goto :show_menu
if "%choice%"=="2" call :activateOffice  & goto :show_menu

if "%choice%"=="3" call :checkStatus     & goto :show_menu
if "%choice%"=="4" call :resetKMS        & goto :show_menu

if "%choice%"=="5" exit /b
echo [ERROR] Error. select (1, 2, 3, 4, 5).
pause

goto :show_menu

:main
call :cmd_clear
if "%_log%"=="1" call :log_clear

color 0A
call :show_banner

call :show_server_menu
exit /b

:checkInstallDocker
docker --version >nul 2>&1

if %errorlevel% neq 0 (
    echo [INFO] Docker not found
    echo [INFO] Install docker via winget...

    winget install --accept-package-agreements --accept-source-agreements -e Docker.DockerDesktop
    timeout /t 10 >nul

    docker --version >nul 2>&1

    if %errorlevel% neq 0 (
        color 0C

        echo [ERROR] Docker not found. Maybe restart system for resolve this shit
        echo or

        echo Install Docker manually: https://www.docker.com/products/docker-desktop
        pause

        exit /b
    )

    call :run_kms_docker
)

color 0A
echo [INFO] Docker installed successfully and ready to use

exit /b

:show_server_menu
echo ╔═══════════════════════════════════════════════╗
echo ║                   KMS Server                  ║
echo ╠═══════════════════════════════════════════════╣
echo ║ 1. Docker vlmcsd server                       ║
echo ║ 2. Online KMS server                          ║
echo ║                                               ║
echo ║ NOTE:                                         ║
echo ║ You use the online server at your own         ║
echo ║ risk                                          ║
echo ╚═══════════════════════════════════════════════╝

set /p choice=Select server: 

if "%choice%"=="1" (
    set "isDocker=1"
    call :checkInstallDocker
    call :show_menu
) else if "%choice%"=="2" (
    set "isDocker=0"
    call :show_menu
) else (
    echo [ERROR] Error. select (1, 2)
    goto :show_server_menu
)

pause
exit /b

:run_kms_docker
echo [INFO] Checking the running KMS Docker container...

docker ps --filter "name=vlmcsd_kms" --format "{{.Names}}" | findstr vlmcsd_kms >nul

if %errorlevel% equ 0 (
    echo [INFO] The vlmcsd_kms container is already running
) else (
    echo [INFO] Container vlmcsd_kms not found. Launching...
    docker image inspect mikolatero/vlmcsd >nul 2>&1

    if %errorlevel% neq 0 (
        echo [INFO] mikolatero/vlmcsd image not found. Downloading...
        docker pull mikolatero/vlmcsd
    )

    docker run -d --name vlmcsd_kms -p 1688:1688 mikolatero/vlmcsd

    if %errorlevel% neq 0 (
        color 0C
        echo [ERROR] Failed to start KMS container

        pause
        exit /b
    )
    echo [INFO] Container KMS run on port: %_vlmcsdPort%
)

color 0A
exit /b

:activateWindows
echo [INFO] Windows activation...

:: Get Windows Edition from registry
for /f "tokens=3" %%i in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID 2^>nul') do (
    set "winEdition=%%i"
)
call :checkEdition

if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /upk

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ipk %KMS_KEY%
if errorlevel 1 (
    echo [ERROR] Failed to install KMS key
) else (
    echo [INFO] KMS key installed successfully
)

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ckms
timeout /t 1 >nul

if "%isDocker%"=="1" (
    echo [INFO] Docker KMS server...

    set "DOCKER_IP=127.0.0.1"
    echo [INFO] We use IP for activation: !DOCKER_IP!

    %_cscript% //nologo "%windir%\System32\slmgr.vbs" /skms !DOCKER_IP!:%_vlmcsdPort%
    %_cscript% //nologo "%windir%\System32\slmgr.vbs" /ato
) else (
    echo [INFO] Online KMS servers...
    call :tryOnlineKMS
)

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /xpr

pause
exit /b

:tryOnlineKMS
setlocal EnableDelayedExpansion

:: Servers list
set "kmsServers=kms.srv.crsoo.com cy2617.jios.org kms.digiboy.ir kms.cangshui.net kms.library.hk hq1.chinancce.com kms.loli.beer kms.v0v.bid 54.223.212.31 kms.jm33.me nb.shenqw.win kms.izetn.cn kms.cin.ink 222.184.9.98 kms.ijio.net fourdeltaone.net:1688 kms.iaini.net kms.cnlic.com kms.51it.wang key.17108.com kms.chinancce.com kms.ddns.net windows.kms.app kms.ddz.red franklv.ddns.net kms.mogeko.me k.zpale.com amrice.top m.zpale.com mvg.zpale.com kms.shuax.com kensol263.imwork.net:1688 xykz.f3322.org kms789.com dimanyakms.sytes.net:1688 kms8.MSGuides.com kms.03k.org:1688 kms.ymgblog.com kms.bige0.com kms9.MSGuides.com kms.cz9.cn kms.lolico.moe kms.ddddg.cn kms.zhuxiaole.org kms.moeclub.org kms.lotro.cc zh.us.to noair.strangled.net:1688"

for %%S in (!kmsServers!) do (
    echo [INFO] Test server: %%S

    %_cscript% //nologo "%windir%\System32\slmgr.vbs" /skms %%S
    %_cscript% //nologo "%windir%\System32\slmgr.vbs" /ato >nul 2>&1

    %_cscript% //nologo "%windir%\System32\slmgr.vbs" /xpr | find /I "permanently" >nul
    if !errorlevel! equ 0 (
        echo [INFO] Activated: %%S

        endlocal
        goto :eof
    ) else (
        echo [WARN] Failure: %%S
    )
)

echo [ERROR] It was not possible to activate the online KMS server because of everything
endlocal
exit /b

:checkEdition
set "winEdition=%winEdition: =%"
echo [INFO] Installed Windows Edition: %winEdition%

echo %winEdition% | find /I "Home" >nul

if %errorlevel% equ 0 (
    color 0C
    echo [ERROR] Windows Home is not supported by this KMS script
    echo Please upgrade to Windows Pro, Education, or Enterprise

    pause
    exit /b
)

:: Get key
set "KMS_KEY="

if /I "%winEdition%"=="Windows10Education"                         set "KMS_KEY=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"
if /I "%winEdition%"=="Windows10EducationN"                        set "KMS_KEY=84NGF-MHBT6-FXBX8-QWJK7-DRR8H"
if /I "%winEdition%"=="Windows10Enterprise"                        set "KMS_KEY=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"
if /I "%winEdition%"=="Windows10EnterpriseN"                       set "KMS_KEY=3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT"
if /I "%winEdition%"=="Windows10EnterpriseLTSB"                    set "KMS_KEY=FWN7H-PF93Q-4GGP8-M8RF3-MDWWW"
if /I "%winEdition%"=="Windows10EnterpriseLTSB2016"                set "KMS_KEY=NK96Y-D9CD8-W44CQ-R8YTK-DYJWX"
if /I "%winEdition%"=="Windows10EnterpriseLTSC2019"                set "KMS_KEY=43TBQ-NH92J-XKTM7-KT3KK-P39PB"
if /I "%winEdition%"=="Windows10EnterpriseNLTSB"                   set "KMS_KEY=NTX6B-BRYC2-K6786-F6MVQ-M7V2X"
if /I "%winEdition%"=="Windows10EnterpriseNLTSB2016"               set "KMS_KEY=2DBW3-N2PJG-MVHW3-G7TDK-9HKR4"
if /I "%winEdition%"=="Windows10IoTEnterprise"                     set "KMS_KEY=XQQYW-NFFMW-XJPBH-K8732-CKFFD"
if /I "%winEdition%"=="Windows10IoTEnterpriseSubscription"         set "KMS_KEY=P8Q7T-WNK7X-PMFXY-VXHBG-RRK69"
if /I "%winEdition%"=="Windows10IoTEnterpriseLTSC2021"             set "KMS_KEY=QPM6N-7J2WJ-P88HH-P3YRH-YY74H"
if /I "%winEdition%"=="Windows10IoTEnterpriseLTSC2024"             set "KMS_KEY=CGK42-GYN6Y-VD22B-BX98W-J8JXD"
if /I "%winEdition%"=="Windows10IoTEnterpriseLTSCSubscription2024" set "KMS_KEY=N979K-XWD77-YW3GB-HBGH6-D32MH"
if /I "%winEdition%"=="Windows10Pro"                               set "KMS_KEY=VK7JG-NPHTM-C97JM-9MPGT-3V66T"
if /I "%winEdition%"=="Windows10ProN"                              set "KMS_KEY=2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"
if /I "%winEdition%"=="Windows10ProEducation"                      set "KMS_KEY=8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"
if /I "%winEdition%"=="Windows10ProEducationN"                     set "KMS_KEY=GJTYN-HDMQY-FRR76-HVGC7-QPF8P"
if /I "%winEdition%"=="Windows10ProWorkstations"                   set "KMS_KEY=DXG7C-N36C4-C4HTG-X4T3X-2YV77"
if /I "%winEdition%"=="Windows10ProNWorkstations"                  set "KMS_KEY=WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"
if /I "%winEdition%"=="Windows10S"                                 set "KMS_KEY=V3WVW-N2PV2-CGWC3-34QGF-VMJ2C"
if /I "%winEdition%"=="Windows10SN"                                set "KMS_KEY=NH9J3-68WK7-6FB93-4K3DF-DJ4F6"
if /I "%winEdition%"=="Windows10SE"                                set "KMS_KEY=KY7PN-VR6RX-83W6Y-6DDYQ-T6R4W"
if /I "%winEdition%"=="Windows10SEN"                               set "KMS_KEY=K9VKN-3BGWV-Y624W-MCRMQ-BHDCD"
if /I "%winEdition%"=="Windows10Team"                              set "KMS_KEY=XKCNC-J26Q9-KFHD2-FKTHY-KD72Y"
if /I "%winEdition%"=="Enterprise"                                 set "KMS_KEY=NPPR9-FWDCX-D2C8J-H872K-2YT43"
if /I "%winEdition%"=="Professional"                               set "KMS_KEY=W269N-WFGWX-YVC9B-4J6C9-T83GX"
if /I "%winEdition%"=="Education"                                  set "KMS_KEY=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"

if not defined KMS_KEY (
    echo [INFO] Edition not recognized, using default Windows 10 Pro key
    set "KMS_KEY=VK7JG-NPHTM-C97JM-9MPGT-3V66T"
)

:resetKMS
echo [INFO] Deleting the key and resetting the KMS settings...

if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /upk
timeout /t 1 >nul

%_cscript% //nologo "%windir%\System32\slmgr.vbs" /ckms
echo [INFO] KMS reset

exit /b

:checkStatus
echo.
echo [INFO] Status Windows:

if exist "%windir%\SysNative\cscript.exe" (
    set "_cscript=%windir%\SysNative\cscript.exe"
) else (
    set "_cscript=%windir%\System32\cscript.exe"
)

%_cscript% //nologo "%windir%\system32\slmgr.vbs" /dli
echo.

echo [INFO] Status Office:
set "officeScript="

for /f "delims=" %%A in ('dir /b /s ospp.vbs 2^>nul') do (
    set "officeScript=%%A"
    goto :officeCheck
)

echo [WARN] ospp.vbs not found
pause

exit /b

:officeCheck
echo --- !officeScript! ---
cscript //nologo "!officeScript!" /dstatus

echo.
pause
exit /b

:activateOffice
echo [INFO] Find Office...
set "officeScript="

for /f "delims=" %%A in ('dir /b /s ospp.vbs 2^>nul') do (
    set "officeScript=%%A"
    goto :officeFound
)

echo [ERROR] ospp.vbs not found
pause

exit /b

:officeFound
echo [INFO] Find: !officeScript!"

goto :officeMenu

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

:show_footer
echo MIT license
echo github.com/fxhxyz4/kms
echo.
echo.
echo.
exit /b