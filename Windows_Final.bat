@echo off
title Windows Activation Menu
cls

:: Check for administrator privileges
net session >nul 2>&1
if errorlevel 1 (
    color 0C
    echo [ERROR] This script requires administrator privileges. Restarting as admin...
    echo.
    powershell -Command "Start-Process cmd.exe -ArgumentList '/c %~dpnx0' -Verb RunAs"
    exit /b
)

:: Define color codes
set "WHITE=0F"
set "GREEN=0A"
set "RED=0C"
set "YELLOW=0E"
set "CYAN=0B"
set "ERROR_COLOR=0C"

:: Main menu
:menu
cls
color 0B
echo ============================================================================ 
echo [MENU] Select which version of Windows to activate or view information: 
echo ============================================================================ 
echo [1] Windows 7 Professional/Enterprise 
echo [2] Windows 8/8.1 
echo [3] Windows 10 
echo [4] Remove Product Key 
echo ============================================================================ 
echo [0] Exit 
echo ============================================================================ 
color 0F
set /p choice="Enter your choice [0-5]: "

if "%choice%"=="1" goto win7
if "%choice%"=="2" goto win8
if "%choice%"=="3" goto win10
if "%choice%"=="4" goto removekey
if "%choice%"=="0" exit
goto invalid

:: Install and retry mechanism for keys
:install_key
set KEYLIST=%~1
set KEY_INDEX=0
set KEY_RESULT=1
for %%K in (%KEYLIST%) do (
    if !KEY_RESULT! NEQ 0 (
        echo Installing key: %%K...
        slmgr /ipk %%K >nul 2>&1
        set KEY_RESULT=!ERRORLEVEL!
    )
)
exit /b !KEY_RESULT!

:: Retry mechanism for KMS servers
:skms
set KMS_SERVERS=kms.lotro.cc mhd.kmdns.net110 kms.digiboy.ir hq1.chinancce.com 54.223.212.31 kms.cnlic.com
set SKMS_RESULT=1
for %%S in (%KMS_SERVERS%) do (
    if !SKMS_RESULT! NEQ 0 (
        echo Trying KMS server: %%S...
        slmgr /skms %%S:1688 >nul 2>&1
        slmgr /ato >nul 2>&1
        set SKMS_RESULT=!ERRORLEVEL!
        if !SKMS_RESULT! EQU 0 echo [SUCCESS] Activated with server %%S.
    )
)
exit /b !SKMS_RESULT!

:win7
color 0E
echo [MENU] Activating Windows 7 Professional/Enterprise...
color 0F
call :install_key "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4 MRPKT-YTG23-K7D7T-X2JMM-QY7MG"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install any product key.
    goto end
)
call :skms
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to activate Windows 7.
    goto end
)
goto ato_success

:win8
color 0E
echo [MENU] Activating Windows 8/8.1...
color 0F
call :install_key "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 FB4WR-32NVD-4RW79-XQFWH-CYQG3"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install any product key.
    goto end
)
call :skms
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to activate Windows 8/8.1.
    goto end
)
goto ato_success

:win10
color 0E
echo [MENU] Activating Windows 10...
color 0F
call :install_key "R6JHF-DN3DX-D7FRF-V3YD3-DV66T MH37W-N47XK-V7XM9-C7227-GCQG9"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install any product key.
    goto end
)
call :skms
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to activate Windows 10.
    goto end
)
goto ato_success

:ato_success
color 0A
echo ============================================================================ 
echo [SUCCESS] Windows activated successfully.
goto end

:ato_failed
color 0C
echo ============================================================================ 
echo [ERROR] Failed to activate Windows.
goto end

:removekey
color 0E
echo [INFO] Removing the product key... 
color 0F
slmgr /upk >nul
slmgr /cpky >nul
echo [SUCCESS] Product key removed.
goto end

:invalid
color 0C
echo Invalid choice! Please select a valid option [0-5].
goto menu

:end
PAUSE
goto menu
