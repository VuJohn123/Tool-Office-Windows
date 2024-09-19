@echo off
title Microsoft Office Activation
cls
color 0F
mode con cols=100 lines=30

:: Define colors
set "ERROR_COLOR=4F"
set "INFO_COLOR=9F"
set "SUCCESS_COLOR=2F"
set "MENU_COLOR=3F"
set "HEADER_COLOR=1F"
set "BACKGROUND_COLOR=F0"

:: Function to display header
:header
cls
color %HEADER_COLOR%
echo ============================================================================  
echo                      Microsoft Office Activation  
echo ============================================================================  
color %BACKGROUND_COLOR%

:: Check for administrator privileges
net session >nul 2>&1
if errorlevel 1 (
    color %ERROR_COLOR%
    echo [ERROR] This script requires administrator privileges. Restarting as admin...
    echo.
    powershell -Command "Start-Process cmd.exe -ArgumentList '/c %~dpnx0' -Verb RunAs"
    exit /b
)

:: Define potential paths for ospp.vbs based on system architecture
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    set "ARCH=64-bit"
    set "OSPP_PATH_32=%ProgramFiles(x86)%\Microsoft Office\Office"
    set "OSPP_PATH_64=%ProgramFiles%\Microsoft Office\Office"
) else (
    set "ARCH=32-bit"
    set "OSPP_PATH_32=%ProgramFiles%\Microsoft Office\Office"
    set "OSPP_PATH_64=%ProgramFiles(x86)%\Microsoft Office\Office"
)

:: Office 2019 specific path
set "OSPP_PATH_2019=%ProgramFiles(x86)%\Microsoft Office\Office16"

:: Check which version of Office is installed and set the appropriate path for ospp.vbs
set "OSPP_PATH="
set "VERSION="
set "KEY="

:: Check for Office versions
for %%v in (2021 2019 2016 2013 2010) do (
    if %%v==2019 (
        if exist "%OSPP_PATH_2019%\ospp.vbs" (
            set "OSPP_PATH=%OSPP_PATH_2019%"
            set "VERSION=Office %%v 32-bit"
            goto :found
        )
    ) else (
        echo Checking 64-bit path: "%OSPP_PATH_64%\Office%%v\ospp.vbs"
        echo Checking 32-bit path: "%OSPP_PATH_32%\Office%%v\ospp.vbs"
        if exist "%OSPP_PATH_64%\Office%%v\ospp.vbs" (
            set "OSPP_PATH=%OSPP_PATH_64%\Office%%v"
            set "VERSION=Office %%v 64-bit"
            goto :found
        ) else if exist "%OSPP_PATH_32%\Office%%v\ospp.vbs" (
            set "OSPP_PATH=%OSPP_PATH_32%\Office%%v"
            set "VERSION=Office %%v 32-bit"
            goto :found
        )
    )
)

color %ERROR_COLOR%
echo [ERROR] ospp.vbs file not found in any known Office installation path.
pause
exit /b

:found
color %INFO_COLOR%
echo [INFO] Detected %VERSION%.
echo.

:: Add the found path of ospp.vbs to both user and system PATH environment variables
setx PATH "%PATH%;%OSPP_PATH%" >nul
for /f "tokens=1,* delims==" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v PATH 2^>nul') do set "systemPath=%%b"
setx PATH "%systemPath%;%OSPP_PATH%" /M >nul

:: Check if Office is already activated
echo ============================================================================  
echo Checking if Office is already activated...  
echo ============================================================================  
cscript //nologo "%OSPP_PATH%\ospp.vbs" /dstatus | find /i "LICENSED" >nul
if %ERRORLEVEL% EQU 0 (
    color %SUCCESS_COLOR%
    echo [INFO] %VERSION% is already activated.

    :: Ask if user wants to remove the product key
    set /p removeKey="Do you want to remove the product key? (y/n): "
    if /i "%removeKey%"=="y" (
        color %INFO_COLOR%
        echo [INFO] Removing product key...
        cscript "%OSPP_PATH%\ospp.vbs" /unpkey:all
        echo [INFO] Product key removed.
    )

    echo [INFO] Press any key to return to the main menu or Ctrl+C to exit.
    pause >nul
    goto :menu
) else (
    color %ERROR_COLOR%
    echo [INFO] %VERSION% is not activated or not detected.
    pause
)

:: Present menu to user for Office version activation
:menu
call :header
color %MENU_COLOR%
echo [MENU] Select which version of Office to activate:
echo ============================================================================  
echo [1] Office Pro Plus 2010
echo [2] Office Pro Plus 2013
echo [3] Office Pro Plus 2016
echo [4] Office Pro Plus 2019
echo [5] Office Pro Plus 2021
echo [6] Office 365 Mondo
echo ============================================================================  
echo [0] Display Office Installation Status
echo ============================================================================  
color %BACKGROUND_COLOR%
set /p version="Enter your choice (0-6): "
if "%version%"=="0" goto :status
if "%version%"=="1" goto :activate2010
if "%version%"=="2" goto :activate2013
if "%version%"=="3" goto :activate2016
if "%version%"=="4" goto :activate2019
if "%version%"=="5" goto :activate2021
if "%version%"=="6" goto :activate365
goto :menu

:status
call :header
color %INFO_COLOR%
echo [INFO] Displaying Office Installation Status...
echo ============================================================================  
for %%v in (2021 2019 2016 2013 2010) do (
    if exist "%OSPP_PATH_64%\Office%%v\ospp.vbs" (
        echo [INFO] Office %%v 64-bit detected.
    ) else if exist "%OSPP_PATH_32%\Office%%v\ospp.vbs" (
        echo [INFO] Office %%v 32-bit detected.
    ) else if %%v==2019 (
        if exist "%OSPP_PATH_2019%\ospp.vbs" (
            echo [INFO] Office %%v 32-bit detected.
        )
    )
)
pause
goto :menu

:activate2010
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office Pro Plus 2010...
echo ============================================================================  
set "KEY=BDD3G-XM7FB-BD2HM-YK63V-VQFDK"
goto :activate

:activate2013
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office Pro Plus 2013...
echo ============================================================================  
set "KEY=YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
goto :activate

:activate2016
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office Pro Plus 2016...
echo ============================================================================  
set "KEY=JNRGM-WHDWX-FJJG3-K47QV-DRTFM"
goto :activate

:activate2019
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office Pro Plus 2019...
echo ============================================================================  
set "KEY=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
goto :activate

:activate2021
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office Pro Plus 2021...
echo ============================================================================  
set "KEY=XM2V9-DN9HH-QB449-XDGKC-W2RMW"
goto :activate

:activate365
color %INFO_COLOR%
echo ============================================================================  
echo Activating Office 365 Mondo...
echo ============================================================================  
set "KEY=DMTCJ-KNRKX-26982-JYCKT-P7KB6"
goto :activate

:activate
color %INFO_COLOR%
echo [INFO] Setting the port to 1688...
cscript "%OSPP_PATH%\ospp.vbs" /setprt:1688
if %ERRORLEVEL% neq 0 (
    color %ERROR_COLOR%
    echo [ERROR] Failed to set the port. Please check the configuration.
    goto :halt
)

echo [INFO] Entering product key...
cscript //nologo "%OSPP_PATH%\ospp.vbs" /inpkey:%KEY%
if %ERRORLEVEL% neq 0 (
    color %ERROR_COLOR%
    echo [ERROR] Failed to enter the product key.
    goto :halt
)

set /p KMS_SERVER="Enter KMS server (default 107.175.77.7): "
if "%KMS_SERVER%"=="" set "KMS_SERVER=107.175.77.7"
echo [INFO] Setting KMS server to %KMS_SERVER%...
cscript "%OSPP_PATH%\ospp.vbs" /sethst:%KMS_SERVER%
if %ERRORLEVEL% neq 0 (
    color %ERROR_COLOR%
    echo [ERROR] Failed to set the KMS server.
    goto :halt
)

echo [INFO] Attempting to activate...
cscript "%OSPP_PATH%\ospp.vbs" /act
if %ERRORLEVEL% neq 0 (
    color %ERROR_COLOR%
    echo [ERROR] Activation failed.
    goto :halt
)

color %SUCCESS_COLOR%
echo [INFO] Activation attempt complete. Please wait for the result...
pause
goto :end

:halt
color %ERROR_COLOR%
echo ============================================================================  
echo Activation process halted.  
echo ============================================================================  
pause
exit /b

:end
color %BACKGROUND_COLOR%
echo ============================================================================  
echo Operation complete. Press any key to exit.  
echo ============================================================================  
pause
exit /b
