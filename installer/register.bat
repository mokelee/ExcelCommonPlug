@echo off
setlocal enabledelayedexpansion

set "INSTALL_DIR=%~dp0"
set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

:: Read BaseName from addin.ini
set "BASE_NAME="
for /f "tokens=1,* delims==" %%a in ('type "%INSTALL_DIR%\addin.ini" ^| findstr /i "BaseName"') do set "BASE_NAME=%%b"
for /f "tokens=*" %%x in ("%BASE_NAME%") do set "BASE_NAME=%%x"
if "%BASE_NAME%"=="" exit /b 1

:: Detect Excel bitness
set "IS64=0"
for /f "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul ^| findstr /i "Platform"') do (
    echo %%b | findstr /i "x64" >nul
    if !errorlevel! equ 0 set "IS64=1"
)
if "!IS64!"=="0" (
    for %%V in (16.0 15.0 14.0) do (
        if "!IS64!"=="0" (
            for /f "tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\%%V\Outlook" /v Bitness 2^>nul ^| findstr /i "Bitness"') do (
                echo %%b | findstr /i "x64" >nul
                if !errorlevel! equ 0 set "IS64=1"
            )
        )
    )
)

if "!IS64!"=="1" (set "XLL_FILE=%BASE_NAME%64.xll") else (set "XLL_FILE=%BASE_NAME%.xll")
set "XLL_PATH=%INSTALL_DIR%\%XLL_FILE%"
echo [register] IS64=!IS64!, XLL=%XLL_FILE% >> "%INSTALL_DIR%\register_debug.log"

:: Register for detected Office version only (check 16.0 first as most common)
for %%V in (16.0 15.0 14.0) do (
    reg query "HKCU\Software\Microsoft\Office\%%V\Excel" >nul 2>&1
    if !errorlevel! equ 0 call :Reg "HKCU\Software\Microsoft\Office\%%V\Excel\Options" "%XLL_PATH%" "%BASE_NAME%"
)
exit /b 0

:Reg
set "RK=%~1"
set "XL=%~2"
set "MT=%~3"
:: Check OPEN through OPEN5 only (enough for most cases)
for %%i in (OPEN OPEN1 OPEN2 OPEN3 OPEN4 OPEN5) do (
    for /f "tokens=2,*" %%a in ('reg query "%RK%" /v "%%i" 2^>nul ^| findstr REG_SZ') do (
        echo %%b | findstr /i "%MT%" >nul
        if !errorlevel! equ 0 (
            reg add "%RK%" /v "%%i" /t REG_SZ /d "/R \"%XL%\"" /f >nul 2>&1
            exit /b 0
        )
    )
)
:: Find empty slot
for %%i in (OPEN OPEN1 OPEN2 OPEN3 OPEN4 OPEN5) do (
    reg query "%RK%" /v "%%i" >nul 2>&1
    if !errorlevel! neq 0 (
        reg add "%RK%" /v "%%i" /t REG_SZ /d "/R \"%XL%\"" /f >nul 2>&1
        exit /b 0
    )
)
exit /b 0
