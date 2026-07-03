@echo off
setlocal enabledelayedexpansion

set "INSTALL_DIR=%~dp0"
set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

set "BASE_NAME="
for /f "tokens=1,* delims==" %%a in ('type "%INSTALL_DIR%\addin.ini" ^| findstr /i "BaseName"') do set "BASE_NAME=%%b"
for /f "tokens=*" %%x in ("%BASE_NAME%") do set "BASE_NAME=%%x"
if "%BASE_NAME%"=="" exit /b 1

for %%V in (16.0 15.0 14.0) do (
    set "RK=HKCU\Software\Microsoft\Office\%%V\Excel\Options"
    for %%i in (OPEN OPEN1 OPEN2 OPEN3 OPEN4 OPEN5) do (
        for /f "tokens=2,*" %%a in ('reg query "!RK!" /v "%%i" 2^>nul ^| findstr REG_SZ') do (
            echo %%b | findstr /i "%BASE_NAME%" >nul
            if !errorlevel! equ 0 reg delete "!RK!" /v "%%i" /f >nul 2>&1
        )
    )
)
exit /b 0
