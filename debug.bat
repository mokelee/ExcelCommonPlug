@echo off
chcp 65001 >nul 2>&1
:: Build and launch Excel with add-in (32-bit)

:: Close Excel if running
tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if %ERRORLEVEL% EQU 0 (
    echo [0] Closing Excel...
    taskkill /IM EXCEL.EXE /F >nul 2>&1
    timeout /t 2 /nobreak >nul
)

:: Remove installed version registry to avoid conflict with debug
echo     Cleaning addin registry...
setlocal enabledelayedexpansion
for %%V in (14.0 15.0 16.0) do (
    set "REGKEY=HKCU\Software\Microsoft\Office\%%V\Excel\Options"
    for /f "tokens=1,2,*" %%a in ('reg query "!REGKEY!" 2^>nul ^| findstr /i "OPEN"') do (
        echo %%c | findstr /i "ExcelCommonTools FiKingdee" >nul
        if !errorlevel! equ 0 (
            reg delete "!REGKEY!" /v "%%a" /f >nul 2>&1
        )
    )
)
endlocal

:: Read version from config.iss
set VERSION=1.0.0
for /f "tokens=3" %%a in ('findstr /c:"#define MyAppVersion" installer\config.iss') do (
    set "VERSION=%%~a"
)

:: Build
echo [1/2] Building (v%VERSION%)...
dotnet build src\ExcelCommonTools.csproj -c Debug --no-incremental /p:Version=%VERSION%
if %ERRORLEVEL% NEQ 0 (
    echo Build FAILED!
    pause
    exit /b 1
)

:: Launch Excel
echo [2/2] Starting Excel with add-in...
start "" "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /x "%~dp0src\bin\Debug\net48\ExcelCommonTools-AddIn64.xll" "D:\Excel插件测试\Excel插件测试文件.xlsx"
