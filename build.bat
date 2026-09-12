@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo   Excel-DNA AddIn Build Script
echo ============================================
echo.

:: 从 CHANGELOG.md 读取版本号（版本号唯一数据源）
:: 首个标题行 "# [x.y.z] - 日期" 即为最新版本
set VERSION=1.0.0
for /f "usebackq tokens=2 delims=[]" %%a in (`findstr /r /c:"^# \[" CHANGELOG.md`) do (
    set "VERSION=%%a"
    goto :got_version
)
:got_version

:: 生成 installer\version.iss，供 setup.iss #include（Inno Setup 使用）
> "%~dp0installer\version.iss" echo ; 本文件由 build.bat 自动生成，请勿手动修改。版本号唯一来源为 CHANGELOG.md
>> "%~dp0installer\version.iss" echo #define MyAppVersion "%VERSION%"

set CONFIG=Release
set SRC_DIR=%~dp0src
set INSTALLER_DIR=%~dp0installer
set UPDATE_CLIENT_DIR=f:\CodeGitHub\UpdateClient
set CLEANER_DIR=%~dp0ExcelCleaner
set CLEANER_EXE_DIR=%~dp0ExcelCleanerExe

:: ============================================
:: 步骤1: 编译 ExcelCommonTools 插件
:: ============================================
echo [1/5] Building ExcelCommonTools (v%VERSION%)...
dotnet build "%SRC_DIR%\ExcelCommonTools.csproj" -c %CONFIG% /p:Version=%VERSION%
if errorlevel 1 (
    echo Error: ExcelCommonTools build failed!
    pause
    exit /b 1
)
echo      OK
echo.

:: ============================================
:: 步骤2: 编译 UpdateClient 升级程序
:: ============================================
echo [2/5] Building UpdateClient...
dotnet build "%UPDATE_CLIENT_DIR%\src\UpdateClient.csproj" -c %CONFIG%
if errorlevel 1 (
    echo Error: UpdateClient build failed!
    pause
    exit /b 1
)
echo      OK
echo.

:: ============================================
:: 步骤3: 校验 ExcelCleaner 后台进程清理工具是否存在
::        (不再每次重新编译，直接使用 ExcelCleanerExe 目录的现成文件)
:: ============================================
echo [3/5] Checking ExcelCleaner.exe...
if not exist "%CLEANER_EXE_DIR%\ExcelCleaner.exe" (
    echo Error: ExcelCleaner.exe not found at "%CLEANER_EXE_DIR%\ExcelCleaner.exe"
    echo Please build it once manually:
    echo   dotnet publish "%CLEANER_DIR%\ExcelCleaner.csproj" -c %CONFIG% -r win-x64
    echo   copy /y "%CLEANER_DIR%\bin\%CONFIG%\net10.0\win-x64\publish\ExcelCleaner.exe" "%CLEANER_EXE_DIR%\ExcelCleaner.exe"
    pause
    exit /b 1
)
echo      OK
echo.

:: ============================================
:: 步骤4: 同步版本号到 installer 文件
:: ============================================
echo [4/5] Writing version %VERSION%...
echo %VERSION%> "%INSTALLER_DIR%\version.txt"
:: 同步到 addin.ini (保留 BaseName，版本由 .version 文件管理)
echo      Done
echo.

:: ============================================
:: Step 5: Package installer (Inno Setup)
:: ============================================
echo [5/5] Packaging installer...

:: 尝试常见的 Inno Setup 安装路径
set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
) else if exist "D:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC=D:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)

if "%ISCC%"=="" (
    echo Warning: Inno Setup 6 not found, skip installer generation.
    echo Download: https://jrsoftware.org/isdl.php
    echo.
    echo Build output:
    echo   Plugin: %SRC_DIR%\bin\%CONFIG%\net48\
    echo   Updater: %UPDATE_CLIENT_DIR%\src\bin\%CONFIG%\net48\UpdateClient.exe
    pause
    exit /b 0
)

"%ISCC%" "%INSTALLER_DIR%\setup.iss"
if errorlevel 1 (
    echo Error: Installer build failed!
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build complete!
echo   Output: installer\output\ExcelCommonTools_Setup_%VERSION%.exe
echo ============================================
echo.
