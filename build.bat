@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo   Excel-DNA 插件 编译打包脚本
echo ============================================
echo.

:: 从 installer\config.iss 读取版本号
set VERSION=1.0.0
for /f "tokens=3" %%a in ('findstr /c:"#define MyAppVersion" installer\config.iss') do (
    set "VERSION=%%~a"
)

set CONFIG=Release
set SRC_DIR=%~dp0src
set INSTALLER_DIR=%~dp0installer
set UPDATE_CLIENT_DIR=f:\CodeGitHub\UpdateClient

:: ============================================
:: 步骤1: 编译 ExcelCommonTools 插件
:: ============================================
echo [1/4] 编译 ExcelCommonTools (v%VERSION%)...
dotnet build "%SRC_DIR%\ExcelCommonTools.csproj" -c %CONFIG% /p:Version=%VERSION%
if errorlevel 1 (
    echo 错误: ExcelCommonTools 编译失败！
    pause
    exit /b 1
)
echo      编译成功 → src\bin\%CONFIG%\net48\
echo.

:: ============================================
:: 步骤2: 编译 UpdateClient 升级程序
:: ============================================
echo [2/4] 编译 UpdateClient...
dotnet build "%UPDATE_CLIENT_DIR%\src\UpdateClient.csproj" -c %CONFIG%
if errorlevel 1 (
    echo 错误: UpdateClient 编译失败！
    pause
    exit /b 1
)
echo      编译成功 → UpdateClient\src\bin\%CONFIG%\net48\UpdateClient.exe
echo.

:: ============================================
:: 步骤3: 同步版本号到 installer 文件
:: ============================================
echo [3/4] 写入版本号 %VERSION%...
echo %VERSION%> "%INSTALLER_DIR%\version.txt"
:: 同步到 addin.ini (保留 BaseName，版本由 .version 文件管理)
echo      完成
echo.

:: ============================================
:: 步骤4: 打包安装程序 (Inno Setup)
:: ============================================
echo [4/4] 打包安装程序...

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
    echo 警告: 未找到 Inno Setup 6，跳过安装包生成。
    echo 请安装 Inno Setup 6: https://jrsoftware.org/isdl.php
    echo 安装后重新运行此脚本即可生成安装包。
    echo.
    echo 编译产物位于:
    echo   插件: %SRC_DIR%\bin\%CONFIG%\net48\
    echo   升级: %UPDATE_CLIENT_DIR%\src\bin\%CONFIG%\net48\UpdateClient.exe
    pause
    exit /b 0
)

"%ISCC%" "%INSTALLER_DIR%\setup.iss"
if errorlevel 1 (
    echo 错误: 安装包生成失败！
    pause
    exit /b 1
)

echo.
echo ============================================
echo   打包完成！
echo   安装包: installer\output\ExcelCommonTools_Setup_%VERSION%.exe
echo ============================================
echo.
pause
