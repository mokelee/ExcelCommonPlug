# Excel-DNA 插件通用安装包方案

## 概述

本项目使用 [Inno Setup 6](https://jrsoftware.org/isdl.php) 生成 Windows 安装程序（.exe），实现：

- 自动检测 Excel 32/64 位，注册对应的 xll 文件
- 安装到用户目录（`%APPDATA%`），无需管理员权限
- 支持 Office 2010/2013/2016/2019/365 多版本
- 集成 UpdateClient 自动升级客户端
- 中文/英文双语安装界面
- 卸载时自动清理注册表和临时文件

## 目录结构

```
installer/
├── config.iss            ← 项目参数配置（每个项目必改）
├── setup.iss             ← 通用安装脚本（无需修改）
├── register.bat          ← 通用 xll 注册脚本（无需修改）
├── unregister.bat        ← 通用 xll 反注册脚本（无需修改）
├── addin.ini             ← 运行时配置，BaseName 须与 config.iss 一致
├── ChineseSimplified.isl ← 中文语言包
├── version.txt           ← 由 build.bat 自动生成
└── output/               ← 安装包输出目录
    └── XXX_Setup_x.x.x.exe
```

## 一键编译

双击项目根目录的 `build.bat`，自动完成：

1. 编译 ExcelCommonTools 插件（dotnet build）
2. 编译 UpdateClient 升级程序（dotnet build）
3. 同步版本号到 version.txt
4. 调用 Inno Setup 生成安装包

最终输出：`installer\output\ExcelCommonTools_Setup_1.0.0.exe`

### 前置依赖

| 工具 | 用途 | 下载地址 |
|------|------|----------|
| .NET SDK | 编译 C# 项目 | https://dotnet.microsoft.com/download |
| Inno Setup 6 | 打包安装程序 | https://jrsoftware.org/isdl.php |

## 新项目复用指南

将 `installer` 目录整体复制到新的 Excel-DNA 插件项目，只需修改以下内容：

### 1. config.iss（必改）

```iss
#define MyAppName        "你的插件显示名称"
#define MyAppVersion     "1.0.0"
#define MyAppPublisher   "发布者"
#define MyAppURL         "项目主页"
#define MyAppId          "生成一个新的 GUID"
#define AddInBaseName    "YourProject-AddIn"
#define AssemblyName     "YourProject"
#define InstallDirName   "YourProject"
#define BuildOutput      "..\src\bin\Release\net48"
#define UpdateClientExe  "..\..\UpdateClient\src\bin\Release\net48\UpdateClient.exe"
#define OutputName       "YourProject_Setup"
```

> GUID 生成工具：https://www.guidgenerator.com

### 2. addin.ini（必改）

```ini
[AddIn]
BaseName=YourProject-AddIn
```

BaseName 的值必须与 config.iss 中的 `AddInBaseName` 一致。

### 3. setup.iss 的依赖 DLL 段（按需修改）

如果新项目引用了不同的 NuGet 包，在 `[Files]` 段的"依赖 DLL"区域增减对应行：

```iss
; === 依赖 DLL (按项目需要增减) ===
Source: "{#BuildOutput}\Your.Dependency.dll"; DestDir: "{app}"; Flags: ignoreversion
```

## 注册机制说明

### Excel 加载 xll 的原理

Excel 通过注册表 `HKCU\Software\Microsoft\Office\{版本}\Excel\Options` 下的 `OPEN`、`OPEN1`、`OPEN2`... 键来加载 xll 插件。

### register.bat 工作流程

1. 从 `addin.ini` 读取 `BaseName`
2. 检测 Excel 位数（ClickToRun → MSI Bitness 两种方式）
3. 32 位选择 `{BaseName}.xll`，64 位选择 `{BaseName}64.xll`
4. 遍历 Office 14.0/15.0/16.0 版本
5. 如已注册则更新路径，否则找空闲的 OPEN 键写入

### unregister.bat 工作流程

遍历所有 Office 版本的 OPEN 键，移除包含 `BaseName` 的条目。

## 常见问题

### Q: 安装后 Excel 提示"文件格式和扩展名不匹配"

**原因**：Excel 位数与注册的 xll 不匹配（如 32 位 Excel 加载了 64 位 xll）。

**解决**：确认 `register.bat` 的位数检测逻辑正确。对于 MSI 安装的旧版 Office，可能需要在 register.bat 中添加更多检测路径。

### Q: 安装后 Excel 没有出现插件

可能原因：
1. Excel 在安装时未关闭，xll 文件被锁定
2. 注册表写入的 Office 版本与实际不匹配
3. 被 Excel 的安全策略禁用

排查方式：
- 打开注册表编辑器，查看 `HKCU\Software\Microsoft\Office\16.0\Excel\Options` 下是否有 OPEN 键
- Excel → 文件 → 选项 → 加载项 → 查看是否在"非活动应用程序加载项"中

### Q: 如何更新版本号

修改 `config.iss` 中的 `MyAppVersion`，然后重新运行 `build.bat`。版本号会自动同步到安装包文件名和 `.version` 文件。

## 自动升级集成

安装包中包含 `UpdateClient.exe`，插件启动时会后台调用升级服务器的 API 检查更新：

```
GET /api/v1/products/{product_id}/check?version={当前版本}
```

如有新版本，提示用户确认后启动 UpdateClient 执行升级。详细流程参见 UpdateClient 项目的 README。

### 相关配置

| 配置项 | 位置 | 说明 |
|--------|------|------|
| product_id | `src\Core\AppConstants.cs` | 服务器端产品标识，须与服务器一致 |
| server_url | `src\Core\AppConstants.cs` | 升级服务器地址 |

## 更换应用图标

安装程序图标和控制面板中显示的图标由 `installer\app.ico` 决定。

### 更换步骤

1. 准备一个 `.ico` 格式的图标文件（建议包含 256×256、48×48、32×32、16×16 多尺寸）
2. 将文件命名为 `app.ico`，放入 `installer\` 目录（覆盖原文件）
3. 重新运行 `build.bat` 打包

### 图标生效位置

| 位置 | 说明 |
|------|------|
| 安装程序 exe 文件图标 | 用户双击安装时看到的图标 |
| 安装向导窗口左上角 | 安装过程中窗口标题栏图标 |
| 控制面板 / 设置 → 应用 | 卸载列表中显示的图标 |

### 图标要求

- 格式：`.ico`（不支持 png/jpg 直接使用）
- 建议尺寸：包含 16×16、32×32、48×48、256×256
- 在线转换工具：https://convertico.com 或 https://icoconvert.com
- 文件大小建议不超过 500KB

### 相关配置项（setup.iss）

```iss
SetupIconFile=app.ico              ; 安装程序 exe 的图标
UninstallDisplayIcon={app}\app.ico ; 控制面板卸载列表中的图标
```

如需更改图标文件名，同时修改 `setup.iss` 中上述两行以及 `[Files]` 段中的 `Source: ".\app.ico"` 行。
